from datetime import timedelta
import json
from pathlib import Path
import unittest
from unittest.mock import patch

from occurrence_comparison import AnalysisCache
from tests import test_paired_adjacent as fixture

DAY, ticket, trace = fixture.DAY, fixture.ticket, fixture.trace


def measured(seconds, modbus, reference='REF', lot='LOT', sn='SN', dims=(104, 50, 195)):
    lines = ticket(seconds, modbus, reference, lot, sn).splitlines(True)
    stamp = (DAY + timedelta(seconds=seconds + 1.5)).strftime('%d|%H:%M:%S.%f')
    lines.insert(2, f'{stamp} IDPK pour {reference}: l={dims[0]} H={dims[1]} L={dims[2]}\n')
    return ''.join(lines)


class ExactComparisonTests(unittest.TestCase):
    setUp = fixture.PairedAdjacentTests.setUp
    write = fixture.PairedAdjacentTests.write
    sync = fixture.PairedAdjacentTests.sync
    select = fixture.PairedAdjacentTests.select
    def test_identity_history_crosses_robots_and_keeps_older_difference(self):
        self.write('R2', trace(measured(1000, '10', dims=(104, 48, 189))))
        self.write('R1', trace(measured(2000, '858'), measured(4000, '1054')))
        self.sync()
        path, number = self.select('REF')
        selected = next((p, n) for p, hits in self.store.search('ModBusId=1054', True) for n, _ in hits)
        result = self.store.comparison_for_occurrence(*selected)
        self.assertTrue(result['different'])
        self.assertEqual(result['basis_label'], 'Mesure antérieure différente')
        self.assertEqual(json.loads(result['pair'][1]['dims'])['L'], 189)
        self.assertIn('R2', str(result['pair'][1]['path']))

    def test_other_sn_lot_sources_and_future_are_excluded(self):
        self.write('R2', trace(measured(1000, '1', sn='OTHER', dims=(1, 1, 1)),
                                measured(1100, '2', lot='OTHER', dims=(1, 1, 1)),
                                measured(5000, '5', dims=(1, 1, 1))))
        self.write('R2', trace(measured(1200, '6', dims=(1, 1, 1))), root=self.other)
        self.write('R1', trace(measured(4000, '1054')))
        self.sync()
        path, hits = self.store.search('ModBusId=1054', True)[0]
        result = self.store.comparison_for_occurrence(path, hits[0][0])
        self.assertEqual(result['status'], 'no_previous')
        self.assertFalse(result['different'])

    def test_selected_measurement_not_first_measurement_of_reference(self):
        self.write('R1', trace(measured(1000, '1', dims=(90, 20, 50)),
                                measured(2000, '2', dims=(104, 50, 195))))
        self.sync()
        path, hits = self.store.search('ModBusId=2', True)[0]
        result = self.store.comparison_for_occurrence(path, hits[0][0])
        self.assertEqual(json.loads(result['current']['dims'])['l'], 104)
        self.assertTrue(result['different'])

    def test_missing_measurement_does_not_borrow_other_identity_or_later_get(self):
        self.write('R2', trace(measured(1000, '1049', sn='OTHER'),
                                ticket(2000, '1052', 'REF', 'LOT', 'SN'),
                                measured(3000, '2000', sn='OTHER', dims=(1, 1, 1))))
        self.sync()
        path, hits = self.store.search('ModBusId=1052', True)[0]
        result = self.store.comparison_for_occurrence(path, hits[0][0])
        self.assertEqual(result['status'], 'missing_measurement')
        self.assertIsNone(result['current'])
        self.assertFalse(result['different'])

    def test_tolerance_incomplete_fields_and_ambiguous_current(self):
        for delta in (2, 3):
            path = self.write('R1', trace(measured(1000, '1', dims=(100, 20, 50)),
                                         measured(2000, '2', dims=(100 + delta, 20, 50))))
            self.sync()
            ref, hits = next((p, h) for p, h in self.store.search('ModBusId=2', True) if not p.historical)
            result = self.store.comparison_for_occurrence(ref, hits[0][0])
            self.assertEqual(result['different'], delta > 2)
        self.write('R2', trace(measured(4000, '3', sn='').replace('SN=<Cr>', '')))
        self.sync()
        path, hits = self.store.search('ModBusId=3', True)[0]
        self.assertEqual(self.store.comparison_for_occurrence(path, hits[0][0])['status'], 'incomplete_identity')
        self.write('R2', trace(measured(4000, '4') + '14|01:06:42.0 IDPK pour REF: l=1 H=2 L=3\n'))
        self.sync()
        path, hits = self.store.search('ModBusId=4', True)[0]
        self.assertEqual(self.store.comparison_for_occurrence(path, hits[0][0])['status'], 'ambiguous_measurement')

    def test_counts_zero_custom_limits_and_warm_cache(self):
        self.write('R1', trace(*(measured(i * 10, str(i), reference='REF' + str(i)) for i in range(80))))
        self.sync()
        path, number = self.select('REF20')
        result = self.store.paired_adjacent(path, '20', number, before_count=12, after_count=47)
        own = result['columns'][0]
        self.assertEqual((len(own['previous']), len(own['following'])), (12, 47))
        with patch.object(self.store, '_rows', wraps=self.store._rows) as rows:
            again = self.store.paired_adjacent(path, '20', number, before_count=12, after_count=47)
            self.assertIs(result, again)
            self.assertEqual(rows.call_count, 0)
        empty = self.store.paired_adjacent(path, '20', number, before_count=0, after_count=0)['columns'][0]
        self.assertFalse(empty['previous'] or empty['following'])
        maximum = self.store.paired_adjacent(path, '20', number, before_count=1000, after_count=1000)['columns'][0]
        self.assertEqual((len(maximum['previous']), len(maximum['following'])), (20, 59))
        for bad in (-1, 1001, '5', 1.5, True):
            with self.assertRaises(ValueError):
                self.store.paired_adjacent(path, '20', number, before_count=bad)

    def test_cache_invalidation_growth_rotation_metadata_and_reopen(self):
        source = self.write('R1', trace(measured(1000, '1')))
        self.sync()
        path, number = self.select('REF')
        old = self.store.comparison_for_occurrence(path, number)
        revision = self.store.analysis_revision()
        self.write('R1', trace(measured(1000, '1'), measured(2000, '2', dims=(1, 1, 1))))
        self.sync()
        self.assertGreater(self.store.analysis_revision(), revision)
        new_path, hits = self.store.search('ModBusId=2', True)[0]
        self.assertTrue(self.store.comparison_for_occurrence(new_path, hits[0][0])['different'])
        self.assertEqual(old['status'], 'no_previous')
        Path(source).unlink()
        self.store.sync([])
        archived = self.store.comparison_for_occurrence(new_path, hits[0][0])
        self.assertTrue(archived['current']['path'].historical)

    def test_small_cache_does_not_disable_large_identity_history(self):
        self.write('R1', trace(*(measured(i * 10, str(i)) for i in range(20))))
        self.sync()
        self.store._analysis_cache = AnalysisCache(budget=5)
        path, hits = self.store.search('ModBusId=19', True)[0]
        result = self.store.comparison_for_occurrence(path, hits[0][0])
        self.assertEqual(result['status'], 'same')
        self.assertLessEqual(self.store._analysis_cache.weight, 5)

    def test_equal_time_previous_candidates_and_missing_header(self):
        self.write('R1', trace(measured(1000,'1',dims=(1,1,1)), measured(3000,'3')))
        self.write('R2', trace(measured(1000,'2',dims=(2,2,2))))
        self.sync()
        path,hits=self.store.search('ModBusId=3',True)[0]
        result=self.store.comparison_for_occurrence(path,hits[0][0])
        self.assertEqual(result['status'],'ambiguous_previous')
        self.assertTrue(result['different'])
        self.assertIsNone(result['pair'])
        self.assertEqual(len(result['candidates']),2)
        self.write('R1', measured(4000,'4'))
        self.sync()
        path,hits=self.store.search('ModBusId=4',True)[0]
        self.assertEqual(self.store.comparison_for_occurrence(path,hits[0][0])['status'],'unknown_time')

    def test_query_index_additive_and_following_link_chain(self):
        self.write('R1', trace(measured(1000,'1')), 'TracOmega017.old')
        self.write('R1', trace(measured(2000,'2'),previous='TracOmega017.old'), 'TracOmega018.old')
        self.write('R1', trace(measured(3000,'3'),previous='TracOmega018.old'))
        self.sync()
        revision=self.store.analysis_revision()
        with self.store.connection() as db:
            count=db.execute('SELECT count(*) FROM lines').fetchone()[0]
        self.store.prepare_analysis()
        with self.store.connection() as db:
            self.assertEqual(db.execute('SELECT count(*) FROM lines').fetchone()[0],count)
            plan=db.execute('''EXPLAIN QUERY PLAN SELECT * FROM lines WHERE file_id=1
                AND modbus IS NOT NULL AND ref IS NOT NULL AND time<=? ORDER BY time DESC LIMIT 1''',(10000000000,)).fetchall()
        self.assertIn('ticket_time',str([tuple(row) for row in plan]))
        self.assertEqual(self.store.analysis_revision(),revision)
        path,hits=self.store.search('ModBusId=1',True)[0]
        result=self.store.paired_adjacent(path,'1',hits[0][0],after_count=30)
        self.assertEqual([row['modbus'] for row in result['columns'][0]['following']],['2','3'])

    def test_get_without_new_ticket_cannot_inherit_earlier_sn_lot(self):
        orphan=''.join(measured(1500,'999',dims=(1,1,1)).splitlines(True)[1:])
        self.write('R1',trace(ticket(1000,'1','REF','LOT','SN'),orphan,measured(2000,'2')))
        self.sync()
        path,hits=self.store.search('ModBusId=1',True)[0]
        self.assertEqual(self.store.comparison_for_occurrence(path,hits[0][0])['status'],'missing_measurement')
        path,hits=self.store.search('ModBusId=2',True)[0]
        result=self.store.comparison_for_occurrence(path,hits[0][0])
        self.assertFalse(result['different'])
        self.assertEqual(result['status'],'no_previous')

    def test_no_shared_dimensions_is_not_reported_as_matching(self):
        first=measured(1000,'1').replace('l=104 H=50 L=195','l=104')
        second=measured(2000,'2').replace('l=104 H=50 L=195','H=50')
        empty=measured(3000,'3').replace('l=104 H=50 L=195','sans dimensions')
        self.write('R1',trace(first,second,empty)); self.sync()
        path,hits=self.store.search('ModBusId=2',True)[0]
        self.assertEqual(self.store.comparison_for_occurrence(path,hits[0][0])['status'],'no_previous')
        path,hits=self.store.search('ModBusId=3',True)[0]
        result=self.store.comparison_for_occurrence(path,hits[0][0])
        self.assertEqual(result['status'],'missing_dimensions')
        self.assertFalse(result['different'])
