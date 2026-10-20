from datetime import datetime
from pathlib import Path
import unittest
from unittest.mock import patch

import trace_sources
from trace_identity import TraceRef
from tracking_view import tracking_values
import tempfile
from trace_store import TraceStore


class TrackingValuesTests(unittest.TestCase):
    def test_tracking_filters_exact_ticket_identity_and_missing_get_ticket(self):
        original = Path('tests/fixtures/current_r2.txt').read_text(encoding='utf8')
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            path = root/'TracOmega.txt'
            second = original.replace('SN=343272', 'SN=3432720').replace('261013.127', '261013.1270')
            orphan = '\n'.join(line for line in original.splitlines() if 'ModBusId=' not in line)
            path.write_text(original+second+orphan, encoding='utf8')
            store = TraceStore(root/'cache.sqlite3')
            store.sync([str(path)])
            self.assertEqual(len(store.tracking('A0061574AC')), 3)
            self.assertEqual(len(store.tracking('A0061574AC', sn='343272')), 1)
            self.assertEqual(len(store.tracking('A0061574AC', sn='3432720', lot='261013.1270')), 1)
            self.assertEqual(store.tracking('A0061574AC', sn='343272', lot='261013.1270'), [])
            self.assertEqual(store.tracking('A0061574AC', sn='34327'), [])
            self.assertEqual(store.tracking('A0061574AC', lot='261013.12'), [])
            with store.snapshot():
                self.assertEqual(len(store.tracking('A0061574AC', lot='261013.127')), 2)

    def test_columns_preserve_fractional_time_and_exact_position(self):
        with patch.object(trace_sources, 'application_directory', return_value=str(Path('project').resolve())), \
                patch.object(trace_sources, 'desktop_directory', return_value=str(Path('desktop').resolve())):
            root = Path('desktop/traces').resolve()
            path = TraceRef(str(root / 'trace r1' / 'TracOmega019.old'), 'saved-generation',
                            datetime(2026, 10, 13).timestamp(), datetime(2026, 10, 15).timestamp())
            for suffix, expected in (('face:1 209', 'Face 1 · 209'), ('face:2 F32', 'Face 2 · F32'),
                                     ('face:1 $0101', 'Face 1 · $0101'), ('face:2', 'Face 2'),
                                     ('no position', '—')):
                row = dict(path=path, time=datetime(2026, 10, 14, 6, 34, 12).timestamp(),
                           text='G 14|06:34:12.2 GRIPPER PUTE 1x boîte(s): Code:REF ' + suffix)
                self.assertEqual(tracking_values(row, [root]),
                                 ('14/10/2026', '06:34:12.2', 'R1', expected, 'TracOmega019.old', '—', '—', '—'))
                self.assertEqual(path.version_id, 'saved-generation')

    def test_multiple_sources_keep_short_provenance(self):
        project = Path('project').resolve()
        desktop = Path('desktop').resolve()
        with patch.object(trace_sources, 'application_directory', return_value=str(project)), \
                patch.object(trace_sources, 'desktop_directory', return_value=str(desktop)):
            row = dict(path=str(desktop / 'traces/R2/TracOmega.txt'), time=0,
                       text='G 01|01:02:03.0 GRIPPER PUTE face:1')
            self.assertEqual(tracking_values(row, [project / 'traces', desktop / 'traces'])[2], 'Bureau R2')


class TrackingCriteriaTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.path = self.root/'traces'/'R1'/'TracOmega.txt'
        self.path.parent.mkdir(parents=True)
        self.store = TraceStore(self.root/'cache.sqlite3', directories=[str(self.root/'traces')])
        original = (Path(__file__).parent/'fixtures/current_r2.txt').read_text(encoding='utf8')
        self.header = '\n'.join(original.splitlines()[:2])+'\n'
        self.ticket = next(line for line in original.splitlines() if 'ModBusId=' in line)
        self.get = next(line for line in original.splitlines() if 'GRIPPER GET' in line)
        self.put = next(line for line in original.splitlines() if 'GRIPPER PUTE' in line)

    def event(self, template, ref, sn='SN1', lot='LOT1'):
        return template.replace('A0061574AC', ref).replace('343272', sn).replace('261013.127', lot)+'\n'

    def publish(self, text):
        self.path.write_text(self.header+text, encoding='utf8')
        self.store.sync([str(self.path)])

    def block(self, ref, sn, lot):
        return ''.join(self.event(t, ref, sn, lot) for t in (self.ticket, self.get, self.put))

    def test_seven_combinations_exact_matching_and_empty_guard(self):
        self.publish(''.join(self.block(*row) for row in (
            ('REF', 'SN1', 'LOT1'), ('REF', 'SN2', 'LOT1'), ('OTHER', 'SN1', 'LOT1'),
            ('OTHER', 'SN1', 'LOT2'), ('REF0', 'SN10', 'LOT10'))))
        for criteria, expected in ((dict(reference='REF'), 2), (dict(sn='SN1'), 3),
                (dict(lot='LOT1'), 3), (dict(reference='REF', sn='SN1'), 1),
                (dict(reference='REF', lot='LOT1'), 2), (dict(sn='SN1', lot='LOT1'), 2),
                (dict(reference='REF', sn='SN1', lot='LOT1'), 1)):
            with self.subTest(criteria=criteria):
                rows = self.store.tracking(**criteria)
                self.assertEqual(len(rows), expected)
                self.assertTrue(all(row['ref'] and row['sn'] and row['lot'] for row in rows))
        self.assertEqual(len(self.store.tracking(sn='  SN1  ')), 3)
        self.assertEqual(self.store.tracking(sn='SN'), [])
        self.assertEqual(self.store.tracking(lot='LOT'), [])
        with patch.object(self.store, '_rows', side_effect=AssertionError('Unbounded query')):
            self.assertEqual(self.store.tracking(), [])
            self.assertEqual(self.store.tracking(' ', sn=' ', lot=' '), [])

    def test_interleaved_references_consume_tickets_and_no_future_identity(self):
        self.publish(self.event(self.ticket, 'A', 'A_SN')+self.event(self.ticket, 'B', 'B_SN')+
                     self.event(self.get, 'A')+self.event(self.get, 'B')+
                     self.event(self.put, 'A')+self.event(self.put, 'B')+
                     self.event(self.get, 'A')+self.event(self.put, 'A')+
                     self.event(self.ticket, 'A', 'FUTURE'))
        self.assertEqual([r['ref'] for r in self.store.tracking(sn='A_SN')], ['A'])
        self.assertEqual([r['ref'] for r in self.store.tracking(sn='B_SN')], ['B'])
        self.assertFalse(self.store.tracking(sn='FUTURE'))
        rows = self.store.tracking('A')
        self.assertEqual(len(rows), 2)
        self.assertIsNone(rows[1]['sn'])

    def test_archived_generations_remain_separate_after_deletion(self):
        self.publish(self.block('A', 'SN1', 'LOT1'))
        self.publish(self.block('A', 'SN2', 'LOT1'))
        self.path.unlink()
        self.store.sync([])
        rows = self.store.tracking(lot='LOT1')
        self.assertEqual({r['sn'] for r in rows}, {'SN1', 'SN2'})
        self.assertEqual(len({r['path'].version_id for r in rows}), 2)
        self.assertTrue(all(not Path(r['path']).exists() for r in rows))
