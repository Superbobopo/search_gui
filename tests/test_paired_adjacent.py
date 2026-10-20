from datetime import datetime, timedelta
from pathlib import Path
import tempfile
import threading
import unittest

from adjacent_view import gap_text
from trace_sources import robot_scope
from trace_store import Cancelled, TraceStore


DAY = datetime(2026, 10, 14)


def ticket(seconds, modbus, reference, lot='LOT', sn='123456'):
    def stamp(offset):
        return (DAY + timedelta(seconds=seconds + offset)).strftime('%d|%H:%M:%S.0')
    return (f'  {stamp(0)} RefProduit={reference}<Cr>SN={sn}<Cr>CodeLot={lot}<Cr>ModBusId={modbus}<Cr>\n'
            f'G {stamp(1)} GRIPPER GET 1x boîte(s): Code:{reference} $0101\n'
            f'G {stamp(2)} GRIPPER PUTE 1x boîte(s): Code:{reference} face:1 601\n')


def trace(*tickets, previous=''):
    return (f'Ancienne trace = {previous}\n'
            'creation de la trace de ... à 00:00:01 Le 14/10/2026\n' + ''.join(tickets))


class PairedAdjacentTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.base = Path(temporary.name)
        self.root = self.base / 'Bureau' / 'traces'
        self.other = self.base / 'Local' / 'traces'
        self.store = TraceStore(self.base / 'index.db', directories=[self.root, self.other])
        self.paths = []

    def write(self, robot, text, name='TracOMEGA.txt', root=None):
        path = (root or self.root) / robot / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding='utf8')
        if str(path) not in self.paths:
            self.paths.append(str(path))
        return str(path)

    def sync(self):
        self.assertFalse(self.store.sync(self.paths)['failures'])

    def select(self, reference):
        ref, hits = self.store.search('', ref=reference)[0]
        return ref, hits[0][0]

    def paired(self, reference, target='10'):
        ref, number = self.select(reference)
        return self.store.paired_adjacent(ref, target, number, reference)

    def test_time_anchor_independent_ids_before_tie_and_source_isolation(self):
        self.write('trace r1', trace(ticket(36000, '10', 'R1_SELECTED')))
        self.write('trace r2', trace(ticket(35990, '10', 'R2_OTHER_PRODUCT'),
                                     ticket(36010, '99', 'R2_AFTER')))
        self.write('R2', trace(ticket(36000, '10', 'WRONG_SOURCE')), root=self.other)
        self.sync()
        data = self.paired('R1_SELECTED')
        first, second = data['columns']
        self.assertEqual((first['robot'], second['robot']), ('R1', 'R2'))
        self.assertEqual(first['anchor']['ref'], 'R1_SELECTED')
        self.assertEqual(second['anchor']['ref'], 'R2_OTHER_PRODUCT')
        self.assertEqual(second['delta'], -10)
        self.assertEqual([r['modbus'] for r in second['following']], ['99'])
        self.assertEqual(robot_scope(second['anchor']['path']), (robot_scope(first['anchor']['path'])[0], 'R2'))
        # A nearer different ID beats a matching ID on the other robot.
        self.write('trace r2', trace(ticket(35980, '10', 'R2_SAME_ID'), ticket(36001, '77', 'R2_NEAREST')))
        self.sync()
        self.assertEqual(self.paired('R1_SELECTED')['columns'][1]['anchor']['modbus'], '77')

    def test_exact_reused_occurrence_and_five_distinct_following_across_archive(self):
        self.write('R1', trace(ticket(35900, '10', 'OLD'), ticket(36000, '10', 'SELECTED'),
                                ticket(36010, '11', 'NEXT'), ticket(36020, '11', 'DUPLICATE')),
                   'TracOmega019.old')
        self.write('R1', trace(*(ticket(36100 + i * 10, str(i), 'NEXT_' + str(i)) for i in range(11, 17)),
                               previous='TracOmega019.old'))
        self.sync()
        own, missing = self.paired('SELECTED')['columns']
        self.assertEqual(own['anchor']['ref'], 'SELECTED')
        self.assertEqual(own['anchor']['number'], self.select('SELECTED')[1])
        self.assertEqual([r['modbus'] for r in own['following']], ['11', '12', '13', '14', '15'])
        self.assertIn('Aucune trace', missing['message'])

    def test_far_robot_still_visible_and_r2_selection_orders_columns(self):
        self.write('R1', trace(ticket(36000, '10', 'ONE')))
        self.write('R2', trace(ticket(40000, '20', 'TWO')))
        self.sync()
        data = self.paired('TWO', '20')
        self.assertEqual(data['selected_robot'], 'R2')
        self.assertEqual([c['robot'] for c in data['columns']], ['R1', 'R2'])
        self.assertEqual(data['columns'][0]['delta'], -4000)
        self.assertEqual(data['columns'][1]['delta'], 0)
        self.assertIn('avant', gap_text(-4000))
        self.assertEqual(gap_text(0), 'même instant')
        self.assertIn('0.2 s après', gap_text(.2))

    def test_ambiguous_generations_offered_and_deleted_sources_stay_available(self):
        first = self.write('R1', trace(ticket(36000, '10', 'ONE')))
        second = self.write('R2', trace(ticket(36002, '20', 'OLD_R2')))
        self.sync()
        self.write('R2', trace(ticket(36002, '21', 'NEW_R2')))
        self.sync()
        data = self.paired('ONE')
        other = data['columns'][1]
        self.assertIsNone(other['anchor'])
        self.assertEqual({row['ref'] for row in other['candidates']}, {'OLD_R2', 'NEW_R2'})
        self.assertEqual(len({r['path'].version_id for r in other['candidates']}), 2)
        for path in (first, second):
            Path(path).unlink()
        self.store.sync([])
        self.store = TraceStore(self.base / 'index.db', directories=[self.root, self.other])
        data = self.paired('ONE')
        self.assertTrue(all(row['path'].historical for row in data['columns'][1]['candidates']))
        self.assertTrue(data['selected']['path'].historical)

    def test_missing_date_or_invalid_time_does_not_use_copy_date(self):
        self.write('R1', trace(ticket(36000, '10', 'ONE')).replace('10:00:00.0', '99:00:00.0'))
        self.write('R2', trace(ticket(36000, '20', 'TWO')))
        self.sync()
        data = self.paired('ONE')
        self.assertIsNone(data['time'])
        self.assertIn('inexploitable', data['columns'][1]['message'])

    def test_headerless_other_robot_and_unknown_selected_robot(self):
        self.write('R1', trace(ticket(36000, '10', 'ONE')))
        self.write('R2', ticket(36000, '20', 'NO_DATE'))
        self.write('unidentified', trace(ticket(36000, '30', 'UNKNOWN')))
        self.sync()
        self.assertIn('date exploitable', self.paired('ONE')['columns'][1]['message'])
        unknown = self.paired('UNKNOWN', '30')
        self.assertIsNone(unknown['selected_robot'])
        self.assertEqual(unknown['columns'][0]['anchor']['ref'], 'UNKNOWN')
        self.assertIn('non reconnu', unknown['columns'][1]['message'])

    def test_cancelled_request_and_wrong_occurrence_return_no_neighbors(self):
        self.write('R1', trace(ticket(36000, '10', 'ONE')))
        self.sync()
        ref, number = self.select('ONE')
        self.assertIsNone(self.store.paired_adjacent(ref, '11', number))
        self.assertIsNone(self.store.paired_adjacent(ref, '10', number, 'WRONG'))
        cancel = threading.Event()
        cancel.set()
        with self.assertRaises(Cancelled):
            self.store.paired_adjacent(ref, '10', number, cancel=cancel)

    def test_fractional_seconds_tie_prefers_earlier_ticket(self):
        self.write('R1', trace(ticket(36000, '10', 'ONE')).replace('10:00:00.0', '10:00:00.2'))
        self.write('R2', trace(ticket(36000, '20', 'BEFORE')).replace('10:00:00.0', '10:00:00.1')
                   + ticket(36000, '21', 'AFTER').replace('10:00:00.0', '10:00:00.3'))
        self.sync()
        self.assertEqual(self.paired('ONE')['columns'][1]['anchor']['ref'], 'BEFORE')

    def test_ambiguous_generation_survives_more_than_one_batch_of_repeated_tickets(self):
        self.write('R1', trace(ticket(36000, '10', 'ONE')))
        repeated = ticket(36000, '20', 'REPEATED') * 40
        self.write('R2', trace(repeated))
        self.sync()
        self.write('R2', trace(ticket(36000, '21', 'ANOTHER_GENERATION')))
        self.sync()
        other = self.paired('ONE')['columns'][1]
        self.assertIsNone(other['anchor'])
        self.assertEqual({row['ref'] for row in other['candidates']}, {'REPEATED', 'ANOTHER_GENERATION'})


if __name__ == '__main__':
    unittest.main()
