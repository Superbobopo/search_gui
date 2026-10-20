from pathlib import Path
import tempfile
import unittest

from trace_parser import read_trace
from trace_store import TraceStore


REFERENCE = 'A0015126AB'
LOT = '260819.198'


def trace_block(lot, hour=6, modbus='1194'):
    original = read_trace(Path(__file__).parent / 'fixtures/current_r2.txt')
    return (original.replace('A0061574AC', REFERENCE).replace('261013.127', lot)
            .replace('06:', f'{hour:02}:').replace('1194', modbus))


class LotLocationsTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.base = Path(temporary.name)
        self.root = self.base / 'traces'
        self.root.mkdir()
        self.database = self.base / 'index.sqlite3'
        self.archive = self.base / 'history.sqlite3'
        self.store = self.open()
        self.paths = []
        # Target lot has an older operation in the very same file: filtering
        # only the unfiltered "last PUTE per file" result would lose it.
        for robot, text in (('R1', trace_block(LOT) + trace_block(LOT + '0', 7, '1195')),
                            ('R2', trace_block(LOT, 8, '1196') + trace_block('OTHER', 9, '1197'))):
            path = self.root / robot / 'TracOmega019.old'
            path.parent.mkdir()
            path.write_text(text, encoding='utf8')
            self.paths.append(str(path))
        self.assertFalse(self.store.sync(self.paths)['failures'])
        self.ref, hits = self.store.search('', ref=REFERENCE, lot=LOT)[0]
        self.number = hits[0][0]

    def open(self):
        return TraceStore(self.database, directories=[self.root], history_database=self.archive)

    def test_exact_lot_is_filtered_before_latest_event_selection(self):
        result = self.store.details(self.ref, self.number, location_lot=LOT)
        self.assertEqual(result['location_lot'], LOT)
        self.assertEqual([event['lot_hint'] for event in result['locations']], [LOT])
        self.assertIn('08:28:07.6', result['locations'][0]['text'])
        self.assertEqual(result['row']['lot'], LOT)
        other = self.store.details(self.ref, self.number, location_lot=LOT + '0')
        self.assertEqual([event['lot_hint'] for event in other['locations']], [LOT + '0'])
        self.assertFalse(self.store.details(self.ref, self.number, location_lot='ABSENT')['locations'])

    def test_without_filter_keeps_existing_last_operation_per_file_rule(self):
        result = self.store.details(self.ref, self.number)
        self.assertIsNone(result['location_lot'])
        self.assertEqual({event['lot_hint'] for event in result['locations']}, {LOT + '0', 'OTHER'})

    def test_lot_filter_survives_source_deletion_and_restart_without_reindex(self):
        for path in self.paths:
            Path(path).unlink()
        self.assertFalse(self.store.sync([])['failures'])
        reopened = self.open()
        ref, hits = reopened.search('', ref=REFERENCE, lot=LOT)[0]
        result = reopened.details(ref, hits[0][0], location_lot=LOT)
        self.assertEqual([event['lot_hint'] for event in result['locations']], [LOT])
        self.assertTrue(result['locations'][0]['path'].historical)
