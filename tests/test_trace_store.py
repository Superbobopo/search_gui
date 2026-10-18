from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import json
from pathlib import Path
import sqlite3
import struct
import tempfile
import threading
import time
import unittest
from unittest.mock import patch

from background_tasks import BackgroundTasks
from trace_parser import read_trace, missing_robot_tickets
from trace_store import TraceStore, Cancelled


FIXTURES = Path(__file__).parent / 'fixtures'


class StoreTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.database = self.root / 'cache' / 'index.db'
        self.store = TraceStore(self.database)

    def write(self, directory='r1', filename='TracOMEGA.txt', text=None):
        path = self.root / directory / filename
        path.parent.mkdir(parents=True, exist_ok=True)
        text = read_trace(FIXTURES / 'current_r2.txt') if text is None else text
        path.write_bytes(text.encode('cp1252'))
        return str(path)

    def test_reopen_does_not_read_unchanged_traces(self):
        path = self.write()
        self.store.sync([path])
        self.assertTrue(self.store.initialized())
        other = TraceStore(self.database)
        with patch('trace_store.detect_encoding', side_effect=AssertionError('Unexpected file read')):
            self.assertEqual(other.sync([path])['changed'], 0)
        self.assertEqual(other.search('ModBusId=1194', True), self.store.search('ModBusId=1194', True))

    def test_add_modify_delete_and_fts_invalidation(self):
        path = self.write(text='RefProduit=OLD<Cr>ModBusId=123<Cr>\n')
        self.store.sync([path]); self.store.file_lines(path)
        self.write(text='RefProduit=NEW_LONGER<Cr>ModBusId=456<Cr>\n')
        second = self.write('r2')
        self.store.sync([path, second])
        self.assertFalse(self.store.search('RefProduit=OLD'))
        self.assertTrue(self.store.search('NEW_LONGER'))
        self.assertIn('NEW_LONGER', self.store.file_lines(path)[0])
        Path(path).unlink(); self.store.sync([second])
        self.assertEqual(self.store.file_lines(path), [])
        self.assertFalse(self.store.search('NEW_LONGER'))
        self.assertEqual(len(self.store.metadata()), 1)

    def test_full_reference_prefix_search_and_exact_advanced_fields(self):
        path = self.write(text='RefProduit=ABC123<Cr>SN=456<Cr>CodeLot=L1<Cr>ModBusId=1194<Cr>\n'
                               'RefProduit=ABC1234<Cr>SN=4567<Cr>CodeLot=L12<Cr>ModBusId=11940<Cr>\n')
        self.store.sync([path])
        self.assertEqual(len(self.store.search('RefProduit=abc123')[0][1]), 2)
        self.assertEqual(len(self.store.search('ModBusId=1194', True)[0][1]), 1)
        self.assertEqual(len(self.store.search('', ref='ABC123', sn='456', lot='L1')[0][1]), 1)
        self.assertFalse(self.store.search('', ref='ABC123', sn='4567'))

    def test_text_search_punctuation_unicode_and_short_terms(self):
        text = 'Étagère [%_]+"Ω" / événement\n*other\n'
        path = self.write(text=text.replace('Ω', 'O'))
        self.store.sync([path])
        for term in ('étagère', '[%_]', '+"O"', 'événement', '*', '%_', 'ABSENT'):
            expected = [line for line in text.replace('Ω', 'O').splitlines() if term.lower() in line.lower()]
            groups = self.store.search(term)
            actual = [line for _, hits in groups for number, line in hits]
            self.assertEqual(actual, expected, term)

    def test_homonyms_rotation_and_line_positions(self):
        paths = [self.write(directory) for directory in ('r1', 'r2')]
        old = self.write('r1', 'TracOmega029.old', 'old\n')
        self.store.sync(paths + [old])
        self.assertEqual(len(self.store.search('ModBusId=1194', True)), 2)
        self.assertEqual(self.store.next_files(), {old: paths[0]})
        for path in paths:
            hit = self.store.search('ModBusId=1194', True, path=path)[0][1][0]
            self.assertEqual(self.store.file_lines(path)[hit[0]], hit[1])

    def test_details_measurements_locations_and_robot_errors_match_fixtures(self):
        path = self.write()
        self.store.sync([path])
        number = self.store.search('ModBusId=1194', True)[0][1][0][0]
        details = self.store.details(path, number)
        self.assertIn('GRIPPER GET', details['get']['text'])
        self.assertIn('GRIPPER PUTE', details['pute']['text'])
        self.assertEqual(details['locations'][0]['lot_hint'], '261013.127')
        measurement = self.store.measurement(path, 'A0061574AC', details['get']['number'])
        self.assertEqual(json.loads(measurement['dims']), {'l': 48, 'H': 21, 'L': 47})
        self.assertEqual(self.store.robot_errors(), [])
        self.assertEqual(missing_robot_tickets(read_trace(path).splitlines()), [])

    def test_missing_robot_does_not_use_reference_prefix_or_prior_get(self):
        path = self.write(text='GRIPPER GET 1x bote(s): Code:ABC $0101\n'
                               'RefProduit=ABC<Cr>ModBusId=12<Cr>\n'
                               'RefProduit=ABC123<Cr>ModBusId=13<Cr>\n'
                               'GRIPPER GET 1x bote(s): Code:ABC123 $0101\n')
        self.store.sync([path])
        self.assertEqual([(row['modbus'], row['ref']) for row in self.store.robot_errors()], [('12', 'ABC')])

    def test_modbus_steps_summary_recap_and_birth_rule(self):
        rows = []
        for values in ({1: '12', 2: '13'}, {6: '12'}, {6: '-14'}, {2: '14'}):
            row = ['0'] * 121; row[0] = '16|12:00:00'
            for col, value in values.items(): row[col] = value
            rows.append(';'.join(row))
        path = self.write(filename='TracMODBUS.txt', text='\n'.join(rows))
        self.store.sync([path])
        self.assertEqual(self.store.convoyage_errors(), [(path, '14')])
        self.assertEqual(self.store.convoyage_steps('12')[0]['etat'], 'Présent')
        self.assertEqual(self.store.convoyage_summary('14')[0]['_checkpoint_actuel'], 'Boite entrée')
        self.assertEqual(self.store.convoyage_recap('14')[6]['etat'], 'Jamais arrivé')

    def test_query_snapshot_remains_consistent_during_file_replacement(self):
        path = self.write(text='RefProduit=OLD\n')
        self.store.sync([path])
        with self.store.snapshot():
            self.assertTrue(self.store.search('OLD'))
            self.write(text='RefProduit=NEW\n')
            with ThreadPoolExecutor(1) as pool:
                pool.submit(self.store.sync, [path]).result(timeout=10)
            self.assertTrue(self.store.search('OLD'))
            self.assertFalse(self.store.search('NEW'))
        self.assertTrue(self.store.search('NEW'))

    def test_file_changing_during_read_does_not_publish_partial_index(self):
        path = self.write(text='RefProduit=OLD\n'); self.store.sync([path])
        self.write(text='RefProduit=CHANGED\n')
        original = self.store._index_snapshot
        def change(filename, *args):
            Path(filename).write_text('RefProduit=REWRITTEN_LONGER\n', encoding='utf8')
            return original(filename, *args)
        with patch.object(self.store, '_index_snapshot', side_effect=change):
            self.assertTrue(self.store.sync([path])['failures'])
        self.assertTrue(self.store.search('OLD'))
        self.assertFalse(self.store.search('CHANGED'))

    def test_live_append_publishes_stable_prefix_then_indexes_the_new_tail(self):
        path = self.write(text='RefProduit=FIRST\n')
        original = self.store._index_snapshot
        def append(filename, *args):
            with open(filename, 'a') as stream: stream.write('RefProduit=SECOND\n')
            return original(filename, *args)
        with patch.object(self.store, '_index_snapshot', side_effect=append):
            self.assertFalse(self.store.sync([path])['failures'])
        self.assertTrue(self.store.search('FIRST'))
        self.assertFalse(self.store.search('SECOND'))
        self.store.sync([path])
        self.assertTrue(self.store.search('SECOND'))

    def test_corrupt_and_incompatible_caches_rebuild(self):
        for contents in ('broken', 'version'):
            database = self.root / (contents + '.db')
            if contents == 'broken': database.write_bytes(b'not a sqlite database')
            else:
                db = sqlite3.connect(database)
                db.execute('PRAGMA user_version=999'); db.close()
            self.assertEqual(TraceStore(database).metadata(), [])

    def test_text_cache_is_limited_to_three_and_only_populated_on_open(self):
        paths = [self.write(str(index)) for index in range(4)]
        self.store.sync(paths)
        self.store.search('ModBusId=1194', True)
        self.store.details(paths[0], 3)
        self.assertFalse(self.store.text_cache)
        for path in paths: self.store.file_lines(path)
        self.assertEqual(len(self.store.text_cache), 3)
        self.assertNotIn(paths[0], self.store.text_cache)

    def test_corruption_detected_during_query_can_be_rebuilt_without_restart(self):
        path = self.write(); self.store.sync([path])
        with self.store.connection() as db:
            page_size = db.execute('PRAGMA page_size').fetchone()[0]
            root = db.execute("SELECT rootpage FROM sqlite_master WHERE name='files'").fetchone()[0]
        with open(self.database, 'r+b') as stream:
            stream.seek((root - 1) * page_size); stream.write(b'\xff')
        with self.assertRaises(sqlite3.DatabaseError): self.store.metadata()
        self.assertTrue(self.store.corrupt.is_set())
        self.store.rebuild(); self.store.sync([path])
        self.assertFalse(self.store.corrupt.is_set())
        self.assertTrue(self.store.search('ModBusId=1194', True))

    def test_cancelled_search_is_not_executed(self):
        cancel = threading.Event(); cancel.set()
        with self.assertRaises(Cancelled): self.store.search('absent', cancel=cancel)

    def test_adjacent_anchors_reused_id_to_selected_line_and_reference(self):
        text = ''.join(f'RefProduit={ref}<Cr>ModBusId={mb}<Cr>\n'
                       for mb, ref in [('12', 'OLD'), ('13', 'BEFORE'), ('12', 'NEW'), ('14', 'AFTER')])
        path = self.write(text=text)
        self.store.sync([path])
        before, after, ref, _, number = self.store.adjacent(path, '12', 2, 'NEW')
        self.assertEqual((ref, number), ('NEW', 2))
        self.assertEqual(before, [('13', 'BEFORE', path, 1)])
        self.assertEqual(after, [('14', 'AFTER', path, 3)])
        self.assertIsNone(self.store.adjacent(path, '12', 2, 'OLD'))
        self.assertIsNone(self.store.adjacent(path, '12', 1))

    def test_adjacent_completes_five_distinct_ids_across_same_robot_archive(self):
        def tickets(ids):
            return ''.join(f'RefProduit=R{mb}<Cr>ModBusId={mb}<Cr>\n' for mb in ids)
        old = self.write(filename='TracOmega029.old', text=tickets([10, 10, 11, 11]))
        current = self.write(text='Ancienne trace = C:\\apoteka\\Traces\\TracOmega029.old\n' + tickets([11, 11, 12, 13, 14, 15, 16]))
        other_robot = self.write('r2', text='Ancienne trace = TracOmega029.old\n' + tickets([999]))
        self.store.sync([old, current, other_robot])
        _, following, _, _, _ = self.store.adjacent(old, '10', 0)
        self.assertEqual([item[0] for item in following], ['11', '12', '13', '14', '15'])
        self.assertTrue(all(item[2] in (old, current) for item in following))
        self.assertEqual(following[0][3], 2)

    def test_adjacent_keeps_five_distinct_preceding_ids_despite_duplicates(self):
        path = self.write(text=''.join(f'RefProduit=R{i}<Cr>ModBusId={i}<Cr>\n' for i in [1, 1, 2, 3, 4, 5, 5, 6]))
        self.store.sync([path])
        before, after, _, _, _ = self.store.adjacent(path, '6', 7)
        self.assertEqual([item[0] for item in before], ['5', '4', '3', '2', '1'])
        self.assertEqual(after, [])

    def test_details_get_and_pute_follow_selected_reused_occurrence(self):
        def occurrence(marker):
            return ('RefProduit=ABC<Cr>ModBusId=12<Cr>\n'
                    f'GRIPPER GET 1x bote(s): Code:ABC $0101 {marker}\n'
                    f'GRIPPER PUTE 1x bote(s): Code:ABC {marker}\n')
        path = self.write(text=occurrence('FIRST') + occurrence('SECOND'))
        self.store.sync([path])
        details = self.store.details(path, 3)
        self.assertEqual(details['get']['number'], 4)
        self.assertEqual(details['pute']['number'], 5)
        self.assertIn('SECOND', details['get']['text'])

    def test_file_view_reads_modified_or_deleted_trace_before_index_refresh(self):
        path = self.write(text='original\ntarget\n')
        self.store.sync([path])
        self.assertEqual(self.store.file_lines(path), ['original', 'target'])
        self.write(text='inserted\noriginal\ntarget\n')
        self.assertEqual(self.store.file_lines(path), ['inserted', 'original', 'target'])
        Path(path).unlink()
        self.assertEqual(self.store.file_lines(path), [])


class BackgroundTests(unittest.TestCase):
    def test_latest_query_wins_and_callbacks_stay_on_calling_thread(self):
        runner = BackgroundTasks(); self.addCleanup(runner.close)
        waiting = threading.Event(); finished = threading.Event(); output = []
        thread = threading.get_ident()
        def slow(cancel):
            waiting.wait(3); finished.set(); return 'old'
        runner.submit('search', slow, lambda value: output.append(value), self.fail)
        runner.submit('search', lambda cancel: 'new', lambda value: output.append((value, threading.get_ident())), self.fail)
        waiting.set(); finished.wait(3)
        deadline = time.monotonic() + 3
        while not output and time.monotonic() < deadline:
            runner.drain(); time.sleep(.01)
        self.assertEqual(output, [('new', thread)])

    def test_icon_has_all_seven_sizes_and_png_frames(self):
        data = (Path(__file__).parents[1] / 'assets' / 'tracomega.ico').read_bytes()
        self.assertEqual(struct.unpack_from('<HHH', data), (0, 1, 7))
        sizes = []
        for index in range(7):
            width, height, _, _, _, _, length, offset = struct.unpack_from('<BBBBHHII', data, 6 + index * 16)
            sizes.append(width or 256)
            self.assertEqual(width, height)
            self.assertEqual(data[offset:offset + 8], b'\x89PNG\r\n\x1a\n')
            self.assertGreater(length, 100)
        self.assertEqual(sizes, [16, 24, 32, 48, 64, 128, 256])
