"""Small, isolated regressions for permanent OMEGA history."""

import os
from contextlib import closing
from pathlib import Path
import sqlite3
import tempfile
import threading
import unittest
from unittest.mock import patch

from omega_history import OmegaHistory, historical_roots
from trace_identity import TraceRef
from trace_parser import read_trace
from trace_sources import available_trace_sources, path_key, trace_label
from trace_store import TraceStore, Cancelled


class HistoryTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.base = Path(temporary.name)
        self.root = self.base / 'traces'
        self.root.mkdir()
        self.database = self.base / 'cache' / 'index.db'
        self.archive = self.base / 'history' / 'omega_history.sqlite3'
        self.store = self.open()

    def open(self, roots=None, database=None):
        return TraceStore(database or self.database, directories=roots or [self.root], history_database=self.archive)

    def write(self, marker='OCTOBRE', name='TracOmega001.old', robot='R1', month=10, encoding='cp1252'):
        path = self.root / robot / name
        path.parent.mkdir(parents=True, exist_ok=True)
        text = (f'Ancienne trace = C:\\apoteka\\Traces\\TracOmega020.old\n'
                f'creation de la trace à 00:00:01 Le 14/{month:02}/2026\n\n'
                f'14|06:27:47 RefProduit={marker}<Cr>SN=343272<Cr>CodeLot=L1<Cr>ModBusId=1194<Cr>\n'
                f'G 14|06:27:52 GRIPPER GET 1x boîte(s): Code:{marker} $0101\n'
                f'14|06:27:59 IDPK pour {marker} : l=48 H=21 L=47\n'
                f'G 14|06:28:00 GRIPPER PUTE 1x boîte(s): Code:{marker}\n'
                '14|06:28:01 café\n14|06:28:02 fin\n')
        path.write_bytes(text.encode(encoding))
        return str(path)

    def refs(self, marker):
        return [ref for ref, hits in self.store.search('RefProduit=' + marker)]

    def sync(self, paths):
        result = self.store.sync(paths)
        self.assertFalse(result['failures'], result)
        return result

    def test_two_numbered_cycles_deletion_restart_and_rebuild(self):
        paths = [self.write('OCT' + str(i), f'TracOmega{i:03}.old') for i in (1, 2)]
        self.sync(paths)
        old = self.refs('OCT1')[0]
        text = self.store.file_lines(old)
        for i in (1, 2):
            self.write('NOV' + str(i), f'TracOmega{i:03}.old', month=11)
        self.sync(paths)
        self.assertNotEqual(old, self.refs('NOV1')[0])
        self.assertEqual(old.path, self.refs('NOV1')[0].path)
        self.assertTrue(self.refs('OCT1')[0].historical)
        self.assertEqual(self.store.file_lines(old), text)
        for path in paths:
            Path(path).unlink()
        self.sync([])
        self.store = self.open()
        self.store.rebuild()
        self.sync([])
        self.assertEqual(len(self.store.metadata()), 4)
        self.assertEqual(self.store.file_lines(old), text)
        self.assertIn('Historique', trace_label(self.refs('OCT1')[0]))

    def test_identical_copy_and_rename_do_not_duplicate(self):
        path = self.write()
        self.sync([path])
        ref = self.refs('OCTOBRE')[0]
        renamed = str(Path(path).with_name('TracOmega020.old'))
        Path(renamed).write_bytes(Path(path).read_bytes())
        self.sync([path, renamed])
        self.assertEqual(self.refs('OCTOBRE'), [ref])
        Path(path).unlink()
        self.sync([renamed])
        self.assertFalse(self.refs('OCTOBRE')[0].historical)
        self.assertEqual(len(self.store.history.entries([self.root])), 1)

    def test_append_extends_generation_and_keeps_incremental_indexing(self):
        path = self.write()
        self.sync([path])
        ref = self.refs('OCTOBRE')[0]
        with open(path, 'ab') as stream:
            stream.write(b'14|06:29:00 RefProduit=AJOUT<Cr>ModBusId=1195<Cr>\n')
        result = self.sync([path])
        self.assertEqual(self.refs('AJOUT'), [ref])
        self.assertEqual(len(self.store.metadata()), 1)
        self.assertEqual(result['indexing'][0]['mode'], 'append')
        self.assertLess(result['indexing'][0]['lines_parsed'], len(self.store.file_lines(ref)))

    def test_truncation_preserves_complete_generation(self):
        path = self.write()
        self.sync([path])
        ref = self.refs('OCTOBRE')[0]
        old = self.store.file_lines(ref)
        Path(path).write_bytes(b'RefProduit=TRONQUE<Cr>ModBusId=7<Cr>\n')
        self.sync([path])
        self.assertEqual(self.store.file_lines(ref), old)
        self.assertNotEqual(self.refs('TRONQUE')[0], ref)

    def test_same_size_and_preserved_mtime_are_checked_on_restart(self):
        path = self.write('OCTOBRE')
        self.sync([path])
        stat = os.stat(path)
        self.write('NOVEMBR')
        os.utime(path, ns=(stat.st_atime_ns, stat.st_mtime_ns))
        self.assertEqual(os.stat(path).st_size, stat.st_size)
        self.store = self.open()
        self.sync([path])
        self.assertTrue(self.refs('OCTOBRE'))
        self.assertTrue(self.refs('NOVEMBR'))

    def test_robots_and_selected_sources_are_isolated(self):
        one = self.write(robot='R1')
        two = self.write(robot='R2')
        self.sync([one, two])
        self.assertEqual(len(self.refs('OCTOBRE')), 2)
        first = self.open([self.root / 'R1'], self.base / 'r1.db')
        first.sync([])
        self.assertEqual(len(first.metadata()), 1)
        self.assertEqual(first.metadata()[0]['path'].path, one)
        other = self.base / 'other'
        other.mkdir()
        none = self.open([other], self.base / 'other.db')
        none.sync([])
        self.assertFalse(none.search('OCTOBRE'))

    def test_modbus_still_replaces_and_deletes(self):
        path = self.root / 'TracMODBUS.txt'
        path.write_text('14|06:00;OLD_MODBUS;1;\n')
        self.sync([str(path)])
        path.write_text('14|06:00;NEW_MODBUS;2;\n')
        self.sync([str(path)])
        self.assertFalse(self.store.search('OLD_MODBUS'))
        self.assertTrue(self.store.search('NEW_MODBUS'))
        path.unlink()
        self.sync([])
        self.assertFalse(self.store.search('NEW_MODBUS'))
        self.assertFalse(self.store.history.entries([self.root]))

    def test_archive_survives_failure_between_capture_and_index(self):
        path = self.write()
        with patch.object(self.store, '_project_history', side_effect=RuntimeError('interruption')):
            with self.assertRaisesRegex(RuntimeError, 'interruption'):
                self.store.sync([path])
        Path(path).unlink()
        self.store = self.open()
        self.sync([])
        self.assertTrue(self.refs('OCTOBRE'))

    def test_legacy_missing_files_migrate_without_touching_cache(self):
        path = self.write()
        legacy_file = self.base / 'legacy.db'
        legacy = TraceStore(legacy_file)
        legacy.sync([path])
        original = legacy_file.read_bytes()
        Path(path).unlink()
        with patch.object(Path, 'glob', return_value=iter([legacy_file])):
            self.store = self.open()
        self.sync([])
        self.assertTrue(self.refs('OCTOBRE'))
        self.assertIn('café', '\n'.join(self.store.file_lines(self.refs('OCTOBRE')[0])))
        self.assertEqual(legacy_file.read_bytes(), original)
        # A source copy with CRLF and its original cp1252 encoding deduplicates
        # against the normalized UTF8 representation rescued from the old index.
        self.write()
        Path(path).write_bytes(Path(path).read_bytes().replace(b'\n', b'\r\n'))
        self.sync([path])
        self.assertEqual(len(self.refs('OCTOBRE')), 1)

    def test_incompatible_history_is_never_deleted(self):
        path = self.write()
        self.sync([path])
        with closing(sqlite3.connect(self.archive)) as db:
            db.execute('PRAGMA user_version=999')
            db.commit()
        before = self.archive.read_bytes()
        with self.assertRaisesRegex(RuntimeError, 'incompatible'):
            self.open()
        self.assertEqual(self.archive.read_bytes(), before)

    def test_corrupt_history_is_never_deleted(self):
        database = self.base / 'corrupt.sqlite3'
        database.write_bytes(b'precious unreadable history')
        with self.assertRaises(sqlite3.DatabaseError):
            OmegaHistory(database)
        self.assertEqual(database.read_bytes(), b'precious unreadable history')

    def test_history_only_source_remains_available(self):
        path = self.write()
        self.sync([path])
        Path(path).unlink()
        self.sync([])
        self.assertEqual(historical_roots(self.archive), {path_key(self.root)})
        with patch('trace_sources.trace_directories', return_value=(self.root, self.base/'absent', self.base/'absent2')), \
                patch('omega_history.history_path', return_value=self.archive):
            self.assertEqual(available_trace_sources(), (('Apoteka (historique)', str(self.root)),))

    def test_archived_get_pute_and_original_positions(self):
        path = self.write()
        self.sync([path])
        ref, hits = self.store.search('ModBusId=1194', True)[0]
        details = self.store.details(ref, hits[0][0])
        self.write('NOVEMBR', month=11)
        self.sync([path])
        for key in ('get', 'pute', 'idpk'):
            event = details[key]
            self.assertEqual(self.store.file_lines(ref)[event['number']], event['text'])
        self.assertTrue(self.store.tracking('OCTOBRE'))
        self.assertIsInstance(ref, TraceRef)

    def test_reimport_of_an_older_observed_prefix_does_not_duplicate(self):
        path = self.write()
        original = Path(path).read_bytes()
        self.sync([path])
        ref = self.refs('OCTOBRE')[0]
        with open(path, 'ab') as stream:
            stream.write(b'14|06:29:00 RefProduit=AJOUT<Cr>\n')
        self.sync([path])
        Path(path).write_bytes(original)
        self.sync([path])
        self.assertEqual(self.refs('OCTOBRE'), [ref])
        self.assertEqual(self.refs('AJOUT'), [ref])
        self.assertEqual(len(self.store.metadata()), 1)
        self.assertTrue(self.refs('OCTOBRE')[0].historical)

    def test_unchanged_files_are_not_recaptured_within_a_session(self):
        path = self.write()
        self.sync([path])
        with patch.object(self.store, '_copy_snapshot', side_effect=AssertionError('unnecessary read')):
            self.assertEqual(self.sync([path])['changed'], 0)

    def test_archive_write_failure_does_not_replace_old_results(self):
        path = self.write()
        self.sync([path])
        ref = self.refs('OCTOBRE')[0]
        original = self.store.file_lines(ref)
        self.write('NOVEMBR', month=11)
        with patch.object(self.store.history, 'publish', side_effect=sqlite3.OperationalError('database full')):
            with self.assertRaisesRegex(sqlite3.OperationalError, 'full'):
                self.store.sync([path])
        self.assertEqual(self.store.file_lines(ref), original)
        self.assertFalse(self.refs('NOVEMBR'))
        self.assertEqual(len(self.store.history.entries([self.root])), 1)

    def test_backup_is_a_complete_independent_sqlite(self):
        path = self.write()
        self.sync([path])
        backup = self.store.history.backup_before_migration()
        other = OmegaHistory(backup)
        self.assertEqual(other.entries([self.root]), self.store.history.entries([self.root]))
        ref = self.refs('OCTOBRE')[0]
        with other.snapshot(ref.version_id) as stream:
            self.assertEqual(stream.read(), Path(path).read_bytes())

    def test_cancelled_sync_keeps_last_committed_history(self):
        path = self.write()
        self.sync([path])
        before = self.store.history.entries([self.root])
        self.write('NOVEMBR', month=11)
        stop = threading.Event()
        stop.set()
        with self.assertRaises(Cancelled):
            self.store.sync([path], stop=stop)
        self.assertEqual(self.store.history.entries([self.root]), before)

    def test_history_cannot_be_opened_as_a_disposable_cache(self):
        path = self.write()
        self.sync([path])
        before = self.archive.read_bytes()
        with self.assertRaisesRegex(RuntimeError, 'permanent'):
            TraceStore(self.archive)
        self.assertEqual(before, self.archive.read_bytes())

    def test_snapshot_integrity_is_checked_before_rebuilding_index(self):
        path = self.write()
        self.sync([path])
        ref = self.refs('OCTOBRE')[0]
        with self.store.history.connection() as db, db:
            db.execute('DELETE FROM chunks WHERE version=?', (ref.version_id,))
        with self.assertRaisesRegex(OSError, 'invalide'):
            with self.store.history.snapshot(ref.version_id):
                pass
        self.assertTrue(self.refs('OCTOBRE'))

    def test_archive_links_follow_each_cycle_and_refuse_ambiguous_generations(self):
        previous = self.write('PREVIOUS', 'TracOmega020.old')
        current = self.write('CURRENT', 'TracOMEGA.txt')
        def previous_day():
            text = Path(previous).read_bytes().replace(b'14/', b'13/').replace(b'14|', b'13|')
            Path(previous).write_bytes(text)
        previous_day()
        self.sync([previous, current])
        old_previous, old_current = self.refs('PREVIOUS')[0], self.refs('CURRENT')[0]
        self.write('PREVIOUS_NEW', 'TracOmega020.old', month=11)
        self.write('CURRENT_NEW', 'TracOMEGA.txt', month=11)
        previous_day()
        self.sync([previous, current])
        new_previous, new_current = self.refs('PREVIOUS_NEW')[0], self.refs('CURRENT_NEW')[0]
        links = self.store.next_files()
        self.assertEqual(links[old_previous], old_current)
        self.assertEqual(links[new_previous], new_current)
        self.write('AMBIGUOUS', 'TracOmega020.old', month=11)
        previous_day()
        self.sync([previous, current])
        self.assertNotIn(new_previous, self.store.next_files())

    def test_migrated_normalized_trace_can_extend_without_duplicate(self):
        path = self.write()
        legacy_file = self.base / 'trace_index_old.sqlite3'
        legacy = TraceStore(legacy_file)
        legacy.sync([path])
        self.store = self.open()
        self.sync([path])
        ref = self.refs('OCTOBRE')[0]
        with open(path, 'ab') as stream:
            stream.write(b'14|06:29:00 RefProduit=AFTER_MIGRATION<Cr>\n')
        self.sync([path])
        self.assertEqual(self.refs('AFTER_MIGRATION'), [ref])
        self.assertEqual(len(self.store.metadata()), 1)
        self.assertEqual(self.store.history.entries([self.root])[0]['source_encoding'], 'cp1252')

    def test_identical_text_recopied_as_utf8_can_then_extend(self):
        path = self.write()
        self.sync([path])
        ref = self.refs('OCTOBRE')[0]
        text = read_trace(path)
        Path(path).write_bytes(text.replace('\n', '\r\n').encode('utf-8-sig'))
        self.sync([path])
        with open(path, 'ab') as stream:
            stream.write(b'14|06:29:00 RefProduit=UTF8_APPEND<Cr>\n')
        self.sync([path])
        self.assertEqual(self.refs('UTF8_APPEND'), [ref])
        self.assertEqual(len(self.store.metadata()), 1)


if __name__ == '__main__':
    unittest.main()
