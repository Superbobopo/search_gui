"""Regression checks for cached startup and atomic batched FTS ingestion."""

from contextlib import contextmanager
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import patch

from omega_history import OmegaHistory
from trace_store import TraceStore, Cancelled


class PerformanceTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.base = Path(temporary.name)
        self.root = self.base / 'traces'
        self.root.mkdir()
        self.database = self.base / 'index.db'
        self.archive = self.base / 'history.db'
        self.path = self.root / 'TracOmega.txt'
        self.text = (Path(__file__).parent / 'fixtures/current_r2.txt').read_bytes()
        self.path.write_bytes(self.text)
        self.store = self.open()

    def open(self, **kwargs):
        return TraceStore(self.database, directories=[self.root], history_database=self.archive, **kwargs)

    def test_unchanged_restart_does_not_decode_parse_compress_or_reindex(self):
        self.store.sync([str(self.path)])
        expected = self.store.search('ModBusId=1194', True)
        reopened = self.open(defer_history=True)
        reopened.prepare_history()
        statements = []
        connection = reopened.history.connection
        @contextmanager
        def observed():
            with connection() as db:
                db.set_trace_callback(statements.append)
                yield db
        with patch.object(reopened.history, 'connection', observed), \
                patch.object(reopened.history, 'publish', side_effect=AssertionError('reparsed')), \
                patch.object(reopened.history, 'snapshot', side_effect=AssertionError('decompressed')), \
                patch.object(reopened, '_copy_snapshot', side_effect=AssertionError('copied/decoded')), \
                patch.object(reopened, '_project_history', side_effect=AssertionError('reindexed')), \
                patch('omega_history.zlib.compress', side_effect=AssertionError('compressed')):
            result = reopened.sync([str(self.path)])
        self.assertEqual(result['changed'], 0)
        self.assertFalse(result['indexing'])
        self.assertEqual(reopened.search('ModBusId=1194', True), expected)
        writes = [sql for sql in statements if sql.lstrip().upper().startswith(('INSERT', 'UPDATE', 'DELETE'))]
        self.assertEqual(writes, [])

    def test_cached_search_is_available_before_legacy_preparation(self):
        self.store.sync([str(self.path)])
        with patch.object(OmegaHistory, 'migrate_caches', side_effect=RuntimeError('deferred work')):
            reopened = self.open(defer_history=True)
            self.assertTrue(reopened.search('ModBusId=1194', True))
            self.assertTrue(reopened.initialized())
            with self.assertRaisesRegex(RuntimeError, 'deferred work'):
                reopened.prepare_history()

    def test_partial_index_reopens_and_projects_only_unpublished_file(self):
        second = self.root / 'TracOmega001.old'
        second.write_bytes(self.text.replace(b'14/10/2026', b'13/10/2026').replace(b'14|', b'13|'))
        original = self.store._project_history
        calls = []
        def interrupted(*args):
            calls.append(args[0])
            if len(calls) == 2:
                raise RuntimeError('stop after first committed file')
            return original(*args)
        with patch.object(self.store, '_project_history', side_effect=interrupted):
            with self.assertRaisesRegex(RuntimeError, 'stop after'):
                self.store.sync([str(self.path), str(second)])
        reopened = self.open(defer_history=True)
        self.assertEqual(len(reopened.metadata()), 1)
        self.assertTrue(reopened.search('ModBusId=1194', True))
        published = reopened.metadata()[0]['path']
        with patch.object(reopened, '_project_history', wraps=reopened._project_history) as project:
            result = reopened.sync([str(self.path), str(second)])
        self.assertEqual(project.call_count, 1)
        self.assertNotEqual(project.call_args.args[0], published)
        self.assertFalse(result['failures'])
        self.assertTrue(reopened.initialized())
        self.assertEqual(len(reopened.metadata()), 2)

    def test_deferred_import_precedes_discard_of_incompatible_legacy_cache(self):
        legacy = TraceStore(self.database)
        legacy.sync([str(self.path)])
        with legacy.connection() as db, db:
            db.execute('PRAGMA user_version=999')
        self.path.unlink()
        reopened = self.open(defer_history=True)
        reopened.sync([])
        self.assertTrue(reopened.search('ModBusId=1194', True))

    def test_atomic_fts_batch_rolls_back_and_restores_trigger_after_cancel(self):
        self.store.sync([str(self.path)])
        ref = self.store.metadata()[0]['path']
        old = self.store.search('A0061574AC')
        with self.path.open('ab') as stream:
            stream.write(b'14|08:00:00 RefProduit=NEW_BATCH_CANCEL<Cr>ModBusId=555<Cr>\n')
        insert = self.store._insert_batch
        def cancel(db, batch, kind):
            insert(db, batch, kind)
            raise Cancelled()
        with patch.object(self.store, '_insert_batch', side_effect=cancel):
            with self.assertRaises(Cancelled):
                self.store.sync([str(self.path)])
        with self.store.connection() as db:
            self.assertEqual(db.execute('SELECT active FROM fts_ingest').fetchone()[0], 0)
        self.assertEqual(self.store.search('A0061574AC'), old)
        self.assertFalse(self.store.search('NEW_BATCH_CANCEL'))
        self.store.sync([str(self.path)])
        self.assertTrue(self.store.search('new_batch_cancel'))
        self.assertEqual(self.store.metadata()[0]['path'], ref)

    def test_additive_trigger_upgrade_preserves_cached_rows_and_fts(self):
        self.store.sync([str(self.path)])
        with self.store.connection() as db, db:
            before = [tuple(row) for row in db.execute('SELECT id,file_id,number,text FROM lines ORDER BY id')]
            db.execute('DROP TRIGGER line_insert')
            db.execute('DROP TABLE fts_ingest')
            db.execute('''CREATE TRIGGER line_insert AFTER INSERT ON lines BEGIN
                INSERT INTO lines_fts(rowid,normalized) VALUES(new.id,new.normalized); END''')
        reopened = self.open(defer_history=True)
        with reopened.connection() as db:
            self.assertEqual([tuple(row) for row in db.execute('SELECT id,file_id,number,text FROM lines ORDER BY id')], before)
        self.assertTrue(reopened.search('a0061574ac'))
        self.assertTrue(reopened.search('ModBusId=1194', True))

    def test_failure_after_fts_publication_rolls_back_lines_and_fts(self):
        self.store.sync([str(self.path)])
        before = self.store.search('A0061574AC')
        with self.path.open('ab') as stream:
            stream.write(b'14|08:00:00 RefProduit=AFTER_FTS_BATCH<Cr>ModBusId=555<Cr>\n')
        publish = self.store._publish_fts_batch
        def fail(db, file_id, first_number):
            publish(db, file_id, first_number)
            raise Cancelled()
        with patch.object(self.store, '_publish_fts_batch', side_effect=fail):
            with self.assertRaises(Cancelled):
                self.store.sync([str(self.path)])
        self.assertEqual(self.store.search('A0061574AC'), before)
        self.assertFalse(self.store.search('AFTER_FTS_BATCH'))
        with self.store.connection() as db:
            self.assertEqual(db.execute('SELECT active FROM fts_ingest').fetchone()[0], 0)
        self.store.sync([str(self.path)])
        self.assertTrue(self.store.search('after_fts_batch'))

    def test_failed_additive_upgrade_does_not_delete_existing_index(self):
        self.store.sync([str(self.path)])
        before = self.store.search('ModBusId=1194', True)
        connection = TraceStore.connection
        class FullDisk:
            def __init__(self, db):
                self.db = db
            def execute(self, *args):
                return self.db.execute(*args)
            def executescript(self, script):
                raise sqlite3.OperationalError('database or disk is full')
        @contextmanager
        def failing(store):
            with connection(store) as db:
                yield FullDisk(db)
        with patch.object(TraceStore, 'connection', failing):
            with self.assertRaisesRegex(sqlite3.OperationalError, 'full'):
                self.open(defer_history=True)
        self.assertTrue(self.database.exists())
        self.assertEqual(self.store.search('ModBusId=1194', True), before)

    def test_batched_search_matches_original_text_and_exact_fields_after_append(self):
        self.path.write_text('RefProduit=ABC123 SN=456 CodeLot=L1 ModBusId=1194\n'
                             'RefProduit=ABC1234 SN=4567 CodeLot=L12 ModBusId=11940\n'
                             'Unicode café ; a/b ; MixedCase\n', encoding='utf-8')
        self.store.sync([str(self.path)])
        for term in ('abc123', 'MixedCASE', 'café', 'a/b', 'xyz_absent'):
            result = self.store.search(term)
            expected = [(i, text) for i, text in enumerate(self.path.read_text(encoding='utf-8').splitlines()) if term.lower() in text.lower()]
            self.assertEqual([hit for ref, hits in result for hit in hits], expected)
        self.assertEqual(sum(len(hits) for ref, hits in self.store.search('ModBusId=1194', True)), 1)
        self.assertEqual(sum(len(hits) for ref, hits in self.store.search('', ref='ABC123', sn='456', lot='L1')), 1)
        with self.path.open('a', encoding='utf-8') as stream:
            stream.write('RefProduit=APPENDED_EXACT SN=789 CodeLot=L2 ModBusId=1195\n')
        self.store.sync([str(self.path)])
        self.assertTrue(self.store.search('appended_exact'))
        self.assertTrue(self.store.search('', ref='APPENDED_EXACT', sn='789', lot='L2'))


if __name__ == '__main__':
    unittest.main()
