"""Append regressions use only small temporary traces and disposable databases."""

from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import json
import os
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import patch

import trace_store
from trace_parser import read_trace
from trace_store import Cancelled, TraceStore


class IncrementalStoreTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.path = self.root / 'r1' / 'TracOMEGA.txt'
        self.path.parent.mkdir()
        self.database = self.root / 'index.db'
        self.store = TraceStore(self.database)
        self.full_counter = 0

    def prefix(self, newline='\n'):
        return newline.join(['Ancienne trace = TracOmega029.old',
                             'creation de la trace Le 31/12/2025'] +
                            [f'D 31|23:59:58 diagnostic {i} accentué' for i in range(8)]) + newline

    def write(self, text, encoding='utf-8'):
        self.path.write_bytes(text.encode(encoding))
        return self.sync()

    def append(self, value, encoding='utf-8'):
        with self.path.open('ab') as stream:
            stream.write(value.encode(encoding) if isinstance(value, str) else value)

    def sync(self):
        return self.store.sync([str(self.path)])

    def rows(self, store=None):
        store = self.store if store is None else store
        with store.connection() as db:
            return [tuple(row) for row in db.execute(
                'SELECT number,text,normalized,ref,sn,lot,modbus,op,op_ref,loading,time,dims,lot_hint '
                'FROM lines ORDER BY number')]

    def state(self):
        with self.store.connection() as db:
            return db.execute('SELECT state FROM file_resume').fetchone()[0]

    def assert_full_equivalent(self, preserve_fallback_start=False):
        self.full_counter += 1
        full = TraceStore(self.root / f'full{self.full_counter}.db')
        self.assertFalse(full.sync([str(self.path)])['failures'])
        self.assertEqual(self.rows(), self.rows(full))
        metadata, full_metadata = self.store.metadata(), full.metadata()
        if preserve_fallback_start:
            # With no trace creation header, append keeps the original fallback
            # date instead of moving it every time filesystem mtime changes.
            full_metadata[0]['start'] = metadata[0]['start']
        self.assertEqual(metadata, full_metadata)
        self.assertEqual(self.store.file_lines(str(self.path)), read_trace(self.path).splitlines())
        for store in (self.store, full):
            with store.connection() as db:
                values = [tuple(row) for row in db.execute(
                    'SELECT l.number,v.col,v.value FROM modbus_values v JOIN lines l ON l.id=v.line_id '
                    'ORDER BY l.number,v.col')]
            if store is self.store:
                expected_values = values
            else:
                self.assertEqual(values, expected_values)
        for term in ('accentué', 'RefProduit=ABC', 'GRIPPER', 'PART', 'not present'):
            self.assertEqual(self.store.search(term), full.search(term))

    def test_append_keeps_prefix_rows_and_only_parses_overlapped_tail(self):
        text = self.prefix() + 'D 31|23:59:59 RefProduit=ABC ModBusId=12\n'
        with patch('trace_store.detect_encoding', side_effect=AssertionError('Second encoding pass')):
            self.write(text)
        before_size = self.path.stat().st_size
        with self.store.connection() as db:
            prefix_ids = [tuple(row) for row in db.execute('SELECT id,number FROM lines WHERE number<10')]
        self.store.file_lines(str(self.path))
        addition = ('G 01|00:00:01 GRIPPER GET 1x boîte(s): Code:ABC $0101\n'
                    'B 01|00:00:02 IDPK pour ABC : l=10 H=11 L=12\n'
                    'G 01|00:00:03 GRIPPER PUTE 1x boîte(s): Code:ABC\n')
        self.append(addition)
        with patch('trace_store.parse_fields', wraps=trace_store.parse_fields) as parse:
            result = self.sync()
        self.assertFalse(result['failures'])
        metrics = result['indexing'][0]
        self.assertEqual(metrics['mode'], 'append')
        self.assertEqual(metrics['lines_parsed'], 4)
        self.assertEqual(parse.call_count, 4)
        self.assertLess(metrics['bytes_parsed'], self.path.stat().st_size)
        self.assertEqual(metrics['bytes_verified'], before_size)
        self.assertEqual(metrics['bytes_read'], self.path.stat().st_size)
        self.assertFalse(self.store.text_cache)
        with self.store.connection() as db:
            self.assertEqual(prefix_ids, [tuple(row) for row in db.execute('SELECT id,number FROM lines WHERE number<10')])
        details = self.store.details(str(self.path), 10)
        self.assertEqual(details['idpk']['number'], 12)
        self.assertEqual(json.loads(details['idpk']['dims']), {'l': 10, 'H': 11, 'L': 12})
        self.assert_full_equivalent()

    def test_newline_and_unterminated_line_variants_match_full_index(self):
        for ending in ('', '\n', '\r', '\r\n', '\v', '\x85', '\u2028'):
            with self.subTest(ending=repr(ending)):
                self.write(self.prefix() + 'D 31|23:59:59 RefProduit=PART' + ending)
                self.append('IAL ModBusId=12\r\n\r\nD 01|00:00:01 RefProduit=ABC ModBusId=13\n')
                result = self.sync()
                self.assertFalse(result['failures'])
                self.assertEqual(result['indexing'][0]['mode'], 'append')
                self.assert_full_equivalent()

    def test_lot_and_date_context_survives_tail_reparse(self):
        text = (self.prefix() + 'D 31|23:59:59 Rech. si ABC (lot:LOT1)\n'
                'D 31|23:59:59 Rech. si ABC (lot:LOT2)\vD 31|23:59:59 Rech. si ABC (lot:LOT3)')
        self.write(text)
        self.append('\nG 01|00:00:01 GRIPPER PUTE 1x boîte(s): Code:ABC\n')
        self.assertEqual(self.sync()['indexing'][0]['mode'], 'append')
        event = self.store.events(str(self.path), 'ABC', 'PUTE')[0]
        self.assertEqual(event['lot_hint'], 'LOT3')
        self.assertEqual(datetime.fromtimestamp(event['time']), datetime(2026, 1, 1, 0, 0, 1))
        self.assert_full_equivalent()

    def test_encodings_and_chunk_boundaries(self):
        for encoding in ('utf-8', 'utf-8-sig', 'cp1252'):
            with self.subTest(encoding=encoding):
                self.write(self.prefix() + 'D 31|23:59:59 final café', encoding)
                # Encoding a suffix with utf-8-sig would insert another actual BOM.
                suffix_encoding = 'utf-8' if encoding == 'utf-8-sig' else encoding
                self.append(' ajouté\r\nD 01|00:00:01 autre référence\n', suffix_encoding)
                self.assertEqual(self.sync()['indexing'][0]['mode'], 'append')
                self.assert_full_equivalent()
        prefix = self.prefix().encode('utf-8')
        self.path.write_bytes(prefix + b'x' * (65535 - len(prefix)) + b'\r')
        self.sync()
        self.append(b'\n' + 'D 01|00:00:01 événement\n'.encode('utf-8'))
        self.assertFalse(self.sync()['failures'])
        self.assert_full_equivalent()

    def test_partial_established_utf8_keeps_old_snapshot_until_completed(self):
        self.write(self.prefix() + 'D 31|23:59:59 final\n', 'utf-8-sig')
        before, state, metadata = self.rows(), self.state(), self.store.metadata()
        self.append('D 01|00:00:01 prix '.encode('utf-8') + b'\xe2\x82')
        result = self.sync()
        self.assertTrue(result['failures'])
        self.assertEqual((self.rows(), self.state(), self.store.metadata()), (before, state, metadata))
        self.append(b'\xac\n')
        self.assertEqual(self.sync()['indexing'][0]['mode'], 'append')
        self.assert_full_equivalent()

    def test_initial_bom_partial_utf8_is_retried_and_cp1252_final_accent_is_kept(self):
        self.path.write_bytes(b'\xef\xbb\xbf' + self.prefix().encode('utf-8') + b'\xc3')
        self.assertTrue(self.sync()['failures'])
        self.assertFalse(self.store.metadata())
        self.append(b'\xa9')
        self.assertFalse(self.sync()['failures'])
        self.assert_full_equivalent()
        self.write(self.prefix() + 'D 31|23:59:59 café', 'cp1252')
        self.assertEqual(self.store.file_lines(str(self.path))[-1], 'D 31|23:59:59 café')

    def test_ascii_prefix_can_transition_to_cp1252_without_retry_loop(self):
        self.write(self.prefix().replace('accentué', 'ascii') + 'ASCII\n')
        self.append(b'caf\xe9')
        result = self.sync()
        self.assertFalse(result['failures'])
        self.assertEqual(result['indexing'][0]['reason'], 'encoding_changed')
        self.assert_full_equivalent()

    def test_changed_prefix_falls_back_without_stale_fts(self):
        self.write(self.prefix() + 'RefProduit=OLD ModBusId=12\nlast\n')
        self.path.write_bytes(self.path.read_bytes().replace(b'OLD', b'NEW') + b'appended\n')
        result = self.sync()
        self.assertEqual((result['indexing'][0]['mode'], result['indexing'][0]['reason']),
                         ('full', 'prefix_rewritten'))
        self.assertFalse(self.store.search('RefProduit=OLD'))
        self.assert_full_equivalent()

    def test_truncation_same_size_rewrite_and_rotation_reindex(self):
        self.write(self.prefix() + 'OLD\n')
        data = self.path.read_bytes().replace(b'OLD', b'NEW')
        previous_stat = self.path.stat()
        self.path.write_bytes(data)
        os.utime(self.path, ns=(previous_stat.st_atime_ns, previous_stat.st_mtime_ns + 1000000))
        self.assertEqual(self.sync()['indexing'][0]['reason'], 'not_append')
        self.assert_full_equivalent()
        self.path.write_bytes(data[:200])
        self.assertEqual(self.sync()['indexing'][0]['reason'], 'not_append')
        self.assert_full_equivalent()
        replacement = self.path.with_suffix('.replacement')
        replacement.write_bytes((self.prefix() + 'different longer replacement\n').encode('utf-8'))
        os.replace(replacement, self.path)
        self.assertEqual(self.sync()['indexing'][0]['reason'], 'file_replaced')
        self.assert_full_equivalent()

    def test_live_append_commits_bounded_snapshot_and_next_pass_adds_rest(self):
        self.write(self.prefix() + 'initial\n')
        self.append('FIRST\n')
        original = self.store._index_snapshot
        def grow(*args):
            self.append('SECOND\n')
            return original(*args)
        with patch.object(self.store, '_index_snapshot', side_effect=grow):
            result = self.sync()
        self.assertFalse(result['failures'])
        self.assertEqual(result['indexing'][0]['mode'], 'append')
        self.assertTrue(self.store.search('FIRST'))
        self.assertFalse(self.store.search('SECOND'))
        self.assertEqual(self.sync()['indexing'][0]['mode'], 'append')
        self.assert_full_equivalent()

    def test_rewrite_during_incremental_parse_rolls_back_rows_and_resume(self):
        self.write(self.prefix() + 'initial\n')
        before, state, metadata = self.rows(), self.state(), self.store.metadata()
        self.append('FIRST\n')
        original = self.store._index_snapshot
        def rewrite(*args):
            previous_stat = self.path.stat()
            self.path.write_bytes(self.path.read_bytes().replace(b'diagnostic 0', b'diagnostic X'))
            os.utime(self.path, ns=(previous_stat.st_atime_ns, previous_stat.st_mtime_ns + 1000000))
            return original(*args)
        with patch.object(self.store, '_index_snapshot', side_effect=rewrite):
            self.assertTrue(self.sync()['failures'])
        self.assertEqual((self.rows(), self.state(), self.store.metadata()), (before, state, metadata))
        self.assertEqual(self.sync()['indexing'][0]['mode'], 'full')
        self.assert_full_equivalent()

    def test_cancellation_after_inserts_rolls_back_and_wal_reader_stays_consistent(self):
        self.write(self.prefix() + 'initial\n')
        before, state = self.rows(), self.state()
        self.append('FIRST\n')
        stop = threading.Event()
        original = self.store._insert_batch
        def cancel_after_insert(*args):
            original(*args)
            stop.set()
        with patch.object(self.store, '_insert_batch', side_effect=cancel_after_insert):
            with self.assertRaises(Cancelled):
                self.store.sync([str(self.path)], stop=stop)
        self.assertEqual((self.rows(), self.state()), (before, state))
        with self.store.snapshot():
            self.assertFalse(self.store.search('FIRST'))
            with ThreadPoolExecutor(1) as pool:
                result = pool.submit(self.sync).result(timeout=10)
            self.assertEqual(result['indexing'][0]['mode'], 'append')
            self.assertFalse(self.store.search('FIRST'))
        self.assertTrue(self.store.search('FIRST'))
        self.assert_full_equivalent()

    def test_modbus_append_replaces_tail_values_and_keeps_register_positions(self):
        self.path = self.path.with_name('TracMODBUS.txt')
        def modbus_row(stamp, values):
            row = ['0'] * 121
            row[0] = stamp
            for col, value in values.items():
                row[col] = value
            return ';'.join(row)
        self.write('\n'.join(modbus_row('31|23:59:59', {2: str(i + 1)}) for i in range(10)) + '\n' +
                   modbus_row('31|23:59:59', {6: '12', 114: '99'}))
        initial_start = self.store.metadata()[0]['start']
        self.append('\r\n' + modbus_row('01|00:00:01', {7: '12', 115: '-99'}) + '\n')
        self.assertEqual(self.sync()['indexing'][0]['mode'], 'append')
        self.assertEqual(self.store.metadata()[0]['start'], initial_start)
        self.assert_full_equivalent(preserve_fallback_start=True)

    def test_v1_cache_gets_additive_resume_without_reading_unchanged_traces(self):
        self.write(self.prefix() + 'initial\n')
        before = self.rows()
        with self.store.connection() as db, db:
            db.execute('DROP TABLE file_resume')
        with patch('trace_store.detect_encoding', side_effect=AssertionError('Unchanged trace read')):
            reopened = TraceStore(self.database)
            self.assertEqual(reopened.sync([str(self.path)])['changed'], 0)
        self.store = reopened
        self.assertEqual(self.rows(), before)
        with self.store.connection() as db:
            self.assertEqual(db.execute('PRAGMA user_version').fetchone()[0], 1)
        self.append('FIRST\n')
        self.assertEqual(self.sync()['indexing'][0]['reason'], 'resume_missing')
        self.append('SECOND\n')
        self.assertEqual(self.sync()['indexing'][0]['mode'], 'append')
        self.assert_full_equivalent()

    def test_homonymous_sources_keep_independent_rows_and_resume_context(self):
        self.write(self.prefix() + 'RefProduit=R1 ModBusId=12\n')
        other = self.root / 'r2' / self.path.name
        other.parent.mkdir()
        other.write_bytes((self.prefix() + 'RefProduit=R2 ModBusId=12\n').encode('utf-8'))
        paths = [str(self.path), str(other)]
        self.store.sync(paths)
        with self.store.connection() as db:
            other_rows = [tuple(row) for row in db.execute(
                'SELECT l.* FROM lines l JOIN files f ON f.id=l.file_id WHERE f.path=? ORDER BY number',
                (str(other),))]
        self.append('RefProduit=APPEND_R1 ModBusId=13\n')
        result = self.store.sync(paths)
        self.assertEqual(len(result['indexing']), 1)
        self.assertEqual(result['indexing'][0]['mode'], 'append')
        self.assertEqual(result['indexing'][0]['path'], str(self.path))
        with self.store.connection() as db:
            self.assertEqual(other_rows, [tuple(row) for row in db.execute(
                'SELECT l.* FROM lines l JOIN files f ON f.id=l.file_id WHERE f.path=? ORDER BY number',
                (str(other),))])
        hits = self.store.search('ModBusId=12', exact_modbus=True)
        self.assertEqual({path for path, _ in hits}, set(paths))
        self.assertEqual(self.store.file_lines(str(other)), read_trace(other).splitlines())

    def test_malformed_resume_falls_back_and_short_header_remains_editable(self):
        self.write(self.prefix() + 'initial\n')
        for key, value in (('number', 1.5), ('offset', '0'), ('lots', []), ('digest', 'bad')):
            with self.subTest(key=key):
                state = json.loads(self.state())
                state[key] = value
                with self.store.connection() as db, db:
                    db.execute('UPDATE file_resume SET state=?', (json.dumps(state),))
                self.append('extra\n')
                self.assertEqual(self.sync()['indexing'][0]['reason'], 'resume_invalid')
                self.assert_full_equivalent()
        self.write('Ancienne trace = Trac')
        self.append('Omega029.old\ncreation de la trace Le 31/12/2025\n')
        self.assertEqual(self.sync()['indexing'][0]['reason'], 'header_unfinished')
        self.assertEqual(self.store.metadata()[0]['previous'], 'TracOmega029.old')
        self.assert_full_equivalent()

    def test_idpk_belongs_to_selected_get_and_stops_at_next_loading_get(self):
        def occurrence(marker, dims=True):
            return ('D 31|23:59:59 RefProduit=ABC ModBusId=12\n'
                    f'G 31|23:59:59 GRIPPER GET 1x boîte(s): Code:ABC $0101 {marker}\n' +
                    (f'B 31|23:59:59 IDPK pour ABC : l={marker} H=10 L=20\n' if dims else ''))
        self.write(self.prefix() + occurrence('1') + occurrence('2', False) + occurrence('3'))
        self.assertEqual(json.loads(self.store.details(str(self.path), 10)['idpk']['dims'])['l'], 1)
        self.assertIsNone(self.store.details(str(self.path), 13)['idpk'])
        self.assertEqual(json.loads(self.store.details(str(self.path), 15)['idpk']['dims'])['l'], 3)


if __name__ == '__main__':
    unittest.main()
