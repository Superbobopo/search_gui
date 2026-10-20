"""Persistent, incrementally refreshed trace index. No Tk dependencies."""

from collections import OrderedDict, defaultdict
from contextlib import contextmanager
from datetime import datetime
import codecs
import hashlib
import json
import ntpath
import os
from pathlib import Path
import re
import sqlite3
import threading
import tempfile
import time
from types import SimpleNamespace

from trace_parser import (parse_fields, parse_idpk, GRIPPER_RE, line_datetime, read_trace, timestamp_parts,
                          trace_start_datetime, parse_modbus_line)
from trace_sources import application_directory, path_key, trace_label, robot_scope
from trace_identity import TraceRef
from omega_history import OmegaHistory, history_path, under, HISTORY_APPLICATION_ID, HistoryUnavailable
from occurrence_comparison import AnalysisCache, comparisons as compare_tickets, metadata as analysis_metadata, timed as analysis_timed

SCHEMA_VERSION = 1
CHECKPOINTS = dict(zip(
    [2, 31, 32, 33, 34, 35, 36, 37, 57, 58, 6, 114, 7, 115],
    ["Boite entrée", "En attente en bas de l’asc POSTE1", "En attente en bas de l’asc POSTE2",
     "En attente en bas de l’asc POSTE3", "En attente en bas de l’asc POSTE4",
     "Boite en train d’être convoyée vers l’ascenseur 1 ou 2",
     "Boite arrivée devant l’ascenseur 1", "Boite arrivée devant l’ascenseur 2",
     "Boite poussée depuis l’ascenseur 1 vers le T5/R1", "Boite poussée depuis l’ascenseur 2 vers le T5/R2",
     "Boite en attente sur T5 / R1", "Unnamed: 114", "Boite en attente sur T5 / R2", "Unnamed: 115"]))
MODBUS_COLUMNS = sorted(set(CHECKPOINTS) | {1, 2, 3, 4})
LOT_RE = re.compile(r"Rech\. si\s+(\S+)\s+\(lot:([^)]*)\)")


class Cancelled(Exception):
    pass


class _EncodingProbe:
    """Validate UTF-8 while the bounded snapshot is copied, without another pass."""
    def __init__(self):
        self.decoder = codecs.getincrementaldecoder('utf-8-sig')('strict')
        self.valid = True
        self.confirmed = False
        self.first = b''

    def feed(self, chunk):
        if len(self.first) < 3:
            self.first += chunk[:3 - len(self.first)]
        if self.valid:
            try:
                text = self.decoder.decode(chunk)
                self.confirmed = self.confirmed or not text.isascii()
            except UnicodeDecodeError:
                self.valid = False

    def finish(self):
        partial = False
        if self.valid:
            try:
                self.decoder.decode(b'', final=True)
            except UnicodeDecodeError as error:
                partial = error.reason == 'unexpected end of data'
                self.valid = False
        self.confirmed = self.confirmed or self.first == codecs.BOM_UTF8
        return ('utf-8-sig' if self.valid else 'cp1252'), partial, self.confirmed


def _physical_lines(stream, base_offset=0):
    """Yield universal-newline physical records and their exact byte offsets."""
    stream.seek(0)
    pending, offset = b'', base_offset
    while True:
        chunk = stream.read(65536)
        data = pending + chunk
        consumed = 0
        for match in re.finditer(br'\r\n|\r|\n', data):
            # A CR at the chunk boundary may be the first half of CRLF.
            if chunk and match.end() == len(data) and match[0] == b'\r':
                break
            raw = data[consumed:match.end()]
            yield offset, raw
            offset += len(raw)
            consumed = match.end()
        pending = data[consumed:]
        if not chunk:
            if pending:
                yield offset, pending
            return


def _decode_physical(raw, encoding, offset):
    # utf-8-sig must not strip a literal BOM occurring inside the file.
    codec = 'utf-8' if encoding == 'utf-8-sig' and offset else encoding
    return raw.decode(codec, errors='replace').replace('\r\n', '\n').replace('\r', '\n')


def cache_path(directories=None):
    identity = hashlib.sha256(path_key(application_directory()).encode()).hexdigest()[:16]
    root = Path(os.environ.get("LOCALAPPDATA", Path.home() / ".cache"))
    filename = "trace_index.sqlite3"
    if directories is not None:
        sources = '\n'.join(sorted({path_key(path) for path in directories}))
        scope = hashlib.sha256(sources.encode()).hexdigest()[:16]
        filename = f"trace_index_sources_{scope}_history.sqlite3"
    return root / "TracOmega" / "cache" / identity / filename


def detect_encoding(path):
    decoder = codecs.getincrementaldecoder("utf-8-sig")("strict")
    owned = not hasattr(path, 'read')
    stream = open(path, 'rb') if owned else path
    try:
        stream.seek(0)
        while chunk := stream.read(65536):
            decoder.decode(chunk)
        decoder.decode(b"", final=True)
        return "utf-8-sig"
    except UnicodeDecodeError:
        return "cp1252"
    finally:
        if owned:
            stream.close()
        else:
            stream.seek(0)


class TraceStore:
    def __init__(self, database=None, *, directories=None, history_database=None, stop=None, defer_history=False):
        self.database = Path(database) if database is not None else cache_path()
        self.database.parent.mkdir(parents=True, exist_ok=True)
        self._writer_lock = threading.Lock()
        self._cache_lock = threading.Lock()
        self._local = threading.local()
        self._connections = threading.Condition()
        self._active_connections = 0
        self._recovery_owner = None
        self.corrupt = threading.Event()
        self.text_cache = OrderedDict()
        self._analysis_cache = AnalysisCache()
        self._analysis_prepared = False
        self.history = None
        self.directories = tuple(directories) if directories is not None else None
        self._history_verified = set()
        self._history_prepared = False
        self._history_prepare_lock = threading.Lock()
        self._initial_stop = stop
        if directories is not None:
            destination = history_database or (history_path() if self.database.parent == cache_path().parent
                                               else self.database.parent / 'omega_history.sqlite3')
            if Path(destination).resolve() == self.database.resolve():
                raise ValueError('L’historique OMEGA et le cache doivent être deux fichiers distincts.')
            self.history = OmegaHistory(destination)
            from trace_sources import trace_directories
            self._history_roots = tuple(dict.fromkeys((*self.directories, *trace_directories())))
            if not defer_history:
                self.prepare_history(stop)
        self._initialize()

    def prepare_history(self, stop=None):
        """Recover legacy snapshots after cache availability, once per session."""
        if self.history is None or self._history_prepared:
            return
        with self._history_prepare_lock:
            if self._history_prepared:
                return
            candidates = [self.database, *self.database.parent.glob('trace_index*.sqlite3')]
            self.history.migrate_caches(candidates, self._history_roots,
                                       lambda: self._check_cancel(stop or self._initial_stop))
            self._history_prepared = True

    def prepare_analysis(self, stop=None):
        """Add query indexes after the existing cache has been made available."""
        if self._analysis_prepared:
            return
        with self._writer_lock, self.connection() as db, db:
            db.set_progress_handler(lambda: int(stop.is_set()) if stop else 0, 1000)
            try:
                self._check_cancel(stop)
                db.execute('''CREATE INDEX IF NOT EXISTS ticket_time ON lines(file_id,time,number)
                    WHERE modbus IS NOT NULL AND ref IS NOT NULL''')
                db.execute('''CREATE INDEX IF NOT EXISTS comparison_events ON lines(op_ref,file_id,number)
                    WHERE loading=1 OR dims IS NOT NULL''')
            finally:
                db.set_progress_handler(None, 0)
        self._analysis_prepared = True

    def analysis_revision(self):
        with self.connection() as db:
            row = db.execute("SELECT value FROM state WHERE key='content_revision'").fetchone()
            return row[0] if row else 0

    @contextmanager
    def connection(self):
        existing = getattr(self._local, 'snapshot', None)
        if existing is not None:
            yield existing
            return
        with self._connections:
            self._connections.wait_for(lambda: self._recovery_owner in (None, threading.get_ident()))
            self._active_connections += 1
        connection = None
        try:
            connection = sqlite3.connect(self.database, timeout=30)
            connection.row_factory = sqlite3.Row
            connection.execute("PRAGMA foreign_keys=ON")
            connection.execute("PRAGMA cache_size=-8192")
            yield connection
        except sqlite3.DatabaseError as error:
            if getattr(error, 'sqlite_errorcode', 0) in (sqlite3.SQLITE_CORRUPT, sqlite3.SQLITE_NOTADB):
                self.corrupt.set()
            raise
        finally:
            if connection is not None:
                connection.close()
            with self._connections:
                self._active_connections -= 1
                self._connections.notify_all()

    @contextmanager
    def snapshot(self):
        with self.connection() as db:
            db.execute('BEGIN')
            self._local.snapshot = db
            try:
                yield
            finally:
                del self._local.snapshot
                db.rollback()

    def _initialize(self):
        try:
            with self.connection() as db:
                if db.execute('PRAGMA application_id').fetchone()[0] == HISTORY_APPLICATION_ID:
                    raise RuntimeError('Cette base est un historique OMEGA permanent, pas un cache ; fichier conservé.')
                version = db.execute("PRAGMA user_version").fetchone()[0]
                db.execute("SELECT name FROM sqlite_master LIMIT 1").fetchone()
                if version not in (0, SCHEMA_VERSION):
                    raise sqlite3.DatabaseError("Incompatible cache version")
                db.execute("PRAGMA journal_mode=WAL")
                db.executescript("""
                    CREATE TABLE IF NOT EXISTS files (
                        id INTEGER PRIMARY KEY, path TEXT UNIQUE, directory TEXT,
                        size INTEGER, mtime INTEGER, encoding TEXT, start REAL, end REAL,
                        previous TEXT, line_count INTEGER, kind TEXT);
                    CREATE TABLE IF NOT EXISTS lines (
                        id INTEGER PRIMARY KEY, file_id INTEGER REFERENCES files ON DELETE CASCADE,
                        number INTEGER, text TEXT, normalized TEXT,
                        ref TEXT, sn TEXT, lot TEXT, modbus TEXT,
                        op TEXT, op_ref TEXT, loading INTEGER, time REAL,
                        dims TEXT, lot_hint TEXT);
                    CREATE UNIQUE INDEX IF NOT EXISTS line_location ON lines(file_id,number);
                    CREATE INDEX IF NOT EXISTS line_modbus ON lines(modbus,file_id,number) WHERE modbus IS NOT NULL;
                    CREATE INDEX IF NOT EXISTS line_ref ON lines(ref,file_id,number) WHERE ref IS NOT NULL;
                    CREATE INDEX IF NOT EXISTS line_sn ON lines(sn,file_id,number) WHERE sn IS NOT NULL;
                    CREATE INDEX IF NOT EXISTS line_lot ON lines(lot,file_id,number) WHERE lot IS NOT NULL;
                    CREATE INDEX IF NOT EXISTS line_event ON lines(file_id,op_ref,op,number) WHERE op IS NOT NULL;
                    CREATE INDEX IF NOT EXISTS event_reference ON lines(op_ref,op,time) WHERE op IS NOT NULL;
                    CREATE INDEX IF NOT EXISTS operation_reference ON lines(file_id,op_ref,number) WHERE op_ref IS NOT NULL;
                    CREATE INDEX IF NOT EXISTS line_dimensions ON lines(file_id,op_ref,number) WHERE dims IS NOT NULL;
                    CREATE TABLE IF NOT EXISTS modbus_values (
                        line_id INTEGER REFERENCES lines ON DELETE CASCADE, col INTEGER, value TEXT);
                    CREATE INDEX IF NOT EXISTS modbus_lookup ON modbus_values(value,line_id,col);
                    CREATE INDEX IF NOT EXISTS modbus_line ON modbus_values(line_id);
                    CREATE VIRTUAL TABLE IF NOT EXISTS lines_fts USING fts5(
                        normalized, content='lines', content_rowid='id',
                        tokenize='trigram', detail='none');
                    CREATE TABLE IF NOT EXISTS fts_ingest (
                        id INTEGER PRIMARY KEY CHECK(id=1), active INTEGER NOT NULL);
                    INSERT OR IGNORE INTO fts_ingest VALUES(1,0);
                    CREATE TRIGGER IF NOT EXISTS line_insert AFTER INSERT ON lines
                    WHEN (SELECT active FROM fts_ingest WHERE id=1)=0 BEGIN
                        INSERT INTO lines_fts(rowid,normalized) VALUES(new.id,new.normalized);
                    END;
                    CREATE TRIGGER IF NOT EXISTS line_delete AFTER DELETE ON lines BEGIN
                        INSERT INTO lines_fts(lines_fts,rowid,normalized) VALUES('delete',old.id,old.normalized);
                    END;
                    CREATE TABLE IF NOT EXISTS state(key TEXT PRIMARY KEY,value INTEGER);
                    CREATE TABLE IF NOT EXISTS file_resume (
                        file_id INTEGER PRIMARY KEY REFERENCES files ON DELETE CASCADE,
                        state TEXT NOT NULL);
                    CREATE TABLE IF NOT EXISTS archive_files (
                        key TEXT PRIMARY KEY, metadata TEXT NOT NULL);
                """)
                # Additive cache migration: existing rows/FTS remain untouched.
                trigger = db.execute("SELECT sql FROM sqlite_master WHERE name='line_insert'").fetchone()
                if trigger and 'fts_ingest' not in trigger[0]:
                    db.execute('BEGIN IMMEDIATE')
                    db.execute('DROP TRIGGER line_insert')
                    db.execute('''CREATE TRIGGER line_insert AFTER INSERT ON lines
                        WHEN (SELECT active FROM fts_ingest WHERE id=1)=0 BEGIN
                        INSERT INTO lines_fts(rowid,normalized) VALUES(new.id,new.normalized); END''')
                db.execute(f"PRAGMA user_version={SCHEMA_VERSION}")
                db.execute("INSERT OR IGNORE INTO state VALUES('content_revision',0)")
                for table in ('files', 'archive_files'):
                    for operation in ('INSERT', 'UPDATE', 'DELETE'):
                        db.execute(f'''CREATE TRIGGER IF NOT EXISTS revision_{table}_{operation.lower()}
                            AFTER {operation} ON {table} BEGIN
                            UPDATE state SET value=value+1 WHERE key='content_revision'; END''')
                db.commit()
        except sqlite3.DatabaseError as error:
            code = getattr(error, 'sqlite_errorcode', 0) & 0xff
            recoverable = code in (sqlite3.SQLITE_CORRUPT, sqlite3.SQLITE_NOTADB) or str(error) == 'Incompatible cache version'
            if not recoverable or getattr(self, '_recovered', False):
                raise
            self._recovered = True
            # A deferred legacy import must finish BEFORE discarding any cache.
            self.prepare_history(self._initial_stop)
            # This directory contains only disposable data owned by this service.
            for suffix in ("", "-wal", "-shm"):
                Path(str(self.database) + suffix).unlink(missing_ok=True)
            self._initialize()
            self.corrupt.clear()

    def rebuild(self):
        """Close readers before removing disposable database files (also on Windows)."""
        with self._writer_lock:
            self.prepare_history(self._initial_stop)
            with self._connections:
                self._recovery_owner = threading.get_ident()
                self._connections.wait_for(lambda: self._active_connections == 0)
            try:
                for suffix in ('', '-wal', '-shm'):
                    Path(str(self.database) + suffix).unlink(missing_ok=True)
                self._recovered = False
                self._initialize()
                self.corrupt.clear()
                with self._cache_lock:
                    self.text_cache.clear()
                self._analysis_cache.clear()
                self._analysis_prepared = False
            finally:
                with self._connections:
                    self._recovery_owner = None
                    self._connections.notify_all()

    def metadata(self):
        with self.connection() as db:
            refs = self._archive_refs(db)
            return [self._with_ref(dict(row), refs) for row in db.execute("SELECT * FROM files ORDER BY end DESC,path DESC")]

    @staticmethod
    def _archive_refs(db, keys=None):
        if keys is None:
            return {row[0]: json.loads(row[1]) for row in db.execute('SELECT key,metadata FROM archive_files')}
        keys = list(keys)
        result = {}
        for start in range(0, len(keys), 500):
            batch = keys[start:start + 500]
            sql = 'SELECT key,metadata FROM archive_files WHERE key IN (' + ','.join('?' for key in batch) + ')'
            result.update((row[0], json.loads(row[1])) for row in db.execute(sql, batch))
        return result

    @staticmethod
    def _with_ref(row, refs):
        path = row.get('path')
        if path in refs:
            entry = refs[path]
            row['path'] = TraceRef(entry['path'], entry['id'], entry['start'], entry['end'], not entry['present'])
            row['version_id'] = entry['id']
            row['aliases'] = entry['aliases']
        return row

    def initialized(self):
        with self.connection() as db:
            row = db.execute("SELECT value FROM state WHERE key='complete'").fetchone()
            return bool(row and row[0])

    def sync(self, paths, progress=None, stop=None):
        """Publish each stable file atomically; readers keep using the WAL snapshot."""
        if self.history is not None:
            self.prepare_history(stop)
            return self._sync_history(paths, progress, stop)
        with self._writer_lock:
            paths = {path_key(path): os.path.abspath(path) for path in paths}
            known = {path_key(item['path']): item for item in self.metadata()}
            changed = []
            for key, path in paths.items():
                try:
                    stat = os.stat(path)
                except OSError:
                    continue
                previous = known.get(key)
                if not previous or (previous['size'], previous['mtime']) != (stat.st_size, stat.st_mtime_ns):
                    changed.append((path, stat))
            # Current files first; for archives retain known chronology when available.
            def priority(item):
                path, stat = item
                date = known.get(path_key(path), {}).get('end')
                if date is None:
                    try:
                        with open(path, 'rb') as stream:
                            header = stream.read(2048).decode('cp1252', errors='replace')
                        date = trace_start_datetime(header, datetime.fromtimestamp(stat.st_mtime)).timestamp()
                    except OSError:
                        date = stat.st_mtime
                return (os.path.basename(path).lower() not in ('tracomega.txt', 'tracmodbus.txt'), -date, path_key(path))
            changed.sort(key=priority)
            removed = [item for key, item in known.items() if key not in paths]
            if removed:
                with self.connection() as db, db:
                    db.executemany("DELETE FROM files WHERE id=?", [(row['id'],) for row in removed])
                with self._cache_lock:
                    for row in removed:
                        self.text_cache.pop(row['path'], None)
            failures = []
            indexing = []
            for index, (path, stat) in enumerate(changed):
                if stop and stop.is_set():
                    raise Cancelled()
                if progress:
                    progress(index, len(changed), path, False)
                try:
                    indexing.append(self._index_file(path, stat, stop))
                except (OSError, UnicodeError) as error:
                    failures.append((path, str(error)))
                if progress:
                    progress(index + 1, len(changed), path, True)
            if not failures and not self.initialized():
                with self.connection() as db, db:
                    db.execute("INSERT OR REPLACE INTO state VALUES('complete',1)")
            return {'changed': len(changed), 'removed': len(removed), 'failures': failures,
                    'indexing': indexing}

    def _sync_history(self, paths, progress=None, stop=None):
        with self._writer_lock:
            paths = list(dict.fromkeys(os.path.abspath(path) for path in paths
                                      if any(under(path, root) for root in self.directories)))
            locations = self.history.locations()
            failures, indexing, changed = [], [], 0
            check = lambda: self._check_cancel(stop)
            omega = [path for path in paths if os.path.basename(path).lower().startswith('tracomega')]
            present_paths = []
            for path in omega:
                check()
                try:
                    stat = os.stat(path)
                    location = locations.get(path_key(path))
                    signature = (stat.st_size, stat.st_mtime_ns, stat.st_dev, stat.st_ino)
                    if path_key(path) in self._history_verified and location and signature == tuple(location[k] for k in ('size', 'mtime', 'device', 'inode')):
                        present_paths.append(path)
                        continue
                    metrics = dict(bytes_read=0, bytes_verified=0, read_seconds=0., verify_seconds=0.)
                    # Verify raw bytes once at startup, without line parsing,
                    # encoding detection, or a temporary copy for known captures.
                    digest = self._fingerprint_source(path, stat, stop, metrics)
                    if self.history.recognize(path, stat, digest.hex(), check) is not None:
                        self._history_verified.add(path_key(path))
                        present_paths.append(path)
                        continue
                    with tempfile.TemporaryFile(dir=self.database.parent) as snapshot:
                        digest, _, encoding, partial, confirmed = self._copy_snapshot(path, stat, snapshot, None, stop, metrics)
                        if partial and confirmed:
                            raise OSError('Écriture UTF-8 incomplète ; archivage à réessayer')
                        self._verify_snapshot(path, stat, digest, stop, metrics)
                        root = next(root for root in self.directories if under(path, root))
                        self.history.publish(path, snapshot, stat, encoding, digest.hex(), root, check)
                    self._history_verified.add(path_key(path))
                    present_paths.append(path)
                except (OSError, UnicodeError) as error:
                    failures.append((path, 'OMEGA non archivée : ' + str(error)))
                except HistoryUnavailable:
                    raise
                except sqlite3.DatabaseError as error:
                    raise HistoryUnavailable('Historique OMEGA non actualisé ; nouvelles traces non archivées. '
                                             'Base conservée. ' + str(error)) from error
            self.history.mark_presence(present_paths, self.directories)
            # Index current MODBUS before the long archive projection so live
            # searches remain available progressively as in the original service.
            modbus = [path for path in paths if os.path.basename(path).lower().startswith('tracmodbus')]
            known = {row['path']: row for row in self.metadata() if not isinstance(row['path'], TraceRef)}
            removed = [row for path, row in known.items() if path not in modbus]
            with self.connection() as db, db:
                db.executemany('DELETE FROM files WHERE id=?', [(row['id'],) for row in removed])
            for position, path in enumerate(modbus):
                check()
                try:
                    stat = os.stat(path)
                    old = known.get(path)
                    if old and (old['size'], old['mtime']) == (stat.st_size, stat.st_mtime_ns):
                        continue
                    if progress:
                        progress(position, len(modbus), path, False)
                    indexing.append(self._index_file(path, stat, stop))
                    changed += 1
                    if progress:
                        progress(position + 1, len(modbus), path, True)
                except (OSError, UnicodeError) as error:
                    failures.append((path, str(error)))
            entries = self.history.entries(self.directories)
            entries.sort(key=lambda row: (os.path.basename(row['path']).lower() != 'tracomega.txt', -row['end']))
            with self.connection() as db:
                projected = {row[0]: (json.loads(row[1]), row[2] is not None)
                             for row in db.execute('''SELECT a.key,a.metadata,f.id FROM archive_files a
                                 LEFT JOIN files f ON f.path=a.key''')}
            # Publish archived generations to the disposable index. A crash here
            # is harmless: the committed history is replayed on the next sync.
            for position, entry in enumerate(entries):
                check()
                ref = TraceRef(entry['path'], entry['id'], entry['start'], entry['end'], not entry['present'])
                old, indexed = projected.get(ref.key, (None, False))
                if old == entry and indexed:
                    continue
                if progress:
                    progress(position, len(entries), ref, False)
                needs_projection = not indexed or not old or old['revision'] != entry['revision']
                if needs_projection:
                    with self.history.snapshot(entry['id'], entry['revision']) as snapshot:
                        stat = SimpleNamespace(st_size=entry['size'], st_mtime_ns=entry['mtime'],
                                               st_mtime=entry['mtime']/1e9, st_dev=entry['device'], st_ino=entry['inode'])
                        resume, reason = self._resume_state(ref, stat)
                        if resume:
                            prefix = hashlib.sha256()
                            remaining = resume['size']
                            while remaining:
                                check()
                                data = snapshot.read(min(65536, remaining))
                                if not data:
                                    resume = None
                                    break
                                prefix.update(data)
                                remaining -= len(data)
                            if resume and prefix.hexdigest() != resume['digest']:
                                resume = None
                        if resume:
                            snapshot.seek(resume['offset'])
                            with tempfile.TemporaryFile(dir=self.database.parent) as tail:
                                while data := snapshot.read(65536):
                                    check()
                                    tail.write(data)
                                tail.seek(0)
                                metrics = self._project_history(ref, stat, tail, entry, resume, reason, stop)
                        else:
                            snapshot.seek(0)
                            metrics = self._project_history(ref, stat, snapshot, entry, None, reason, stop)
                        indexing.append(metrics)
                if not needs_projection:
                    with self.connection() as db, db:
                        db.execute('INSERT OR REPLACE INTO archive_files VALUES(?,?)', (ref.key, json.dumps(entry)))
                changed += 1
                with self._cache_lock:
                    self.text_cache.pop(ref, None)
                if progress:
                    progress(position + 1, len(entries), ref, True)
            if not failures and not self.initialized():
                with self.connection() as db, db:
                    db.execute("INSERT OR REPLACE INTO state VALUES('complete',1)")
            return dict(changed=changed, removed=len(removed), failures=failures, indexing=indexing)

    def _fingerprint_source(self, path, stat, stop, metrics):
        digest = hashlib.sha256()
        with open(path, 'rb') as source:
            opened = os.fstat(source.fileno())
            if (opened.st_dev, opened.st_ino) != (stat.st_dev, stat.st_ino):
                raise OSError('Fichier remplacé avant la vérification')
            remaining = stat.st_size
            while remaining:
                self._check_cancel(stop)
                data = source.read(min(1024 * 1024, remaining))
                if not data:
                    raise OSError('Fichier tronqué pendant la vérification')
                digest.update(data)
                remaining -= len(data)
        self._verify_snapshot(path, stat, digest.digest(), stop, metrics)
        return digest.digest()

    def _project_history(self, ref, stat, snapshot, entry, resume, reason, stop):
        metrics = dict(path=ref, mode='append' if resume else 'full', reason=reason,
                       bytes_read=0, bytes_verified=0, bytes_parsed=0, lines_parsed=0,
                       read_seconds=0., parse_write_seconds=0., verify_seconds=0.)
        self._index_snapshot(ref, stat, snapshot, bytes.fromhex(entry['digest']), stop,
                             entry['encoding'], resume, metrics, False, history_entry=entry)
        return metrics

    def _index_file(self, path, stat, stop):
        started = time.perf_counter()
        resume, reason = self._resume_state(path, stat)
        metrics = {'path': path, 'mode': 'append' if resume else 'full', 'reason': reason,
                   'bytes_read': 0, 'bytes_verified': 0, 'bytes_parsed': 0, 'lines_parsed': 0,
                   'read_seconds': 0.0, 'parse_write_seconds': 0.0, 'verify_seconds': 0.0}
        # Hash all prior bytes, but copy and parse only the final physical line and
        # the appended bytes. Readers keep the previous WAL snapshot until commit.
        with tempfile.TemporaryFile(dir=self.database.parent) as snapshot:
            digest, prefix, encoding, partial, confirmed = self._copy_snapshot(
                path, stat, snapshot, resume, stop, metrics)
            established_utf8 = (resume and prefix == resume['digest']
                                and resume['encoding'] == 'utf-8-sig' and resume['utf8_confirmed'])
            if partial and (established_utf8 or (not resume and confirmed)):
                raise OSError('Écriture UTF-8 incomplète en fin de trace ; nouvelle tentative au prochain passage')
            if resume and (prefix != resume['digest'] or encoding != resume['encoding']):
                metrics['reason'] = 'prefix_rewritten' if prefix != resume['digest'] else 'encoding_changed'
                metrics['mode'] = 'full'
                resume = None
                original_digest = digest
                digest, _, encoding, _, confirmed = self._copy_snapshot(
                    path, stat, snapshot, None, stop, metrics)
                if original_digest != digest:
                    raise OSError('Contenu réécrit pendant la lecture ; nouvelle tentative au prochain passage')
            self._index_snapshot(path, stat, snapshot, digest, stop, encoding, resume, metrics, confirmed)
        metrics['total_seconds'] = time.perf_counter() - started
        return metrics

    def _resume_state(self, path, stat):
        with self.connection() as db:
            meta = db.execute('SELECT * FROM files WHERE path=?', (path,)).fetchone()
            if meta is None:
                return None, 'new_file'
            saved = db.execute('SELECT state FROM file_resume WHERE file_id=?', (meta['id'],)).fetchone()
        if saved is None:
            return None, 'resume_missing'
        try:
            state = json.loads(saved[0])
            integer_fields = ('version', 'size', 'mtime', 'device', 'inode', 'offset', 'number', 'physical')
            if (not isinstance(state, dict) or any(type(state[key]) is not int for key in integer_fields)
                    or state['version'] != 1 or state['size'] != meta['size'] or state['mtime'] != meta['mtime']
                    or state['encoding'] != meta['encoding'] or not 0 <= state['offset'] <= state['size']
                    or not 0 <= state['number'] <= meta['line_count'] or state['physical'] < 0
                    or len(bytes.fromhex(state['digest'])) != 32
                    or not isinstance(state['lots'], dict)
                    or not all(isinstance(k, str) and isinstance(v, str) for k, v in state['lots'].items())):
                return None, 'resume_invalid'
            datetime.fromtimestamp(state['time'])
            if state['device'] != stat.st_dev or state['inode'] != stat.st_ino:
                return None, 'file_replaced'
            if stat.st_size <= state['size']:
                return None, 'not_append'
            # An unfinished first-eight-line header can still change start/previous.
            if state['physical'] < 8:
                return None, 'header_unfinished'
            state['utf8_confirmed'] = bool(state['utf8_confirmed'])
        except (ValueError, KeyError, TypeError, OverflowError, OSError):
            return None, 'resume_invalid'
        state['metadata'] = dict(meta)
        return state, 'verified_growth'

    @staticmethod
    def _check_cancel(stop):
        if stop and stop.is_set():
            raise Cancelled()

    def _copy_snapshot(self, path, stat, snapshot, resume, stop, metrics):
        started = time.perf_counter()
        snapshot.seek(0)
        snapshot.truncate()
        digest, probe = hashlib.sha256(), _EncodingProbe()
        offset = resume['offset'] if resume else 0
        old_size = resume['size'] if resume else 0
        prefix = None
        position = 0
        with open(path, 'rb') as source:
            opened = os.fstat(source.fileno())
            if (opened.st_dev, opened.st_ino) != (stat.st_dev, stat.st_ino):
                raise OSError('Fichier remplacé avant la lecture')
            while position < stat.st_size:
                self._check_cancel(stop)
                amount = min(65536, stat.st_size - position)
                if position < old_size:
                    amount = min(amount, old_size - position)
                chunk = source.read(amount)
                if not chunk:
                    raise OSError('Fichier tronqué pendant la lecture')
                metrics['bytes_read'] += len(chunk)
                digest.update(chunk)
                probe.feed(chunk)
                if position + len(chunk) > offset:
                    snapshot.write(chunk[max(0, offset - position):])
                position += len(chunk)
                if resume and position == old_size:
                    prefix = digest.hexdigest()
                    metrics['bytes_verified'] += old_size
        encoding, partial, confirmed = probe.finish()
        metrics['read_seconds'] += time.perf_counter() - started
        snapshot.seek(0)
        return digest.digest(), prefix, encoding, partial, confirmed

    def _verify_snapshot(self, path, stat, digest, stop, metrics):
        started = time.perf_counter()
        self._check_cancel(stop)
        final = os.stat(path)
        if ((final.st_dev, final.st_ino) != (stat.st_dev, stat.st_ino)
                or final.st_size < stat.st_size):
            raise OSError('Fichier remplacé ou tronqué ; nouvelle tentative au prochain passage')
        if (final.st_size, final.st_mtime_ns) != (stat.st_size, stat.st_mtime_ns):
            check = hashlib.sha256()
            with open(path, 'rb') as source:
                opened = os.fstat(source.fileno())
                if (opened.st_dev, opened.st_ino) != (stat.st_dev, stat.st_ino):
                    raise OSError('Fichier remplacé pendant la vérification')
                remaining = stat.st_size
                while remaining:
                    self._check_cancel(stop)
                    chunk = source.read(min(65536, remaining))
                    if not chunk:
                        raise OSError('Fichier tronqué pendant la vérification')
                    check.update(chunk)
                    metrics['bytes_read'] += len(chunk)
                    metrics['bytes_verified'] += len(chunk)
                    remaining -= len(chunk)
            if check.digest() != digest:
                raise OSError('Contenu réécrit ; nouvelle tentative au prochain passage')
            checked = os.stat(path)
            if ((checked.st_dev, checked.st_ino) != (stat.st_dev, stat.st_ino)
                    or checked.st_size < stat.st_size):
                raise OSError('Fichier remplacé ou tronqué pendant la vérification')
        metrics['verify_seconds'] += time.perf_counter() - started

    def _index_snapshot(self, path, stat, snapshot, digest, stop, encoding, resume, metrics, confirmed, history_entry=None):
        started = time.perf_counter()
        if resume:
            meta = resume['metadata']
            start = datetime.fromtimestamp(meta['start'])
            current = datetime.fromtimestamp(resume['time'])
            previous, kind = meta['previous'], meta['kind']
            lots = dict(resume['lots'])
            number, physical_count, base_offset = resume['number'], resume['physical'], resume['offset']
        else:
            head = []
            for offset, raw in _physical_lines(snapshot):
                head.append(_decode_physical(raw, encoding, offset))
                if len(head) == 8:
                    break
            head = ''.join(head)
            start = trace_start_datetime(head, datetime.fromtimestamp(stat.st_mtime))
            current = start
            previous = head.splitlines()[0].split('=', 1)[1].strip() if head.startswith('Ancienne trace =') else ''
            kind = 'modbus' if os.path.basename(path).lower().startswith('tracmodbus') else 'omega'
            lots = {}
            number, physical_count, base_offset = 0, 0, 0
        self._check_cancel(stop)
        resume_line = {'offset': base_offset, 'number': number, 'physical': physical_count,
                       'time': current.timestamp()}
        last_lot_changes = []
        try:
            with self.connection() as db, db:
                if stop:
                    db.set_progress_handler(lambda: int(stop.is_set()), 1000)
                db.execute("INSERT OR IGNORE INTO files(path) VALUES(?)", (path,))
                file_id = db.execute("SELECT id FROM files WHERE path=?", (path,)).fetchone()[0]
                db.execute('DELETE FROM lines WHERE file_id=? AND number>=?', (file_id, number))
                db.execute('UPDATE fts_ingest SET active=1 WHERE id=1')
                first_number = number
                batch = []
                for offset, raw in _physical_lines(snapshot, base_offset):
                    resume_line = {'offset': offset, 'number': number, 'physical': physical_count,
                                   'time': current.timestamp()}
                    last_lot_changes = []
                    physical_count += 1
                    physical = _decode_physical(raw, encoding, offset)
                    for line in physical.splitlines():
                        self._check_cancel(stop)
                        fields = parse_fields(line)
                        op = GRIPPER_RE.search(line) if 'GRIPPER' in line else None
                        measurement = parse_idpk(line)
                        lot_match = LOT_RE.search(line) if 'Rech. si' in line else None
                        if lot_match:
                            last_lot_changes.append((lot_match[1], lot_match[1] in lots, lots.get(lot_match[1])))
                            lots[lot_match[1]] = lot_match[2]
                        current = line_datetime(line, current)
                        op_ref = op[2] if op else measurement[0] if measurement else None
                        batch.append((file_id, number, line, line.lower(), fields.get('RefProduit'),
                                      fields.get('SN'), fields.get('CodeLot'), fields.get('ModBusId'),
                                      op[1] if op else None, op_ref,
                                      int(bool(op and re.search(r'\$0101(?!\w)', line))),
                                      current.timestamp(), json.dumps(measurement[1]) if measurement else None,
                                      lots.get(op_ref) if op and op[1] == 'PUTE' else None))
                        number += 1
                        metrics['lines_parsed'] += 1
                        if len(batch) >= 1000:
                            self._insert_batch(db, batch, kind)
                            batch.clear()
                self._insert_batch(db, batch, kind)
                self._publish_fts_batch(db, file_id, first_number)
                if not isinstance(path, TraceRef):
                    self._verify_snapshot(path, stat, digest, stop, metrics)
                db.execute("""UPDATE files SET directory=?,size=?,mtime=?,encoding=?,start=?,end=?,
                           previous=?,line_count=?,kind=? WHERE id=?""",
                           (path_key(os.path.dirname(path)), stat.st_size, stat.st_mtime_ns, encoding,
                            start.timestamp(), current.timestamp(), previous, number, kind, file_id))
                # Undo only the last physical record's lot updates. Copying the
                # entire lot map on every line would make indexing quadratic.
                saved_lots = dict(lots)
                for reference, existed, value in reversed(last_lot_changes):
                    if existed:
                        saved_lots[reference] = value
                    else:
                        saved_lots.pop(reference, None)
                state = dict(resume_line, version=1, size=stat.st_size, mtime=stat.st_mtime_ns,
                             device=stat.st_dev, inode=stat.st_ino, digest=digest.hex(), encoding=encoding,
                             lots=saved_lots, utf8_confirmed=confirmed)
                db.execute('INSERT OR REPLACE INTO file_resume VALUES(?,?)',
                           (file_id, json.dumps(state, ensure_ascii=False)))
                if history_entry is not None:
                    db.execute('INSERT OR REPLACE INTO archive_files VALUES(?,?)',
                               (path.key, json.dumps(history_entry)))
                self._check_cancel(stop)
                db.set_progress_handler(None, 0)
        except sqlite3.OperationalError:
            self._check_cancel(stop)
            raise
        metrics['bytes_parsed'] = stat.st_size - base_offset
        metrics['parse_write_seconds'] = time.perf_counter() - started - metrics['verify_seconds']
        with self._cache_lock:
            self.text_cache.pop(path, None)

    @staticmethod
    def _publish_fts_batch(db, file_id, first_number):
        db.execute('''INSERT INTO lines_fts(rowid,normalized)
            SELECT id,normalized FROM lines WHERE file_id=? AND number>=?''', (file_id, first_number))
        db.execute('UPDATE fts_ingest SET active=0 WHERE id=1')

    @staticmethod
    def _insert_batch(db, batch, kind):
        if not batch:
            return
        first = db.execute("SELECT coalesce(max(id),0)+1 FROM lines").fetchone()[0]
        db.executemany("INSERT INTO lines VALUES(NULL,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", batch)
        if kind == 'modbus':
            values = []
            for offset, row in enumerate(batch):
                columns = parse_modbus_line(row[2])
                for col in MODBUS_COLUMNS:
                    if col < len(columns) and columns[col] not in ('', '0'):
                        values.append((first + offset, col, columns[col]))
            db.executemany("INSERT INTO modbus_values VALUES(?,?,?)", values)

    def _rows(self, sql, parameters=(), cancel=None):
        if cancel and cancel.is_set():
            raise Cancelled()
        with self.connection() as db:
            if cancel:
                db.set_progress_handler(lambda: int(cancel.is_set()), 1000)
            try:
                rows = [dict(row) for row in db.execute(sql, parameters)]
                refs = self._archive_refs(db, {row['path'] for row in rows if isinstance(row.get('path'), str) and row['path'].startswith('omega:')})
                return [self._with_ref(row, refs) for row in rows]
            except sqlite3.OperationalError:
                if cancel and cancel.is_set():
                    raise Cancelled()
                raise

    def search(self, keyword, exact_modbus=False, ref=None, sn=None, lot=None, path=None, cancel=None):
        conditions, params = [], []
        if path:
            conditions.append('f.path=?'); params.append(path)
        if exact_modbus:
            match = re.fullmatch(r'ModBusId=(\d+)', keyword)
            if not match:
                return []
            conditions.append('l.modbus=?'); params.append(match[1])
        elif ref is None:
            normalized = keyword.lower()
            if len(normalized) >= 3:
                # Single trigram terms allow compact detail=none, including punctuation.
                grams = sorted({normalized[i:i + 3] for i in range(len(normalized) - 2)})
                expression = ' AND '.join('"' + gram.replace('"', '""') + '"' for gram in grams)
                conditions.append('l.id IN (SELECT rowid FROM lines_fts WHERE lines_fts MATCH ?)')
                params.append(expression)
            conditions.append('instr(l.normalized,?)>0'); params.append(normalized)
        for column, value in [('ref', ref), ('sn', sn), ('lot', lot)]:
            if value is not None:
                conditions.append(f'l.{column}=?'); params.append(value)
        rows = self._rows("SELECT f.path,l.number,l.text FROM lines l JOIN files f ON f.id=l.file_id WHERE "
                          + ' AND '.join(conditions or ['1']) + ' ORDER BY f.end DESC,f.path DESC,l.number', params, cancel)
        grouped = OrderedDict()
        for row in rows:
            grouped.setdefault(row['path'], []).append((row['number'], row['text']))
        return list(grouped.items())

    def events(self, path, reference, operation=None, after=-1, before=None, reverse=False, loading=False):
        where, args = ['f.path=?', 'l.op_ref=?', 'l.number>?'], [path, reference, after]
        if operation:
            where.append('l.op=?'); args.append(operation)
        if before is not None:
            where.append('l.number<?'); args.append(before)
        if loading:
            where.append('l.loading=1')
        return self._rows('SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id WHERE '
                          + ' AND '.join(where) + ' ORDER BY l.number ' + ('DESC' if reverse else 'ASC'), args)

    def measurement(self, path, reference, get_number):
        rows = self.events(path, reference, after=get_number)
        for row in rows:
            if row['op'] == 'GET' and row['loading']:
                return None
            if row['dims']:
                return row
        return None

    def comparison(self, path, reference):
        gets = self.events(path, reference, 'GET', loading=True)
        current = self.measurement(path, reference, gets[0]['number']) if gets else None
        if not current:
            return None
        metadata = self.metadata()
        paths = [row['path'] for row in metadata]
        for previous in paths[paths.index(path) + 1:]:
            if path_key(os.path.dirname(previous)) != path_key(os.path.dirname(path)):
                continue
            gets = self.events(previous, reference, 'GET', loading=True)
            other = self.measurement(previous, reference, gets[0]['number']) if gets else None
            if other:
                return current, other
        return None

    def differs(self, path, reference):
        pair = self.comparison(path, reference)
        if not pair:
            return False
        first, second = (json.loads(row['dims']) for row in pair)
        return any(abs(first[key] - second[key]) > 2 for key in ('l', 'H', 'L') if key in first and key in second)

    def details(self, path, number, anchor=True, *, location_lot=None):
        selected = self._rows('SELECT l.*,f.start FROM lines l JOIN files f ON f.id=l.file_id WHERE f.path=? AND l.number=?', (path, number))
        if not selected:
            return None
        row = selected[0]; reference = row['ref']
        row['time'] = line_datetime(row['text'], datetime.fromtimestamp(row['start'])).timestamp()
        gets = self.events(path, reference, 'GET', after=number if anchor else -1, loading=True)
        get = gets[0] if gets else None
        putes = self.events(path, reference, 'PUTE', after=get['number']) if get else []
        # With a lot criterion, choose the last PUTE for that exact lot before
        # deduplication; a later operation for another lot must not hide it.
        lot_where = ' AND l.lot_hint=?' if location_lot is not None else ''
        lot_latest = ' AND x.lot_hint=l.lot_hint' if location_lot is not None else ''
        parameters = (reference, location_lot) if location_lot is not None else (reference,)
        locations = self._rows("""SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id
            WHERE l.op_ref=? AND l.op='PUTE'""" + lot_where + """ AND l.number=(SELECT max(x.number) FROM lines x
            WHERE x.file_id=l.file_id AND x.op_ref=l.op_ref AND x.op='PUTE'""" + lot_latest
            + ') ORDER BY l.time DESC', parameters)
        lots = {}
        for location in locations:
            if location['lot_hint']:
                lots.setdefault(location['lot_hint'], location)
        idpk = self.measurement(path, reference, get['number']) if get else None
        return {'row': row, 'path': path, 'get': get, 'pute': putes[0] if putes else None,
                'idpk': idpk, 'locations': list(lots.values()), 'location_lot': location_lot}

    def tracking(self, reference):
        return self._rows("SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id WHERE op='PUTE' AND op_ref=? ORDER BY time DESC", (reference,))

    def robot_errors(self):
        return self._rows("""SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id
            WHERE l.modbus IS NOT NULL AND l.modbus<>'' AND l.modbus NOT GLOB '*[^0-9]*'
            AND l.ref IS NOT NULL AND l.number=(SELECT min(t.number)
            FROM lines t WHERE t.file_id=l.file_id AND t.modbus=l.modbus AND t.ref=l.ref)
            AND NOT EXISTS(SELECT 1 FROM lines g WHERE g.file_id=l.file_id AND g.op_ref=l.ref
            AND g.op='GET' AND g.loading=1 AND g.number>l.number)
            ORDER BY f.end DESC,f.path DESC,l.number""")

    def modbus_rows(self, reference, path=None):
        sql = """SELECT DISTINCT l.*,f.path FROM modbus_values v JOIN lines l ON l.id=v.line_id
                 JOIN files f ON f.id=l.file_id WHERE v.value IN (?,?)"""
        params = [str(reference), '-' + str(reference)]
        if path:
            sql += ' AND f.path=?'; params.append(path)
        return self._rows(sql + ' ORDER BY f.end DESC,f.path DESC,l.number', params)

    def convoyage_steps(self, reference, path=None):
        found, result = set(), []
        for row in self.modbus_rows(reference, path):
            values = parse_modbus_line(row['text'])
            for col, name in CHECKPOINTS.items():
                key = (row['path'], col)
                if key not in found and col < len(values) and values[col] in (reference, '-' + reference):
                    found.add(key)
                    result.append({'horodatage': values[0], 'checkpoint': name,
                                   'etat': 'Présent' if values[col] == reference else 'Jamais arrivé',
                                   'file': trace_label(row['path'])})
        return result

    def convoyage_summary(self, reference, path=None):
        results, positives = OrderedDict(), {}
        for row in self.modbus_rows(reference, path):
            values = parse_modbus_line(row['text'])
            positive = next((col for col in CHECKPOINTS if col < len(values) and values[col] == reference), None)
            negative = next((col for col in CHECKPOINTS if col < len(values) and values[col] == '-' + reference), None)
            col = positive if positive is not None else negative
            if col is None:
                continue
            item = {'HORODATAGE': values[0], '_filepath': row['path'], '_col_idx': col,
                    '_file': trace_label(row['path']), CHECKPOINTS[col]: values[col]}
            if positive is not None:
                item['_checkpoint_actuel'] = CHECKPOINTS[col]
                positives[row['path']] = {'horodatage': values[0], 'checkpoint': CHECKPOINTS[col], 'col_idx': col}
            else:
                item['_checkpoint_attente'] = CHECKPOINTS[col]
                if row['path'] in positives:
                    item['_last_positive'] = positives[row['path']]
            results[row['path']] = item
        return list(results.values())

    def convoyage_recap(self, reference):
        result = {col: {'etat': None, 'horodatage': '', 'file': ''} for col in CHECKPOINTS}
        for row in self.modbus_rows(reference):
            values = parse_modbus_line(row['text'])
            for col in CHECKPOINTS:
                if col >= len(values):
                    continue
                if values[col] == reference or (values[col] == '-' + reference and result[col]['etat'] is None):
                    result[col] = {'etat': 'Présent' if values[col] == reference else 'Jamais arrivé',
                                   'horodatage': values[0], 'file': trace_label(row['path'])}
        return result

    def convoyage_errors(self):
        rows = self._rows("""SELECT v.col,v.value,l.id,f.path FROM modbus_values v
            JOIN lines l ON l.id=v.line_id JOIN files f ON f.id=l.file_id WHERE v.col IN (1,2,3,4,6,7)
            ORDER BY f.end DESC,f.path DESC,l.number,v.col""")
        born, arrived = defaultdict(set), defaultdict(set)
        # First non-zero input column per row is the original detector's rule.
        seen = set()
        for row in rows:
            value = row['value']
            if not value.lstrip('-').isdigit():
                continue
            if row['col'] in (6, 7) and not value.startswith('-'):
                arrived[row['path']].add(value)
            elif row['col'] in (1, 2, 3, 4):
                if row['id'] not in seen:
                    born[row['path']].add(value.lstrip('-'))
                    seen.add(row['id'])
        return [(path, value) for path, values in born.items() for value in sorted(values - arrived[path])]

    def next_files(self):
        rows = self.metadata()
        if self.history is None:
            lookup = {(row['directory'], os.path.basename(row['path']).lower()): row['path'] for row in rows}
            return {lookup[(row['directory'], ntpath.basename(row['previous']).lower())]: row['path']
                    for row in rows if row['previous'] and (row['directory'], ntpath.basename(row['previous']).lower()) in lookup
                    and lookup[(row['directory'], ntpath.basename(row['previous']).lower())] != row['path']}
        lookup = defaultdict(list)
        for row in rows:
            for name in {os.path.basename(alias).lower() for alias in row.get('aliases', [row['path']])}:
                lookup[(row['directory'], name)].append(row)
        links = defaultdict(list)
        for row in rows:
            if not row['previous']:
                continue
            previous_name = ntpath.basename(row['previous']).lower()
            if previous_name in {os.path.basename(alias).lower() for alias in row.get('aliases', [row['path']])}:
                continue
            candidates = [candidate for candidate in lookup[(row['directory'], previous_name)]
                          if candidate['path'] != row['path'] and candidate['start'] <= row['start'] and candidate['end'] < row['end']]
            if candidates:
                latest = max(candidate['end'] for candidate in candidates)
                candidates = [candidate for candidate in candidates if candidate['end'] == latest]
                if len(candidates) == 1:
                    links[candidates[0]['path']].append(row)
        result = {}
        for previous, following in links.items():
            earliest = min(row['start'] for row in following)
            candidates = [row for row in following if row['start'] == earliest]
            if len(candidates) == 1:
                result[previous] = candidates[0]['path']
        return result

    def comparison_for_occurrence(self, path, number, cancel=None):
        if getattr(self._local, 'snapshot', None) is None:
            with self.snapshot():
                return self.comparison_for_occurrence(path, number, cancel)
        self._check_cancel(cancel)
        revision = self.analysis_revision()
        cached = self._analysis_cache.get(('comparison', revision, path, number))
        if cached is not None:
            return cached
        rows = self._rows('''SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id
            WHERE f.path=? AND l.number=?''', (path, number), cancel)
        if not rows or not rows[0]['ref']:
            return dict(status='unavailable', different=False, current=None, pair=None,
                        message='Occurrence indisponible.', candidates=[])
        return compare_tickets(self, rows, revision, cancel)[(path, number)]

    @staticmethod
    def _neighbor_counts(before_count, after_count):
        if any(type(n) is not int or not 0 <= n <= 1000 for n in (before_count, after_count)):
            raise ValueError('Les nombres de voisins doivent être des entiers entre 0 et 1000.')

    def _ticket_rows(self, path, revision, cancel=None):
        key = ('tickets', revision, path)
        rows = self._analysis_cache.get(key)
        if rows is None:
            rows = self._rows('''SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id
                WHERE f.path=? AND modbus IS NOT NULL AND ref IS NOT NULL ORDER BY number''', (path,), cancel)
            self._analysis_cache.put(key, rows, len(rows))
        return rows

    def adjacent(self, path, target, number=None, reference=None, *, before_count=5, after_count=5):
        self._neighbor_counts(before_count, after_count)
        neighbors = self._collect_adjacent(path, target, number, reference, before_count, after_count, self.analysis_revision())
        if neighbors is None:
            return None
        previous, following, reference, number = neighbors
        differences = {(file, ref): self.differs(file, ref) for _, ref, file, _ in previous + following + [(str(target), reference, path, number)]}
        return previous, following, reference, differences, number

    def _collect_adjacent(self, path, target, number, reference, before_count, after_count, revision, cancel=None):
        target = str(target)
        if number is None:
            hits = self.search('ModBusId=' + target, True, path=path)
            if not hits:
                return None
            number, text = next(((n, text) for n, text in hits[0][1] if parse_fields(text).get('RefProduit')), (None, None))
        else:
            selected = self._rows("SELECT l.* FROM lines l JOIN files f ON f.id=l.file_id WHERE f.path=? AND l.number=?", (path, number))
            if not selected or selected[0]['modbus'] != target or not selected[0]['ref']:
                return None
            if reference is not None and selected[0]['ref'] != reference:
                return None
            text = selected[0]['text']
        if number is None:
            return None
        rows = self._ticket_rows(path, revision, cancel)
        def collect(items, file, limit=5, exclude=()):
            result, seen = [], {target, *exclude}
            if not limit:
                return result
            for item in items:
                self._check_cancel(cancel)
                if item['modbus'].isdigit() and item['ref'] and item['modbus'] not in seen:
                    result.append((item['modbus'], item['ref'], file, item['number'])); seen.add(item['modbus'])
                    if len(result) == limit:
                        break
            return result
        previous = collect(reversed([row for row in rows if row['number'] < number]), path, before_count)
        following = collect([row for row in rows if row['number'] > number], path, after_count)
        links = self._analysis_cache.get(('next_files', revision))
        if links is None:
            links = self._analysis_cache.put(('next_files', revision), self.next_files())
        next_path = links.get(path)
        visited = {path}
        while len(following) < after_count and next_path and next_path not in visited:
            visited.add(next_path)
            more = self._ticket_rows(next_path, revision, cancel)
            following.extend(collect(more, next_path, after_count - len(following), {item[0] for item in following}))
            next_path = links.get(next_path)
        reference = parse_fields(text)['RefProduit']
        return previous, following, reference, number

    def paired_adjacent(self, path, target, number, reference=None, cancel=None, *, before_count=5, after_count=5):
        """Two independent robot neighborhoods, anchored to the selected ticket's time.

        The index supplies all data, including deleted archives. No cross-source
        ID matching and no arbitrary choice among equally dated generations.
        """
        if getattr(self._local, 'snapshot', None) is None:
            with self.snapshot():
                return self.paired_adjacent(path, target, number, reference, cancel,
                                            before_count=before_count, after_count=after_count)
        self._check_cancel(cancel)
        self._neighbor_counts(before_count, after_count)
        revision = self.analysis_revision()
        cache_key = ('paired', revision, path, str(target), number, reference, before_count, after_count)
        cached = self._analysis_cache.get(cache_key)
        if cached is not None:
            return cached
        selected = self._rows('''SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id
            WHERE f.path=? AND l.number=?''', (path, number), cancel)
        if not selected or selected[0]['modbus'] != str(target) or not selected[0]['ref']:
            return None
        selected = selected[0]
        if reference is not None and selected['ref'] != reference:
            return None
        scope, robot = robot_scope(path)
        def timed(row):
            return analysis_timed(self, row, revision)

        def column(row, name):
            self._check_cancel(cancel)
            neighbors = self._collect_adjacent(row['path'], row['modbus'], row['number'], row['ref'],
                                               before_count, after_count, revision, cancel)
            if neighbors is None:
                return {'robot': name, 'message': 'Occurrence indisponible.', 'anchor': None}
            before, after, _, _ = neighbors
            identities = before + [(row['modbus'], row['ref'], row['path'], row['number'])] + after
            by_file = defaultdict(list)
            for _, _, file, n in identities:
                by_file[file].append(n)
            tickets = {}
            for file, numbers in by_file.items():
                needed = set(numbers)
                for item in self._ticket_rows(file, revision, cancel):
                    if item['number'] in needed:
                        tickets[(item['path'], item['number'])] = dict(item)
            def ticket(identity):
                return tickets[(identity[2], identity[3])]
            return {'robot': name, 'anchor': tickets[(row['path'], row['number'])],
                    'previous': [ticket(item) for item in reversed(before)],
                    'following': [ticket(item) for item in after], 'candidates': [], 'delta': None}

        own = column(selected, robot or 'Trace sélectionnée')
        has_time = timed(selected)
        result = {'selected': selected, 'selected_robot': robot,
                  'time': selected['time'] if has_time else None, 'columns': []}
        other = {'robot': 'R2' if robot == 'R1' else 'R1' if robot == 'R2' else 'Autre robot',
                 'anchor': None, 'candidates': []}
        if not robot:
            other['message'] = 'Robot non reconnu : rapprochement R1/R2 impossible.'
        elif not has_time:
            other['message'] = 'Date ou heure inexploitable dans la trace sélectionnée.'
        else:
            metas = [meta for meta in analysis_metadata(self, revision) if meta['kind'] == 'omega'
                     and robot_scope(meta['path']) == (scope, other['robot'])]
            from occurrence_comparison import dated
            files = [meta['path'] for meta in metas if dated(self, meta['path'], revision)]
            if not metas:
                other['message'] = 'Aucune trace de ce robot dans cet emplacement.'
            elif not files:
                other['message'] = 'Aucune trace de ce robot avec une date exploitable.'
            else:
                # Two bounded chronological lookups per generation, no global ABS sort.
                closest = []
                best = None
                for file in files:
                    for relation, direction in (('<=', 'DESC'), ('>=', 'ASC')):
                        offset = 0
                        while True:
                            rows = self._rows('''SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id
                                WHERE f.path=? AND l.modbus IS NOT NULL AND l.ref IS NOT NULL
                                AND l.modbus<>'' AND l.modbus NOT GLOB '*[^0-9]*' AND l.time ''' + relation + ''' ?
                                ORDER BY l.time ''' + direction + ',l.number LIMIT 32 OFFSET ?',
                                (file, selected['time'], offset), cancel)
                            row = next((r for r in rows if timed(r)), None)
                            if row is not None:
                                key = (round(abs(row['time'] - selected['time']), 6), row['time'])
                                if best is None or key < best:
                                    closest, best = [], key
                                if key == best and all(item['path'] != row['path'] for item in closest):
                                    closest.append(row)
                                break
                            if len(rows) < 32:
                                break
                            offset += len(rows)
                if not closest:
                    other['message'] = 'Aucune occurrence avec une date et une heure exploitables.'
                elif len(closest) > 1:
                    other.update(candidates=closest, delta=closest[0]['time'] - selected['time'],
                                 message='Plusieurs générations au même instant : choisissez une occurrence.')
                else:
                    other = column(closest[0], other['robot'])
                    other['delta'] = closest[0]['time'] - selected['time']
        own['delta'] = 0 if has_time else None
        if not has_time:
            own['message'] = 'Date ou heure inexploitable ; voisins dans la trace sélectionnée.'
        result['columns'] = ([own, other] if robot != 'R2' else [other, own])
        tickets = []
        for col in result['columns']:
            if col.get('anchor'):
                tickets.extend(col['previous'] + [col['anchor']] + col['following'])
            tickets.extend(col.get('candidates', []))
        flags = compare_tickets(self, tickets, revision, cancel)
        for row in tickets:
            row['comparison'] = flags[(row['path'], row['number'])]
            row['different'] = row['comparison']['different']
        result.update(revision=revision, before_count=before_count, after_count=after_count)
        return self._analysis_cache.put(cache_key, result, len(tickets) + 1)

    def file_lines(self, path):
        with self.connection() as db:
            meta = db.execute('SELECT mtime,size FROM files WHERE path=?', (path,)).fetchone()
            if not meta:
                return []
            if isinstance(path, TraceRef):
                with self._cache_lock:
                    cached = self.text_cache.get(path)
                    if cached and cached[0] == tuple(meta):
                        self.text_cache.move_to_end(path)
                        return cached[1]
                lines = [row[0] for row in db.execute('SELECT text FROM lines WHERE file_id=(SELECT id FROM files WHERE path=?) ORDER BY number', (path,))]
                with self._cache_lock:
                    self.text_cache[path] = tuple(meta), lines
                    while len(self.text_cache) > 3:
                        self.text_cache.popitem(last=False)
                return lines
            try:
                stat = os.stat(path)
            except FileNotFoundError:
                return []
            key = (stat.st_mtime_ns, stat.st_size)
            with self._cache_lock:
                cached = self.text_cache.get(path)
                if cached and cached[0] == key:
                    self.text_cache.move_to_end(path)
                    return cached[1]
            if key == tuple(meta):
                lines = [row[0] for row in db.execute('SELECT text FROM lines WHERE file_id=(SELECT id FROM files WHERE path=?) ORDER BY number', (path,))]
            else:
                lines = read_trace(path).splitlines()
        with self._cache_lock:
            self.text_cache[path] = key, lines
            self.text_cache.move_to_end(path)
            while len(self.text_cache) > 3:
                self.text_cache.popitem(last=False)
        return lines
