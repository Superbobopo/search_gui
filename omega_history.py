"""Authoritative OMEGA snapshots. Never delete or rebuild this database.

The search index is a disposable projection. Compressed raw chunks here are
committed first and are sufficient to recreate it without the source files.
"""

from contextlib import contextmanager, closing
from collections import defaultdict
from datetime import datetime
import hashlib
import json
import os
from pathlib import Path
import sqlite3
import tempfile
from types import SimpleNamespace
import uuid
import zlib

from trace_sources import application_directory, path_key, trace_directories
from trace_parser import trace_start_datetime, line_datetime
from trace_identity import TraceRef


HISTORY_VERSION = 1
HISTORY_APPLICATION_ID = 0x4F4D4547
CHUNK = 1024 * 1024


class HistoryUnavailable(sqlite3.OperationalError):
    """Stop projection when the authoritative history cannot be committed."""


def history_path():
    identity = hashlib.sha256(path_key(application_directory()).encode()).hexdigest()[:16]
    root = Path(os.environ.get('LOCALAPPDATA', Path.home() / '.cache'))
    return root / 'TracOmega' / 'history' / identity / 'omega_history.sqlite3'


def under(path, root):
    try:
        return os.path.commonpath((path_key(path), path_key(root))) == path_key(root)
    except ValueError:
        return False


def historical_roots(database=None):
    """Read only the source catalogue before source choice, never open an index."""
    database = Path(database) if database else history_path()
    if not database.exists():
        return set()
    with closing(sqlite3.connect(database.resolve().as_uri() + '?mode=ro', uri=True)) as db:
        if db.execute('PRAGMA user_version').fetchone()[0] != HISTORY_VERSION:
            raise RuntimeError('Version de l’historique OMEGA incompatible ; fichier conservé.')
        return {row[0] for row in db.execute('SELECT DISTINCT root FROM versions')}


class OmegaHistory:
    def __init__(self, database):
        self.database = Path(database)
        self.database.parent.mkdir(parents=True, exist_ok=True)
        with self.connection() as db:
            version = db.execute('PRAGMA user_version').fetchone()[0]
            tables = db.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
            if ((version != HISTORY_VERSION and (version != 0 or tables))
                    or (version == HISTORY_VERSION and 'versions' not in {row[0] for row in tables})):
                raise RuntimeError('Version de l’historique OMEGA incompatible ; fichier conservé.')
            # This initial schema needs no migration. Future migrations MUST use
            # backup_before_migration before changing any schema or stored data.
            db.execute('PRAGMA journal_mode=WAL')
            db.executescript('''
                CREATE TABLE IF NOT EXISTS versions (
                    id TEXT PRIMARY KEY, root TEXT NOT NULL, directory TEXT NOT NULL,
                    path TEXT NOT NULL, encoding TEXT NOT NULL, size INTEGER NOT NULL,
                    mtime INTEGER NOT NULL, device INTEGER NOT NULL, inode INTEGER NOT NULL,
                    digest TEXT NOT NULL, canonical TEXT NOT NULL, head TEXT NOT NULL,
                    start REAL NOT NULL, end REAL NOT NULL, revision INTEGER NOT NULL,
                    source_encoding TEXT NOT NULL, line_count INTEGER NOT NULL,
                    normalized INTEGER NOT NULL);
                CREATE INDEX IF NOT EXISTS history_digest ON versions(directory,digest);
                CREATE INDEX IF NOT EXISTS history_canonical ON versions(directory,canonical);
                CREATE TABLE IF NOT EXISTS chunks (
                    version TEXT REFERENCES versions(id), number INTEGER, data BLOB NOT NULL,
                    PRIMARY KEY(version,number));
                CREATE TABLE IF NOT EXISTS locations (
                    path TEXT PRIMARY KEY, original TEXT NOT NULL, version TEXT NOT NULL,
                    size INTEGER, mtime INTEGER, device INTEGER, inode INTEGER,
                    canonical TEXT NOT NULL,
                    present INTEGER NOT NULL);
                CREATE TABLE IF NOT EXISTS aliases (
                    version TEXT NOT NULL, path TEXT NOT NULL, PRIMARY KEY(version,path));
                CREATE TABLE IF NOT EXISTS imports (
                    cache TEXT NOT NULL, file INTEGER NOT NULL, signature TEXT NOT NULL,
                    PRIMARY KEY(cache,file));
                CREATE TABLE IF NOT EXISTS captures (
                    directory TEXT NOT NULL, digest TEXT NOT NULL, canonical TEXT NOT NULL,
                    version TEXT NOT NULL REFERENCES versions(id), PRIMARY KEY(directory,digest));
                CREATE INDEX IF NOT EXISTS capture_text ON captures(directory,canonical);
            ''')
            db.execute(f'PRAGMA user_version={HISTORY_VERSION}')
            db.execute(f'PRAGMA application_id={HISTORY_APPLICATION_ID}')
            db.commit()

    @contextmanager
    def connection(self):
        db = sqlite3.connect(self.database, timeout=30)
        db.row_factory = sqlite3.Row
        try:
            db.execute('PRAGMA foreign_keys=ON')
            db.execute('PRAGMA synchronous=FULL')
            yield db
        except sqlite3.DatabaseError as error:
            raise HistoryUnavailable('Historique OMEGA non actualisé ; nouvelles traces non archivées. '
                                     'Base conservée. ' + str(error)) from error
        finally:
            db.close()

    def backup_before_migration(self):
        destination = self.database.with_name(self.database.name + '.' + uuid.uuid4().hex + '.bak')
        with self.connection() as source, closing(sqlite3.connect(destination)) as backup:
            source.backup(backup)
        return destination

    def locations(self):
        with self.connection() as db:
            return {row['path']: dict(row) for row in db.execute('SELECT * FROM locations')}

    def mark_presence(self, paths, roots):
        present = {path_key(path) for path in paths}
        with self.connection() as db, db:
            for row in db.execute('SELECT path,present FROM locations').fetchall():
                if any(under(row['path'], root) for root in roots):
                    value = int(row['path'] in present)
                    if value != row['present']:
                        db.execute('UPDATE locations SET present=? WHERE path=?', (value, row['path']))

    def entries(self, roots):
        with self.connection() as db:
            rows = [dict(row) for row in db.execute('SELECT * FROM versions')
                    if any(under(row['directory'], root) for root in roots)]
            present = {(r[0], r[1]) for r in db.execute('SELECT version,canonical FROM locations WHERE present=1')}
            aliases = defaultdict(list)
            selected = {row['id'] for row in rows}
            for version, path in db.execute('SELECT version,path FROM aliases ORDER BY version,path'):
                if version in selected:
                    aliases[version].append(path)
            for row in rows:
                row['present'] = (row['id'], row['canonical']) in present
                row['aliases'] = aliases[row['id']]
            return rows

    @staticmethod
    def _record_location(db, path, version, canonical, stat, live):
        original = os.path.abspath(path)
        if not db.execute('SELECT 1 FROM aliases WHERE version=? AND path=?', (version, original)).fetchone():
            db.execute('INSERT OR IGNORE INTO aliases VALUES(?,?)', (version, original))
        if live:
            values = (path_key(path), original, version, stat.st_size, stat.st_mtime_ns,
                      stat.st_dev, stat.st_ino, canonical, 1)
            previous = db.execute('SELECT * FROM locations WHERE path=?', (values[0],)).fetchone()
            if previous is None or tuple(previous) != values:
                db.execute('INSERT OR REPLACE INTO locations VALUES(?,?,?,?,?,?,?,?,?)', values)

    def recognize(self, path, stat, digest, check=lambda: None, live=True):
        """Reuse a raw capture before decoding, parsing or touching its chunks."""
        with self.connection() as db, db:
            row = db.execute('SELECT version,canonical FROM captures WHERE directory=? AND digest=?',
                             (path_key(os.path.dirname(path)), digest)).fetchone()
            if row is None:
                return None
            self._record_location(db, path, row['version'], row['canonical'], stat, live)
            check()
            return row['version']

    @contextmanager
    def snapshot(self, version, revision=None):
        with tempfile.TemporaryFile(dir=self.database.parent) as stream:
            with self.connection() as db:
                db.execute('BEGIN')
                metadata = db.execute('SELECT size,digest,revision FROM versions WHERE id=?', (version,)).fetchone()
                if metadata is None or (revision is not None and metadata['revision'] != revision):
                    raise OSError('Historique OMEGA actualisé pendant la lecture ; réessayez.')
                digest = hashlib.sha256()
                for row in db.execute('SELECT data FROM chunks WHERE version=? ORDER BY number', (version,)):
                    data = zlib.decompress(row[0])
                    digest.update(data)
                    stream.write(data)
                if stream.tell() != metadata['size'] or digest.hexdigest() != metadata['digest']:
                    raise OSError('Contenu de l’historique OMEGA invalide ; base conservée sans modification.')
                db.rollback()
            stream.seek(0)
            yield stream

    def publish(self, path, stream, stat, encoding, digest, root, check=lambda: None, live=True, source_encoding=None):
        # Import locally to keep discovery free of an import cycle.
        from trace_store import _physical_lines, _decode_physical
        recognized = self.recognize(path, stat, digest, check, live)
        if recognized is not None:
            return recognized
        canonical = hashlib.sha256()
        header, count, start, current = [], 0, None, None
        fallback = datetime.fromtimestamp(stat.st_mtime)
        for offset, raw in _physical_lines(stream):
            check()
            for line in _decode_physical(raw, encoding, offset).splitlines():
                canonical.update(line.encode('utf-8') + b'\n')
                if start is None:
                    header.append(line)
                    if len(header) == 8:
                        start = trace_start_datetime('\n'.join(header), fallback)
                        current = start
                        for record in header:
                            current = line_datetime(record, current)
                else:
                    current = line_datetime(line, current)
                count += 1
        head = '\n'.join(header)
        if start is None:
            start = trace_start_datetime(head, fallback)
            current = start
            for record in header:
                current = line_datetime(record, current)
        directory, canonical = path_key(os.path.dirname(path)), canonical.hexdigest()
        with self.connection() as db, db:
            # Serialize capture decisions across concurrent application instances.
            db.execute('BEGIN IMMEDIATE')
            exact = db.execute('SELECT v.* FROM versions v JOIN captures c ON c.version=v.id WHERE c.directory=? AND (c.digest=? OR c.canonical=?) ORDER BY v.revision DESC',
                               (directory, digest, canonical)).fetchall()
            version = exact[0]['id'] if exact else None
            candidate = None
            if not version:
                location = db.execute('SELECT v.* FROM versions v JOIN locations l ON l.version=v.id WHERE l.path=?', (path_key(path),)).fetchone()
                candidates = [location] if location else db.execute('SELECT * FROM versions WHERE directory=? AND head=? AND size<?', (directory, head, stat.st_size)).fetchall()
                compatible = []
                for previous in candidates:
                    if previous['size'] >= stat.st_size and not previous['normalized']:
                        continue
                    stream.seek(0)
                    prefix, remaining = hashlib.sha256(), previous['size']
                    while remaining:
                        check()
                        data = stream.read(min(CHUNK, remaining))
                        if not data:
                            break
                        prefix.update(data)
                        remaining -= len(data)
                    if not remaining and prefix.hexdigest() == previous['digest']:
                        compatible.append(previous)
                    elif previous['normalized']:
                        # Old caches saved decoded lines, not raw newlines/BOM.
                        # Match their complete original records before extending.
                        prefix = hashlib.sha256()
                        records = 0
                        for offset, raw in _physical_lines(stream):
                            check()
                            for line in _decode_physical(raw, encoding, offset).splitlines():
                                if records < previous['line_count']:
                                    prefix.update(line.encode('utf-8') + b'\n')
                                    records += 1
                            if records >= previous['line_count']:
                                break
                        if records == previous['line_count'] and prefix.hexdigest() == previous['canonical']:
                            compatible.append(previous)
                if len(compatible) == 1:
                    candidate = compatible[0]
                    version = candidate['id']
            if exact and live and exact[0]['canonical'] == canonical and exact[0]['digest'] != digest:
                # Same complete text copied with another encoding/newlines.
                # Refresh its raw representation so the next append can resume.
                candidate, exact = exact[0], []
            if version is None:
                version = uuid.uuid4().hex
            if not exact:
                revision = candidate['revision'] + 1 if candidate else 1
                db.execute('''INSERT INTO versions VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    ON CONFLICT(id) DO UPDATE SET path=excluded.path, encoding=excluded.encoding,
                    size=excluded.size,mtime=excluded.mtime,device=excluded.device,inode=excluded.inode,
                    digest=excluded.digest,canonical=excluded.canonical,head=excluded.head,
                    start=excluded.start,end=excluded.end,revision=excluded.revision,
                    source_encoding=excluded.source_encoding,line_count=excluded.line_count,
                    normalized=excluded.normalized''',
                    (version, path_key(root), directory, os.path.abspath(path), encoding, stat.st_size,
                     stat.st_mtime_ns, stat.st_dev, stat.st_ino, digest, canonical, head,
                     start.timestamp(), current.timestamp(), revision, source_encoding or encoding, count, int(not live)))
                db.execute('DELETE FROM chunks WHERE version=?', (version,))
                stream.seek(0)
                number = 0
                while data := stream.read(CHUNK):
                    check()
                    db.execute('INSERT INTO chunks VALUES(?,?,?)', (version, number, zlib.compress(data)))
                    number += 1
            db.execute('INSERT OR IGNORE INTO captures VALUES(?,?,?,?)', (directory, digest, canonical, version))
            self._record_location(db, path, version, canonical, stat, live)
            check()
        return version

    def migrate_caches(self, databases, roots, check=lambda: None):
        """Read legacy rows, including vanished files, before any cache rebuild."""
        for database in dict.fromkeys(Path(p) for p in databases):
            if not database.is_file():
                continue
            cache = str(database.resolve())
            with self.connection() as db:
                imported = {row[0]: row[1] for row in db.execute('SELECT file,signature FROM imports WHERE cache=?', (cache,))}
            with closing(sqlite3.connect(database.resolve().as_uri() + '?mode=ro', uri=True)) as source:
                source.row_factory = sqlite3.Row
                if not source.execute("SELECT 1 FROM sqlite_master WHERE name='files'").fetchone():
                    continue
                for row in source.execute("SELECT * FROM files WHERE kind='omega'"):
                    check()
                    path = row['path']
                    if path.startswith('omega:'):
                        continue
                    matching = [root for root in roots if under(path, root)]
                    if not matching:
                        continue
                    signature = json.dumps([row['size'], row['mtime'], row['line_count'], row['end']])
                    if imported.get(row['id']) == signature:
                        continue
                    with tempfile.TemporaryFile(dir=self.database.parent) as stream:
                        digest = hashlib.sha256()
                        for line in source.execute('SELECT text FROM lines WHERE file_id=? ORDER BY number', (row['id'],)):
                            check()
                            data = (line[0] + '\n').encode('utf-8')
                            stream.write(data)
                            digest.update(data)
                        stat = SimpleNamespace(st_size=stream.tell(), st_mtime_ns=row['mtime'],
                                               st_mtime=row['mtime']/1e9, st_dev=0, st_ino=0)
                        self.publish(path, stream, stat, 'utf-8', digest.hexdigest(), matching[0], check,
                                     live=False, source_encoding=row['encoding'])
                    with self.connection() as db, db:
                        db.execute('INSERT OR REPLACE INTO imports VALUES(?,?,?)', (cache, row['id'], signature))
