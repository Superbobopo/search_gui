"""Compare store versions on fresh databases and disposable trace copies.

The application cache and original traces are never opened for writing. Run each
version separately so that both measurements start from an empty SQLite cache.
"""
import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import shutil
import sqlite3
import sys
import tempfile
import time

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
from tools.benchmark_traces import memory_mb
from trace_sources import discover_trace_files


def load_store(path):
    name = '_benchmark_store_' + hashlib.sha256(str(path).encode()).hexdigest()[:12]
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def fingerprint(paths):
    return {str(path): [path.stat().st_size, path.stat().st_mtime_ns] for path in paths}


def storage(database):
    sizes = {path.name: path.stat().st_size for path in database.parent.glob(database.name + '*')}
    result = {'files_bytes': sizes, 'total_bytes': sum(sizes.values())}
    with sqlite3.connect(database) as db:
        result['page_size'] = db.execute('PRAGMA page_size').fetchone()[0]
        result['page_count'] = db.execute('PRAGMA page_count').fetchone()[0]
        result['freelist_count'] = db.execute('PRAGMA freelist_count').fetchone()[0]
        result['rows'] = db.execute('SELECT count(*) FROM lines').fetchone()[0]
        columns = {row[1] for row in db.execute('PRAGMA table_info(lines)')}
        for column in ('text', 'normalized'):
            if column in columns:
                result[column + '_logical_utf8_bytes'] = db.execute(
                    f'SELECT sum(length(CAST({column} AS BLOB))) FROM lines').fetchone()[0]
        try:
            result['dbstat_bytes'] = dict(db.execute('SELECT name,sum(pgsize) FROM dbstat GROUP BY name'))
        except sqlite3.OperationalError as error:
            result['dbstat_unavailable'] = str(error)
    return result


def run(args):
    output = args.output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    database = output.with_suffix('.sqlite3')
    if output.exists() or any(database.parent.glob(database.name + '*')):
        raise SystemExit('Refusing to reuse an existing report/database: ' + str(output))
    originals = [Path(path) for path in discover_trace_files()]
    if not originals:
        raise SystemExit('No real trace files found; no application cache was touched.')
    before = fingerprint(originals)
    report = {
        'store_module': str(args.store.resolve()), 'python': sys.version,
        'sqlite': sqlite3.sqlite_version, 'source_files': len(originals),
        'source_bytes': sum(size for size, mtime in before.values()),
        'source_manifest': before, 'database_was_new': True,
        'first_index_is_os_cold': False,
        'note': 'Fresh SQLite database; Windows filesystem cache was not flushed. All sources are disposable copies.',
        'memory_initial_mb': memory_mb(), 'status': 'started',
    }
    def save():
        output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf8')
    save()
    module = load_store(args.store.resolve())
    with tempfile.TemporaryDirectory(prefix='tracomega-benchmark-') as temporary:
        copied = []
        copy_start = time.perf_counter()
        for original in originals:
            group = hashlib.sha256(str(original.parent).lower().encode()).hexdigest()[:16]
            target = Path(temporary) / group / original.name
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(original, target)
            copied.append(str(target))
        report['copy_seconds_excluded_from_index'] = time.perf_counter() - copy_start
        start = time.perf_counter()
        store = module.TraceStore(database)
        report['new_database_open_seconds'] = time.perf_counter() - start
        progress_rows = []
        def progress(done, total, path, published):
            item = {'done': done, 'total': total, 'file': Path(path).name,
                    'published': bool(published), 'seconds': time.perf_counter() - start,
                    'memory_mb': memory_mb()}
            progress_rows.append(item)
            report['first_index_progress'] = progress_rows
            save()
            print(f"{done}/{total} {Path(path).name}: {item['seconds']:.1f}s RAM={item['memory_mb']}MB", flush=True)
        start = time.perf_counter()
        result = store.sync(copied, progress)
        report['first_index_seconds'] = time.perf_counter() - start
        report['first_index_result'] = result
        report['first_index_memory_mb'] = memory_mb()
        report['storage'] = storage(database)
        report['cache_to_source_ratio'] = report['storage']['total_bytes'] / report['source_bytes']
        save()
        queries = [
            ('id_present', 'ModBusId=1194', True),
            ('reference_present', 'RefProduit=A0061574AC', False),
            ('id_absent', 'ModBusId=99999999999999', True),
            ('reference_absent', 'RefProduit=REF_ABSENTE_BENCH_20261017', False),
        ]
        report['queries'] = {}
        for label, keyword, exact in queries:
            samples = []
            for _ in range(5):
                start = time.perf_counter()
                with store.snapshot():
                    groups = store.search(keyword, exact)
                    searched = time.perf_counter()
                    details = store.details(groups[0][0], groups[0][1][0][0]) if groups else None
                    detailed = time.perf_counter()
                samples.append({'search_ms': (searched - start) * 1000,
                                'details_ms': (detailed - searched) * 1000,
                                'total_ms': (time.perf_counter() - start) * 1000,
                                'files': len(groups), 'occurrences': sum(len(rows) for _, rows in groups),
                                'details_found': details is not None})
            report['queries'][label] = {'keyword': keyword, 'samples': samples}
        start = time.perf_counter()
        reopened = module.TraceStore(database)
        report['reopen_constructor_seconds'] = time.perf_counter() - start
        start = time.perf_counter()
        report['reopen_sync_result'] = reopened.sync(copied)
        report['unchanged_sync_seconds'] = time.perf_counter() - start
        report['reopen_memory_mb'] = memory_mb()
        save()
        active = [Path(path) for path in copied if Path(path).name.lower() == 'tracomega.txt']
        target = max(active or [Path(path) for path in copied if Path(path).name.lower().startswith('tracomega')], key=lambda path: path.stat().st_size)
        marker = 'BENCH_APPEND_20261017'
        addition = (f'\n  17|10:00:00.0 RefProduit={marker}<Cr>ModBusId=987654321<Cr>\n'
                    f'G 17|10:00:01.0 GRIPPER GET 1x boite(s): Code:{marker} $0101\n'
                    f'G 17|10:00:02.0 GRIPPER PUTE 1x boite(s): Code:{marker}\n').encode('ascii')
        report['append'] = {'original_copy_bytes': target.stat().st_size, 'added_bytes': len(addition), 'file': target.name}
        with target.open('ab') as stream:
            stream.write(addition)
        start = time.perf_counter()
        report['append']['sync_result'] = reopened.sync(copied)
        report['append']['seconds'] = time.perf_counter() - start
        with reopened.snapshot():
            matches = reopened.search('ModBusId=987654321', True, path=str(target))
            assert matches and len(matches[0][1]) == 1, 'Appended ticket missing or duplicated'
            detail = reopened.details(matches[0][0], matches[0][1][0][0])
            assert detail['get'] is not None and detail['pute'] is not None, 'Appended event details missing'
        report['append']['ticket_and_events_verified'] = True
        report['memory_final_mb'] = memory_mb()
    report['source_manifest_after'] = fingerprint(originals)
    report['original_sources_unchanged'] = before == report['source_manifest_after']
    assert report['original_sources_unchanged'], 'Real sources changed during the benchmark'
    report['status'] = 'complete'
    save()
    print(json.dumps({key: report[key] for key in ('status', 'source_files', 'source_bytes', 'first_index_seconds', 'cache_to_source_ratio', 'unchanged_sync_seconds', 'append', 'original_sources_unchanged')}, ensure_ascii=False, indent=2), flush=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--store', type=Path, default=PROJECT / 'trace_store.py')
    parser.add_argument('--output', type=Path, required=True)
    run(parser.parse_args())


if __name__ == '__main__':
    main()
