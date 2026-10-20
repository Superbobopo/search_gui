"""Compare history-aware startup/indexing on identical disposable real copies."""

import argparse
from datetime import datetime
import hashlib
import json
from pathlib import Path
import shutil
import statistics
import sys
import tempfile
import time


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--module-dir', type=Path)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    project = Path(__file__).resolve().parents[1]
    sys.path.insert(0, str(project))
    if args.module_dir:
        sys.path.insert(0, str(args.module_dir.resolve()))
    import trace_store
    from trace_sources import discover_trace_files
    from tools.benchmark_traces import memory_mb
    output = args.output.resolve()
    if output.exists():
        raise SystemExit('Rapport déjà présent ; choisir un nouveau chemin.')
    output.parent.mkdir(parents=True, exist_ok=True)
    originals = sorted(Path(p) for p in discover_trace_files()
                       if Path(p).parent.name.lower() in ('r1', 'r2', 'trace r1', 'trace r2')
                       and Path(p).name.lower() in ('tracomega.txt', 'tracmodbus.txt'))
    if len(originals) != 4:
        raise SystemExit('Ce comparatif exige les quatre fichiers courants R1/R2 OMEGA/MODBUS.')
    manifest = {str(p): (p.stat().st_size, p.stat().st_mtime_ns) for p in originals}
    report = dict(status='started', source_bytes=sum(v[0] for v in manifest.values()),
                  source_files=len(originals), store=str(Path(trace_store.__file__).resolve()),
                  os_cache_cold=False, memory_initial_mb=memory_mb())

    def save():
        output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')

    def storage(database):
        return sum(p.stat().st_size for p in database.parent.glob(database.name + '*') if p.is_file())

    save()
    with tempfile.TemporaryDirectory(prefix='omega-startup-') as folder:
        base = Path(folder)
        root = base / 'traces'
        copies = []
        for original in originals:
            target = root / original.parent.name / original.name
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(original, target)
            assert hashlib.sha256(target.read_bytes()).digest() == hashlib.sha256(original.read_bytes()).digest()
            copies.append(str(target))
        database, archive = base / 'index.db', base / 'history.db'
        before = time.perf_counter()
        store = trace_store.TraceStore(database, directories=[root], history_database=archive)
        report['new_open_seconds'] = time.perf_counter() - before
        before = time.perf_counter()
        result = store.sync(copies, lambda done, total, path, published:
                            print(f'{done}/{total} {Path(path).name}', flush=True) if published else None)
        report['first_index_seconds'] = time.perf_counter() - before
        assert not result['failures'], result
        report['first_index_memory_mb'] = memory_mb()
        report['cache_bytes'] = storage(database)
        report['archive_bytes'] = storage(archive)
        with store.connection() as db:
            ticket = db.execute("SELECT modbus,ref FROM lines WHERE modbus IS NOT NULL AND ref IS NOT NULL LIMIT 1").fetchone()
        assert ticket
        queries = [('id_present', 'ModBusId=' + ticket['modbus'], True),
                   ('reference_present', 'RefProduit=' + ticket['ref'], False),
                   ('id_absent', 'ModBusId=999999999999999', True),
                   ('reference_absent', 'RefProduit=ABSENT_PERF_999999999999999', False)]
        report['queries'] = {}
        for name, text, exact in queries:
            searches, details = [], []
            fingerprint = hashlib.sha256()
            for iteration in range(5):
                start = time.perf_counter()
                groups = store.search(text, exact)
                searches.append((time.perf_counter() - start) * 1000)
                start = time.perf_counter()
                data = store.details(groups[0][0], groups[0][1][0][0]) if groups else None
                details.append((time.perf_counter() - start) * 1000)
                if iteration == 0:
                    for ref, hits in groups:
                        fingerprint.update(str(Path(ref).relative_to(root)).encode())
                        for number, line in hits:
                            fingerprint.update(f'{number}:{line}\n'.encode('utf-8'))
                    if data:
                        for key in ('get', 'pute', 'idpk'):
                            event = data[key]
                            if event:
                                fingerprint.update(f"{key}:{event['number']}:{event['text']}".encode('utf-8'))
            report['queries'][name] = dict(search_ms=statistics.median(searches),
                                           details_ms=statistics.median(details), groups=len(groups), digest=fingerprint.hexdigest())
        before = time.perf_counter()
        store = trace_store.TraceStore(database, directories=[root], history_database=archive)
        metadata = store.metadata()
        report['reopen_cache_available_seconds'] = time.perf_counter() - before
        report['reopen_cached_files'] = len(metadata)
        before = time.perf_counter()
        result = store.sync(copies)
        report['reopen_verification_seconds'] = time.perf_counter() - before
        assert not result['failures'] and not result['changed'], result
        before = time.perf_counter()
        result = store.sync(copies)
        report['steady_sync_seconds'] = time.perf_counter() - before
        assert not result['changed'], result
        omega = next(p for p in copies if Path(p).parent.name.lower() in ('r2', 'trace r2') and Path(p).name.lower() == 'tracomega.txt')
        day = datetime.fromtimestamp(max(r['end'] for r in store.metadata())).day
        with open(omega, 'ab') as stream:
            stream.write(f' {day}|23:59:59 RefProduit=PERF_APPEND_987654321<Cr>ModBusId=987654321<Cr>\n'.encode())
        before = time.perf_counter()
        result = store.sync(copies)
        report['append_seconds'] = time.perf_counter() - before
        assert not result['failures'] and store.search('PERF_APPEND_987654321'), result
        report['append_modes'] = [item['mode'] for item in result['indexing']]
        report['memory_final_mb'] = memory_mb()
        report['final_cache_bytes'], report['final_archive_bytes'] = storage(database), storage(archive)
    assert manifest == {str(p): (p.stat().st_size, p.stat().st_mtime_ns) for p in originals}
    report['sources_unchanged'] = True
    report['status'] = 'complete'
    save()
    print(json.dumps({key: report[key] for key in ('first_index_seconds', 'reopen_cache_available_seconds',
                                                 'reopen_verification_seconds', 'append_seconds')}, indent=2))


if __name__ == '__main__':
    main()
