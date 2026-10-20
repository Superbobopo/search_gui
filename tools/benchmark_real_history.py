"""Read-only comparisons on an existing history cache; never syncs or migrates."""
from contextlib import contextmanager
from pathlib import Path
from unittest.mock import patch
import argparse
import json
import os
import sqlite3
import sys
import threading
import time

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
parser = argparse.ArgumentParser()
parser.add_argument('--output', required=True)
parser.add_argument('--database')
parser.add_argument('--module-dir')
parser.add_argument('--compare')
args = parser.parse_args() if __name__ == '__main__' else None
if args is not None and args.module_dir:
    sys.path.insert(0, str(Path(args.module_dir).resolve()))
from trace_store import TraceStore, cache_path
import get_diagnostics
import trace_store


class ReadOnlyStore(TraceStore):
    def _initialize(self):
        pass

    @contextmanager
    def connection(self):
        if getattr(self._local, 'snapshot', None) is not None:
            yield self._local.snapshot
            return
        db = sqlite3.connect(self.database.as_uri()+'?mode=ro', uri=True, timeout=2)
        db.row_factory = sqlite3.Row
        db.execute('PRAGMA query_only=ON')
        db.set_progress_handler(lambda: int(stop.is_set()), 1000)
        try:
            yield db
        finally:
            db.close()

    def sync(self, *args, **kwargs):
        raise AssertionError('No synchronization on real data')

    def prepare_history(self, *args, **kwargs):
        raise AssertionError('No history preparation on real data')

    def prepare_analysis(self, *args, **kwargs):
        raise AssertionError('No schema changes on real data')


def normalize(value):
    if hasattr(value, 'version_id'):
        return dict(path=os.fspath(value), generation=value.version_id, historical=value.historical)
    if isinstance(value, dict):
        return {str(k): normalize(v) for k,v in value.items()}
    if isinstance(value, (tuple, list)):
        return [normalize(v) for v in value]
    return value


stop = threading.Event()


def main():
    output = Path(args.output)
    if output.exists():
        raise FileExistsError(output)
    database = Path(args.database) if args.database else cache_path([str(PROJECT/'traces')])
    if not database.exists():
        raise FileNotFoundError(database)
    before = database.stat()
    store = ReadOnlyStore(database)
    revision = store.analysis_revision()
    groups = store.search('ModBusId=1054', True)
    path, number = next((p,n) for p,hits in groups for n,text in hits
                        if Path(p).name.lower()=='tracomega019.old'
                        and 'trace r1' in str(p).lower() and 'A0118427AA' in text)
    measurements, results = {}, {}
    original = store._rows
    sql = []
    def rows(*a, **kw):
        start = time.perf_counter()
        value = original(*a, **kw)
        sql.append((time.perf_counter()-start, len(value)))
        return value
    def measure(name, work):
        sql.clear(); spy.reset_mock(); stop.clear()
        timer = threading.Timer(30, stop.set); timer.start()
        start = time.perf_counter()
        try:
            value = work()
        finally:
            timer.cancel()
        measurements[name] = dict(ms=(time.perf_counter()-start)*1000, queries=len(sql),
                                  rows=sum(n for _,n in sql), get_diagnostic_calls=spy.call_count,
                                  diagnostic_files=len({call.args[1] for call in spy.call_args_list}))
        results[name] = normalize(value)
        print(name, json.dumps(measurements[name]), flush=True)
        return value
    # _file is the shared entry point, including the selected-neighbor diagnostics.
    with patch.object(store, '_rows', side_effect=rows), \
         patch.object(get_diagnostics, '_file', wraps=get_diagnostics._file) as spy:
        measure('details', lambda: store.details(path, number))
        measure('current', lambda: store.current_measurement_for_occurrence(path, number, stop))
        measure('comparison', lambda: store.comparison_for_occurrence(path, number, stop))
        adjacent = measure('adjacent', lambda: store.paired_adjacent(path, '1054', number, cancel=stop))
        own = next(c for c in adjacent['columns'] if c.get('anchor') and c['anchor']['path']==path)
        destinations = own['following'][:2]+[own['anchor']]
        for i, row in enumerate(destinations):
            def click(row=row):
                p, n = row['path'], row['number']
                return dict(details=store.details(p,n),
                            current=store.current_measurement_for_occurrence(p,n,stop),
                            comparison=store.comparison_for_occurrence(p,n,stop),
                            adjacent=store.paired_adjacent(p,row['modbus'],n,cancel=stop))
            measure(f'click_{i+1}_{row["modbus"]}', click)
    if store.analysis_revision()!=revision:
        raise AssertionError('Cache revision changed during measurement')
    after = database.stat()
    assert (before.st_size,before.st_mtime_ns)==(after.st_size,after.st_mtime_ns)
    report = dict(database=str(database), bytes=before.st_size, revision=revision,
                  metadata_files=len(store.metadata()), measurements=measurements, results=results)
    if args.compare:
        baseline = json.loads(Path(args.compare).read_text(encoding='utf8'))
        assert baseline['revision']==revision, 'Baseline revision differs'
        assert baseline['results']==results, 'Business results changed'
        report['identical_to_baseline'] = True
        print('All business results exactly identical; cache size and mtime unchanged', flush=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf8')


if __name__ == '__main__':
    main()
