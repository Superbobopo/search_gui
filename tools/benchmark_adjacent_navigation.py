"""Compare real adjacent clicks on temporary traces, on a private desktop."""
from contextlib import ExitStack, contextmanager
import argparse
import ctypes
from ctypes import wintypes
from datetime import datetime, timedelta
import json
import os
from pathlib import Path
import statistics
import sys
import tempfile
import threading
import time
from unittest.mock import patch

PROJECT = Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser()
parser.add_argument('--output', required=True)
parser.add_argument('--stress', action='store_true')
args = parser.parse_args()
sys.path.insert(0, str(PROJECT))
import search_gui as gui
import trace_sources as sources
from tools.ui_test_support import run_gui_check


def rss():
    class Counters(ctypes.Structure):
        _fields_ = [('cb', wintypes.DWORD), ('faults', wintypes.DWORD)] + [
            (name, ctypes.c_size_t) for name in ('peak', 'working', 'poolpeak', 'pool',
                                               'nonpoolpeak', 'nonpool', 'page', 'pagepeak')]
    kernel = ctypes.WinDLL('kernel32')
    kernel.GetCurrentProcess.restype = wintypes.HANDLE
    process = ctypes.WinDLL('psapi')
    process.GetProcessMemoryInfo.argtypes = [wintypes.HANDLE, ctypes.c_void_p, wintypes.DWORD]
    counter = Counters()
    counter.cb = ctypes.sizeof(counter)
    if not process.GetProcessMemoryInfo(kernel.GetCurrentProcess(), ctypes.byref(counter), counter.cb):
        raise ctypes.WinError()
    return counter.working


def main():
    output = Path(args.output)
    assert not output.exists(), output
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        for robot, offset in (('R1', 0), ('R2', 2)):
            folder = base / 'traces' / robot
            folder.mkdir(parents=True)
            for name, hour in ((('TracOmega019.old', 8),) if args.stress else (('TracOmega018.old', 8), ('TracOmega019.old', 12))):
                lines = ['creation de la trace de ... Le 14/10/2026\n']
                for i in range(5000 if args.stress else 300):
                    stamp = datetime(2026, 10, 14, hour) + timedelta(seconds=i * 10 + offset)
                    def at(extra):
                        return (stamp + timedelta(seconds=extra)).strftime('%d|%H:%M:%S.0')
                    ref = f'REF{i:04}'
                    lines.extend((f'{at(0)} RefProduit={ref}<Cr>SN={i}<Cr>CodeLot=LOT{i}<Cr>ModBusId={1000+i}<Cr>\n',
                                  f'G {at(1)} GRIPPER GET 1x boîte(s): Code:{ref} $0101\n',
                                  f'{at(2)} IDPK pour {ref}: l={100 if hour==12 else 90} H=20 L=50\n',
                                  f'G {at(3)} GRIPPER PUTE 1x boîte(s): Code:{ref} face:1 601\n'))
                (folder / name).write_text(''.join(lines), encoding='utf8')
        stack.enter_context(patch.object(sources, 'application_directory', return_value=str(base)))
        stack.enter_context(patch.object(sources, 'desktop_directory', return_value=str(base / 'missing')))
        stack.enter_context(patch.object(sources, 'FIXED_DIRECTORY', str(base / 'absent')))
        stack.enter_context(patch.dict(os.environ, LOCALAPPDATA=str(base / 'AppData')))
        app = gui.SearchApp()
        errors = []
        app.report_callback_exception = lambda *error: errors.append(error)
        def pump(condition=lambda: True):
            deadline = time.perf_counter() + 20
            while time.perf_counter() < deadline:
                app.update()
                if errors:
                    raise errors[0][1]
                if condition():
                    return
                time.sleep(.002)
            raise AssertionError('Benchmark timeout: ' + app.status.get())
        queries = [0]
        transferred = [0]
        lock = threading.Lock()
        def trace_sql(sql):
            with lock:
                queries[0] += 1
        try:
            pump(lambda: app._sources_verified)
            stack.enter_context(patch.object(app.store, 'sync', return_value=dict(changed=0, removed=0, failures=[])))
            original = app.store.connection
            @contextmanager
            def counted():
                with original() as db:
                    db.set_trace_callback(trace_sql)
                    yield db
            stack.enter_context(patch.object(app.store, 'connection', counted))
            original_rows = app.store._rows
            def counted_rows(*a, **kw):
                result = original_rows(*a, **kw)
                with lock:
                    transferred[0] += len(result)
                return result
            stack.enter_context(patch.object(app.store, '_rows', counted_rows))
            app.mode_radios['ref'].invoke()
            app.keyword.insert(0, 'REF0250')
            app.btn_search.invoke()
            pump(lambda: app.selection is not None and '_summary_comparison' in app.selection.details)
            # Choose the current R1 file through the actual file navigation.
            while not (sources.robot_scope(app.selection.path)[1] == 'R1'
                       and Path(app.selection.path).name == 'TracOmega019.old'):
                prior = app.selection
                app.btn_next_file.invoke()
                pump(lambda: app.selection is not None and app.selection is not prior)
            start, q, r = time.perf_counter(), queries[0], transferred[0]
            app.btn_adjacent.invoke()
            pump(lambda: app.adjacent_view.data is not None)
            cold = dict(ms=(time.perf_counter()-start)*1000, queries=queries[0]-q, transferred_rows=transferred[0]-r)
            samples = []
            for target in (['1251', '1252', '1253', '1252', '1251', '1250'] if args.stress else ['1251', '1252', '1251', '1250'] * 3):
                view = app.adjacent_view
                before_widgets = tuple(str(tree) for tree in view.trees.values())
                key = next(key for key, row in view.rows['R1'].items() if row['modbus']==target)
                tree = view.trees['R1']
                tree.see(key)
                pump()
                x, y, width, height = tree.bbox(key)
                start, q, r = time.perf_counter(), queries[0], transferred[0]
                tree.event_generate('<ButtonPress-1>', x=x+10, y=y+height//2)
                tree.event_generate('<ButtonRelease-1>', x=x+10, y=y+height//2)
                pump(lambda: app.selection is not None and app.selection.modbus==target
                     and view.data is not None and view.data['selected']['modbus']==target
                     and '_summary_comparison' in app.selection.details)
                samples.append(dict(ms=(time.perf_counter()-start)*1000, queries=queries[0]-q, transferred_rows=transferred[0]-r,
                                    same_widgets=before_widgets==tuple(str(t) for t in view.trees.values())))
            report = dict(corpus=('2 temporary OMEGA files, 10000 tickets, 5000 identities, 2 robots' if args.stress else '4 temporary OMEGA files, 1200 tickets, 300 identities, 2 robots'),
                          cold=cold, navigation=samples,
                          median_ms=statistics.median(s['ms'] for s in samples),
                          median_queries=statistics.median(s['queries'] for s in samples),
                          cache_weight=app.store._analysis_cache.weight, rss_bytes=rss(), database_bytes=app.store.database.stat().st_size)
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_text(json.dumps(report, indent=2), encoding='utf8')
            print(json.dumps({k: v for k, v in report.items() if k!='navigation'}))
        finally:
            app._close_indexed_app()
            app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)


if __name__ == '__main__':
    run_gui_check(main)
