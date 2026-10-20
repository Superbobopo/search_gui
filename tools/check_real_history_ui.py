"""Actual clicks against the real history cache, read-only, private desktop only."""
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import patch
import argparse
import json
import os
import sys
import tempfile
import time

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
from tools.benchmark_real_history import ReadOnlyStore, normalize
from tools.ui_test_support import run_gui_check
from tools.native_test_capture import widget_image
from tools.summary_test_support import click_summary
import indexed_ui
import search_gui as gui


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--baseline', required=True)
    parser.add_argument('--output', required=True)
    args = parser.parse_args()
    baseline = json.loads(Path(args.baseline).read_text(encoding='utf8'))
    output = Path(args.output)
    if output.exists():
        raise FileExistsError(output)
    database = Path(baseline['database'])
    before = database.stat()
    store = ReadOnlyStore(database)
    assert store.analysis_revision()==baseline['revision']
    metadata = store.metadata()
    def ready(app):
        app._index_messages.put(('ready', store, metadata, True))
        app._index_messages.put(('complete', dict(changed=0, removed=0, failures=[]), metadata))
    with ExitStack() as stack:
        folder = stack.enter_context(tempfile.TemporaryDirectory())
        stack.enter_context(patch.dict(os.environ, LOCALAPPDATA=folder))
        stack.enter_context(patch.object(indexed_ui.IndexedAppMixin, '_index_loop', ready))
        app = gui.SearchApp()
        app.geometry('1090x820+0+0')
        errors = []
        app.report_callback_exception = lambda *error: errors.append(error)
        def pump(condition=lambda: True):
            until = time.perf_counter()+15
            while time.perf_counter()<until:
                app.update()
                if errors:
                    raise errors[0][1]
                if condition():
                    return
                time.sleep(.003)
            raise AssertionError(app.status.get())
        timings = {}
        try:
            pump(lambda: app._sources_verified)
            app.modbus.set(True); app.reference.set(False)
            app.keyword.insert(0, '1054')
            start = time.perf_counter(); app.btn_search.invoke()
            pump(lambda: app.selection is not None and '_summary_comparison' in app.selection.details)
            timings['search_ms'] = (time.perf_counter()-start)*1000
            while not (Path(app.selection.path).name.lower()=='tracomega019.old'
                       and 'trace r1' in str(app.selection.path).lower()):
                previous = app.selection
                app.btn_next_file.invoke()
                pump(lambda: app.selection is not None and app.selection is not previous
                     and '_summary_comparison' in app.selection.details)
            assert normalize(app.selection.details['_summary_comparison'])==baseline['results']['comparison']
            start = time.perf_counter(); app.btn_adjacent.invoke()
            pump(lambda: getattr(app, 'adjacent_view', None) is not None and app.adjacent_view.data is not None)
            timings['adjacent_ms'] = (time.perf_counter()-start)*1000
            view = app.adjacent_view
            assert normalize(view.data)==baseline['results']['adjacent']
            expected = baseline['results']['click_1_1062']
            key, row = next((key,row) for key,row in view.rows['R1'].items() if row['modbus']=='1062')
            tree = view.trees['R1']; tree.see(key); pump()
            x,y,w,h = tree.bbox(key)
            start = time.perf_counter()
            tree.event_generate('<ButtonPress-1>', x=x+10, y=y+h//2)
            tree.event_generate('<ButtonRelease-1>', x=x+10, y=y+h//2)
            pump(lambda: app.selection is not None and app.selection.number==row['number']
                 and app.selection.path==row['path'] and '_summary_comparison' in app.selection.details
                 and view.data['selected']['number']==row['number'])
            timings['neighbor_click_ms'] = (time.perf_counter()-start)*1000
            assert normalize(view.data)==expected['adjacent']
            assert normalize(app.selection.details['_summary_comparison'])==expected['comparison']
            event = app.selection.details['get']
            click_summary(app, 'get', stamp=900000)
            def traces():
                return [w for w in app.winfo_children() if getattr(w, 'trace_controller', None)]
            pump(lambda: traces() and traces()[0].trace_controller.loaded)
            controller = traces()[0].trace_controller
            assert controller.text.get('highlight.first','highlight.last')==event['text']
            assert controller.text.index('highlight.first')==str(event['number']+1)+'.0'
            traces()[0].destroy()
            app._change_theme('Sombre'); pump()
            widget_image(app).save(output.with_suffix('.png'))
            widget_image(app.adj_popup).save(output.with_name(output.stem+'_neighbors.png'))
            output.write_text(json.dumps(timings, indent=2), encoding='utf8')
            print('OK real read-only cache, actual search/adjacent/click, exact baseline results, GET highlight:', timings)
        finally:
            app._close_indexed_app()
            app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)
    after = database.stat()
    assert (before.st_size,before.st_mtime_ns)==(after.st_size,after.st_mtime_ns)


if __name__ == '__main__':
    run_gui_check(main)
