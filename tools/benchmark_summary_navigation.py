"""Measure actual navigation clicks and UI stalls; existing cache is read-only."""
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import patch
import argparse
import json
import os
import statistics
import sys
import tempfile
import time
import tkinter as tk

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
from tools.benchmark_real_history import ReadOnlyStore, normalize
from trace_store import cache_path
from tools.ui_test_support import run_gui_check
from tools.summary_test_support import click_summary
from tools.native_test_capture import widget_image


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--output', required=True)
    parser.add_argument('--module-dir')
    parser.add_argument('--compare')
    parser.add_argument('--database')
    args = parser.parse_args()
    output = Path(args.output).resolve()
    if output.exists():
        raise FileExistsError(output)
    if args.module_dir:
        sys.path.insert(0, str(Path(args.module_dir).resolve()))
    import indexed_ui
    import search_gui as gui
    database = Path(args.database) if args.database else cache_path([str(PROJECT/'traces')])
    if not database.exists():
        raise FileNotFoundError(database)
    before = database.stat()
    store = ReadOnlyStore(database)
    revision = store.analysis_revision()
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
        errors, beats, renders, counts = [], [], [], dict(created=0, destroyed=0)
        app.report_callback_exception = lambda *error: errors.append(error)
        original_render = app._render_summary

        def render(data):
            start = time.perf_counter()
            try:
                return original_render(data)
            finally:
                renders.append(dict(stage='empty' if not data else 'final' if '_summary_comparison' in data
                                    else 'current' if '_summary_current' in data else 'initial',
                                    ms=(time.perf_counter()-start)*1000))

        app._render_summary = render

        def heartbeat():
            beats.append(time.perf_counter())
            app.after(8, heartbeat)

        app.after(8, heartbeat)

        def pump(condition=lambda: True, settle=0):
            until, stable = time.perf_counter()+20, None
            while time.perf_counter()<until:
                app.update()
                if errors:
                    raise errors[0][1]
                if condition():
                    stable = stable or time.perf_counter()
                    if time.perf_counter()-stable >= settle:
                        return
                else:
                    stable = None
                time.sleep(.001)
            raise AssertionError(app.status.get())

        def final():
            return app.selection is not None and '_summary_comparison' in app.selection.details

        def click(button):
            button._canvas.event_generate('<ButtonPress-1>', x=5, y=5)
            button._canvas.event_generate('<ButtonRelease-1>', x=5, y=5)

        samples, results = [], []

        def sample(name, action, done=final):
            beats.clear(); renders.clear()
            counts.update(created=0, destroyed=0)
            beats.append(time.perf_counter())
            start = time.perf_counter()
            action()
            callback = time.perf_counter()
            pump(done)
            completed = time.perf_counter()
            pump(settle=.12)
            samples.append(dict(name=name, callback_ms=(callback-start)*1000,
                                complete_ms=(completed-start)*1000,
                                heartbeat_max_ms=max((b-a for a,b in zip(beats, beats[1:])), default=0)*1000,
                                widgets=dict(counts), renders=list(renders)))
            data = app.selection.details
            results.append(normalize(dict(details={k:v for k,v in data.items() if not k.startswith('_')},
                                          comparison=data['_summary_comparison'], groups=app._groups)))

        setup, destroy = tk.BaseWidget._setup, tk.BaseWidget.destroy

        def counted_setup(widget, *args, **kw):
            counts['created'] += 1
            return setup(widget, *args, **kw)

        def counted_destroy(widget):
            counts['destroyed'] += 1
            return destroy(widget)

        try:
            pump(lambda: app._sources_verified)
            app.modbus.set(True); app.reference.set(False)
            app.keyword.delete(0, 'end'); app.keyword.insert(0, '1054')
            click(app.btn_search); pump(final, settle=.15)
            with patch.object(tk.BaseWidget, '_setup', counted_setup), patch.object(tk.BaseWidget, 'destroy', counted_destroy):
                for cycle in range(5):
                    for number, button in ((1053, app.btn_up), (1052, app.btn_up),
                                           (1053, app.btn_down), (1054, app.btn_down)):
                        sample(f'id_{number}_cycle_{cycle}', lambda button=button: click(button))
                        assert app.keyword.get()==str(number)
                sample('reference_link', lambda: click_summary(app, 'refprod', stamp=900000, double=False),
                       lambda: final() and app._active_mode=='ref')
            widget_image(app).save(output.with_suffix('.png'))
            app._change_theme('Sombre'); pump(settle=.1)
            widget_image(app).save(output.with_name(output.stem+'_dark.png'))
            if args.compare:
                previous = json.loads(Path(args.compare).read_text(encoding='utf8'))
                assert previous['revision']==revision and previous['results']==results, 'Business results changed'
            report = dict(revision=revision, database=str(database), samples=samples, results=results)
            output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf8')
            for kind, selected in (('first cycle', samples[:4]), ('warm cycles', samples[4:20]), ('reference', samples[20:])):
                print(kind, {key:round(statistics.median(s[key] for s in selected),2)
                             for key in ('callback_ms','complete_ms','heartbeat_max_ms')},
                      'widgets', [{k:s['widgets'][k] for k in ('created','destroyed')} for s in selected[:4]], flush=True)
        finally:
            app._close_indexed_app(); app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)
    after = database.stat()
    assert (before.st_size,before.st_mtime_ns)==(after.st_size,after.st_mtime_ns)
    assert store.analysis_revision()==revision


if __name__=='__main__':
    run_gui_check(main)
