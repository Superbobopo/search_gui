"""Prove cached searches and GET/PUTE work while history preparation is blocked."""

from contextlib import ExitStack
import os
from pathlib import Path
import sys
import tempfile
import threading
import time
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import indexed_ui
import search_gui as gui
import trace_sources as sources
from trace_parser import read_trace
from trace_store import TraceStore, Cancelled
from tools.check_omega_history import descendants
from tools.ui_test_support import run_gui_check


def main():
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        root = base / 'traces'
        paths = []
        for robot in ('r1', 'r2'):
            path = root / robot.upper() / 'TracOmega.txt'
            path.parent.mkdir(parents=True)
            text = read_trace(Path(__file__).resolve().parents[1] / f'tests/fixtures/current_{robot}.txt')
            path.write_bytes(text.encode('cp1252'))
            paths.append(str(path))
        database, archive = base / 'index.db', base / 'history.db'
        seeded = TraceStore(database, directories=[root], history_database=archive)
        assert not seeded.sync(paths)['failures']
        stack.enter_context(patch.object(sources, 'application_directory', return_value=str(base)))
        stack.enter_context(patch.object(sources, 'desktop_directory', return_value=str(base / 'absent')))
        stack.enter_context(patch.object(sources, 'FIXED_DIRECTORY', str(base / 'absent2')))
        stack.enter_context(patch.dict(os.environ, LOCALAPPDATA=str(base / 'AppData')))
        opening = []
        def open_cache(database_arg, **kwargs):
            opening.append(time.perf_counter())
            return TraceStore(database, history_database=archive, **kwargs)
        stack.enter_context(patch.object(indexed_ui, 'TraceStore', side_effect=open_cache))
        gate, preparing = threading.Event(), threading.Event()
        original_prepare = TraceStore.prepare_history
        def blocked(store, stop=None):
            if threading.current_thread().name == 'trace-index':
                preparing.set()
                while not gate.wait(.01):
                    if stop and stop.is_set():
                        raise Cancelled()
            return original_prepare(store, stop)
        stack.enter_context(patch.object(TraceStore, 'prepare_history', blocked))
        before = time.perf_counter()
        app = gui.SearchApp()
        app.geometry('1100x900')
        errors = []
        app.report_callback_exception = lambda *error: errors.append(error)
        def pump(condition, seconds=5):
            deadline = time.monotonic() + seconds
            while time.monotonic() < deadline:
                app.update()
                if errors:
                    raise errors[0][1]
                if condition():
                    return
                time.sleep(.005)
            raise AssertionError(app.status.get() + ' / ' + app.index_status.get())
        try:
            pump(lambda: hasattr(app, 'store') and len(app.file_list) == 2 and preparing.is_set())
            startup = time.perf_counter() - before
            available = time.perf_counter() - opening[0]
            assert available < 1, available
            assert 'Cache disponible' in app.index_status.get()
            assert not app._sources_verified
            assert app.btn_erreurs.cget('state') == 'disabled'
            app.keyword.insert(0, '1194')
            app.modbus.set(True)
            app.on_modbus_check()
            app.btn_search.invoke()
            pump(lambda: app.selection is not None)
            selected = app.selection
            app.result_switch._buttons_dict['Classique'].invoke()
            pump(lambda: app.classic_frame.winfo_ismapped())
            expected = app.store.file_lines(selected.path)
            for index, key in enumerate(('get', 'pute')):
                event = selected.details[key]
                old = {w for w in app.winfo_children() if isinstance(w, gui.tk.Toplevel)}
                position = f"{app.txt.tag_ranges('clickable_' + key)[0]} +1lines +5chars"
                app.txt.see(position)
                app.update()
                box = app.txt.bbox(position)
                assert box
                x, y = box[0] + 2, box[1] + 2
                assert 'clickable_' + key in app.txt.tag_names(f'@{x},{y}')
                app.txt.event_generate('<Motion>', x=x, y=y)
                for click in range(2):
                    stamp = 1000 + index * 1000 + click * 100
                    app.txt.event_generate('<ButtonPress-1>', x=x, y=y, time=stamp)
                    app.txt.event_generate('<ButtonRelease-1>', x=x, y=y, time=stamp + 20)
                def opened():
                    return {w for w in app.winfo_children() if isinstance(w, gui.tk.Toplevel)} - old
                pump(lambda: bool(opened()))
                window = opened().pop()
                text = next(w for w in descendants(window) if isinstance(w, gui.tk.Text))
                pump(lambda: bool(text.tag_ranges('highlight')))
                start, end = text.tag_ranges('highlight')
                assert text.get(start, end) == event['text']
                assert int(str(start).split('.')[0]) == event['number'] + 1
                assert text.get('1.0', 'end-1c').splitlines() == expected
                window.destroy()
            assert not gate.is_set() and not app._sources_verified
            gate.set()
            pump(lambda: app._sources_verified)
            assert app.selection.path == selected.path and app.selection.number == selected.number
            assert app.btn_erreurs.cget('state') == 'normal'
            print(f'OK cache disponible en {available:.3f} s apres debut ouverture (construction interface comprise : {startup:.3f} s) ; recherche et vrais doubles-clics GET/PUTE pendant preparation bloquee ; selection stable apres verification')
        finally:
            gate.set()
            app._close_indexed_app()
            app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)


if __name__ == '__main__':
    run_gui_check(main)
