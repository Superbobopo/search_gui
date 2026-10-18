"""Check ID 683 and display timing using the existing cache in read-only mode.

Never initializes or synchronizes the application cache. Requires local traces.
"""
from contextlib import contextmanager
from pathlib import Path
import sqlite3
import sys
import time
from unittest.mock import patch

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
import indexed_ui
import search_gui as gui
from trace_store import TraceStore, cache_path
from tools.check_interactions import descendants


class ReadOnlyStore(TraceStore):
    def _initialize(self):
        pass

    @contextmanager
    def connection(self):
        existing = getattr(self._local, 'snapshot', None)
        if existing is not None:
            yield existing
            return
        db = sqlite3.connect(self.database.as_uri() + '?mode=ro', uri=True)
        db.row_factory = sqlite3.Row
        try:
            yield db
        finally:
            db.close()


def main():
    path = cache_path()
    if not path.is_file():
        raise SystemExit('Cache existant requis ; aucune indexation ne sera lancée.')
    before_cache = path.stat()
    store = ReadOnlyStore(path)
    started = time.perf_counter()
    metadata = store.metadata()
    before_files = {row['path']: Path(row['path']).stat() for row in metadata}
    print(f'Reouverture et stat de {len(metadata)} fichiers: {time.perf_counter() - started:.4f} s', flush=True)
    hits = store.search('ModBusId=683', True, ref='A0061256AB')
    source = next(path for path, _ in hits if Path(path).parent.name.lower() == 'trace r2')
    number = next(hits_ for path, hits_ in hits if path == source)[0][0]
    def ready(app):
        app._index_messages.put(('ready', store, metadata, True))
        app._index_messages.put(('complete', {'changed': 0, 'removed': 0, 'failures': []}, metadata))
    with patch.object(indexed_ui.IndexedAppMixin, '_index_loop', ready):
        started = time.perf_counter()
        app = gui.SearchApp(); app.geometry('1100x900+18000+18000')
        errors = []
        app.report_callback_exception = lambda *args: errors.append(args)
        def pump(condition, seconds=15):
            deadline = time.monotonic() + seconds
            while time.monotonic() < deadline:
                app.update()
                for window in app.winfo_children():
                    if isinstance(window, gui.tk.Toplevel) and window.winfo_x() < 10000:
                        window.geometry('+18000+18000')
                if errors:
                    raise errors[0][1]
                if condition(): return
                time.sleep(.002)
            raise AssertionError('Timeout: ' + app.status.get())
        try:
            pump(lambda: hasattr(app, 'store'))
            print(f'Interface disponible: {time.perf_counter() - started:.4f} s', flush=True)
            for value, mode in [('683', 'id'), ('A0061256AB', 'ref'), ('REF_ABSENTE_999', 'ref')]:
                started = time.perf_counter()
                app._launch_search(value, mode, None, None)
                pump(lambda: 'Recherche…' not in app.status.get())
                print(f'Recherche {value}, affichage compris: {(time.perf_counter() - started) * 1000:.1f} ms', flush=True)
            app._focus_indexed_ticket(source, '683', 'A0061256AB', number=number)
            pump(lambda: app.selection is not None)
            assert app.selection.number == 65168
            assert '11:00:32' in app.txt.get('1.0', 'end')
            assert '11:00:36.9' in app.selection.details['get']['text']
            assert '11:00:52.3' in app.selection.details['pute']['text']
            app.btn_adjacent.invoke(); pump(lambda: app.adj_popup is not None)
            assert '683' in ' '.join(w.cget('text') for w in app.current_frame.winfo_children())
            print('Cas capture: Local R2, ticket ligne 65169, GET 11:00:36.9, PUTE 11:00:52.3', flush=True)
            # Real doubles, complete contents and exact source line on the actual large trace.
            stamp = 2000
            for key in ('get', 'pute'):
                event = app.selection.details[key]
                pos = f"{app.txt.tag_ranges('clickable_' + key)[0]} +1lines +5chars"
                app.txt.see(pos); app.update()
                box = app.txt.bbox(pos); assert box
                x, y = box[0] + 2, box[1] + 2
                app.txt.event_generate('<Motion>', x=x, y=y)
                previous = set(app.winfo_children())
                for i in range(2):
                    app.txt.event_generate('<ButtonPress-1>', x=x, y=y, time=stamp + i * 100)
                    app.txt.event_generate('<ButtonRelease-1>', x=x, y=y, time=stamp + i * 100 + 20)
                stamp += 1000
                def opened():
                    return [w for w in app.winfo_children() if w not in previous and isinstance(w, gui.tk.Toplevel)]
                pump(lambda: bool(opened()))
                view = opened()[0]
                widget = next(w for w in descendants(view) if isinstance(w, gui.tk.Text))
                pump(lambda: bool(widget.tag_ranges('highlight')))
                begin, end = widget.tag_ranges('highlight')
                assert int(str(begin).split('.')[0]) == event['number'] + 1
                assert widget.get(begin, end) == event['text']
                assert len(widget.get('1.0', 'end-1c').splitlines()) == next(row['line_count'] for row in metadata if row['path'] == source)
                view.destroy()
            print('Doubles-clics reels GET/PUTE sur les 66 traces: fichier et surlignage exacts', flush=True)
            assert not errors
        finally:
            app._close_indexed_app()
            app._index_thread.join(timeout=5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)
    assert (path.stat().st_size, path.stat().st_mtime_ns) == (before_cache.st_size, before_cache.st_mtime_ns)
    for name, before in before_files.items():
        after = Path(name).stat()
        assert (before.st_size, before.st_mtime_ns) == (after.st_size, after.st_mtime_ns)
    print('Cache version 1 et traces inchanges, aucune synchronisation/reindexation', flush=True)


if __name__ == '__main__':
    from tools.ui_test_support import run_gui_check
    run_gui_check(main)
