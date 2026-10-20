"""Tracking table interactions on temporary traces and a private Windows desktop."""

from contextlib import ExitStack
import os
from pathlib import Path
import sys
import tempfile
import time
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import search_gui as gui
import trace_sources as sources
from trace_parser import read_trace
from tools.ui_test_support import run_gui_check


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def main():
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        original = read_trace(Path(__file__).resolve().parents[1] / 'tests/fixtures/current_r2.txt')
        content = {}
        for robot, name, old in (('trace r1', 'TracOMEGA.txt', False),
                                  ('trace r2', 'TracOMEGA.txt', False),
                                  ('trace r2', 'TracOmega029.old', True)):
            path = base / 'traces' / robot / name
            path.parent.mkdir(parents=True, exist_ok=True)
            value = original.replace('14/10/2026', '13/10/2026').replace('14|', '13|') if old else original
            path.write_text(value, encoding='utf8')
            content[str(path)] = value
        stack.enter_context(patch.object(sources, 'application_directory', return_value=str(base)))
        stack.enter_context(patch.object(sources, 'desktop_directory', return_value=str(base / 'absent')))
        stack.enter_context(patch.object(sources, 'FIXED_DIRECTORY', str(base / 'missing')))
        stack.enter_context(patch.dict(os.environ, LOCALAPPDATA=str(base / 'AppData')))
        app = gui.SearchApp()
        errors = []
        app.report_callback_exception = lambda *error: errors.append(error)

        def pump(condition=lambda: True):
            deadline = time.monotonic() + 8
            while time.monotonic() < deadline:
                app.update()
                if errors:
                    raise errors[0][1]
                if condition():
                    return
                time.sleep(.005)
            raise AssertionError(app.status.get())

        try:
            pump(lambda: app._sources_verified)
            app.btn_tracking_robot.invoke()
            view = app.tracking_view
            app.keyword.delete(0, 'end')
            app.keyword.insert(0, 'A0061574AC')
            app.btn_search.invoke()
            pump(lambda: len(view.rows) == 3 and view._pending is None)
            assert [view.rows[k]['time'] for k in view.rows] == sorted(
                [r['time'] for r in view.rows.values()], reverse=True)
            assert {view.values[k][2] for k in view.rows} == {'R1', 'R2'}
            assert view.tree['columns'] == ('date', 'hour', 'robot', 'location', 'trace')
            assert view.tree.item('0', 'tags') != view.tree.item('1', 'tags')
            assert '3 dépôts' in view.count.cget('text') and '3 traces' in app.status.get()
            old_key = next(k for k, row in view.rows.items() if str(row['path']).endswith('.old'))
            event = view.rows[old_key]
            box = view.tree.bbox(old_key, 'location')
            assert box
            x, y = box[0] + 15, box[1] + 12
            view.tree.event_generate('<ButtonPress-1>', x=x, y=y, time=1000)
            view.tree.event_generate('<ButtonRelease-1>', x=x, y=y, time=1020)
            pump(lambda: '13/10/2026' in view.detail_title.cget('text'))
            assert event['text'] == view.detail_event.selection_text.get('1.0', 'end-1c')
            assert '.old' in view.detail_source.cget('text')
            with patch.object(view, 'clipboard_clear') as clear, patch.object(view, 'clipboard_append') as copy:
                view.tree.focus_force()
                view.tree.event_generate('<Control-KeyPress-c>')
                pump()
                clear.assert_called_once()
                copy.assert_called_once_with('\t'.join(view.values[old_key]))
                with patch.object(view.menu, 'tk_popup') as popup:
                    view.tree.event_generate('<Button-3>', x=x, y=y)
                    popup.assert_called_once()
                view.menu.invoke(0)
                assert copy.call_args.args == ('Face 2 · F32',)
            field = view.detail_title.selection_text
            field.tag_add('sel', '1.0', '1.6')
            with patch.object(field, 'clipboard_clear'), patch.object(field, 'clipboard_append') as copy:
                field.event_generate('<<Copy>>')
                copy.assert_called_once_with('Face 2')
            # Same names in R1/R2: opening uses the saved row reference, never its label.
            stack.enter_context(patch.object(app.store, 'sync', return_value=dict(changed=0, removed=0, failures=[])))
            Path(os.fspath(event['path'])).unlink()
            for click in range(2):
                when = 3000 + click * 100
                view.tree.event_generate('<ButtonPress-1>', x=x, y=y, time=when)
                view.tree.event_generate('<ButtonRelease-1>', x=x, y=y, time=when + 20)
            def windows():
                return [w for w in app.winfo_children() if isinstance(w, gui.tk.Toplevel)
                        and w.title().startswith('Vue de')]
            pump(lambda: bool(windows()))
            window = windows()[0]
            text = next(w for w in descendants(window) if isinstance(w, gui.tk.Text))
            pump(lambda: bool(text.tag_ranges('highlight')))
            begin, end = text.tag_ranges('highlight')
            assert text.get(begin, end) == event['text']
            assert int(str(begin).split('.')[0]) == event['number'] + 1
            assert text.get('1.0', 'end-1c').splitlines() == content[os.fspath(event['path'])].splitlines()
            window.destroy()
            app._change_theme('Sombre')
            pump()
            assert view.tree.tag_configure('even')['background'] == app._palette['surface']
            assert view.tree.selection() == (old_key,)
            # Batch population is cancellable and does not append rows from the previous search.
            rows = list(view.rows.values())
            view.begin('large')
            view.show_results(rows * 200)
            assert len(view.rows) == 200 and view._pending is not None
            view.begin('empty')
            pump()
            assert not view.rows and not view.tree.get_children()
            app.keyword.delete(0, 'end')
            app.keyword.insert(0, 'ABSENT')
            app.btn_search.invoke()
            pump(lambda: 'Aucun dépôt' in view.count.cget('text'))
            assert not view.detail_event.selection_text.get('1.0', 'end-1c')
            # Dispose during a scheduled append: no Tk callback may reach destroyed widgets.
            view.begin('large')
            view.show_results(rows * 200)
            app.btn_search_menu.invoke()
            pump()
            print('OK suivi : colonnes, 3 depots R1/R2 chronologiques, alternance, detail copiable, '
                  'Ctrl+C/cellule, vrai double-clic archive supprimee et ligne exacte ; theme, vide, '
                  'annulation des lots et changement de vue sans callback obsolete')
        finally:
            app._close_indexed_app()
            app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)


if __name__ == '__main__':
    run_gui_check(main)
