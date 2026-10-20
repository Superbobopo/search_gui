"""Reference click and drag/copy regression on a private Windows desktop."""

from contextlib import ExitStack
import os
from pathlib import Path
import sys
import tempfile
import time
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import customtkinter as ctk
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
        for robot, name, contents in (
                ('R2', 'TracOMEGA.txt', original),
                ('R2', 'TracOmega029.old', original.replace('14/10/2026', '13/10/2026').replace('14|', '13|')),
                ('R1', 'TracOMEGA.txt', original.replace('A0061574AC', 'OTHER_R1')
                    .replace('14/10/2026', '12/10/2026').replace('14|', '12|'))):
            path = base / 'traces' / robot / name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(contents, encoding='utf8')
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

        def ref_widget():
            return next(w.selection_text for w in descendants(app.summary_frame)
                        if isinstance(w, ctk.CTkLabel) and w.cget('text') == 'A0061574AC')

        try:
            pump(lambda: app._sources_verified)
            app.keyword.insert(0, '1194')
            app.modbus.set(True)
            app.on_modbus_check()
            app.btn_search.invoke()
            pump(lambda: app.selection is not None and '_summary_comparison' in app.selection.details)
            selected = app.selection
            assert selected.path.path == str(base / 'traces/R2/TracOMEGA.txt')
            query = stack.enter_context(patch.object(app.store, 'search', wraps=app.store.search))
            widget = ref_widget()
            first, last = widget.bbox('1.1'), widget.bbox('1.5')
            assert first and last
            widget.event_generate('<ButtonPress-1>', x=first[0], y=first[1] + 2, time=1000)
            widget.event_generate('<B1-Motion>', x=last[0], y=last[1] + 2, state=256, time=1300)
            widget.event_generate('<ButtonRelease-1>', x=last[0], y=last[1] + 2, time=1400)
            pump()
            assert widget.get('sel.first', 'sel.last') == '0061'
            with patch.object(widget, 'clipboard_clear'), patch.object(widget, 'clipboard_append') as copy:
                widget.focus_force()
                widget.event_generate('<Control-KeyPress-c>')
                pump()
                copy.assert_called_once_with('0061')
            assert app.selection is selected and app.modbus.get() and app.keyword.get() == '1194'
            query.assert_not_called()
            widget.event_generate('<ButtonPress-1>', x=4, y=4, time=3000)
            widget.event_generate('<ButtonRelease-1>', x=4, y=4, time=3040)
            assert app.keyword.get() == 'A0061574AC'
            assert app.search_mode.get() == 'ref' and app.reference.get() and not app.modbus.get()
            pump(lambda: app.selection is not None and app._active_mode == 'ref')
            assert (app.selection.path, app.selection.number) == (selected.path, selected.number)
            assert app.selection.version_id == selected.version_id
            assert len(app._groups) == 2
            assert all('/R1/' not in os.fspath(path).replace('\\', '/') for path, _ in app._groups)
            count = query.call_count
            widget = ref_widget()
            widget.event_generate('<ButtonPress-1>', x=4, y=4, time=5000)
            widget.event_generate('<ButtonRelease-1>', x=4, y=4, time=5040)
            pump()
            assert app.reference.get() and query.call_count == count
            # Classic keeps its historical reverse link; the preferred occurrence stays exact.
            app.result_switch._buttons_dict['Classique'].invoke()
            pump()
            pos = str(app.txt.tag_ranges('clickable_refprod')[0]) + ' +5chars'
            box = app.txt.bbox(pos)
            assert box
            app.txt.event_generate('<Motion>', x=box[0] + 2, y=box[1] + 2)
            app.txt.event_generate('<ButtonPress-1>', x=box[0] + 2, y=box[1] + 2, time=7000)
            app.txt.event_generate('<ButtonRelease-1>', x=box[0] + 2, y=box[1] + 2, time=7040)
            pump(lambda: app.selection is not None and app._active_mode == 'id')
            assert app.keyword.get() == '1194' and app.search_mode.get() == 'id'
            assert (app.selection.path, app.selection.number) == (selected.path, selected.number)
            print('OK Synthese : vrai clic reference ID -> reference, champ/radio/recherche et occurrence exacte ; '
                  'glisser/copier sans recherche, clic deja en reference neutre ; lien Classique preserve, '
                  'fichiers homonymes et ID partage R1/R2')
        finally:
            app._close_indexed_app()
            app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)


if __name__ == '__main__':
    run_gui_check(main)
