"""Real archived OMEGA selection/double-click checks on a private desktop."""

from contextlib import ExitStack
import os
from pathlib import Path
import sys
import tempfile
import time
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import indexed_ui
import search_gui as gui
import trace_sources as sources
from trace_parser import read_trace
from trace_store import TraceStore
from tools.ui_test_support import run_gui_check


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def main():
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        directory = base / 'traces' / 'R2'
        directory.mkdir(parents=True)
        path = directory / 'TracOMEGA.txt'
        original = read_trace(Path(__file__).resolve().parents[1] / 'tests/fixtures/current_r2.txt')
        path.write_bytes(original.encode('cp1252'))
        stack.enter_context(patch.object(sources, 'application_directory', return_value=str(base)))
        stack.enter_context(patch.object(sources, 'desktop_directory', return_value=str(base / 'absent')))
        stack.enter_context(patch.object(sources, 'FIXED_DIRECTORY', str(base / 'absent2')))
        stack.enter_context(patch.dict(os.environ, LOCALAPPDATA=str(base / 'AppData')))
        stack.enter_context(patch.object(indexed_ui, 'TraceStore', side_effect=lambda database, **kwargs:
                                        TraceStore(base / 'index.sqlite3', **kwargs)))
        app = gui.SearchApp()
        app.geometry('1100x900')
        errors = []
        app.report_callback_exception = lambda *error: errors.append(error)

        def pump(condition=lambda: True, seconds=8):
            deadline = time.monotonic() + seconds
            while time.monotonic() < deadline:
                app.update()
                if errors:
                    raise errors[0][1]
                if condition():
                    return
                time.sleep(.005)
            raise AssertionError(app.status.get() + ' / ' + app.index_status.get())

        def windows():
            return {w for w in app.winfo_children() if isinstance(w, gui.tk.Toplevel) and w.title().startswith('Vue de')}

        stamp = 1000
        def double_click(key):
            nonlocal stamp
            event = app.selection.details[key]
            before = windows()
            start = app.txt.tag_ranges('clickable_' + key)[0]
            position = f'{start} +1lines +5chars'
            app.txt.see(position)
            pump()
            box = app.txt.bbox(position)
            assert box, (key, position)
            x, y = box[0] + 2, box[1] + 2
            assert 'clickable_' + key in app.txt.tag_names(f'@{x},{y}')
            app.txt.event_generate('<Motion>', x=x, y=y)
            for count in range(2):
                app.txt.event_generate('<ButtonPress-1>', x=x, y=y, time=stamp + count * 100)
                app.txt.event_generate('<ButtonRelease-1>', x=x, y=y, time=stamp + count * 100 + 20)
            stamp += 1000
            pump(lambda: bool(windows() - before))
            window = (windows() - before).pop()
            text = next(w for w in descendants(window) if isinstance(w, gui.tk.Text))
            pump(lambda: bool(text.tag_ranges('highlight')))
            begin, end = text.tag_ranges('highlight')
            assert text.get(begin, end) == event['text']
            assert int(str(begin).split('.')[0]) == event['number'] + 1
            assert text.get('1.0', 'end-1c').splitlines() == original.splitlines()
            # Presence may change during refresh; the generation must stay exact.
            assert app.selection.version_id[:8] in window.title()
            assert path.name in window.title()
            assert app.selection.path.path == str(path)
            window.destroy()

        try:
            pump(lambda: app._indexed_complete)
            app.keyword.insert(0, '1194')
            app.modbus.set(True)
            app.on_modbus_check()
            app.btn_search.invoke()
            pump(lambda: app.selection is not None)
            version = app.selection.version_id
            compact = sources.compact_trace_label(app.selection.path, app._selected_directories)
            assert app.file_label.cget('text') == 'Fichier : ' + compact
            assert app.source_label.cget('text') == 'Local R2'
            assert str(path) in app._path_tooltip.text and version in app._path_tooltip.text
            app.file_label.event_generate('<Enter>')
            pump(lambda: app._path_tooltip.tipwindow is not None)
            tooltip = app._path_tooltip.tipwindow
            assert any(version in str(w.cget('text')) for w in tooltip.winfo_children())
            app.file_label.event_generate('<Leave>')
            # Exercise the actual copy button without touching the shared Windows clipboard.
            with patch.object(app, 'clipboard_clear') as clear, patch.object(app, 'clipboard_append') as copy:
                app.btn_copy_path.invoke()
                clear.assert_called_once_with()
                copy.assert_called_once_with(str(path))
            assert app.selection.details['locations'], 'Fixture requires a known lot location'
            for event in app.selection.details['locations']:
                expected = sources.compact_trace_label(event['path'], app._selected_directories)
                assert expected in app.txt.get('1.0', 'end')
            app.result_switch._buttons_dict['Synthèse'].invoke()
            pump(lambda: app.result_mode.get() == 'Synthèse')
            summary_text = '\n'.join(str(w.cget('text')) for w in descendants(app.summary_frame)
                                     if isinstance(w, gui.ctk.CTkLabel))
            assert 'Derniers emplacements' in summary_text and expected in summary_text
            assert 'OMEGA conservée' not in summary_text
            assert all(getattr(w, '_selectable_display', False) for w in descendants(app.summary_frame)
                       if isinstance(w, gui.tk.Text))
            app.result_switch._buttons_dict['Classique'].invoke()
            # Same pathname, different month and dimensions. The old selection
            # must still open its own snapshot even before the next refresh.
            path.write_bytes(original.replace('14/10/2026', '14/11/2026').replace('l=48', 'l=88').encode('cp1252'))
            double_click('get')
            pump(lambda: len([row for row in app.store.metadata() if row['kind'] == 'omega']) == 2)
            pump(lambda: app.selection is not None and app.selection.version_id == version)
            pump(lambda: 'Historique' in app._path_tooltip.text)
            assert 'Historique' not in app.file_label.cget('text')
            double_click('pute')
            path.unlink()
            pump(lambda: all(row['path'].historical for row in app.store.metadata()))
            double_click('get')
            # Navigate to the other generation and back through the real button.
            app.btn_prev_file.invoke()
            pump(lambda: app.selection is not None and app.selection.version_id != version)
            app.btn_next_file.invoke()
            pump(lambda: app.selection is not None and app.selection.version_id == version)
            double_click('pute')
            app.btn_tracking_robot.invoke()
            app.keyword.delete(0, 'end')
            app.keyword.insert(0, 'A0061574AC')
            app.btn_search.invoke()
            pump(lambda: 'PUTE' in app.tracking_view.export_text())
            tracking = app.tracking_view.export_text()
            assert 'R2 \\TracOMEGA.txt — 14/10/2026 – 14/10/2026' in tracking
            assert 'OMEGA conservée' not in tracking and 'Historique' not in tracking
            assert version[:8] not in tracking and 'trace r2' not in tracking
            print('OK libelles compacts : en-tete, badge, infobulle reelle, bouton copie (presse-papiers simule), Classique/Synthese et suivi')
            print('OK OMEGA: recherche, generation conservee apres rotation/suppression, navigation, vrais doubles-clics GET/PUTE, texte integral et surlignages')
        finally:
            app._close_indexed_app()
            app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)


if __name__ == '__main__':
    run_gui_check(main)
