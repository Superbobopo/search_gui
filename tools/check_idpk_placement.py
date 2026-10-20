"""Exercise the comparison button and native popup bounds on a private desktop."""

from contextlib import ExitStack
from ctypes import wintypes
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
from window_placement import monitor_work_area, native_window_bounds, windows_api
from tools.ui_test_support import run_gui_check


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def main():
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        directory = base / 'traces' / 'trace r2'
        directory.mkdir(parents=True)
        text = read_trace(Path(__file__).resolve().parents[1] / 'tests/fixtures/current_r2.txt')
        (directory / 'TracOMEGA.txt').write_text(text, encoding='utf8')
        old = text.replace('14/10/2026', '13/10/2026').replace('14|', '13|')
        old = old.replace('l=48 H=21 L=47', 'l=55 H=21 L=47')
        (directory / 'TracOmega029.old').write_text(old, encoding='utf8')
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
                time.sleep(.01)
            raise AssertionError('IDPK check timed out')

        try:
            pump(lambda: app._sources_verified)
            app.keyword.insert(0, 'A0061574AC')
            app.modbus.set(False)
            app.reference.set(True)
            app.btn_search.invoke()
            pump(lambda: app.selection is not None)
            area = monitor_work_area(app)
            # Both left-edge and bottom-right placements reproduced the old defect.
            for position in ((area[0], area[1]), (area[2] - 600, area[3] - 400)):
                app.geometry(f'600x400+{max(0, position[0])}+{max(0, position[1])}')
                pump()
                selected = app.selection
                app.btn_compare_idpk.invoke()
                def comparisons():
                    return [w for w in app.winfo_children() if isinstance(w, gui.tk.Toplevel)
                            and w.title().startswith('Comparaison IDPK')]
                pump(lambda: bool(comparisons()) and comparisons()[0].winfo_ismapped())
                popup = comparisons()[0]
                # Let map/configure events settle before checking the whole decorated window.
                for _ in range(5):
                    pump()
                _, rect = native_window_bounds(popup)
                work = monitor_work_area(app)
                assert rect[0] >= work[0] and rect[1] >= work[1], (rect, work)
                assert rect[2] <= work[2] and rect[3] <= work[3], (rect, work)
                assert rect[2] - rect[0] >= 425 and rect[3] - rect[1] >= 255, rect
                assert str(popup.transient()) == str(app)
                user = windows_api()
                user.GetWindow.argtypes = [wintypes.HWND, wintypes.UINT]
                user.GetWindow.restype = wintypes.HWND
                popup_handle, _ = native_window_bounds(popup)
                parent_handle, _ = native_window_bounds(app)
                above_parent = []
                handle = user.GetWindow(parent_handle, 3)  # GW_HWNDPREV
                while handle:
                    above_parent.append(handle)
                    handle = user.GetWindow(handle, 3)
                assert popup_handle in above_parent, 'Comparison must be above its parent'
                labels = ' '.join(str(w.cget('text')) for w in descendants(popup)
                                  if isinstance(w, gui.tk.Label))
                assert 'l = 48' in labels and 'l = 55' in labels and 'Δ l = -7' in labels
                assert app.selection is selected
                close = popup.diagnostic_view.close_button
                close.invoke()
                pump(lambda: not comparisons())
            print('OK IDPK : vrai bouton, fenetre entiere dans la zone utile aux bords, '
                  'parent transient, mesures/delta, selection et bouton Fermer')
        finally:
            app._close_indexed_app()
            app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)


if __name__ == '__main__':
    run_gui_check(main)
