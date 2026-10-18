"""Check only the new source chooser and selected-source search on a private desktop."""

from contextlib import ExitStack
import ctypes
from ctypes import wintypes
import os
from pathlib import Path
import sys
import tempfile
import time
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import search_gui
import trace_sources
import trace_store
from tools.ui_test_support import run_gui_check


def main():
    # Verify that the OS placed this thread on the requested private desktop.
    user = ctypes.WinDLL('user32', use_last_error=True)
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    user.GetThreadDesktop.argtypes = [wintypes.DWORD]
    user.GetThreadDesktop.restype = wintypes.HANDLE
    user.GetUserObjectInformationW.argtypes = [wintypes.HANDLE, ctypes.c_int, ctypes.c_void_p,
                                              wintypes.DWORD, ctypes.POINTER(wintypes.DWORD)]
    name = ctypes.create_unicode_buffer(256)
    needed = wintypes.DWORD()
    desktop = user.GetThreadDesktop(kernel.GetCurrentThreadId())
    assert user.GetUserObjectInformationW(desktop, 2, name, ctypes.sizeof(name), ctypes.byref(needed))
    assert name.value == os.environ['TRACOMEGA_TEST_DESKTOP']
    print('OK bureau Windows isole confirme ; aucun SwitchDesktop ni activation du bureau utilisateur')
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        installed, project, desktop_root = base / 'Apoteka', base / 'Project', base / 'Desktop'
        for root, marker in ((installed, 'INSTALLED'), (project / 'traces', 'PROJECT'),
                             (desktop_root / 'traces', 'DESKTOP')):
            for robot in ('R1', 'R2'):
                directory = root / robot
                directory.mkdir(parents=True)
                (directory / 'TracOmega.txt').write_text(
                    f'RefProduit={marker}<Cr>ModBusId=123<Cr>\n', encoding='utf8')
        stack.enter_context(patch.object(trace_sources, 'FIXED_DIRECTORY', str(installed)))
        stack.enter_context(patch.object(trace_sources, 'application_directory', return_value=str(project)))
        stack.enter_context(patch.object(trace_sources, 'desktop_directory', return_value=str(desktop_root)))
        stack.enter_context(patch.dict(os.environ, LOCALAPPDATA=str(base / 'Cache')))
        app = search_gui.SearchApp()
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
            raise AssertionError('Attente de la selection des sources expiree')
        try:
            pump(lambda: app._source_dialog is not None)
            dialog = app._source_dialog
            assert len(dialog.checkboxes) == 3
            assert dialog.open_button.cget('state') == 'disabled'
            assert not hasattr(app, 'store')
            assert not list((base / 'Cache').rglob('*.sqlite3'))
            dialog.checkboxes[2]._canvas.event_generate('<Button-1>', x=8, y=8)
            pump(lambda: bool(dialog.selected_directories()))
            assert dialog.selected_directories() == (str(desktop_root / 'traces'),)
            dialog.open_button.invoke()
            pump(lambda: app._indexed_complete)
            expected = {str(desktop_root / 'traces' / robot / 'TracOmega.txt') for robot in ('R1', 'R2')}
            assert set(app.file_list) == expected
            assert app.store.database == trace_store.cache_path([str(desktop_root / 'traces')])
            app.modbus.set(False)
            app.reference.set(True)
            app.keyword.insert(0, 'DESKTOP')
            app.btn_search.invoke()
            pump(lambda: app.selection is not None)
            assert {path for path, hits in app._groups} == expected
            assert 'Bureau R' in app.file_label.cget('text')
            assert not app.store.search('PROJECT') and not app.store.search('INSTALLED')
            print('OK choix par vrais boutons : Bureau seul, R1/R2, recherche et provenance ; autres sources absentes')
        finally:
            app._close_indexed_app()
            app._index_thread.join(3)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)


if __name__ == '__main__':
    run_gui_check(main)
