"""Exercise the actual Tk actions with temporary traces and a temporary index."""

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
from trace_store import TraceStore


def main():
    fixtures = Path(__file__).resolve().parents[1] / 'tests' / 'fixtures'
    with tempfile.TemporaryDirectory() as folder:
        base = Path(folder)
        for robot, fixture in [('trace r1', 'current_r1.txt'), ('trace r2', 'current_r2.txt')]:
            directory = base / 'traces' / robot
            directory.mkdir(parents=True)
            text = read_trace(fixtures / fixture)
            (directory / 'TracOMEGA.txt').write_bytes(text.encode('cp1252'))
            values = ['0'] * 121
            values[0] = '14|06:27:47'; values[2] = '1194'; values[6] = '1194'; values[-1] = ''
            (directory / 'TracMODBUS.txt').write_text(';'.join(values) + '\n', encoding='utf8')
        old = read_trace(fixtures / 'current_r2.txt').replace('14/10/2026', '13/10/2026').replace('14|', '13|').replace('l=48 H=21 L=47', 'l=55 H=21 L=47')
        (base / 'traces' / 'trace r2' / 'TracOmega029.old').write_bytes(old.encode('cp1252'))
        archive_gate = threading.Event()
        original_index = TraceStore._index_snapshot
        def pause_archive(store, path, *args):
            if path.lower().endswith('.old'):
                archive_gate.wait(5)
            return original_index(store, path, *args)
        with patch.object(sources, 'application_directory', return_value=folder), \
             patch.object(sources, 'FIXED_DIRECTORY', str(base / 'absent')), \
             patch.object(TraceStore, '_index_snapshot', pause_archive), \
             patch.object(indexed_ui, 'TraceStore', side_effect=lambda: TraceStore(base / 'index.db')):
            errors = []
            before = time.perf_counter()
            app = gui.SearchApp(); app.withdraw()
            print('Construction interface:', round(time.perf_counter() - before, 3), 's', flush=True)
            app.report_callback_exception = lambda *error: errors.append(error)
            def pump(condition=lambda: True, seconds=5):
                deadline = time.monotonic() + seconds
                while time.monotonic() < deadline:
                    app.update()
                    for widget in app.winfo_children():
                        if isinstance(widget, gui.tk.Toplevel): widget.withdraw()
                    if errors:
                        raise errors[0][1]
                    if condition(): return
                    time.sleep(.01)
                raise AssertionError('UI operation did not complete: ' + app.status.get())
            try:
                pump(lambda: len(app.file_list) == 4)
                assert not app._indexed_complete
                app.keyword.insert(0, '1194'); app.modbus.set(True); app.on_modbus_check()
                app.perform_search()
                pump(lambda: bool(getattr(app, '_display_data', None)))
                assert 'partiels' in app.status.get()
                assert not app.files_loaded
                print('Recherche disponible pendant la premiere indexation: OK', flush=True)
                archive_gate.set()
                pump(lambda: app._indexed_complete)
                app.keyword.delete(0, 'end'); app.keyword.insert(0, '1194'); app.modbus.set(True); app.on_modbus_check()
                app.perform_search()
                pump(lambda: bool(getattr(app, '_display_data', None)))
                assert '<Cr>' not in app.txt.get('1.0', 'end')
                assert 'GRIPPER GET' in app.txt.get('1.0', 'end')
                app.btn_adjacent.invoke()
                pump(lambda: app.adj_popup is not None)
                def titles():
                    return [widget.title() for widget in app.winfo_children() if isinstance(widget, gui.tk.Toplevel)]
                app.compare_previous_idpk_popup(); pump(lambda: any(title.startswith('Comparaison IDPK') for title in titles()))
                app.open_gripper_line_view('GRIPPER GET'); pump(lambda: any(title.startswith('Vue de') for title in titles()))
                views = sum(title.startswith('Vue de') for title in titles())
                app.open_file_view(); pump(lambda: sum(title.startswith('Vue de') for title in titles()) > views)
                app.convoyage_action('1194'); pump(lambda: bool(getattr(app, 'convoyage_window', None)))
                app.show_convoyage_details('1194'); app.convoyage_errors_action(); app.check_errors()
                pump(lambda: any(title.startswith('Parcours détaillé') for title in titles())
                     and 'Erreurs Convoyage' in titles() and 'Vérification erreurs Robot' in titles())
                app.show_tracking_robot_panel(); app.tracking_keyword.insert(0, 'A0061574AC')
                app.perform_tracking_search()
                pump(lambda: '14/10/26' in app.tracking_txt.get('1.0', 'end'))
                app.show_search_panel(); app.keyword.delete(0, 'end'); app.keyword.insert(0, 'A0061574AC')
                app.recherche_avancee_criteria('A0061574AC', sn='343272', codelot='261013.127')
                pump(lambda: len(app._groups) == 2 and getattr(app, '_active_mode', None) == 'advanced')
                app.on_right(); pump(lambda: 'TracOmega029.old' in app.file_label.cget('text'))
                app.on_left(); pump(lambda: 'TracOMEGA.txt' in app.file_label.cget('text'))
                # Superseded searches and view changes must not redraw a deleted widget.
                app.keyword.delete(0, 'end'); app.keyword.insert(0, 'REF_ABSENTE_999'); app.perform_search()
                app.show_tracking_robot_panel(); pump(seconds=.2)
                print('UI OK: recherche, IDs adjacents, IDPK, vues, convoyage, erreurs robot, tracking, recherche avancee, navigation et changement de vue', flush=True)
            finally:
                archive_gate.set()
                app._close_indexed_app()
                app._index_thread.join(timeout=5)
                app.tasks.executor.shutdown(wait=True, cancel_futures=True)


if __name__ == '__main__':
    main()
