"""Real dual robot popup interactions, exclusively on a private Windows desktop."""

from contextlib import ExitStack
import os
from pathlib import Path
import sys
import tempfile
import threading
import time
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import search_gui as gui
import trace_sources as sources
from tests.test_paired_adjacent import ticket, trace
from tools.summary_test_support import rendered_summary_text, click_summary
from tools.ui_test_support import run_gui_check
from window_placement import monitor_work_area, native_window_bounds


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def main():
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        files, texts = {}, {}
        for robot, offset in (('R1', 0), ('R2', 2)):
            path = base / 'traces' / robot / 'TracOmega019.old'
            path.parent.mkdir(parents=True)
            text = trace(*(ticket(36000 + offset + i * 10, str(10 + i),
                                  robot + ('_MAIN' if i == 0 else '_N' + str(i)),
                                  lot=robot + '_LOT', sn=robot + '_SN') for i in range(-5, 6)),
                         ticket(50000, '90', 'R1_FAR') if robot == 'R1' else '')
            path.write_text(text, encoding='utf8')
            files[robot], texts[robot] = path, text
        for name, body in (
            ('TracOmega005.old', trace(*(ticket(35000+i*10, str(i), 'BOUNDARY_'+str(i)) for i in range(439,447)))),
            ('TracOmega006.old', trace(*(ticket(40000+i*10, str(i), 'BOUNDARY_'+str(i)) for i in range(511,517)), previous='TracOmega005.old'))):
            (base/'traces'/'R1'/name).write_text(body, encoding='utf8')
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
            raise AssertionError('Dual popup timeout: ' + app.status.get())

        def search(reference):
            prior = app.selection
            app.mode_radios['ref'].invoke()
            app.keyword.delete(0, 'end')
            app.keyword.insert(0, reference)
            app.btn_search.invoke()
            pump(lambda: app.selection is not None and app.selection is not prior)

        def ready():
            return (getattr(app, 'adjacent_view', None) and app.adjacent_view.data is not None
                    and not str(app.adjacent_view.note.cget('text')).startswith('Actualisation'))

        def columns():
            return {column['robot']: column for column in app.adjacent_view.data['columns']}

        def click_row(robot, predicate):
            view = app.adjacent_view
            tree = view.trees[robot]
            key, target = next((key, row) for key, row in view.rows[robot].items() if predicate(row))
            view._yview('moveto', 0)
            tree.see(key)
            pump()
            box = tree.bbox(key)
            assert box, (robot, key)
            x, y = box[0] + 12, box[1] + box[3] // 2
            tree.event_generate('<ButtonPress-1>', x=x, y=y)
            tree.event_generate('<ButtonRelease-1>', x=x, y=y)
            pump(lambda: app.selection is not None and app.selection.path == target['path']
                 and app.selection.number == target['number'] and ready()
                 and app.adjacent_view.data['selected']['path'] == target['path'])
            return target

        def popup_bounds():
            pump()
            _, rect = native_window_bounds(app.adj_popup)
            work = monitor_work_area(app)
            assert work[0] <= rect[0] < rect[2] <= work[2], (rect, work)
            assert work[1] <= rect[1] < rect[3] <= work[3], (rect, work)
            maximum = app.adj_popup.maxsize()
            assert maximum[0] <= work[2] - work[0] and maximum[1] <= work[3] - work[1]

        release = threading.Event()
        try:
            pump(lambda: app._sources_verified)
            sync = app.store.sync
            stack.enter_context(patch.object(app.store, 'sync', return_value=dict(changed=0, removed=0, failures=[])))
            search('R1_MAIN')
            selected = app.selection
            area = monitor_work_area(app)
            for x, y in ((area[0], area[1]), (area[2] - 600, area[3] - 400)):
                app.geometry(f'600x400+{max(0, x)}+{max(0, y)}')
                pump()
                app.btn_adjacent.invoke()
                pump(ready)
                popup_bounds()
                assert app.selection is selected
                assert columns()['R1']['anchor']['ref'] == 'R1_MAIN'
                assert columns()['R2']['anchor']['ref'] == 'R2_MAIN'
                assert columns()['R2']['delta'] == 2
                assert len(app.adjacent_view.rows['R1']) == len(app.adjacent_view.rows['R2']) == 11
                for robot, tree in app.adjacent_view.trees.items():
                    assert tree.item('5', 'values')[1:5] == ('10', robot + '_MAIN', robot + '_SN', robot + '_LOT')
                    assert ('current' if robot == 'R1' else 'nearest') in tree.item('5', 'tags')
                app.adjacent_view.close_button.invoke()
                pump(lambda: not app.adj_popup.winfo_exists())
            app.btn_adjacent.invoke()
            pump(ready)
            click_row('R2', lambda row: row['ref'] == 'R2_MAIN')
            assert app.adjacent_view.data['selected_robot'] == 'R2'
            assert columns()['R1']['delta'] == -2
            # The two fixed-height tables scroll together; the canvas is horizontal only.
            app.adj_popup.geometry('780x350')
            pump()
            view = app.adjacent_view
            initial = view.trees['R1'].yview()
            view.trees['R1'].event_generate('<MouseWheel>', delta=-120)
            pump()
            assert view.trees['R1'].yview()[0] > initial[0]
            assert view.trees['R1'].yview()[0] == view.trees['R2'].yview()[0]
            assert view.canvas.xview()[1] < 1
            app.adj_popup.geometry('1180x720')
            pump()
            view._yview('moveto', 0)
            # Copy commands are real; intercept only native menu display/clipboard writes.
            tree = view.trees['R2']
            key = '5'
            x, y, width, height = tree.bbox(key, 'sn')
            copied = []
            with patch.object(view.menu, 'tk_popup'), patch.object(view, 'clipboard_clear'), \
                    patch.object(view, 'clipboard_append', side_effect=copied.append):
                tree.event_generate('<Button-3>', x=x + width // 2, y=y + height // 2)
                view.menu.invoke(0)
                assert copied[-1] == 'R2_SN'
                tree.focus_set()
                tree.event_generate('<Control-c>')
                assert 'R2_LOT' in copied[-1]
                view.menu.invoke(2)
                assert str(files['R2']) in copied[-1] and 'Ligne ' in copied[-1]
            app._change_theme('Sombre')
            assert app.adjacent_view.data['selected_robot'] == 'R2'
            assert view.canvas.cget('background') == app._palette['bg']
            for panel in view.panels:
                for name in ('status','source'):
                    label=panel[name]
                    assert label.cget('background') == app._palette['surface']
                    assert label.selection_text.cget('foreground') == label.cget('foreground')
            # The popup also identifies partial results without disabling existing snapshots.
            verified = app._sources_verified
            app._sources_verified = False
            app.btn_adjacent.invoke()
            pump(ready)
            assert 'Vérification en cours' in view.note.cget('text')
            app._sources_verified = verified
            search('R1_FAR')
            pump(ready)
            assert abs(columns()['R2']['delta']) > 300 and columns()['R2']['anchor']
            messages = ' '.join(str(w.cget('text')) for w in descendants(view) if isinstance(w, gui.tk.Label))
            assert 'Aucune occurrence à moins de 5 minutes' in messages
            search('R1_MAIN')
            pump(ready)
            # Rotate the same R2 path into a second generation at the same instant.
            files['R2'].write_text(texts['R2'].replace('R2_MAIN', 'R2_NEW'), encoding='utf8')
            assert not sync([str(path) for path in files.values()])['failures']
            files['R2'].unlink()
            assert not sync([str(files['R1'])])['failures']
            app.btn_adjacent.invoke()
            pump(ready)
            assert columns()['R2']['anchor'] is None
            assert {row['ref'] for row in columns()['R2']['candidates']} == {'R2_MAIN', 'R2_NEW'}
            assert all(row['path'].historical for row in columns()['R2']['candidates'])
            target = click_row('R2', lambda row: row['ref'] == 'R2_MAIN')
            assert target['path'].historical and app.selection.path.version_id == target['path'].version_id
            assert app.selection.details['get']['text'].endswith('Code:R2_MAIN $0101')
            # Actual GET double-click after rotation/deletion, checking complete saved content and tag.
            pump(lambda: '_summary_comparison' in app.selection.details)
            app.geometry('1200x900+0+0')
            pump()
            widget = app.summary_targets['get'].selection_text
            widget.see('1.0')
            pump()
            box = widget.bbox('1.5')
            assert box
            x, y = box[0] + 2, box[1] + 2
            for i in range(2):
                widget.event_generate('<ButtonPress-1>', x=x, y=y, time=20000 + i * 100)
                widget.event_generate('<ButtonRelease-1>', x=x, y=y, time=20020 + i * 100)
            def trace_windows():
                return [w for w in app.winfo_children() if isinstance(w, gui.tk.Toplevel) and w.title().startswith('Vue de')]
            pump(lambda: bool(trace_windows()) and any(isinstance(w, gui.tk.Text) and w.tag_ranges('highlight')
                                                      for w in descendants(trace_windows()[0])))
            text = next(w for w in descendants(trace_windows()[0]) if isinstance(w, gui.tk.Text))
            assert text.get('1.0', 'end-1c').splitlines() == texts['R2'].splitlines()
            event = app.selection.details['get']
            assert text.index(text.tag_ranges('highlight')[0]) == f'{event["number"] + 1}.0'
            trace_windows()[0].destroy()
            # A worker from the old selection cannot populate the dual popup.
            started, finished = threading.Event(), threading.Event()
            paired = app.store.paired_adjacent
            def slow(*args, **kwargs):
                started.set()
                release.wait(5)
                try:
                    return paired(*args, **kwargs)
                finally:
                    finished.set()
            with patch.object(app.store, 'paired_adjacent', slow):
                app.btn_adjacent.invoke()
                pump(started.is_set)
                app.keyword.delete(0, 'end')
                app.keyword.insert(0, 'ABSENT')
                app.btn_search.invoke()
                pump(lambda: app.selection is None and 'Aucun' in rendered_summary_text(app))
                release.set()
                pump(finished.is_set)
                pump()
                assert app.adjacent_view.data is None and not app.adjacent_view.rows
            search('R1_MAIN')
            pump(ready)
            # Closing while a result is pending must not reopen the window.
            release.clear()
            started.clear()
            finished.clear()
            with patch.object(app.store, 'paired_adjacent', slow):
                app.btn_adjacent.invoke()
                pump(started.is_set)
                popup = app.adj_popup
                app.adjacent_view.close_button.invoke()
                release.set()
                pump(finished.is_set)
                pump()
                assert not popup.winfo_exists()
            search('BOUNDARY_444')
            app.btn_adjacent.invoke(); pump(ready)
            for ident in ('445', '446', '511'):
                click_row('R1', lambda row, ident=ident: row['modbus']==ident)
            assert [r['modbus'] for r in columns()['R1']['previous']]==['442','443','444','445','446']
            assert [r['modbus'] for r in columns()['R1']['following']]==['512','513','514','515','516']
            assert app.selection.path.path.endswith('TracOmega006.old')
            assert '5 précédents / 5 demandés' in app.adjacent_view.panels[0]['footer'].cget('text')
            print('OK R1/R2 : vrais boutons/clics et 444/445/446/511 entre archives, 11 voisins independants, champs SN/lot, '
                  'recalage, ecart >5 min, generations candidates et archive supprimee, GET complet/surlignage, '
                  'copie, theme, scroll commun, placement aux bords, reponses obsoletes et fermeture')
        finally:
            release.set()
            app._close_indexed_app()
            app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)


if __name__ == '__main__':
    run_gui_check(main)
