"""Real card clicks, isolated file queries, rendered colors and responsive layout."""
from contextlib import ExitStack
import math
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
from tools.native_test_capture import widget_image, color_pixels
from tools.ui_test_support import run_gui_check
from window_placement import monitor_work_area, native_window_bounds


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def records(robot):
    rows = []
    def row(ident, col, hour, negative=False):
        values = ['0']*121
        values[0], values[col] = '13|'+hour, ('-' if negative else '')+str(ident)
        rows.append(';'.join(values))
    for ident in range(1000, 1060):
        if ident==1055:
            row(ident, 2, '10:02:00', True)
        elif ident==1056:
            row(ident, 1, '10:02:00')  # Input without a documented checkpoint.
        else:
            row(ident, 2, '10:00:00')
            if ident==1057:
                row(ident, 114, '10:01:00')
            else:
                for col, hour in ((31, '10:00:15'), (35, '10:00:30'),
                                  (36 if robot=='R1' else 37, '10:01:00'),
                                  (57 if robot=='R1' else 58, '10:01:30')):
                    row(ident, col, hour)
                row(ident, 6 if robot=='R1' else 7, '10:02:00', True)
    return '\n'.join(rows)+'\n'


def main():
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        local, desktop = base/'traces', base/'Desktop'
        files = []
        for root, robot in ((local, 'R1'), (local, 'R2'), (desktop/'traces', 'R1')):
            path = root/robot/'TracMODBUS.txt'
            path.parent.mkdir(parents=True)
            path.write_text(records(robot), encoding='utf8')
            files.append(str(path))
        stack.enter_context(patch.object(sources, 'application_directory', return_value=str(base)))
        stack.enter_context(patch.object(sources, 'desktop_directory', return_value=str(desktop)))
        stack.enter_context(patch.object(sources, 'FIXED_DIRECTORY', str(base/'missing')))
        stack.enter_context(patch.dict(os.environ, LOCALAPPDATA=str(base/'AppData')))
        app = gui.SearchApp()
        errors, views = [], []
        app.report_callback_exception = lambda *error: errors.append(error)
        releases = []
        output = Path(__file__).resolve().parents[1]/'.benchmark'
        output.mkdir(exist_ok=True)

        def pump(condition=lambda: True):
            deadline = time.monotonic()+12
            while time.monotonic()<deadline:
                app.update()
                if errors:
                    raise errors[0][1]
                if condition():
                    return
                time.sleep(.005)
            raise AssertionError('Conveyor cards timeout: '+app.status.get())

        def windows():
            return [w for w in app.winfo_children() if getattr(w, 'convoyage_errors_view', None)]

        def index(path, ident):
            return next(i for i, row in enumerate(view.errors) if row==(path, str(ident)))

        def card_at(i):
            view.cards.see(i); pump()
            return next(card for card, item in view.cards.cards
                        if card.index==i and view.cards.canvas.itemcget(item, 'state')!='hidden')

        def click(path, ident, wait=True):
            i = index(path, ident)
            card = card_at(i)
            before = view.cards.canvas.yview()
            card.title._label.event_generate('<ButtonPress-1>', x=12, y=10)
            card.title._label.event_generate('<ButtonRelease-1>', x=12, y=10)
            if wait:
                pump(lambda: view.selected_index==i and view.state=='ready')
            assert view.cards.canvas.yview()==before, 'Click must keep the list scroll position'
            return i

        def color(card, role):
            assert card.cget('fg_color')==app._palette[role+'_bg']
            assert color_pixels(widget_image(card.value.selection_text), app._palette[role+'_text'])>3

        def snapshot(name):
            pump()
            image = widget_image(view)
            assert color_pixels(image, app._palette['text'])>50
            if app._appearance=='Dark':
                assert color_pixels(image, '#ffffff')<1000, 'No stale white containers in Dark mode'
            image.save(output/name)

        def inside():
            _, bounds = native_window_bounds(view.winfo_toplevel())
            work = monitor_work_area(app)
            assert work[0]<=bounds[0]<bounds[2]<=work[2] and work[1]<=bounds[1]<bounds[3]<=work[3]

        try:
            pump(lambda: getattr(app, '_source_dialog', None) is not None)
            for checkbox in app._source_dialog.checkboxes:
                checkbox._canvas.event_generate('<Button-1>', x=8, y=8)
            app._source_dialog.open_button.invoke()
            pump(lambda: app._indexed_complete and app._sources_verified)
            app.btn_convoyage_menu.invoke()
            # Legacy popup entry point. The new page is exercised by check_convoyage_workspace.
            app.convoyage_errors_action(); pump(lambda: bool(windows()))
            view = windows()[0].convoyage_errors_view; views.append(view)
            pump(lambda: view.state=='ready')
            assert len(view.errors)==180 and view.selected_index==0
            assert '3 fichiers' in view.summary.cget('text')
            assert view.panes.sash_coord(0)[0]==340 and view._vertical is False
            inside()
            initial = app.selection
            detail_widgets = (str(view.positive), str(view.waiting), str(view.steps_frame))
            with patch.object(app.store, 'convoyage_summary', wraps=app.store.convoyage_summary) as summary, \
                 patch.object(app.store, 'convoyage_steps', wraps=app.store.convoyage_steps) as steps:
                for path in files:
                    click(path, 1000)
                    assert view.data[0][0]['_filepath']==path
                    assert all(call.args[1]==path for call in summary.call_args_list[-1:])
                    assert all(call.args[1]==path for call in steps.call_args_list[-1:])
                    assert len(view.data[1])==6
                    assert ('T5 / R2' if path==files[1] else 'T5 / R1') in view.waiting.value.cget('text')
                    assert len(windows())==1 and not getattr(app, 'convoyage_window', None)
                assert 'Bureau' in card_at(view.selected_index).source.cget('text')
                before = summary.call_count, steps.call_count
                click(files[2], 1000)
                assert (summary.call_count, steps.call_count)==before
            assert app.selection is initial
            assert detail_widgets==(str(view.positive), str(view.waiting), str(view.steps_frame))
            color(view.positive, 'positive'); color(view.waiting, 'error')
            assert all(row.checkpoint.selection_text.cget('background')==app._palette['surface'] for row in view._step_rows)
            snapshot('convoyage_cards_light_20261019.png')
            # Visible cards are pooled; scrolling changes neither detail nor query.
            before = view.selected_index
            detail_y = view.detail_scroll._parent_canvas.yview()
            view.cards.canvas.event_generate('<MouseWheel>', delta=-120); pump()
            assert view.selected_index==before and view.detail_scroll._parent_canvas.yview()==detail_y
            assert len(view.cards.cards)<=math.ceil(view.cards.canvas.winfo_height()/view.cards.row_height)+4
            view.cards.canvas.focus_force()
            view.cards.canvas.event_generate('<Down>')
            pump(lambda: view.selected_index==before+1 and view.state=='ready')
            # Right-click copy and fragment copy never write the shared clipboard.
            target = card_at(view.selected_index)
            copied = []
            with patch.object(view.menu, 'tk_popup'), patch.object(view, 'clipboard_clear'), \
                 patch.object(view, 'clipboard_append', side_effect=copied.append):
                target.title._label.event_generate('<Button-3>', x=12, y=10)
                view.menu.invoke(0); assert copied[-1]==view.errors[view.selected_index][1]
                view.menu.invoke(2); assert copied[-1]==view.errors[view.selected_index][0]
                view.path_button.invoke(); assert copied[-1]==view.errors[view.selected_index][0]
                view.cards.canvas.focus_force(); view.cards.canvas.event_generate('<Control-c>')
                assert '\t' in copied[-1]
            value = view.waiting.value.selection_text
            with patch.object(value, 'clipboard_clear'), patch.object(value, 'clipboard_append', side_effect=copied.append):
                value.tag_add('sel', '1.0', '1.5'); value.event_generate('<<Copy>>')
                assert copied[-1]==value.get('1.0', '1.5')
            view.provenance_button.invoke(); pump()
            assert view.path_label.winfo_manager() and view.errors[view.selected_index][0] in view.path_label.cget('text')
            click(files[0], 1055)
            assert view.positive.role=='neutral' and view.waiting.role=='waiting'
            assert 'Aucune position positive' in view.positive.value.cget('text')
            click(files[0], 1056)
            assert view.positive.role==view.waiting.role=='neutral' and not view.data[1]
            assert 'Aucune étape indexée' in view.message.cget('text')
            click(files[0], 1057)
            assert view.positive.value.cget('text')=='Unnamed: 114' and view.waiting.role=='neutral'
            # A canceled response cannot replace a newer robot, even after panel navigation.
            started, finished, release = threading.Event(), threading.Event(), threading.Event()
            releases.append(release)
            original = app.store.convoyage_summary
            def slow(ident, path=None):
                if ident=='1002' and path==files[0]:
                    started.set(); release.wait(8)
                try:
                    return original(ident, path)
                finally:
                    if ident=='1002' and path==files[0]: finished.set()
            with patch.object(app.store, 'convoyage_summary', slow):
                click(files[0], 1002, wait=False); pump(started.is_set)
                click(files[1], 1003)
                app.btn_search_menu.invoke()
                release.set(); pump(finished.is_set)
                pump()
                assert view.errors[view.selected_index]==(files[1], '1003') and view.state=='ready'
                assert view.data[0][0]['_filepath']==files[1]
            # The latest request survives a main-panel change while it is pending.
            started.clear(); finished.clear(); release.clear()
            with patch.object(app.store, 'convoyage_summary', slow):
                click(files[0], 1002, wait=False); pump(started.is_set)
                app.btn_convoyage_menu.invoke()
                release.set(); pump(lambda: view.state=='ready')
                assert view.data[0][0]['_filepath']==files[0]
            # Failure is local, neutral and recoverable through the real button.
            with patch.object(app.store, 'convoyage_summary', side_effect=OSError('Test indisponible')):
                click(files[0], 1004, wait=False); pump(lambda: view.state=='error')
            assert 'Test indisponible' in view.message.cget('text') and view.positive.role=='neutral'
            view.retry.invoke(); pump(lambda: view.state=='ready')
            click(files[0], 1000)
            view.provenance_button.invoke() if view.path_label.winfo_manager() else None
            app._change_theme('Sombre'); pump()
            color(view.positive, 'positive'); color(view.waiting, 'error')
            snapshot('convoyage_cards_dark_20261019.png')
            view.winfo_toplevel().geometry('780x740'); pump()
            assert view._vertical and str(view.panes.cget('orient'))=='vertical'
            assert len(view.cards.cards)<16
            snapshot('convoyage_cards_narrow_20261019.png')
            view.winfo_toplevel().geometry('560x640'); pump()
            assert view.waiting.grid_info()['row']==1
            assert view.positive.winfo_width()>view.states.winfo_width()*.9
            snapshot('convoyage_cards_small_20261019.png')
            left = view.cards.canvas.yview()
            view.waiting.value.selection_text.event_generate('<MouseWheel>', delta=-120); pump()
            assert view.detail_scroll._parent_canvas.yview()[0]>0 and view.cards.canvas.yview()==left
            view.winfo_toplevel().geometry('1180x740'); pump()
            assert not view._vertical and view.panes.sash_coord(0)[0]==340
            # Close an actual in-flight query and verify cancellation/cleanup.
            started.clear(); finished.clear(); release.clear()
            with patch.object(app.store, 'convoyage_summary', slow):
                click(files[0], 1002, wait=False); pump(started.is_set)
                view.close_button.invoke(); pump()
                assert view._closed and view.tasks.closed
                assert view.tasks.latest[view.channel].is_set()
                release.set(); pump(finished.is_set)
            # Empty and very large lists: no invented detail or per-result widgets.
            app.btn_convoyage_menu.invoke()
            with patch.object(app.store, 'convoyage_errors', return_value=[]):
                app.convoyage_errors_action(); pump(lambda: bool(windows()))
            empty = windows()[0].convoyage_errors_view; views.append(empty)
            assert empty.state=='empty' and not empty.cards.cards and empty.selected_index is None
            empty.close_button.invoke()
            app._render_convoyage_errors([(files[0], str(i)) for i in range(10000)])
            large = windows()[0].convoyage_errors_view; views.append(large)
            pump(lambda: large.state=='ready')
            assert len(large.cards.cards)<16
            large.close_button.invoke(); pump()
            assert large.tasks.closed
            print('OK convoyage cartes : 180 anomalies/3 sources robot, fichier exact des deux requetes, '
                  'pas de popup, selection/clavier/copie, widgets conserves, vert/rouge rendus Clair/Sombre, '
                  'responsive 1180/780/560, scrolls independants, etats manquants/registres inconnus, '
                  'reponses depassees/erreur/reprise/fermeture et 10000 lignes virtualisees.')
        finally:
            for release in releases: release.set()
            app._close_indexed_app(); app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)
            for item in views:
                item.tasks.executor.shutdown(wait=True, cancel_futures=True)


if __name__=='__main__':
    run_gui_check(main)
