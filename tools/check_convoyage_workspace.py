"""Actual events and rendered snapshots of the embedded conveyor workspace."""
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
from tools.check_convoyage_cards import records
from tools.ui_test_support import run_gui_check
from tools.native_test_capture import widget_image, color_pixels


def main():
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        desktop = base/'Desktop'
        paths = []
        for root, robot in ((base/'traces', 'trace r1'), (base/'traces', 'R2'), (desktop/'traces', 'R1')):
            path = (root/'R1'/'copies' if robot=='R2' else root/robot)/'TracMODBUS.txt'
            path.parent.mkdir(parents=True)
            path.write_text(records('R2' if robot=='R2' else 'R1'), encoding='utf8')
            paths.append(str(path))
        stack.enter_context(patch.object(sources, 'application_directory', return_value=str(base)))
        stack.enter_context(patch.object(sources, 'desktop_directory', return_value=str(desktop)))
        stack.enter_context(patch.object(sources, 'FIXED_DIRECTORY', str(base/'absent')))
        stack.enter_context(patch.dict(os.environ, LOCALAPPDATA=str(base/'AppData')))
        app = gui.SearchApp()
        exceptions, releases, runners = [], [], []
        app.report_callback_exception = lambda *args: exceptions.append(args)
        output = Path(__file__).resolve().parents[1]/'.benchmark'
        output.mkdir(exist_ok=True)

        def pump(condition=lambda: True, timeout=12):
            deadline = time.monotonic()+timeout
            while time.monotonic()<deadline:
                app.update()
                if exceptions:
                    raise exceptions[0][1]
                if condition():
                    return
                time.sleep(.005)
            raise AssertionError('Workspace timeout: '+str(workspace.states if 'workspace' in locals() else app.index_status.get()))

        def settle(seconds=.2):
            end = time.monotonic()+seconds
            pump(lambda: time.monotonic()>=end)

        def capture(name):
            settle()
            image = widget_image(app)
            assert color_pixels(image, app._palette['text'])>200
            if app._appearance=='Dark':
                assert color_pixels(widget_image(workspace), '#ffffff')<1000, 'White containers in Dark mode'
            image.save(output/f'convoyage_workspace_{name}_20261020.png')

        def ready(tab):
            return (workspace.states[tab]['loaded'] and workspace.states[tab]['pending'] is None
                    and (tab not in workspace.views or workspace.views[tab].state in ('ready', 'empty')))

        def search(ident, enter=False):
            workspace.entry.delete(0, 'end')
            workspace.entry.insert(0, ident)
            if enter:
                workspace.entry.focus_force()  # Only the isolated test desktop.
                pump(lambda: app.focus_get() is workspace.entry._entry)
                workspace.entry._entry.event_generate('<Return>')
            else:
                workspace.search_button.invoke()

        def click(view, index):
            if view.parcours:
                view.file_picker._dropdown_menu._button_callback(view.file_picker.cget('values')[index])
                pump(lambda: view.selected_index==index and view.state=='ready')
                return
            view.cards.see(index); settle(.03)
            card = next(card for card, item in view.cards.cards if card.index==index)
            card.title._label.event_generate('<ButtonPress-1>', x=12, y=10)
            card.title._label.event_generate('<ButtonRelease-1>', x=12, y=10)
            pump(lambda: view.selected_index==index and view.state=='ready')

        try:
            pump(lambda: app._source_dialog is not None)
            for checkbox in app._source_dialog.checkboxes:
                checkbox._canvas.event_generate('<Button-1>', x=8, y=8)
            app._source_dialog.open_button.invoke()
            pump(lambda: app._sources_verified and app._indexed_complete)
            app.geometry('1340x880+0+0')
            app.keyword.insert(0, 'SEARCH_SAVED')
            app.btn_convoyage_menu.invoke(); settle()
            workspace = app.convoyage_workspace
            runners.append(workspace.tasks)
            assert workspace.active_tab=='parcours'
            assert not app.search_bar.winfo_ismapped()
            assert workspace.pages['parcours']['title'].cget('text')=='Suivre une boîte'
            assert not workspace.views
            assert workspace.search_button.cget('font')[1]==14
            capture('initial_light')
            app._change_theme('Sombre'); capture('initial_dark')
            app._change_theme('Clair')
            with patch.object(app.store, 'convoyage_summary', wraps=app.store.convoyage_summary) as summary, \
                 patch.object(app.store, 'convoyage_details', wraps=app.store.convoyage_details) as steps, \
                 patch.object(app.store, 'convoyage_errors', wraps=app.store.convoyage_errors) as anomalies:
                search('1000', enter=True)
                pump(lambda: ready('parcours'))
                view = workspace.views['parcours']
                assert len(view.errors)==3 and view.selected_index==0
                assert not view.header.winfo_ismapped() and not view.close_button.winfo_ismapped()
                assert not view._vertical
                for i in range(3):
                    click(view, i)
                    path = view.errors[i][0]
                    assert view.data[0][0]['_filepath']==path
                    assert steps.call_args_list[-1].args==('1000', path)
                    card = next(card for card, _ in view.cards.cards if card.index==i)
                    assert card.robot.cget('text')=='R1/R2'
                view.reset_items(view.errors[:1], force=True)
                view.configure_parcours()
                pump(lambda: view.state=='ready'); settle()
                assert len(view.panes.panes())==1 and not view.file_picker.winfo_ismapped()
                capture('single_file_light')
                view.reset_items([(p, '1000') for p in paths], force=True)
                view.configure_parcours()
                pump(lambda: view.state=='ready')
                click(view, 2)
                capture('parcours_light')
                app._change_theme('Sombre'); capture('parcours_dark')
                app._change_theme('Clair')
                popup_count = len([w for w in app.winfo_children() if isinstance(w, gui.tk.Toplevel)])
                assert popup_count==0
                calls = summary.call_count, steps.call_count
                click(view, 2)
                search('1000'); settle()
                assert calls==(summary.call_count, steps.call_count), 'Identical selection/search recomputed'
                # Anomalies must wait for complete source verification.
                app._sources_verified = False
                workspace.buttons['anomalies'].invoke(); settle(.02)
                assert anomalies.call_count==0
                assert workspace.refresh_button.cget('state')=='disabled'
                app._sources_verified = True
                pump(lambda: ready('anomalies'))
                errors_view = workspace.views['anomalies']
                assert len(errors_view.errors)==180
                assert len(errors_view.cards.cards)<24
                assert '180 anomalies' in workspace.count.cget('text')
                assert color_pixels(widget_image(errors_view.positive.value.selection_text), app._palette['positive_text'])>3
                assert color_pixels(widget_image(errors_view.waiting.value.selection_text), app._palette['error_text'])>3
                capture('anomalies_light')
                click(errors_view, 130)
                errors_view.cards.canvas.yview_moveto(.64)
                errors_view.detail_scroll._parent_canvas.yview_moveto(.15)
                settle()
                selected = errors_view.errors[errors_view.selected_index]
                list_scroll = errors_view.cards.canvas.yview()[0]
                detail_scroll = errors_view.detail_scroll._parent_canvas.yview()[0]
                calls = summary.call_count, steps.call_count, anomalies.call_count
                workspace.buttons['parcours'].invoke(); settle()
                assert view.errors[view.selected_index][0]==paths[2] or view.selected_index==2
                workspace.buttons['anomalies'].invoke(); settle()
                assert calls==(summary.call_count, steps.call_count, anomalies.call_count)
                app.btn_search_menu.invoke(); settle()
                assert app.search_bar.winfo_ismapped() and app.keyword.get()=='SEARCH_SAVED'
                app.btn_tracking_robot.invoke(); settle()
                app.keyword.insert(0, 'TRACK_SAVED')
                app.btn_convoyage_menu.invoke(); settle()
                assert not app.search_bar.winfo_ismapped()
                assert calls==(summary.call_count, steps.call_count, anomalies.call_count)
                assert errors_view.errors[errors_view.selected_index]==selected
                assert abs(errors_view.cards.canvas.yview()[0]-list_scroll)<.01
                assert abs(errors_view.detail_scroll._parent_canvas.yview()[0]-detail_scroll)<.01
                # Theme changes preserve data/widgets and do not execute queries.
                app._change_theme('Sombre'); settle()
                assert calls==(summary.call_count, steps.call_count, anomalies.call_count)
                assert workspace.views['anomalies'] is errors_view
                capture('anomalies_dark')
                app.geometry('1090x820+0+0'); settle()
                assert not errors_view._vertical
                assert len(errors_view.panes.panes())==2
                saved_errors = list(errors_view.errors)
                errors_view.reset_items(saved_errors[:7], force=True)
                workspace.count.configure(text='7 anomalies · 1 fichier')
                pump(lambda: errors_view.state=='ready'); settle()
                assert errors_view.cards.canvas.winfo_height()>=7*errors_view.cards.row_height
                rows = [r for r in errors_view._step_rows if r.winfo_ismapped()]
                assert len(rows)>=5
                assert sum(r.arrow.winfo_ismapped() for r in rows)==len(rows)-1
                assert not rows[-1].arrow.winfo_ismapped()
                for row in rows[:-1]:
                    assert abs(row.arrow.winfo_x()+row.arrow.winfo_width()/2-row.winfo_width()/2)<3
                    assert abs(row.checkpoint.winfo_x()+row.checkpoint.winfo_width()/2-row.winfo_width()/2)<3
                    assert row.checkpoint.cget('anchor')=='center' and row.checkpoint.cget('justify')=='center'
                arrow_image = widget_image(rows[0].arrow._label)
                arrow_image.save(output/'convoyage_arrow_probe_20261020.png')
                assert color_pixels(arrow_image, app._palette['positive_text'])>0
                title = errors_view.detail_title
                header = title.master
                assert abs(title.winfo_x()+title.winfo_width()/2-header.winfo_width()/2)<3
                canvas = errors_view.detail_scroll._parent_canvas
                assert rows[4].winfo_rooty()+rows[4].winfo_height()<=canvas.winfo_rooty()+canvas.winfo_height()
                capture('compact_seven_dark')
                app.geometry('800x820+0+0'); settle()
                assert len(errors_view.panes.panes())==1
                errors_view.back_button.invoke(); settle()
                assert str(errors_view.list_panel) in [str(p) for p in errors_view.panes.panes()]
                before_calls = summary.call_count, steps.call_count
                click(errors_view, errors_view.selected_index); settle()
                assert str(errors_view.detail_panel) in [str(p) for p in errors_view.panes.panes()]
                assert before_calls==(summary.call_count, steps.call_count)
                errors_view.back_button.invoke(); settle()
                click(errors_view, 2); settle()
                assert str(errors_view.detail_panel) in [str(p) for p in errors_view.panes.panes()]
                errors_view.reset_items(saved_errors, preserve=True, force=True)
                workspace.count.configure(text='180 anomalies · 3 fichiers')
                pump(lambda: errors_view.state=='ready')
                capture('narrow_dark')
                app._change_theme('Clair'); settle()
                capture('narrow_light')
                app.geometry('1340x880+0+0'); settle()
                click(errors_view, errors_view.errors.index(selected))
                errors_view.cards.canvas.yview_moveto(list_scroll)
                settle()
                # Copy uses the exact original path (native clipboard writes mocked).
                copied = []
                with patch.object(errors_view, 'clipboard_clear'), patch.object(errors_view, 'clipboard_append', side_effect=copied.append):
                    errors_view.copy('path')
                assert copied==[selected[0]]
                # Refresh preserves the selection and list position; no popup.
                workspace.refresh_button.invoke(); pump(lambda: ready('anomalies'))
                assert errors_view.errors[errors_view.selected_index]==selected
                # A verification cycle without a content change must clear the waiting state.
                app._sources_verified = False
                settle(.12)
                assert workspace.refresh_button.cget('state')=='disabled'
                app._sources_verified = True
                pump(lambda: ready('anomalies') and 'Contrôle terminé' in workspace.pages['anomalies']['status'].cget('text'))
                assert errors_view.errors[errors_view.selected_index]==selected
                assert abs(errors_view.cards.canvas.yview()[0]-list_scroll)<.01
                # A failed refresh keeps the last validated detail and can be retried.
                app.geometry('970x880+0+0'); settle()
                with patch.object(app.store, 'convoyage_errors', side_effect=OSError('temporary refresh failure')):
                    workspace.refresh_button.invoke()
                    pump(lambda: workspace.states['anomalies']['error'])
                    assert errors_view.errors[errors_view.selected_index]==selected
                    assert 'Actualisation impossible' in workspace.pages['anomalies']['status'].cget('text')
                    assert workspace.pages['anomalies']['status'].cget('wraplength')<=workspace.winfo_width()
                    capture('refresh_error_narrow_light')
                app.geometry('1340x880+0+0'); settle()
                workspace.refresh_button.invoke(); pump(lambda: ready('anomalies'))
                assert errors_view.errors[errors_view.selected_index]==selected
                # No anomaly clears the exact destination; refresh restores the same widgets.
                with patch.object(app.store, 'convoyage_errors', return_value=[]):
                    workspace.refresh_button.invoke(); pump(lambda: ready('anomalies'))
                    assert workspace.pages['anomalies']['title'].cget('text')=='Aucune anomalie détectée'
                    assert errors_view.selected_index is None
                    assert workspace.count.cget('text')=='0 anomalies · 0 fichiers'
                workspace.refresh_button.invoke(); pump(lambda: ready('anomalies'))
                assert workspace.views['anomalies'] is errors_view and errors_view.selected_index==0
                restored_index = errors_view.errors.index(selected)
                click(errors_view, restored_index)
                # A previous request cannot overwrite a newer ID or change the tab.
                workspace.buttons['parcours'].invoke()
                started, release = threading.Event(), threading.Event(); releases.append(release)
                original = app.store.convoyage_summary
                def delayed(ident, path=None):
                    if ident=='1001' and path is None:
                        started.set(); release.wait(5)
                    return original(ident, path)
                with patch.object(app.store, 'convoyage_summary', side_effect=delayed):
                    search('1001'); pump(started.is_set)
                    workspace.buttons['anomalies'].invoke()
                    release.set(); pump(lambda: ready('parcours'))
                    assert workspace.active_tab=='anomalies'
                    workspace.buttons['parcours'].invoke()
                    assert view.errors[0][1]=='1001'
                started, release = threading.Event(), threading.Event(); releases.append(release)
                with patch.object(app.store, 'convoyage_summary', side_effect=delayed):
                    search('1001')  # Same ID: already cached, no task.
                    assert not started.is_set()
                    search('1002'); pump(lambda: ready('parcours'))
                    search('1001'); pump(started.is_set)
                    search('1003'); pump(lambda: ready('parcours'))
                    release.set(); settle()
                    assert all(ident=='1003' for _, ident in view.errors)
                # Empty query/invalid ID do not dispatch a request or erase validated results.
                count = summary.call_count
                search('wrong')
                assert summary.call_count==count
                assert 'numérique' in workspace.pages['parcours']['status'].cget('text')
                search('999999'); pump(lambda: ready('parcours'))
                assert workspace.pages['parcours']['title'].cget('text')=='Aucune position retrouvée'
                assert view.selected_index is None and not hasattr(view, 'path_button')
                # Failure on a new query has an on-page retry, no messagebox.
                with patch.object(app.store, 'convoyage_summary', side_effect=RuntimeError('synthetic failure')):
                    search('1004'); pump(lambda: workspace.states['parcours']['error'])
                    assert workspace.pages['parcours']['retry'].winfo_ismapped()
                workspace.pages['parcours']['retry'].invoke(); pump(lambda: ready('parcours'))
                assert view.errors[0][1]=='1004'
                # Actual publication in the temporary files refreshes both tabs lazily.
                workspace.buttons['anomalies'].invoke(); settle()
                before = len(errors_view.errors)
                extra = ['0']*121; extra[0], extra[2], extra[6] = '13|12:00:00', '8080', '-8080'
                with open(paths[0], 'a', encoding='utf8') as handle:
                    handle.write(';'.join(extra)+'\n')
                pump(lambda: ready('anomalies') and len(errors_view.errors)==before+1)
                assert errors_view.errors[errors_view.selected_index]==selected
                assert not [w for w in app.winfo_children() if isinstance(w, gui.tk.Toplevel)]
            # Legacy entry points still open their own windows; no embedded close handler.
            app.convoyage_action('1000')
            pump(lambda: getattr(app, 'convoyage_window', None) is not None)
            popup = app.convoyage_window
            runners.append(popup.convoyage_errors_view.tasks)
            assert not popup.convoyage_errors_view.embedded
            popup.convoyage_errors_view.close_button.invoke()
            assert app.winfo_exists()
            # Close the application while the workspace's request is blocked.
            workspace.buttons['parcours'].invoke()
            started, release = threading.Event(), threading.Event(); releases.append(release)
            original = app.store.convoyage_summary
            def blocked(ident, path=None):
                if path is None:
                    started.set(); release.wait(5)
                return original(ident, path)
            with patch.object(app.store, 'convoyage_summary', side_effect=blocked):
                search('1005'); pump(started.is_set)
                app._close_indexed_app()
                release.set()
            assert workspace.tasks.closed
            assert all(v.tasks.closed for v in workspace.views.values())
            print('OK: inline Parcours/Anomalies, real buttons/Enter/cards, no popup, exact paths across '
                  '3 sources, shared R1/R2, rendered colors, light/dark/wide/narrow captures, '
                  'verification gate, retained results/scroll, stale replies, errors/retry, '
                  'publication refresh, legacy popup and shutdown during loading.')
        finally:
            for release in releases:
                release.set()
            try:
                if app.winfo_exists(): app._close_indexed_app()
            except gui.tk.TclError:
                pass
            app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)
            if hasattr(app, 'convoyage_workspace'):
                runners.extend(v.tasks for v in app.convoyage_workspace.views.values())
            for runner in runners:
                runner.executor.shutdown(wait=True, cancel_futures=True)


if __name__=='__main__':
    run_gui_check(main)
