"""Actual report/neighbor interactions, temporary data, isolated Windows desktop."""
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import patch
from ctypes import wintypes
import ctypes
import os
import sys
import tempfile
import threading
import time

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0,str(PROJECT))
import search_gui as gui
import trace_sources as sources
from tests.test_get_diagnostics import HEADER,ticket,get,FAILURE,OK
from tools.ui_test_support import run_gui_check
from tools.native_test_capture import widget_image,color_pixels
from window_placement import native_window_bounds
from get_errors_view import error_key


def main():
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        traces = base/'traces'
        files = []
        for robot,ref in (('R1','FIRST'),('R2','SECOND')):
            path = traces/robot/'TracOmega.txt'; path.parent.mkdir(parents=True)
            path.write_text(HEADER+ticket(ref=ref)+get(ref)+FAILURE
                            +get(ref+'SHELF').replace('$0101','F32')+FAILURE,encoding='utf8')
            files.append(path)
        unknown=traces/'TracOMEGA.txt'
        unknown.write_text(HEADER+'G 13|13:37:23.9 Err: réponse gripper GET: Err NOK (-7)\n',encoding='utf8')
        files.append(unknown)
        stack.enter_context(patch.object(sources,'application_directory',return_value=str(base)))
        stack.enter_context(patch.object(sources,'desktop_directory',return_value=str(base/'absent')))
        stack.enter_context(patch.object(sources,'FIXED_DIRECTORY',str(base/'missing')))
        stack.enter_context(patch.dict(os.environ,LOCALAPPDATA=str(base/'AppData')))
        app = gui.SearchApp(); app.geometry('1090x820+0+0')
        errors,gates,views = [],[],[]
        app.report_callback_exception = lambda *error:errors.append(error)
        def pump(condition=lambda:True,settle=0):
            until = time.perf_counter()+15; stable = None
            while time.perf_counter()<until:
                app.update()
                if errors: raise errors[0][1]
                if condition():
                    stable = stable or time.perf_counter()
                    if time.perf_counter()-stable>=settle: return
                else: stable = None
                time.sleep(.003)
            raise AssertionError('UI timeout '+app.status.get())
        def click(button):
            button._canvas.event_generate('<ButtonPress-1>',x=5,y=5)
            button._canvas.event_generate('<ButtonRelease-1>',x=5,y=5)
        def choose(widget,value):
            menu = widget._dropdown_menu
            index = widget.cget('values').index(value)
            with patch.object(menu,'post',side_effect=lambda *a:menu.invoke(index)):
                widget._canvas.event_generate('<ButtonPress-1>',x=5,y=5)
            pump()
        stamp = 1000
        def tree_click(tree,iid,double=False):
            nonlocal stamp
            tree.see(iid); pump()
            x,y,w,h = tree.bbox(iid)
            for i in range(2 if double else 1):
                tree.event_generate('<ButtonPress-1>',x=x+15,y=y+h//2,time=stamp+i*100)
                tree.event_generate('<ButtonRelease-1>',x=x+15,y=y+h//2,time=stamp+i*100+20)
            stamp+=1000
        def trace_windows():
            return [w for w in app.winfo_children() if getattr(w,'trace_controller',None)]
        def verify_trace(event):
            pump(lambda:trace_windows() and trace_windows()[0].trace_controller.loaded)
            window=trace_windows()[0]; controller=window.trace_controller
            assert controller.text.get('highlight.first','highlight.last')==event['text']
            assert controller.text.index('highlight.first')==str(event['number']+1)+'.0'
            assert controller.text.get('1.0','end-1c').splitlines()==app.store.file_lines(event['path'])
            window.destroy();pump()
        user=ctypes.WinDLL('user32',use_last_error=True)
        user.GetWindow.argtypes=[wintypes.HWND,wintypes.UINT];user.GetWindow.restype=wintypes.HWND
        def above(first,second):
            handle=native_window_bounds(second)[0]
            while handle:
                handle=user.GetWindow(handle,3)  # GW_HWNDPREV
                if handle==native_window_bounds(first)[0]: return True
            return False
        try:
            pump(lambda:app._sources_verified)
            app._sources_verified=False;pump(settle=.05)
            assert app.btn_get_errors.cget('state')=='disabled'
            click(app.btn_get_errors);assert not getattr(app,'get_errors_popup',None)
            app._sources_verified=True;pump(settle=.05)
            original=app.store.get_errors
            with patch.object(app.store,'get_errors',wraps=original) as queries:
                click(app.btn_get_errors)
                popup=app.get_errors_popup;view=popup.get_errors_view;views.append(view)
                pump(lambda:view.data is not None and view._insert_timer is None)
                assert len(view.rows)==5 and popup.transient()==''
                assert not user.GetWindow(native_window_bounds(popup)[0],4)  # GW_OWNER
                assert app.selection is None,'Selecting a report must not change the search'
                app._sources_verified=False;pump(settle=.05)
                assert view.open_button.cget('state')=='disabled'
                with patch.object(app,'_open_indexed_file') as opening:
                    view.tree.focus_force();view.tree.event_generate('<KeyPress-Return>')
                    opening.assert_not_called()
                app._sources_verified=True
                pump(lambda:not view._loading and view._marker==app._index_content_revision)
                calls=queries.call_count
                choose(view.robot_filter,'Non identifié');choose(view.location_filter,'Inconnu')
                assert len(view.rows)==1 and 'Commande GET non rattachée' in view.command.cget('text')
                orphan=next(iter(view.rows.values()))
                tree_click(view.tree,'0',double=True);verify_trace(orphan['error'])
                choose(view.robot_filter,'R1')
                assert not view.rows and 'Aucune erreur GET' in view.detail_title.cget('text')
                choose(view.robot_filter,'R1');choose(view.location_filter,'$0101')
                assert len(view.rows)==1 and next(iter(view.rows.values()))['robot']=='R1'
                choose(view.robot_filter,'R2')
                assert len(view.rows)==1 and next(iter(view.rows.values()))['robot']=='R2'
                row=next(iter(view.rows.values())); selected=error_key(row)
                tree_click(view.tree,'0');assert row['error']['text'] in view.responses.cget('text')
                assert app.selection is None
                # A native heading click sorts without querying again.
                view.tree.event_generate('<ButtonPress-1>',x=20,y=8)
                view.tree.event_generate('<ButtonRelease-1>',x=20,y=8);pump()
                assert queries.call_count==calls
                tree_click(view.tree,'0',double=True);verify_trace(row['error'])
                widget=view.command.selection_text
                for i in range(2):
                    widget.event_generate('<ButtonPress-1>',x=5,y=5,time=stamp+i*100)
                    widget.event_generate('<ButtonRelease-1>',x=5,y=5,time=stamp+i*100+20)
                stamp+=1000;verify_trace(row['get'])
                view.tree.focus_force();view.tree.event_generate('<KeyPress-Return>');verify_trace(row['error'])
                for theme in ('Clair','Sombre'):
                    app._change_theme(theme);popup.geometry('1090x740');pump(settle=.1)
                    image=widget_image(popup)
                    assert color_pixels(image,app._palette['warning_text'])>10
                    image.save(PROJECT/'.benchmark'/f'get_errors_{theme}_20261020.png')
                popup.geometry('780x640');pump(settle=.1)
                widget_image(popup).save(PROJECT/'.benchmark/get_errors_narrow_20261020.png')
                assert queries.call_count==calls
                # Population is batched and interrupted by a new filter.
                original_data=view.data
                view.data=dict(original_data,errors=original_data['errors']*1000)
                choose(view.robot_filter,'Tous');choose(view.location_filter,'Tous')
                view._display()
                assert len(view.rows)==200 and view._insert_timer is not None
                view.data=original_data
                choose(view.robot_filter,'R2');choose(view.location_filter,'$0101')
                pump(settle=.1)
                assert len(view.rows)==1 and view._insert_timer is None
                assert queries.call_count==calls
                # Publication refreshes the report, retaining both filters and selection.
                files[1].write_text(files[1].read_text(encoding='utf8')+ticket('1053',ref='NEW')+get('NEW')+FAILURE,encoding='utf8')
                report=app.store.sync([str(p) for p in files])
                app._index_messages.put(('complete',report,app.store.metadata()))
                pump(lambda:view.data['revision']==app.store.analysis_revision() and view._insert_timer is None and not view._loading)
                assert view.robot_filter.get()=='R2' and view.location_filter.get()=='$0101'
                assert view._selected_identity==selected
            # A stale response and a recoverable error never replace the current report.
            gate=threading.Event();gates.append(gate);entered=threading.Event();calls=[0]
            def delayed(cancel=None):
                calls[0]+=1
                if calls[0]==1:
                    entered.set();gate.wait(5)
                    return dict(revision=-1,errors=[])
                return original(cancel)
            with patch.object(app.store,'get_errors',side_effect=delayed):
                click(view.refresh_button);pump(entered.is_set)
                with patch.object(app,'_open_indexed_file') as opening:
                    view.tree.focus_force();view.tree.event_generate('<KeyPress-Return>')
                    opening.assert_not_called()
                app._index_content_revision+=1;gate.set()
                pump(lambda:not view._loading and view._marker==app._index_content_revision and calls[0]>=2)
                assert view.data['errors'] and view.data['revision']!=-1
            with patch.object(app.store,'get_errors',side_effect=ValueError('temporary diagnostic error')) as query:
                click(view.refresh_button);pump(lambda:'temporary diagnostic error' in view.header.subtitle.cget('text'))
                count=query.call_count;pump(settle=.1);assert query.call_count==count
                with patch.object(app,'_open_indexed_file') as opening:
                    view.open_button.invoke()
                    view.tree.focus_force();view.tree.event_generate('<KeyPress-Return>')
                    opening.assert_not_called()
            click(view.refresh_button);pump(lambda:not view._loading and view._marker==app._index_content_revision)
            # Neighbor windows are modeless, and late results never raise either window.
            popup.destroy();app.modbus.set(True);app.reference.set(False);app.on_modbus_check()
            app.keyword.insert(0,'1052');click(app.btn_search)
            pump(lambda:app.selection is not None and '_summary_comparison' in app.selection.details)
            click(app.btn_adjacent);pump(lambda:getattr(app,'adjacent_view',None) is not None and app.adjacent_view.data is not None)
            adjacent=app.adj_popup
            assert adjacent.transient()=='' and not user.GetWindow(native_window_bounds(adjacent)[0],4)
            app.lift();app.focus_force();pump()
            assert above(app,adjacent),'Main window must be able to rise above neighbors'
            adjacent.lift();pump();assert above(adjacent,app)
            column=next(c for c in app.adjacent_view.data['columns'] if c['robot']=='R2')
            anchor=column['anchor'];iid=next(i for i,r in app.adjacent_view.rows['R2'].items() if r['number']==anchor['number'] and r['path']==anchor['path'])
            tree_click(app.adjacent_view.trees['R2'],iid)
            pump(lambda:app.selection is not None and app.selection.path==anchor['path'] and app.selection.number==anchor['number'])
            assert above(adjacent,app),'Selecting a neighbor must not force the main window forward'
            adjacent.destroy()
            # Closing a loading report cancels only that window's worker.
            gate=threading.Event();gates.append(gate);entered=threading.Event()
            def waiting(cancel=None):
                entered.set()
                while not gate.wait(.01): app.store._check_cancel(cancel)
                return original(cancel)
            with patch.object(app.store,'get_errors',side_effect=waiting):
                click(app.btn_get_errors);view=app.get_errors_popup.get_errors_view;views.append(view)
                pump(entered.is_set);app.get_errors_popup.destroy();assert view.tasks.closed
                gate.set();pump(settle=.1)
            print('OK GET report buttons/filters/sort, complete sources, exact error/GET/Enter openings, '
                  'R1/R2 same location, themes/narrow captures, revision/stale/error/retry/close, '
                  'modeless native ownership/z-order and neighbor clicks without forced raising')
        finally:
            for gate in gates:gate.set()
            app._close_indexed_app();app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True,cancel_futures=True)
            for view in views:view.tasks.executor.shutdown(wait=True,cancel_futures=True)


if __name__=='__main__':run_gui_check(main)
