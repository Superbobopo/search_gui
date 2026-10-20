"""Reusable summary lifecycle and actual clicks, exclusively on a private desktop."""
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import patch
import os
import sys
import tempfile
import threading
import time

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
import search_gui as gui
import trace_sources as sources
from trace_parser import read_trace
from tools.ui_test_support import run_gui_check
from tools.summary_test_support import click_summary, rendered_summary_text
from tools.native_test_capture import widget_image


def main():
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        directory = base/'traces'/'R2'
        directory.mkdir(parents=True)
        original = read_trace(PROJECT/'tests/fixtures/current_r2.txt')
        second = original.splitlines()[3:]
        second = '\n'.join(second).replace('1194','1195').replace('A0061574AC','SECOND_REF')\
            .replace('343272','SECOND_SN').replace('06:','07:')
        path = directory/'TracOMEGA.txt'
        path.write_text(original+'\n'+second, encoding='utf8')
        stack.enter_context(patch.object(sources, 'application_directory', return_value=str(base)))
        stack.enter_context(patch.object(sources, 'desktop_directory', return_value=str(base/'absent')))
        stack.enter_context(patch.object(sources, 'FIXED_DIRECTORY', str(base/'missing')))
        stack.enter_context(patch.dict(os.environ, LOCALAPPDATA=str(base/'AppData')))
        app = gui.SearchApp(); app.geometry('1090x820+0+0')
        errors = []
        app.report_callback_exception = lambda *error: errors.append(error)

        def pump(condition=lambda: True):
            until = time.perf_counter()+12
            while time.perf_counter()<until:
                app.update()
                if errors:
                    raise errors[0][1]
                if condition():
                    return
                time.sleep(.002)
            raise AssertionError(app.status.get())

        def final():
            return app.selection is not None and '_summary_comparison' in app.selection.details

        def click(button):
            button._canvas.event_generate('<ButtonPress-1>', x=5, y=5)
            button._canvas.event_generate('<ButtonRelease-1>', x=5, y=5)

        def search(value):
            app.keyword.delete(0,'end'); app.keyword.insert(0,value); click(app.btn_search)

        gates = []
        try:
            pump(lambda: app._sources_verified)
            app.modbus.set(True); app.reference.set(False); app.on_modbus_check()
            search('1194'); pump(final)
            first = app.selection
            stable = {key:app.summary_targets[key] for key in ('get','pute','idpk','location')}
            pool = app.summary_frame._summary_widgets
            binding = stable['get'].master.bind('<Configure>')
            for _ in range(3):
                click(app.btn_down); pump(final)
                assert app.selection.modbus=='1195'
                assert all(app.summary_targets[k] is w for k,w in stable.items())
                click(app.btn_up); pump(final)
            assert stable['get'].master.bind('<Configure>')==binding

            gate, entered = threading.Event(), threading.Event(); gates.append(gate)
            original_search = app.store.search

            def delayed(*args, **kw):
                if args[0]=='ModBusId=1195':
                    entered.set()
                    while not gate.wait(.01):
                        app.store._check_cancel(kw.get('cancel'))
                return original_search(*args, **kw)

            with patch.object(app.store, 'search', side_effect=delayed), \
                 patch.object(app, '_open_indexed_file') as opening:
                old_text = stable['get'].cget('text')
                click(app.btn_down); pump(entered.is_set)
                assert app.selection is None and pool.loading
                assert stable['get'].cget('text')==old_text and stable['get'].winfo_ismapped()
                assert 'Recherche' in app.status.get() or 'Chargement' in app.status.get()
                widget_image(app).save(PROJECT/'.benchmark/summary_navigation_loading_20261020.png')
                click_summary(app, 'get', stamp=150000)
                opening.assert_not_called()
                click(app.btn_up); pump(final)
                gate.set(); pump()
                assert app.selection.modbus=='1194' and not pool.loading

            with patch.object(app.store, 'search', side_effect=ValueError('temporary search failure')):
                click(app.btn_down)
                pump(lambda:'temporary search failure' in app.status.get())
                assert pool.loading and app.selection is None and stable['get'].winfo_exists()
            click(app.btn_search); pump(final)
            assert app.selection.modbus=='1195'
            event = app.selection.details['get']
            click_summary(app, 'get', stamp=250000)
            def traces():
                return [w for w in app.winfo_children() if getattr(w,'trace_controller',None)]
            pump(lambda:traces() and traces()[0].trace_controller.loaded)
            controller = traces()[0].trace_controller
            assert controller.text.get('highlight.first','highlight.last')==event['text']
            assert controller.text.index('highlight.first')==str(event['number']+1)+'.0'
            traces()[0].destroy()

            with patch.object(app.store, 'current_measurement_for_occurrence', wraps=app.store.current_measurement_for_occurrence) as local, \
                 patch.object(app.store, 'comparison_for_occurrence', wraps=app.store.comparison_for_occurrence) as comparison:
                selected = app.selection
                click_summary(app,'refprod',stamp=350000,double=False)
                pump(lambda:final() and app._active_mode=='ref')
                assert (app.selection.path,app.selection.number)==(selected.path,selected.number)
                local.assert_not_called(); comparison.assert_not_called()

            selected = app.selection
            with patch.object(app.store, 'search', wraps=app.store.search) as query:
                click(app.btn_tracking_robot); pump()
                click(app.btn_search_menu); pump()
                assert app.selection is selected and app.summary_targets['get'] is stable['get']
                query.assert_not_called()
            gate, entered = threading.Event(), threading.Event(); gates.append(gate)
            original_compare = app.store.comparison_for_occurrence

            def delayed_compare(path, number, cancel=None):
                entered.set()
                while not gate.wait(.01):
                    app.store._check_cancel(cancel)
                return original_compare(path, number, cancel)

            with patch.object(app.store, 'comparison_for_occurrence', side_effect=delayed_compare):
                click(app.btn_search)
                pump(lambda:app.selection is not None and '_summary_current' in app.selection.details and entered.is_set())
                local_widget = app.summary_targets['idpk']
                click(app.btn_tracking_robot); pump()
                click(app.btn_search_menu); pump()
                gate.set(); pump(final)
                assert app.summary_targets['idpk'] is local_widget
            # Changing the source publication forbids reuse of the previous analysis.
            path.write_text(original+'\n'+second.replace('l=48 H=21 L=47','l=55 H=21 L=47'), encoding='utf8')
            app.store.sync([str(path)])
            with patch.object(app.store, 'comparison_for_occurrence', wraps=app.store.comparison_for_occurrence) as comparison:
                click(app.btn_search); pump(final)
                assert comparison.call_count>=1
                assert app.selection.details['_summary_comparison']['revision']==app.store.analysis_revision()

            # Hide every stale field when a real result has no locations, then restore it.
            data = app.selection.details; locations = data['locations']
            data['locations'] = []; app._render_summary(data); pump()
            assert 'Aucun emplacement connu' in rendered_summary_text(app)
            assert not any('Lot ' in w.cget('text') for (_,kind,_),w in pool.widgets.items() if kind=='label')
            data['locations'] = locations; app._render_summary(data); pump()
            # Expanded cards above the twenty-card reserve are reclaimed asynchronously.
            event = locations[0]
            data['locations'] = [dict(event,lot_hint=f'LOT-{i}') for i in range(45)]
            app._render_summary(data); pump()
            group = next(w for w in pool.widgets.values() if hasattr(w,'_summary_more'))
            click(group._summary_more); pump(); click(group._summary_more); pump()
            assert group._summary_shown==45
            data['locations'] = locations; app._render_summary(data)
            pump(lambda:not any(getattr(w,'_summary_disposable',False) and not w._summary_active for w in pool.widgets.values()))
            app.modbus.set(True); app.reference.set(False); app.on_modbus_check()
            search('999999'); pump(lambda:app.selection is None and not app._groups and '0 fichier' in app.status.get())
            assert not app.summary_targets and 'Aucune occurrence sélectionnée' in rendered_summary_text(app)
            assert 'SECOND_REF' not in rendered_summary_text(app)
            print('OK reusable fields, bounded pool, stable bindings, loading/error/retry, blocked stale openings, '
                  'rapid +/- cancellation, reference analysis reuse, source revision, empty result, pages and exact GET highlight')
        finally:
            for gate in gates:
                gate.set()
            app._close_indexed_app(); app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True,cancel_futures=True)


if __name__=='__main__':
    run_gui_check(main)
