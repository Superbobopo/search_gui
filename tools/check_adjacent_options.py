"""Exact comparisons, actual rendered colors and persistent adjacent tables."""
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
from tests.test_occurrence_comparison import measured, ticket, trace
from tools.ui_test_support import run_gui_check
from tools.native_test_capture import widget_image, color_pixels


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def main():
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        files = []
        for robot in ('R1','R2'):
            folder = base/'traces'/robot
            folder.mkdir(parents=True)
            rows = []
            for ident in range(1040,1081):
                sec = 36000 + (ident-1040)*10 + (2 if robot=='R2' else 0)
                if robot=='R1' and ident==1054:
                    row = measured(sec,str(ident),'A0118427AA',sn='292226',lot='260210.004')
                elif robot=='R2' and ident==1052:
                    row = ticket(sec,str(ident),'A0110543AC','261012.095','343132')
                else:
                    row = measured(sec,str(ident),robot+'_REF'+str(ident),sn='SN'+str(ident),lot='LOT'+str(ident))
                rows.append(row)
            path=folder/'TracOmega019.old'
            path.write_text(trace(*rows),encoding='utf8')
            files.append(path)
        old=base/'traces'/'R2'/'TracOmega018.old'
        old.write_text(trace(measured(1000,'1','A0118427AA',sn='292226',lot='260210.004',dims=(104,48,189)),
                             measured(1100,'2','A0110543AC',sn='OTHER',lot='261012.095',dims=(1,1,1))),encoding='utf8')
        files.append(old)
        prior=base/'traces'/'R1'/'TracOmega018.old'
        prior.write_text(trace(measured(2000,'858','A0118427AA',sn='292226',lot='260210.004')),encoding='utf8')
        files.append(prior)
        stack.enter_context(patch.object(sources,'application_directory',return_value=str(base)))
        stack.enter_context(patch.object(sources,'desktop_directory',return_value=str(base/'absent')))
        stack.enter_context(patch.object(sources,'FIXED_DIRECTORY',str(base/'missing')))
        stack.enter_context(patch.dict(os.environ,LOCALAPPDATA=str(base/'AppData')))
        app=gui.SearchApp()
        errors=[]
        app.report_callback_exception=lambda *error: errors.append(error)
        release=threading.Event()
        def pump(condition=lambda:True):
            deadline=time.monotonic()+12
            while time.monotonic()<deadline:
                app.update()
                if errors: raise errors[0][1]
                if condition(): return
                time.sleep(.004)
            raise AssertionError('Adjacent options timeout: '+app.status.get())
        def ready(robot=None,modbus=None):
            view=getattr(app,'adjacent_view',None)
            return (view and view.data is not None and (robot is None or view.data['selected_robot']==robot)
                    and (modbus is None or view.data['selected']['modbus']==modbus))
        def row(robot,modbus):
            view=app.adjacent_view
            assert callable(view.after), 'Keep the native Tk scheduler available'
            return next((key,item) for key,item in view.rows[robot].items() if item['modbus']==modbus)
        def click(robot,modbus):
            view=app.adjacent_view
            key,target=row(robot,modbus)
            tree=view.trees[robot]
            tree.see(key); pump()
            x,y,width,height=tree.bbox(key)
            before=tuple(str(p['tree']) for p in view.panels)
            tree.event_generate('<ButtonPress-1>',x=x+10,y=y+height//2)
            tree.event_generate('<ButtonRelease-1>',x=x+10,y=y+height//2)
            # Old rows stay present while details are in flight.
            assert tuple(str(p['tree']) for p in view.panels)==before and all(p['tree'].get_children() for p in view.panels)
            pump(lambda:app.selection is not None and app.selection.path==target['path']
                 and app.selection.number==target['number'] and ready(robot,modbus))
            assert tuple(str(p['tree']) for p in view.panels)==before
            return target
        def rendered(robot,modbus,different,expected_red=None):
            view=app.adjacent_view
            key,item=row(robot,modbus)
            assert item['different'] is different
            visible_red=different if expected_red is None else expected_red
            tree=view.trees[robot]
            tree.see(key); pump()
            x,y,width,height=tree.bbox(key,'ref')
            image=widget_image(tree).crop((x,y,x+width,y+height))
            reds=color_pixels(image,app._palette['error_text'])
            expected=app._palette['error_text'] if visible_red else app._palette['text']
            foreground=color_pixels(image,expected)
            if foreground<3 or (not visible_red and reds):
                folder=Path(__file__).resolve().parents[1]/'.benchmark'
                folder.mkdir(exist_ok=True)
                image.save(folder/'adjacent_capture_failure.png')
                raise AssertionError(('Actual pixels',robot,modbus,different,reds,foreground))
            assert not tree.selection(), 'Native selection must not mask row colors'
        def compare_windows():
            return [w for w in app.winfo_children() if isinstance(w,gui.tk.Toplevel) and w.title().startswith('Comparaison IDPK')]
        def close_comparisons():
            for window in compare_windows(): window.destroy()
        try:
            pump(lambda:app._sources_verified)
            sync=app.store.sync
            stack.enter_context(patch.object(app.store,'sync',return_value=dict(changed=0,removed=0,failures=[])))
            app.mode_radios['ref'].invoke()
            app.keyword.insert(0,'R1_REF1050')
            app.btn_search.invoke()
            pump(lambda:app.selection is not None)
            app.btn_adjacent.invoke()
            pump(lambda:ready('R1','1050'))
            view=app.adjacent_view
            rendered('R1','1054',True)
            for target in ('1051','1055','1054'):
                click('R1',target)
                rendered('R1','1054',True)
            pump(lambda:'_summary_comparison' in app.selection.details)
            result=app.selection.details['_summary_comparison']
            assert result['different'] and result['pair'][1]['path'].path==str(old)
            assert app.summary_targets['idpk'].cget('text_color')==('#b42318','#ff9b91')
            app.btn_compare_idpk.invoke()
            pump(lambda:bool(compare_windows()))
            labels=' '.join(str(w.cget('text')) for w in descendants(compare_windows()[0]) if isinstance(w,gui.tk.Label))
            assert 'Mesure antérieure différente' in labels and '292226' in labels and '260210.004' in labels and 'Δ L = +6' in labels
            close_comparisons()
            # Current click is a true no-op, including all SQL-facing APIs.
            with patch.object(app.store,'paired_adjacent',wraps=app.store.paired_adjacent) as paired, \
                 patch.object(app.store,'comparison_for_occurrence',wraps=app.store.comparison_for_occurrence) as comparison:
                selected=app.selection
                click('R1','1054')
                assert app.selection is selected and paired.call_count==comparison.call_count==0
            # R2/1052 must be neutral both as neighbor of 1049 and as current.
            click('R2','1049')
            rendered('R2','1052',False)
            click('R2','1052')
            rendered('R2','1052',False)
            pump(lambda:'_summary_comparison' in app.selection.details)
            assert app.selection.details['_summary_comparison']['status']=='missing_measurement'
            assert 'rattachable' in view.detail.cget('text')
            app.btn_compare_idpk.invoke(); pump(lambda:bool(compare_windows()))
            labels=' '.join(str(w.cget('text')) for w in descendants(compare_windows()[0]) if isinstance(w,gui.tk.Label))
            assert 'rattachable' in labels and '343132' in labels
            close_comparisons()
            for counts in ((5,10),(5,30),(5,5)):
                view.presets[counts].invoke()
                pump(lambda:view.data and (view.data['before_count'],view.data['after_count'])==counts)
                assert app._adjacent_counts==counts
            for value in ('-1','1001','3.5','abc'):
                view.before_var.set(value); view.apply_button.invoke(); pump()
                assert view.validation.cget('text') and app._adjacent_counts==(5,5)
            for counts in ((0,0),(12,47),(1000,1000)):
                view.before_var.set(str(counts[0])); view.after_var.set(str(counts[1])); view.apply_button.invoke()
                pump(lambda:view.data and (view.data['before_count'],view.data['after_count'])==counts)
                assert all(len(p['tree'].get_children())<=42 for p in view.panels)
            # One scrollbar, synchronized vertical tables, header/control geometry fixed.
            header_y=view.heading.winfo_rooty()
            tree=view.trees['R1']; first=tree.yview()[0]
            tree.event_generate('<MouseWheel>',delta=-120); pump()
            assert tree.yview()[0]>first and abs(tree.yview()[0]-view.trees['R2'].yview()[0])<.00001
            assert view.heading.winfo_rooty()==header_y
            view.presets[(5,10)].invoke(); pump(lambda:view.data and view.data['after_count']==10)
            view.close_button.invoke(); app.btn_adjacent.invoke(); pump(lambda:ready('R2','1052'))
            view=app.adjacent_view
            assert (view.before_var.get(),view.after_var.get())==('5','10')
            view.presets[(5,5)].invoke(); pump(lambda:view.data and view.data['after_count']==5)
            click('R1','1054')
            for theme in ('Sombre','Clair'):
                app._change_theme(theme); pump()
                rendered('R1','1054',True)
            app._sources_verified=False
            app.btn_adjacent.invoke()
            pump(lambda:'Vérification en cours' in view.note.cget('text'))
            rendered('R1','1054',True,expected_red=False)
            assert 'après vérification' in view.detail.cget('text')
            app._index_messages.put(('complete',dict(changed=0,removed=0,failures=[]),app.store.metadata()))
            pump(lambda:app._sources_verified and 'Vérification en cours' not in view.note.cget('text')
                 and not str(view.note.cget('text')).startswith('Actualisation'))
            rendered('R1','1054',True)
            # Actual copy of a neutral cell, without touching the shared clipboard.
            key,item=row('R1','1054'); tree=view.trees['R1']; tree.see(key); pump()
            x,y,width,height=tree.bbox(key,'sn'); copied=[]
            with patch.object(view.menu,'tk_popup'),patch.object(view,'clipboard_clear'),patch.object(view,'clipboard_append',side_effect=copied.append):
                tree.event_generate('<Button-3>',x=x+width//2,y=y+height//2)
                view.menu.invoke(0)
                assert copied[-1]=='292226'
                rendered('R1','1054',True)
            # Archived rows and new revision invalidate old comparisons without rebuilding tables.
            for path in files: path.unlink()
            assert not sync([])['failures']
            app.btn_adjacent.invoke(); pump(lambda:view.data and view.data['selected']['path'].historical)
            rendered('R1','1054',True)
            # A cancelled range request cannot overwrite a later setting.
            started,finished=threading.Event(),threading.Event()
            original=app.store.paired_adjacent
            def slow(*args,**kwargs):
                started.set(); release.wait(5)
                try:return original(*args,**kwargs)
                finally:finished.set()
            with patch.object(app.store,'paired_adjacent',slow):
                view.presets[(5,30)].invoke(); pump(started.is_set)
                view.presets[(5,10)].invoke()
                release.set(); pump(finished.is_set)
                pump(lambda:view.data and view.data['after_count']==10)
                assert view.data['after_count']==10
            print('OK options/IDPK : reglages 0..1000, widgets conserves, scroll commun, '
                  'pixels rouges R1/1054 a toutes positions/themes et neutres R2/1052, '
                  'trois affichages coherents, copie, archive et requetes depassees')
        finally:
            release.set()
            app._close_indexed_app(); app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True,cancel_futures=True)


if __name__=='__main__':
    run_gui_check(main)
