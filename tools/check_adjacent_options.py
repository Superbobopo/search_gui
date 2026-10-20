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
                if (robot, ident) in (('R1', 1054), ('R2', 1052)):
                    lines = row.splitlines(True)
                    lines.insert(2, "r [Tgripper = 5s rep ='Err NOK (-32092)']\nG 14|10:02:00.0 Err: reponse gripper GET: Err NOK (-32092)\n")
                    row = ''.join(lines)
                if (robot, ident)==('R1', 1050):
                    lines = row.splitlines(True)
                    lines.insert(2, "r [Tgripper = 5s rep ='OK']\n")
                    row = ''.join(lines)
                if (robot, ident)==('R2', 1051):
                    lines = row.splitlines(True)
                    context = "b 'ChargeBoitesDansGripperMB' 1 boite(s) (IDb:142509(R2_REF1051)) en $0101\n"
                    retry = lines[1].replace('10:01:53.0', '10:01:56.0')
                    row = (lines[0]+context+lines[1]+"r [Tgripper = 1s rep ='Err NOK (-7)']\n"
                           +"G GRIPPER REF\nr [Tgripper = 1s rep ='OK']\n"+context+retry
                           +"r [Tgripper = 1s rep ='OK']\n"
                           +lines[2].replace('10:01:53.500000', '10:01:57.500000')
                           +lines[-1].replace('10:01:54.0', '10:01:58.0'))
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
        def manual(counts):
            view.before_var.set(str(counts[0])); view.after_var.set(str(counts[1]))
            view.apply_button.invoke()
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
            failed = item.get('get_diagnostic', {}).get('failed') and app._sources_verified and app._indexed_complete
            if failed and visible_red:
                assert tree.item(key, 'values')[-1]=='IDPK + échec GET'
            expected=app._palette['error_text'] if visible_red else app._palette['warning_text'] if failed else app._palette['text']
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
            pump(lambda: app.selection and app.selection.modbus=='1050')
            assert app.selection.details['get_diagnostic']['status']=='success'
            assert 'get_status' not in app.summary_targets
            assert 'get_response_0' not in app.summary_targets
            view=app.adjacent_view
            from adjacent_view import listing_text
            view.copy_button.invoke()
            app.update()
            copied = view.clipboard_get()
            assert copied==listing_text(view.data, app._selected_directories)
            assert copied.index('R1')<copied.index('R2')
            assert 'SN :' in copied and 'Lot :' in copied and '🟢 ➜ ' in copied
            assert 'TracOmega' not in copied and 'Repère' not in copied
            assert sum(' | ID :' in line for line in copied.splitlines())==sum(
                len(c.get('previous', []))+1+len(c.get('following', []))
                for c in view.data['columns'] if c.get('anchor'))
            published = view.data
            view.loading()
            assert view.copy_button.cget('state')=='disabled'
            view.copy_lists()
            assert view.clipboard_get()==copied
            view.show(published)
            assert view.copy_button.cget('state')=='normal'
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
            assert app.selection.details['get_diagnostic']['failed']
            assert app.selection.details['pute'] is not None
            assert '-32092' in view.get_detail.cget('text')
            pump(lambda: 'get_status' in app.summary_targets)
            assert not any(key.startswith('get_response_') for key in app.summary_targets)
            for theme in ('Clair', 'Sombre'):
                app._change_theme(theme); pump()
                rendered('R2','1052',False)
                assert color_pixels(widget_image(app.summary_targets['get_status'].selection_text), app._palette['warning_text'])>3
                assert color_pixels(widget_image(view.get_detail.selection_text), app._palette['warning_text'])>3
                output = Path(__file__).resolve().parents[1]/'.benchmark'
                output.mkdir(exist_ok=True)
                widget_image(view).save(output/f'get_neighbors_{theme}_20261020.png')
                widget_image(app).save(output/f'get_summary_{theme}_20261020.png')
            def traces():
                return [w for w in app.winfo_children() if getattr(w, 'trace_controller', None)]
            pump(lambda:'_summary_comparison' in app.selection.details)
            assert app.selection.details['_summary_comparison']['status']=='missing_measurement'
            assert 'rattachable' in view.detail.cget('text')
            app.btn_compare_idpk.invoke(); pump(lambda:bool(compare_windows()))
            labels=' '.join(str(w.cget('text')) for w in descendants(compare_windows()[0]) if isinstance(w,gui.tk.Label))
            assert 'rattachable' in labels and '343132' in labels
            close_comparisons()
            gate = threading.Event()
            original_compare = app.store.comparison_for_occurrence
            def delayed_compare(path, number, cancel=None):
                while not gate.wait(.01):
                    app.store._check_cancel(cancel)
                return original_compare(path, number, cancel)
            with patch.object(app.store, 'comparison_for_occurrence', side_effect=delayed_compare):
                try:
                    start = time.perf_counter()
                    click('R2','1051')
                    pump(lambda: '_summary_current' in app.selection.details)
                    assert '_summary_comparison' not in app.selection.details
                    assert app.selection.details['_summary_current']['current'] is not None
                    assert 'idpk' in app.summary_targets and 'previous_idpk' not in app.summary_targets
                    assert app.summary_targets['idpk'].cget('text_color')==app._palette['text']
                    print('Mesure actuelle apres clic voisin:', round((time.perf_counter()-start)*1000, 1), 'ms')
                finally:
                    gate.set()
                pump(lambda: '_summary_comparison' in app.selection.details)
            pump(lambda: 'get_attempt_1' in app.summary_targets)
            diagnostic = app.selection.details['get_diagnostic']
            assert diagnostic['status']=='recovered' and diagnostic['failed']
            recovered_key = next(key for key, row in view.rows['R2'].items() if row['modbus']=='1051')
            assert view.trees['R2'].item(recovered_key, 'values')[-1]=='Échec GET'
            assert len(diagnostic['attempts'])==2
            pump(lambda: '_summary_comparison' in app.selection.details)
            assert app.selection.details['_summary_comparison']['current'] is not None
            assert 'idpk' in app.summary_targets
            assert app.summary_targets['get_status'].cget('text')=='Tentative 1 — Échec GET'
            assert app.summary_targets['get_attempt_status_1'].cget('text')=='Tentative 2 — GET réussi'
            assert app.summary_targets['get'].master != app.summary_targets['get_attempt_1'].master
            retry = diagnostic['attempts'][1]['get']
            assert app.summary_targets['get_attempt_1'].cget('text')==retry['text']
            assert not any(key.startswith('get_response_') for key in app.summary_targets)
            visible = ' '.join(str(w.cget('text')) for w in descendants(app.summary_frame) if isinstance(w, gui.ctk.CTkLabel))
            assert 'Tgripper' not in visible and 'reponse gripper' not in visible
            widget = app.summary_targets['get_attempt_1'].selection_text
            for i in range(2):
                widget.event_generate('<ButtonPress-1>', x=15, y=8, time=400000+i*100)
                widget.event_generate('<ButtonRelease-1>', x=15, y=8, time=400020+i*100)
            pump(lambda: bool(traces()) and traces()[0].trace_controller.loaded)
            assert traces()[0].trace_controller.text.index('highlight.first')==str(retry['number']+1)+'.0'
            traces()[0].destroy()
            output = Path(__file__).resolve().parents[1]/'.benchmark'
            for theme in ('Clair', 'Sombre'):
                app._change_theme(theme); pump()
                rendered('R2','1051',False)
                assert color_pixels(widget_image(app.summary_targets['get_attempt_1'].selection_text), '#8ac7ff' if theme=='Sombre' else '#00589b')>3
                widget_image(app).save(output/f'get_retry_{theme}_20261020.png')
            click('R2','1052')
            for counts in ((5,10),(5,30),(5,5)):
                manual(counts)
                pump(lambda:view.data and (view.data['before_count'],view.data['after_count'])==counts)
                assert app._adjacent_counts==counts
            for value in ('-1','1001','3.5','abc'):
                view.before_var.set(value); view.apply_button.invoke(); pump()
                assert view.validation.cget('text') and app._adjacent_counts==(5,5)
            for counts in ((0,0),(12,47),(1000,1000)):
                view.before_var.set(str(counts[0])); view.after_var.set(str(counts[1])); view.apply_button.invoke()
                pump(lambda:view.data and (view.data['before_count'],view.data['after_count'])==counts)
                assert all(len(p['tree'].get_children())<=42 for p in view.panels)
            view.copy_button.invoke(); app.update()
            assert view.clipboard_get()==listing_text(view.data, app._selected_directories)
            assert view.clipboard_get().count(' | ID :')>40, 'Off-screen rows missing'
            view.copy_button._canvas.event_generate('<Enter>'); pump()
            assert view._copy_tip is not None
            view.copy_button._canvas.event_generate('<Leave>'); pump()
            output = Path(__file__).resolve().parents[1]/'.benchmark'
            output.mkdir(exist_ok=True)
            widget_image(view).save(output/'adjacent_copy_20261020.png')
            # One scrollbar, synchronized vertical tables, header/control geometry fixed.
            header_y=view.heading.winfo_rooty()
            tree=view.trees['R1']; first=tree.yview()[0]
            tree.event_generate('<MouseWheel>',delta=-120); pump()
            assert tree.yview()[0]>first and abs(tree.yview()[0]-view.trees['R2'].yview()[0])<.00001
            assert view.heading.winfo_rooty()==header_y
            manual((5,10)); pump(lambda:view.data and view.data['after_count']==10)
            view.close_button.invoke(); app.btn_adjacent.invoke(); pump(lambda:ready('R2','1052'))
            view=app.adjacent_view
            assert (view.before_var.get(),view.after_var.get())==('5','10')
            manual((5,5)); pump(lambda:view.data and view.data['after_count']==5)
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
                manual((5,30)); pump(started.is_set)
                manual((5,10))
                release.set(); pump(finished.is_set)
                pump(lambda:view.data and view.data['after_count']==10)
                assert view.data['after_count']==10
            print('OK options/IDPK : reglages 0..1000, widgets conserves, scroll commun, '
                  'pixels rouges IDPK + GET R1/1054 et orange GET R2/1052, '
                  'preuves GET ouvrables en Synthese, copie, archive et requetes depassees')
        finally:
            release.set()
            app._close_indexed_app(); app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True,cancel_futures=True)


if __name__=='__main__':
    run_gui_check(main)
