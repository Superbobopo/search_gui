"""GET report and neighbor performance on the existing cache, read-only/private UI."""
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import patch
import argparse
import json
import os
import sys
import tempfile
import time

PROJECT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(PROJECT))
from tools.benchmark_real_history import ReadOnlyStore
from tools.ui_test_support import run_gui_check
from tools.native_test_capture import widget_image
from trace_store import cache_path
from get_errors_view import error_key
import indexed_ui
import search_gui as gui


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--output',required=True)
    parser.add_argument('--database')
    args=parser.parse_args();output=Path(args.output).resolve()
    if output.exists():raise FileExistsError(output)
    database=Path(args.database) if args.database else cache_path([str(PROJECT/'traces')])
    before=database.stat();store=ReadOnlyStore(database);revision=store.analysis_revision()
    groups=store.search('ModBusId=1054',True)
    path,number=next((p,h[0][0]) for p,h in groups if Path(p).name.lower()=='tracomega019.old' and 'trace r1' in str(p).lower())
    report={};start=time.perf_counter();neighbors=store.paired_adjacent(path,'1054',number)
    report['neighbors_first_ms']=(time.perf_counter()-start)*1000
    start=time.perf_counter();assert store.paired_adjacent(path,'1054',number)==neighbors
    report['neighbors_warm_before_ms']=(time.perf_counter()-start)*1000
    calls=[];original=store._rows
    def rows(*args,**kw):
        value=original(*args,**kw);calls.append(len(value));return value
    with patch.object(store,'_rows',side_effect=rows):
        start=time.perf_counter();data=store.get_errors()
        report.update(report_ms=(time.perf_counter()-start)*1000,errors=len(data['errors']),queries=len(calls),transferred_lines=sum(calls))
    start=time.perf_counter();assert store.paired_adjacent(path,'1054',number)==neighbors
    report['neighbors_warm_after_ms']=(time.perf_counter()-start)*1000
    report['cache_weight']=store._analysis_cache.weight
    metadata=store.metadata()
    def ready(app):
        app._index_messages.put(('ready',store,metadata,True))
        app._index_messages.put(('complete',dict(changed=0,removed=0,failures=[]),metadata))
    with ExitStack() as stack:
        folder=stack.enter_context(tempfile.TemporaryDirectory())
        stack.enter_context(patch.dict(os.environ,LOCALAPPDATA=folder))
        stack.enter_context(patch.object(indexed_ui.IndexedAppMixin,'_index_loop',ready))
        app=gui.SearchApp();app.geometry('1090x820+0+0');errors=[];beats=[]
        app.report_callback_exception=lambda *error:errors.append(error)
        def pump(condition=lambda:True,settle=0):
            until=time.perf_counter()+20;stable=None
            while time.perf_counter()<until:
                app.update()
                if errors:raise errors[0][1]
                if condition():
                    stable=stable or time.perf_counter()
                    if time.perf_counter()-stable>=settle:return
                else:stable=None
                time.sleep(.002)
            raise AssertionError(app.status.get())
        def beat():
            beats.append(time.perf_counter());app.after(8,beat)
        app.after(8,beat)
        view=None
        try:
            pump(lambda:app._sources_verified)
            start=time.perf_counter();beats.clear();beats.append(start)
            app.btn_get_errors._canvas.event_generate('<ButtonPress-1>',x=5,y=5)
            popup=app.get_errors_popup;view=popup.get_errors_view
            pump(lambda:view.data is not None and view._insert_timer is None)
            report['ui_report_ms']=(time.perf_counter()-start)*1000
            report['ui_heartbeat_max_ms']=max(b-a for a,b in zip(beats,beats[1:]))*1000
            assert view.data==data
            row=next(r for r in data['errors'] if r['robot']=='R2' and r['modbus']=='1052'
                     and Path(r['path']).name.lower()=='tracomega029.old' and r['ref']=='A0110543AC')
            iid=next(i for i,r in view.rows.items() if error_key(r)==error_key(row))
            view.tree.see(iid);pump();x,y,w,h=view.tree.bbox(iid)
            view.tree.event_generate('<ButtonPress-1>',x=x+10,y=y+h//2,time=5000)
            view.tree.event_generate('<ButtonRelease-1>',x=x+10,y=y+h//2,time=5020)
            view.tree.focus_force();view.tree.event_generate('<KeyPress-Return>')
            def traces():return [w for w in app.winfo_children() if getattr(w,'trace_controller',None)]
            pump(lambda:traces() and traces()[0].trace_controller.loaded)
            controller=traces()[0].trace_controller
            assert controller.text.get('highlight.first','highlight.last')==row['error']['text']
            assert controller.text.index('highlight.first')==str(row['error']['number']+1)+'.0'
            traces()[0].destroy()
            app._change_theme('Sombre');pump(settle=.1)
            widget_image(popup).save(output.with_suffix('.png'))
            report.update(database=str(database),revision=revision,files=len(metadata),exact_neighbors_unchanged=True,
                          actual_R2_1052_error_highlight=True,source_size=before.st_size)
        finally:
            app._close_indexed_app();app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True,cancel_futures=True)
            if view:view.tasks.executor.shutdown(wait=True,cancel_futures=True)
    assert (before.st_size,before.st_mtime_ns)==(database.stat().st_size,database.stat().st_mtime_ns)
    assert store.analysis_revision()==revision
    output.write_text(json.dumps(report,indent=2),encoding='utf8');print(json.dumps(report,indent=2))


if __name__=='__main__':run_gui_check(main)
