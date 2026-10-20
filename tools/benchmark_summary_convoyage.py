"""Temporary-data measurements on the private Windows desktop."""
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import argparse
import json
import statistics
import sys
import tempfile
import time

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import customtkinter as ctk
import trace_store
from trace_store import TraceStore
from summary_view import render_summary
from ui_theme import LIGHT
from tools.ui_test_support import run_gui_check


def widgets(parent):
    return [child for w in parent.winfo_children() for child in [w, *widgets(w)]]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--output', required=True)
    output = Path(parser.parse_args().output)
    if output.exists():
        raise FileExistsError(output)
    root = ctk.CTk()
    root.geometry('1090x820')
    root.report_callback_exception = lambda *error: (_ for _ in ()).throw(error[1])
    frame = ctk.CTkScrollableFrame(root)
    frame.pack(fill='both', expand=True)
    app = SimpleNamespace(summary_frame=frame, _palette=LIGHT, _sources_verified=True,
                          _indexed_complete=True, modbus=ctk.BooleanVar(value=False), selection=None,
                          _compact_trace_label=lambda path: str(path), _load_summary_comparison=lambda data: None,
                          _open_indexed_file=lambda *args: None)
    event = dict(path='C:/temporary/R1/TracOmega.txt', number=3, time=1791960000,
                 text='GET reference TEMP', dims={'l':104, 'H':50, 'L':195}, lot_hint='LOT')
    data = dict(row=dict(ref='TEMP', modbus='1054', sn='SN', lot='LOT'), get=event, pute=event,
                get_diagnostic=dict(status='success', failed=False),
                locations=[dict(event, lot_hint=str(i)) for i in range(45)])
    local = dict(current=event, status='current', revision=1)
    full = dict(local, status='same', different=False, pair=(event, event), message='')
    times, destroyed, retained = [], [], []
    try:
        for _ in range(3):
            data = dict(data)
            data.pop('_summary_current', None); data.pop('_summary_comparison', None)
            render_summary(app, data); root.update()
            stable = app.summary_targets['get']
            for key, result in (('_summary_current', local), ('_summary_comparison', full)):
                old = widgets(frame)
                data[key] = result
                start = time.perf_counter()
                render_summary(app, data); root.update_idletasks()
                times.append((time.perf_counter()-start)*1000)
                destroyed.append(sum(not w.winfo_exists() for w in old))
                retained.append(stable.winfo_exists())
        report = dict(summary=dict(update_ms=statistics.median(times), destroyed_widgets=statistics.median(destroyed),
                                   get_preserved=all(retained), visible_lots=20, total_lots=45))
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)/'R1'/'TracMODBUS.txt'
            path.parent.mkdir()
            lines = []
            for i in range(1000):
                values = ['0']*120
                values[0] = f'13|10:{i%60:02}:00'
                values[2] = '1054' if i%2 else '-1054'
                values[31] = '1054'
                lines.append(';'.join(values))
            path.write_text('\n'.join(lines), encoding='utf8')
            store = TraceStore(Path(folder)/'cache.sqlite3'); store.sync([str(path)])
            elapsed, queries, decodes = [], [], []
            for _ in range(3):
                with patch.object(store, 'modbus_rows', wraps=store.modbus_rows) as rows, \
                     patch.object(trace_store, 'parse_modbus_line', wraps=trace_store.parse_modbus_line) as parse:
                    start = time.perf_counter()
                    if hasattr(store, 'convoyage_details'):
                        result = store.convoyage_details('1054', str(path))
                    else:
                        with store.snapshot():
                            result = store.convoyage_summary('1054', str(path)), store.convoyage_steps('1054', str(path))
                    assert result[0] and result[1]
                    elapsed.append((time.perf_counter()-start)*1000)
                    queries.append(rows.call_count); decodes.append(parse.call_count)
            report['convoyage'] = dict(load_ms=statistics.median(elapsed), queries=statistics.median(queries),
                                       decodes=statistics.median(decodes), source_rows=1000)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(report, indent=2), encoding='utf8')
        print(json.dumps(report))
    finally:
        root.destroy()


if __name__ == '__main__':
    run_gui_check(main)
