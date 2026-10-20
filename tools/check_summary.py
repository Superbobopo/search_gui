"""Focused summary interactions, entirely on a private Windows desktop."""

from contextlib import ExitStack
import os
from pathlib import Path
import sys
import tempfile
import time
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import customtkinter as ctk
import search_gui as gui
import trace_sources as sources
from trace_parser import read_trace
from tools.ui_test_support import run_gui_check


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def main():
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        directory = base / 'traces' / 'trace r2'
        directory.mkdir(parents=True)
        original = read_trace(Path(__file__).resolve().parents[1] / 'tests/fixtures/current_r2.txt')
        current = directory / 'TracOMEGA.txt'
        current.write_text(original, encoding='utf8')
        old = original.replace('14/10/2026', '13/10/2026').replace('14|', '13|')
        old = old.replace('l=48 H=21 L=47', 'l=55 H=21 L=47')
        previous = directory / 'TracOmega029.old'
        previous.write_text(old, encoding='utf8')
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
            raise AssertionError(app.status.get())

        def labels():
            return [w for w in descendants(app.summary_frame) if isinstance(w, ctk.CTkLabel)]

        def text():
            return '\n'.join(str(w.cget('text')) for w in labels())

        def buttons():
            return [w for w in descendants(app.summary_frame) if isinstance(w, ctk.CTkButton)]

        def windows(title):
            return [w for w in app.winfo_children() if isinstance(w, gui.tk.Toplevel)
                    and w.title().startswith(title)]

        stamp = 1000
        def check_open(target, event, contents):
            nonlocal stamp
            pump()
            canvas = app.summary_frame._parent_canvas
            offset = target.winfo_rooty() - canvas.winfo_rooty() + canvas.canvasy(0)
            canvas.yview_moveto(max(0, (offset - 60) / max(1, app.summary_frame.winfo_height())))
            pump()
            widget = target.selection_text
            widget.event_generate('<Motion>', x=4, y=4)
            for click in range(2):
                when = stamp + click * 100
                widget.event_generate('<ButtonPress-1>', x=4, y=4, time=when)
                widget.event_generate('<ButtonRelease-1>', x=4, y=4, time=when + 20)
            stamp += 1000
            pump(lambda: bool(windows('Vue de')))
            view = windows('Vue de')[0]
            widget = next(w for w in descendants(view) if isinstance(w, gui.tk.Text))
            pump(lambda: bool(widget.tag_ranges('highlight')))
            begin, end = widget.tag_ranges('highlight')
            assert widget.get(begin, end) == event['text']
            assert int(str(begin).split('.')[0]) == event['number'] + 1
            assert widget.get('1.0', 'end-1c').splitlines() == contents.splitlines()
            view.destroy()

        def check_copy(target, begin=0, end=None, expected=None):
            nonlocal stamp
            pump()
            widget = target.selection_text
            canvas = app.summary_frame._parent_canvas
            offset = target.winfo_rooty() - canvas.winfo_rooty() + canvas.canvasy(0)
            canvas.yview_moveto(max(0, (offset - 60) / max(1, app.summary_frame.winfo_height())))
            pump()
            content = widget.get('1.0', 'end-1c')
            end = len(content) if end is None else end
            first = widget.bbox(f'1.0+{begin}c')
            last = widget.bbox(f'1.0+{end}c')
            assert first and last, (content, first, last)
            # Drag over the visible text, exactly as a user selects a substring.
            widget.event_generate('<ButtonPress-1>', x=first[0], y=first[1] + 3, time=stamp)
            widget.event_generate('<B1-Motion>', x=last[0], y=last[1] + 3, state=256, time=stamp + 400)
            widget.event_generate('<ButtonRelease-1>', x=last[0], y=last[1] + 3, time=stamp + 500)
            stamp += 1500
            selected = widget.get('sel.first', 'sel.last')
            assert selected == (content[begin:end] if expected is None else expected), (content, selected)
            with patch.object(widget, 'clipboard_clear') as clear, patch.object(widget, 'clipboard_append') as copy:
                # Private desktop only: give the test keyboard focus, then send Ctrl+C.
                widget.focus_force()
                pump()
                widget.event_generate('<Control-KeyPress-c>')
                pump()
                clear.assert_called_once_with()
                copy.assert_called_once_with(selected)
                # Native Windows menus wait for input; dispatch right-click with only
                # their native display mocked, then invoke the real copy command.
                with patch.object(widget.copy_menu, 'tk_popup') as popup:
                    widget.event_generate('<Button-3>', x=3, y=3)
                    pump()
                    popup.assert_called_once()
                assert widget.copy_menu.entrycget(0, 'label') == 'Copier'
                widget.copy_menu.invoke(0)
                widget.copy_menu.unpost()
                assert copy.call_count == 2 and copy.call_args.args == (selected,)
            before = widget.get('1.0', 'end-1c')
            widget.event_generate('<KeyPress-x>')
            widget.event_generate('<KeyPress-Delete>')
            pump()
            assert widget.get('1.0', 'end-1c') == before
            assert not windows('Vue de'), 'Dragging a selection must not open a trace'

        try:
            assert app.result_mode.get() == 'Synthèse'
            pump(lambda: app._sources_verified)
            spy = stack.enter_context(patch.object(app.store, 'comparison_for_occurrence', wraps=app.store.comparison_for_occurrence))
            app.keyword.insert(0, 'A0061574AC')
            app.modbus.set(False)
            app.reference.set(True)
            app.btn_search.invoke()
            pump(lambda: app.selection is not None and '_summary_comparison' in app.selection.details)
            pump(lambda: 'Précédente' in text())
            selection = app.selection
            for value in ('Référence', 'A0061574AC', 'ID', '1194', 'SN', '343272', 'CodeLot', '261013.127'):
                assert value in text(), value
            assert not any(isinstance(w, ctk.CTkButton) and 'Ouvrir' in str(w.cget('text')) for w in buttons())
            assert 'Ligne originale' not in text() and 'Écart' not in text()
            assert not any(w.cget('text') in ('GET', 'PUTE', 'IDPK') for w in labels())
            assert app.summary_targets['get'].winfo_rooty() < app.summary_targets['pute'].winfo_rooty()
            assert app.summary_targets['pute'].winfo_rooty() < app.summary_targets['idpk'].winfo_rooty()
            recent = app.summary_targets['idpk']
            previous_values = app.summary_targets['previous_idpk']
            assert recent.winfo_rooty() == previous_values.winfo_rooty()
            assert recent.winfo_rootx() < previous_values.winfo_rootx()
            assert recent.cget('text_color') == ('#b42318', '#ff9b91')
            assert previous_values.cget('text_color') != recent.cget('text_color')
            for name in ('Référence', 'ID', 'SN', 'CodeLot'):
                widget = next(w for w in labels() if w.cget('text') == name)
                assert widget.cget('anchor') == 'center' and widget.cget('justify') == 'center'
            for value in ('A0061574AC', '1194', '343272', '261013.127'):
                target = next(w for w in labels() if w.cget('text') == value)
                check_copy(target)
            check_copy(next(w for w in labels() if w.cget('text') == '343272'), 1, 4, '432')
            check_copy(app.summary_targets['idpk'], 0, 6, 'l = 48')
            get_text = app.summary_targets['get'].cget('text')
            begin = get_text.index('A0061574AC')
            check_copy(app.summary_targets['get'], begin, begin + 10, 'A0061574AC')
            for key in ('get', 'pute'):
                widget = app.summary_targets[key].selection_text
                assert widget.dlineinfo('end-1c') is not None, 'Wrapped text must remain fully visible'
            pair = selection.details['_summary_comparison']['pair']
            assert pair and pair[1]['path'].path == str(previous)
            # Saved generations must open even after their original paths disappear.
            stack.enter_context(patch.object(app.store, 'sync', return_value=dict(changed=0, removed=0, failures=[])))
            current.unlink()
            previous.unlink()
            check_open(app.summary_targets['previous_idpk'], pair[1], old)
            for key in ('get', 'idpk', 'pute'):
                check_open(app.summary_targets[key], pair[0] if key == 'idpk' else selection.details[key], original)
            app.btn_compare_idpk.invoke()
            pump(lambda: bool(windows('Comparaison IDPK')))
            popup = windows('Comparaison IDPK')[0]
            popup_text = ' '.join(str(w.cget('text')) for w in descendants(popup) if isinstance(w, gui.tk.Label))
            assert 'Δ l = -7' in popup_text
            numeric = next(w for w in descendants(popup) if isinstance(w, gui.tk.Label)
                           and 'l = 48' in str(w.cget('text')))
            numeric.selection_text.tag_add('sel', '1.0', 'end-1c')
            with patch.object(numeric.selection_text, 'clipboard_clear'), \
                    patch.object(numeric.selection_text, 'clipboard_append') as copy:
                numeric.selection_text.event_generate('<<Copy>>')
                copy.assert_called_once_with(numeric.cget('text'))
            popup.destroy()
            calls = spy.call_count
            # Exactly two units is accepted, like the popup; changing theme must preserve this.
            matching = [dict(pair[0], dims=dict(l=53, H=21, L=47)), pair[1]]
            original_result = selection.details['_summary_comparison']
            selection.details['_summary_comparison'] = dict(original_result, pair=matching, current=matching[0], different=False)
            app._render_summary(selection.details)
            assert app.summary_targets['idpk'].cget('text_color') != ('#b42318', '#ff9b91')
            selection.details['_summary_comparison'] = original_result
            app._sources_verified = False
            app._render_summary(selection.details)
            assert 'Comparaison disponible après vérification' in text()
            assert spy.call_count == calls
            app._sources_verified = True
            app._render_summary(selection.details)
            app.result_switch._buttons_dict['Classique'].invoke()
            pump(lambda: app.classic_frame.winfo_ismapped())
            assert 'RefProduit' in app.txt.get('1.0', 'end')
            app.result_switch._buttons_dict['Synthèse'].invoke()
            app._change_theme('Sombre')
            pump()
            assert app.selection is selection and spy.call_count == calls
            assert all(getattr(w, '_selectable_display', False) for w in descendants(app.summary_frame)
                       if isinstance(w, gui.tk.Text))
            check_copy(next(w for w in labels() if w.cget('text') == '261013.127'))
            # Many lots expand in the single outer scroll area; none is discarded.
            event = selection.details['locations'][0]
            selection.details['locations'] = [dict(event, lot_hint=f'LOT-{i:02}') for i in range(45)]
            app._render_summary(selection.details)
            assert 'LOT-19' in text() and 'LOT-20' not in text()
            next(w for w in buttons() if str(w.cget('text')).startswith('Afficher plus')).invoke()
            assert 'LOT-39' in text() and 'LOT-40' not in text()
            next(w for w in buttons() if str(w.cget('text')).startswith('Afficher plus')).invoke()
            assert 'LOT-44' in text()
            check_open(app.summary_targets['location'], event, original)
            target = next(w for w in labels() if str(w.cget('text')).startswith('Lot LOT-00'))
            check_copy(target, 4, 10, 'LOT-00')
            # Wheel input over a selectable field scrolls only the outer panel.
            canvas = app.summary_frame._parent_canvas
            canvas.yview_moveto(0)
            pump()
            target = next(w for w in labels() if w.cget('text') == '343272')
            inner = target.selection_text.yview()
            target.selection_text.event_generate('<MouseWheel>', delta=-120)
            pump()
            assert canvas.yview()[0] > 0 and target.selection_text.yview() == inner
            assert all(getattr(w, '_selectable_display', False) for w in descendants(app.summary_frame)
                       if isinstance(w, gui.tk.Text))
            # An older file has no predecessor: say so rather than inventing a comparison.
            app.btn_next_file.invoke()
            pump(lambda: app.selection is not None and app.selection.path != selection.path
                 and '_summary_comparison' in app.selection.details)
            assert 'Aucune comparaison disponible' in text()
            assert app.summary_targets['idpk'].cget('text_color') != ('#b42318', '#ff9b91')
            print('OK GET/PUTE/IDPK compacts dans cet ordre, sans titres/boutons/ligne originale/tableau, '
                  'mesures cote a cote, rouge selon seuil >2, vrais doubles-clics sur archives absentes '
                  'et sur lots avec texte/ligne/surlignage exacts ; theme/Classique et 45 lots conserves')
            print('OK selection par glisser, copie Ctrl+C et menu Copier : reference/ID/SN/CodeLot/mesures '
                  'et popup ; sous-chaine, lecture seule et theme sombre ; presse-papiers simule')
        finally:
            app._close_indexed_app()
            app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)


if __name__ == '__main__':
    run_gui_check(main)
