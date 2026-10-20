"""Exercise the shared UI on a private Windows desktop and temporary indexes."""
from contextlib import ExitStack
import os
from pathlib import Path
import sys
import tempfile
import time
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import search_gui as gui
import trace_sources as sources
from ui_components import ActionButton
from tests.test_occurrence_comparison import measured, ticket, trace
from tools.ui_test_support import run_gui_check
from tools.native_test_capture import widget_image, color_pixels
from window_placement import monitor_work_area, native_window_bounds


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def main():
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        paths = {}
        for robot in ('R1', 'R2'):
            path = base/'traces'/robot/'TracOmega019.old'
            path.parent.mkdir(parents=True)
            missing = ticket(3600, '900', 'MISSING', sn='SN_'+robot, lot='LOT_'+robot).splitlines(True)[0]
            text = trace(missing, measured(7200, '101', 'MEASURE', sn='MATCH', lot='SAME',
                         dims=(104, 50, 195) if robot=='R1' else (104, 48, 189)))
            if robot=='R2':
                text = text.replace('14/10/2026', '13/10/2026').replace('14|', '13|')
            path.write_text(text, encoding='utf8')
            paths[robot] = path
        modbus = base/'traces'/'R1'/'TracMODBUS.txt'
        lines = []
        for ident, destination in (('904', 6), ('905', 7)):
            for col, value, hour in ((2, ident, '10:00:00'), (36, ident, '10:01:00'), (destination, '-'+ident, '10:02:00')):
                row = ['0']*121; row[0], row[col] = '14|'+hour, value
                lines.append(';'.join(row))
        modbus.write_text('\n'.join(lines)+'\n', encoding='utf8')
        stack.enter_context(patch.object(sources, 'application_directory', return_value=str(base)))
        stack.enter_context(patch.object(sources, 'desktop_directory', return_value=str(base/'absent')))
        stack.enter_context(patch.object(sources, 'FIXED_DIRECTORY', str(base/'missing')))
        stack.enter_context(patch.dict(os.environ, LOCALAPPDATA=str(base/'AppData')))
        app = gui.SearchApp()
        errors = []
        app.report_callback_exception = lambda *error: errors.append(error)
        output = Path(__file__).resolve().parents[1]/'.benchmark'
        output.mkdir(exist_ok=True)

        def pump(condition=lambda: True):
            deadline = time.monotonic()+12
            while time.monotonic()<deadline:
                app.update()
                if errors: raise errors[0][1]
                if condition(): return
                time.sleep(.004)
            raise AssertionError('Unified UI timeout: '+app.status.get())

        def popups(attr):
            return [w for w in app.winfo_children() if getattr(w, attr, None) is not None]

        def screenshot(view, name):
            pump()
            image = widget_image(view)
            assert color_pixels(image, app._palette['text'])>30
            if app._appearance=='Dark':
                assert color_pixels(image, '#ffffff')<1000, 'White background in Dark mode'
            image.save(output/name)

        def card_click(view, index, double=False):
            view.cards.see(index); pump()
            card = next(card for card, _ in view.cards.cards if card.index==index)
            widget = card.title._label
            widget.event_generate('<ButtonPress-1>', x=12, y=10)
            widget.event_generate('<ButtonRelease-1>', x=12, y=10)
            if double:
                for i in range(2):
                    widget.event_generate('<ButtonPress-1>', x=12, y=10, time=100000+i*100)
                    widget.event_generate('<ButtonRelease-1>', x=12, y=10, time=100020+i*100)
            pump()

        def button_style(root):
            controls = [w for w in descendants(root) if isinstance(w, ActionButton)]
            assert controls
            for button in controls:
                assert button.cget('height')==32 and button.cget('corner_radius')==6
                assert button.cget('font')[1] in (12, 14)  # Readable embedded conveyor controls.
                if button.cget('state')=='disabled':
                    assert button.cget('fg_color')==app._palette['raised']

        def inside(window):
            _, rect = native_window_bounds(window)
            area = monitor_work_area(app)
            assert rect[0]>=area[0] and rect[1]>=area[1] and rect[2]<=area[2] and rect[3]<=area[3]

        def close_extras():
            for window in list(app.winfo_children()):
                if isinstance(window, gui.tk.Toplevel): window.destroy()
            pump()

        try:
            pump(lambda: app._sources_verified and app._indexed_complete)
            sync = app.store.sync
            stack.enter_context(patch.object(app.store, 'sync', return_value=dict(changed=0, removed=0, failures=[])))
            app.btn_erreurs.invoke(); pump(lambda: bool(popups('robot_errors_view')))
            view = popups('robot_errors_view')[0].robot_errors_view
            assert len(view.errors)==2 and view.selected_index==0
            assert {view.card_content(i)['badge'] for i in range(2)}=={'R1', 'R2'}
            assert view.values['sn'].cget('text') in ('SN_R1', 'SN_R2')
            assert view.values['lot'].cget('text') in ('LOT_R1', 'LOT_R2')
            assert 'RefProduit=MISSING' in view.ticket.cget('text')
            original = app.selection
            card_click(view, 1)
            assert app.selection is original
            detail_ids = str(view.fields), str(view.ticket), str(view.status)
            for theme in ('Clair', 'Sombre'):
                app._change_theme(theme); pump()
                assert color_pixels(widget_image(view.status.value.selection_text), app._palette['error_text'])>3
                screenshot(view, f'robot_errors_{theme}_20261019.png')
                assert detail_ids==(str(view.fields), str(view.ticket), str(view.status))
                button_style(view)
            value = view.values['sn'].selection_text
            copied = []
            with patch.object(value, 'clipboard_clear'), patch.object(value, 'clipboard_append', side_effect=copied.append):
                value.tag_add('sel', '1.0', '1.2'); value.event_generate('<<Copy>>')
                assert copied[-1]=='SN'
            view.provenance_button.invoke(); pump()
            assert 'Génération' in view.provenance.cget('text')
            selected = view.errors[view.selected_index]
            # The original source disappears; opening must still select the exact archived generation.
            for path in paths.values(): path.unlink()
            assert not sync([str(modbus)])['failures']
            view.open_button.invoke()
            pump(lambda: app.selection is not None and app.selection.path==selected['filepath'] and app.selection.number==selected['number'])
            pump(lambda: getattr(app, 'adjacent_view', None) and app.adjacent_view.data is not None)
            assert app.selection.reference=='MISSING'
            other = view.errors[0]
            card_click(view, 0, double=True)
            pump(lambda: app.selection is not None and app.selection.path==other['filepath']
                 and app.selection.number==other['number'])
            view.winfo_toplevel().geometry('780x600'); pump()
            assert view.split.vertical
            close_extras()
            # Empty diagnostics and a long list reuse bounded card instances.
            with patch.object(app.store, 'robot_errors', return_value=[]):
                app.btn_erreurs.invoke(); pump(lambda: bool(popups('robot_errors_view')))
            empty = popups('robot_errors_view')[0].robot_errors_view
            assert empty.selected_index is None and empty.status.role=='positive'
            empty.close_button.invoke(); pump()
            app._render_robot_errors([dict(selected, modbus_id=str(i)) for i in range(10000)], [])
            large = popups('robot_errors_view')[0].robot_errors_view
            pump(); assert len(large.cards.cards)<16
            large.close_button.invoke(); pump()
            # Common MODBUS file stored in R1 contains both robot destinations.
            app.btn_convoyage_menu.invoke()
            workspace = app.convoyage_workspace
            workspace.buttons['anomalies'].invoke()
            pump(lambda: 'anomalies' in workspace.views and workspace.views['anomalies'].state=='ready')
            convoy = workspace.views['anomalies']
            for index in range(2):
                card_click(convoy, index); pump(lambda: convoy.state=='ready')
                assert convoy.card_content(index)['badge']=='R1/R2'
                assert convoy.detail_source.cget('text')=='TracMODBUS.txt'
                assert ('T5 / R1' if convoy.errors[index][1]=='904' else 'T5 / R2') in convoy.waiting.value.cget('text')
            screenshot(convoy, 'modbus_shared_20261019.png')
            # Normal path and recap use the same skin and no basename filtering.
            workspace.buttons['parcours'].invoke()
            workspace.entry.insert(0, '904')
            workspace.search_button.invoke()
            pump(lambda: 'parcours' in workspace.views and workspace.views['parcours'].state=='ready')
            normal = workspace.views['parcours']
            assert normal.card_content(0)['badge']=='R1/R2'
            assert len(normal.data[1])==3
            app.show_convoyage_details('904'); pump(lambda: bool(popups('diagnostic_view')))
            recap = popups('diagnostic_view')[0].diagnostic_view
            assert len(recap.rows)==3
            recap.origins_toggle.invoke(); pump()
            assert any(str(modbus) in label.cget('text') for label in recap.sources)
            screenshot(recap, 'convoyage_recap_20261019.png')
            close_extras()
            # Actual comparison button, exact identity, colors, deltas and trace highlights.
            app.btn_search_menu.invoke()
            app.mode_radios['ref'].invoke(); app.keyword.delete(0,'end'); app.keyword.insert(0,'MEASURE')
            app.btn_search.invoke(); pump(lambda: app.selection is not None and app.selection.reference=='MEASURE' and '_summary_comparison' in app.selection.details)
            app.btn_compare_idpk.invoke(); pump(lambda: bool(popups('diagnostic_view')))
            comparison = popups('diagnostic_view')[0].diagnostic_view
            assert comparison.context['different'] and comparison.cards[0].role=='waiting'
            assert any('Δ L = +6' in label.cget('text') for label in comparison.labels)
            screenshot(comparison, 'idpk_unified_20261019.png')
            button_style(comparison)
            comparison.close_button.invoke(); pump()
            app.btn_open_trace.invoke(); pump(lambda: bool(popups('trace_controller')))
            trace_window = popups('trace_controller')[0]
            controller = trace_window.trace_controller
            pump(lambda: controller.loaded)
            assert controller.text.get('1.0','end-1c').splitlines()==app.store.file_lines(app.selection.path)
            button_style(trace_window)
            controller.buttons['get'][1].invoke(); pump()
            assert controller.text.tag_ranges('getline')
            assert controller.text.tag_cget('getline', 'foreground')==app._palette['text']
            x,y,width,height = controller.text.bbox('getline.first')
            rendered = widget_image(controller.text).crop((0,y,controller.text.winfo_width(),y+height))
            assert color_pixels(rendered, app._palette['text'])>3
            screenshot(trace_window, 'trace_unified_toolbar_20261019.png')
            close_extras()
            get = app.summary_targets['get'].selection_text
            for _ in range(2):
                get.event_generate('<ButtonPress-1>', x=15, y=10)
                get.event_generate('<ButtonRelease-1>', x=15, y=10)
            pump(lambda: bool(popups('trace_controller')))
            highlighted = popups('trace_controller')[0].trace_controller
            pump(lambda: highlighted.loaded)
            expected = app.selection.get_event['number'] if hasattr(app.selection, 'get_event') else app.selection.details['get']['number']
            assert highlighted.text.index('highlight.first')==f'{expected+1}.0'
            assert highlighted.text.get('highlight.first','highlight.last')==app.selection.details['get']['text']
            assert highlighted.text.tag_cget('highlight', 'background')==app._palette['warning_bg']
            assert highlighted.text.tag_cget('highlight', 'foreground')==app._palette['text']
            close_extras()
            app.btn_tracking_robot.invoke(); pump()
            app.tracking_keyword.delete(0,'end'); app.tracking_keyword.insert(0,'MEASURE')
            app.btn_search.invoke(); pump(lambda: bool(app.tracking_view.rows))
            assert len(app.tracking_view.rows)==2
            button_style(app)
            print('OK UI commune : erreurs robot R1/R2, detail/SN/lot/copie, archives et navigation exacte, '
                  '10000 cartes bornees, MODBUS commun dans R1 vers deux destinations, parcours et recap, '
                  'IDPK/rouge/delta, trace integrale/GET, suivi, boutons communs, Clair/Sombre et fenetre etroite.')
        finally:
            app._close_indexed_app(); app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)


if __name__=='__main__':
    run_gui_check(main)
