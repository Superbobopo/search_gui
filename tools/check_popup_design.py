"""Adjacent presentation checks; conveyor cards are in check_convoyage_cards.py."""
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
from tests.test_occurrence_comparison import measured, trace
from tools.native_test_capture import widget_image, color_pixels
from tools.ui_test_support import run_gui_check
from window_placement import native_window_bounds, monitor_work_area


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def main():
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        for robot in ('R1', 'R2'):
            directory = base/'traces'/robot
            directory.mkdir(parents=True)
            (directory/'TracOmega019.old').write_text(trace(*(
                measured(36000+i*10, str(1000+i), 'REFERENCE_'+str(i), sn='SN'+str(i), lot='260210.004')
                for i in range(40))), encoding='utf8')
            lines = []
            for i in range(60):
                values = ['0']*121
                values[0], values[2] = '13|10:00:00', str(1000+i)
                lines.append(';'.join(values))
            (directory/'TracMODBUS.txt').write_text('\n'.join(lines)+'\n', encoding='utf8')
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
                if errors:
                    raise errors[0][1]
                if condition():
                    return
                time.sleep(.005)
            raise AssertionError('Popup design timeout: '+app.status.get())

        def in_screen(window):
            _, bounds = native_window_bounds(window)
            work = monitor_work_area(window)
            assert bounds[0]>=work[0] and bounds[1]>=work[1] and bounds[2]<=work[2] and bounds[3]<=work[3], (bounds, work)

        def snapshot(view, name):
            pump()
            assert view.title.cget('foreground')==app._palette['text']
            image = widget_image(view)
            assert color_pixels(image, app._palette['text'])>50, 'Capture must contain rendered text'
            image.save(output/name)

        try:
            pump(lambda: app._sources_verified and app._indexed_complete)
            app.mode_radios['ref'].invoke()
            app.keyword.insert(0, 'REFERENCE_10')
            app.btn_search.invoke()
            pump(lambda: app.selection is not None)
            app.btn_adjacent.invoke()
            pump(lambda: getattr(app, 'adjacent_view', None) and app.adjacent_view.data is not None)
            adjacent = app.adjacent_view
            in_screen(app.adj_popup)
            assert gui.ttk.Style(app).lookup('Adjacent.Treeview', 'borderwidth')==0
            assert not adjacent.validation.winfo_ismapped(), 'No empty validation row'
            identities = [str(panel['tree']) for panel in adjacent.panels]
            adjacent.presets[(5, 30)].invoke()
            pump(lambda: adjacent.data['after_count']==30)
            assert identities==[str(panel['tree']) for panel in adjacent.panels]
            snapshot(adjacent, 'popup_adjacent_light_20261019.png')
            app.adj_popup.geometry('780x560'); pump()
            assert adjacent._count_controls.grid_info()['row']==1
            assert adjacent.apply_button.winfo_rootx()+adjacent.apply_button.winfo_width()<app.adj_popup.winfo_rootx()+app.adj_popup.winfo_width()
            heading_y = adjacent.trees['R1'].winfo_rooty()
            adjacent.trees['R1'].event_generate('<MouseWheel>', delta=-120); pump()
            assert adjacent.trees['R1'].winfo_rooty()==heading_y
            assert adjacent.trees['R1'].yview()==adjacent.trees['R2'].yview()
            adjacent.close_button.invoke()
            print('OK adjacent presentation: styles, responsive controls, persistent tables, scrolling, private captures.')
        finally:
            app._close_indexed_app(); app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)


if __name__=='__main__':
    run_gui_check(main)
