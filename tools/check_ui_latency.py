"""Measure UI transitions on a private desktop, with temporary data only."""

from contextlib import ExitStack
import cProfile
import io
import json
import os
from pathlib import Path
import pstats
import statistics
import sys
import tempfile
import time
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import search_gui as gui
import trace_sources as sources
from trace_parser import read_trace
from tools.ui_test_support import run_gui_check


def main():
    output = Path(sys.argv[sys.argv.index('--output') + 1])
    assert not output.exists(), output
    report = {}
    with ExitStack() as stack:
        base = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        directory = base / 'traces' / 'R2'
        directory.mkdir(parents=True)
        text = read_trace(Path(__file__).resolve().parents[1] / 'tests/fixtures/current_r2.txt')
        (directory / 'TracOMEGA.txt').write_text(text, encoding='utf8')
        (directory / 'TracOmega029.old').write_text(text.replace('14/10/2026', '13/10/2026')
            .replace('14|', '13|').replace('l=48 H=21 L=47', 'l=55 H=21 L=47'), encoding='utf8')
        stack.enter_context(patch.object(sources, 'application_directory', return_value=str(base)))
        stack.enter_context(patch.object(sources, 'desktop_directory', return_value=str(base / 'absent')))
        stack.enter_context(patch.object(sources, 'FIXED_DIRECTORY', str(base / 'missing')))
        stack.enter_context(patch.dict(os.environ, LOCALAPPDATA=str(base / 'AppData')))
        app = gui.SearchApp()
        errors = []
        app.report_callback_exception = lambda *error: errors.append(error)

        def pump(condition=lambda: True):
            deadline = time.perf_counter() + 10
            while time.perf_counter() < deadline:
                app.update()
                if errors:
                    raise errors[0][1]
                if condition():
                    return
                time.sleep(.005)
            raise AssertionError('UI timeout')

        def sample(name, action):
            start = time.perf_counter()
            action()
            callback = time.perf_counter()
            pump()
            # Drain queued map/theme work, measured separately from the callback.
            app.after(20, lambda: None)
            end = time.perf_counter() + .025
            while time.perf_counter() < end:
                pump()
                time.sleep(.002)
            settled = time.perf_counter()
            report.setdefault(name, []).append(dict(callback_ms=(callback - start) * 1000,
                                                     settle_ms=(settled - start) * 1000))

        try:
            pump(lambda: app._sources_verified)
            app.keyword.insert(0, 'A0061574AC')
            app.modbus.set(False)
            app.reference.set(True)
            app.btn_search.invoke()
            pump(lambda: app.selection is not None and '_summary_comparison' in app.selection.details)
            selection = app.selection
            event = selection.details['locations'][0]
            selection.details['locations'] = [dict(event, lot_hint=f'LOT-{i}') for i in range(20)]
            app._render_summary(selection.details)
            pump()
            targets = dict(app.summary_targets)
            source_search = app.store.search
            search_calls = stack.enter_context(patch.object(app.store, 'search', wraps=source_search))
            compare_calls = stack.enter_context(patch.object(app.store, 'comparison_for_occurrence', wraps=app.store.comparison_for_occurrence))
            field = next(w for w in app.summary_frame.winfo_children()[0].winfo_children()[2].winfo_children()
                         if hasattr(w, 'selection_text') and w.cget('text') == '343272')
            field.selection_text.tag_add('sel', '1.0', '1.3')
            for _ in range(3):
                sample('summary_to_classic', app.result_switch._buttons_dict['Classique'].invoke)
                sample('classic_to_summary', app.result_switch._buttons_dict['Synthèse'].invoke)
                assert app.selection is selection
                assert app.summary_targets == targets
                assert field.selection_text.get('sel.first', 'sel.last') == '343'
            profiler = cProfile.Profile()
            profiler.enable()
            app.result_switch._buttons_dict['Classique'].invoke()
            app.result_switch._buttons_dict['Synthèse'].invoke()
            pump()
            profiler.disable()
            stream = io.StringIO()
            pstats.Stats(profiler, stream=stream).strip_dirs().sort_stats('cumtime').print_stats(14)
            report['profile_modes'] = stream.getvalue()
            for _ in range(3):
                sample('search_to_tracking', app.btn_tracking_robot.invoke)
                sample('tracking_to_convoyage', app.btn_convoyage_menu.invoke)
                sample('convoyage_to_search', app.btn_search_menu.invoke)
                assert app.selection is selection and app.summary_targets == targets
                assert len(app._panels) == 3
            assert search_calls.call_count == 0 and compare_calls.call_count == 0
            assert field.selection_text.get('sel.first', 'sel.last') == '343'
            profiler = cProfile.Profile()
            profiler.enable()
            app.btn_tracking_robot.invoke()
            pump()
            app.btn_convoyage_menu.invoke()
            pump()
            app.btn_search_menu.invoke()
            pump()
            profiler.disable()
            stream = io.StringIO()
            pstats.Stats(profiler, stream=stream).strip_dirs().sort_stats('cumtime').print_stats(18)
            report['profile_navigation'] = stream.getvalue()
            # A theme change recolors the same fields, preserving copy/selection and pagination.
            app._change_theme('Sombre')
            pump()
            assert app.summary_targets == targets
            assert field.selection_text.get('sel.first', 'sel.last') == '343'
            assert search_calls.call_count == 0 and compare_calls.call_count == 0
            # Cache reuse must never hide a new source revision after returning to Search.
            old_hits = sum(len(hits) for _, hits in app._groups)
            revision = app._index_content_revision
            app.btn_tracking_robot.invoke()
            with (directory / 'TracOMEGA.txt').open('a', encoding='utf8') as stream:
                stream.write('\n  14|09:00:00.0 RefProduit=A0061574AC<Cr>ModBusId=2222<Cr>SN=NEW<Cr>\n')
            pump(lambda: app._index_content_revision > revision)
            app.btn_search_menu.invoke()
            pump(lambda: app.selection is not None and sum(len(hits) for _, hits in app._groups) > old_hits)
            assert search_calls.call_count >= 1
            report['regressions'] = 'cached widgets, selection/copy, bounded panels, no redundant SQL, theme, source revision refresh'
            print('PROFILE NAVIGATION\n' + report['profile_navigation'])
            print('PROFILE MODES\n' + report['profile_modes'])
            for name, samples in report.items():
                if isinstance(samples, list):
                    print(name, 'callback median', round(statistics.median(s['callback_ms'] for s in samples), 2),
                          'ms ; settled median', round(statistics.median(s['settle_ms'] for s in samples), 2), 'ms')
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_text(json.dumps(report, indent=2), encoding='utf8')
        finally:
            app._close_indexed_app()
            app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True, cancel_futures=True)


if __name__ == '__main__':
    run_gui_check(main)
