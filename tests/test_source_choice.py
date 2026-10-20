"""Targeted regressions for source selection; no graphical window is opened."""

from pathlib import Path
from queue import Queue
from types import SimpleNamespace
import tempfile
import json
import threading
import unittest
from unittest.mock import Mock, patch

import indexed_ui
import trace_sources
from trace_store import cache_path
from omega_history import HistoryUnavailable


class SourceChoiceTests(unittest.TestCase):
    def setUp(self):
        preference = patch.object(indexed_ui, 'preferred_trace_directories', return_value=None)
        self.preferred = preference.start()
        self.addCleanup(preference.stop)

    def test_local_desktop_policy_bypasses_choice_and_other_source_discovery(self):
        self.preferred.return_value = ('desktop/traces',)
        with patch.object(indexed_ui, 'available_trace_sources') as available:
            self._run_worker((), ('desktop/traces',))
            available.assert_not_called()

    def test_preference_is_local_to_machine_and_tracks_redirected_desktop(self):
        with tempfile.TemporaryDirectory() as folder:
            filename = Path(folder) / 'source_preferences.json'
            with patch.object(trace_sources, 'source_preferences_path', return_value=str(filename)), \
                    patch.object(trace_sources, 'machine_identifier', return_value='this-pc'), \
                    patch.object(trace_sources, 'desktop_directory', return_value=folder):
                self.assertIsNone(trace_sources.preferred_trace_directories())
                settings = dict(version=1, machine_id='this-pc', startup_source='desktop')
                filename.write_text(json.dumps(settings), encoding='utf8')
                self.assertEqual(trace_sources.preferred_trace_directories(), (str(Path(folder) / 'traces'),))
                with patch.object(trace_sources, 'machine_identifier', return_value='another-pc'):
                    self.assertIsNone(trace_sources.preferred_trace_directories())
                with patch.object(trace_sources, 'machine_identifier', return_value=None):
                    self.assertIsNone(trace_sources.preferred_trace_directories())
                filename.write_text('invalid json', encoding='utf8')
                self.assertIsNone(trace_sources.preferred_trace_directories())

    def test_only_populated_distinct_roots_are_offered(self):
        with tempfile.TemporaryDirectory() as folder:
            roots = [Path(folder) / name for name in ('apoteka', 'project', 'desktop')]
            for root in roots:
                root.mkdir()
            (roots[0] / 'TracOmega.txt').write_text('one')
            (roots[1] / 'TracOmega.sav').write_text('ignored')
            (roots[2] / 'R2').mkdir()
            (roots[2] / 'R2' / 'TracMODBUS.txt').write_text('two')
            with patch.object(trace_sources, 'trace_directories', return_value=roots):
                self.assertEqual(trace_sources.available_trace_sources(),
                                 (('Apoteka', str(roots[0])), ('Bureau', str(roots[2]))))
            with patch.object(trace_sources, 'trace_directories', return_value=[roots[0]] * 3):
                self.assertEqual(len(trace_sources.available_trace_sources()), 1)

    def test_cache_is_scoped_to_the_selected_roots(self):
        first, second = str(Path('first').resolve()), str(Path('second').resolve())
        self.assertEqual(cache_path([first, second]), cache_path([second, first, first]))
        self.assertNotEqual(cache_path([first]), cache_path([second]))
        self.assertNotEqual(cache_path([first]), cache_path())

    def test_worker_waits_for_choice_then_only_discovers_selected_root(self):
        self._run_worker((('Apoteka', 'first'), ('Bureau', 'second')), ('second',))

    def test_one_source_loads_automatically(self):
        self._run_worker((('Bureau', 'second'),), ('second',))

    def test_no_sources_still_watches_configured_roots(self):
        with patch.object(indexed_ui, 'trace_directories', return_value=('first', 'second')):
            self._run_worker((), ('first', 'second'))

    def test_cancel_while_waiting_does_not_open_any_index(self):
        app = self._app()
        with patch.object(indexed_ui, 'available_trace_sources', return_value=(('A', 'one'), ('B', 'two'))), \
                patch.object(indexed_ui, 'TraceStore') as store:
            worker = threading.Thread(target=indexed_ui.IndexedAppMixin._index_loop, args=(app,))
            worker.start()
            self.assertEqual(app._index_messages.get(timeout=2)[0], 'sources')
            app._stop_index.set()
            worker.join(2)
            self.assertFalse(worker.is_alive())
            store.assert_not_called()

    def test_archive_write_failure_is_reported_and_retried(self):
        app = self._app()
        store = Mock()
        store.corrupt.is_set.return_value = False
        store.metadata.return_value = []
        attempts = []
        def sync(*args):
            attempts.append(True)
            if len(attempts) == 1:
                raise HistoryUnavailable('Nouvelles traces non archivées')
            app._stop_index.set()
            return dict(changed=0, removed=0, failures=[])
        store.sync.side_effect = sync
        with patch.object(indexed_ui, 'available_trace_sources', return_value=(('Bureau', 'second'),)), \
                patch.object(indexed_ui, 'discover_trace_files', return_value=[]), \
                patch.object(indexed_ui, 'TraceStore', return_value=store), \
                patch.object(app._stop_index, 'wait', return_value=False):
            indexed_ui.IndexedAppMixin._index_loop(app)
        self.assertEqual(len(attempts), 2)
        messages = [app._index_messages.get_nowait() for i in range(app._index_messages.qsize())]
        self.assertEqual([message[0] for message in messages],
                         ['ready', 'phase', 'phase', 'history_error', 'phase', 'complete'])
        store.rebuild.assert_not_called()

    def test_ready_is_published_before_preparation_failure_and_retry(self):
        app = self._app()
        store = Mock()
        store.corrupt.is_set.return_value = False
        store.metadata.return_value = [{'path': 'cached'}]
        store.prepare_history.side_effect = [HistoryUnavailable('archive occupée'), None]
        store.sync.side_effect = lambda *args: app._stop_index.set() or dict(changed=0, removed=0, failures=[])
        with patch.object(indexed_ui, 'available_trace_sources', return_value=(('Bureau', 'second'),)), \
                patch.object(indexed_ui, 'discover_trace_files', return_value=[]), \
                patch.object(indexed_ui, 'TraceStore', return_value=store), \
                patch.object(app._stop_index, 'wait', return_value=False):
            indexed_ui.IndexedAppMixin._index_loop(app)
        messages = [app._index_messages.get_nowait() for i in range(app._index_messages.qsize())]
        self.assertEqual([message[0] for message in messages], ['ready', 'phase', 'history_error', 'phase', 'complete'])
        self.assertEqual(messages[0][2], [{'path': 'cached'}])
        self.assertEqual(store.prepare_history.call_count, 2)
        store.rebuild.assert_not_called()

    def test_global_diagnostics_wait_for_source_verification(self):
        app = SimpleNamespace(_indexed_complete=True, _sources_verified=False)
        with patch.object(indexed_ui.messagebox, 'showinfo') as info:
            self.assertFalse(indexed_ui.IndexedAppMixin._diagnostic_ready(app))
            info.assert_called_once()
        app._sources_verified = True
        self.assertTrue(indexed_ui.IndexedAppMixin._diagnostic_ready(app))

    @staticmethod
    def _app():
        return SimpleNamespace(_index_messages=Queue(), _stop_index=threading.Event(),
                               _sources_chosen=threading.Event(), _selected_directories=None)

    def _run_worker(self, sources, selected):
        app = self._app()
        store = Mock()
        store.corrupt.is_set.return_value = False
        store.metadata.return_value = []
        store.sync.side_effect = lambda *args: app._stop_index.set() or {'changed': 0, 'failures': []}
        with patch.object(indexed_ui, 'available_trace_sources', return_value=sources), \
                patch.object(indexed_ui, 'discover_trace_files', return_value=['chosen-file']) as discover, \
                patch.object(indexed_ui, 'TraceStore', return_value=store) as constructor:
            worker = threading.Thread(target=indexed_ui.IndexedAppMixin._index_loop, args=(app,))
            worker.start()
            try:
                if len(sources) > 1:
                    self.assertEqual(app._index_messages.get(timeout=2), ('sources', sources))
                    constructor.assert_not_called()
                    app._selected_directories = selected
                    app._sources_chosen.set()
                worker.join(2)
                self.assertFalse(worker.is_alive())
                constructor.assert_called_once_with(cache_path(selected), directories=selected,
                                                    stop=app._stop_index, defer_history=True)
                discover.assert_called_once_with(directories=selected)
                store.sync.assert_called_once_with(['chosen-file'], unittest.mock.ANY, app._stop_index)
            finally:
                app._stop_index.set()
                app._sources_chosen.set()
                worker.join(2)


if __name__ == '__main__':
    unittest.main()
