"""Targeted regressions for source selection; no graphical window is opened."""

from pathlib import Path
from queue import Queue
from types import SimpleNamespace
import tempfile
import threading
import unittest
from unittest.mock import Mock, patch

import indexed_ui
import trace_sources
from trace_store import cache_path


class SourceChoiceTests(unittest.TestCase):
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

    @staticmethod
    def _app():
        return SimpleNamespace(_index_messages=Queue(), _stop_index=threading.Event(),
                               _sources_chosen=threading.Event(), _selected_directories=None)

    def _run_worker(self, sources, selected):
        app = self._app()
        store = Mock()
        store.corrupt.is_set.return_value = False
        store.metadata.return_value = []
        store.sync.side_effect = lambda *args: app._stop_index.set() or {'changed': 0}
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
                constructor.assert_called_once_with(cache_path(selected))
                discover.assert_called_once_with(directories=selected)
                store.sync.assert_called_once_with(['chosen-file'], unittest.mock.ANY, app._stop_index)
            finally:
                app._stop_index.set()
                app._sources_chosen.set()
                worker.join(2)


if __name__ == '__main__':
    unittest.main()
