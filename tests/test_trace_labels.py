"""Compact presentation must not replace saved trace identities."""

from datetime import datetime
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import trace_sources as sources
from trace_identity import TraceRef


class TraceLabelTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.base = Path(temporary.name)
        for name, value in (('application_directory', str(self.base / 'app')),
                            ('desktop_directory', str(self.base / 'desktop'))):
            mocked = patch.object(sources, name, return_value=value)
            mocked.start()
            self.addCleanup(mocked.stop)
        mocked = patch.object(sources, 'FIXED_DIRECTORY', str(self.base / 'apoteka'))
        mocked.start()
        self.addCleanup(mocked.stop)
        self.desktop = self.base / 'desktop' / 'traces'
        self.local = self.base / 'app' / 'traces'

    def test_robot_aliases_and_exact_period(self):
        for folder, robot in (('trace r1', 'R1'), ('R1', 'R1'), ('trace r2', 'R2'), ('r2', 'R2')):
            with self.subTest(folder=folder):
                ref = TraceRef(str(self.desktop / folder / 'TracOmega019.old'), 'generation123',
                               datetime(2026, 10, 13).timestamp(), datetime(2026, 10, 15).timestamp())
                self.assertEqual(sources.compact_trace_label(ref, [self.desktop]),
                                 f'{robot} \\TracOmega019.old — 13/10/2026 – 15/10/2026')
                self.assertEqual(ref.version_id, 'generation123')

    def test_multiple_sources_are_distinguished_and_duplicate_roots_count_once(self):
        path = self.desktop / 'R1' / 'TracOmega.txt'
        self.assertEqual(sources.compact_trace_label(path, [self.desktop, self.local]),
                         'Bureau R1 \\TracOmega.txt')
        self.assertEqual(sources.compact_trace_label(self.local / 'R1' / path.name,
                                                    [self.desktop, self.local]),
                         'Local R1 \\TracOmega.txt')
        self.assertEqual(sources.compact_trace_label(path, [self.desktop, self.desktop]),
                         'R1 \\TracOmega.txt')

    def test_unknown_robot_and_missing_dates_have_no_invented_values(self):
        ref = TraceRef(str(self.desktop / 'other' / 'TracOmega.txt'), 'unknown')
        self.assertEqual(sources.compact_trace_label(ref, [self.desktop]), 'TracOmega.txt')
        self.assertEqual(sources.compact_trace_label(ref, [self.desktop, self.local]),
                         'Bureau \\TracOmega.txt')
        self.assertEqual(sources.compact_trace_label(self.base / 'elsewhere' / 'TracMODBUS.txt'),
                         'TracMODBUS.txt')

    def test_tooltip_retains_original_path_full_generation_and_history_status(self):
        ref = TraceRef(str(self.desktop / 'R2' / 'TracOmega.txt'), 'full-generation-identifier',
                       datetime(2026, 10, 13).timestamp(), datetime(2026, 10, 15).timestamp(), True)
        detail = sources.trace_tooltip(ref)
        for expected in ('Bureau R2', ref.path, ref.version_id, 'Historique', '13/10/2026'):
            self.assertIn(expected, detail)
        self.assertNotIn('Historique', sources.compact_trace_label(ref))
        self.assertIn('OMEGA conservée', sources.trace_tooltip(TraceRef(ref.path, ref.version_id)))

    def test_shared_modbus_does_not_inherit_robot_from_its_folder(self):
        for root, origin in ((self.desktop, 'Bureau'), (self.local, 'Local')):
            for robot in ('trace r1', 'R2'):
                path = root / robot / 'TracMODBUS.txt'
                self.assertEqual(sources.trace_badge(path), 'R1/R2')
                self.assertEqual(sources.trace_label(path), f'{origin} / TracMODBUS.txt')
                self.assertEqual(sources.compact_trace_label(path, [root]), 'TracMODBUS.txt')
                self.assertEqual(sources.compact_trace_label(path, [self.desktop, self.local]), f'{origin} \\TracMODBUS.txt')
                self.assertIn(str(path), sources.trace_tooltip(path))


if __name__ == '__main__':
    unittest.main()
