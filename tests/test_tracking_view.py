from datetime import datetime
from pathlib import Path
import unittest
from unittest.mock import patch

import trace_sources
from trace_identity import TraceRef
from tracking_view import tracking_values


class TrackingValuesTests(unittest.TestCase):
    def test_columns_preserve_fractional_time_and_exact_position(self):
        with patch.object(trace_sources, 'application_directory', return_value=str(Path('project').resolve())), \
                patch.object(trace_sources, 'desktop_directory', return_value=str(Path('desktop').resolve())):
            root = Path('desktop/traces').resolve()
            path = TraceRef(str(root / 'trace r1' / 'TracOmega019.old'), 'saved-generation',
                            datetime(2026, 10, 13).timestamp(), datetime(2026, 10, 15).timestamp())
            for suffix, expected in (('face:1 209', 'Face 1 · 209'), ('face:2 F32', 'Face 2 · F32'),
                                     ('face:1 $0101', 'Face 1 · $0101'), ('face:2', 'Face 2'),
                                     ('no position', '—')):
                row = dict(path=path, time=datetime(2026, 10, 14, 6, 34, 12).timestamp(),
                           text='G 14|06:34:12.2 GRIPPER PUTE 1x boîte(s): Code:REF ' + suffix)
                self.assertEqual(tracking_values(row, [root]),
                                 ('14/10/2026', '06:34:12.2', 'R1', expected, 'TracOmega019.old'))
                self.assertEqual(path.version_id, 'saved-generation')

    def test_multiple_sources_keep_short_provenance(self):
        project = Path('project').resolve()
        desktop = Path('desktop').resolve()
        with patch.object(trace_sources, 'application_directory', return_value=str(project)), \
                patch.object(trace_sources, 'desktop_directory', return_value=str(desktop)):
            row = dict(path=str(desktop / 'traces/R2/TracOmega.txt'), time=0,
                       text='G 01|01:02:03.0 GRIPPER PUTE face:1')
            self.assertEqual(tracking_values(row, [project / 'traces', desktop / 'traces'])[2], 'Bureau R2')
