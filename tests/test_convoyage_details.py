"""Combined convoyage reads preserve legacy transformations and snapshots."""
from concurrent.futures import ThreadPoolExecutor
from unittest.mock import patch
import unittest

from tests import test_trace_store as fixture
import trace_store


def line(stamp, values):
    columns = ['0'] * 121
    columns[0] = stamp
    for col, value in values.items():
        columns[col] = value
    return ';'.join(columns)


class ConvoyageDetailsTests(unittest.TestCase):
    setUp = fixture.StoreTests.setUp
    write = fixture.StoreTests.write

    def test_modbus_old_only_and_different_current_paths_ignore_r2(self):
        from trace_sources import discover_trace_files
        old = self.write('traces/R1', 'TracMODBUS.old', line('16|12:00:00', {2:'903'}))
        ignored = self.write('traces/R2', 'TracMODBUS.old', line('16|12:00:00', {2:'999'}))
        paths = discover_trace_files(directories=[str(self.root/'traces')])
        self.assertEqual(paths, [old])
        self.assertNotIn(ignored, paths)
        self.store.sync(paths)
        summary, steps = self.store.convoyage_details('903', old)
        self.assertTrue(summary and steps)
        current = self.write('traces/R1', 'TracMODBUS.txt', line('16|13:00:00', {31:'904'}))
        self.store.sync(discover_trace_files(directories=[str(self.root/'traces')]))
        self.assertEqual(self.store.convoyage_details('903', old), (summary, steps))
        self.assertEqual(self.store.convoyage_details('903', current), ([], []))
        self.assertTrue(self.store.convoyage_details('904', current)[1])

    def test_combined_outputs_exact_paths_order_and_single_decode(self):
        text = '\n'.join((line('16|12:00:00', {2:'1054'}),
                          line('16|12:00:01', {31:'1054'}),
                          line('16|12:00:02', {114:'-1054'}),
                          line('16|12:00:03', {1:'1054'})))
        first = self.write('Local/R1', 'TracMODBUS.txt', text)
        other = self.write('Bureau/R1', 'TracMODBUS.txt', line('16|13:00:00', {6:'1054'}))
        self.store.sync([first, other])
        for path in (first, other, None):
            expected = self.store.convoyage_summary('1054', path), self.store.convoyage_steps('1054', path)
            count = len(self.store.modbus_rows('1054', path))
            with patch.object(self.store, 'modbus_rows', wraps=self.store.modbus_rows) as rows, \
                 patch.object(trace_store, 'parse_modbus_line', wraps=trace_store.parse_modbus_line) as parse:
                result = self.store.convoyage_details('1054', path)
            self.assertEqual(result, expected)
            self.assertEqual(rows.call_count, 1)
            self.assertEqual(parse.call_count, count)
            if path:
                self.assertTrue(all(item['_filepath']==path for item in result[0]))
        summary, steps = self.store.convoyage_details('1054', first)
        self.assertEqual(summary[0]['HORODATAGE'], '16|12:00:02')
        self.assertEqual(summary[0]['_last_positive']['horodatage'], '16|12:00:01')
        self.assertEqual([item['horodatage'] for item in steps], ['16|12:00:00','16|12:00:01','16|12:00:02'])
        self.assertEqual([item['etat'] for item in steps], ['Présent','Présent','Jamais arrivé'])
        self.assertEqual(self.store.convoyage_details('999', first), ([], []))

    def test_publication_between_transformations_keeps_same_snapshot(self):
        path = self.write(filename='TracMODBUS.txt', text=line('16|12:00:00', {2:'1054'}))
        self.store.sync([path])
        expected = self.store.convoyage_details('1054', path)
        transform = self.store._convoyage_summary
        def publish(reference, parsed):
            self.write(filename='TracMODBUS.txt', text=line('16|13:00:00', {31:'1054'}))
            with ThreadPoolExecutor(1) as pool:
                pool.submit(self.store.sync, [path]).result(timeout=10)
            return transform(reference, parsed)
        with patch.object(self.store, '_convoyage_summary', side_effect=publish):
            result = self.store.convoyage_details('1054', path)
        self.assertEqual(result, expected)
        self.assertNotEqual(self.store.convoyage_details('1054', path), expected)
