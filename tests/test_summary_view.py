import unittest
from types import SimpleNamespace
from unittest.mock import Mock

from summary_view import comparison_rows, comparison_mismatch
from indexed_ui import IndexedAppMixin
from summary_widgets import reuse_summary_analysis
from trace_identity import TraceRef


class SummaryComparisonTests(unittest.TestCase):
    def test_reuse_is_limited_to_same_occurrence_revision_and_detail_inputs(self):
        previous = dict(path='R1/archive.old', row=dict(number=17, ref='ABC'),
                        locations=[], location_lot=None, _summary_revision=4,
                        _summary_comparison=dict(revision=4, status='same'))
        for patch in ({}, {'path':'R2/archive.old'}, {'row':dict(number=18, ref='ABC')},
                      {'location_lot':'LOT'}, {'locations':[{'number':24}]}):
            with self.subTest(patch=patch):
                data = {k:v for k,v in previous.items() if not k.startswith('_')}
                data.update(patch)
                reuse_summary_analysis(previous, data, 4)
                self.assertEqual('_summary_comparison' in data, not patch)
        data = {k:v for k,v in previous.items() if not k.startswith('_')}
        reuse_summary_analysis(previous, data, 5)
        self.assertNotIn('_summary_comparison', data)
        previous['_summary_comparison'] = dict(revision=3, status='same')
        reuse_summary_analysis(previous, data, 4)
        self.assertNotIn('_summary_comparison', data)

    def test_reuse_never_crosses_generation_or_copies_an_error(self):
        previous = dict(path=TraceRef('R1/archive.old', 'first'), row=dict(number=17),
                        _summary_revision=4, _summary_comparison=dict(revision=4, status='same'))
        data = dict(path=TraceRef('R1/archive.old', 'second'), row=dict(number=17))
        reuse_summary_analysis(previous, data, 4)
        self.assertNotIn('_summary_comparison', data)
        data['path'] = TraceRef('R2/archive.old', 'first')
        reuse_summary_analysis(previous, data, 4)
        self.assertNotIn('_summary_comparison', data)
        data['path'] = previous['path']
        previous['_summary_comparison'] = dict(revision=4, error='interrupted')
        reuse_summary_analysis(previous, data, 4)
        self.assertNotIn('_summary_comparison', data)

    def test_red_values_follow_two_unit_tolerance(self):
        for delta in (-3, -2, 0, 2, 3):
            with self.subTest(delta=delta):
                pair = ({'dims': {'l': 48 + delta, 'H': 21, 'L': 47}},
                        {'dims': {'l': 48, 'H': 21, 'L': 47}})
                self.assertEqual(comparison_mismatch(pair), abs(delta) > 2)
        self.assertFalse(comparison_mismatch(({'dims': {'l': 48}}, {'dims': {'H': 21}})))

    def test_old_summary_comparison_is_discarded_after_selection_changes(self):
        app = SimpleNamespace(store=Mock(), tasks=Mock(), _view_generation=1, _selection_serial=1)
        done = Mock()
        IndexedAppMixin._submit(app, 'summary_comparison', Mock(), done)
        callback = app.tasks.submit.call_args.args[2]
        app._selection_serial = 2
        callback({'pair': 'obsolete'})
        done.assert_not_called()

    def test_dimensions_match_current_minus_previous_in_both_storage_formats(self):
        pair = ({'dims': '{"l":48,"H":21,"L":47}'}, {'dims': {'l': 55, 'H': 20, 'L': 49}})
        self.assertEqual(comparison_rows(pair), [('l', 48, 55, -7), ('H', 21, 20, 1), ('L', 47, 49, -2)])

    def test_missing_measurements_do_not_invent_zero_or_delta(self):
        self.assertEqual(comparison_rows(({'dims': {'l': 0}}, {'dims': {'H': 21}})),
                         [('l', 0, None, None), ('H', None, 21, None), ('L', None, None, None)])
