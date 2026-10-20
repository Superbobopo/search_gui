"""Text sharing keeps complete published occurrences and explicit missing matches."""
import unittest
from adjacent_view import listing_text


class ListingTests(unittest.TestCase):
    def row(self, number, **changes):
        row = dict(path='C:/traces/R1/TracOmega.txt', number=number,
                   text='10:12:13 GET', time=1792404733, modbus=str(number),
                   ref='REFERENCE_COMPLETE', sn='SN_COMPLETE', lot='LOT_COMPLETE')
        row.update(changes)
        return row

    def test_all_rows_duplicates_full_values_and_order(self):
        anchor = self.row(40)
        previous = [self.row(i) for i in range(40)]
        following = [self.row(41, sn=None, lot=None)]
        data = dict(columns=[dict(robot='R2', message='Pas de date', anchor=None),
                             dict(robot='R1', anchor=anchor, previous=previous,
                                  following=following)])
        text = listing_text(data, [])
        self.assertTrue(text.startswith('R1\n'))
        self.assertEqual(text.count('Réf : REFERENCE_COMPLETE'), 42)
        self.assertEqual(text.count('🟢 ➜ '), 1)
        self.assertNotIn('Repère', text)
        self.assertNotIn('TracOmega', text)
        self.assertNotIn('C:/traces', text)
        self.assertIn('SN : — | Lot : —', text)
        self.assertIn('SN : SN_COMPLETE | Lot : LOT_COMPLETE', text)
        self.assertLess(text.index('ID : 0 |'), text.index('ID : 40 |'))
        self.assertIn('\n\nR2\nPas de date', text)

    def test_ambiguous_candidates_are_not_exported_as_neighbors(self):
        text = listing_text(dict(columns=[dict(robot='R2', anchor=None,
                                                candidates=[self.row(1)])]), [])
        self.assertIn('ambigu', text)
        self.assertNotIn('REFERENCE_COMPLETE', text)

    def test_empty_listing(self):
        self.assertEqual(listing_text(dict(columns=[]), []), '')
