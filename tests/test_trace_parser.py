from datetime import datetime
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock

import search_gui as gui
import trace_parser as parser
import test_trace_sources as source_tests


FIXTURES = Path(__file__).parent / "fixtures"


class TraceParserTests(unittest.TestCase):
    def test_utf8_and_windows1252_keep_accents_and_line_positions(self):
        sample = (FIXTURES / "current_r2.txt").read_text(encoding="utf-8")
        with tempfile.TemporaryDirectory() as folder:
            for encoding in ("utf-8", "utf-8-sig", "cp1252"):
                path = Path(folder) / encoding
                path.write_bytes(sample.replace("\n", "\r\n").encode(encoding))
                self.assertEqual(parser.read_trace(path), sample)
                self.assertIn("boîte(s)", parser.read_trace(path))

    def test_current_and_legacy_ticket_fields(self):
        expected = {"RefProduit": "A0061574AC", "SN": "343272", "CodeLot": "261013.127", "ModBusId": "1194"}
        samples = [
            "RefProduit=A0061574AC/SN=343272/CodeLot=261013.127/ModBusId=1194",
            "RefProduit=A0061574AC;SN=343272;CodeLot=261013.127;ModBusId=1194",
            "RefProduit = A0061574AC SN = 343272 CodeLot = 261013.127 ModBusId = 1194",
            (FIXTURES / "current_r2.txt").read_text(encoding="utf-8").splitlines()[3],
        ]
        for sample in samples:
            with self.subTest(sample=sample):
                self.assertEqual(parser.parse_fields(sample), expected)

    def test_reference_and_modbus_prefixes_are_not_matches(self):
        line = "RefProduit=ABC1234<Cr>ModBusId=11940<Cr>SN=1234<Cr>"
        self.assertFalse(parser.field_matches(line, "RefProduit", "ABC123"))
        self.assertFalse(parser.field_matches(line, "ModBusId", "1194"))
        self.assertEqual(gui.find_lines_with_keyword(line, "ModBusId=1194", True), [])
        self.assertEqual(gui.find_lines_with_keyword(line, "ModBusId=11940", True), [line])
        self.assertEqual(gui.find_lines_with_keyword("ModBusId=1194A", "ModBusId=1194", True), [])

    def test_display_fields_stop_at_cr_and_show_prefixed_timestamp(self):
        line = (FIXTURES / "current_r2.txt").read_text(encoding="utf-8").splitlines()[3]
        text, fields = parser.format_ticket_fields(line)
        self.assertTrue(text.startswith("14 | 06:27:47.8"))
        self.assertEqual(fields["refprod"], "A0061574AC")
        self.assertNotIn("<Cr>", text)
        self.assertEqual(parser.timestamp_parts("G 16|13:44:48.8 événement"), ("16", "13:44:48.8"))

    def test_gripper_current_and_historical_spelling_exact_reference_and_location(self):
        for spelling in ("boîte(s)", "boite(s)", "bote(s)"):
            line = f"G 16|13:44:48.8 GRIPPER GET 1x {spelling}: Code:ABC123 l:0 face:2 $0101"
            self.assertTrue(parser.is_gripper(line, "GET", "ABC123", True))
            self.assertFalse(parser.is_gripper(line, "GET", "ABC12", True))
            self.assertFalse(parser.is_gripper(line, "PUTE", "ABC123"))
            self.assertFalse(parser.is_gripper(line.replace("$0101", "$01010"), "GET", "ABC123", True))

    def test_idpk_accepts_accents_and_preserves_dimension_case(self):
        for prefix in ("Résultat", "Resultat", "Rsultat"):
            line = f"B 16|13:45:01.1 {prefix} de la mesure IDPK pour ABC123 : L=98 l=70 H=18"
            self.assertEqual(parser.parse_idpk(line), ("ABC123", {"L": 98, "l": 70, "H": 18}))
            self.assertFalse(parser.is_idpk_for(line, "ABC12"))
            self.assertEqual(gui.SearchApp.extract_idpk_dimensions(None, line), {"L": "98", "l": "70", "H": "18"})

    def test_measurement_after_many_diagnostics(self):
        for filename, reference in (("current_r1.txt", "A0149813AA"), ("current_r2.txt", "A0061574AC")):
            text = (FIXTURES / filename).read_text(encoding="utf-8")
            lines = text.splitlines()
            get_index, _ = gui.SearchApp.find_gripper_line(None, text, reference)
            index, line = gui.SearchApp.find_result_idpk_line(None, text, reference, get_index)
            self.assertGreater(index - get_index, 6)
            self.assertTrue(parser.is_idpk_for(line, reference))
            # The previous occurrence uses the same unbounded diagnostic window.
            self.assertEqual(gui.SearchApp.find_previous_get_and_idpk(None, text, reference, len(lines)), (index, line))
            self.assertIsNotNone(gui.SearchApp.find_gripper_pute_from(None, text, reference, get_index)[0])

    def test_measurement_does_not_belong_to_a_later_get(self):
        get = "G 16|13:44:48 GRIPPER GET 1x boîte(s): Code:ABC $0101"
        measurement = "B 16|13:45:01 Résultat de la mesure IDPK pour ABC : l=1 H=2 L=3"
        self.assertEqual(parser.find_idpk_after([get, get, measurement], "ABC", 0), (None, None))

    def test_robot_error_detection_uses_exact_reference_and_accepts_current_get(self):
        text = (FIXTURES / "current_r2.txt").read_text(encoding="utf-8")
        self.assertEqual(parser.missing_robot_tickets(text.splitlines()), [])
        lines = ["ModBusId=12<Cr>RefProduit=ABC<Cr>", "ModBusId=13<Cr>RefProduit=ABC123<Cr>",
                 "GRIPPER GET 1x boîte(s): Code:ABC123 $0101"]
        self.assertEqual(parser.missing_robot_tickets(lines), [("12", "ABC", 0)])
        self.assertEqual(parser.missing_robot_tickets(list(reversed(lines))), [("13", "ABC123", 1), ("12", "ABC", 2)])

    def test_trace_dates_take_precedence_over_copy_dates(self):
        text = (FIXTURES / "current_r2.txt").read_text(encoding="utf-8")
        fallback = datetime(2030, 1, 1)
        self.assertEqual(parser.trace_start_datetime(text, fallback), datetime(2026, 10, 14))
        self.assertEqual(parser.trace_end_datetime(text, fallback), datetime(2026, 10, 14, 6, 28, 7, 600000))
        self.assertEqual(parser.trace_start_datetime("no header", fallback), fallback)

    def test_month_and_year_rollover(self):
        text = "creation de la trace à 23:59:58 Le 31/12/2025\nG 31|23:59:59 événement\nG 01|00:00:01 événement"
        dates = [dt for dt, line in parser.iter_timed_lines(text, datetime(2030, 1, 1))]
        self.assertEqual(dates[-1], datetime(2026, 1, 1, 0, 0, 1))
        self.assertEqual(parser.trace_end_datetime(text, datetime(2030, 1, 1)), dates[-1])

    def test_modbus_different_row_lengths_do_not_shift_registers(self):
        for width in (120, 121):
            row = ["0"] * width
            row[0], row[6], row[7], row[114], row[115], row[-1] = "16|14:57:15", "821", "1194", "1", "0", ""
            values = parser.parse_modbus_line(";".join(row) + "\r\n")
            self.assertEqual(len(values), width)
            self.assertEqual((values[6], values[7], values[114], values[115]), ("821", "1194", "1", "0"))


class CurrentTraceIntegrationTests(source_tests.TraceFixture):
    def test_real_format_indexes_tracking_and_archive_order(self):
        for directory, fixture in ((self.r1, "current_r1.txt"), (self.r2, "current_r2.txt")):
            text = (FIXTURES / fixture).read_text(encoding="utf-8")
            self.write(directory, text=text)
            (directory / "TracOMEGA.txt").write_bytes(text.encode("cp1252"))
        self.model._refresh_traces()
        self.assertEqual(set(self.model.ref_index), {"A0149813AA", "A0061574AC"})
        self.assertEqual(set(self.model.sn_index), {"272106", "343272"})
        self.assertEqual(set(self.model.codelot_index), {"251114.019", "261013.127"})
        self.assertEqual(self.model.idpk_index["A0061574AC"][0][1], {"l": 48, "H": 21, "L": 47})
        self.assertEqual(self.model.file_list[0], str(self.r1 / "TracOMEGA.txt"))
        self.model.tracking_keyword = Mock()
        self.model.tracking_keyword.get.return_value = "A0061574AC"
        self.model.tracking_txt = Mock()
        gui.SearchApp.perform_tracking_search(self.model)
        result = self.model.tracking_txt.insert.call_args[0][1]
        self.assertIn("14/10/26 | 06:28:07.6", result)
        self.assertIn("Local R2", result)


if __name__ == "__main__":
    unittest.main()
