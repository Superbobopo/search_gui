import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch

import search_gui as gui
import trace_sources as sources


class TraceModel:
    _refresh_traces = gui.SearchApp._refresh_traces
    _rebuild_trace_indexes = gui.SearchApp._rebuild_trace_indexes
    _build_idpk_index = gui.SearchApp._build_idpk_index
    get_convoyage_steps = gui.SearchApp.get_convoyage_steps


class TraceFixture(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.base = Path(temporary.name)
        self.installed = self.base / "Apoteka" / "traces"
        self.r1 = self.base / "traces" / "trace r1"
        self.r2 = self.base / "traces" / "trace r2"
        self.application_patch = patch.object(sources, "application_directory", return_value=str(self.base))
        self.application_patch.start()
        self.addCleanup(self.application_patch.stop)
        directory_patch = patch.object(sources, "FIXED_DIRECTORY", str(self.installed))
        directory_patch.start()
        self.addCleanup(directory_patch.stop)
        gui.BUFFER = {}
        gui.BUFFER_MTIME = {}
        gui.current_file = None
        self.addCleanup(self.clear_buffers)
        self.model = TraceModel()

    @staticmethod
    def clear_buffers():
        gui.BUFFER = {}
        gui.BUFFER_MTIME = {}
        gui.current_file = None

    def write(self, directory, name="TracOMEGA.txt", text="RefProduit=COMMON SN=123 CodeLot=LOT\n"):
        directory.mkdir(parents=True, exist_ok=True)
        path = directory / name
        path.write_text(text, encoding="utf-8")
        return str(path)


class TraceSourcesTests(TraceFixture):
    def test_all_sources_keep_homonyms_and_search_indexes(self):
        paths = [self.write(folder) for folder in (self.installed, self.r1, self.r2)]
        self.model._refresh_traces()
        self.assertEqual(set(gui.BUFFER), set(paths))
        self.assertEqual({p for p, _ in self.model.ref_index["COMMON"]}, set(paths))
        self.assertEqual(len(self.model.sn_index["123"]), 3)
        self.assertEqual(len(self.model.codelot_index["LOT"]), 3)
        for path in paths:
            self.assertTrue(gui.search_file_for_keyword(path, "COMMON"))
        self.assertTrue(sources.trace_label(paths[0]).startswith("Apoteka /"))
        self.assertTrue(sources.trace_label(paths[1]).startswith("Local R1 /"))
        self.assertTrue(sources.trace_label(paths[2]).startswith("Local R2 /"))

    def test_missing_installed_directory_and_filters(self):
        paths = {self.write(self.r1), self.write(self.r2, "TracMODBUS.txt")}
        self.write(self.r1, "TracOmega.sav")
        self.write(self.r2, "TracMODBUS.SAV")
        self.write(self.r2, "other.txt")
        archive = self.write(self.r1, "TracOmega001.old")
        paths.add(archive)
        self.model._refresh_traces()
        self.assertEqual(set(gui.BUFFER), paths)
        self.assertEqual(sources.discover_trace_files("tracmodbus"), [str(self.r2 / "TracMODBUS.txt")])

    def test_empty_sources_and_open_without_current_file(self):
        self.model._refresh_traces()
        self.assertEqual(self.model.file_list, [])
        self.assertEqual(self.model.next_file_map, {})
        self.assertEqual(self.model.get_convoyage_steps("123"), [])
        self.model.keyword = Mock()
        self.model.keyword.get.return_value = "123"
        with patch.object(gui.messagebox, "showinfo") as info:
            gui.SearchApp.open_file_view(self.model)
        info.assert_called_once()

    def test_refresh_add_modify_remove_and_clear_stale_selection(self):
        first = self.write(self.r1)
        self.model._refresh_traces()
        gui.current_file = first
        self.model.lines = ["stale result"]
        self.write(self.r1, text="RefProduit=CHANGED SN=456 CodeLot=NEW\nIDPK pour ITEM: l=1 H=2 L=3\n")
        stat = os.stat(first)
        os.utime(first, ns=(stat.st_atime_ns, stat.st_mtime_ns + 1_000_000_000))
        added = self.write(self.r2)
        self.model._refresh_traces()
        self.assertEqual(self.model.ref_index["CHANGED"][0][0], first)
        self.assertEqual([p for p, _ in self.model.ref_index["COMMON"]], [added])
        self.assertTrue(self.model.idpk_index["ITEM"])
        Path(first).unlink()
        self.model.advanced_candidates = [(first, ["old"])]
        self.model._refresh_traces()
        self.assertNotIn(first, gui.BUFFER)
        self.assertNotIn(first, gui.BUFFER_MTIME)
        self.assertEqual(self.model.ref_index["CHANGED"], [])
        self.assertEqual(self.model.idpk_index, {})
        self.assertEqual(self.model.advanced_candidates, [])
        self.assertIsNone(gui.current_file)
        self.assertEqual(self.model.lines, [])

    def test_unchanged_refresh_reuses_indexes(self):
        self.write(self.r1)
        self.model._refresh_traces()
        index = self.model.ref_index
        self.model._refresh_traces()
        self.assertIs(self.model.ref_index, index)

    def test_rotated_links_stay_in_the_same_robot_directory(self):
        old1 = self.write(self.r1, "TracOmega001.old")
        old2 = self.write(self.r2, "TracOmega001.old")
        header = "Ancienne trace = C:\\Apoteka\\traces\\TracOmega001.old\n"
        new1 = self.write(self.r1, text=header)
        new2 = self.write(self.r2, text=header)
        self.model._refresh_traces()
        self.assertEqual(self.model.next_file_map, {old1: new1, old2: new2})
        self.assertFalse(sources.same_trace_directory(old1, new2))

    def test_convoyage_detail_uses_full_path(self):
        first = self.write(self.r1, "TracMODBUS.txt", "R1;0;123\n")
        second = self.write(self.r2, "TracMODBUS.txt", "R2;0;123\n")
        all_steps = self.model.get_convoyage_steps("123")
        self.assertEqual(len(all_steps), 2)
        for path, expected in ((first, "R1"), (second, "R2")):
            steps = self.model.get_convoyage_steps("123", only_filepath=path)
            self.assertEqual(len(steps), 1)
            self.assertEqual(steps[0]["horodatage"], expected)
            self.assertEqual(steps[0]["file"], sources.trace_label(path))

    def test_robot_error_click_opens_exact_homonym(self):
        first = self.write(self.r1)
        second = self.write(self.r2)
        self.model._refresh_traces()
        for attribute in ("reference", "modbus", "btn_up", "btn_down", "keyword",
                          "btn_compare_idpk", "btn_adjacent", "focus_on_ticket_and_get",
                          "show_adjacent_modbusids_window"):
            setattr(self.model, attribute, Mock())
        gui.SearchApp._on_error_line_click(self.model, second, "123", "COMMON")
        self.assertEqual(gui.current_file, second)
        self.model.focus_on_ticket_and_get.assert_called_once_with(second, "123", "COMMON")
        self.model.show_adjacent_modbusids_window.assert_called_once_with("123", second)
        self.assertNotEqual(first, second)

    def test_current_working_directory_does_not_change_sources(self):
        path = self.write(self.r1)
        initial_directory = os.getcwd()
        try:
            os.chdir(self.base.parent)
            self.assertEqual(sources.discover_trace_files(), [path])
        finally:
            os.chdir(initial_directory)

    def test_script_and_executable_base_directory(self):
        self.application_patch.stop()
        self.assertEqual(sources.application_directory(), str(Path(sources.__file__).resolve().parent))
        with patch.object(sys, "frozen", True, create=True), patch.object(
            sys, "executable", str(self.base / "dist" / "search_gui.exe")
        ):
            self.assertEqual(sources.application_directory(), str(self.base / "dist"))
            self.assertEqual(sources.trace_directories()[1], str(self.base / "dist" / "traces"))

    def test_overlapping_roots_do_not_load_a_file_twice(self):
        path = self.write(self.r1)
        self.assertEqual(sources.discover_trace_files(directories=[self.base / "traces", self.r1]), [path])


if __name__ == "__main__":
    unittest.main()
