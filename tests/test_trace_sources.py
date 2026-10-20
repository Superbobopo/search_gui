import os
import ctypes
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch

import search_gui as gui
import trace_sources as sources
from trace_store import TraceStore


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
        self.desktop = self.base / "OneDrive" / "Bureau"
        self.desktop_patch = patch.object(sources, "desktop_directory", return_value=str(self.desktop))
        self.desktop_patch.start()
        self.addCleanup(self.desktop_patch.stop)
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
    def test_r2_modbus_is_ignored_in_all_roots_and_nested_folders(self):
        expected = set()
        for root in (self.installed, self.base / 'traces', self.desktop / 'traces'):
            expected.add(self.write(root, 'TracMODBUS.txt'))
            for name in ('R1', 'trace r1', 'R2', 'trace r2'):
                folder = root / name / 'nested'
                expected.add(self.write(folder, 'TracOmega.txt'))
                for filename in ('TracMODBUS.txt', 'tracmodbus001.old'):
                    path = self.write(folder, filename)
                    if name.lower() in ('r1', 'trace r1'):
                        expected.add(path)
        self.assertEqual(set(sources.discover_trace_files()), expected)
        self.assertEqual(set(sources.discover_trace_files('tracmodbus')),
                         {p for p in expected if Path(p).name.lower().startswith('tracmodbus')})
        self.assertEqual(sources.discover_trace_files('tracmodbus', directories=[self.r2]), [])

    def test_cached_r2_modbus_removed_before_active_store_is_published(self):
        first = self.write(self.r1, 'TracMODBUS.txt', 'R1_SHARED_MARKER\n')
        second = self.write(self.r2, 'TracMODBUS.txt', 'R2_IGNORED_MARKER\n')
        omega = self.write(self.r2, text='OMEGA_R2_KEPT\n')
        cache = self.base / 'cache.sqlite3'
        legacy = TraceStore(cache)
        legacy.sync([first, second, omega])
        self.assertTrue(legacy.search('R2_IGNORED_MARKER'))
        active = TraceStore(cache, directories=[str(self.base / 'traces')], defer_history=True)
        self.assertFalse(active.search('R2_IGNORED_MARKER'))
        self.assertTrue(active.search('R1_SHARED_MARKER'))
        self.assertTrue(active.search('OMEGA_R2_KEPT'))
        self.assertTrue(Path(second).exists())

    def test_desktop_sources_keep_homonyms_labels_and_archive_links(self):
        folders = (self.installed, self.r1, self.r2,
                   self.desktop / "traces" / "R1", self.desktop / "traces" / "R2")
        paths = [self.write(folder) for folder in folders]
        archives = [self.write(folder, "TracOmega001.old") for folder in folders]
        header = "Ancienne trace = C:\\Apoteka\\traces\\TracOmega001.old\n"
        for folder in folders:
            self.write(folder, text=header + "RefProduit=COMMON<Cr>ModBusId=123<Cr>\n")
        self.write(folders[-1], "TracOmega.sav")
        self.assertEqual(set(sources.discover_trace_files()), set(paths + archives))
        self.assertTrue(sources.trace_label(paths[-2]).startswith("Bureau R1 /"))
        self.assertTrue(sources.trace_label(paths[-1]).startswith("Bureau R2 /"))
        store = TraceStore(self.base / "index.sqlite3")
        store.sync(sources.discover_trace_files())
        self.assertEqual({path for path, hits in store.search("COMMON")}, set(paths + archives))
        self.assertEqual(store.next_files(), dict(zip(archives, paths)))
        texts = {path: Path(path).read_text(encoding="utf-8") for path in paths + archives}
        self.assertEqual(sources.build_next_file_map(texts), dict(zip(archives, paths)))

    def test_desktop_refresh_add_remove_modbus_and_overlapping_roots(self):
        root = self.desktop / "traces"
        self.assertEqual(sources.discover_trace_files(), [])
        path = self.write(root / "R1")
        modbus = self.write(root / "R1", "TracMODBUS.txt")
        self.write(root / "R2", "TracMODBUS.SAV")
        self.assertEqual(set(sources.discover_trace_files()), {path, modbus})
        self.assertEqual(sources.discover_trace_files("tracmodbus"), [modbus])
        with patch.object(sources, "application_directory", return_value=str(self.desktop)):
            self.assertEqual(set(sources.discover_trace_files()), {path, modbus})
        Path(path).unlink()
        self.assertEqual(sources.discover_trace_files(), [modbus])

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
        paths = {self.write(self.r1), self.write(self.r1, "TracMODBUS.txt")}
        self.write(self.r1, "TracOmega.sav")
        self.write(self.r2, "TracMODBUS.SAV")
        self.write(self.r2, "other.txt")
        archive = self.write(self.r1, "TracOmega001.old")
        paths.add(archive)
        self.model._refresh_traces()
        self.assertEqual(set(gui.BUFFER), paths)
        self.assertEqual(sources.discover_trace_files("tracmodbus"), [str(self.r1 / "TracMODBUS.txt")])

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
        second = self.write(self.desktop / 'traces' / 'R1', "TracMODBUS.txt", "R2;0;123\n")
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


class DesktopDirectoryTests(unittest.TestCase):
    def test_windows_redirected_desktop(self):
        redirected = os.path.join("C:" + os.sep, "Users", "Example", "OneDrive", "Bureau")
        def locate(window, folder, token, flags, buffer):
            self.assertEqual(folder, 0x10)
            buffer.value = redirected
            return 0
        shell = Mock()
        shell.shell32.SHGetFolderPathW.side_effect = locate
        with patch.object(sources.os, "name", "nt"), patch.object(ctypes, "windll", shell, create=True):
            self.assertEqual(sources.desktop_directory(), redirected)

    def test_windows_lookup_failure_falls_back_to_profile(self):
        shell = Mock()
        for failure in (1, OSError("Unavailable")):
            with self.subTest(failure=failure):
                shell.shell32.SHGetFolderPathW.reset_mock(side_effect=True)
                if isinstance(failure, OSError):
                    shell.shell32.SHGetFolderPathW.side_effect = failure
                else:
                    shell.shell32.SHGetFolderPathW.return_value = failure
                with patch.object(sources.os, "name", "nt"), patch.object(ctypes, "windll", shell, create=True):
                    self.assertEqual(sources.desktop_directory(), os.path.join(os.path.expanduser("~"), "Desktop"))


if __name__ == "__main__":
    unittest.main()
