"""Discovery and display names for installed and local trace files."""

import ntpath
import ctypes
import os
import sys


FIXED_DIRECTORY = r"C:/Apoteka/traces"


def application_directory():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def desktop_directory():
    """Locate the Windows Desktop, including redirected/OneDrive desktops."""
    if os.name == "nt":
        try:
            buffer = ctypes.create_unicode_buffer(32768)
            # CSIDL_DESKTOPDIRECTORY: physical user Desktop, not the shell view.
            if ctypes.windll.shell32.SHGetFolderPathW(None, 0x10, None, 0, buffer) == 0 and buffer.value:
                return buffer.value
        except (AttributeError, OSError):
            pass
    return os.path.join(os.path.expanduser("~"), "Desktop")


def trace_directories():
    return (FIXED_DIRECTORY, os.path.join(application_directory(), "traces"),
            os.path.join(desktop_directory(), "traces"))


def path_key(path):
    return os.path.normcase(os.path.realpath(path))


def discover_trace_files(prefix=None, directories=None):
    """Return distinct absolute paths; missing or inaccessible roots are skipped."""
    paths = {}
    for directory in trace_directories() if directories is None else directories:
        for root, _, files in os.walk(directory):
            for filename in files:
                name = filename.lower()
                if name.endswith(".sav"):
                    continue
                if not name.startswith(("tracomega", "tracmodbus")):
                    continue
                if prefix and not name.startswith(prefix.lower()):
                    continue
                path = os.path.abspath(os.path.join(root, filename))
                paths.setdefault(path_key(path), path)
    return sorted(paths.values(), key=path_key)


def trace_label(path):
    roots = ((os.path.join(application_directory(), "traces"), "Local"),
             (os.path.join(desktop_directory(), "traces"), "Bureau"))
    for root, origin in roots:
        try:
            relative = os.path.relpath(path, root)
            if relative != os.pardir and not relative.startswith(os.pardir + os.sep):
                robot = {"trace r1": "R1", "trace r2": "R2", "r1": "R1", "r2": "R2"}.get(
                    relative.split(os.sep)[0].lower()
                )
                source = f"{origin} {robot}" if robot else origin
                return f"{source} / {relative}"
        except ValueError:
            pass
    return f"Apoteka / {os.path.relpath(path, FIXED_DIRECTORY)}"


def same_trace_directory(first, second):
    return path_key(os.path.dirname(first)) == path_key(os.path.dirname(second))


def build_next_file_map(buffer):
    """Resolve archived original paths within each copied trace directory."""
    by_directory_and_name = {
        (path_key(os.path.dirname(path)), os.path.basename(path).lower()): path
        for path in buffer
    }
    links = {}
    for path, text in buffer.items():
        lines = text.splitlines()
        if lines and lines[0].startswith("Ancienne trace ="):
            previous = lines[0].split("=", 1)[1].strip()
            key = (path_key(os.path.dirname(path)), ntpath.basename(previous).lower())
            candidate = by_directory_and_name.get(key)
            if candidate and candidate != path:
                links[candidate] = path
    return links
