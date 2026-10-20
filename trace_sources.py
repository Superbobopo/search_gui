"""Discovery and display names for installed and local trace files."""

import ntpath
import ctypes
import hashlib
import json
import os
import sys


FIXED_DIRECTORY = r"C:/Apoteka/traces"


def machine_identifier():
    """Bind local preferences to Windows itself, never to a copied project."""
    if os.name != 'nt':
        return None
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r'SOFTWARE\Microsoft\Cryptography') as key:
            value = winreg.QueryValueEx(key, 'MachineGuid')[0]
        return hashlib.sha256(str(value).encode('utf8')).hexdigest() if value else None
    except OSError:
        return None


def source_preferences_path():
    base = os.environ.get('LOCALAPPDATA')
    return os.path.join(base, 'TracOmega', 'source_preferences.json') if base else None


def preferred_trace_directories():
    """None retains normal source choice; a matching local policy selects Desktop."""
    filename = source_preferences_path()
    if not filename:
        return None
    try:
        with open(filename, encoding='utf8') as stream:
            settings = json.load(stream)
        machine = machine_identifier()
        if (isinstance(settings, dict) and settings.get('version') == 1 and machine
                and settings.get('machine_id') == machine
                and settings.get('startup_source') == 'desktop'):
            return (os.path.join(desktop_directory(), 'traces'),)
    except (OSError, ValueError):
        pass
    return None


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


def robot_scope(path):
    """Return the source above R1/R2 and the robot, without using display labels."""
    directory = os.path.dirname(os.path.abspath(os.fspath(path)))
    robots = {'r1': 'R1', 'trace r1': 'R1', 'r2': 'R2', 'trace r2': 'R2'}
    while directory:
        robot = robots.get(os.path.basename(directory).lower())
        parent = os.path.dirname(directory)
        if robot:
            return path_key(parent), robot
        if parent == directory:
            break
        directory = parent
    return path_key(os.path.dirname(os.fspath(path))), None


def available_trace_sources():
    """List nonempty roots once, preserving the source and full root path."""
    found = []
    seen = set()
    from omega_history import historical_roots
    archived = historical_roots()
    for name, directory in zip(("Apoteka", "À côté du programme", "Bureau"), trace_directories()):
        key = path_key(directory)
        live = discover_trace_files(directories=[directory])
        if key not in seen and (live or key in archived):
            found.append((name if live else name + ' (historique)', os.path.abspath(directory)))
            seen.add(key)
    return tuple(found)


def ignored_r2_modbus(path):
    """R1 supplies the shared MODBUS trace; never ingest the R2 copy."""
    return (ntpath.basename(os.fspath(path)).lower().startswith('tracmodbus')
            and robot_scope(path)[1] == 'R2')


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
                if ignored_r2_modbus(path):
                    continue
                paths.setdefault(path_key(path), path)
    return sorted(paths.values(), key=path_key)


def is_modbus_trace(path):
    return ntpath.basename(os.fspath(path)).lower().startswith('tracmodbus')


def trace_badge(path):
    """MODBUS is shared by both robots, regardless of its storage folder."""
    return 'R1/R2' if is_modbus_trace(path) else robot_scope(path)[1] or 'Trace'


def trace_label(path):
    from trace_identity import TraceRef
    if isinstance(path, TraceRef):
        from datetime import datetime
        period = f'{datetime.fromtimestamp(path.start):%d/%m/%Y} – {datetime.fromtimestamp(path.end):%d/%m/%Y}'
        suffix = 'Historique' if path.historical else 'OMEGA conservée'
        return f'{trace_label(path.path)} — {suffix} · {period} · {path.version_id[:8]}'
    roots = ((os.path.join(application_directory(), "traces"), "Local"),
             (os.path.join(desktop_directory(), "traces"), "Bureau"))
    for root, origin in roots:
        try:
            relative = os.path.relpath(path, root)
            if relative != os.pardir and not relative.startswith(os.pardir + os.sep):
                robot = {"trace r1": "R1", "trace r2": "R2", "r1": "R1", "r2": "R2"}.get(
                    relative.split(os.sep)[0].lower()
                )
                if is_modbus_trace(path):
                    return f'{origin} / {ntpath.basename(os.fspath(path))}'
                source = f"{origin} {robot}" if robot else origin
                return f"{source} / {relative}"
        except ValueError:
            pass
    return f"Apoteka / {ntpath.basename(os.fspath(path)) if is_modbus_trace(path) else os.path.relpath(path, FIXED_DIRECTORY)}"


def compact_trace_label(path, directories=None):
    """Presentation only: never use this label to resolve a trace identity."""
    from datetime import datetime
    from trace_identity import TraceRef
    original = os.fspath(path)
    origin, relative = '', original
    for root, name in ((os.path.join(application_directory(), 'traces'), 'Local'),
                       (os.path.join(desktop_directory(), 'traces'), 'Bureau'),
                       (FIXED_DIRECTORY, 'Apoteka')):
        try:
            candidate = os.path.relpath(original, root)
        except ValueError:
            continue
        if candidate != os.pardir and not candidate.startswith(os.pardir + os.sep):
            origin, relative = name, candidate
            break
    parts = relative.replace('\\', '/').split('/')
    robots = {'r1': 'R1', 'trace r1': 'R1', 'r2': 'R2', 'trace r2': 'R2'}
    robot = next((robots[part.lower()] for part in parts[:-1] if part.lower() in robots), '')
    if is_modbus_trace(path):
        robot = ''
    multiple = len({path_key(root) for root in directories or ()}) > 1
    prefix = ' '.join(part for part in (origin if multiple else '', robot) if part)
    label = f'{prefix} \\{parts[-1]}' if prefix else parts[-1]
    if isinstance(path, TraceRef) and path.start and path.end:
        label += f' — {datetime.fromtimestamp(path.start):%d/%m/%Y} – {datetime.fromtimestamp(path.end):%d/%m/%Y}'
    return label


def trace_tooltip(path):
    from trace_identity import TraceRef
    detail = f'{trace_label(path)}\n{os.fspath(path)}'
    if isinstance(path, TraceRef):
        detail += f'\nGénération : {path.version_id}'
    return detail


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
