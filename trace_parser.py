"""Read robot traces without changing their original text or line positions."""

from datetime import datetime
import re

GRIPPER_COMMAND_RE = re.compile(r'\bGRIPPER\s+([A-Z]+)\b')
GRIPPER_REPLY_RE = re.compile(r"\bTgripper\b.*?\brep\s*=\s*['\"]([^'\"]*)['\"]", re.I)
GET_ERROR_RE = re.compile(r'\bErr\s*:\s*r[eé]ponse\s+gripper\s+GET\s*:\s*(.+)', re.I)
GET_BOX_RE = re.compile(r'\bChargeBoitesDansGripperMB\b.*?\b1\s+bo[iî]tes?(?:\(s\))?\s+\(IDb:(\d+)\(([^()\s]+)\)\)', re.I)


def get_box_context(text):
    """Single-box loading context provides an explicit internal box identity."""
    match = GET_BOX_RE.search(text)
    return (match[2], match[1]) if match else None


def gripper_reply(text):
    """An explicit GET error or the reply to the active gripper command."""
    error = GET_ERROR_RE.search(text)
    match = error or GRIPPER_REPLY_RE.search(text)
    if not match:
        return None
    value = match[1].strip()
    failed = bool(error or re.search(r'\b(?:err|error|nok|erreur)\b', value, re.I))
    status = 'failed' if failed else 'success' if value.casefold() == 'ok' else 'unknown'
    code = re.search(r'\((-?\d+)\)', value)
    return dict(status=status, message=value, code=code[1] if code else None,
                explicit_get=bool(error))


FIELD_RE = re.compile(r"\b(RefProduit|SN|CodeLot|ModBusId)\s*=\s*([^<>\s/;]+)")
STAMP_RE = re.compile(r"^\s*(?:[^\d\s]\s+)?(\d{1,2})\|(\d{2}:\d{2}:\d{2}(?:\.\d+)?)")
GRIPPER_RE = re.compile(r"\bGRIPPER\s+(GET|PUTE)\s+1x\s+[^:\s]+:\s*Code:\s*([^<>\s;]+)")
IDPK_RE = re.compile(r"\bIDPK\s+pour\s+([^\s:<>]+)\s*:\s*(.*)")
DIMENSION_RE = re.compile(r"\b(l|h|H|L)\s*=\s*(\d+)")


def read_trace(path):
    with open(path, "rb") as stream:
        data = stream.read()
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("cp1252", errors="replace")
    return text.replace("\r\n", "\n").replace("\r", "\n")


def parse_fields(line):
    if ("RefProduit" not in line and "SN" not in line
            and "CodeLot" not in line and "ModBusId" not in line):
        return {}
    return dict(FIELD_RE.findall(line))


def field_matches(line, name, value):
    return parse_fields(line).get(name) == str(value)


def timestamp_parts(line):
    match = STAMP_RE.search(line)
    return match.groups() if match else ("", "")


def format_ticket_fields(line):
    fields = parse_fields(line)
    day, hour = timestamp_parts(line)
    parts = [f"{day} | {hour}"] if day else []
    parts.extend(f"{key}: {fields[key]}" for key in ("RefProduit", "SN", "CodeLot", "ModBusId")
                 if key in fields)
    return "    ".join(parts), {
        "refprod": fields.get("RefProduit"), "sn": fields.get("SN"),
        "codelot": fields.get("CodeLot"), "modbusid": fields.get("ModBusId"),
    }


def is_gripper(line, operation, reference=None, loading_only=False):
    if "GRIPPER" not in line:
        return False
    match = GRIPPER_RE.search(line)
    return bool(match and match[1] == operation
                and (reference is None or match[2] == reference)
                and (not loading_only or re.search(r"\$0101(?![\w])", line)))


def parse_idpk(line):
    if "IDPK" not in line:
        return None
    match = IDPK_RE.search(line)
    if not match:
        return None
    dimensions = {"H" if key == "h" else key: int(value)
                  for key, value in DIMENSION_RE.findall(match[2])}
    return match[1], dimensions


def is_idpk_for(line, reference):
    if "IDPK" not in line:
        return False
    match = IDPK_RE.search(line)
    return bool(match and match[1] == reference)


def find_idpk_after(lines, reference, get_index):
    """Measurement logs can be separated from GET by many diagnostic lines."""
    for index in range(get_index + 1, len(lines)):
        line = lines[index]
        if is_gripper(line, "GET", reference, loading_only=True):
            break
        if is_idpk_for(line, reference):
            return index, line
    return None, None


def trace_start_datetime(text, fallback):
    # Creation headers retain the real date when archives are copied locally.
    for line in text.split("\n", 8)[:8]:
        if "creation de la trace" in line.lower() or "OUVERTURE DE LA TRACE" in line:
            match = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", line)
            if match:
                try:
                    return datetime.strptime(match[1], "%d/%m/%Y")
                except ValueError:
                    pass
    return fallback


def line_datetime(line, base):
    day, hour = timestamp_parts(line)
    if not day:
        return base
    try:
        # Rotating traces can cross month/year boundaries.
        year, month = base.year, base.month
        if int(day) < base.day - 15:
            year, month = (year + 1, 1) if month == 12 else (year, month + 1)
        elif int(day) > base.day + 15:
            year, month = (year - 1, 12) if month == 1 else (year, month - 1)
        hours, minutes, seconds = hour.split(":")
        seconds, _, fraction = seconds.partition(".")
        return datetime(year, month, int(day), int(hours), int(minutes), int(seconds),
                        int((fraction + "000000")[:6]))
    except ValueError:
        return base


def iter_timed_lines(text, fallback):
    current = trace_start_datetime(text, fallback)
    for line in text.splitlines():
        current = line_datetime(line, current)
        yield current, line


def trace_end_datetime(text, fallback):
    start = trace_start_datetime(text, fallback)
    # The final timestamp, rather than filesystem mtime, orders copied archives.
    for line in reversed(text.splitlines()):
        if timestamp_parts(line)[0]:
            return line_datetime(line, start)
    return start


def parse_modbus_line(line):
    # Preserve empty and trailing columns: their positions are register indices.
    return [value.strip() for value in line.rstrip("\r\n").split(";")]


def missing_robot_tickets(lines):
    """First occurrence of each ticket with no subsequent loading GET."""
    first_tickets = {}
    latest_get = {}
    for index, line in enumerate(lines):
        fields = parse_fields(line)
        if fields.get("ModBusId", "").isdigit() and fields.get("RefProduit"):
            first_tickets.setdefault((fields["ModBusId"], fields["RefProduit"]), index)
        if is_gripper(line, "GET", loading_only=True):
            latest_get[GRIPPER_RE.search(line)[2]] = index
    return [(modbus, reference, index)
            for (modbus, reference), index in first_tickets.items()
            if latest_get.get(reference, -1) <= index]
