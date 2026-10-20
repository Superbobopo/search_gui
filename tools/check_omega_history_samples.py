"""Exercise short real R1/R2 extracts using temporary history/index databases."""

import hashlib
from pathlib import Path
import sys
import tempfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from trace_parser import parse_fields
from trace_sources import discover_trace_files
from trace_store import TraceStore


def main():
    paths = discover_trace_files('tracomega')
    samples = []
    for robot in ('r1', 'r2'):
        choices = [Path(path) for path in paths if Path(path).parent.name.lower() in (robot, 'trace ' + robot)]
        choices.sort(key=lambda path: (path.name.lower() != 'tracomega.txt', str(path)))
        if not choices:
            raise SystemExit('Extraits réels R1/R2 indisponibles ; contrôle non exécuté.')
        for path in choices:
            stat = path.stat()
            with path.open('rb') as stream:
                data = stream.read(65536)
            if len(data) == 65536:
                data = data[:data.rfind(b'\n') + 1]
            if b'RefProduit=' in data:
                samples.append((robot, path, stat, data))
                break
        else:
            raise SystemExit('Aucun ticket dans les courts extraits ' + robot + ' ; contrôle non exécuté.')
    with tempfile.TemporaryDirectory() as folder:
        base = Path(folder)
        root = base / 'traces'
        copies = []
        for robot, original, stat, data in samples:
            destination = root / robot / original.name
            destination.parent.mkdir(parents=True)
            destination.write_bytes(data)
            copies.append(str(destination))
        store = TraceStore(base / 'index.sqlite3', directories=[root], history_database=base / 'history.sqlite3')
        assert not store.sync(copies)['failures']
        originals = []
        for row in store.metadata():
            ref = row['path']
            lines = store.file_lines(ref)
            number = next(i for i, line in enumerate(lines) if parse_fields(line).get('RefProduit'))
            details = store.details(ref, number)
            assert details['row']['text'] == lines[number]
            originals.append((ref, lines, number))
            with open(ref.path, 'wb') as stream:
                stream.write(b'RefProduit=GENERATION_SUIVANTE<Cr>ModBusId=999999<Cr>\n')
        assert not store.sync(copies)['failures']
        for path in copies:
            Path(path).unlink()
        store.rebuild()
        assert not store.sync([])['failures']
        for ref, lines, number in originals:
            assert store.file_lines(ref) == lines
            assert store.details(ref, number)['row']['text'] == lines[number]
        assert len(store.metadata()) == 4
    for robot, original, stat, data in samples:
        after = original.stat()
        assert (after.st_size, after.st_mtime_ns) == (stat.st_size, stat.st_mtime_ns)
        with original.open('rb') as stream:
            assert hashlib.sha256(stream.read(len(data))).digest() == hashlib.sha256(data).digest()
    print('OK extraits reels R1/R2 (64 Kio maximum chacun) : rotation, suppression, reconstruction et lignes originales ; sources inchangees, bases temporaires uniquement')


if __name__ == '__main__':
    main()
