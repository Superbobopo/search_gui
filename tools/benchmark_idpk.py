"""Reproducible service timings; temporary traces and databases only."""
from pathlib import Path
import argparse
import json
import statistics
import sys
import tempfile
import time
import tracemalloc

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from trace_store import TraceStore


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--output', required=True)
    args = parser.parse_args()
    output = Path(args.output)
    if output.exists():
        raise FileExistsError(output)
    with tempfile.TemporaryDirectory() as folder:
        root = Path(folder)
        paths = []
        for file in range(8):
            path = root / 'traces' / ('R1' if file % 2 else 'R2') / f'TracOmega{file:03}.old'
            path.parent.mkdir(parents=True, exist_ok=True)
            lines = ['creation de la trace de ... Le 14/10/2026\n']
            for i in range(60):
                lines.extend(['B 14|09:00:00.0 routine sans evenement utile\n'] * 100)
                lines.extend((f'14|10:{i:02}:00.0 RefProduit=REF<Cr>SN=SN<Cr>CodeLot=LOT<Cr>ModBusId={i}<Cr>\n',
                              'G 14|10:00:01.0 GRIPPER GET 1x boîte(s): Code:REF $0101\n',
                              "r [Tgripper = 1s rep ='OK']\n",
                              f'B 14|10:{i:02}:02.0 IDPK pour REF : l={100+file} H=50 L=195\n',
                              'G 14|10:00:03.0 GRIPPER PUTE 1x boîte(s): Code:REF D23\n'))
            path.write_text(''.join(lines), encoding='utf8')
            paths.append(str(path))
        store = TraceStore(root / 'cache.sqlite3')
        store.sync(paths)
        path, hits = store.search('ModBusId=59', True)[-1]
        number = hits[0][0]
        cold, current, warm, rows, peaks = [], [], [], [], []
        original = store._rows
        for _ in range(3):
            store._analysis_cache.clear()
            count = [0]
            def counted(*args, **kwargs):
                result = original(*args, **kwargs)
                count[0] += len(result)
                return result
            store._rows = counted
            tracemalloc.start()
            start = time.perf_counter()
            if hasattr(store, 'current_measurement_for_occurrence'):
                local = store.current_measurement_for_occurrence(path, number)
                assert local['current'] is not None
                current.append((time.perf_counter()-start)*1000)
            result = store.comparison_for_occurrence(path, number)
            assert result['current'] is not None
            cold.append((time.perf_counter()-start)*1000)
            if len(current) < len(cold):
                current.append(cold[-1])
            rows.append(count[0])
            peaks.append(tracemalloc.get_traced_memory()[1])
            tracemalloc.stop()
            start = time.perf_counter()
            store.comparison_for_occurrence(path, number)
            warm.append((time.perf_counter()-start)*1000)
        report = dict(corpus=dict(files=8, tickets=480, noise_lines=48000),
                      current_ms=statistics.median(current), complete_ms=statistics.median(cold),
                      warm_ms=statistics.median(warm), transferred_rows=statistics.median(rows),
                      peak_python_bytes=statistics.median(peaks), cache_weight=store._analysis_cache.weight)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(report, indent=2), encoding='utf8')
        print(json.dumps(report))


if __name__ == '__main__':
    main()
