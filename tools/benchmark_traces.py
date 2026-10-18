"""Reproducible local benchmark; the database is a disposable workspace artifact."""

import argparse
import ctypes
from ctypes import wintypes
from pathlib import Path
import sys
import time

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from trace_store import TraceStore, cache_path
from trace_sources import discover_trace_files


def memory_mb():
    if sys.platform != 'win32':
        return None
    class Counters(ctypes.Structure):
        _fields_ = [('cb', wintypes.DWORD), ('faults', wintypes.DWORD)] + [
            (name, ctypes.c_size_t) for name in ('peak', 'working', 'pool_peak', 'pool', 'non_peak', 'non', 'page', 'page_peak')]
    counters = Counters(); counters.cb = ctypes.sizeof(counters)
    ctypes.windll.kernel32.GetCurrentProcess.restype = wintypes.HANDLE
    ctypes.windll.psapi.GetProcessMemoryInfo.argtypes = [wintypes.HANDLE, ctypes.POINTER(Counters), wintypes.DWORD]
    if not ctypes.windll.psapi.GetProcessMemoryInfo(ctypes.windll.kernel32.GetCurrentProcess(), ctypes.byref(counters), counters.cb):
        return None
    return round(counters.working / 1048576, 1)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--database', type=Path, default=Path('.benchmark/trace_index.sqlite3'))
    parser.add_argument('--application-cache', action='store_true', help='Mesurer le cache utilisé par l’application')
    args = parser.parse_args()
    if args.application_cache:
        args.database = cache_path()
    start = time.perf_counter()
    store = TraceStore(args.database)
    print('Ouverture:', round(time.perf_counter() - start, 3), 's', flush=True)
    def progress(done, total, path, published):
        if published:
            print(f'{done}/{total} {Path(path).name} {time.perf_counter() - start:.1f}s RAM={memory_mb()} Mo', flush=True)
    result = store.sync(discover_trace_files(), progress)
    print('Synchronisation:', result, 'en', round(time.perf_counter() - start, 3), 's', flush=True)
    for key, exact in [('ModBusId=1194', True), ('RefProduit=A0061574AC', False),
                       ('ModBusId=999999999', True), ('RefProduit=REF_ABSENTE_999', False)]:
        timings = []
        for _ in range(3):
            before = time.perf_counter()
            with store.snapshot():
                groups = store.search(key, exact)
                if groups:
                    details = store.details(groups[0][0], groups[0][1][0][0])
            timings.append(round((time.perf_counter() - before) * 1000, 1))
        print(key, 'recherche + details (ms):', timings, 'fichiers:', len(groups), flush=True)
    before = time.perf_counter()
    reopened = TraceStore(args.database)
    reopened.sync(discover_trace_files())
    print('Reouverture + verification:', round(time.perf_counter() - before, 3), 's', flush=True)
    print('Cache total:', round(sum(p.stat().st_size for p in args.database.parent.glob(args.database.name + '*')) / 1048576, 1), 'Mo', flush=True)
    print('RAM finale:', memory_mb(), 'Mo', flush=True)


if __name__ == '__main__':
    main()
