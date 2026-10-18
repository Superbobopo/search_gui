"""Run GUI checks on an isolated Windows desktop, never on the user's desktop."""

import ctypes
from contextlib import redirect_stdout, redirect_stderr
from ctypes import wintypes
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import traceback
import uuid


def run_gui_check(main):
    if os.environ.get('TRACOMEGA_TEST_DESKTOP'):
        with open(os.environ['TRACOMEGA_TEST_LOG'], 'w', encoding='utf-8') as log, \
                redirect_stdout(log), redirect_stderr(log):
            try:
                main()
            except Exception:
                traceback.print_exc()
                raise SystemExit(1)
        return
    if os.name != 'nt':
        raise SystemExit('Contrôle graphique réservé à un bureau Windows isolé.')
    raise SystemExit(_launch_on_private_desktop())


def _launch_on_private_desktop():
    class StartupInfo(ctypes.Structure):
        _fields_ = [('cb', wintypes.DWORD), ('lpReserved', wintypes.LPWSTR),
                    ('lpDesktop', wintypes.LPWSTR), ('lpTitle', wintypes.LPWSTR),
                    ('dwX', wintypes.DWORD), ('dwY', wintypes.DWORD),
                    ('dwXSize', wintypes.DWORD), ('dwYSize', wintypes.DWORD),
                    ('dwXCountChars', wintypes.DWORD), ('dwYCountChars', wintypes.DWORD),
                    ('dwFillAttribute', wintypes.DWORD), ('dwFlags', wintypes.DWORD),
                    ('wShowWindow', wintypes.WORD), ('cbReserved2', wintypes.WORD),
                    ('lpReserved2', ctypes.c_void_p), ('hStdInput', wintypes.HANDLE),
                    ('hStdOutput', wintypes.HANDLE), ('hStdError', wintypes.HANDLE)]

    class ProcessInfo(ctypes.Structure):
        _fields_ = [('hProcess', wintypes.HANDLE), ('hThread', wintypes.HANDLE),
                    ('dwProcessId', wintypes.DWORD), ('dwThreadId', wintypes.DWORD)]

    user = ctypes.WinDLL('user32', use_last_error=True)
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    user.CreateDesktopW.argtypes = [wintypes.LPCWSTR, ctypes.c_void_p, ctypes.c_void_p,
                                    wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p]
    user.CreateDesktopW.restype = wintypes.HANDLE
    user.CloseDesktop.argtypes = [wintypes.HANDLE]
    kernel.CreateProcessW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, ctypes.c_void_p,
                                      ctypes.c_void_p, wintypes.BOOL, wintypes.DWORD,
                                      ctypes.c_void_p, wintypes.LPCWSTR,
                                      ctypes.POINTER(StartupInfo), ctypes.POINTER(ProcessInfo)]
    kernel.CreateProcessW.restype = wintypes.BOOL
    kernel.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    kernel.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    name = 'TracOmegaTest_' + uuid.uuid4().hex
    desktop = user.CreateDesktopW(name, None, None, 0, 0x10000000, None)
    if not desktop:
        raise ctypes.WinError(ctypes.get_last_error())
    process = ProcessInfo()
    log_folder = tempfile.TemporaryDirectory()
    log_path = Path(log_folder.name) / 'gui-check.log'
    try:
        startup = StartupInfo(cb=ctypes.sizeof(StartupInfo), lpDesktop=name)
        environment = dict(os.environ, TRACOMEGA_TEST_DESKTOP=name, TRACOMEGA_TEST_LOG=str(log_path))
        block = ctypes.create_unicode_buffer('\0'.join(f'{k}={v}' for k, v in sorted(environment.items())) + '\0\0')
        command = ctypes.create_unicode_buffer(subprocess.list2cmdline(
            [sys.executable, '-B', str(Path(sys.argv[0]).resolve()), *sys.argv[1:]]))
        if not kernel.CreateProcessW(None, command, None, None, False,
                                     0x08000400, block, os.getcwd(), ctypes.byref(startup), ctypes.byref(process)):
            raise ctypes.WinError(ctypes.get_last_error())
        kernel.WaitForSingleObject(process.hProcess, 0xFFFFFFFF)
        code = wintypes.DWORD()
        if not kernel.GetExitCodeProcess(process.hProcess, ctypes.byref(code)):
            raise ctypes.WinError(ctypes.get_last_error())
        return code.value
    finally:
        if process.hThread:
            kernel.CloseHandle(process.hThread)
        if process.hProcess:
            kernel.CloseHandle(process.hProcess)
        user.CloseDesktop(desktop)
        if log_path.exists():
            print(log_path.read_text(encoding='utf-8'), end='')
        log_folder.cleanup()
