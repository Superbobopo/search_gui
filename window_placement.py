"""Keep secondary windows inside their parent's monitor work area."""

import ctypes
from ctypes import wintypes
import os


def centered_bounds(parent, size, work):
    left, top, right, bottom = work
    width = min(size[0], right - left)
    height = min(size[1], bottom - top)
    x = max(left, min((parent[0] + parent[2] - width) // 2, right - width))
    y = max(top, min((parent[1] + parent[3] - height) // 2, bottom - height))
    return x, y, width, height


def windows_api():
    user = ctypes.WinDLL('user32', use_last_error=True)
    user.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
    user.GetAncestor.restype = wintypes.HWND
    user.MonitorFromWindow.argtypes = [wintypes.HWND, wintypes.DWORD]
    user.MonitorFromWindow.restype = wintypes.HANDLE
    user.GetMonitorInfoW.argtypes = [wintypes.HANDLE, ctypes.c_void_p]
    user.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
    user.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int,
                                  ctypes.c_int, ctypes.c_int, ctypes.c_int, wintypes.UINT]
    return user


def native_window_bounds(widget):
    user = windows_api()
    handle = user.GetAncestor(widget.winfo_id(), 2)  # GA_ROOT: includes decorations.
    rect = wintypes.RECT()
    if not user.GetWindowRect(handle, ctypes.byref(rect)):
        raise ctypes.WinError(ctypes.get_last_error())
    return handle, (rect.left, rect.top, rect.right, rect.bottom)


def monitor_work_area(widget):
    if os.name == 'nt':
        class MonitorInfo(ctypes.Structure):
            _fields_ = [('cbSize', wintypes.DWORD), ('rcMonitor', wintypes.RECT),
                        ('rcWork', wintypes.RECT), ('dwFlags', wintypes.DWORD)]
        user = windows_api()
        handle, _ = native_window_bounds(widget)
        info = MonitorInfo()
        info.cbSize = ctypes.sizeof(info)
        monitor = user.MonitorFromWindow(handle, 2)  # nearest monitor
        if not user.GetMonitorInfoW(monitor, ctypes.byref(info)):
            raise ctypes.WinError(ctypes.get_last_error())
        rect = info.rcWork
        return rect.left, rect.top, rect.right, rect.bottom
    return 0, 0, widget.winfo_screenwidth(), widget.winfo_screenheight()


def show_centered_popup(popup, parent, width, height, *, transient=True):
    popup.transient(parent if transient else '')
    popup.geometry(f'{width}x{height}')
    popup.update_idletasks()
    work = monitor_work_area(parent)
    if os.name == 'nt':
        handle, rect = native_window_bounds(popup)
        _, parent_rect = native_window_bounds(parent)
        x, y, outer_width, outer_height = centered_bounds(
            parent_rect, (rect[2] - rect[0], rect[3] - rect[1]), work)
        # Native coordinates support monitors to the left/above the primary display.
        if not windows_api().SetWindowPos(handle, None, x, y, outer_width, outer_height, 0x0010):
            raise ctypes.WinError(ctypes.get_last_error())  # SWP_NOACTIVATE
    else:
        rect = (parent.winfo_rootx(), parent.winfo_rooty(),
                parent.winfo_rootx() + parent.winfo_width(),
                parent.winfo_rooty() + parent.winfo_height())
        x, y, width, height = centered_bounds(rect, (width, height), work)
        popup.geometry(f'{width}x{height}+{x}+{y}')
    popup.deiconify()
    popup.lift()
