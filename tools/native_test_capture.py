"""Capture one native widget on the isolated test desktop, never the user desktop."""
import ctypes
from ctypes import wintypes
import os

from PIL import Image


def widget_image(widget):
    if not os.environ.get('TRACOMEGA_TEST_DESKTOP'):
        raise RuntimeError('Native captures require the private test desktop.')
    user, gdi = ctypes.WinDLL('user32'), ctypes.WinDLL('gdi32')
    user.GetDC.argtypes = [wintypes.HWND]
    user.GetDC.restype = wintypes.HDC
    user.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
    user.PrintWindow.argtypes = [wintypes.HWND, wintypes.HDC, wintypes.UINT]
    gdi.CreateCompatibleDC.argtypes = [wintypes.HDC]
    gdi.CreateCompatibleDC.restype = wintypes.HDC
    gdi.CreateCompatibleBitmap.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int]
    gdi.CreateCompatibleBitmap.restype = wintypes.HBITMAP
    gdi.SelectObject.argtypes = [wintypes.HDC, wintypes.HANDLE]
    gdi.SelectObject.restype = wintypes.HANDLE
    gdi.GetDIBits.argtypes = [wintypes.HDC, wintypes.HBITMAP, wintypes.UINT, wintypes.UINT,
                              ctypes.c_void_p, ctypes.c_void_p, wintypes.UINT]
    gdi.DeleteObject.argtypes = [wintypes.HANDLE]
    gdi.DeleteDC.argtypes = [wintypes.HDC]
    gdi.BitBlt.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                           wintypes.HDC, ctypes.c_int, ctypes.c_int, wintypes.DWORD]
    class BitmapInfo(ctypes.Structure):
        _fields_ = [('size', wintypes.DWORD), ('width', wintypes.LONG), ('height', wintypes.LONG),
                    ('planes', wintypes.WORD), ('bits', wintypes.WORD), ('compression', wintypes.DWORD),
                    ('image_size', wintypes.DWORD), ('xppm', wintypes.LONG), ('yppm', wintypes.LONG),
                    ('used', wintypes.DWORD), ('important', wintypes.DWORD)]
    widget.update_idletasks()
    width, height = widget.winfo_width(), widget.winfo_height()
    handle = widget.winfo_id()
    dc = user.GetDC(handle)
    memory = gdi.CreateCompatibleDC(dc)
    bitmap = gdi.CreateCompatibleBitmap(dc, width, height)
    old = gdi.SelectObject(memory, bitmap)
    try:
        if not gdi.BitBlt(memory, 0, 0, width, height, dc, 0, 0, 0x00CC0020):
            raise RuntimeError('Private widget capture failed; no desktop fallback.')
        info = BitmapInfo(ctypes.sizeof(BitmapInfo), width, -height, 1, 32)
        buffer = ctypes.create_string_buffer(width*height*4)
        if not gdi.GetDIBits(memory, bitmap, 0, height, buffer, ctypes.byref(info), 0):
            raise ctypes.WinError()
        return Image.frombytes('RGB', (width,height), buffer.raw, 'raw', 'BGRX')
    finally:
        gdi.SelectObject(memory, old)
        gdi.DeleteObject(bitmap)
        gdi.DeleteDC(memory)
        user.ReleaseDC(handle, dc)


def color_pixels(image, color):
    rgb = tuple(int(color[i:i+2],16) for i in (1,3,5))
    return sum(1 for pixel in image.getdata() if max(abs(a-b) for a,b in zip(pixel,rgb)) <= 12)
