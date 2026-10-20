"""Actual Tk and CTk window icons, on a private Windows desktop only."""
from pathlib import Path
from unittest.mock import patch
import ctypes
from ctypes import wintypes
import os
import struct
import sys
import tempfile
import time

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0,str(PROJECT))
import customtkinter as ctk
import tkinter as tk
import search_gui as gui
from app_assets import ICON_FILENAME,resource_path,apply_icon
from tools.ui_test_support import run_gui_check


def pixels(user, icon):
    class IconInfo(ctypes.Structure):
        _fields_ = [('icon',wintypes.BOOL),('x',wintypes.DWORD),('y',wintypes.DWORD),
                    ('mask',wintypes.HBITMAP),('color',wintypes.HBITMAP)]

    class Bitmap(ctypes.Structure):
        _fields_ = [('type',wintypes.LONG),('width',wintypes.LONG),('height',wintypes.LONG),
                    ('stride',wintypes.LONG),('planes',wintypes.WORD),('bits',wintypes.WORD),('data',ctypes.c_void_p)]

    gdi = ctypes.WinDLL('gdi32',use_last_error=True)
    user.GetIconInfo.argtypes = [wintypes.HICON,ctypes.POINTER(IconInfo)]
    gdi.GetObjectW.argtypes = [wintypes.HANDLE,ctypes.c_int,ctypes.c_void_p]
    gdi.CreateCompatibleDC.argtypes = [wintypes.HDC]; gdi.CreateCompatibleDC.restype = wintypes.HDC
    gdi.GetDIBits.argtypes = [wintypes.HDC,wintypes.HBITMAP,wintypes.UINT,wintypes.UINT,
                             ctypes.c_void_p,ctypes.c_void_p,wintypes.UINT]
    gdi.DeleteObject.argtypes = [wintypes.HANDLE]; gdi.DeleteDC.argtypes = [wintypes.HDC]
    info, bitmap = IconInfo(), Bitmap()
    assert user.GetIconInfo(icon,ctypes.byref(info))
    dc = gdi.CreateCompatibleDC(None)
    try:
        assert gdi.GetObjectW(info.color,ctypes.sizeof(bitmap),ctypes.byref(bitmap))
        width,height = bitmap.width,bitmap.height
        header = ctypes.create_string_buffer(struct.pack('<IiiHHIIiiII',40,width,-height,1,32,0,0,0,0,0,0))
        buffer = ctypes.create_string_buffer(width*height*4)
        assert gdi.GetDIBits(dc,info.color,0,height,buffer,header,0)
        return width,height,buffer.raw
    finally:
        gdi.DeleteDC(dc); gdi.DeleteObject(info.color); gdi.DeleteObject(info.mask)


def main():
    with tempfile.TemporaryDirectory() as folder, patch.dict(os.environ,LOCALAPPDATA=folder), \
         patch.object(gui.IndexedAppMixin,'_index_loop',lambda app:None):
        app = gui.SearchApp()
        errors = []
        app.report_callback_exception = lambda *error:errors.append(error)
        plain = tk.Toplevel(app); plain.title('Icon test Tk')
        modern = ctk.CTkToplevel(app); modern.title('Icon test CTk')
        user = ctypes.WinDLL('user32',use_last_error=True)
        user.GetParent.argtypes = [wintypes.HWND]; user.GetParent.restype = wintypes.HWND
        user.SendMessageW.argtypes = [wintypes.HWND,wintypes.UINT,wintypes.WPARAM,wintypes.LPARAM]
        user.SendMessageW.restype = wintypes.LPARAM
        user.LoadImageW.argtypes = [wintypes.HINSTANCE,wintypes.LPCWSTR,wintypes.UINT,
                                    ctypes.c_int,ctypes.c_int,wintypes.UINT]
        user.LoadImageW.restype = wintypes.HANDLE
        user.DestroyIcon.argtypes = [wintypes.HICON]
        try:
            # Wait beyond CTk's own delayed icon callback, including after a theme change.
            for theme in ('Clair','Sombre'):
                app._change_theme(theme)
                until = time.perf_counter()+.4
                while time.perf_counter()<until:
                    app.update()
                    if errors:
                        raise errors[0][1]
                    time.sleep(.004)
                for window in (app,plain,modern):
                    assert window._tracomega_icon_applied
                    handle = user.GetParent(window.winfo_id())
                    for size in (0,1):
                        icon = user.SendMessageW(handle,0x7f,size,0)
                        assert icon, 'Missing native icon'
                        actual = pixels(user,icon)
                        expected = user.LoadImageW(None,str(resource_path(ICON_FILENAME)),1,actual[0],actual[1],0x10)
                        assert expected
                        try:
                            expected_pixels = pixels(user,expected)
                            from PIL import Image,ImageChops,ImageStat
                            images = [Image.frombytes('RGBA',v[:2],v[2],'raw','BGRA') for v in (actual,expected_pixels)]
                            # Tk and LoadImage can select/resample different ICO frames.
                            difference = ImageStat.Stat(ImageChops.difference(*(im.resize((16,16)) for im in images))).mean
                            assert max(difference)<16, (theme,window,size,difference,'Wrong icon pixels')
                            if theme=='Sombre' and size==1:
                                images[0].save(PROJECT/'.benchmark'/f'icon_native_{type(window).__name__}_20261020.png')
                        finally:
                            user.DestroyIcon(expected)
            apply_icon(modern)
            assert modern._tracomega_icon_timer is not None
            modern.destroy()
            assert modern._tracomega_icon_timer is None
            until = time.perf_counter()+.3
            while time.perf_counter()<until:
                app.update(); time.sleep(.004)
            assert not errors
            print('OK active ICO on main/Tk/CTk windows, small/large native handles, both themes, '
                  'unchanged beyond CTk delayed initialization, restore cancelled on close; '
                  'no indexing or user desktop focus')
        finally:
            app._close_indexed_app(); app._index_thread.join(5)
            app.tasks.executor.shutdown(wait=True,cancel_futures=True)


if __name__=='__main__':
    run_gui_check(main)
