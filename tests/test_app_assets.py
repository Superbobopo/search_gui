from pathlib import Path
from unittest.mock import Mock, patch
import struct
import tempfile
import tkinter as tk
import unittest

from PIL import Image
import app_assets


class IconTests(unittest.TestCase):
    def test_active_icon_has_seven_png_frames_and_transparent_source(self):
        data = app_assets.resource_path(app_assets.ICON_FILENAME).read_bytes()
        self.assertEqual(struct.unpack_from('<HHH',data), (0,1,7))
        sizes = []
        for index in range(7):
            width,height,_,_,_,_,length,offset = struct.unpack_from('<BBBBHHII',data,6+index*16)
            self.assertEqual(width,height)
            self.assertEqual(data[offset:offset+8],b'\x89PNG\r\n\x1a\n')
            self.assertGreater(length,100)
            sizes.append(width or 256)
        self.assertEqual(sizes,[16,24,32,48,64,128,256])
        with Image.open(app_assets.resource_path('tracomega-v2.png')) as source:
            self.assertEqual(source.mode,'RGBA')
            self.assertEqual(source.getextrema()[3],(0,255))

    def test_pyinstaller_resource_path_and_ctk_delayed_icon_are_handled(self):
        with tempfile.TemporaryDirectory() as folder, patch.object(app_assets.sys,'_MEIPASS',folder,create=True):
            self.assertEqual(app_assets.resource_path(app_assets.ICON_FILENAME),Path(folder)/'assets'/app_assets.ICON_FILENAME)
        window = Mock(_iconbitmap_method_called=False)
        app_assets.apply_icon(window)
        window.iconbitmap.assert_called_once_with(str(app_assets.resource_path(app_assets.ICON_FILENAME)))
        self.assertTrue(window._tracomega_icon_applied)
        self.assertTrue(window._iconbitmap_method_called)

    def test_unavailable_icon_does_not_crash_window_creation(self):
        window = Mock()
        window.iconbitmap.side_effect = tk.TclError('unavailable icon')
        app_assets.apply_icon(window)
        self.assertNotIn('_tracomega_icon_applied',window.__dict__)
