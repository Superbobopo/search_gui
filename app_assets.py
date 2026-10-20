"""Resolve packaged resources in both Python and PyInstaller launches."""

from pathlib import Path
import sys
import tkinter as tk

ICON_FILENAME = 'tracomega-v2.ico'


def resource_path(name):
    root = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))
    return root / 'assets' / name


def apply_icon(window):
    try:
        window.iconbitmap(str(resource_path(ICON_FILENAME)))
        # CTkToplevel aliases Tk's iconbitmap and otherwise replaces it after 200 ms.
        if hasattr(window, '_iconbitmap_method_called'):
            window._iconbitmap_method_called = True
        window._tracomega_icon_applied = True
        if isinstance(window, tk.Toplevel) and hasattr(window, '_windows_set_titlebar_icon'):
            # Some CTk releases also install an unconditional default-icon lambda.
            # Restore our icon once after that initializer, without touching other timers.
            root = window._root()

            def restore():
                window._tracomega_icon_timer = None
                if window.winfo_exists():
                    try:
                        window.iconbitmap(str(resource_path(ICON_FILENAME)))
                    except (tk.TclError, OSError):
                        pass

            previous = getattr(window, '_tracomega_icon_timer', None)
            if previous is not None:
                root.after_cancel(previous)
            window._tracomega_icon_timer = root.after(220, restore)
            if not getattr(window, '_tracomega_icon_destroy_bound', False):
                window._tracomega_icon_destroy_bound = True

                def destroyed(event):
                    timer = getattr(window, '_tracomega_icon_timer', None)
                    if event.widget is window and timer is not None:
                        root.after_cancel(timer)
                        window._tracomega_icon_timer = None

                window.bind('<Destroy>', destroyed, add='+')
    except (tk.TclError, OSError):
        pass
