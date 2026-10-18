"""Resolve packaged resources in both Python and PyInstaller launches."""

from pathlib import Path
import sys
import tkinter as tk


def resource_path(name):
    root = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))
    return root / 'assets' / name


def apply_icon(window):
    try:
        window.iconbitmap(str(resource_path('tracomega.ico')))
    except (tk.TclError, OSError):
        pass
