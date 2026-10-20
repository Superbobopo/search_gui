"""Check navigation icon clicks on a private Windows desktop."""

from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import customtkinter as ctk
from ui_theme import NavigationButton, navigation_icon
from ui_test_support import run_gui_check


def main():
    root = ctk.CTk()
    errors = []
    root.report_callback_exception = lambda *error: errors.append(error)
    calls = []
    try:
        for name in ('search', 'tracking', 'convoyage'):
            button = NavigationButton(root, text=name,
                                      command=lambda name=name: calls.append(name))
            button.pack()
            icon = navigation_icon(button, name)
            root.update()
            icon.event_generate('<Enter>')
            icon.event_generate('<Button-1>', x=8, y=8)
            icon.event_generate('<Leave>')
            root.update()
            assert calls == [name], calls
            calls.clear()
            button.configure(state='disabled')
            icon.event_generate('<Button-1>', x=8, y=8)
            root.update()
            assert not calls, calls
            button.destroy()
        assert not errors, errors
        print('Navigation icons: creation, click, hover and disabled state OK')
    finally:
        root.destroy()


if __name__ == '__main__':
    run_gui_check(main)
