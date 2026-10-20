"""Reproduce resize delivery during destruction, on a private Windows desktop."""
from pathlib import Path
import sys
import threading
import time
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import customtkinter as ctk
from convoyage_workspace import ConvoyageWorkspace
from tools.ui_test_support import run_gui_check
from ui_components import ActionButton
from ui_theme import LIGHT


def main():
    root = ctk.CTk()
    errors, delivered = [], []
    root.report_callback_exception = lambda *args: errors.append(args)
    app = SimpleNamespace(_palette=LIGHT, _active_panel='search')
    workspace = ConvoyageWorkspace(root, app)
    workspace.pack(fill='both', expand=True)
    started, release = threading.Event(), threading.Event()

    def pending(cancel):
        started.set()
        release.wait(5)
        return 'late result'

    workspace.tasks.submit('pending', pending, delivered.append, delivered.append)
    close = ActionButton(root, text='Fermer', command=workspace.destroy)
    close.pack()
    root.geometry('900x700')
    root.update()
    assert started.wait(2)
    # Native labels disappear before their parent receives <Destroy>. Deliver a
    # geometry event in that interval, as nested Tk layout processing can do.
    native = workspace.pages['parcours']['status']._label
    native.bind('<Destroy>', lambda event: workspace.event_generate(
        '<Configure>', width=400, height=300), add='+')
    try:
        close.invoke()
        release.set()
        workspace.tasks.executor.shutdown(wait=True, cancel_futures=True)
        end = time.monotonic()+.2
        while time.monotonic()<end:
            root.update()
            time.sleep(.005)
        workspace.tasks.drain()
        assert not errors, errors
        assert workspace._closed and workspace.tasks.closed
        assert not delivered
        assert not workspace.winfo_exists()
        print('OK: destruction via button, resize during child destruction, cancelled timer and late result')
    finally:
        release.set()
        if workspace.winfo_exists():
            workspace.destroy()
        workspace.tasks.executor.shutdown(wait=True, cancel_futures=True)
        root.destroy()


if __name__=='__main__':
    run_gui_check(main)
