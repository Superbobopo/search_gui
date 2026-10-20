"""Selectable read-only text over a label, preserving its layout and styling."""

import tkinter as tk
import customtkinter as ctk


def make_selectable(label):
    if hasattr(label, 'selection_text'):
        return label
    native = label._label if isinstance(label, ctk.CTkLabel) else label
    text = tk.Text(native.master, width=1, height=1, borderwidth=0, highlightthickness=0,
                   padx=0, pady=0, wrap='word', exportselection=False, cursor='xterm',
                   takefocus=True, undo=False)
    text._selectable_display = True
    label.selection_text = text
    previous_style = None
    previous_value = None

    def sync(event=None):
        nonlocal previous_style, previous_value
        if not text.winfo_exists() or not native.winfo_exists():
            return
        value = str(native.cget('text'))
        style = tuple(native.cget(option) for option in ('font', 'background', 'foreground', 'justify'))
        if previous_value == value and previous_style == style:
            return
        if previous_value != value:
            text.configure(state='normal')
            text.delete('1.0', 'end')
            text.insert('1.0', value)
            text.configure(state='disabled')
        if previous_style != style:
            text.configure(font=style[0], background=style[1], foreground=style[2],
                           selectbackground='#176baf', selectforeground='#ffffff',
                           inactiveselectbackground='#176baf')
            text.tag_configure('alignment', justify=style[3])
        text.tag_add('alignment', '1.0', 'end')
        text.lift(native)
        previous_value, previous_style = value, style

    def copy(event=None):
        try:
            selected = text.get('sel.first', 'sel.last')
        except tk.TclError:
            return 'break'
        text.clipboard_clear()
        text.clipboard_append(selected)
        return 'break'

    label.sync_selection_text = sync

    def select_all(event=None):
        text.tag_add('sel', '1.0', 'end-1c')
        return 'break'

    menu = tk.Menu(text, tearoff=False)
    menu.add_command(label='Copier', command=copy)
    menu.add_command(label='Tout sélectionner', command=select_all)
    text.copy_menu = menu

    def context(event):
        text.focus_set()
        menu.entryconfigure(0, state='normal' if text.tag_ranges('sel') else 'disabled')
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return 'break'

    def wheel(event):
        ancestor = label.master
        while ancestor is not None:
            if isinstance(ancestor, ctk.CTkScrollableFrame):
                ancestor._mouse_wheel_all(event)
                break
            ancestor = getattr(ancestor, 'master', None)
        return 'break'

    text.bind('<Button-1>', lambda event: text.focus_set(), add='+')
    text.bind('<<Copy>>', copy)
    text.bind('<Control-a>', select_all)
    text.bind('<Control-A>', select_all)
    text.bind('<Button-3>', context)
    text.bind('<MouseWheel>', wheel)
    native.bind('<Configure>', sync, add='+')
    text.bind('<Map>', sync, add='+')
    text.place(in_=native, x=0, y=0, relwidth=1, relheight=1)
    sync()
    return label
