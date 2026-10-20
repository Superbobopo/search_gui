"""Local palette adapter for the mixed Tk/ttk/CustomTkinter interface."""

import tkinter as tk
from tkinter import ttk
from tkinter import font as tkfont
import customtkinter as ctk


LIGHT = dict(bg='#f4f6fa', surface='#ffffff', raised='#edf2f8', text='#1e293b',
             muted='#526176', border='#cbd5e1', accent='#0f766e', hover='#0b5c56',
             accent_text='#ffffff', sidebar='#142c48', sidebar_active='#294367',
             info_bg='#e5efff', info_text='#174d99', positive_bg='#e6f7ef',
             positive_text='#146044', warning_bg='#fff3d6', warning_text='#805000',
             error_bg='#fdebec', error_text='#9f2431')
DARK = dict(bg='#111827', surface='#1b2535', raised='#253247', text='#e5edf8',
            muted='#aab8cc', border='#3b4a60', accent='#35d6c6', hover='#70ddd1',
            accent_text='#142c48', sidebar='#142c48', sidebar_active='#294367',
            info_bg='#1d3455', info_text='#b3d0ff', positive_bg='#173b32',
            positive_text='#a7eed4', warning_bg='#463516', warning_text='#ffd58a',
            error_bg='#442629', error_text='#ffb2b2')


def palette_for(mode):
    return dict(DARK if mode.lower() == 'dark' else LIGHT)


class ManagedTextbox(ctk.CTkTextbox):
    """Cancel CTk's recurring scrollbar checks when a view is rebuilt."""

    def __init__(self, *args, **kwargs):
        self._pending_callbacks = set()
        super().__init__(*args, **kwargs)

    def after(self, ms, func=None, *args):
        if func is None:
            return super().after(ms)
        handle = None
        def run():
            self._pending_callbacks.discard(handle)
            func(*args)
        handle = super().after(ms, run)
        self._pending_callbacks.add(handle)
        return handle

    def after_cancel(self, handle):
        self._pending_callbacks.discard(handle)
        return super().after_cancel(handle)

    def destroy(self):
        for handle in tuple(self._pending_callbacks):
            self.after_cancel(handle)
        super().destroy()


class NavigationButton(ctk.CTkButton):
    """Keep the Canvas icon in step with CTk hover, redraw and scaling."""

    def _sync_navigation_icon(self):
        icon = getattr(self, '_navigation_icon', None)
        label = getattr(self, '_text_label', None)
        if icon is not None and label is not None:
            icon.configure(bg=label.cget('bg'))

    def _draw(self, no_color_updates=False):
        super()._draw(no_color_updates)
        self._sync_navigation_icon()

    def _on_enter(self, event=None):
        super()._on_enter(event)
        self._sync_navigation_icon()

    def _on_leave(self, event=None):
        super()._on_leave(event)
        self._sync_navigation_icon()

    def _set_scaling(self, *args, **kwargs):
        super()._set_scaling(*args, **kwargs)
        icon = getattr(self, '_navigation_icon', None)
        if icon is not None:
            scale = self._get_widget_scaling()
            icon.scale('all', 0, 0, scale / icon._trace_icon_scale,
                       scale / icon._trace_icon_scale)
            icon._trace_icon_scale = scale
            icon.configure(width=round(22 * scale), height=round(22 * scale))
            icon.place_configure(x=round(10 * scale))


def navigation_icon(master, name):
    """Draw matching monochrome line icons without platform-dependent emoji."""
    picture = tk.Canvas(master, width=22, height=22, bg=LIGHT['sidebar'],
                        highlightthickness=0, borderwidth=0)
    color, width = '#d8e4f5', 1.6
    if name == 'search':
        picture.create_oval(3, 2.5, 12.5, 12, outline=color, width=width)
        picture.create_line(11.5, 11, 18, 18, fill=color, width=width)
    elif name == 'tracking':
        picture.create_line(3, 16, 3, 5, 11, 5, 11, 14, 18, 14, fill=color, width=width)
        for x, y in ((3, 16), (11, 5), (18, 14)):
            picture.create_oval(x-1.5, y-1.5, x+1.5, y+1.5, fill=color, outline=color)
    else:
        for x in (2, 8, 14):
            picture.create_rectangle(x, 6, x+4, 12, outline=color, width=1)
        picture.create_line(2, 16, 19, 16, fill=color, width=width)
        picture.create_line(16, 13, 19, 16, 16, 19, fill=color, width=1)
    picture.place(x=10, rely=.5, anchor='w')
    scale = master._get_widget_scaling()
    if scale != 1:
        picture.scale('all', 0, 0, scale, scale)
        picture.configure(width=round(22 * scale), height=round(22 * scale))
        picture.place_configure(x=round(10 * scale))
    picture._trace_icon_scale = scale
    master._navigation_icon = picture
    picture.bind('<Enter>', master._on_enter)
    picture.bind('<Leave>', master._on_leave)
    picture.bind('<Button-1>', lambda event: master.invoke())
    master._sync_navigation_icon()
    return picture


_ROLES = {
    '#f7f9fa': 'bg', '#f4f6fb': 'bg', '#fafbfc': 'surface', '#f5f7fa': 'raised',
    '#ffffff': 'surface', 'white': 'surface', '#e0e7ef': 'raised',
    '#dbeafe': 'raised', '#d1e3f5': 'info_bg', '#e3f2fd': 'info_bg',
    '#b3e5fc': 'info_bg', '#d2e3fc': 'border', '#cfd8dc': 'border',
    '#2196f3': 'accent', '#1565c0': 'info_text', '#00589b': 'info_text',
    '#1a5276': 'info_text', '#003566': 'info_text', '#314c7f': 'sidebar_active',
    '#182a44': 'sidebar', '#233557': 'sidebar', '#405060': 'muted',
    '#37474f': 'text', '#172a3a': 'text', '#333': 'text', '#333333': 'text', '#555': 'muted',
    '#555555': 'muted', 'black': 'text', '#000000': 'text',
    '#909090': 'muted', '#adb5bd': 'muted', '#a3a3a3': 'muted',
    '#c7c7c7': 'border', '#00876c': 'positive_text', '#e9fdf4': 'positive_bg',
    '#055300': 'positive_text', '#b66800': 'warning_text', '#9f8700': 'warning_text',
    '#fff4e3': 'warning_bg', '#fffde3': 'warning_bg', '#ffffe0': 'warning_bg',
    '#ffcdd2': 'error_bg', '#c62828': 'error_text', '#b71c1c': 'error_text',
    '#b31c1c': 'error_text', '#d02b20': 'error_text', '#fff59d': 'warning_bg',
    'yellow': 'warning_bg', '#ffff00': 'warning_bg', '#ffe66d': 'warning_bg',
}
for _palette in (LIGHT, DARK):
    for _role, _color in _palette.items():
        _ROLES.setdefault(_color, _role)


def themed_color(value, colors, foreground=False):
    if not isinstance(value, str):
        return value
    key = value.lower()
    if foreground and key in ('white', '#ffffff'):
        return colors['text']
    role = _ROLES.get(key)
    return colors.get(role, value)


def _keyboard_control(widget, colors):
    if not isinstance(widget, (ctk.CTkButton, ctk.CTkRadioButton, ctk.CTkOptionMenu)):
        return
    canvas = widget._canvas
    background = widget._apply_appearance_mode(widget.cget('bg_color'))
    if background == 'transparent':
        background = colors['surface']
    canvas.configure(highlightthickness=1, highlightbackground=background,
                     highlightcolor=colors['accent'])
    if getattr(widget, '_trace_keyboard_ready', False):
        return
    widget._trace_keyboard_ready = True
    canvas.configure(takefocus=lambda path=None: int(widget.cget('state') != 'disabled'))
    def activate(event=None):
        if widget.cget('state') != 'disabled':
            if isinstance(widget, ctk.CTkOptionMenu):
                widget._open_dropdown_menu()
            else:
                widget.invoke()
        return 'break'
    canvas.bind('<Return>', activate, add='+')
    canvas.bind('<space>', activate, add='+')


def apply_widget_theme(widget, colors):
    """Apply on the Tk thread; the historical result subtree stays untouched."""
    if (not widget.winfo_exists() or getattr(widget, '_preserve_classic_theme', False)
            or getattr(widget, '_theme_deferred', False)):
        return
    if isinstance(widget, ctk.CTkTextbox) and not hasattr(widget, '_x_scrollbar'):
        return
    if isinstance(widget, (ctk.CTkBaseClass, ctk.CTkScrollableFrame)):
        originals = getattr(widget, '_trace_theme_colors', {})
        options = {}
        for option in ('fg_color', 'bg_color', 'text_color', 'border_color', 'hover_color'):
            try:
                value = originals.setdefault(option, widget.cget(option))
                if isinstance(value, (tuple, list)) or value == 'transparent':
                    continue
                if option == 'text_color' and isinstance(widget, ctk.CTkButton) and value in ('white', '#ffffff'):
                    options[option] = colors['accent_text']
                else:
                    options[option] = themed_color(value, colors, option == 'text_color')
            except (ValueError, tk.TclError, AttributeError):
                pass
        widget._trace_theme_colors = originals
        if options:
            widget.configure(**options)
        _keyboard_control(widget, colors)
        private = {id(getattr(widget, name, None)) for name in
                   ('_canvas', '_entry', '_label', '_text_label', '_image_label', '_textbox', '_scrollbar')}
        for child in widget.winfo_children():
            if id(child) not in private:
                apply_widget_theme(child, colors)
        return
    if not isinstance(widget, ttk.Widget):
        originals = getattr(widget, '_trace_theme_colors', {})
        options = {}
        for option in ('background', 'foreground', 'insertbackground', 'highlightbackground',
                       'highlightcolor', 'selectbackground', 'selectforeground',
                       'activebackground', 'activeforeground', 'disabledforeground'):
            try:
                value = originals.setdefault(option, widget.cget(option))
                if option in ('foreground', 'activeforeground') and isinstance(widget, tk.Button):
                    options[option] = colors['text']
                else:
                    options[option] = themed_color(value, colors, 'foreground' in option)
            except (tk.TclError, ValueError):
                pass
        widget._trace_theme_colors = originals
        if options:
            widget.configure(**options)
        if isinstance(widget, (tk.Label, tk.Button, tk.LabelFrame)):
            try:
                font = tkfont.Font(root=widget, font=widget.cget('font'))
                if font.actual('family').lower() == 'arial':
                    widget.configure(font=('Segoe UI', font.actual('size'),
                                           font.actual('weight'), font.actual('slant')))
            except tk.TclError:
                pass
        if isinstance(widget, tk.Text):
            tags = getattr(widget, '_trace_theme_tags', {})
            for tag in widget.tag_names():
                if tag == 'sel':
                    continue
                original = tags.setdefault(tag, {key: widget.tag_cget(tag, key)
                                                for key in ('background', 'foreground')})
                mapped = {key: themed_color(value, colors, key == 'foreground')
                          for key, value in original.items() if value}
                if mapped:
                    if 'background' in mapped and 'foreground' not in mapped:
                        mapped['foreground'] = colors['text']
                    widget.tag_configure(tag, **mapped)
            widget._trace_theme_tags = tags
    for child in widget.winfo_children():
        apply_widget_theme(child, colors)
