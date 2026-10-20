"""Shared, scoped styles for the robot context and conveyor diagnostics."""
from tkinter import ttk


class PopupTreeview(ttk.Treeview):
    def focus(self, item=None):
        # CustomTkinter may restore native focus after the old window was closed.
        return '' if getattr(self, '_disposed', False) else super().focus(item)

    def destroy(self):
        self._disposed = True
        super().destroy()


def configure_popup_style(widget, colors):
    from ui_components import font
    style = ttk.Style(widget)
    style.configure('Popup.TFrame', background=colors['surface'])
    style.configure('Popup.TLabel', background=colors['surface'], foreground=colors['muted'])
    style.configure('Popup.TButton', padding=(12, 7), borderwidth=0,
                    background=colors['raised'], foreground=colors['text'], relief='flat')
    style.map('Popup.TButton', background=[('active', colors['border'])])
    style.configure('Popup.Active.TButton', padding=(12, 7), borderwidth=0,
                    background=colors['accent'], foreground=colors['accent_text'], relief='flat')
    style.map('Popup.Active.TButton', background=[('disabled', colors['raised']), ('active', colors['hover'])],
              foreground=[('disabled', colors['muted']), ('active', colors['accent_text'])])
    for direction, sticky in (('Vertical', 'ns'), ('Horizontal', 'we')):
        name = f'Popup.{direction}.TScrollbar'
        style.layout(name, [(f'{direction}.Scrollbar.trough', {'sticky': sticky, 'children': [
            (f'{direction}.Scrollbar.thumb', {'expand': True, 'sticky': 'nswe'})]})])
        style.configure(name, width=10, borderwidth=0, gripcount=0,
                        background=colors['border'], troughcolor=colors['bg'],
                        lightcolor=colors['border'], darkcolor=colors['border'])
        style.map(name, background=[('active', colors['muted'])])
    for name in ('Adjacent.Treeview', 'Convoyage.Treeview', 'GetErrors.Treeview'):
        style.configure(name, font=font(native=True), rowheight=32, borderwidth=0,
                        relief='flat', background=colors['surface'],
                        fieldbackground=colors['surface'], foreground=colors['text'])
        style.layout(name, [('Treeview.treearea', {'sticky': 'nswe'})])
        style.configure(name + '.Heading', font=font('meta', True, native=True), padding=(8, 10),
                        background=colors['raised'], foreground=colors['muted'],
                        borderwidth=0, relief='flat')
        style.map(name + '.Heading', background=[('active', colors['raised'])])
    style.map('Convoyage.Treeview', background=[('selected', colors['info_bg'])],
              foreground=[('selected', colors['text'])])
    return style
