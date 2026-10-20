"""Robot diagnostics preserve indexed occurrences, including archived generations."""
import os
import tkinter as tk

import customtkinter as ctk

from trace_sources import compact_trace_label, trace_badge, trace_tooltip
from ui_components import (WindowShell, SplitView, VirtualCardList, StateCard,
                           ActionButton, selectable, font, style_labels)


class RobotErrorsView(WindowShell):
    def __init__(self, parent, app, errors):
        self.errors = list(errors)
        self.selected_index = None
        count = len(self.errors)
        files = len({(os.fspath(item['filepath']), getattr(item['filepath'], 'version_id', None)) for item in self.errors})
        super().__init__(parent, app, 'Erreurs robot',
                         f'{count} erreur'+('s' if count!=1 else '')+f' · {files} trace'+('s' if files!=1 else '')+' concernée'+('s' if files!=1 else ''))
        self.body.configure(fg_color=app._palette['bg'])
        self.split = SplitView(self.body)
        self.split.grid(row=0, column=0, sticky='nsew')
        self.left, self.right = self.split.left, self.split.right
        for pane in (self.left, self.right):
            pane.grid_columnconfigure(0, weight=1)
            pane.grid_rowconfigure(1, weight=1)
        self.list_title = ctk.CTkLabel(self.left, text='À EXAMINER', anchor='w', font=font('meta', True))
        self.list_title.grid(row=0, column=0, sticky='ew', padx=16, pady=(12, 8))
        self.cards = VirtualCardList(self.left, self)
        self.cards.grid(row=1, column=0, sticky='nsew', padx=(8, 6), pady=(0, 8))
        self.detail_title = selectable(self.right, text='', font=font('title', True), wraplength=600)
        self.detail_title.grid(row=0, column=0, sticky='ew', padx=20, pady=(16, 12))
        self.scroll = ctk.CTkScrollableFrame(self.right, corner_radius=0)
        self.scroll.grid(row=1, column=0, sticky='nsew', padx=(14, 6), pady=(0, 12))
        self.scroll.grid_columnconfigure(0, weight=1)
        self.status = StateCard(self.scroll, 'PRISE ROBOT')
        self.status.grid(row=0, column=0, sticky='ew', padx=4, pady=(0, 16))
        self.fields = ctk.CTkFrame(self.scroll, corner_radius=0, fg_color='transparent')
        self.fields.grid(row=1, column=0, sticky='ew', padx=6)
        self.fields.grid_columnconfigure(1, weight=1)
        self.values, self.captions = {}, []
        for i, (key, title) in enumerate((('ref', 'Référence'), ('modbus_id', 'ModBusId'), ('sn', 'SN'),
                                         ('lot', 'CodeLot'), ('when', 'Date et heure'), ('source', 'Trace'))):
            caption = ctk.CTkLabel(self.fields, text=title, anchor='w', font=font('meta'))
            caption.grid(row=i, column=0, sticky='nw', padx=(0, 16), pady=8)
            self.captions.append(caption)
            value = selectable(self.fields, text='', font=font(bold=key in ('ref', 'modbus_id')), wraplength=500)
            value.grid(row=i, column=1, sticky='ew', pady=8)
            self.values[key] = value
        self.open_button = ActionButton(self.scroll, text='Afficher dans la recherche', role='primary', command=self.activate)
        self.open_button.grid(row=2, column=0, sticky='w', padx=6, pady=(16, 12))
        self.ticket_heading = ctk.CTkLabel(self.scroll, text='Ticket original', anchor='w', font=font('section', True))
        self.ticket_heading.grid(row=3, column=0, sticky='ew', padx=6, pady=(12, 8))
        self.ticket = selectable(self.scroll, text='', font=('Consolas', 11), wraplength=600)
        self.ticket.grid(row=4, column=0, sticky='ew', padx=6, pady=(0, 12))
        self.provenance_button = ActionButton(self.scroll, text='Afficher la provenance', role='quiet', command=self.toggle_provenance)
        self.provenance_button.grid(row=5, column=0, sticky='w', padx=4, pady=8)
        self.provenance = selectable(self.scroll, text='', font=font('meta'), wraplength=600)
        self.provenance.grid(row=6, column=0, sticky='ew', padx=6, pady=4)
        self.provenance.grid_remove()
        self.menu = tk.Menu(self, tearoff=False)
        for title, kind in (('Copier le ModBusId', 'id'), ('Copier la ligne', 'row'), ('Copier le chemin complet', 'path')):
            self.menu.add_command(label=title, command=lambda value=kind: self.copy(value))
        self._width = None
        self.scroll.bind('<Configure>', self._resize_detail)
        if count:
            self.select(0)
        else:
            self.detail_title.configure(text='Aucune erreur détectée')
            self.detail_title.sync_selection_text()
            self.status.show('positive', 'Aucune prise GRIPPER GET manquante détectée')
            for widget in (self.fields, self.open_button, self.ticket_heading, self.ticket, self.provenance_button):
                widget.grid_remove()
        self.apply_theme()

    def card_content(self, index):
        item = self.errors[index]
        source = compact_trace_label(item['filepath'], self.app._selected_directories)
        return dict(title=f'{item["ref"]} · ID {item["modbus_id"]}',
                    subtitle=f'{item.get("date", "")} {item.get("heure", "")} · {source}',
                    badge=trace_badge(item['filepath']))

    def select(self, index):
        if index is None or index==self.selected_index or not 0<=index<len(self.errors):
            return 'break'
        self.selected_index = index
        item = self.errors[index]
        when = ' '.join(str(item.get(key) or '') for key in ('date', 'heure')).strip() or '—'
        self.detail_title.configure(text=f'ID {item["modbus_id"]} · {trace_badge(item["filepath"])}')
        self.detail_title.sync_selection_text()
        self.status.show('waiting', 'GRIPPER GET non retrouvé après ce ticket', when)
        values = dict(item, when=when, source=compact_trace_label(item['filepath'], self.app._selected_directories))
        for key, label in self.values.items():
            label.configure(text=str(values.get(key) or '—'))
            label.sync_selection_text()
        self.ticket.configure(text=item.get('text') or 'Ticket original indisponible.')
        self.ticket.sync_selection_text()
        self.provenance.configure(text=trace_tooltip(item['filepath'])+f'\nLigne originale : {item["number"]+1}')
        self.provenance.sync_selection_text()
        self.scroll._parent_canvas.yview_moveto(0)
        self.cards.draw()
        self.status.apply_theme(self.app._palette)
        return 'break'

    def activate(self, index=None):
        if index is not None:
            self.select(index)
        if self.selected_index is not None:
            item = self.errors[self.selected_index]
            self.app._on_error_line_click(item['filepath'], item['modbus_id'], item['ref'], item['number'])
        return 'break'

    def toggle_provenance(self):
        visible = bool(self.provenance.winfo_manager())
        self.provenance.grid_remove() if visible else self.provenance.grid()
        self.provenance_button.configure(text='Afficher la provenance' if visible else 'Masquer la provenance')

    def context(self, index, event):
        self.select(index)
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()
        return 'break'

    def copy(self, kind):
        if self.selected_index is not None:
            item = self.errors[self.selected_index]
            text = (os.fspath(item['filepath']) if kind=='path' else str(item['modbus_id']) if kind=='id' else
                    f'{item["ref"]}\t{item["modbus_id"]}\t{self.values["when"].cget("text")}\t{self.values["source"].cget("text")}')
            self.clipboard_clear(); self.clipboard_append(text)
        return 'break'

    def _resize_detail(self, event):
        if event.width==self._width:
            return
        self._width = event.width
        for label in (self.ticket, self.provenance, self.status.value):
            label.configure(wraplength=max(100, event.width-40))
        for label in self.values.values():
            label.configure(wraplength=max(100, event.width-150))

    def apply_theme(self):
        WindowShell.apply_theme(self)
        if not hasattr(self, 'split'):
            return
        colors = self.app._palette
        self.body.configure(fg_color=colors['bg'])
        self.split.apply_theme(colors)
        self.scroll.configure(fg_color=colors['surface'], bg_color=colors['surface'],
                              scrollbar_button_color=colors['border'], scrollbar_button_hover_color=colors['muted'])
        self.fields.configure(fg_color=colors['surface'], bg_color=colors['surface'])
        style_labels([self.detail_title, self.ticket_heading, self.ticket, *self.values.values()], colors)
        style_labels([self.list_title, self.provenance, *self.captions], colors, 'muted')
        self.status.apply_theme(colors)
        self.cards.apply_theme()
        for button in (self.open_button, self.provenance_button):
            button.apply_theme(colors)
