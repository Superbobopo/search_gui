"""Readable deposit tracking table; row identities always retain TraceRef."""

from datetime import datetime
import os
import re
import tkinter as tk
from tkinter import ttk
import customtkinter as ctk
from ui_components import font, ActionButton

from selectable_text import make_selectable
from trace_parser import timestamp_parts
from trace_sources import compact_trace_label


def tracking_values(row, directories=None):
    when = datetime.fromtimestamp(row['time'])
    _, hour = timestamp_parts(row['text'])
    face = re.search(r'\bface\s*:\s*(\d+)(?:[ \t]+([A-Za-z$]*\d+)\b)?', row['text'])
    location = ('Face ' + face[1] + (' · ' + face[2] if face[2] else '')) if face else '—'
    compact = compact_trace_label(row['path'], directories)
    robot = compact.split(' \\', 1)[0] if ' \\' in compact else '—'
    return (when.strftime('%d/%m/%Y'), hour, robot, location, os.path.basename(row['path']),
            row.get('ref') or row.get('op_ref') or '—', row.get('sn') or '—', row.get('lot') or '—')


class TrackingView(ctk.CTkFrame):
    columns = ('date', 'hour', 'robot', 'location', 'trace', 'ref', 'sn', 'lot')

    def __init__(self, parent, app):
        super().__init__(parent, fg_color='transparent')
        self.app = app
        self.rows = {}
        self.values = {}
        self.reference = ''
        self._pending = None
        self._closed = False
        self.form = ctk.CTkFrame(self, corner_radius=8)
        self.form.pack(fill='x', pady=(4, 10))
        self.form.grid_columnconfigure((0, 1, 2), weight=1, uniform='criteria')
        self.form_fields = []
        for column, caption in enumerate(('Référence', 'SN', 'CodeLot')):
            field = ctk.CTkFrame(self.form, fg_color='transparent')
            field.grid(row=0, column=column, sticky='ew', padx=10, pady=8)
            field.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(field, text=caption, anchor='w', font=font('meta')).grid(row=0, column=0, sticky='w')
            entry = ctk.CTkEntry(field, width=100, height=32, font=font())
            entry.grid(row=1, column=0, sticky='ew')
            entry.bind('<Return>', lambda event: self.app.perform_tracking_search())
            self.form_fields.append((field, entry))
        self.reference_entry, self.sn_entry, self.lot_entry = [entry for _, entry in self.form_fields]
        self.search_button = ActionButton(self.form, text='Rechercher', role='primary',
                                         command=app.perform_tracking_search)
        self.search_button.grid(row=0, column=3, sticky='s', padx=10, pady=8)
        self._form_narrow = None
        tk.Misc.bind(self.form, '<Configure>', self._resize_form, add='+')
        self.heading = ctk.CTkLabel(self, text='Suivi des dépôts', anchor='w', font=font('title', True))
        self.heading.pack(fill='x', pady=(2, 0))
        self.count = ctk.CTkLabel(self, text='Saisissez au moins une référence, un SN ou un CodeLot', anchor='w',
                                  font=('Segoe UI', 11))
        self.count.pack(fill='x', pady=(0, 10))
        table = ctk.CTkFrame(self, corner_radius=8)
        table.pack(fill='both', expand=True)
        table.grid_columnconfigure(0, weight=1)
        table.grid_rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(table, columns=self.columns, show='headings',
                                  displaycolumns=tuple(key for key in self.columns if key!='trace'),
                                  selectmode='extended', style='Tracking.Treeview')
        for key, title, width, minimum in zip(self.columns,
                ('DATE', 'HEURE', 'ROBOT', 'EMPLACEMENT', 'TRACE', 'RÉFÉRENCE', 'SN', 'CODELOT'),
                (100, 100, 65, 135, 160, 140, 110, 130), (85, 85, 55, 100, 120, 110, 85, 100)):
            self.tree.heading(key, text=title, anchor='w' if key in ('location', 'trace') else 'center')
            self.tree.column(key, width=width, minwidth=minimum,
                             anchor='w' if key in ('location', 'trace') else 'center', stretch=key in ('location','ref'))
        self.tree.grid(row=0, column=0, sticky='nsew', padx=(1, 0), pady=(1, 0))
        scrollbar = ttk.Scrollbar(table, orient='vertical', command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky='ns', pady=(1, 0))
        horizontal = ttk.Scrollbar(table, orient='horizontal', command=self.tree.xview)
        horizontal.grid(row=1, column=0, sticky='ew', padx=(1, 0))
        self.tree.configure(yscrollcommand=scrollbar.set, xscrollcommand=horizontal.set)
        self.detail = ctk.CTkFrame(self, corner_radius=8)
        self.detail.pack(fill='x', pady=(10, 0))
        self.detail_title = ctk.CTkLabel(self.detail, text='Sélectionnez un dépôt pour consulter le détail.',
                                         anchor='w', justify='left', font=('Segoe UI', 12, 'bold'))
        self.detail_title.pack(fill='x', padx=12, pady=(8, 2))
        self.detail_source = ctk.CTkLabel(self.detail, text='', anchor='w', justify='left', font=('Segoe UI', 10))
        self.detail_source.pack(fill='x', padx=12, pady=2)
        self.detail_identity = ctk.CTkLabel(self.detail, text='', anchor='w', justify='left', font=font())
        self.detail_identity.pack(fill='x', padx=12, pady=2)
        self.detail_event = ctk.CTkLabel(self.detail, text='', anchor='w', justify='left',
                                         font=('Consolas', 10), wraplength=600)
        self.detail_event.pack(fill='x', padx=12, pady=(2, 8))
        for widget in (self.detail_title, self.detail_source, self.detail_identity, self.detail_event):
            make_selectable(widget)
        self.detail.bind('<Configure>', self._resize_detail, add='+')
        self.tree.bind('<<TreeviewSelect>>', self._selection_changed)
        self.tree.bind('<Double-1>', self._open)
        self.tree.bind('<Return>', self._open)
        self.tree.bind('<Control-c>', self._copy_rows)
        self.tree.bind('<Control-a>', self._select_all)
        self.tree.bind('<Button-3>', self._menu)
        self.menu = tk.Menu(self, tearoff=False)
        self.menu.add_command(label='Copier la cellule', command=self._copy_cell)
        self.menu.add_command(label='Copier la sélection', command=self._copy_rows)
        self.menu.add_separator()
        self.menu.add_command(label='Ouvrir la trace', command=self._open)
        self._cell = None
        self.bind('<Destroy>', self._destroyed, add='+')
        self.apply_theme()

    def _resize_form(self, event):
        if self._closed or event.widget is not self.form:
            return
        width = round(event.width/self._get_widget_scaling())
        self.heading.configure(wraplength=max(150, width-16))
        self.count.configure(wraplength=max(150, width-16))
        narrow = width < 700
        if narrow == self._form_narrow:
            return
        self._form_narrow = narrow
        self.form.grid_columnconfigure(2, weight=0 if narrow else 1, uniform='' if narrow else 'criteria')
        for i, (field, _) in enumerate(self.form_fields):
            field.grid_configure(row=1 if narrow and i==2 else 0,
                                 column=0 if narrow and i==2 else i)
        self.search_button.grid_configure(row=1 if narrow else 0, column=1 if narrow else 3, sticky='se')

    def _resize_detail(self, event):
        if self._closed:
            return
        for widget in (self.detail_title, self.detail_source, self.detail_identity, self.detail_event):
            widget.configure(wraplength=max(150, event.width - 28))

    def apply_theme(self):
        colors = self.app._palette
        self.form.configure(fg_color=colors['surface'])
        for field, entry in self.form_fields:
            entry.configure(fg_color=colors['bg'], border_color=colors['border'], text_color=colors['text'])
        self.search_button.apply_theme(colors)
        self.tree.tag_configure('even', background=colors['surface'], foreground=colors['text'])
        self.tree.tag_configure('odd', background=colors['bg'], foreground=colors['text'])

    def _cancel(self):
        if self._pending:
            self.after_cancel(self._pending)
            self._pending = None

    def _destroyed(self, event):
        if event.widget is self:
            self._cancel()

    def destroy(self):
        self._closed = True
        self._cancel()
        super().destroy()

    def begin(self, reference='', *, sn='', lot=''):
        self._cancel()
        self.reference = reference
        self.rows.clear()
        self.values.clear()
        self.tree.delete(*self.tree.get_children())
        criteria = [f'{label} : {value}' for label, value in (('Référence', reference), ('SN', sn), ('CodeLot', lot)) if value]
        self.heading.configure(text='Dépôts · ' + ' · '.join(criteria) if criteria else 'Suivi des dépôts')
        self.count.configure(text='Recherche en cours…' if criteria else 'Saisissez au moins une référence, un SN ou un CodeLot')
        self._show_detail(None)

    def show_results(self, rows):
        self._cancel()
        rows = iter(rows)
        def append():
            self._pending = None
            for _ in range(200):
                row = next(rows, None)
                if row is None:
                    count = len(self.rows)
                    files = len({r['path'] for r in self.rows.values()})
                    total = f'{count} dépôt{"s" if count != 1 else ""} · {files} trace{"s" if files != 1 else ""}'
                    self.count.configure(text=total + ' · du plus récent au plus ancien'
                                         if count else 'Aucun dépôt trouvé pour ces critères')
                    self.app.status.set(total)
                    return
                key = str(len(self.rows))
                values = tracking_values(row, self.app._selected_directories)
                self.rows[key], self.values[key] = row, values
                self.tree.insert('', 'end', iid=key, values=values,
                                 tags=('even' if int(key) % 2 == 0 else 'odd',))
                if key == '0':
                    self.tree.selection_set(key)
                    self.tree.focus(key)
                    self._show_detail(row)
            self.count.configure(text=f'{len(self.rows)} dépôts affichés…')
            self._pending = self.after(1, append)
        append()

    def _show_detail(self, row):
        if row is None:
            title, source, identity, text = 'Sélectionnez un dépôt pour consulter le détail.', '', '', ''
        else:
            date, hour, robot, location, _, ref, sn, lot = tracking_values(row, self.app._selected_directories)
            title = f'{location}   ·   {robot}   ·   {date} à {hour}'
            source = self.app._compact_trace_label(row['path'])
            identity = f'Référence : {ref}   ·   SN : {sn}   ·   CodeLot : {lot}'
            text = row['text']
        for widget, value in zip((self.detail_title, self.detail_source, self.detail_identity, self.detail_event), (title, source, identity, text)):
            widget.configure(text=value)
            widget.sync_selection_text()

    def _selection_changed(self, event=None):
        selected = self.tree.selection()
        self._show_detail(self.rows.get(selected[0]) if selected else None)

    def _open(self, event=None):
        key = self.tree.identify_row(event.y) if event and hasattr(event, 'y') and event.type != tk.EventType.KeyPress else self.tree.focus()
        row = self.rows.get(key)
        if row:
            self.app._open_indexed_file(row['path'], row['number'], row['text'])
        return 'break'

    def _select_all(self, event=None):
        self.tree.selection_set(self.tree.get_children())
        return 'break'

    def _copy(self, value):
        if value:
            self.clipboard_clear()
            self.clipboard_append(value)

    def _copy_rows(self, event=None):
        self._copy('\n'.join('\t'.join(self.values[key]) for key in self.tree.selection()))
        return 'break'

    def _copy_cell(self):
        if self._cell:
            key, column = self._cell
            self._copy(self.values[key][column])

    def _menu(self, event):
        key = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if key not in self.rows or not column or column == '#0':
            return
        self._cell = (key, self.columns.index(self.tree.column(column, 'id')))
        self.tree.focus(key)
        if key not in self.tree.selection():
            self.tree.selection_set(key)
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    def export_text(self):
        return '\n'.join(' | '.join(self.values[key]) + ' | PUTE | '
                         + self.app._compact_trace_label(row['path']) for key, row in self.rows.items())
