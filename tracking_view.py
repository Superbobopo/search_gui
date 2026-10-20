"""Readable deposit tracking table; row identities always retain TraceRef."""

from datetime import datetime
import os
import re
import tkinter as tk
from tkinter import ttk
import customtkinter as ctk
from ui_components import font

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
    return when.strftime('%d/%m/%Y'), hour, robot, location, os.path.basename(row['path'])


class TrackingView(ctk.CTkFrame):
    columns = ('date', 'hour', 'robot', 'location', 'trace')

    def __init__(self, parent, app):
        super().__init__(parent, fg_color='transparent')
        self.app = app
        self.rows = {}
        self.values = {}
        self.reference = ''
        self._pending = None
        self.heading = ctk.CTkLabel(self, text='Suivi des dépôts', anchor='w', font=font('title', True))
        self.heading.pack(fill='x', pady=(2, 0))
        self.count = ctk.CTkLabel(self, text='Recherchez une référence pour afficher ses dépôts.', anchor='w',
                                  font=('Segoe UI', 11))
        self.count.pack(fill='x', pady=(0, 10))
        table = ctk.CTkFrame(self, corner_radius=8)
        table.pack(fill='both', expand=True)
        table.grid_columnconfigure(0, weight=1)
        table.grid_rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(table, columns=self.columns, show='headings',
                                  selectmode='extended', style='Tracking.Treeview')
        for key, title, width, minimum in zip(self.columns,
                ('DATE', 'HEURE', 'ROBOT', 'EMPLACEMENT', 'TRACE'),
                (108, 110, 90, 145, 230), (90, 85, 60, 100, 130)):
            self.tree.heading(key, text=title, anchor='w' if key in ('location', 'trace') else 'center')
            self.tree.column(key, width=width, minwidth=minimum,
                             anchor='w' if key in ('location', 'trace') else 'center', stretch=key == 'trace')
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
        self.detail_event = ctk.CTkLabel(self.detail, text='', anchor='w', justify='left',
                                         font=('Consolas', 10), wraplength=600)
        self.detail_event.pack(fill='x', padx=12, pady=(2, 8))
        for widget in (self.detail_title, self.detail_source, self.detail_event):
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

    def _resize_detail(self, event):
        for widget in (self.detail_title, self.detail_source, self.detail_event):
            widget.configure(wraplength=max(150, event.width - 28))

    def apply_theme(self):
        colors = self.app._palette
        self.tree.tag_configure('even', background=colors['surface'], foreground=colors['text'])
        self.tree.tag_configure('odd', background=colors['bg'], foreground=colors['text'])

    def _cancel(self):
        if self._pending:
            self.after_cancel(self._pending)
            self._pending = None

    def _destroyed(self, event):
        if event.widget is self:
            self._cancel()

    def begin(self, reference):
        self._cancel()
        self.reference = reference
        self.rows.clear()
        self.values.clear()
        self.tree.delete(*self.tree.get_children())
        self.heading.configure(text='Dépôts · ' + reference if reference else 'Suivi des dépôts')
        self.count.configure(text='Recherche en cours…' if reference else 'Saisissez une référence.')
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
                                         if count else 'Aucun dépôt trouvé pour cette référence.')
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
            title, source, text = 'Sélectionnez un dépôt pour consulter le détail.', '', ''
        else:
            date, hour, robot, location, _ = tracking_values(row, self.app._selected_directories)
            title = f'{location}   ·   {robot}   ·   {date} à {hour}'
            source = self.app._compact_trace_label(row['path'])
            text = row['text']
        for widget, value in zip((self.detail_title, self.detail_source, self.detail_event), (title, source, text)):
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
        self._cell = (key, int(column[1:]) - 1)
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
