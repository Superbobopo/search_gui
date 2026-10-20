"""Persistent robot tables, one shared scroll and exact occurrence targets."""
from datetime import datetime
import json
import tkinter as tk
from tkinter import ttk

from selectable_text import make_selectable
from ui_components import ActionButton, font as ui_font
from popup_style import configure_popup_style, PopupTreeview
from trace_parser import timestamp_parts
from trace_sources import compact_trace_label, trace_tooltip, trace_label


def gap_text(seconds):
    if seconds is None:
        return ''
    if not seconds:
        return 'même instant'
    days, remainder = divmod(abs(seconds), 86400)
    hours, remainder = divmod(remainder, 3600)
    minutes, remainder = divmod(remainder, 60)
    parts = ([f'{int(days)} j'] if days else []) + ([f'{int(hours)} h'] if hours else [])
    if minutes:
        parts.append(f'{int(minutes)} min')
    if remainder or not parts:
        parts.append(f'{remainder:.1f}'.rstrip('0').rstrip('.') + ' s')
    return ' '.join(parts) + (' après' if seconds > 0 else ' avant')


def listing_text(data, directories):
    """Plain text export of the published result, never a second lookup."""
    blocks = []
    columns = data.get('columns', [])
    for col in sorted(columns, key=lambda c: {'R1': 0, 'R2': 1}.get(c['robot'], 2)):
        anchor = col.get('anchor')
        lines = [col['robot']]
        if not anchor:
            lines.append(col.get('message') or ('Rapprochement ambigu : choisir une génération.'
                         if col.get('candidates') else 'Aucun rapprochement disponible.'))
        else:
            when = anchor.get('time')
            lines.append(datetime.fromtimestamp(when).strftime('%d/%m/%Y') if when is not None else 'Date inconnue')
            for row in [*col.get('previous', []), anchor, *col.get('following', [])]:
                central = row['path']==anchor['path'] and row['number']==anchor['number']
                hour = timestamp_parts(row['text'])[1] or '—'
                values = [str(row.get(key) or '—') for key in ('modbus', 'ref', 'sn', 'lot')]
                lines.append(('🟢 ➜ ' if central else '') + hour + ' | ' + ' | '.join(
                    f'{name} : {value}' for name, value in zip(('ID', 'Réf', 'SN', 'Lot'), values)))
        blocks.append('\n'.join(lines))
    return '\n\n'.join(blocks)


class AdjacentView(ttk.Frame):
    fields = ('hour', 'id', 'ref', 'sn', 'lot', 'state')

    def __init__(self, parent, app):
        super().__init__(parent, padding=(20, 16))
        self._theme_deferred = True  # This view owns its surface and semantic colors.
        self.app, self.data = app, None
        self.trees, self.rows, self.panels = {}, {}, []
        self._updating = False
        self._scrolling = False
        self._hover_key = None
        self._cell = None
        self._wrap_width = None
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        header = ttk.Frame(self)
        header.grid(row=0, column=0, sticky='ew', pady=(0, 14))
        header.grid_columnconfigure(1, weight=1)
        self.title = tk.Label(header, text='IDs adjacents', anchor='w', font=ui_font('title', True, native=True))
        self.title.grid(row=0, column=0, sticky='w')
        self.heading = tk.Label(header, anchor='w', font=('Segoe UI', 10), justify='left')
        self.heading.grid(row=0, column=1, sticky='ew', padx=(24, 0))
        make_selectable(self.heading)
        controls = ttk.Frame(self, style='Popup.TFrame', padding=(14, 10))
        controls.grid(row=1, column=0, sticky='ew', pady=(0, 10))
        self._heading_controls = ttk.Frame(controls, style='Popup.TFrame')
        self._heading_controls.grid(row=0, column=0, sticky='w')
        self._count_controls = ttk.Frame(controls, style='Popup.TFrame')
        self._count_controls.grid(row=0, column=1, sticky='w', padx=(22, 0))
        ttk.Label(self._heading_controls, text='Voisinage', style='Popup.TLabel',
                  font=('Segoe UI', 10, 'bold')).grid(row=0, column=0, padx=(0, 16))
        value = getattr(app, '_adjacent_counts', (5, 5))
        self.before_var, self.after_var = tk.StringVar(value=str(value[0])), tk.StringVar(value=str(value[1]))
        ttk.Label(self._count_controls, text='Avant', style='Popup.TLabel').grid(row=0, column=0, padx=(0, 6))
        self.before_entry = ttk.Entry(self._count_controls, textvariable=self.before_var, width=6)
        self.before_entry.grid(row=0, column=1)
        ttk.Label(self._count_controls, text='Après', style='Popup.TLabel').grid(row=0, column=2, padx=(14, 6))
        self.after_entry = ttk.Entry(self._count_controls, textvariable=self.after_var, width=6)
        self.after_entry.grid(row=0, column=3)
        self.apply_button = ActionButton(self._count_controls, text='Appliquer', style='Popup.TButton', command=self.apply_counts)
        self.apply_button.grid(row=0, column=4, padx=(14, 0))
        self.copy_button = ActionButton(self._count_controls, text='⧉', width=32, role='quiet',
                                        command=self.copy_lists, state='disabled')
        self.copy_button.grid(row=0, column=5, padx=(6, 0))
        self.copy_button.bind('<Enter>', self._show_copy_tip, add='+')
        self.copy_button.bind('<Leave>', self._hide_copy_tip, add='+')
        self._copy_tip = None
        self._copy_ready = False
        self.validation = tk.Label(controls, text='', anchor='w', font=('Segoe UI', 9))
        self.validation.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(6, 0))
        self.validation.grid_remove()
        for entry in (self.before_entry, self.after_entry):
            entry.bind('<Return>', lambda event: self.apply_counts())
        self.note = ttk.Label(self, text='', font=('Segoe UI', 9))
        self.note.grid(row=2, column=0, sticky='ew', pady=(0, 12))
        # Horizontal overflow only. Vertical scrolling is shared by the two
        # fixed-height native tables, so thousands of rows need no giant canvas.
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self.canvas.grid(row=3, column=0, sticky='nsew')
        self.vertical = ttk.Scrollbar(self, orient='vertical', style='Popup.Vertical.TScrollbar', command=self._yview)
        self.vertical.grid(row=3, column=1, sticky='ns')
        self.horizontal = ttk.Scrollbar(self, orient='horizontal', style='Popup.Horizontal.TScrollbar', command=self.canvas.xview)
        self.horizontal.grid(row=4, column=0, sticky='ew')
        self.canvas.configure(xscrollcommand=self._xscroll)
        self.content = ttk.Frame(self.canvas)
        self.content.grid_rowconfigure(0, weight=1)
        self.window = self.canvas.create_window((0, 0), window=self.content, anchor='nw')
        self.canvas.bind('<Configure>', self._canvas_size)
        self.detail_card = ttk.Frame(self, style='Popup.TFrame', padding=(14, 10))
        self.detail_card.grid(row=5, column=0, sticky='ew', pady=(12, 0))
        self.detail_card.grid_columnconfigure(0, weight=1)
        self.detail = tk.Label(self.detail_card, text='', height=2, anchor='nw', justify='left', font=('Segoe UI', 10))
        self.detail.grid(row=0, column=0, sticky='ew')
        make_selectable(self.detail)
        self.get_detail = tk.Label(self.detail_card, text='', anchor='w', justify='left', font=('Segoe UI', 10, 'bold'))
        self.get_detail.grid(row=1, column=0, sticky='ew', pady=(4, 0))
        make_selectable(self.get_detail)
        self.close_button = ActionButton(self, text='Fermer', style='Popup.TButton')
        self.close_button.grid(row=6, column=0, sticky='e', pady=(12, 0))
        for index in range(2):
            self._make_panel(index)
        self.menu = tk.Menu(self, tearoff=False)
        for title, kind in (('Copier la cellule', 'cell'), ('Copier la ligne', 'row'),
                            ('Copier la provenance complète', 'source')):
            self.menu.add_command(label=title, command=lambda value=kind: self._copy(value))
        self.bind('<Configure>', self._resize)
        self.status('Recherche des voisins…')
        self.apply_theme()

    def _make_panel(self, index):
        self.content.grid_columnconfigure(index, weight=1, uniform='robots')
        frame = ttk.Frame(self.content, style='Popup.TFrame', padding=12)
        frame.grid(row=0, column=index, sticky='nsew', padx=(0, 8) if index==0 else (8, 0))
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(1, weight=1)
        meta = ttk.Frame(frame, style='Popup.TFrame')
        meta.grid(row=0, column=0, sticky='ew', pady=(0, 10))
        meta.grid_columnconfigure(1, weight=1)
        title = ttk.Label(meta, style='Popup.TLabel', font=ui_font('section', True, native=True))
        title.grid(row=0, column=0, sticky='nw', padx=(0, 14))
        status = tk.Label(meta, anchor='nw', justify='left', wraplength=430, height=2, font=('Segoe UI', 9))
        status.grid(row=0, column=1, sticky='ew', pady=(4, 0))
        source = tk.Label(meta, anchor='nw', justify='left', wraplength=490, height=1, font=('Segoe UI', 9))
        source.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(4, 0))
        for label in (status, source):
            make_selectable(label)
            label.selection_text.bind('<MouseWheel>', self._wheel)
        tree = PopupTreeview(frame, columns=self.fields, show='headings', height=11,
                                 selectmode='none', style='Adjacent.Treeview')
        for field, caption, width in zip(self.fields, ('Heure', 'ID', 'Référence', 'SN', 'CodeLot', 'État'),
                                        (80, 45, 110, 75, 95, 120)):
            tree.heading(field, text=caption)
            tree.column(field, width=width, minwidth=width-10, anchor='center')
        tree.grid(row=1, column=0, sticky='nsew')
        footer = ttk.Label(frame, style='Popup.TLabel', font=('Segoe UI', 9))
        footer.grid(row=2, column=0, sticky='w', pady=(10, 0))
        picker = ttk.Combobox(frame, state='readonly', width=50)
        picker.grid(row=3, column=0, sticky='ew')
        picker.grid_remove()
        panel = dict(frame=frame, title=title, status=status, source=source, tree=tree,
                     footer=footer, picker=picker, robot='', candidates=[], warning=False)
        self.panels.append(panel)
        picker.bind('<<ComboboxSelected>>', lambda event, p=panel: self._pick(p))
        tree.configure(yscrollcommand=lambda first, last, p=panel: self._scrolled(p, first, last))
        tree.bind('<MouseWheel>', self._wheel)
        tree.bind('<ButtonRelease-1>', lambda event, p=panel: self._open(p['robot'], event))
        tree.bind('<Return>', lambda event, p=panel: self._open(p['robot']))
        tree.bind('<Control-c>', lambda event, p=panel: self._copy_selected(p['robot']))
        tree.bind('<Button-3>', lambda event, p=panel: self._context(p['robot'], event))
        tree.bind('<Motion>', lambda event, p=panel: self._hover(p['robot'], event))

    def _canvas_size(self, event):
        self.canvas.itemconfigure(self.window, width=max(1040, event.width), height=event.height)
        self.canvas.configure(scrollregion=(0, 0, max(1040, event.width), event.height))

    def _xscroll(self, first, last):
        self.horizontal.set(first, last)
        if float(first)<=0 and float(last)>=1:
            self.horizontal.grid_remove()
        else:
            self.horizontal.grid()

    def _resize(self, event):
        if event.widget is self:
            width = max(150, event.width-32)
            if width != self._wrap_width:
                self._wrap_width = width
                self.heading.configure(wraplength=max(150, width-200))
                for label in (self.note, self.detail, self.get_detail):
                    label.configure(wraplength=width)
                narrow = event.width < 860
                self._count_controls.grid_configure(row=1 if narrow else 0, column=0 if narrow else 1,
                                                     padx=0 if narrow else (22, 0), pady=(8, 0) if narrow else 0)

    def apply_counts(self):
        try:
            counts = int(self.before_var.get().strip()), int(self.after_var.get().strip())
            self.app.store._neighbor_counts(*counts)
        except ValueError:
            self.validation.configure(text='Saisissez deux entiers entre 0 et 1 000.')
            self.validation.grid()
            return
        self.validation.configure(text='')
        self.validation.grid_remove()
        self.app._adjacent_counts = counts
        self.app.open_adjacent_for_current_occurrence()

    def _yview(self, *args):
        self._scrolling = True
        try:
            for panel in self.panels:
                panel['tree'].yview(*args)
            self.vertical.set(*self.panels[0]['tree'].yview())
        finally:
            self._scrolling = False

    def _scrolled(self, panel, first, last):
        if not self._updating and not self._scrolling:
            for other in self.panels:
                if other is not panel and abs(other['tree'].yview()[0]-float(first)) > .00001:
                    other['tree'].yview_moveto(first)
            self.vertical.set(first, last)

    def _wheel(self, event):
        self._yview('scroll', -int(event.delta/120), 'units')
        return 'break'

    def _show_copy_tip(self, event=None):
        self._hide_copy_tip()
        self._copy_tip = tk.Toplevel(self)
        self._copy_tip.wm_overrideredirect(True)
        self._copy_tip.wm_geometry(f'+{self.copy_button.winfo_rootx()}+{self.copy_button.winfo_rooty()+36}')
        ttk.Label(self._copy_tip, text='Copier les listes R1 et R2', padding=6).pack()

    def _hide_copy_tip(self, event=None):
        if self._copy_tip is not None:
            self._copy_tip.destroy()
            self._copy_tip = None

    def copy_lists(self):
        if not self._copy_ready or not self.data:
            return
        text = listing_text(self.data, self.app._selected_directories)
        if text:
            self.clipboard_clear()
            self.clipboard_append(text)
            self.note.configure(text='Listes R1 et R2 copiées')
        self._hide_copy_tip()

    def loading(self, message='Actualisation des voisins…'):
        self._copy_ready = False
        self.copy_button.configure(state='disabled')
        self.note.configure(text=message)

    def status(self, message):
        self._copy_ready = False
        self.copy_button.configure(state='disabled')
        self.data = None
        self.rows.clear()
        self._cell = self._hover_key = None
        self.heading.configure(text=message)
        self.heading.sync_selection_text()
        self.note.configure(text='')
        self.detail.configure(text='')
        self.detail.sync_selection_text()
        self.get_detail.configure(text='')
        self.get_detail.sync_selection_text()
        self._updating = True
        for panel in self.panels:
            tree = panel['tree']
            if tree.get_children():
                tree.delete(*tree.get_children())
            for key in ('title', 'status', 'source', 'footer'):
                panel[key].configure(text='')
            panel['picker'].grid_remove()
            for key in ('status', 'source'):
                panel[key].sync_selection_text()
        self._updating = False

    def show(self, data):
        self._copy_ready = any(col.get('anchor') for col in data['columns'])
        self.copy_button.configure(state='normal' if self._copy_ready else 'disabled')
        previous = self.data
        self.data = data
        self._cell = self._hover_key = None
        self.trees.clear()
        self.rows.clear()
        selected = data['selected']
        name = data['selected_robot'] or 'Trace sélectionnée'
        when = datetime.fromtimestamp(data['time']).strftime('%d/%m/%Y %H:%M:%S.%f').rstrip('0').rstrip('.') if data['time'] is not None else 'horaire inexploitable'
        self.heading.configure(text=f'Repère : {when} · {name} · ID {selected["modbus"]}')
        self.heading.sync_selection_text()
        self.note.configure(text='Rouge : écart IDPK > 2 · Orange : échec GET confirmé · Voisinages indépendants.'
                            + (' · Vérification en cours : alertes provisoires.' if not self.app._sources_verified else ''))
        columns = data['columns']
        verified = self.app._sources_verified and self.app._indexed_complete
        before = max((len(col.get('previous', [])) for col in columns), default=0)
        after = max((len(col.get('following', [])) for col in columns), default=0)
        size = max(before+1+after, *(len(col.get('candidates', [])) for col in columns))
        self._updating = True
        try:
            for panel, col in zip(self.panels, columns):
                robot = panel['robot'] = col['robot']
                tree = self.trees[robot] = panel['tree']
                self.rows[robot] = {}
                panel['title'].configure(text=robot)
                anchor = col.get('anchor')
                delta = col.get('delta')
                warning = panel['warning'] = delta is not None and abs(delta)>300
                message = col.get('message', '')
                if anchor:
                    date = datetime.fromtimestamp(anchor['time']).strftime('%d/%m/%Y %H:%M:%S')
                    message = message or (date + (' · sélection actuelle' if robot==name else ' · '+gap_text(delta)))
                if warning:
                    message += '\nAucune occurrence à moins de 5 minutes dans les traces disponibles.'
                panel['status'].configure(text=message)
                source = compact_trace_label(anchor['path'], self.app._selected_directories) if anchor else ''
                if anchor and getattr(anchor['path'], 'historical', False):
                    source += ' · Historique'
                panel['source'].configure(text=source)
                for key in ('status', 'source'):
                    panel[key].sync_selection_text()
                candidates = panel['candidates'] = col.get('candidates', [])
                if candidates:
                    panel['picker'].configure(values=[trace_label(row['path']) for row in candidates])
                    panel['picker'].set('Choisir une génération…')
                    panel['picker'].grid()
                else:
                    panel['picker'].grid_remove()
                if anchor:
                    entries = [None]*(before-len(col['previous'])) + col['previous'] + [anchor] + col['following']
                else:
                    entries = list(candidates)
                entries += [None]*(size-len(entries))
                keep = set()
                for i, row in enumerate(entries):
                    key = str(i)
                    keep.add(key)
                    diagnostic = row.get('get_diagnostic', {}) if row else {}
                    get_failed = verified and diagnostic.get('failed')
                    different = verified and row and row.get('different')
                    caption = {'failed': 'Échec GET', 'recovered': 'Échec GET',
                               'success': 'GET réussi', 'unknown': 'GET inconnu', 'missing': 'GET absent'}
                    state = (('IDPK + échec GET' if different and get_failed else caption.get(diagnostic.get('status'), 'GET inconnu'))
                             if verified else 'Vérification en cours')
                    values = ((timestamp_parts(row['text'])[1] or '—', row['modbus'], row['ref'], row['sn'] or '—', row['lot'] or '—', state)
                              if row else ('', '', '', '', '', ''))
                    central = row is not None and anchor is not None and row['path']==anchor['path'] and row['number']==anchor['number']
                    background = ('current' if robot==name else 'nearest') if central else ('even' if i%2==0 else 'odd')
                    tag = background + ('_different' if different else '_get_failed' if get_failed else '')
                    if tree.exists(key):
                        if tuple(tree.item(key, 'values')) != tuple(values) or tree.item(key, 'tags') != (tag,):
                            tree.item(key, values=values, tags=(tag,))
                    else:
                        tree.insert('', 'end', iid=key, values=values, tags=(tag,))
                    if row:
                        self.rows[robot][key] = row
                remove = [key for key in tree.get_children() if key not in keep]
                if remove:
                    tree.delete(*remove)
                if tree.selection():
                    tree.selection_remove(*tree.selection())
                panel['footer'].configure(text=(f'{len(col.get("previous", []))} précédents / {data["before_count"]} demandés'
                    f' · {len(col.get("following", []))} suivants / {data["after_count"]} demandés') if anchor else '')
            self.apply_theme()
        finally:
            self._updating = False
        # Keep the repère visible when a requested range grows substantially.
        if previous is None or previous.get('before_count') != data['before_count'] or previous.get('after_count') != data['after_count']:
            self._yview('moveto', max(0, (before-4)/max(1, size)))
        for panel in self.panels:
            if panel['robot']==name:
                row = columns[self.panels.index(panel)].get('anchor')
                if row:
                    self._detail(row)

    def apply_theme(self):
        colors = self.app._palette
        configure_popup_style(self, colors)
        self.canvas.configure(background=colors['bg'])
        self.title.configure(background=colors['bg'], foreground=colors['text'])
        self.heading.configure(background=colors['bg'], foreground=colors['muted'])
        self.heading.sync_selection_text()
        self.note.configure(foreground=colors['muted'])
        self.get_detail.configure(background=colors['surface'], foreground=colors['warning_text'])
        self.get_detail.sync_selection_text()
        for label in (self.detail, self.validation):
            label.configure(background=colors['surface'], foreground=colors['error_text'] if label is self.validation else colors['text'])
        self.detail.sync_selection_text()
        for panel in self.panels:
            panel['title'].configure(foreground=colors['accent'])
            for label, role in ((panel['status'], 'warning_text' if panel['warning'] else 'muted'), (panel['source'], 'text')):
                label.configure(background=colors['surface'], foreground=colors[role])
                label.sync_selection_text()
            for tag, background in (('even','surface'), ('odd','bg'), ('current','positive_bg'), ('nearest','info_bg')):
                for different in (False, True):
                    panel['tree'].tag_configure(tag + ('_different' if different else ''), background=colors[background],
                                                foreground=colors['error_text'] if different else colors['text'])
                panel['tree'].tag_configure(tag+'_get_failed', background=colors[background], foreground=colors['warning_text'])

    def _detail(self, row):
        result = row.get('comparison', {})
        message = result.get('message', '')
        pair = result.get('pair')
        if pair:
            def dims(event):
                values = json.loads(event['dims'])
                return ' / '.join(str(values.get(key, '—')) for key in ('l','H','L'))
            other = pair[1]
            date = datetime.fromtimestamp(other['time']).strftime('%d/%m/%Y %H:%M:%S')
            message = dims(pair[0]) + ' contre ' + dims(other) + ' · ' + result['basis_label'] + f' · {date} · ' + compact_trace_label(other['path'], self.app._selected_directories)
        if not self.app._sources_verified or not self.app._indexed_complete:
            message = 'Comparaison disponible après vérification des sources.'
        self.detail.configure(text=f'ID {row["modbus"]} · {row["ref"]} · SN {row["sn"] or "—"} · Lot {row["lot"] or "—"}\n' + message)
        self.detail.sync_selection_text()
        verified = self.app._sources_verified and self.app._indexed_complete
        self.get_detail.configure(text=row.get('get_diagnostic', {}).get('message', 'Résultat GET inconnu') if verified else '')
        self.get_detail.sync_selection_text()

    def _hover(self, robot, event):
        if not self.data or robot not in self.trees:
            return
        key = self.trees[robot].identify_row(event.y)
        row = self.rows.get(robot, {}).get(key)
        if row and self._hover_key != (robot,key):
            self._hover_key = robot,key
            self._detail(row)

    def _pick(self, panel):
        index = panel['picker'].current()
        if 0 <= index < len(panel['candidates']):
            self._navigate(panel['candidates'][index])

    def _navigate(self, row):
        if not self.data:
            return
        selected = self.app.selection
        if selected and selected.path==row['path'] and selected.number==row['number']:
            return
        self.app._on_adjacent_ref_click(row['modbus'], row['ref'], row['path'], row['number'])

    def _open(self, robot, event=None):
        if robot not in self.trees:
            return 'break'
        tree = self.trees[robot]
        key = tree.identify_row(event.y) if event else tree.focus()
        row = self.rows.get(robot, {}).get(key)
        if row and (event is None or tree.identify_region(event.x,event.y)=='cell'):
            self._navigate(row)
        return 'break'

    def _copy_selected(self, robot):
        key = self.trees[robot].focus()
        if key:
            self._cell = robot,key,None
            self._copy('row')
        return 'break'

    def _context(self, robot, event):
        tree = self.trees[robot]
        key = tree.identify_row(event.y)
        if key in self.rows.get(robot, {}):
            tree.focus(key)
            self._cell = robot,key,tree.identify_column(event.x)
            try:
                self.menu.tk_popup(event.x_root,event.y_root)
            finally:
                self.menu.grab_release()
        return 'break'

    def _copy(self, kind):
        if not self._cell:
            return
        robot,key,column = self._cell
        row = self.rows.get(robot, {}).get(key)
        if not row:
            return
        values = self.trees[robot].item(key,'values')
        text = (trace_tooltip(row['path']) + f'\nLigne {row["number"]+1}\n' + row['text'] if kind=='source' else
                values[int(column[1:])-1] if kind=='cell' and column and column!='#0' else '\t'.join(values))
        self.clipboard_clear()
        self.clipboard_append(text)
