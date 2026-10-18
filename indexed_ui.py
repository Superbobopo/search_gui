"""Indexed actions for the existing Tk interface; workers never touch widgets."""

from datetime import datetime
import json
import os
from queue import Queue, Empty
import re
import sqlite3
import threading
import time
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk

from background_tasks import BackgroundTasks
from trace_parser import format_ticket_fields, parse_fields, timestamp_parts
from trace_sources import discover_trace_files, trace_label
from trace_store import TraceStore, Cancelled, CHECKPOINTS
from app_assets import apply_icon
from selection_state import OccurrenceSelection
from ui_theme import ManagedTextbox


class IndexedAppMixin:
    def _start_indexed_service(self):
        self.tasks = BackgroundTasks()
        self._index_messages = Queue()
        self._stop_index = threading.Event()
        self._groups = []
        self._group_index = 0
        self._selection_numbers = []
        self.selection = None
        self._selection_serial = 0
        self.btn_adjacent.configure(state='disabled')
        self._view_generation = 0
        self._indexed_complete = False
        self._query_serial = 0
        self.status = tk.StringVar(value='Prêt pour une recherche')
        self.index_status = tk.StringVar(value='Ouverture de l’index…')
        if hasattr(self, '_install_status_widgets'):
            self._install_status_widgets()
        else:
            for variable in (self.status, self.index_status):
                tk.Label(self, textvariable=variable, anchor='w', bg='#f7f9fa', fg='#405060').pack(side='bottom', fill='x')
        self.protocol('WM_DELETE_WINDOW', self._close_indexed_app)
        apply_icon(self)
        self.bind('<Map>', self._icon_on_map, add='+')
        self.after(250, lambda: apply_icon(self))
        self._index_thread = threading.Thread(target=self._index_loop, daemon=True, name='trace-index')
        self._index_thread.start()
        self.after(30, self._poll_indexed_service)

    def _icon_on_map(self, event):
        if isinstance(event.widget, (tk.Toplevel, tk.Tk)):
            apply_icon(event.widget)

    def _index_loop(self):
        try:
            store = TraceStore()
            try:
                metadata = store.metadata()
            except sqlite3.DatabaseError:
                if not store.corrupt.is_set():
                    raise
                store.rebuild(); metadata = []
            self._index_messages.put(('ready', store, metadata, store.initialized()))
            while not self._stop_index.is_set():
                if store.corrupt.is_set():
                    self._index_messages.put(('rebuilding',))
                    store.rebuild()
                    self._index_messages.put(('ready', store, [], False))
                paths = discover_trace_files()
                def progress(done, total, path, published):
                    self._index_messages.put(('progress', done, total, path, published,
                                              store.metadata() if published else None))
                try:
                    result = store.sync(paths, progress, self._stop_index)
                except sqlite3.DatabaseError:
                    if store.corrupt.is_set():
                        continue
                    raise
                self._index_messages.put(('complete', result, store.metadata()))
                self._stop_index.wait(2)
        except Cancelled:
            pass
        except Exception as error:
            self._index_messages.put(('error', error))

    def _poll_indexed_service(self):
        if self._stop_index.is_set():
            return
        try:
            while True:
                item = self._index_messages.get_nowait()
                if item[0] == 'ready':
                    self.store = item[1]
                    self._indexed_complete = item[3]
                    self._install_metadata(item[2])
                elif item[0] == 'progress':
                    _, done, total, path, published, metadata = item
                    self.index_status.set(f'Indexation {done}/{total} : {os.path.basename(path)} — résultats partiels'
                                    if not self._indexed_complete else f'Actualisation {done}/{total} : {os.path.basename(path)}')
                    if published:
                        self._install_metadata(metadata)
                        if getattr(self, '_last_search', None):
                            self._launch_search(*self._last_search)
                elif item[0] == 'complete':
                    changed = item[1]['changed'] or item[1]['removed']
                    self._install_metadata(item[2])
                    was_complete = self._indexed_complete
                    self._indexed_complete = was_complete or not bool(item[1]['failures'])
                    self.files_loaded = self._indexed_complete
                    failures = item[1]['failures']
                    if failures:
                        suffix = 'actualisation incomplète' if was_complete else 'résultats partiels'
                        self.index_status.set(f'{len(self.file_list)} fichiers prêts — {suffix}, '
                                              f'{len(failures)} fichier(s) à réessayer')
                    else:
                        self.index_status.set(f'{len(self.file_list)} fichiers prêts')
                    if getattr(self, 'btn_erreurs', None) and self.btn_erreurs.winfo_exists():
                        self.btn_erreurs.configure(state='normal' if self._indexed_complete else 'disabled')
                    self.title('Recherche TracOmega')
                    if (changed or not was_complete) and getattr(self, '_last_search', None):
                        self._launch_search(*self._last_search)
                elif item[0] == 'error':
                    self.index_status.set(f'Index indisponible : {item[1]}')
                elif item[0] == 'rebuilding':
                    self.index_status.set('Reconstruction de l’index local…')
                    self._indexed_complete = False
                    self._invalidate_view()
        except Empty:
            pass
        try:
            self.tasks.drain()
        finally:
            self.after(30, self._poll_indexed_service)

    def _install_metadata(self, metadata):
        self.file_list = [row['path'] for row in metadata]
        self._file_metadata = {row['path']: row for row in metadata}
        if self.selection and self.selection.path not in self._file_metadata:
            self._clear_selection()
            self._groups = []
            if self.file_label.winfo_exists():
                self.file_label.configure(text='Fichier : Aucun')
        self.files_loaded = self._indexed_complete

    def _submit(self, channel, work, done):
        if not hasattr(self, 'store'):
            self.status.set('Ouverture de l’index en cours…')
            return
        def compute(cancel):
            with self.store.snapshot():
                with self.store.connection() as db:
                    db.set_progress_handler(lambda: int(cancel.is_set()), 1000)
                    return work(cancel)
        def failed(error):
            self.status.set(f'Opération interrompue : {error}')
        view, serial = self._view_generation, self._selection_serial
        def apply(result):
            if view != self._view_generation:
                return
            if channel in ('details', 'adjacent', 'comparison', 'file_view') and serial != self._selection_serial:
                return
            done(result)
        return self.tasks.submit(channel, compute, apply, failed)

    @property
    def _display_data(self):
        return self.selection.details if self.selection else None

    def _begin_selection_change(self):
        self._selection_serial += 1
        self.selection = None
        self.btn_adjacent.configure(state='disabled')
        for channel in ('details', 'adjacent', 'comparison', 'file_view'):
            self.tasks.cancel(channel)
        self._clear_adjacent('Recherche en cours…')
        if hasattr(self, '_update_context_header'):
            self._update_context_header()
        self._render_summary(None)

    def _clear_adjacent(self, message='Aucune occurrence sélectionnée.'):
        popup = getattr(self, 'adj_popup', None)
        if popup is not None and popup.winfo_exists():
            for frame in (self.prev_frame, self.current_frame, self.next_frame):
                for widget in frame.winfo_children():
                    widget.destroy()
            tk.Label(self.current_frame, text=message).pack(padx=5, pady=5)

    def _clear_selection(self):
        self._begin_selection_change()
        self._groups = []
        self._group_index = 0
        self.lines = []
        self.line_index = 0
        self._selection_numbers = []
        self._clear_adjacent()

    def _invalidate_view(self):
        self._view_generation += 1
        self._query_serial += 1
        self._last_search = None
        self._clear_selection()
        for channel in list(self.tasks.latest):
            self.tasks.cancel(channel)

    def _close_indexed_app(self):
        self._stop_index.set()
        self.tasks.close()
        self.destroy()

    def _diagnostic_ready(self):
        if not self._indexed_complete:
            messagebox.showinfo('Indexation en cours', 'Ce contrôle sera disponible après l’indexation complète des traces.')
            return False
        return True

    def perform_search(self):
        raw = self.keyword.get().strip()
        if not raw:
            messagebox.showwarning('Attention', 'Entrez un mot-clé ou un numéro valide !')
            return
        mode = 'id' if self.modbus.get() else 'ref' if self.reference.get() else 'text'
        self.btn_compare_idpk.configure(state='normal' if mode == 'ref' else 'disabled')
        self._launch_search(raw, mode, None, None)

    def recherche_avancee_criteria(self, ref, sn=None, codelot=None):
        self._launch_search(ref, 'advanced', sn, codelot)

    def _launch_search(self, raw, mode, sn, lot, preferred=None):
        self._last_search = (raw, mode, sn, lot, preferred)
        self._query_serial += 1
        serial = self._query_serial
        self._begin_selection_change()
        self._groups = []
        self.lines = []
        self._selection_numbers = []
        self.status.set('Recherche…' + (' résultats partiels' if not self._indexed_complete else ''))
        def search(cancel):
            groups = self.store.search('ModBusId=' + raw if mode == 'id' else 'RefProduit=' + raw if mode == 'ref' else raw,
                                       mode == 'id', ref=raw if mode == 'advanced' else None, sn=sn, lot=lot, cancel=cancel)
            index, number = 0, groups[0][1][0][0] if groups else None
            if preferred:
                for i, (path, hits) in enumerate(groups):
                    if path == preferred.path and any(n == preferred.number for n, _ in hits):
                        index, number = i, preferred.number
                        break
            details = self.store.details(groups[index][0], number) if groups else None
            return groups, index, details
        def apply(result):
            if serial != self._query_serial:
                return
            self._groups, self._group_index, details = result
            self._active_mode = mode
            self._apply_selection(details)
            self.status.set(('Résultats partiels — indexation en cours' if not self._indexed_complete else
                             f'{len(self._groups)} fichier(s) correspondant(s)'))
        self._submit('search', search, apply)

    def _apply_selection(self, details):
        if not details:
            self._clear_selection()
            self.txt.configure(state='normal'); self.txt.delete('1.0', 'end')
            self.txt.insert('end', 'Aucun résultat dans les fichiers déjà indexés.\n' if not self._indexed_complete else 'Aucun fichier trouvé avec ce critère.\n')
            self.txt.configure(state='disabled'); self.file_label.configure(text='Fichier : Aucun')
            self._render_summary(None)
            return
        path, hits = self._groups[self._group_index]
        self._selection_numbers = [number for number, text in hits]
        self.lines = [text for number, text in hits]
        self.file_index = self.file_list.index(path) if path in self.file_list else 0
        self.line_index = self._selection_numbers.index(details['row']['number'])
        self._render_indexed_details(details)

    def display_line(self):
        if not self._groups or not self.lines or not self._selection_numbers:
            return
        path, number = self._groups[self._group_index][0], self._selection_numbers[self.line_index]
        self._begin_selection_change()
        self._submit('details', lambda cancel: self.store.details(path, number), self._render_indexed_details)

    def _render_indexed_details(self, data):
        if not data:
            self._apply_selection(None)
            return
        if not self.txt.winfo_exists():
            return
        self.selection = OccurrenceSelection.from_details(data)
        text = self.txt
        text.configure(state='normal'); text.delete('1.0', 'end')
        text.tag_configure('date', font=('Segoe UI', 10), foreground='#909090', spacing3=0, justify='center')
        text.tag_configure('occ', font=('Segoe UI', 9), foreground='#adb5bd', spacing3=2, justify='center')
        text.tag_configure('divider', font=('Segoe UI', 9), foreground='#c7c7c7', spacing3=1)
        text.tag_configure('brute', font=('Consolas', 9), background='#f5f7fa', foreground='#1e293b', spacing3=2, lmargin1=7, lmargin2=7)
        for tag, color, background in [('getcard', '#00589b', '#e3f2fd'), ('putecard', '#b66800', '#fff4e3'), ('lastpute', '#9f8700', '#fffde3')]:
            text.tag_configure(tag, font=('Segoe UI', 10, 'bold'), foreground=color, background=background, spacing1=4, spacing3=8, lmargin1=10, lmargin2=10)
        text.tag_configure('none', font=('Segoe UI', 9, 'italic'), foreground='#a3a3a3', spacing3=2, justify='center')
        text.tag_configure('refcard', font=('Segoe UI', 10, 'bold'), foreground='#00876c', background='#e9fdf4', spacing1=6, spacing3=6, justify='center', lmargin1=6, lmargin2=6)
        for tag, color in [('clickable_get', '#00589b'), ('clickable_pute', '#b66800'), ('clickable_refprod', '#00876c')]:
            text.tag_configure(tag, underline=1, foreground=color)
        row = data['row']; when = datetime.fromtimestamp(row['time'])
        text.insert('end', f'{when:%d/%m/%Y}\n', 'date')
        text.insert('end', f'Occurrence {self.line_index + 1}/{len(self.lines)}\n', 'occ')
        text.insert('end', '─' * 56 + '\n', 'divider')
        formatted, fields = format_ticket_fields(row['text'])
        text.insert('end', formatted + '\n', 'brute')
        text.insert('end', '─' * 56 + '\n', 'divider')
        if fields['refprod']:
            text.insert('end', f"RefProduit trouvé : {fields['refprod']}\n", ('refcard', 'clickable_refprod'))
            text.insert('end', '─' * 56 + '\n', 'divider')
        for name, title, tag in [('get', 'GRIPPER GET', 'getcard'), ('pute', 'GRIPPER PUTE', 'putecard')]:
            event = data[name]
            if event:
                text.insert('end', f"{title} :\n{event['text']}\n", (tag, 'clickable_' + name))
                text.insert('end', '─' * 56 + '\n', 'divider')
            elif name == 'get':
                text.insert('end', 'Aucune ligne GRIPPER GET trouvée.\n', 'none')
                text.insert('end', '─' * 56 + '\n', 'divider')
        if data['locations']:
            text.insert('end', '\nDerniers emplacements connus par CodeLot (tous fichiers) :\n', 'lastpute')
            for event in data['locations']:
                when = datetime.fromtimestamp(event['time'])
                text.insert('end', f"Lot {event['lot_hint']} — {trace_label(event['path'])} @ {when:%d/%m/%Y %H:%M:%S}\n{event['text']}\n", 'lastpute')
        else:
            text.insert('end', 'Aucun emplacement connu trouvé pour cette référence (tous fichiers).\n', 'none')
        self.file_label.configure(text='Fichier : ' + trace_label(data['path']))
        text.configure(state='disabled')
        if hasattr(self, '_update_context_header'):
            self._update_context_header(data['path'], f'{self.line_index + 1}/{len(self.lines)}')
        self._render_summary(data)
        self.btn_adjacent.configure(state='normal' if self.selection.modbus else 'disabled')
        popup = getattr(self, 'adj_popup', None)
        if popup is not None and popup.winfo_exists():
            if self.selection.modbus:
                self.open_adjacent_for_current_occurrence()
            else:
                self._clear_adjacent('Cette ligne ne contient pas de ModBusId.')

    def _change_result_mode(self, value):
        """Changing the presentation never changes or re-queries the selection."""
        self.result_mode.set(value)
        self._sync_result_visibility()
        self._render_summary(self.selection.details if self.selection else None)

    def _render_summary(self, data):
        container = getattr(self, 'summary_frame', None)
        if container is None or not container.winfo_exists():
            return
        for child in container.winfo_children():
            child.destroy()
        self.summary_buttons = {}
        if getattr(self, 'result_mode', None) and self.result_mode.get() != 'Synthèse':
            return
        palette = getattr(self, '_palette', {})
        foreground = palette.get('text', '#172a3a')
        muted = palette.get('muted', '#526477')
        surface = palette.get('surface', '#ffffff')
        if not data:
            ctk.CTkLabel(container, text='Aucune occurrence sélectionnée.',
                         text_color=muted, font=('Segoe UI', 12)).pack(padx=16, pady=24)
            return
        row = data['row']
        title = f"{row.get('ref') or 'Sans référence'}   ·   ID {row.get('modbus') or '—'}"
        ctk.CTkLabel(container, text=title, font=('Segoe UI', 15, 'bold'),
                     text_color=foreground).pack(anchor='w', padx=16, pady=(12, 8))
        # The indexed data supplies the real timestamps. Missing events remain explicit.
        events = [(name, data.get(name)) for name in ('get', 'idpk', 'pute')]
        events.sort(key=lambda pair: (pair[1] is None,
                                     pair[1]['time'] if pair[1] else float('inf'),
                                     pair[1]['number'] if pair[1] else 0))
        for name, event in events:
            card = ctk.CTkFrame(container, fg_color=surface, corner_radius=8,
                                border_width=1, border_color=palette.get('border', '#d6dfe8'))
            card.pack(fill='x', padx=12, pady=5)
            title = name.upper()
            if event:
                title += '   ' + datetime.fromtimestamp(event['time']).strftime('%d/%m/%Y %H:%M:%S.%f')[:-3]
            ctk.CTkLabel(card, text=title, text_color=foreground,
                         font=('Segoe UI', 12, 'bold')).pack(anchor='w', padx=12, pady=(10, 4))
            if event:
                detail = event['text']
                if name == 'idpk' and event.get('dims'):
                    dimensions = json.loads(event['dims']) if isinstance(event['dims'], str) else event['dims']
                    detail = '   '.join(f'{key} = {value}' for key, value in dimensions.items()) + '\n' + detail
                label = ctk.CTkLabel(card, text=detail, text_color=foreground, anchor='w',
                                     justify='left', font=('Consolas', 11), wraplength=600)
                label.pack(fill='x', padx=12, pady=4)
                card.bind('<Configure>', lambda event_, label_=label: label_.configure(
                    wraplength=max(120, event_.width - 32)), add='+')
                button = ctk.CTkButton(card, text='Ouvrir la ligne', width=125,
                                       command=lambda key=name: self._open_selected_event(key))
                button.pack(anchor='w', padx=12, pady=(4, 10))
                self.summary_buttons[name] = button
            else:
                ctk.CTkLabel(card, text='Non trouvé pour cette occurrence', text_color=muted,
                             font=('Segoe UI', 11)).pack(anchor='w', padx=12, pady=(0, 10))
        raw = ManagedTextbox(container, height=90, font=('Consolas', 11), fg_color=surface,
                              text_color=foreground, wrap='word')
        raw.pack(fill='x', padx=12, pady=(12, 5))
        raw.insert('end', 'Ligne originale\n' + row['text'])
        raw.configure(state='disabled')
        locations = data.get('locations', [])
        location_text = '\n\n'.join(
            f"Lot {event['lot_hint']} · {trace_label(event['path'])}\n{event['text']}"
            for event in locations)
        location_box = ManagedTextbox(container, height=130, font=('Consolas', 11),
                                       fg_color=surface, text_color=foreground, wrap='word')
        location_box.pack(fill='x', padx=12, pady=(5, 12))
        location_box.insert('end', 'Derniers emplacements connus par lot\n' +
                            (location_text or 'Aucun emplacement connu pour cette référence.'))
        location_box.configure(state='disabled')

    def _open_selected_event(self, key):
        selection = self.selection
        if selection and selection.details.get(key):
            event = selection.details[key]
            self._open_indexed_file(event.get('path', selection.path), event['number'], event['text'])

    def _move_group(self, offset):
        if not self._groups:
            return
        target = self._group_index + offset
        if getattr(self, '_active_mode', None) == 'advanced':
            target %= len(self._groups)
        elif not 0 <= target < len(self._groups):
            return
        self._group_index = target
        path, hits = self._groups[self._group_index]
        self._begin_selection_change()
        self._submit('details', lambda cancel: self.store.details(path, hits[0][0]), self._apply_selection)

    def on_right(self):
        self._move_group(1)

    def on_left(self):
        self._move_group(-1)

    def on_up(self):
        self._change_search_number(-1)

    def on_down(self):
        self._change_search_number(1)

    def _change_search_number(self, offset):
        if self.modbus.get():
            try:
                value = int(self.keyword.get()) + offset
            except ValueError:
                return
            self.keyword.delete(0, 'end'); self.keyword.insert(0, str(value)); self.perform_search()
        elif self.lines:
            self.line_index = (self.line_index + offset) % len(self.lines)
            self.display_line()

    def search_ref_in_current_file(self, ref):
        if not self.selection:
            return
        path = self.selection.path
        self._begin_selection_change()
        def work(cancel):
            groups = self.store.search('RefProduit=' + ref, path=path, cancel=cancel)
            return groups, self.store.details(path, groups[0][1][0][0]) if groups else None
        def done(result):
            self._groups = result[0]; self._group_index = 0; self._apply_selection(result[1])
        self._submit('search', work, done)

    def _focus_indexed_ticket(self, path, modbus, ref, show_adjacent=False, number=None):
        self.tasks.cancel('search')
        self._query_serial += 1
        self._begin_selection_change()
        self._groups = []
        self.lines = []
        self._selection_numbers = []
        self.reference.set(True); self.modbus.set(False)
        self.on_reference_check()
        self.keyword.delete(0, 'end'); self.keyword.insert(0, ref)
        self.btn_compare_idpk.configure(state='normal')
        def work(cancel):
            groups = self.store.search('ModBusId=' + modbus, True, ref=ref, path=path, cancel=cancel)
            if groups:
                hits = groups[0][1]
                hits = [(n, text) for n, text in hits if number is None or n == number]
                groups = [(path, hits[:1])] if hits else []
            return groups, self.store.details(path, groups[0][1][0][0], anchor=True) if groups else None
        def done(result):
            self._groups, details = result; self._group_index = 0; self._apply_selection(details)
            self._active_mode = 'ref'
            self._last_search = (ref, 'ref', None, None, self.selection)
            self.status.set(f'{len(self._groups)} fichier(s) correspondant(s)')
            if show_adjacent and details:
                self.open_adjacent_for_current_occurrence()
        self._submit('details', work, done)

    def _on_adjacent_ref_click(self, modbus_id, refprod, filepath, number=None):
        self._focus_indexed_ticket(filepath, modbus_id, refprod, number=number)

    def on_click_refprod(self, event):
        selection = self.selection
        if selection is None or not selection.reference:
            return
        if self.reference.get():
            if not selection.modbus:
                messagebox.showinfo('Info', 'Cette occurrence ne contient pas de ModBusId.')
                return
            self.modbus.set(True); self.reference.set(False)
            self.on_modbus_check()
            value = selection.modbus
        else:
            self.reference.set(True); self.modbus.set(False)
            self.on_reference_check()
            value = selection.reference
        self.keyword.delete(0, 'end'); self.keyword.insert(0, value)
        mode = 'id' if self.modbus.get() else 'ref'
        self.btn_compare_idpk.configure(state='normal' if mode == 'ref' else 'disabled')
        self._launch_search(value, mode, None, None, preferred=selection)

    def compare_previous_idpk_popup(self):
        if not self._diagnostic_ready():
            return
        if not self.selection:
            return
        ref, path = self.selection.reference, self.selection.path
        def done(pair):
            if not pair:
                messagebox.showinfo('Comparaison IDPK', 'Aucune occurrence précédente trouvée dans les fichiers plus anciens.')
                return
            self._render_indexed_comparison(ref, pair)
        self._submit('comparison', lambda cancel: self.store.comparison(path, ref), done)

    def open_adjacent_for_current_occurrence(self):
        if self.selection and self.selection.modbus:
            self.show_adjacent_modbusids_window(self.selection.modbus, self.selection.path,
                                               self.selection.number, self.selection.reference)

    def show_adjacent_modbusids_window(self, target_id, filepath, number=None, reference=None):
        if not filepath:
            return
        selection = self.selection
        if selection and selection.path == filepath and selection.modbus == str(target_id):
            number, reference = selection.number, selection.reference
        def done(data):
            if data:
                self._render_indexed_adjacent(str(target_id), filepath, data)
            else:
                self._clear_adjacent('Occurrence indisponible ; relancez la recherche.')
        self._submit('adjacent', lambda cancel: self.store.adjacent(filepath, str(target_id), number, reference), done)

    def _open_exact_gripper_line(self, event, gripper_type):
        key = 'get' if gripper_type == 'GRIPPER GET' else 'pute'
        self._open_selected_event(key)

    def open_gripper_line_view(self, gripper_type):
        self._open_exact_gripper_line(None, gripper_type)

    def open_gripper_line_view_exact(self, gripper_type, target_line):
        self._open_exact_gripper_line(None, gripper_type)

    def _open_indexed_file(self, path, highlight=None, expected_line=None):
        if not path:
            messagebox.showinfo('Info', 'Aucun fichier de trace disponible.')
            return
        def work(cancel):
            if not os.path.isfile(path):
                return None, 'Le fichier de trace a été supprimé ou déplacé.'
            lines = self.store.file_lines(path)
            selected = highlight
            if expected_line and (selected is None or selected >= len(lines) or lines[selected] != expected_line):
                matches = [i for i, line in enumerate(lines) if line == expected_line]
                if len(matches) != 1:
                    return None, 'La trace a changé et la ligne ne peut plus être identifiée. Relancez la recherche.'
                selected = matches[0]
            return (lines, selected), None
        def done(result):
            content, error = result
            if error:
                messagebox.showinfo('Trace modifiée', error)
            else:
                self._render_indexed_file_view(path, *content)
        self._submit('file_view', work, done)

    def _indexed_open_file_view(self, highlight_line=None):
        path = self.selection.path if self.selection else None
        if highlight_line is None and not self.keyword.get().strip():
            path = next((p for p in self.file_list if os.path.basename(p).lower().startswith('tracomega')
                         and not os.path.basename(p).lower().startswith('tracomegaperm')), None)
        self._open_indexed_file(path, highlight_line)

    def _show_trace_window(self, path, highlight_line_num):
        self._open_indexed_file(path, highlight_line_num)

    def _insert_text_batches(self, widget, lines, done=None):
        if not hasattr(self, '_text_tokens'):
            self._text_tokens = {}
        token = object()
        self._text_tokens[str(widget)] = token
        iterator = iter(lines)
        scheduled = [None]
        def cleanup(event=None):
            if event is not None and event.widget is not widget:
                return
            if self._text_tokens.get(str(widget)) is token:
                self._text_tokens.pop(str(widget), None)
            if scheduled[0] is not None:
                try:
                    self.after_cancel(scheduled[0])
                except tk.TclError:
                    pass
                scheduled[0] = None
        widget.bind('<Destroy>', cleanup, add='+')
        def chunk():
            scheduled[0] = None
            if self._text_tokens.get(str(widget)) is not token or not widget.winfo_exists():
                return
            batch = []
            for _ in range(300):
                try:
                    batch.append(next(iterator))
                except StopIteration:
                    if batch:
                        widget.insert('end', ''.join(batch))
                    if done:
                        done()
                    return
            widget.insert('end', ''.join(batch))
            scheduled[0] = self.after(1, chunk)
        chunk()

    def _render_items(self, items, render, owner):
        if not hasattr(self, '_item_tokens'):
            self._item_tokens = {}
        token = object()
        self._item_tokens[str(owner)] = token
        iterator = iter(items)
        scheduled = [None]
        def cleanup(event=None):
            if event is not None and event.widget is not owner:
                return
            if self._item_tokens.get(str(owner)) is token:
                self._item_tokens.pop(str(owner), None)
            if scheduled[0] is not None:
                try:
                    self.after_cancel(scheduled[0])
                except tk.TclError:
                    pass
                scheduled[0] = None
        owner.bind('<Destroy>', cleanup, add='+')
        def chunk():
            scheduled[0] = None
            if self._item_tokens.get(str(owner)) is not token or not owner.winfo_exists():
                return
            for _ in range(20):
                try:
                    item = next(iterator)
                except StopIteration:
                    return
                render(item)
            scheduled[0] = self.after(1, chunk)
        chunk()

    def _indexed_tracking(self):
        ref = self.tracking_keyword.get().strip()
        if not ref:
            return
        def done(rows):
            widget = self.tracking_txt
            if not widget.winfo_exists():
                return
            widget.configure(state='normal'); widget.delete('1.0', 'end')
            def texts():
                for row in rows:
                    when = datetime.fromtimestamp(row['time']); _, hour = timestamp_parts(row['text'])
                    face = re.search(r'(face:\d+\s*\d*)', row['text'])
                    yield f"{when:%d/%m/%y} | {hour} | {trace_label(row['path'])} | PUTE | {face[1] if face else ''}\n"
            self._insert_text_batches(widget, texts() if rows else ['Aucune opération PUTE trouvée pour cette référence.\n'],
                                      lambda: widget.configure(state='disabled'))
        self._submit('tracking', lambda cancel: self.store.tracking(ref), done)

    def check_errors(self):
        if not self._diagnostic_ready():
            return
        def done(rows):
            errors = []
            for row in rows:
                match = re.search(r'DateREQ=(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2}:\d{2})', row['text'])
                when = datetime.fromtimestamp(row['time'])
                errors.append({'file': trace_label(row['path']), 'filepath': row['path'], 'modbus_id': row['modbus'],
                               'ref': row['ref'], 'number': row['number'], 'date': match[1] if match else when.strftime('%d/%m/%Y'),
                               'heure': match[2] if match else when.strftime('%H:%M:%S')})
            self._render_robot_errors(errors, [trace_label(p) for p in self.file_list])
        self._submit('robot_errors', lambda cancel: self.store.robot_errors(), done)

    def convoyage_action(self, ref=None, geometry_override=None, only_filepath=None):
        ref = ref if ref is not None else self.keyword.get().strip()
        if not ref:
            messagebox.showwarning('Attention', 'Veuillez entrer une référence à rechercher.'); return
        self.last_convoyage_ref = ref
        self._submit('convoyage', lambda cancel: (self.store.convoyage_summary(ref, only_filepath), self.store.convoyage_steps(ref)),
                     lambda result: self._render_indexed_convoyage(ref, geometry_override, *result))

    def convoyage_errors_action(self):
        if self._diagnostic_ready():
            self._submit('convoyage_errors', lambda cancel: self.store.convoyage_errors(), self._render_convoyage_errors)

    def show_convoyage_details(self, ref):
        self._submit('convoyage_details', lambda cancel: self.store.convoyage_recap(ref),
                     lambda recap: self._render_indexed_convoyage_recap(ref, recap))
