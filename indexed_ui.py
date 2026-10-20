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
from trace_sources import discover_trace_files, trace_label, available_trace_sources, trace_directories
from trace_sources import preferred_trace_directories, compact_trace_label
from trace_store import TraceStore, Cancelled, CHECKPOINTS, cache_path
from source_choice import SourceChoiceDialog
from app_assets import apply_icon
from selection_state import OccurrenceSelection
from ui_theme import ManagedTextbox
from trace_identity import TraceRef
from omega_history import HistoryUnavailable


class IndexedAppMixin:
    def _compact_trace_label(self, path):
        return compact_trace_label(path, getattr(self, '_selected_directories', None))

    def _start_indexed_service(self):
        self.tasks = BackgroundTasks()
        self._index_messages = Queue()
        self._stop_index = threading.Event()
        self._sources_chosen = threading.Event()
        self._selected_directories = None
        self._source_dialog = None
        self._groups = []
        self._group_index = 0
        self._selection_numbers = []
        self.selection = None
        self._selection_serial = 0
        self.btn_adjacent.configure(state='disabled')
        self._view_generation = 0
        self._indexed_complete = False
        self._sources_verified = False
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
        self.bind_class('Toplevel', '<Map>', self._icon_on_map, add='+')
        self.after(250, lambda: apply_icon(self))
        self._index_thread = threading.Thread(target=self._index_loop, daemon=True, name='trace-index')
        self._index_thread.start()
        self.after(30, self._poll_indexed_service)

    def _icon_on_map(self, event):
        if isinstance(event.widget, (tk.Toplevel, tk.Tk)):
            apply_icon(event.widget)

    def _index_loop(self):
        try:
            preferred = preferred_trace_directories()
            sources = available_trace_sources() if preferred is None else ()
            if preferred is not None:
                self._selected_directories = preferred
            elif len(sources) > 1:
                self._index_messages.put(('sources', sources))
                while not self._sources_chosen.wait(.1):
                    if self._stop_index.is_set():
                        return
            else:
                self._selected_directories = (tuple(directory for name, directory in sources)
                                              if sources else tuple(trace_directories()))
            if self._stop_index.is_set():
                return
            store = TraceStore(cache_path(self._selected_directories), directories=self._selected_directories,
                               stop=self._stop_index, defer_history=True)
            try:
                metadata = store.metadata()
            except sqlite3.DatabaseError:
                if not store.corrupt.is_set():
                    raise
                store.rebuild(); metadata = []
            self._index_messages.put(('ready', store, metadata, store.initialized()))
            store.prepare_analysis(self._stop_index)
            self._index_messages.put(('phase', 'Préparation de l’historique'))
            while not self._stop_index.is_set():
                try:
                    store.prepare_history(self._stop_index)
                    break
                except HistoryUnavailable as error:
                    self._index_messages.put(('history_error', str(error)))
                    self._stop_index.wait(2)
            verified_once = False
            while not self._stop_index.is_set():
                if store.corrupt.is_set():
                    self._index_messages.put(('rebuilding',))
                    store.rebuild()
                    self._index_messages.put(('ready', store, [], False))
                paths = discover_trace_files(directories=self._selected_directories)
                if not verified_once:
                    self._index_messages.put(('phase', 'Vérification des sources'))
                def progress(done, total, path, published):
                    self._index_messages.put(('progress', done, total, path, published,
                                              store.metadata() if published else None))
                try:
                    result = store.sync(paths, progress, self._stop_index)
                except HistoryUnavailable as error:
                    self._index_messages.put(('history_error', str(error)))
                    self._stop_index.wait(2)
                    continue
                except sqlite3.DatabaseError:
                    if store.corrupt.is_set():
                        continue
                    raise
                self._index_messages.put(('complete', result, store.metadata()))
                verified_once = not bool(result['failures'])
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
                if item[0] == 'sources':
                    self.index_status.set('Choisissez les sources de traces à charger')
                    self._source_dialog = SourceChoiceDialog(
                        self, item[1], self._accept_trace_sources, self._close_indexed_app)
                elif item[0] == 'ready':
                    self._index_content_revision = getattr(self, '_index_content_revision', 0) + 1
                    self.store = item[1]
                    self._indexed_complete = item[3]
                    self._install_metadata(item[2])
                    self._sources_verified = False
                    self.index_status.set('Cache disponible — vérification des sources en cours'
                                          if item[2] else 'Préparation de l’index')
                    if getattr(self, 'btn_erreurs', None) and self.btn_erreurs.winfo_exists():
                        self.btn_erreurs.configure(state='disabled')
                elif item[0] == 'phase':
                    prefix = 'Cache disponible — ' if self.file_list else ''
                    self.index_status.set(prefix + item[1])
                elif item[0] == 'progress':
                    _, done, total, path, published, metadata = item
                    self.index_status.set(f'Indexation {done}/{total} : {os.path.basename(path)} — résultats partiels'
                                    if not self._indexed_complete or not self._sources_verified else f'Actualisation {done}/{total} : {os.path.basename(path)}')
                    if published:
                        self._install_metadata(metadata)
                        if getattr(self, '_last_search', None):
                            self._refresh_indexed_search()
                elif item[0] == 'complete':
                    previously_verified = self._sources_verified
                    changed = item[1]['changed'] or item[1]['removed']
                    if changed:
                        self._index_content_revision = getattr(self, '_index_content_revision', 0) + 1
                    self._install_metadata(item[2])
                    was_complete = self._indexed_complete
                    self._indexed_complete = was_complete or not bool(item[1]['failures'])
                    self._sources_verified = not bool(item[1]['failures'])
                    self.files_loaded = self._indexed_complete
                    if self._sources_verified and getattr(self, 'selection', None):
                        if not previously_verified:
                            self._render_indexed_details(self.selection.details)
                        elif '_summary_comparison' not in self.selection.details:
                            self._render_summary(self.selection.details)
                    failures = item[1]['failures']
                    if failures:
                        self.index_status.set(f'Indexation incomplète — {len(self.file_list)} fichiers disponibles, '
                                              f'{len(failures)} fichier(s) à réessayer')
                    else:
                        self.index_status.set(f'{len(self.file_list)} fichiers prêts')
                    if getattr(self, 'btn_erreurs', None) and self.btn_erreurs.winfo_exists():
                        self.btn_erreurs.configure(state='normal' if self._indexed_complete and self._sources_verified else 'disabled')
                    self.title('Recherche TracOmega')
                    if (changed or not was_complete) and getattr(self, '_last_search', None):
                        self._refresh_indexed_search()
                    elif self._sources_verified and not previously_verified:
                        popup = getattr(self, 'adj_popup', None)
                        if popup is not None and popup.winfo_exists() and self.selection:
                            self.open_adjacent_for_current_occurrence()
                elif item[0] == 'error':
                    self.index_status.set(f'Index indisponible : {item[1]}')
                elif item[0] == 'history_error':
                    self._sources_verified = False
                    if getattr(self, 'btn_erreurs', None) and self.btn_erreurs.winfo_exists():
                        self.btn_erreurs.configure(state='disabled')
                    self.index_status.set(item[1])
                elif item[0] == 'rebuilding':
                    self.index_status.set('Reconstruction de l’index local…')
                    self._indexed_complete = False
                    self._sources_verified = False
                    self._invalidate_view()
        except Empty:
            pass
        try:
            self.tasks.drain()
        finally:
            self.after(30, self._poll_indexed_service)

    def _accept_trace_sources(self, directories):
        if not directories:
            return
        self._selected_directories = tuple(directories)
        self._source_dialog.destroy()
        self._source_dialog = None
        self.index_status.set('Ouverture de l’index…')
        self._sources_chosen.set()

    def _refresh_indexed_search(self):
        raw, mode, sn, lot, preferred = self._last_search
        self._launch_search(raw, mode, sn, lot, preferred=self.selection or preferred, _refresh=True)

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
            if channel == 'adjacent':
                self._clear_adjacent(f'Voisins indisponibles : {error}')
            elif channel == 'tracking':
                self._tracking_loading = False
                widget = getattr(self, 'tracking_view', None)
                if widget is not None and widget.winfo_exists():
                    widget.count.configure(text='Suivi indisponible. Relancez la recherche pour réessayer.')
        view, serial = self._view_generation, self._selection_serial
        def apply(result):
            if view != self._view_generation:
                return
            if channel in ('details', 'adjacent', 'comparison', 'summary_comparison', 'file_view') and serial != self._selection_serial:
                return
            done(result)
        return self.tasks.submit(channel, compute, apply, failed)

    @property
    def _display_data(self):
        return self.selection.details if self.selection else None

    def _begin_selection_change(self):
        self._selection_serial += 1
        self.selection = None
        self._summary_comparison_pending = None
        self.btn_adjacent.configure(state='disabled')
        for channel in ('details', 'adjacent', 'comparison', 'summary_comparison', 'file_view'):
            self.tasks.cancel(channel)
        self._clear_adjacent('Recherche en cours…')
        if not getattr(self, '_suspend_summary', False):
            from summary_widgets import set_summary_loading
            set_summary_loading(self)

    def _clear_adjacent(self, message='Aucune occurrence sélectionnée.'):
        popup = getattr(self, 'adj_popup', None)
        if popup is not None and popup.winfo_exists():
            if message == 'Recherche en cours…':
                self.adjacent_view.loading()
            else:
                self.adjacent_view.status(message)

    def _clear_selection(self):
        self._begin_selection_change()
        self._groups = []
        self._group_index = 0
        self.lines = []
        self.line_index = 0
        self._selection_numbers = []
        self._clear_adjacent()
        if hasattr(self, '_update_context_header') and not getattr(self, '_suspend_summary', False):
            self._update_context_header()
        self._render_summary(None)

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
        if not self._indexed_complete or not self._sources_verified:
            messagebox.showinfo('Indexation en cours', 'Ce contrôle sera disponible après la vérification et l’indexation complète des traces.')
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

    def _location_options(self):
        query = getattr(self, '_last_search', None)
        return {'location_lot': query[3]} if query and query[1] == 'advanced' and query[3] is not None else {}

    def _launch_search(self, raw, mode, sn, lot, preferred=None, _refresh=False):
        previous_details = self.selection.details if self.selection else None
        self._last_search = (raw, mode, sn, lot, preferred)
        self._query_serial += 1
        serial = self._query_serial
        keep_selection = _refresh and self.selection is not None
        if not keep_selection:
            self._begin_selection_change()
            self._groups = []
            self.lines = []
            self._selection_numbers = []
        selection_serial = self._selection_serial
        pool = getattr(getattr(self, 'summary_frame', None), '_summary_widgets', None)
        self.status.set(('Chargement…' if pool is not None and pool.data else 'Recherche…')
                        + (' résultats partiels' if not self._indexed_complete else ''))
        def search(cancel):
            groups = self.store.search('ModBusId=' + raw if mode == 'id' else 'RefProduit=' + raw if mode == 'ref' else raw,
                                       mode == 'id', ref=raw if mode == 'advanced' else None, sn=sn, lot=lot, cancel=cancel)
            index, number = 0, groups[0][1][0][0] if groups else None
            if preferred:
                for i, (path, hits) in enumerate(groups):
                    if path == preferred.path and any(n == preferred.number for n, _ in hits):
                        index, number = i, preferred.number
                        break
            options = {'location_lot': lot} if mode == 'advanced' and lot is not None else {}
            details = self._details_with_revision(groups[index][0], number, **options) if groups else None
            if details and preferred:
                from summary_widgets import reuse_summary_analysis
                reuse_summary_analysis(previous_details, details, details['_summary_revision'])
            return groups, index, details
        def apply(result):
            if serial != self._query_serial:
                return
            if keep_selection and selection_serial != self._selection_serial:
                return
            if keep_selection:
                details = result[2]
                if not details or details['path'] != preferred.path or details['row']['number'] != preferred.number:
                    self._begin_selection_change()
            self._groups, self._group_index, details = result
            self._active_mode = mode
            self._apply_selection(details)
            self.status.set(('Résultats partiels — indexation en cours' if not self._indexed_complete else
                             f'{len(self._groups)} fichier(s) correspondant(s)'))
        self._submit('search', search, apply)

    def _details_with_revision(self, path, number, **options):
        # Called inside the worker snapshot, so data and revision describe one publication.
        details = self.store.details(path, number, **options)
        if details:
            details['_summary_revision'] = self.store.analysis_revision()
        return details

    def _apply_selection(self, details):
        if not details:
            self._clear_selection()
            self.file_label.configure(text='Fichier : Aucun')
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
        options = self._location_options()
        self._begin_selection_change()
        self._submit('details', lambda cancel: self._details_with_revision(path, number, **options), self._render_indexed_details)

    def _render_indexed_details(self, data):
        if not data:
            self._apply_selection(None)
            return
        if not self.summary_frame.winfo_exists():
            return
        self.selection = OccurrenceSelection.from_details(data)
        self.file_label.configure(text='Fichier : ' + self._compact_trace_label(data['path']))
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

    def _render_summary(self, data):
        if getattr(self, '_suspend_summary', False):
            return
        from summary_view import render_summary
        render_summary(self, data)

    def _load_summary_comparison(self, data):
        if getattr(self, '_summary_comparison_pending', None) is data:
            return
        self._summary_comparison_pending = data
        local = '_summary_current' not in data
        def compute(cancel):
            try:
                method = self.store.current_measurement_for_occurrence if local else self.store.comparison_for_occurrence
                return method(data['path'], data['row']['number'], cancel)
            except (sqlite3.DatabaseError, ValueError) as error:
                return {'error': str(error)}
        def done(result):
            self._summary_comparison_pending = None
            data['_summary_current' if local else '_summary_comparison'] = result
            if not local:
                # The complete result replaces the local stage, including after
                # a publication between the two worker snapshots.
                data.pop('_summary_current', None)
            if self.selection and self.selection.details is data:
                self._render_summary(data)
        self._submit('summary_comparison', compute, done)

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
        options = self._location_options()
        self._begin_selection_change()
        self._submit('details', lambda cancel: self._details_with_revision(path, hits[0][0], **options), self._apply_selection)

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
            return groups, self._details_with_revision(path, groups[0][1][0][0]) if groups else None
        def done(result):
            self._groups = result[0]; self._group_index = 0; self._apply_selection(result[1])
        self._submit('search', work, done)

    def _focus_indexed_ticket(self, path, modbus, ref, show_adjacent=False, number=None):
        selected = self.selection
        if selected and selected.path == path and selected.number == number:
            return
        if getattr(self, '_active_panel', 'search') != 'search':
            self.show_search_panel()
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
            details = self._details_with_revision(path, groups[0][1][0][0], anchor=True) if groups else None
            return groups, details
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
        ref, path, number = self.selection.reference, self.selection.path, self.selection.number
        self._submit('comparison', lambda cancel: self.store.comparison_for_occurrence(path, number, cancel),
                     lambda result: self._render_indexed_comparison(ref, result))

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
        from adjacent_view import AdjacentView
        from window_placement import show_centered_popup, monitor_work_area, native_window_bounds
        popup = getattr(self, 'adj_popup', None)
        if popup is None or not popup.winfo_exists():
            popup = self.adj_popup = tk.Toplevel(self)
            popup.title('IDs adjacents — R1 / R2')
            apply_icon(popup)
            self.adjacent_view = AdjacentView(popup, self)
            self.adjacent_view.pack(fill='both', expand=True)
            def close():
                self.tasks.cancel('adjacent')
                popup.destroy()
            popup.protocol('WM_DELETE_WINDOW', close)
            self.adjacent_view.close_button.configure(command=close)
            show_centered_popup(popup, self, 1180, 720)
            work = monitor_work_area(self)
            extra_width = extra_height = 0
            if os.name == 'nt':
                _, bounds = native_window_bounds(popup)
                extra_width = bounds[2] - bounds[0] - popup.winfo_width()
                extra_height = bounds[3] - bounds[1] - popup.winfo_height()
            popup.maxsize(work[2] - work[0] - extra_width, work[3] - work[1] - extra_height)
        counts = getattr(self, '_adjacent_counts', (5, 5))
        self.adjacent_view.loading('Actualisation des voisins R1 / R2…')
        def done(data):
            if self.adj_popup is not popup or not popup.winfo_exists():
                return
            if counts != getattr(self, '_adjacent_counts', (5, 5)):
                self.open_adjacent_for_current_occurrence()
                return
            if data:
                self._render_indexed_adjacent(str(target_id), filepath, data)
            else:
                self._clear_adjacent('Occurrence indisponible ; relancez la recherche.')
        self._submit('adjacent', lambda cancel: self.store.paired_adjacent(
            filepath, str(target_id), number, reference, cancel,
            before_count=counts[0], after_count=counts[1]), done)

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
            if not isinstance(path, TraceRef) and not os.path.isfile(path):
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

    def _indexed_tracking(self, refresh=False):
        if refresh and getattr(self, '_tracking_criteria', None):
            ref, sn, lot = self._tracking_criteria
        else:
            ref = self.tracking_view.reference_entry.get().strip()
            sn = self.tracking_view.sn_entry.get().strip()
            lot = self.tracking_view.lot_entry.get().strip()
            self._tracking_criteria = ref, sn, lot
        widget = self.tracking_view
        widget.begin(ref, sn=sn, lot=lot)
        self._tracking_loading = any((ref, sn, lot))
        if not self._tracking_loading:
            self.tasks.cancel('tracking')
            return
        revision = getattr(self, '_index_content_revision', 0)
        def done(rows):
            if not widget.winfo_exists():
                return
            widget.show_results(rows)
            self._tracking_loading = False
            self._tracking_revision = revision
        self._submit('tracking', lambda cancel: self.store.tracking(ref, sn=sn or None, lot=lot or None, cancel=cancel), done)

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
                               'heure': match[2] if match else when.strftime('%H:%M:%S'),
                               'sn': row.get('sn'), 'lot': row.get('lot'), 'text': row['text']})
            self._render_robot_errors(errors, [trace_label(p) for p in self.file_list])
        self._submit('robot_errors', lambda cancel: self.store.robot_errors(), done)

    def convoyage_action(self, ref=None, geometry_override=None, only_filepath=None):
        ref = ref if ref is not None else self.keyword.get().strip()
        if not ref:
            messagebox.showwarning('Attention', 'Veuillez entrer une référence à rechercher.'); return
        self.last_convoyage_ref = ref
        self._submit('convoyage', lambda cancel: (self.store.convoyage_summary(ref, only_filepath), []),
                     lambda result: self._render_indexed_convoyage(ref, geometry_override, *result))

    def convoyage_errors_action(self):
        if self._diagnostic_ready():
            self._submit('convoyage_errors', lambda cancel: self.store.convoyage_errors(), self._render_convoyage_errors)

    def show_convoyage_details(self, ref):
        def work(cancel):
            recap = self.store.convoyage_recap(ref)
            paths = list(dict.fromkeys(row['path'] for row in self.store.modbus_rows(ref)))
            return recap, paths
        self._submit('convoyage_details', work,
                     lambda result: self._render_indexed_convoyage_recap(ref, *result))
