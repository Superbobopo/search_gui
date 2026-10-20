"""Conveyor anomalies: virtual cards and an independent exact-file detail view."""
import ntpath
import os
import tkinter as tk

import customtkinter as ctk

from background_tasks import BackgroundTasks
from selectable_text import make_selectable
from trace_sources import compact_trace_label, trace_badge, trace_tooltip
from ui_components import (ActionButton, WindowHeader, SplitView, VirtualCardList,
                           StateCard, selectable, register_view)
from trace_store import Cancelled


def anomaly_caption(path, directories):
    """Display names only; preserve the original path as the navigation identity."""
    robot = trace_badge(path)
    label = compact_trace_label(path, directories)
    prefix = label.split(' \\', 1)[0] if ' \\' in label else ''
    origin = prefix.removesuffix(robot).strip()
    return robot, ntpath.basename(os.fspath(path)) + (' · ' + origin if origin else '')


class _StepRow(ctk.CTkFrame):
    def __init__(self, parent, embedded=False):
        super().__init__(parent, corner_radius=0, fg_color='transparent')
        self.grid_columnconfigure(2, weight=1)
        self.dot = ctk.CTkLabel(self, text='●', width=18, font=('Segoe UI', 12))
        self.dot.grid(row=0, column=0, sticky='n', padx=(0, 10), pady=10)
        self.time = selectable(self, text='', font=('Segoe UI', 12 if embedded else 11))
        self.time.grid(row=0, column=1, sticky='nw', padx=(0, 16), pady=13)
        self.checkpoint = selectable(self, text='', wraplength=350, font=('Segoe UI', 14 if embedded else 12))
        self.checkpoint.grid(row=0, column=2, sticky='ew', pady=10)
        self.status = selectable(self, text='', font=('Segoe UI', 12 if embedded else 10, 'bold'))
        self.status.grid(row=1, column=2, sticky='ew', pady=(0, 10))
        self.step = None

    def show(self, step):
        self.step = step
        for label, key in ((self.time, 'horodatage'), (self.checkpoint, 'checkpoint'), (self.status, 'etat')):
            label.configure(text=step.get(key) or '—')
            label.sync_selection_text()

    def apply_theme(self, colors):
        role = 'positive_text' if self.step and self.step['etat']=='Présent' else 'error_text'
        self.configure(fg_color=colors['surface'], bg_color=colors['surface'])
        self.dot.configure(text_color=colors[role], bg_color=colors['surface'])
        for label, color in ((self.time, 'muted'), (self.checkpoint, 'text'), (self.status, role)):
            label.configure(text_color=colors[color], bg_color=colors['surface'])
            label.sync_selection_text()


class ConvoyageErrorsView(ctk.CTkFrame):
    def __init__(self, parent, app, errors, *, embedded=False):
        super().__init__(parent, corner_radius=0)
        self._theme_deferred = True
        self.app, self.errors, self.embedded = app, list(errors), embedded
        self.selected_index, self.data, self.state = None, None, 'empty'
        self.tasks = BackgroundTasks()
        self.channel = f'convoyage_anomaly:{id(self)}'
        self._closed, self._poll_id, self._layout_id = False, None, None
        self._detail_width = None
        self._request, self._step_rows = 0, []
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        files, count = len({os.fspath(path) for path, _ in self.errors}), len(self.errors)
        self.header = WindowHeader(self, 'Anomalies de convoyage',
            f'{count} anomalie'+('s' if count!=1 else '')+f' · {files} fichier'+('s' if files!=1 else ''), parent.destroy)
        self.header.grid(row=0, column=0, sticky='ew', padx=20, pady=(16, 14))
        if embedded:
            self.header.grid_remove()
        self.title, self.summary, self.close_button = self.header.title, self.header.subtitle, self.header.close_button
        self.split = SplitView(self, breakpoint=900 if embedded else 860)
        self.split.grid(row=1, column=0, sticky='nsew', padx=20, pady=(0, 18))
        if embedded:
            self.split.grid_configure(padx=0, pady=0)
        self.panes = self.split.panes
        self.list_panel, self.detail_panel = self.split.left, self.split.right
        self.list_panel.grid_columnconfigure(0, weight=1)
        self.list_panel.grid_rowconfigure(1, weight=1)
        self.list_heading = ctk.CTkLabel(self.list_panel, text='ANOMALIES', anchor='w', font=('Segoe UI', 11, 'bold'))
        self.list_heading.grid(row=0, column=0, sticky='ew', padx=16, pady=(12, 8))
        self.cards = VirtualCardList(self.list_panel, self)
        self.cards.grid(row=1, column=0, sticky='nsew', padx=(8, 6), pady=(0, 8))
        self.detail_panel.grid_columnconfigure(0, weight=1)
        self.detail_panel.grid_rowconfigure(1, weight=1)
        register_view(app, self)
        detail_header = ctk.CTkFrame(self.detail_panel, fg_color='transparent')
        detail_header.grid(row=0, column=0, sticky='ew', padx=20, pady=(16, 10))
        detail_header.grid_columnconfigure(0, weight=1)
        self.detail_title = selectable(detail_header, text='Aucune anomalie détectée' if not count else '', font=('Segoe UI', 20, 'bold'))
        self.detail_title.grid(row=0, column=0, sticky='ew')
        self.detail_source = selectable(detail_header, text='', font=('Segoe UI', 11), wraplength=600)
        self.detail_source.grid(row=1, column=0, sticky='ew', pady=(5, 0))
        self.path_button = ActionButton(detail_header, text='Copier chemin', role='quiet', width=100, height=28,
            font=('Segoe UI', 10), text_size=14 if embedded else None, command=lambda: self.copy('path'))
        self.path_button.grid(row=0, column=1, padx=(12, 0))
        self.detail_scroll = ctk.CTkScrollableFrame(self.detail_panel, corner_radius=0, fg_color='transparent')
        self.detail_scroll.grid(row=1, column=0, sticky='nsew', padx=(14, 6), pady=(0, 12))
        self.detail_scroll.grid_columnconfigure(0, weight=1)
        self.states = ctk.CTkFrame(self.detail_scroll, fg_color='transparent')
        self.states.grid(row=0, column=0, sticky='ew', padx=4, pady=(6, 18))
        self.states.grid_columnconfigure((0, 1), weight=1, uniform='states')
        self.positive = StateCard(self.states, 'DERNIÈRE POSITION CONNUE')
        self.waiting = StateCard(self.states, 'ATTENTE ENREGISTRÉE')
        self.positive.grid(row=0, column=0, sticky='nsew', padx=(0, 6))
        self.waiting.grid(row=0, column=1, sticky='nsew', padx=(6, 0))
        self.steps_heading = ctk.CTkLabel(self.detail_scroll, text='Étapes relevées', anchor='w', font=('Segoe UI', 14, 'bold'))
        self.steps_heading.grid(row=1, column=0, sticky='ew', padx=6, pady=(0, 4))
        self.steps_frame = ctk.CTkFrame(self.detail_scroll, fg_color='transparent')
        self.steps_frame.grid(row=2, column=0, sticky='ew', padx=6)
        self.steps_frame.grid_columnconfigure(0, weight=1)
        self.message = selectable(self.detail_scroll, text='', font=('Segoe UI', 12), wraplength=600)
        self.message.grid(row=3, column=0, sticky='ew', padx=6, pady=8)
        self.retry = ActionButton(self.detail_scroll, text='Réessayer', width=110, text_size=14 if embedded else None,
                                  command=lambda: self.select(self.selected_index, force=True))
        self.retry.grid(row=4, column=0, sticky='w', padx=6, pady=8)
        self.retry.grid_remove()
        self.provenance_button = ActionButton(self.detail_scroll, text='Afficher la provenance', role='quiet',
            anchor='w', height=28, font=('Segoe UI', 10), text_size=14 if embedded else None, command=self.toggle_provenance)
        self.provenance_button.grid(row=5, column=0, sticky='w', padx=4, pady=(16, 4))
        self.path_label = selectable(self.detail_scroll, text='', font=('Segoe UI', 10), wraplength=600)
        self.path_label.grid(row=6, column=0, sticky='ew', padx=6, pady=4)
        self.path_label.grid_remove()
        if embedded:
            for label in (self.list_heading, self.detail_source, self.path_label):
                label.configure(font=('Segoe UI', 12))
            self.message.configure(font=('Segoe UI', 14))
            for card in (self.positive, self.waiting):
                card.caption.configure(font=('Segoe UI', 12, 'bold'))
                card.value.configure(font=('Segoe UI', 14, 'bold'))
                card.time.configure(font=('Segoe UI', 12))
        self.menu = tk.Menu(self, tearoff=False)
        for title, kind in (('Copier le ModBusId', 'id'), ('Copier la ligne', 'row'), ('Copier le chemin complet', 'path')):
            self.menu.add_command(label=title, command=lambda value=kind: self.copy(value))
        self.detail_scroll.bind('<Configure>', self._resize_detail, add='+')
        tk.Misc.bind(self, '<Destroy>', self._destroyed, add='+')
        self.apply_theme()
        self._poll_id = self.after(30, self._poll)
        if self.errors:
            self.select(0)
        else:
            self.summary.configure(text='Aucune anomalie détectée')
            self.message.configure(text='Toutes les arrivées sont retrouvées dans les fichiers contrôlés.')
            self.message.sync_selection_text()
            for widget in (self.states, self.steps_heading, self.provenance_button):
                widget.grid_remove()
            self.path_button.configure(state='disabled')

    @property
    def _vertical(self):
        return self.split.vertical

    def card_content(self, index):
        path, modbus = self.errors[index]
        badge, subtitle = anomaly_caption(path, self.app._selected_directories)
        return dict(title=f'ID {modbus}', subtitle=subtitle, badge=badge)

    def activate(self, index):
        return self.select(index)

    def reset_items(self, items, *, preserve=False, force=False):
        items = list(items)
        if self.errors==items and not force:
            return
        selected = self.errors[self.selected_index] if self.selected_index is not None else None
        position = self.cards.canvas.yview()[0]
        self.tasks.cancel(self.channel)
        self._request += 1
        self.errors = items
        self.selected_index = None
        self.cards._draw_key = None
        for card, _ in self.cards.cards:
            card._content = None
        canvas = self.cards.canvas
        canvas.configure(scrollregion=(0, 0, canvas.winfo_width(), max(canvas.winfo_height(), len(items)*self.cards.row_height)))
        if not items:
            self.show_empty('Aucun résultat', 'Aucune donnée disponible dans les fichiers sélectionnés.')
            return
        index = items.index(selected) if preserve and selected in items else 0
        if preserve:
            canvas.yview_moveto(position)
        else:
            canvas.yview_moveto(0)
            self.cards.see(index)
        self.select(index, preserve_scroll=preserve and selected==items[index])

    def show_empty(self, title, message):
        """Clear the exact-file destination as well as its rendered contents."""
        self.tasks.cancel(self.channel)
        self._request += 1
        self.errors, self.selected_index, self.data, self.state = [], None, None, 'empty'
        self.cards._draw_key = None
        self.cards.draw()
        self.detail_title.configure(text=title)
        self.detail_source.configure(text='')
        self.path_label.configure(text='')
        self.path_label.grid_remove()
        self.path_button.configure(state='disabled')
        self._show_steps([])
        for widget in (self.states, self.steps_heading, self.provenance_button, self.retry):
            widget.grid_remove()
        self.message.configure(text=message)
        for label in (self.detail_title, self.detail_source, self.message, self.path_label):
            label.sync_selection_text()

    def _resize_detail(self, event):
        width = event.width
        if width == self._detail_width:
            return
        self._detail_width = width
        stacked = width<580
        self.waiting.grid_configure(row=1 if stacked else 0, column=0 if stacked else 1,
            padx=0 if stacked else (6, 0), pady=(8, 0) if stacked else 0)
        self.positive.grid_configure(padx=0 if stacked else (0, 6))
        self.states.grid_columnconfigure(0, weight=1, uniform='' if stacked else 'states')
        self.states.grid_columnconfigure(1, weight=0 if stacked else 1, uniform='' if stacked else 'states')
        for card in (self.positive, self.waiting):
            card.value.configure(wraplength=max(100, width-46 if stacked else width//2-40))
        for label in (self.message, self.path_label):
            label.configure(wraplength=max(100, width-24))
        self.detail_source.configure(wraplength=max(100, width-30))
        for row in self._step_rows:
            row.checkpoint.configure(wraplength=max(100, width-150))

    def _poll(self):
        self._poll_id = None
        if not self._closed:
            self.tasks.drain()
            self._poll_id = self.after(30, self._poll)

    def _destroyed(self, event):
        if event.widget is self:
            self._closed = True
            self.tasks.close()
            for scheduled in (self._poll_id, self._layout_id):
                if scheduled is not None:
                    self.after_cancel(scheduled)

    def select(self, index, force=False, preserve_scroll=False):
        if index is None or not 0<=index<len(self.errors) or (index==self.selected_index and not force):
            return 'break'
        self.selected_index = index
        self._request += 1
        serial = self._request
        scroll = self.detail_scroll._parent_canvas.yview()[0] if preserve_scroll else 0
        path, modbus = self.errors[index]
        self.state, self.data = 'loading', None
        self.detail_title.configure(text=f'ModBusId {modbus}')
        self.detail_source.configure(text=compact_trace_label(path, self.app._selected_directories))
        self.path_label.configure(text=trace_tooltip(path))
        for label in (self.detail_title, self.detail_source, self.path_label):
            label.sync_selection_text()
        self.positive.show('neutral', 'Chargement du parcours…')
        self.waiting.show('neutral', 'Chargement du parcours…')
        self._show_steps([])
        self.message.configure(text='')
        self.message.sync_selection_text()
        self.retry.grid_remove()
        for widget in (self.states, self.steps_heading, self.provenance_button):
            widget.grid()
        self.path_button.configure(state='normal')
        self.detail_scroll._parent_canvas.yview_moveto(scroll)
        self.cards.draw()
        self._paint_detail()
        # Own runner: changing the main panel cannot invalidate this diagnostic.
        store = self.app.store
        def work(cancel):
            with store.snapshot():
                with store.connection() as db:
                    db.set_progress_handler(lambda: int(cancel.is_set()), 1000)
                    summary = store.convoyage_summary(modbus, path)
                    if cancel.is_set():
                        raise Cancelled()
                    return summary, store.convoyage_steps(modbus, path)
        def done(result):
            if self._closed or serial!=self._request:
                return
            self.data, self.state = result, 'ready'
            summary, steps = result
            item = next((item for item in summary if item['_filepath']==path), {})
            positive = item.get('_last_positive')
            if '_checkpoint_actuel' in item:
                positive = {'checkpoint': item['_checkpoint_actuel'], 'horodatage': item.get('HORODATAGE')}
            self.positive.show('positive' if positive else 'neutral',
                positive['checkpoint'] if positive else 'Aucune position positive enregistrée',
                positive.get('horodatage') if positive else '')
            waiting = item.get('_checkpoint_attente')
            self.waiting.show('waiting' if waiting else 'neutral',
                waiting or 'Aucune attente enregistrée', item.get('HORODATAGE') if waiting else '')
            self._show_steps(steps)
            self.message.configure(text='' if steps else 'Aucune étape indexée pour cet ID dans ce fichier.')
            self.message.sync_selection_text()
            self._paint_detail()
            if preserve_scroll:
                self.detail_scroll.update_idletasks()
                self.detail_scroll._parent_canvas.yview_moveto(scroll)
        def failed(error):
            if self._closed or serial!=self._request:
                return
            self.state = 'error'
            self.positive.show('neutral', 'Position indisponible')
            self.waiting.show('neutral', 'Attente indisponible')
            self.message.configure(text=f'Le parcours n’a pas pu être chargé : {error}')
            self.message.sync_selection_text()
            self.retry.grid()
            self._paint_detail()
        self.tasks.submit(self.channel, work, done, failed)
        return 'break'

    def _show_steps(self, steps):
        for index, step in enumerate(steps):
            if index==len(self._step_rows):
                self._step_rows.append(_StepRow(self.steps_frame, self.embedded))
            row = self._step_rows[index]
            row.grid(row=index, column=0, sticky='ew')
            row.show(step)
            if self._detail_width:
                row.checkpoint.configure(wraplength=max(100, self._detail_width-150))
        for row in self._step_rows[len(steps):]:
            row.grid_remove()

    def toggle_provenance(self):
        visible = bool(self.path_label.winfo_manager())
        self.path_label.grid_remove() if visible else self.path_label.grid()
        self.provenance_button.configure(text='Afficher la provenance' if visible else 'Masquer la provenance')

    def context(self, index, event):
        self.select(index)
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()
        return 'break'

    def copy(self, kind):
        if self.selected_index is None:
            return 'break'
        path, modbus = self.errors[self.selected_index]
        text = (os.fspath(path) if kind=='path' else str(modbus) if kind=='id' else
                f'{modbus}\t{compact_trace_label(path, self.app._selected_directories)}')
        self.clipboard_clear()
        self.clipboard_append(text)
        return 'break'

    def _paint_detail(self):
        colors = self.app._palette
        for card in (self.positive, self.waiting):
            card.apply_theme(colors)
        for row in self._step_rows:
            row.apply_theme(colors)

    def apply_theme(self):
        colors = self.app._palette
        self.configure(fg_color=colors['bg'])
        self.split.apply_theme(colors)
        self.header.apply_theme(colors)
        for panel in (self.list_panel, self.detail_panel):
            panel.configure(fg_color=colors['surface'], bg_color=colors['bg'])
        self.detail_scroll.configure(fg_color=colors['surface'], bg_color=colors['surface'])
        for frame in (self.states, self.steps_frame):
            frame.configure(fg_color=colors['surface'], bg_color=colors['surface'])
        self.steps_heading.configure(fg_color=colors['surface'], bg_color=colors['surface'])
        for label in (self.message, self.path_label):
            label.configure(bg_color=colors['surface'])
        for label in (self.title, self.detail_title, self.steps_heading):
            label.configure(text_color=colors['text'])
        for label in (self.summary, self.list_heading, self.detail_source, self.message, self.path_label):
            label.configure(text_color=colors['muted'])
        for button in (self.path_button, self.provenance_button):
            button.configure(fg_color='transparent', text_color=colors['muted'], hover_color=colors['raised'])
        for button in (self.close_button, self.retry):
            button.configure(fg_color=colors['raised'], text_color=colors['text'], hover_color=colors['border'])
        self.detail_scroll.configure(scrollbar_button_color=colors['border'], scrollbar_button_hover_color=colors['muted'])
        for label in (self.detail_title, self.detail_source, self.message, self.path_label):
            label.sync_selection_text()
        self.cards.apply_theme()
        self._paint_detail()
