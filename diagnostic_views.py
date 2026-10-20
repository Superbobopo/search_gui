"""Shared diagnostic shells for measurement comparison and conveyor checkpoints."""
from datetime import datetime
import json

import customtkinter as ctk

from ui_components import WindowShell, StateCard, ActionButton, selectable, font, style_labels
from trace_sources import compact_trace_label, trace_tooltip
from trace_store import CHECKPOINTS
from convoyage_errors_view import _StepRow


def dimensions(event):
    values = event.get('dims') or {}
    return json.loads(values) if isinstance(values, str) else values


def measurement_text(event):
    values = dimensions(event)
    return '    '.join(f'{key} = {values.get(key, "—")}' for key in ('l', 'H', 'L'))


def event_time(event):
    return datetime.fromtimestamp(event['time']).strftime('%d/%m/%Y %H:%M:%S') if event.get('time') is not None else 'Horodatage indisponible'


class ComparisonView(WindowShell):
    def __init__(self, parent, app, reference, pair, context=None, choose=None):
        if isinstance(pair, dict):
            context, pair = pair, pair.get('pair')
        self.context = context or {}
        super().__init__(parent, app, 'Comparaison IDPK', reference)
        self.scroll = ctk.CTkScrollableFrame(self.body, corner_radius=0)
        self.scroll.grid(row=0, column=0, sticky='nsew', padx=(14, 6), pady=14)
        self.scroll.grid_columnconfigure(0, weight=1)
        self.labels, self.cards, self.deltas, self.buttons = [], [], [], []
        identity = self.context.get('identity') or (reference, None, None)
        self._label(f'Référence : {identity[0]} · SN : {identity[1] or "—"} · CodeLot : {identity[2] or "—"}', bold=True)
        if pair:
            current, previous = pair
            for title, event, role in (
                ('Mesure actuelle', current, 'waiting' if self.context.get('different') else 'info'),
                (self.context.get('basis_label') or 'Dernière mesure comparable', previous, 'neutral')):
                card = StateCard(self.scroll, title)
                card.grid(row=len(self.labels)+len(self.cards), column=0, sticky='ew', padx=4, pady=(8, 2))
                card.show(role, measurement_text(event), event_time(event))
                self.cards.append(card)
                self._label(compact_trace_label(event['path'], app._selected_directories), small=True)
            self._label('Écarts des dimensions', bold=True)
            a, b = dimensions(current), dimensions(previous)
            for key in ('l', 'H', 'L'):
                if key in a and key in b:
                    delta = a[key]-b[key]
                    label = self._label(f'Δ {key} = {delta:+g}')
                    self.deltas.append((label, abs(delta)>2))
                else:
                    self._label(f'Δ {key} : dimension non comparable', small=True)
            self._label('Mesures originales', bold=True)
            for event in pair:
                self._label(event.get('text') or 'Ligne originale indisponible.', raw=True)
        else:
            current = self.context.get('current')
            if current:
                self._label('Mesure actuelle : '+measurement_text(current), bold=True)
            self._label(self.context.get('message') or 'Comparaison indisponible.')
            for candidate in self.context.get('candidates', []):
                button = ActionButton(self.scroll, text=compact_trace_label(candidate['path'], app._selected_directories)+' · '+event_time(candidate),
                    command=lambda value=candidate: choose(current, value) if choose else None)
                button.grid(row=len(self.labels)+len(self.cards)+len(self.buttons), column=0, sticky='ew', padx=4, pady=4)
                self.buttons.append(button)
        self._width = None
        self.scroll.bind('<Configure>', self._resize)
        self.apply_theme()

    def _label(self, text, bold=False, small=False, raw=False):
        label = selectable(self.scroll, text=text, font=('Consolas', 11) if raw else font('meta' if small else 'body', bold), wraplength=600)
        label.grid(row=len(self.labels)+len(self.cards), column=0, sticky='ew', padx=6, pady=6)
        self.labels.append(label)
        return label

    def _resize(self, event):
        if event.width!=self._width:
            self._width = event.width
            for label in self.labels+[card.value for card in self.cards]:
                label.configure(wraplength=max(100, event.width-36))

    def apply_theme(self):
        WindowShell.apply_theme(self)
        if not hasattr(self, 'scroll'):
            return
        colors = self.app._palette
        self.scroll.configure(fg_color=colors['surface'], bg_color=colors['surface'], scrollbar_button_color=colors['border'])
        style_labels(self.labels, colors)
        for card in self.cards:
            card.apply_theme(colors)
        for label, different in self.deltas:
            style_labels([label], colors, 'error_text' if different else 'muted')
        for button in self.buttons:
            button.apply_theme(colors)


class ConvoyageRecapView(WindowShell):
    def __init__(self, parent, app, reference, recap, paths=()):
        super().__init__(parent, app, 'Parcours détaillé du ModBusId '+str(reference), 'Étapes enregistrées dans les fichiers MODBUS · R1/R2')
        self.scroll = ctk.CTkScrollableFrame(self.body, corner_radius=0)
        self.scroll.grid(row=0, column=0, sticky='nsew', padx=(16, 6), pady=14)
        self.scroll.grid_columnconfigure(0, weight=1)
        self.rows, self.sources = [], []
        for col, name in CHECKPOINTS.items():
            item = recap[col]
            if not item['etat']:
                continue
            row = _StepRow(self.scroll)
            row.grid(row=len(self.rows)*2, column=0, sticky='ew')
            row.show(dict(item, checkpoint=name))
            self.rows.append(row)
            source = selectable(self.scroll, text=item.get('file') or 'Provenance indisponible', font=font('meta'), wraplength=600)
            source.grid(row=len(self.rows)*2-1, column=0, sticky='ew', padx=(28, 6), pady=(0, 8))
            self.sources.append(source)
        self.origins = ctk.CTkFrame(self.scroll, corner_radius=0)
        self.origins.grid_columnconfigure(0, weight=1)
        for index, path in enumerate(paths):
            label = selectable(self.origins, text=trace_tooltip(path), font=font('meta'), wraplength=600)
            label.grid(row=index, column=0, sticky='ew', pady=8)
            self.sources.append(label)
        self.origins_toggle = ActionButton(self.scroll, role='quiet', text='Afficher les sources consultées',
                                           command=self.toggle_origins)
        self.origins_toggle.grid(row=len(self.rows)*2, column=0, sticky='w', pady=12)
        self.origins_row = len(self.rows)*2+1
        self.scroll.bind('<Configure>', self._resize)
        self._width = None
        self.apply_theme()

    def toggle_origins(self):
        if self.origins.winfo_manager():
            self.origins.grid_remove()
            self.origins_toggle.configure(text='Afficher les sources consultées')
        else:
            self.origins.grid(row=self.origins_row, column=0, sticky='ew', padx=6)
            self.origins_toggle.configure(text='Masquer les sources consultées')

    def _resize(self, event):
        if event.width!=self._width:
            self._width = event.width
            for row in self.rows:
                row.checkpoint.configure(wraplength=max(100, event.width-150))
            for label in self.sources:
                label.configure(wraplength=max(100, event.width-42))

    def apply_theme(self):
        WindowShell.apply_theme(self)
        if not hasattr(self, 'scroll'):
            return
        colors = self.app._palette
        self.scroll.configure(fg_color=colors['surface'], bg_color=colors['surface'], scrollbar_button_color=colors['border'])
        self.origins.configure(fg_color=colors['surface'], bg_color=colors['surface'])
        self.origins_toggle.apply_theme(colors)
        for row in self.rows:
            row.apply_theme(colors)
        style_labels(self.sources, colors, 'muted')
