"""Persistent conveyor workspace. All queries and destinations retain exact paths."""
import tkinter as tk

import customtkinter as ctk

from background_tasks import BackgroundTasks
from convoyage_errors_view import ConvoyageErrorsView
from ui_components import ActionButton, register_view


class ConvoyageWorkspace(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent, corner_radius=0)
        self.app, self._theme_deferred = app, True
        self.tasks = BackgroundTasks()
        self.active_tab, self.query, self._closed = 'parcours', '', False
        self.pages, self.buttons, self.views = {}, {}, {}
        self.states = {name: dict(loaded=False, revision=None, pending=None, error=False)
                       for name in ('parcours', 'anomalies')}
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self.heading = ctk.CTkFrame(self, fg_color='transparent')
        self.heading.grid(row=0, column=0, sticky='ew', pady=(12, 16))
        self.title = ctk.CTkLabel(self.heading, text='Convoyage', font=('Segoe UI', 22, 'bold'))
        self.title.pack(side='left')
        self.badge = ctk.CTkLabel(self.heading, text='MODBUS · R1/R2', corner_radius=6,
                                 font=('Segoe UI', 12, 'bold'), height=26)
        self.badge.pack(side='left', padx=16, ipadx=10)
        self.tabs = ctk.CTkFrame(self, corner_radius=8)
        self.tabs.grid(row=1, column=0, sticky='w', pady=(0, 16))
        self.content = ctk.CTkFrame(self, corner_radius=0, fg_color='transparent')
        self.content.grid(row=2, column=0, sticky='nsew')
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)
        for name, caption in (('parcours', 'Parcours'), ('anomalies', 'Anomalies')):
            button = ActionButton(self.tabs, text=caption, text_size=14, width=140,
                                  command=lambda value=name: self.show_tab(value))
            button.pack(side='left', padx=4, pady=4)
            self.buttons[name] = button
            page = ctk.CTkFrame(self.content, fg_color='transparent', corner_radius=0)
            page.grid(row=0, column=0, sticky='nsew')
            page.grid_columnconfigure(0, weight=1)
            page.grid_rowconfigure(2, weight=1)
            toolbar = ctk.CTkFrame(page, corner_radius=10)
            toolbar.grid(row=0, column=0, sticky='ew', pady=(0, 12))
            toolbar.grid_columnconfigure(1, weight=1)
            status = ctk.CTkLabel(page, text='', anchor='w', justify='left', font=('Segoe UI', 12))
            status.grid(row=1, column=0, sticky='ew', padx=2, pady=(0, 8))
            body = ctk.CTkFrame(page, corner_radius=12)
            body.grid(row=2, column=0, sticky='nsew')
            body.grid_columnconfigure(0, weight=1)
            body.grid_rowconfigure(0, weight=1)
            empty = ctk.CTkFrame(body, fg_color='transparent')
            empty.grid(row=0, column=0, padx=24, pady=24)
            empty_title = ctk.CTkLabel(empty, text='', font=('Segoe UI', 20, 'bold'))
            empty_title.pack()
            empty_message = ctk.CTkLabel(empty, text='', font=('Segoe UI', 14),
                                          wraplength=440, justify='center')
            empty_message.pack(pady=(10, 0))
            retry = ActionButton(empty, text='Réessayer', text_size=14,
                                 command=lambda value=name: self.load(value, force=True))
            self.pages[name] = dict(frame=page, toolbar=toolbar, status=status, body=body,
                                    empty=empty, title=empty_title, message=empty_message, retry=retry)
        bar = self.pages['parcours']['toolbar']
        self.id_label = ctk.CTkLabel(bar, text='ModBusId', font=('Segoe UI', 14))
        self.id_label.grid(row=0, column=0, sticky='w', padx=(16, 12), pady=16)
        self.entry = ctk.CTkEntry(bar, placeholder_text='Saisir un identifiant', height=36,
                                  corner_radius=6, font=('Segoe UI', 14), width=280)
        self.entry.grid(row=0, column=1, sticky='ew', pady=16, padx=(0, 12))
        self.search_button = ActionButton(bar, text='Afficher le parcours', text_size=14,
                                         role='primary', width=180, command=self.search)
        self.search_button.grid(row=0, column=2, padx=(0, 16))
        self.entry.bind('<Return>', self.search)
        bar = self.pages['anomalies']['toolbar']
        self.count = ctk.CTkLabel(bar, text='Anomalies de convoyage', font=('Segoe UI', 14, 'bold'))
        self.count.grid(row=0, column=0, sticky='w', padx=16, pady=16)
        self.refresh_button = ActionButton(bar, text='Actualiser', text_size=14,
                                          command=lambda: self.load('anomalies', force=True))
        self.refresh_button.grid(row=0, column=2, padx=16)
        self._empty('parcours', 'Suivre une boîte',
                    'Saisissez son ModBusId pour consulter les positions et les attentes enregistrées.')
        self._empty('anomalies', 'Anomalies de convoyage',
                    'Les arrivées non retrouvées seront affichées ici.')
        register_view(app, self)
        tk.Misc.bind(self, '<Destroy>', self._destroyed, add='+')
        tk.Misc.bind(self, '<Configure>', self._resize, add='+')
        self.show_tab('parcours')
        self.apply_theme()
        self._timer = self.after(80, self._poll)

    def _resize(self, event):
        width = event.width/self._get_widget_scaling()
        wrap = max(220, round(width)-8)
        if wrap!=getattr(self, '_message_wrap', None):
            self._message_wrap = wrap
            for page in self.pages.values():
                page['status'].configure(wraplength=wrap)
                page['message'].configure(wraplength=min(440, max(200, wrap-48)))
        narrow = width<610
        if narrow==getattr(self, '_narrow', None):
            return
        self._narrow = narrow
        self.entry.grid_configure(columnspan=2 if narrow else 1, padx=(0, 16) if narrow else (0, 12))
        self.search_button.grid_configure(row=1 if narrow else 0, column=1 if narrow else 2,
                                         sticky='w' if narrow else '', pady=(0, 14) if narrow else 0)

    def _empty(self, tab, title, message, retry=False):
        page = self.pages[tab]
        if tab in self.views:
            self.views[tab].grid_remove()
        page['empty'].grid()
        page['title'].configure(text=title)
        page['message'].configure(text=message)
        page['retry'].pack(pady=(16, 0)) if retry else page['retry'].pack_forget()

    def show_tab(self, tab):
        self.active_tab = tab
        self.pages[tab]['frame'].tkraise()
        for name, button in self.buttons.items():
            button.configure(role='primary' if name==tab else 'quiet')
        self.observe_index()

    def search(self, event=None):
        query = self.entry.get().strip()
        if not query or not query.isascii() or not query.isdecimal():
            self.pages['parcours']['status'].configure(text='Saisissez un ModBusId numérique pour afficher son parcours.')
            return 'break'
        if query!=self.query:
            self.query = query
            self.states['parcours']['loaded'] = False
            self.states['parcours']['error'] = False
            # A new ID must never leave the previous file as a copy destination.
            if 'parcours' in self.views:
                self.views['parcours'].show_empty('Chargement du parcours', '')
        self.load('parcours')
        return 'break'

    def observe_index(self):
        if self._closed or getattr(self.app, '_active_panel', None)!='convoyage':
            return
        tab = self.active_tab
        if tab=='anomalies':
            ready = getattr(self.app, '_indexed_complete', False) and getattr(self.app, '_sources_verified', False)
            self.refresh_button.configure(state='normal' if ready else 'disabled')
            if not ready:
                if not getattr(self, '_anomalies_waiting', False):
                    self.pages[tab]['status'].configure(text='Vérification des sources en cours — contrôle disponible à la fin.')
                    self._anomalies_waiting = True
                return
            if getattr(self, '_anomalies_waiting', False):
                self._anomalies_waiting = False
                self.states[tab]['revision'] = None
        if tab=='parcours' and not self.query:
            return
        state = self.states[tab]
        revision = getattr(self.app, '_index_content_revision', 0)
        if not state['error'] and (not state['loaded'] or state['revision']!=revision):
            self.load(tab)

    def load(self, tab, force=False):
        state, page = self.states[tab], self.pages[tab]
        if tab=='parcours' and not self.query:
            return
        if tab=='anomalies' and not (getattr(self.app, '_indexed_complete', False) and
                                     getattr(self.app, '_sources_verified', False)):
            self.observe_index()
            return
        store = getattr(self.app, 'store', None)
        if store is None:
            page['status'].configure(text='Ouverture de l’index en cours…')
            return
        revision = getattr(self.app, '_index_content_revision', 0)
        query = self.query
        preserve = state['loaded']
        key = revision, query if tab=='parcours' else None
        if not force and not state['error'] and (state['pending']==key or (state['loaded'] and state['revision']==revision)):
            return
        state['pending'], state['error'] = key, False
        page['status'].configure(text='Actualisation en cours…' if state['loaded'] else 'Chargement en cours…')
        if not state['loaded']:
            self._empty(tab, 'Chargement du parcours' if tab=='parcours' else 'Vérification des arrivées',
                        'Lecture des données indexées…')
        def work(cancel):
            with store.snapshot():
                with store.connection() as db:
                    db.set_progress_handler(lambda: int(cancel.is_set()), 1000)
                    if tab=='anomalies':
                        return store.convoyage_errors()
                    return [(row['_filepath'], query) for row in store.convoyage_summary(query)]
        def done(items):
            state.update(loaded=True, revision=revision, pending=None, error=False)
            files = len({path for path, _ in items})
            files_label = f'{files} fichier'+('s' if files!=1 else '')
            page['status'].configure(text=(f'ModBusId {query} · {files_label}' if tab=='parcours' else
                                           'Contrôle terminé · sources sélectionnées'))
            if tab=='anomalies':
                self.count.configure(text=f'{len(items)} anomalie'+('s' if len(items)!=1 else '')+f' · {files_label}')
            if items:
                view = self.views.get(tab)
                if view is None:
                    view = self.views[tab] = ConvoyageErrorsView(page['body'], self.app, items, embedded=True)
                else:
                    view.reset_items(items, preserve=preserve, force=True)
                view.list_heading.configure(text='FICHIERS' if tab=='parcours' else 'À EXAMINER')
                page['empty'].grid_remove()
                view.grid(row=0, column=0, sticky='nsew')
            else:
                if tab in self.views:
                    self.views[tab].show_empty('Aucun résultat', '')
                self._empty(tab, 'Aucune position retrouvée' if tab=='parcours' else 'Aucune anomalie détectée',
                            f'Aucune position indexée pour le ModBusId {query} dans les sources sélectionnées.'
                            if tab=='parcours' else 'Toutes les arrivées sont retrouvées dans les fichiers contrôlés.')
        def failed(error):
            state.update(pending=None, error=True)
            page['status'].configure(text=f'Chargement interrompu : {error}')
            # A retry remains available even when the old validated results are visible.
            if not state['loaded']:
                self._empty(tab, 'Données indisponibles', 'Le chargement n’a pas pu aboutir.', retry=True)
            else:
                retry = 'Cliquez sur Actualiser pour réessayer.' if tab=='anomalies' else 'Relancez le parcours pour réessayer.'
                page['status'].configure(text=f'Actualisation impossible : {error}. '
                                               f'Les résultats précédents restent affichés. {retry}')
        self.tasks.submit(tab, work, done, failed)

    def _poll(self):
        if self._closed:
            return
        self.tasks.drain()
        self.observe_index()
        self._timer = self.after(80, self._poll)

    def _destroyed(self, event):
        if event.widget is self:
            self._closed = True
            self.tasks.close()
            self.after_cancel(self._timer)

    def apply_theme(self):
        colors = self.app._palette
        self.configure(fg_color=colors['bg'])
        self.title.configure(text_color=colors['text'])
        self.badge.configure(fg_color=colors['raised'], text_color=colors['accent'])
        self.tabs.configure(fg_color=colors['raised'])
        for button in (*self.buttons.values(), self.search_button, self.refresh_button):
            button.apply_theme(colors)
        self.id_label.configure(text_color=colors['text'])
        self.entry.configure(fg_color=colors['bg'], text_color=colors['text'], border_color=colors['border'])
        self.count.configure(text_color=colors['text'])
        for page in self.pages.values():
            for name in ('body', 'toolbar'):
                page[name].configure(fg_color=colors['surface'])
            page['title'].configure(text_color=colors['text'])
            page['message'].configure(text_color=colors['muted'])
            page['status'].configure(text_color=colors['muted'])
            page['retry'].apply_theme(colors)
