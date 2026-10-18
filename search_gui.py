import os
import threading
import tkinter as tk
from tkinter import ttk
import customtkinter as ctk
import re
import datetime
from tkinter import messagebox, scrolledtext

from style_config import configure_style
from ui_theme import LIGHT, DARK, palette_for, NavigationButton, navigation_icon, apply_widget_theme
from indexed_ui import IndexedAppMixin
from app_assets import apply_icon
from trace_store import CHECKPOINTS
import json



# Sources installées et locales, partagées par toutes les vues.
from trace_sources import (
    discover_trace_files, trace_label, build_next_file_map,
    path_key, same_trace_directory,
)
from trace_parser import (
    read_trace, parse_fields, field_matches, timestamp_parts, format_ticket_fields,
    is_gripper, parse_idpk, find_idpk_after, is_idpk_for, parse_modbus_line,
    trace_start_datetime, trace_end_datetime, line_datetime, iter_timed_lines,
    missing_robot_tickets,
)
TRACE_LOCK = threading.RLock()
BUFFER = {}
current_file = None
BUFFER_MTIME = {}
BUFFER_DATE = {}
BUFFER_START = {}

def trace_file_datetime(path):
    if path in BUFFER_DATE:
        return BUFFER_DATE[path]
    fallback = datetime.datetime.fromtimestamp(os.path.getmtime(path))
    return trace_end_datetime(read_trace(path), fallback)


def trace_file_sort_key(path):
    return trace_file_datetime(path), path_key(path)


def trace_line_datetime(path, line):
    return line_datetime(line, BUFFER_START.get(path, trace_file_datetime(path)))


class ToolTip:
    """
    Affiche une petite bulle d'aide (tooltip) lorsque la souris passe
    sur un widget, sans recouvrir le widget lui-même.
    """
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        # si déjà affiché ou pas de texte, on n'affiche pas
        if self.tipwindow or not self.text:
            return
        # calcule la position : tout à droite du widget, un petit offset
        x = self.widget.winfo_rootx() + self.widget.winfo_width() + 10
        y = self.widget.winfo_rooty() + 5
        # crée la fenêtre sans bordure et toujours au-dessus
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_attributes("-topmost", True)
        tw.wm_geometry(f"+{x}+{y}")
        # le label de la bulle
        label = tk.Label(
            tw, text=self.text, justify="left",
            background="#ffffe0", relief="solid", borderwidth=1,
            font=("Segoe UI", 9)
        )
        label.pack(ipadx=1, ipady=1)

    def hide_tip(self, event=None):
        if self.tipwindow:
            self.tipwindow.destroy()
        self.tipwindow = None



def reload_modified_files():
    """Refresh all available sources and report whether cached contents changed."""
    global BUFFER, BUFFER_MTIME, BUFFER_DATE, BUFFER_START
    with TRACE_LOCK:
        new_buffer = {}
        new_mtime = {}
        new_date = {}
        new_start = {}
        paths = discover_trace_files()
        for path in paths:
            try:
                mtime = os.stat(path).st_mtime_ns
                if path in BUFFER and BUFFER_MTIME.get(path) == mtime:
                    text = BUFFER[path]
                else:
                    text = read_trace(path)
                new_buffer[path] = text
                new_mtime[path] = mtime
                if BUFFER_MTIME.get(path) == mtime and path in BUFFER_DATE:
                    new_date[path] = BUFFER_DATE[path]
                    new_start[path] = BUFFER_START[path]
                else:
                    fallback = datetime.datetime.fromtimestamp(mtime / 1_000_000_000)
                    new_start[path] = trace_start_datetime(text, fallback)
                    new_date[path] = trace_end_datetime(text, fallback)
            except OSError:
                pass
        changed = new_mtime != BUFFER_MTIME
        BUFFER = new_buffer
        BUFFER_MTIME = new_mtime
        BUFFER_DATE = new_date
        BUFFER_START = new_start
        return changed


def load_files_into_buffer():
    reload_modified_files()
    return dict(BUFFER)


def find_lines_with_keyword(text, keyword, exact_modbus=False):
    lines = text.splitlines()
    if not exact_modbus:
        key_lower = keyword.lower()
        return [line for line in lines if key_lower in line.lower()]
    else:
        # keyword du style ModBusId=124
        m = re.match(r"ModBusId=(\d+)", keyword)
        if m:
            modbus_id = m.group(1)
            # \D ou fin de ligne pour ne pas matcher ModBusId=1241 si on cherche 124
            return [line for line in lines if field_matches(line, "ModBusId", modbus_id)]
        else:
            return []


def search_file_for_keyword(path, keyword, exact_modbus=False):
    text = BUFFER.get(path, "")
    return find_lines_with_keyword(text, keyword, exact_modbus)


def extract_columns_from_line(line, indices, column_names=None):
    valeurs = parse_modbus_line(line)
    extraites = [valeurs[i] if i < len(valeurs) else '' for i in indices]
    if column_names:
        return {column_names[n]: extraites[n] for n in range(len(indices))}
    else:
        return extraites




class SearchApp(IndexedAppMixin, ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        ctk.ThemeManager.theme['CTkButton'].update(
            fg_color=(LIGHT['accent'], DARK['accent']),
            hover_color=(LIGHT['hover'], DARK['hover']),
            text_color=(LIGHT['accent_text'], DARK['accent_text']))
        self.theme_mode = tk.StringVar(value="Système")
        self.result_mode = tk.StringVar(value="Classique")
        self.search_mode = tk.StringVar(value="id")
        self._palette = palette_for(ctk.get_appearance_mode())
        self._appearance = ctk.get_appearance_mode()
        self._active_panel = None
        self._view_queries = {}
        self._context_path = None
        self._theme_pending = None
        self.title("Recherche TracOmega")
        scale = self._get_window_scaling()
        available_width = max(540, int(self.winfo_screenwidth() / scale) - 60)
        available_height = max(420, int(self.winfo_screenheight() / scale) - 80)
        self.geometry(f"{min(1100, available_width)}x{min(800, available_height)}")
        self.minsize(min(860, available_width), min(560, available_height))
        self.configure(fg_color=(LIGHT['bg'], DARK['bg']))
        self.resultats = []

        self.sidebar = ctk.CTkFrame(self, fg_color=(LIGHT['sidebar'], DARK['sidebar']),
                                    width=184, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        ctk.CTkLabel(self.sidebar, text="TracOmega", text_color=("#e5edf8", "#e5edf8"),
                     font=("Segoe UI", 20, "bold")).pack(anchor="w", padx=18, pady=(22, 0))
        ctk.CTkLabel(self.sidebar, text="Analyse des traces", text_color=("#aab8cc", "#aab8cc"),
                     font=("Segoe UI", 11)).pack(anchor="w", padx=18, pady=(0, 24))
        nav_options = dict(corner_radius=6, fg_color="transparent", hover_color="#294367",
                           text_color=("#e5edf8", "#e5edf8"), anchor="w", height=38,
                           font=("Segoe UI", 12), border_width=0)
        self.btn_search_menu = NavigationButton(self.sidebar, text="         Recherche",
                                            command=self.show_search_panel, **nav_options)
        self.btn_tracking_robot = NavigationButton(self.sidebar, text="         Suivi référence",
                                               command=self.show_tracking_robot_panel, **nav_options)
        self.btn_convoyage_menu = NavigationButton(self.sidebar, text="         Convoyage",
                                               command=self.show_convoyage_tools, **nav_options)
        for button in (self.btn_search_menu, self.btn_tracking_robot, self.btn_convoyage_menu):
            button.pack(fill="x", padx=10, pady=3)
        self._nav_icons = {name: navigation_icon(button, name) for name, button in
                           (('search', self.btn_search_menu), ('tracking', self.btn_tracking_robot),
                            ('convoyage', self.btn_convoyage_menu))}
        theme_box = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        theme_box.pack(side="bottom", fill="x", padx=12, pady=18)
        ctk.CTkLabel(theme_box, text="Apparence", text_color=("#aab8cc", "#aab8cc"),
                     font=("Segoe UI", 11)).pack(anchor="w", padx=4)
        self.theme_menu = ctk.CTkOptionMenu(theme_box, values=["Système", "Clair", "Sombre"],
                                           variable=self.theme_mode, command=self._change_theme,
                                           fg_color="#294367", button_color="#294367",
                                           text_color=("#e5edf8", "#e5edf8"), font=("Segoe UI", 12))
        self.theme_menu.pack(fill="x")

        self.workspace = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        self.workspace.pack(side="left", fill="both", expand=True)
        self.search_bar = ctk.CTkFrame(self.workspace, fg_color=(LIGHT['surface'], DARK['surface']), corner_radius=8)
        self.search_bar.pack(side="top", fill="x", padx=12, pady=(12, 8))
        self.search_bar.grid_columnconfigure(1, weight=1)
        self.keyword_label = ctk.CTkLabel(self.search_bar, text="Recherche", font=("Segoe UI", 11))
        self.keyword_label.grid(row=0, column=0, padx=(12, 8), pady=12)
        self.keyword = ctk.CTkEntry(self.search_bar, font=("Segoe UI", 13), corner_radius=5,
                                   width=120, height=34, placeholder_text="ID, référence ou texte")
        self.keyword.grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=12)
        self.tracking_keyword = self.keyword
        self.modbus = tk.BooleanVar(value=True)
        self.reference = tk.BooleanVar(value=False)
        self.mode_frame = ctk.CTkFrame(self.search_bar, fg_color="transparent")
        self.mode_frame.grid(row=0, column=2, padx=(0, 6))
        self.mode_radios = {}
        for column, (value, label) in enumerate((('id', 'ModBusId'), ('ref', 'Référence'), ('text', 'Texte'))):
            radio = ctk.CTkRadioButton(self.mode_frame, text=label, variable=self.search_mode, value=value,
                                        command=self._choose_search_mode, font=("Segoe UI", 11),
                                        radiobutton_width=16, radiobutton_height=16,
                                        width=82 if value != 'text' else 58,
                                        fg_color=(LIGHT['accent'], DARK['accent']),
                                        hover_color=(LIGHT['hover'], DARK['hover']))
            radio.grid(row=0, column=column, padx=2)
            self.mode_radios[value] = radio
        self.btn_gear = ctk.CTkButton(self.search_bar, text="Filtres", width=68, height=34,
                                     command=self.toggle_advanced_search, font=("Segoe UI", 11),
                                     fg_color=(LIGHT['raised'], DARK['raised']),
                                     hover_color=(LIGHT['border'], DARK['border']),
                                     text_color=(LIGHT['text'], DARK['text']))
        self.btn_gear.grid(row=0, column=3, padx=(0, 6))
        self.btn_search = ctk.CTkButton(self.search_bar, text="Rechercher", width=96, height=34,
                                       command=self._submit_visible_search, font=("Segoe UI", 12),
                                       fg_color=(LIGHT['accent'], DARK['accent']),
                                       hover_color=(LIGHT['hover'], DARK['hover']),
                                       text_color=(LIGHT['accent_text'], DARK['accent_text']))
        self.btn_search.grid(row=0, column=4, padx=(0, 12))
        self.keyword.bind('<Return>', lambda event: self._submit_visible_search())
        self.action_bar = ctk.CTkFrame(self.search_bar, fg_color="transparent")
        self.action_bar.grid(row=1, column=0, columnspan=5, sticky="ew", padx=12, pady=(0, 10))
        self.occurrence_caption = ctk.CTkLabel(self.action_bar, text="Occurrence", font=("Segoe UI", 11))
        self.occurrence_caption.pack(side="left", padx=(0, 6))
        secondary = dict(height=30, corner_radius=5, font=("Segoe UI", 11),
                         fg_color=(LIGHT['raised'], DARK['raised']),
                         hover_color=(LIGHT['border'], DARK['border']),
                         text_color=(LIGHT['text'], DARK['text']))
        self.btn_up = ctk.CTkButton(self.action_bar, text="↑", width=30, command=self.on_up, **secondary)
        self.btn_down = ctk.CTkButton(self.action_bar, text="↓", width=30, command=self.on_down, **secondary)
        self.btn_compare_idpk = ctk.CTkButton(self.action_bar, text="Comparer IDPK", width=108,
                                             command=self.compare_previous_idpk_popup, **secondary)
        self.btn_adjacent = ctk.CTkButton(self.action_bar, text="IDs adjacents", width=108,
                                         command=self.open_adjacent_for_current_occurrence, **secondary)
        self.btn_open_trace = ctk.CTkButton(self.action_bar, text="Ouvrir trace", width=96,
                                           command=self.open_file_view, **secondary)
        self.btn_erreurs = ctk.CTkButton(self.action_bar, text="Erreurs Robot", width=108,
                                        command=self.check_errors, state='disabled', **secondary)
        for button in (self.btn_up, self.btn_down, self.btn_compare_idpk, self.btn_adjacent,
                       self.btn_open_trace, self.btn_erreurs):
            button.pack(side="left", padx=(0, 5))
        ToolTip(self.btn_erreurs, "Vérifie les prises du robot après l’arrivée sur T5.\nDisponible après l’indexation complète.")
        self.advanced_frame = ctk.CTkFrame(self.workspace, fg_color=(LIGHT['surface'], DARK['surface']))
        self.main_panel = ctk.CTkFrame(self.workspace, fg_color="transparent", corner_radius=0)
        self.main_panel.pack(side="top", fill="both", expand=True, padx=12, pady=(0, 8))
        self.show_search_panel()
        self.bind('<Control-f>', self._focus_search, add='+')
        self.search_bar.bind('<Configure>', self._adapt_command_layout, add='+')
        self.bind_all('<Map>', self._theme_new_window, add='+')
        self._apply_ui_theme()
        self.after(1000, self._watch_system_theme)

        # Async load buffers
        import threading
        self.files_loaded = False

        # Vars internes
        self.lines = []
        self.line_index = 0
        self.file_list = []
        self.file_index = 0

        # Adjacent popup placeholders
        self.adj_popup = None
        self.adj_prev_label = None
        self.adj_next_label = None
        self.title("Recherche TracOmega (chargement...)")
        self._start_indexed_service()

    def _theme_window(self, window):
        if window.winfo_exists():
            apply_widget_theme(window, self._palette)

    def _apply_ui_theme(self):
        self._appearance = ctk.get_appearance_mode()
        self._palette = palette_for(self._appearance)
        configure_style(self, self._palette)
        self._theme_window(self)
        self._set_active_navigation()
        frame = getattr(self, 'summary_frame', None)
        if frame is not None and frame.winfo_exists() and hasattr(self, '_render_summary'):
            selection = getattr(self, 'selection', None)
            self._render_summary(selection.details if selection else None)

    def _change_theme(self, value):
        self.theme_mode.set(value)
        ctk.set_appearance_mode({'Clair': 'Light', 'Sombre': 'Dark', 'Système': 'System'}[value])
        self._apply_ui_theme()

    def _watch_system_theme(self):
        if self.winfo_exists():
            if ctk.get_appearance_mode() != self._appearance:
                self._apply_ui_theme()
            self.after(1000, self._watch_system_theme)

    def _theme_new_window(self, event):
        if isinstance(event.widget, (tk.Tk, tk.Toplevel, ctk.CTkButton, ctk.CTkRadioButton,
                                     ctk.CTkOptionMenu)) and self._theme_pending is None:
            # CTk constructors drain idle callbacks before their internal widgets exist.
            self._theme_pending = self.after(1, self._theme_mapped_windows)

    def _theme_mapped_windows(self):
        self._theme_pending = None
        if self.winfo_exists():
            self._theme_window(self)
            self._set_active_navigation()

    def _set_active_navigation(self):
        for name, button in (('search', self.btn_search_menu), ('tracking', self.btn_tracking_robot),
                             ('convoyage', self.btn_convoyage_menu)):
            selected = name == self._active_panel
            button.configure(fg_color='#294367' if selected else 'transparent',
                             border_width=1 if selected else 0,
                             border_color='#35d6c6' if selected else '#294367')
            self._nav_icons[name].configure(bg='#294367' if selected else self._palette['sidebar'])

    def _adapt_command_layout(self, event=None):
        """Keep search and action rows usable in a narrow/high-DPI workspace."""
        width = self.search_bar.winfo_width() / self.search_bar._get_widget_scaling()
        compact = width < 820
        if compact == getattr(self, '_compact_commands', None):
            return
        self._compact_commands = compact
        self.keyword_label.grid_remove() if compact else self.keyword_label.grid()
        self.keyword.grid_configure(column=0 if compact else 1, columnspan=2 if compact else 1,
                                    padx=(10, 6) if compact else (0, 10))
        self.search_bar.grid_columnconfigure(0, weight=1 if compact else 0)
        self.search_bar.grid_columnconfigure(1, weight=0 if compact else 1)
        for value, radio in self.mode_radios.items():
            radio.configure(text={'id': 'ID', 'ref': 'Réf.', 'text': 'Texte'}[value]
                            if compact else {'id': 'ModBusId', 'ref': 'Référence', 'text': 'Texte'}[value],
                            width=46 if compact else 82 if value != 'text' else 58)
        self.btn_search.configure(width=82 if compact else 96)
        self.btn_gear.configure(width=58 if compact else 68)
        self.occurrence_caption.pack_forget() if compact else self.occurrence_caption.pack(
            side='left', before=self.btn_up, padx=(0, 6))
        for button, normal, small in ((self.btn_compare_idpk, 'Comparer IDPK', 'IDPK'),
                                      (self.btn_adjacent, 'IDs adjacents', 'Voisins'),
                                      (self.btn_open_trace, 'Ouvrir trace', 'Trace'),
                                      (self.btn_erreurs, 'Erreurs Robot', 'Erreurs Robot')):
            button.configure(text=small if compact else normal, width=82 if compact else
                             96 if button is self.btn_open_trace else 108)

    def _activate_panel(self, name):
        old = self._active_panel
        if old and old != name:
            self._view_queries[old] = self.keyword.get()
        if old != name:
            self.keyword.delete(0, 'end')
            self.keyword.insert(0, self._view_queries.get(name, ''))
        self._active_panel = name
        self._set_active_navigation()
        self.keyword_label.configure(text={'search': 'Recherche', 'tracking': 'Référence',
                                           'convoyage': 'ModBusId'}[name])
        self.keyword.configure(placeholder_text='ID, référence ou texte' if name == 'search' else
                               'Référence produit' if name == 'tracking' else 'Identifiant ModBus')
        if name == 'search':
            self.mode_frame.grid()
            self.btn_gear.grid()
            self.action_bar.grid()
        else:
            self.mode_frame.grid_remove()
            self.btn_gear.grid_remove()
            self.action_bar.grid_remove()
            self.advanced_frame.pack_forget()
            self.advanced_open = False
        self.keyword.unbind('<Return>')
        self.keyword.bind('<Return>', lambda event: self._submit_visible_search())

    def _submit_visible_search(self):
        if self._active_panel == 'tracking':
            return self.perform_tracking_search()
        if self._active_panel == 'convoyage':
            return self.convoyage_action()
        if getattr(self, 'advanced_open', False):
            return self.on_advanced_search()
        return self.perform_search()

    def _choose_search_mode(self):
        mode = self.search_mode.get()
        self.modbus.set(mode == 'id')
        self.reference.set(mode == 'ref')
        if mode == 'id':
            self.on_modbus_check()
        elif mode == 'ref':
            self.on_reference_check()
        else:
            self.btn_up.configure(state='normal')
            self.btn_down.configure(state='normal')
            self.btn_compare_idpk.configure(state='disabled')

    def _focus_search(self, event=None):
        self.keyword.focus_set()
        return 'break'

    def _install_status_widgets(self):
        self.status_frame = tk.Frame(self.workspace, bg='#f7f9fa')
        self.status_frame.pack(side='bottom', fill='x', padx=14, pady=(0, 8))
        tk.Label(self.status_frame, textvariable=self.status, anchor='w', font=('Segoe UI', 9),
                 bg='#f7f9fa', fg='#405060').pack(fill='x')
        self.index_status_label = tk.Label(self.workspace, textvariable=self.index_status,
                                           anchor='w', font=('Segoe UI', 9),
                                           bg='#f7f9fa', fg='#405060')
        self.index_status_label.pack(before=self.main_panel, fill='x', padx=14, pady=(0, 6))
        self._theme_window(self.status_frame)
        self._theme_window(self.index_status_label)

    def _update_context_header(self, path=None, occurrence=''):
        self._context_path = path
        label = getattr(self, 'file_label', None)
        if label is None or not label.winfo_exists():
            return
        label.configure(text='Fichier : ' + (trace_label(path) if path else 'Aucun'))
        self.source_label.configure(text=trace_label(path).split(' / ', 1)[0] if path else 'Aucune source')
        self.occurrence_label.configure(text=occurrence)
        self.btn_copy_path.configure(state='normal' if path else 'disabled')
        self._path_tooltip.text = path or ''

    def _copy_context_path(self):
        if self._context_path:
            self.clipboard_clear()
            self.clipboard_append(self._context_path)

    def _sync_result_visibility(self):
        classic = getattr(self, 'classic_frame', None)
        summary = getattr(self, 'summary_frame', None)
        if classic is None or not classic.winfo_exists() or summary is None or not summary.winfo_exists():
            return
        for value, button in self.result_switch._buttons_dict.items():
            button.configure(text_color=(LIGHT['accent_text'], DARK['accent_text'])
                             if value == self.result_mode.get() else (LIGHT['text'], DARK['text']))
        if self.result_mode.get() == 'Synthèse':
            classic.pack_forget()
            summary.pack(fill='both', expand=True)
        else:
            summary.pack_forget()
            classic.pack(fill='both', expand=True)

    def toggle_advanced_search(self):
        """
        Affiche ou masque la frame avancée (SN, CodeLot) juste SOUS la barre de recherche
        et AVANT le main_panel.
        """
        import customtkinter as ctk

        # Si on ouvre la zone avancée :
        if not hasattr(self, "advanced_open") or not self.advanced_open:
            # On pack juste AVANT self.main_panel (donc sous search_bar)
            self.advanced_frame.pack(before=self.main_panel, fill="x", padx=32, pady=(2, 6))
            # On vide d’éventuels anciens widgets
            for w in self.advanced_frame.winfo_children():
                w.destroy()
            # On crée les champs SN / CodeLot
            self.sn_label = ctk.CTkLabel(
                self.advanced_frame, text="SN :", font=("Segoe UI", 11),
                text_color=(LIGHT['text'], DARK['text'])
            )
            self.sn_label.pack(side="left", padx=(8, 2), pady=6)
            self.sn_entry = ctk.CTkEntry(
                self.advanced_frame, width=140, font=("Segoe UI", 13)
            )
            self.sn_entry.pack(side="left", padx=(0, 18), pady=6)

            self.codelot_label = ctk.CTkLabel(
                self.advanced_frame, text="CodeLot :", font=("Segoe UI", 11),
                text_color=(LIGHT['text'], DARK['text'])
            )
            self.codelot_label.pack(side="left", padx=(0, 2), pady=6)
            self.codelot_entry = ctk.CTkEntry(
                self.advanced_frame, width=140, font=("Segoe UI", 13)
            )
            self.codelot_entry.pack(side="left", padx=(0, 0), pady=6)

            # Bind <Return> sur les deux
            self.sn_entry.bind('<Return>', self.on_advanced_search)
            self.codelot_entry.bind('<Return>', self.on_advanced_search)

            self.advanced_open = True
            self._theme_window(self.advanced_frame)

        # Si on referme la zone avancée :
        else:
            self.advanced_frame.pack_forget()
            self.advanced_open = False


    def show_convoyage_tools(self):
        if getattr(self, 'tasks', None) is not None:
            self._invalidate_view()
        self._activate_panel('convoyage')
        # Efface le main_panel pour le remplir avec les outils convoyage
        for widget in self.main_panel.winfo_children():
            widget.destroy()

        # Titre principal
        title = tk.Label(self.main_panel, text="Outils Convoyage", font=("Segoe UI", 16, "bold"),
                         bg="#f7f9fa", fg="#00589b")
        title.pack(pady=(20, 10))

        # Cadre regroupant les boutons convoyage
        frame = tk.Frame(self.main_panel, bg="#f7f9fa")
        frame.pack(pady=(10, 12))

        # Bouton « Voir le parcours »
        btn1 = ctk.CTkButton(
            frame,
            text="🔍 Voir le parcours d’un ModBusId",
            corner_radius=8,
            fg_color="#2196f3",
            hover_color="#1565c0",
            text_color="white",
            font=("Segoe UI", 10, "bold"),
            command=self.convoyage_action
        )
        btn1.pack(fill="x", pady=(0, 8))

        # Explications sous chaque bouton
        tk.Label(frame, text="Affiche la dernière position d’un ModBusId dans TracMODBUS.",
                 font=("Segoe UI", 9), bg="#f7f9fa", fg="#555").pack(pady=(0, 14))

        # Bouton « Lister les erreurs »
        btn2 = ctk.CTkButton(
            frame,
            text="🚩 Lister les erreurs de convoyage",
            corner_radius=8,
            fg_color="#2196f3",
            hover_color="#1565c0",
            text_color="white",
            font=("Segoe UI", 10, "bold"),
            command=self.convoyage_errors_action
        )
        btn2.pack(fill="x", pady=(0, 8))

        # Explications sous chaque bouton
        tk.Label(frame, text="Liste tous les ModBusId jamais arrivés à destination dans les fichiers TracMODBUS.",
                 font=("Segoe UI", 9), bg="#f7f9fa", fg="#b31c1c").pack()

        # Bouton retour vers la recherche principale
        ctk.CTkButton(
            self.main_panel,
            text="← Retour à la recherche",
            corner_radius=8,
            fg_color="#2196f3",
            hover_color="#1565c0",
            text_color="white",
            font=("Segoe UI", 10, "bold"),
            command=self.show_search_panel
        ).pack(pady=8)

        self.keyword.unbind('<Return>')  # Retire l'ancien comportement
        self.keyword.bind('<Return>', self._convoyage_entry_enter)
        self._theme_window(self.main_panel)


    def show_search_panel(self):
        if getattr(self, 'tasks', None) is not None:
            self._invalidate_view()
        self._activate_panel('search')
        for widget in self.main_panel.winfo_children():
            widget.destroy()
        header = ctk.CTkFrame(self.main_panel, fg_color='transparent')
        header.pack(fill='x', pady=(0, 8))
        header.grid_columnconfigure(2, weight=1)
        self.source_label = ctk.CTkLabel(header, text='Aucune source', corner_radius=5,
                                        fg_color=(LIGHT['raised'], DARK['raised']),
                                        font=('Segoe UI', 10, 'bold'), padx=8)
        self.source_label.grid(row=0, column=0, sticky='w', padx=(0, 8))
        self.btn_prev_file = ttk.Button(header, text='‹', width=2, command=self.on_left)
        self.btn_prev_file.grid(row=0, column=1)
        self.file_label = tk.Label(header, text='Fichier : Aucun', font=('Segoe UI', 11, 'bold'),
                                   bg='#f7f9fa', fg='#405060', anchor='w', wraplength=460)
        self.file_label.grid(row=0, column=2, sticky='ew', padx=8)
        self._path_tooltip = ToolTip(self.file_label, '')
        self.btn_next_file = ttk.Button(header, text='›', width=2, command=self.on_right)
        self.btn_next_file.grid(row=0, column=3)
        self.btn_copy_path = ctk.CTkButton(header, text='Copier chemin', width=94, height=28,
                                          command=self._copy_context_path, state='disabled',
                                          fg_color=(LIGHT['raised'], DARK['raised']),
                                          hover_color=(LIGHT['border'], DARK['border']),
                                          text_color=(LIGHT['text'], DARK['text']), font=('Segoe UI', 10))
        self.btn_copy_path.grid(row=0, column=4, padx=(8, 0))
        self.occurrence_label = ctk.CTkLabel(header, text='', font=('Segoe UI', 10),
                                            text_color=(LIGHT['muted'], DARK['muted']))
        self.occurrence_label.grid(row=1, column=0, columnspan=3, sticky='w')
        self.result_switch = ctk.CTkSegmentedButton(header, values=['Classique', 'Synthèse'],
                                                    variable=self.result_mode,
                                                    command=self._change_result_mode,
                                                    font=('Segoe UI', 11), height=28,
                                                    fg_color=(LIGHT['raised'], DARK['raised']),
                                                    selected_color=(LIGHT['accent'], DARK['accent']),
                                                    selected_hover_color=(LIGHT['hover'], DARK['hover']),
                                                    unselected_color=(LIGHT['raised'], DARK['raised']),
                                                    unselected_hover_color=(LIGHT['border'], DARK['border']),
                                                    text_color=(LIGHT['text'], DARK['text']))
        self.result_switch.grid(row=1, column=3, columnspan=2, sticky='e', pady=(5, 0))
        body = ctk.CTkFrame(self.main_panel, fg_color='transparent', corner_radius=0)
        body.pack(fill='both', expand=True)
        self.classic_frame = tk.Frame(body, bg='#ffffff', bd=1, relief='groove',
                                      highlightthickness=1, highlightbackground='#cfd8dc')
        self.classic_frame._preserve_classic_theme = True
        self.txt = scrolledtext.ScrolledText(self.classic_frame, wrap='word', font=('Consolas', 11),
                                             bg='#fafbfc', borderwidth=0)
        self.txt.pack(fill='both', expand=True, padx=8, pady=8)
        self.summary_frame = ctk.CTkScrollableFrame(body, fg_color=(LIGHT['surface'], DARK['surface']),
                                                    corner_radius=6)
        self.txt.tag_bind('clickable_get', '<Double-1>', self.on_double_click_get)
        self.txt.tag_bind('clickable_pute', '<Double-1>', self.on_double_click_pute)
        self.txt.tag_bind('clickable_refprod', '<Button-1>', self.on_click_refprod)
        self._sync_result_visibility()
        self.btn_erreurs.configure(state='normal' if getattr(self, 'files_loaded', False) else 'disabled')
        self._theme_window(self.main_panel)
        if hasattr(self, '_render_summary'):
            self._render_summary(None)

    def focus_on_ticket_and_get(self, filename, modbus_id, refprod):
        if getattr(self, 'store', None) is not None:
            return self._focus_indexed_ticket(filename, modbus_id, refprod)
        import re, os, datetime, tkinter as tk
        # Charge le contenu du fichier
        content = BUFFER.get(filename, "")
        lines = content.splitlines()

        # Styles (une seule fois suffit, mais on remet ici pour être safe)
        self.txt.tag_configure("date", font=("Segoe UI", 10), foreground="#909090", justify="center")
        self.txt.tag_configure("occ", font=("Segoe UI", 9), foreground="#adb5bd", justify="center")
        self.txt.tag_configure("divider", foreground="#c7c7c7", font=("Segoe UI", 9))
        self.txt.tag_configure("brute", font=("Consolas", 9), background="#f5f7fa", foreground="#1e293b", lmargin1=7,
                               lmargin2=7)
        self.txt.tag_configure("refcard", font=("Segoe UI", 10, "bold"), foreground="#00876c", background="#e9fdf4",
                               lmargin1=6, lmargin2=6, spacing1=6, spacing3=6, justify="center")
        self.txt.tag_configure("none", font=("Segoe UI", 9, "bold", "italic"), foreground="#b71c1c", justify="center",
                               spacing1=8, spacing3=8)
        self.txt.tag_configure("clickable_refprod", foreground="#00876c", underline=1, font=("Segoe UI", 10, "bold"))
        self.txt.tag_configure("getcard", font=("Segoe UI", 10, "bold"), foreground="#00589b", background="#e3f2fd",
                               lmargin1=10, lmargin2=10)

        # 1) Trouver la ligne ticket (RefProduit=refprod et ModBusId=modbus_id)
        ticket_idx = None
        for idx, line in enumerate(lines):
            if field_matches(line, "RefProduit", refprod) and field_matches(line, "ModBusId", modbus_id):
                ticket_idx = idx
                break
        if ticket_idx is None:
            self.txt.config(state='normal')
            self.txt.delete("1.0", tk.END)
            self.txt.insert(tk.END, "Ticket introuvable dans ce fichier.\n")
            self.txt.config(state='disabled')
            return

        # Affichage principal (comme dans display_line)
        self.txt.config(state='normal')
        self.txt.delete(1.0, tk.END)

        # Date fichier
        file_date = trace_line_datetime(filename, lines[ticket_idx])
        self.txt.insert(tk.END, f"{file_date.strftime('%d/%m/%Y')}\n", "date")
        self.txt.insert(tk.END, f"Occurrence 1/1\n", "occ")
        self.txt.insert(tk.END, "─" * 56 + "\n", "divider")

        # Infos du ticket
        line = lines[ticket_idx]
        self.lines = [line]
        self.line_index = 0

        # Extraction des champs
        extract_fields_brut = format_ticket_fields

        brut_text, line_fields = extract_fields_brut(line)
        self.txt.insert(tk.END, brut_text + "\n", "brute")
        self.txt.insert(tk.END, "─" * 56 + "\n", "divider")

        # Bloc RefProduit cliquable
        if line_fields["refprod"]:
            start_r = self.txt.index(tk.END)
            self.txt.insert(tk.END, f"RefProduit trouvé : {line_fields['refprod']}\n", ("refcard", "clickable_refprod"))
            end_r = self.txt.index(tk.END)
            self.txt.tag_add("clickable_refprod", start_r, end_r)
            self.txt.insert(tk.END, "─" * 56 + "\n", "divider")

        # 3) Chercher le GRIPPER GET après le ticket pour cette ref
        get_line = None
        for l in lines[ticket_idx + 1:]:
            if is_gripper(l, "GET", refprod, loading_only=True):
                get_line = l
                break

        if get_line:
            s1 = self.txt.index(tk.END)
            self.txt.insert(tk.END, f"GRIPPER GET :\n{get_line}\n", ("getcard", "clickable_get"))
            e1 = self.txt.index(tk.END)
            self.txt.tag_add("clickable_get", s1, e1)
            self.txt.insert(tk.END, "─" * 56 + "\n", "divider")
        else:
            # Message d'erreur, rouge, centré, espacé
            self.txt.insert(tk.END, "\n", "none")
            self.txt.insert(tk.END, f"Aucun Gripper GET trouvé pour {refprod}\n", "none")

        self.file_label.config(text=f"Fichier : {trace_label(filename)}")
        self.txt.config(state='disabled')

        # Bind pour clics (si besoin)
        self.txt.tag_bind("clickable_get", "<Double-1>", self.on_double_click_get)
        self.txt.tag_bind("clickable_refprod", "<Button-1>", self.on_click_refprod)

    def _on_error_line_click(self, filename, modbus_id, refprod, number=None):
        if getattr(self, 'store', None) is not None:
            return self._focus_indexed_ticket(filename, modbus_id, refprod, show_adjacent=True, number=number)
        """
        Quand on clique sur une ligne d’erreur Robot :
        - on charge le bon fichier,
        - on passe en mode Référence et on recherche la RefProduit,
        - on ouvre la popup IDs adjacents pour le ModBusId.
        """
        import os
        from tkinter import messagebox

        # 1) Trouver le chemin complet depuis self.file_list
        full_path = next((p for p in self.file_list if path_key(p) == path_key(filename)), None)
        if not full_path:
            messagebox.showerror("Erreur", f"Fichier {filename} introuvable.")
            return

        # 2) Mettre à jour current_file et index
        global current_file
        current_file = full_path
        self.file_index = self.file_list.index(full_path)

        # 3) Passer en mode Référence
        self.reference.set(True)
        self.modbus.set(False)
        self.btn_up.configure(state="disabled")
        self.btn_down.configure(state="disabled")

        # 4) Remplir l’Entry avec la RefProduit et activer Comparer IDPK
        self.keyword.delete(0, tk.END)
        self.keyword.insert(0, refprod)
        self.btn_compare_idpk.configure(state="normal")

        # 5) Affiche ticket et GRIPPER GET correspondant
        self.focus_on_ticket_and_get(full_path, modbus_id, refprod)

        # 6) Activer et ouvrir la popup “IDs adjacents” pour ce ModBusId
        self.btn_adjacent.configure(state="normal")
        self.show_adjacent_modbusids_window(modbus_id, full_path)

    def on_modbus_check(self):
        if self.modbus.get():
            if hasattr(self, 'search_mode'):
                self.search_mode.set('id')
            self.reference.set(False)
            self.btn_up.configure(state="normal")
            self.btn_down.configure(state="normal")
            self.btn_compare_idpk.configure(state="disabled")

    def on_reference_check(self):
        if self.reference.get():
            if hasattr(self, 'search_mode'):
                self.search_mode.set('ref')
            self.modbus.set(False)
            self.btn_up.configure(state="disabled")
            self.btn_down.configure(state="disabled")
            self.btn_compare_idpk.configure(state="normal")

    def find_gripper_line(self, content, refprod):
        # cherche la ligne GET dans tout le fichier
        for i, line in enumerate(content.splitlines()):
            if is_gripper(line, "GET", refprod, loading_only=True):
                return i, line  # retourne index et ligne
        return None, None

    def find_previous_get_and_idpk(self, content, refprod, current_get_index):
        """
        Parcours à l'envers jusqu'au GRIPPER GET $0101 précédent, puis cherche le Résultat IDPK qui suit.
        """
        lines = content.splitlines()
        for i in range(current_get_index - 1, -1, -1):
            line = lines[i]
            if (is_gripper(line, "GET", refprod, loading_only=True)):
                # Trouvé le GET précédent
                # Cherche l'IDPK juste après (dans les 5 lignes suivantes, typiquement)
                result = find_idpk_after(lines, refprod, i)
                if result[0] is not None:
                    return result
                # Si pas trouvé, continue à chercher un GET encore plus ancien
        return None, None

    def find_previous_get_and_idpk_across_files(self, refprod, current_file, current_get_index):
        """
        Cherche l’occurrence précédente de GRIPPER GET pour la même refprod (avec $0101),
        puis cherche l’IDPK juste après, dans les fichiers SUIVANTS dans la liste (plus anciens).
        """
        # 1. Dans le fichier courant, chercher en arrière à partir de current_get_index - 1
        content = BUFFER.get(current_file, "")
        lines = content.splitlines()
        for i in range(current_get_index - 1, -1, -1):
            line = lines[i]
            if is_gripper(line, "GET", refprod, loading_only=True):
                j, next_line = find_idpk_after(lines, refprod, i)
                if j is not None:
                    return j, next_line, current_file

        # 2. Chercher dans les fichiers SUIVANTS (plus anciens)
        try:
            current_idx = self.file_list.index(current_file)
        except ValueError:
            return None, None, None  # Sécurité

        for prev_file in self.file_list[current_idx + 1:]:
            if not same_trace_directory(prev_file, current_file):
                continue
            content_prev = BUFFER.get(prev_file, "")
            lines_prev = content_prev.splitlines()
            for i in range(len(lines_prev) - 1, -1, -1):
                line = lines_prev[i]
                if is_gripper(line, "GET", refprod, loading_only=True):
                    j, next_line = find_idpk_after(lines_prev, refprod, i)
                    if j is not None:
                        return j, next_line, prev_file
        return None, None, None

    def find_result_idpk_line(self, content, refprod, get_index):
        """
        Recherche la première ligne Rsultat de la mesure IDPK pour refprod APRES la ligne GRIPPER GET.
        """
        return find_idpk_after(content.splitlines(), refprod, get_index)

    def find_gripper_pute_line(self, content, refprod, start_index):
        # cherche la ligne PUTE à partir de start_index + 1
        lines = content.splitlines()
        for i in range(start_index + 1, len(lines)):
            line = lines[i]
            if is_gripper(line, "PUTE", refprod):
                return i, line
        return None, None

    def find_gripper_line_from(self, content, refprod, start_idx):
        """
        Cherche la ligne GRIPPER GET (avec $0101) pour refprod
        **après** start_idx dans le contenu.
        """
        lines = content.splitlines()
        for i in range(start_idx + 1, len(lines)):
            l = lines[i]
            if is_gripper(l, "GET", refprod, loading_only=True):
                return i, l
        return None, None

    def find_gripper_pute_from(self, content, refprod, get_idx):
        """
        Cherche la ligne GRIPPER PUTE pour refprod
        **après** get_idx dans le contenu.
        """
        lines = content.splitlines()
        for i in range(get_idx + 1, len(lines)):
            l = lines[i]
            if is_gripper(l, "PUTE", refprod):
                return i, l
        return None, None


    def get_separator(self):
        # Largeur actuelle de la zone de texte (en caractères)
        # On prend la largeur du widget (en pixels) divisé par la taille moyenne d’un caractère (8-9 px pour Courier 12)
        width_px = self.txt.winfo_width()
        if width_px < 400:  # Valeur par défaut au lancement
            width_px = 700
        char_count = int(width_px / 8.5)
        return "-" * max(25, char_count - 2)









    # Dans la classe SearchApp

    def on_double_click_get(self, event):
        """
        Quand on double‐clic sur une ligne GRIPPER GET, on ouvre la vue complète
        avec cette ligne surlignée.
        """
        self._open_exact_gripper_line(event, "GRIPPER GET")

    def on_double_click_pute(self, event):
        """
        Quand on double‐clic sur une ligne GRIPPER PUTE, on ouvre la vue complète
        avec cette ligne surlignée.
        """
        self._open_exact_gripper_line(event, "GRIPPER PUTE")



    def open_file_view(self, highlight_line=None):
        if getattr(self, 'store', None) is not None:
            return self._indexed_open_file_view(highlight_line)
        """
        Ouvre la fenêtre de trace complète, identique au bouton “Ouvrir trace”,
        avec barre de recherche locale, flèches ↑/↓, navigation GET/PUTE, et
        surlignage facultatif de la ligne highlight_line (0-based).
        """
        import os, re
        from tkinter import scrolledtext, messagebox

        # 1) Choix du fichier à afficher
        if highlight_line is None and not self.keyword.get().strip():
            trac_files = [
                path for path in discover_trace_files("tracomega")
                if not os.path.basename(path).lower().startswith("tracomegaperm")
            ]
            if not trac_files:
                messagebox.showinfo("Info", "Aucun fichier TracOmega trouvé.")
                return
            trac_files.sort(key=trace_file_sort_key, reverse=True)
            path = trac_files[0]
        else:
            path = current_file

        if not path or not os.path.isfile(path):
            messagebox.showinfo("Info", "Aucun fichier de trace disponible.")
            return

        # 2) Lecture du contenu
        lines = read_trace(path).splitlines()

        # 3) Création de la fenêtre
        top = tk.Toplevel(self)
        apply_icon(top)
        top.title(f"Vue de {trace_label(path)}")
        top.geometry("900x650")
        top.configure(bg="#f7f9fa")

        # 4) Barre de recherche locale + navigation ↑/↓
        search_frame = tk.Frame(top, bg="#f7f9fa")
        search_frame.pack(fill='x', padx=18, pady=(3, 8))
        tk.Label(search_frame, text="Recherche dans ce fichier :", font=("Segoe UI", 10),
                 bg="#f7f9fa", fg="#37474f").pack(side='left', padx=(2, 6))
        local_search = ttk.Entry(search_frame, font=("Segoe UI", 11), width=30)
        local_search.pack(side='left', padx=(2, 8), ipadx=3)

        # flèches haut/bas
        nav_btns = tk.Frame(search_frame, bg="#f7f9fa")
        nav_btns.pack(side='left', padx=(2, 3))
        btn_search_up = ttk.Button(nav_btns, text="↑", width=2)
        btn_search_up.pack(side='top', pady=(0, 1))
        btn_search_down = ttk.Button(nav_btns, text="↓", width=2)
        btn_search_down.pack(side='top', pady=(1, 0))

        # 5) Flèches GET
        get_grp = tk.Frame(search_frame, bg="#f7f9fa")
        get_grp.pack(side='left', padx=(18, 0))
        tk.Label(get_grp, text="GET", font=("Segoe UI", 9, "bold"),
                 bg="#b3e5fc", fg="#1565c0", padx=6, pady=1, relief="groove").pack(side='left', padx=(0, 3))
        btn_prev_get = ttk.Button(get_grp, text="⟨", width=2)
        btn_prev_get.pack(side='left', padx=1)
        btn_next_get = ttk.Button(get_grp, text="⟩", width=2)
        btn_next_get.pack(side='left', padx=1)

        # 6) Flèches PUTE
        pute_grp = tk.Frame(search_frame, bg="#f7f9fa")
        pute_grp.pack(side='left', padx=(8, 0))
        tk.Label(pute_grp, text="PUTE", font=("Segoe UI", 9, "bold"),
                 bg="#ffcdd2", fg="#c62828", padx=6, pady=1, relief="groove").pack(side='left', padx=(0, 3))
        btn_prev_pute = ttk.Button(pute_grp, text="⟨", width=2)
        btn_prev_pute.pack(side='left', padx=1)
        btn_next_pute = ttk.Button(pute_grp, text="⟩", width=2)
        btn_next_pute.pack(side='left', padx=1)

        # 7) Zone texte principale
        text_widget = scrolledtext.ScrolledText(top, wrap='word',
                                                font=("Consolas", 11),
                                                bg="#fafbfc", borderwidth=0)
        text_widget.pack(fill='both', expand=True, padx=12, pady=10)

        # 8) Configuration des tags
        text_widget.tag_configure("mainline", background="#e3f2fd", font=("Consolas", 11, "bold"))
        text_widget.tag_configure("getline", background="#b3e5fc")
        text_widget.tag_configure("puteline", background="#ffcdd2")
        text_widget.tag_configure("search", background="#fff59d")
        text_widget.tag_configure("highlight", background="yellow")

        # 9) Insertion des lignes et surlignage si demandé
        for i, l in enumerate(lines):
            text_widget.insert(tk.END, l + "\n")
        if highlight_line is not None and 0 <= highlight_line < len(lines):
            start = f"{highlight_line + 1}.0"
            text_widget.tag_add("highlight", start, f"{start} lineend")
            text_widget.see(start)

        # 10) Fonctions de recherche et navigation
        search_state = {'positions': [], 'current': -1}

        def update_search_positions():
            term = local_search.get().strip()
            search_state['positions'].clear()
            text_widget.tag_remove("search", "1.0", tk.END)
            if not term:
                return
            idx = "1.0"
            while True:
                idx = text_widget.search(term, idx, nocase=1, stopindex=tk.END)
                if not idx:
                    break
                search_state['positions'].append(idx)
                text_widget.tag_add("search", idx, f"{idx}+{len(term)}c")
                idx = f"{idx}+{len(term)}c"

        def goto_search(offset):
            if not search_state['positions']:
                return
            search_state['current'] = (search_state['current'] + offset) % len(search_state['positions'])
            pos = search_state['positions'][search_state['current']]
            text_widget.see(pos)
            text_widget.tag_remove("mainline", "1.0", tk.END)
            text_widget.tag_add("mainline", pos, f"{pos} lineend")

        def nav_get(offset):
            get_lines = [i for i, l in enumerate(lines) if "GRIPPER GET" in l]
            if not get_lines:
                return
            current = int(text_widget.index(tk.INSERT).split('.')[0]) - 1
            candidates = [i for i in get_lines if (i > current if offset > 0 else i < current)]
            idx = (candidates[0] if offset > 0 and candidates else
                   candidates[-1] if offset < 0 and candidates else
                   get_lines[0] if offset > 0 else get_lines[-1])
            pos = f"{idx + 1}.0"
            text_widget.see(pos)
            text_widget.tag_remove("getline", "1.0", tk.END)
            text_widget.tag_add("getline", pos, f"{pos} lineend")

        def nav_pute(offset):
            pute_lines = [i for i, l in enumerate(lines) if "GRIPPER PUTE" in l]
            if not pute_lines:
                return
            current = int(text_widget.index(tk.INSERT).split('.')[0]) - 1
            candidates = [i for i in pute_lines if (i > current if offset > 0 else i < current)]
            idx = (candidates[0] if offset > 0 and candidates else
                   candidates[-1] if offset < 0 and candidates else
                   pute_lines[0] if offset > 0 else pute_lines[-1])
            pos = f"{idx + 1}.0"
            text_widget.see(pos)
            text_widget.tag_remove("puteline", "1.0", tk.END)
            text_widget.tag_add("puteline", pos, f"{pos} lineend")

        # 11) Bindings
        btn_search_up.config(command=lambda: goto_search(-1))
        btn_search_down.config(command=lambda: goto_search(1))
        local_search.bind('<KeyRelease>', lambda e: update_search_positions())
        local_search.bind('<Return>', lambda e: goto_search(1))

        btn_prev_get.config(command=lambda: nav_get(-1))
        btn_next_get.config(command=lambda: nav_get(1))
        btn_prev_pute.config(command=lambda: nav_pute(-1))
        btn_next_pute.config(command=lambda: nav_pute(1))

        # Initialisation de la recherche
        update_search_positions()

        top.lift()


        # --- Reste de la logique de recherche locale + nav GET/PUTE identique à open_file_view —
        # tu peux directement copier-coller les handlers de open_file_view

    def get_convoyage_popup_geometry(self):
        # Ici, tu définis largeur/hauteur fixes :
        largeur = 400
        hauteur = 405
        self.update_idletasks()
        main_x = self.winfo_rootx()
        main_y = self.winfo_rooty()
        popup_x = main_x + self.winfo_width() + 40
        popup_y = main_y + 30
        return f"{largeur}x{hauteur}+{popup_x}+{popup_y}"

    def on_click_refprod(self, event):
        return IndexedAppMixin.on_click_refprod(self, event)


    def extract_idpk_dimensions(self, idpk_line):
        """
        Extrait les dimensions l, H, L d'une ligne Résultat IDPK.
        Exemple ligne : 'B 14|13:38:24.0 Rsultat de la mesure IDPK pour PLH000155.00 : l=64 H=41 L=91'
        """
        parsed = parse_idpk(idpk_line)
        return {key: str(value) for key, value in parsed[1].items()} if parsed else {}


    def _convoyage_entry_enter(self, event=None):
        modbusid = self.keyword.get().strip()
        if modbusid.isdigit():  # Simple vérif, adapte selon ton format d’id si besoin
            self.convoyage_action(modbusid)


    def _convoyage_details_click(self, ref, event):
        self.show_convoyage_details(ref)

    def get_convoyage_steps(self, ref, only_filepath=None):
        if getattr(self, 'store', None) is not None:
            return self.store.convoyage_steps(ref, only_filepath)
        import os
        checkpoint_indices = [2, 31, 32, 33, 34, 35, 36, 37, 57, 58, 6, 114, 7, 115]
        checkpoint_noms = [
            'Boite entrée', 'En attente en bas de l’asc POSTE1', 'En attente en bas de l’asc POSTE2',
            'En attente en bas de l’asc POSTE3', 'En attente en bas de l’asc POSTE4',
            'Boite en train d’être convoyée vers l’ascenseur 1 ou 2 ',
            'Boite arrivée devant l’ascenseur 1',
            'Boite arrivée devant l’ascenseur 2',
            "Boite poussée depuis l’ascenseur 1 vers le T5/R1",
            "Boite poussée depuis l’ascenseur 2 vers le T5/R2",
            "Boite en attente sur T5 / R1", "Unnamed: 114", "Boite en attente sur T5 / R2", "Unnamed: 115"
        ]
        checkpoints = dict(zip(checkpoint_indices, checkpoint_noms))
        recap = []
        tracmodbus_files = discover_trace_files("tracmodbus")
        tracmodbus_files.sort(key=trace_file_sort_key, reverse=True)
        for file in tracmodbus_files:
            if only_filepath and path_key(file) != path_key(only_filepath):
                continue
            try:
                lignes = read_trace(file).splitlines()
            except Exception:
                continue
            found = {col: False for col in checkpoint_indices}
            for l in lignes:
                valeurs = parse_modbus_line(l)
                horo = valeurs[0] if len(valeurs) > 0 else ''
                for col in checkpoint_indices:
                    if col < len(valeurs):
                        v = valeurs[col]
                        if not found[col] and v == ref:
                            recap.append({
                                "horodatage": horo,
                                "checkpoint": checkpoints[col],
                                "etat": "Présent",
                                "file": trace_label(file)
                            })
                            found[col] = True
                        elif not found[col] and v == f"-{ref}":
                            recap.append({
                                "horodatage": horo,
                                "checkpoint": checkpoints[col],
                                "etat": "Jamais arrivé",
                                "file": trace_label(file)
                            })
                            found[col] = True
        return recap

    def find_adjacent_modbusids(self, full_lines, target_idx, target_id, count=5):
        ids_before, seen_before = [], set()
        for i in range(target_idx - 1, -1, -1):
            m2 = re.search(r"ModBusId=(\d+)(\D|$)", full_lines[i])
            if m2:
                mb = m2.group(1)
                if mb != target_id and mb not in seen_before:
                    ids_before.append(mb); seen_before.add(mb)
                    if len(ids_before) >= count: break

        ids_after, seen_after = [], set()
        for i in range(target_idx + 1, len(full_lines)):
            m2 = re.search(r"ModBusId=(\d+)(\D|$)", full_lines[i])
            if m2:
                mb = m2.group(1)
                if mb != target_id and mb not in seen_after:
                    ids_after.append(mb); seen_after.add(mb)
                    if len(ids_after) >= count: break

        return ids_before, ids_after

    def _check_idpk_diff(self, refprod):
        """
        Retourne True si la différence IDPK (l, H ou L) entre l'occurence courante
        et la précédente dépasse 2 unités, False sinon.
        """
        # 1) Récupère la dimension courante dans current_file
        content_cur = BUFFER.get(current_file, "")
        get_idx, _ = self.find_gripper_line(content_cur, refprod)
        if get_idx is None:
            return False
        _, idpk_line = self.find_result_idpk_line(content_cur, refprod, get_idx)
        if not idpk_line:
            return False
        dims_cur = self.extract_idpk_dimensions(idpk_line)

        # 2) Recherche la dimension précédente dans les fichiers plus anciens
        prev_dims = None
        try:
            start = self.file_list.index(current_file) + 1
        except ValueError:
            return False
        for path in self.file_list[start:]:
            if not same_trace_directory(path, current_file):
                continue
            cont = BUFFER.get(path, "")
            gi, _ = self.find_gripper_line(cont, refprod)
            if gi is not None:
                _, prev_line = self.find_result_idpk_line(cont, refprod, gi)
                if prev_line:
                    prev_dims = self.extract_idpk_dimensions(prev_line)
                    break
        if not prev_dims:
            return False

        # 3) Compare l, H et L avec un seuil de 2
        for key in ("l", "H", "L"):
            try:
                if abs(int(dims_cur.get(key, 0)) - int(prev_dims.get(key, 0))) > 2:
                    return True
            except ValueError:
                continue

        return False



    def get_adjacent_popup_geometry(self):
        """
        Calcule et renvoie la géométrie de la popup (taille + position à droite).
        """
        self.update_idletasks()
        x = self.winfo_rootx() + self.winfo_width() + 10
        y = self.winfo_rooty()
        width, height = 240, 160
        return f"{width}x{height}+{x}+{y}"

    def _change_id(self, delta):
        """
        Incrémente ou décrémente l’ID de delta (±1) et rafraîchit la fenêtre.
        """
        try:
            cur = int(self.adj_id_var.get())
        except ValueError:
            return
        new = cur + delta
        self.adj_id_var.set(str(new))
        # relance le calcul avec le même fichier courant
        if getattr(self, 'store', None):
            path = self.selection.path if self.selection else None
        else:
            path = current_file
        if path:
            self.show_adjacent_modbusids_window(str(new), path)

    def _build_idpk_index(self):
        """
        Parcourt BUFFER une fois pour récupérer pour chaque RefProduit
        la liste (triée) des (mtime, dims) pour les lignes IDPK.
        """
        import re, os
        self.idpk_index = {}  # ref -> [ (mtime, {'l':..,'H':..,'L':..}), ... ]

        for path, text in BUFFER.items():
            for line in text.splitlines():
                parsed = parse_idpk(line)
                if parsed and all(key in parsed[1] for key in ('l', 'H', 'L')):
                    ref, dims = parsed
                    when = line_datetime(line, BUFFER_START[path]).timestamp()
                    self.idpk_index.setdefault(ref, []).append((when, dims))

        # trier chaque liste par date décroissante
        for lst in self.idpk_index.values():
            lst.sort(key=lambda x: x[0], reverse=True)


    def _refresh_traces(self):
        with TRACE_LOCK:
            changed = reload_modified_files()
            if changed or not hasattr(self, "ref_index"):
                self._rebuild_trace_indexes()

    def _rebuild_trace_indexes(self):
        from collections import defaultdict
        global current_file

        self.next_file_map = build_next_file_map(BUFFER)
        # 4) Construire les index RefProduit, SN et CodeLot
        self.ref_index = defaultdict(list)  # ref -> list of (path, line)
        self.sn_index = defaultdict(list)  # sn  -> list of (path, line)
        self.codelot_index = defaultdict(list)  # cl  -> list of (path, line)
        indexes = {'RefProduit': self.ref_index, 'SN': self.sn_index, 'CodeLot': self.codelot_index}
        for path, text in BUFFER.items():
            for line in text.splitlines():
                for field, value in parse_fields(line).items():
                    if field in indexes:
                        indexes[field][value].append((path, line))

        # 5) Reconstruire l’index IDPK
        self._build_idpk_index()

        # 6) Reconstruire la liste de fichiers triée
        self.file_list = sorted(
            BUFFER.keys(),
            key=trace_file_sort_key,
            reverse=True
        )

        self.advanced_candidates = []
        self.advanced_index = 0
        if current_file not in BUFFER:
            current_file = None
            self.lines = []
            self.line_index = 0
            self.file_index = 0
        else:
            self.file_index = self.file_list.index(current_file)

    def show_tracking_robot_panel(self):
        if getattr(self, 'tasks', None) is not None:
            self._invalidate_view()
        self._activate_panel('tracking')
        for widget in self.main_panel.winfo_children():
            widget.destroy()
        ctk.CTkLabel(self.main_panel, text='Suivi des dépôts par référence',
                     font=('Segoe UI', 16, 'bold'), anchor='w').pack(fill='x', pady=(6, 12))
        frame = tk.Frame(self.main_panel, bg='#ffffff', bd=1, relief='flat')
        frame.pack(fill='both', expand=True)
        self.tracking_txt = scrolledtext.ScrolledText(frame, wrap='word', font=('Consolas', 11),
                                                       bg='#fafbfc', borderwidth=0)
        self.tracking_txt.pack(fill='both', expand=True, padx=10, pady=10)
        self._theme_window(self.main_panel)

    def perform_tracking_search(self):
        if getattr(self, 'store', None) is not None:
            return self._indexed_tracking()
        import re
        ref = self.tracking_keyword.get().strip()
        if not ref:
            self.tracking_txt.config(state="normal")
            self.tracking_txt.delete(1.0, tk.END)
            self.tracking_txt.insert(tk.END, "Veuillez entrer une référence.")
            self.tracking_txt.config(state="disabled")
            return

        self._refresh_traces()
        results = []
        for file_path, content in BUFFER.items():
            for dt, line in iter_timed_lines(content, BUFFER_START[file_path]):
                if is_gripper(line, "PUTE", ref):
                    _, heure = timestamp_parts(line)
                    mface = re.search(r"(face:\d+\s*\d*)", line)
                    face = mface.group(1) if mface else ""
                    results.append((dt, f"{dt:%d/%m/%y} | {heure} | {trace_label(file_path)} | PUTE | {face}"))

        # Tri croissant par date+heure
        results.sort(reverse=True)

        self.tracking_txt.config(state="normal")
        self.tracking_txt.delete(1.0, tk.END)
        if results:
            for _dt, text in results:
                self.tracking_txt.insert(tk.END, text + "\n")
        else:
            self.tracking_txt.insert(tk.END, "Aucune opération PUTE trouvée pour cette référence.")
        self.tracking_txt.config(state="disabled")

    def on_advanced_search(self, event=None):
        """
        Recherche par RefProduit (obligatoire), et éventuellement SN et/ou CodeLot.
        Utilise les index pré-construits pour des recherches instantanées.
        """
        from tkinter import messagebox

        # 1) Récupère les critères saisis
        ref = self.keyword.get().strip()
        sn = getattr(self, "sn_entry", None)
        sn_value = sn.get().strip() if sn else ""
        cl = getattr(self, "codelot_entry", None)
        cl_value = cl.get().strip() if cl else ""

        # 2) Validation RefProduit
        if not ref:
            messagebox.showwarning("Recherche", "Veuillez saisir une référence au minimum !")
            return

        # 3) Lancer la recherche avancée en s'appuyant sur nos index
        self.recherche_avancee_criteria(
            ref=ref,
            sn=sn_value or None,
            codelot=cl_value or None
        )



    def _render_robot_errors(self, errors, files_checked):
        import tkinter as tk

        WIDTH = 500  # Largeur de base de la popup

        top = tk.Toplevel(self)
        apply_icon(top)
        top.title("Vérification erreurs Robot")
        top.resizable(True, True)
        top.geometry(f"{WIDTH}x400")

        frame = tk.Frame(top, bg="#f7fafd")
        frame.pack(fill="both", expand=True, padx=18, pady=18)

        # --- Titre qui s'ajuste dynamiquement ---
        titre = tk.Label(
            frame,
            text="🤖 Résultat de la vérification des erreurs Robot",
            font=("Arial", 16, "bold"),
            fg="#00589b",
            bg="#f7fafd",
            wraplength=WIDTH - 30,
            justify="left"
        )
        titre.pack(pady=(0, 12), fill="x", expand=False)

        def update_title_wrap(event):
            titre.config(wraplength=event.width - 30)

        frame.bind("<Configure>", update_title_wrap)

        if errors:
            # --- Encadré scrollable pour les erreurs ---
            container = tk.Frame(frame, bg="#ffeaea", bd=2, relief="groove")
            container.pack(fill="both", expand=True, pady=(0, 0))

            scroll_canvas = tk.Canvas(container, bg="#ffeaea", highlightthickness=0, bd=0)
            scroll_canvas.pack(side="left", fill="both", expand=True)
            scrollbar = tk.Scrollbar(container, orient="vertical", command=scroll_canvas.yview)
            scrollbar.pack(side="right", fill="y")
            scroll_canvas.configure(yscrollcommand=scrollbar.set)

            inner_frame = tk.Frame(scroll_canvas, bg="#ffeaea")
            window_id = scroll_canvas.create_window((0, 0), window=inner_frame, anchor="nw")

            label_refs = []
            # Ajout des labels dans inner_frame
            def add_error(err):
                dateheure = ""
                if err.get("date") or err.get("heure"):
                    dateheure = f" | {err.get('date', '')} {err.get('heure', '')}".strip()
                err_label = tk.Label(
                    inner_frame,
                    text=f"❌ Fichier: {err['file']} | ModBusId={err['modbus_id']} | RefProduit={err['ref']}{dateheure} non prise par GRIPPER GET",
                    font=("Consolas", 11, "bold"),
                    fg="#b71c1c",
                    bg="#ffeaea",
                    anchor="w",
                    justify="left",
                    wraplength=WIDTH - 40,
                    cursor="hand2"
                )
                err_label.pack(fill="x", padx=8, pady=2)
                # binding pour naviguer vers le fichier / ref / popup
                err_label.bind(
                    "<Button-1>",
                    lambda e,
                           fname=err['filepath'],
                           mb=err['modbus_id'],
                           rp=err['ref'],
                           number=err.get('number'):
                    self._on_error_line_click(fname, mb, rp, number)
                )
                label_refs.append(err_label)

            self._render_items(errors, add_error, top)

            # affichage direct
            scroll_canvas.update_idletasks()
            scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))

            # Ajuste wrap + zone scroll au resize
            def on_frame_configure(event):
                scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))
                scroll_canvas.itemconfig(window_id, width=scroll_canvas.winfo_width())
                for lbl in label_refs:
                    lbl.config(wraplength=max(scroll_canvas.winfo_width() - 32, 60))

            inner_frame.bind("<Configure>", on_frame_configure)

            # Molette souris
            def _on_mousewheel(event):
                scroll_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

            def _on_mousewheel_linux(event):
                if event.num == 4:
                    scroll_canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    scroll_canvas.yview_scroll(1, "units")

            def bind_mousewheel(event):
                scroll_canvas.bind_all("<MouseWheel>", _on_mousewheel)
                scroll_canvas.bind_all("<Button-4>", _on_mousewheel_linux)
                scroll_canvas.bind_all("<Button-5>", _on_mousewheel_linux)

            def unbind_mousewheel(event):
                scroll_canvas.unbind_all("<MouseWheel>")
                scroll_canvas.unbind_all("<Button-4>")
                scroll_canvas.unbind_all("<Button-5>")

            scroll_canvas.bind("<Enter>", bind_mousewheel)
            scroll_canvas.bind("<Leave>", unbind_mousewheel)

        else:
            ok = tk.Label(
                frame,
                text="✅ Aucun problème détecté sur tous les fichiers analysés.",
                font=("Arial", 13, "bold"),
                fg="#228B22",
                bg="#f7fafd"
            )
            ok.pack(pady=12)

        details = tk.Label(
            frame,
            text=f"{len(files_checked)} fichier(s) analysé(s).",
            font=("Arial", 10, "italic"),
            fg="#606060",
            bg="#f7fafd",
            wraplength=WIDTH - 30,
            justify="left"
        )
        details.pack(pady=(4, 0), fill="x", expand=False)

        note = tk.Label(
            frame,
            text="Les erreurs signalent les ModBusId/RefProduit en attente sur T5 et jamais pris GRIPPER GET ($0101).",
            font=("Arial", 10, "italic"),
            fg="#8a6d3b",
            bg="#f7fafd",
            wraplength=WIDTH - 30,
            justify="left"
        )
        note.pack(pady=(12, 0), fill="x", expand=False)

        top.lift()


    def _render_indexed_convoyage(self, ref, geometry_override, all_results, indexed_steps):
        if not hasattr(self, 'convoyage_window'):
            self.convoyage_window = None
        def make_fully_clickable(frame, callback):
            frame.bind('<Button-1>', callback)
            for child in frame.winfo_children():
                child.bind('<Button-1>', callback)
        if not all_results:
            tk.messagebox.showinfo("Résultat",
                                   f"Aucune position trouvée pour la référence {ref} dans le convoyage (tous fichiers).")
            return

        # ==== Création / update de la fenêtre convoyage stylée ====
        if self.convoyage_window and tk.Toplevel.winfo_exists(self.convoyage_window):
            top = self.convoyage_window
            for widget in top.winfo_children():
                widget.destroy()
            # ← ajoute ça ici
            top.update_idletasks()
            top.geometry("")
        else:
            self.convoyage_window = tk.Toplevel(self)
            apply_icon(self.convoyage_window)
            top = self.convoyage_window
            top.title(f"Convoyage - Résultats pour la référence {ref}")
            self.convoyage_window.resizable(True, True)

        frame = tk.Frame(top, bg="#fff8f0")
        frame.pack(fill="both", expand=True, padx=18, pady=18)
        frame.update_idletasks()
        top.update_idletasks()
        top.minsize(0, 0)
        top.geometry("")

        label_refs = []

        def add_label(*args, **kwargs):
            lbl = tk.Label(*args, **kwargs)
            label_refs.append(lbl)
            return lbl

        def show_convoyage_steps_in_bloc(container, ref, filepath, geometry_override=None):
            filename = trace_label(filepath)
            for w in container.winfo_children():
                w.destroy()

            titre = tk.Label(
                container,
                text=f"📦 Parcours détaillé du ModBusId {ref} dans {filename}",
                font=("Segoe UI", 13, "bold"),
                fg="#1a5276",
                bg="#f5fafd"
            )
            titre.pack(anchor="w", pady=(0, 10), padx=10)

            headers = ["Horodatage", "Checkpoint", "État", "Fichier"]
            col_widths = [15, 60, 12, 22]  # Ajuste si besoin

            # --- Crée le tableau (une frame pour grid)
            table = tk.Frame(container, bg="#e3f2fd")
            table.pack(anchor="center", padx=10, pady=3, fill="x")

            # --- En-têtes
            for i, titre_col in enumerate(headers):
                tk.Label(
                    table, text=titre_col, font=("Consolas", 10, "bold"),
                    bg="#d1e3f5", anchor="w", width=col_widths[i], relief="flat"
                ).grid(row=0, column=i, sticky="nsew", padx=1, pady=1)

            # --- Données
            steps = [step for step in indexed_steps if step['file'] == trace_label(filepath)]
            for r, etape in enumerate(steps):
                tk.Label(
                    table, text=etape['horodatage'], bg="#e3f2fd",
                    font=("Consolas", 10), anchor="w", width=col_widths[0]
                ).grid(row=r + 1, column=0, sticky="nsew", padx=1, pady=1)
                tk.Label(
                    table, text=etape['checkpoint'], bg="#e3f2fd",
                    font=("Consolas", 10), anchor="w", width=col_widths[1]
                ).grid(row=r + 1, column=1, sticky="nsew", padx=1, pady=1)
                tk.Label(
                    table, text=etape['etat'], bg="#e3f2fd",
                    font=("Consolas", 10, "bold"),
                    fg="#055300" if etape['etat'] == "Présent" else "#d02b20",
                    anchor="w", width=col_widths[2]
                ).grid(row=r + 1, column=2, sticky="nsew", padx=1, pady=1)
                tk.Label(
                    table, text=etape['file'], bg="#e3f2fd",
                    font=("Consolas", 10), anchor="w", width=col_widths[3]
                ).grid(row=r + 1, column=3, sticky="nsew", padx=1, pady=1)

            for i in range(4):
                table.grid_columnconfigure(i, weight=1)

            top = self.convoyage_window
            top.update_idletasks()
            top.geometry("")

            def revenir_resume():
                geo = top.geometry()  # ex : '730x410+622+246'
                top.destroy()
                import re
                m = re.match(r"\d+x\d+\+(-?\d+)\+(-?\d+)", geo)
                if m:
                    x, y = m.group(1), m.group(2)
                    # Lance la fenêtre résumé, elle prendra sa taille naturelle, mais on lui applique la position
                    self.convoyage_action(ref, geometry_override=f"+{x}+{y}")
                else:
                    # fallback : position auto
                    self.convoyage_action(ref, geometry_override=None)

            # --- Bouton retour
            # On récupère la géométrie de la fenêtre actuelle pour la conserver !
            tk.Button(
                container, text="Revenir au résumé", font=("Segoe UI", 10, "bold"),
                command=revenir_resume
            ).pack(pady=12)


        # Ajoute un bloc pour chaque résultat (1 par fichier)
        def add_result(res):
            filepath = res['_filepath']
            filename = trace_label(filepath)
            bloc = tk.Frame(frame, bd=2, relief="groove", bg="#f5fafd")
            bloc.pack(fill="x", expand=True, padx=10, pady=12)

            titre = add_label(
                bloc,
                text=f"📦 Fichier : {filename}",
                font=("Arial", 13, "bold"),
                fg="#00589b",
                bg="#f5fafd",
                wraplength=700,
                justify="left"
            )
            titre.pack(anchor="w", pady=(2, 7), fill="x", expand=True)

            def show_details_in_this_bloc(event=None, container=bloc, ref=ref, filepath=filepath):
                # Récupère la position actuelle de la fenêtre
                top = self.convoyage_window
                geo = top.geometry()
                show_convoyage_steps_in_bloc(container, ref, filepath, geometry_override=geo)

            if '_checkpoint_attente' in res:
                last_pos = res.get('_last_positive')
                if last_pos:
                    cadre_positif = tk.Frame(bloc, bd=1, relief="groove", bg="#e6ffe6")
                    cadre_positif.pack(fill="x", expand=True, padx=12, pady=(0, 7))
                    cadre_positif.config(cursor="hand2")
                    # Ajout des labels
                    lbl = add_label(
                        cadre_positif,
                        text="Dernière position positive connue :",
                        font=("Arial", 11, "underline"),
                        bg="#e6ffe6",
                        wraplength=660,
                        justify="left"
                    )
                    lbl.pack(anchor="w", padx=8, pady=(4, 0), fill="x", expand=True)
                    horo = add_label(
                        cadre_positif,
                        text=f"HORODATAGE: {last_pos['horodatage']}",
                        font=("Arial", 11),
                        bg="#e6ffe6",
                        wraplength=660,
                        justify="left"
                    )
                    horo.pack(anchor="w", padx=8, fill="x", expand=True)
                    chkpt = add_label(
                        cadre_positif,
                        text=f"{last_pos['checkpoint']}: {ref}",
                        font=("Arial", 11, "underline"),
                        fg="green",
                        bg="#e6ffe6",
                        wraplength=660,
                        justify="left"
                    )
                    chkpt.pack(anchor="w", padx=8, pady=(0, 4), fill="x", expand=True)
                    # CLIC PARTOUT SUR LE CARD
                    make_fully_clickable(cadre_positif, lambda event, container=bloc, ref=ref, filepath=filepath: show_details_in_this_bloc(event, container, ref, filepath))


                checkpoint = res['_checkpoint_attente']
                val = res.get(checkpoint, f"-{ref}")
                horodatage = res.get('HORODATAGE', '')
                cadre_attente = tk.Frame(bloc, bd=1, relief="groove", bg="#ffe6e6")
                cadre_attente.pack(fill="x", expand=True, padx=12, pady=(0, 7))
                cadre_attente.config(cursor="hand2")
                horo_a = add_label(
                    cadre_attente,
                    text=f"HORODATAGE: {horodatage}",
                    font=("Arial", 11),
                    bg="#ffe6e6",
                    wraplength=660,
                    justify="left"
                )
                horo_a.pack(anchor="w", padx=8, pady=(4, 0), fill="x", expand=True)
                chkpt_a = add_label(
                    cadre_attente,
                    text=f"{checkpoint}: {val} (en attente)",
                    font=("Arial", 11, "underline"),
                    fg="red",
                    bg="#ffe6e6",
                    wraplength=660,
                    justify="left"
                )
                chkpt_a.pack(anchor="w", padx=8, pady=(0, 4), fill="x", expand=True)
                # CLIC PARTOUT SUR LE CARD
                make_fully_clickable(cadre_positif, lambda event, container=bloc, ref=ref, filepath=filepath: show_details_in_this_bloc(event, container, ref, filepath))



            elif '_checkpoint_actuel' in res:
                checkpoint = res['_checkpoint_actuel']
                val = res.get(checkpoint, ref)
                horodatage = res.get('HORODATAGE', '')
                cadre_positif = tk.Frame(bloc, bd=1, relief="groove", bg="#e6ffe6")
                cadre_positif.pack(fill="x", expand=True, padx=12, pady=(0, 7))
                cadre_positif.config(cursor="hand2")
                horo = add_label(
                    cadre_positif,
                    text=f"HORODATAGE: {horodatage}",
                    font=("Arial", 11),
                    bg="#e6ffe6",
                    wraplength=660,
                    justify="left"
                )
                horo.pack(anchor="w", padx=8, pady=(4, 0), fill="x", expand=True)
                chkpt = add_label(
                    cadre_positif,
                    text=f"{checkpoint}: {val} (dernière position connue)",
                    font=("Arial", 11, "underline"),
                    fg="green",
                    bg="#e6ffe6",
                    wraplength=660,
                    justify="left"
                )
                chkpt.pack(anchor="w", padx=8, pady=(0, 4), fill="x", expand=True)
                # CLIC PARTOUT SUR LE CARD
                make_fully_clickable(cadre_positif, lambda event, container=bloc, ref=ref,
                                                           filepath=filepath: show_details_in_this_bloc(event,
                                                                                                        container, ref,
                                                                                                        filepath))

        self._render_items(all_results, add_result, top)

        # Note finale globale
        note = add_label(
            frame,
            text="Pour plus de détails, vérifiez aussi l’historique dans les fichiers TracModbus.",
            font=("Arial", 10, "italic"),
            fg="#8a6d3b",
            bg="#fff8f0",
            wraplength=700,
            justify="left"
        )
        note.pack(pady=(16, 2), fill="x", expand=True)

        def update_all_wrap(event):
            new_width = event.width - 40
            for lbl in label_refs:
                try:
                    lbl.config(wraplength=max(new_width, 80))
                except tk.TclError:
                    pass

        frame.bind("<Configure>", update_all_wrap)

        if geometry_override:
            self.convoyage_window.geometry(geometry_override)
        else:
            # Si on affiche le résumé (et pas les détails), on veut toujours forcer la taille auto compacte
            self.convoyage_window.update_idletasks()
            self.convoyage_window.geometry("")  # Force la taille minimale du contenu

        self.convoyage_window.lift()


    def _render_convoyage_errors(self, erreurs_modbus):
        WIDTH = 500

        top = tk.Toplevel(self)
        apply_icon(top)
        top.title("Erreurs Convoyage")
        top.resizable(True, True)
        top.geometry(f"{WIDTH}x500")

        frame = tk.Frame(top, bg="#fff8f0")
        frame.pack(fill="both", expand=True, padx=18, pady=18)

        titre = tk.Label(
            frame,
            text="🚩 Anomalies de Convoyage",
            font=("Arial", 16, "bold"),
            fg="#C81B16",
            bg="#fff8f0"
        )
        titre.pack(pady=(0, 12))

        if erreurs_modbus:
            expl = tk.Label(
                frame,
                text="Certains ModBusId ne sont jamais arrivés à destination (T5/R1 ou T5/R2) :",
                font=("Arial", 12),
                fg="#2a2a2a",
                bg="#fff8f0",
                wraplength=480,
                justify="left"
            )
            expl.pack(anchor="w", padx=10, fill="x", expand=True)

            def update_wrap(event):
                expl.config(wraplength=event.width - 40)

            frame.bind("<Configure>", update_wrap)
            top.bind("<Configure>", update_wrap)

            # Bloc en surbrillance, SCROLLABLE
            box = tk.Frame(frame, bd=2, relief="groove", bg="#ffebee")

            if len(erreurs_modbus) <= 2:
                box.pack(fill="x", expand=False, pady=16, padx=12)
            else:
                box.pack(fill="both", expand=True, pady=16, padx=12)

            canvas = tk.Canvas(box, bg="#ffebee", highlightthickness=0)
            scrollbar = tk.Scrollbar(box, orient="vertical", command=canvas.yview)
            inner_frame = tk.Frame(canvas, bg="#ffebee")

            # Ajuste la hauteur du canvas selon le nombre de résultats (1 ligne ≈ 52px, ajuste si besoin)
            if len(erreurs_modbus) == 1:
                canvas.config(height=52)
            elif len(erreurs_modbus) == 2:
                canvas.config(height=96)
            # sinon ne rien faire, il prendra la taille par défaut et scrollera si besoin

            canvas.configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side="right", fill="y")
            canvas.pack(side="left", fill="both", expand=True)
            window = canvas.create_window((0, 0), window=inner_frame, anchor="nw")

            label_refs = []
            def add_convoyage_error(item):
                filepath, mb = item
                filename = trace_label(filepath)
                lbl = tk.Label(
                    inner_frame,
                    text=f"🡺  {filename}   |   ModBusId {mb}",
                    font=("Segoe UI", 12, "bold"),
                    fg="#1960D6",
                    bg="#e8f0fe",
                    anchor="w",
                    cursor="hand2",
                    padx=12, pady=5,
                    relief="flat",
                    borderwidth=1,
                    wraplength=WIDTH - 60
                )
                lbl.pack(fill="x", padx=8, pady=4, ipady=2)
                label_refs.append(lbl)

                def on_enter(e): e.widget.configure(bg="#d1e0ff", fg="#0b3d91")

                def on_leave(e): e.widget.configure(bg="#e8f0fe", fg="#1960D6")

                lbl.bind("<Enter>", on_enter)
                lbl.bind("<Leave>", on_leave)

                def on_click_modbus(event, mbid=mb, fname=filepath):
                    parent = event.widget.winfo_toplevel()
                    parent.update_idletasks()
                    parent_x = parent.winfo_rootx()
                    parent_y = parent.winfo_rooty()
                    parent_w = parent.winfo_width()
                    parent_h = parent.winfo_height()
                    popup_x = parent_x
                    popup_y = parent_y + parent_h + 16
                    self.convoyage_action(mbid, geometry_override=f"+{popup_x}+{popup_y}", only_filepath=fname)

                lbl.bind("<Button-1>", on_click_modbus)

            self._render_items(erreurs_modbus, add_convoyage_error, top)

            # Ajuste scroll + wrap dynamique
            def on_frame_configure(event):
                canvas.configure(scrollregion=canvas.bbox("all"))
                canvas.itemconfig(window, width=canvas.winfo_width())
                for lbl in label_refs:
                    lbl.config(wraplength=max(canvas.winfo_width() - 32, 60))

            inner_frame.bind("<Configure>", on_frame_configure)

            # >>> Correction affichage direct :
            inner_frame.update_idletasks()
            on_frame_configure(None)
            # Hack ultime : forcer le recalc du scrollregion après un délai
            top.after(200, lambda: on_frame_configure(None))

            # Molette souris intuitive
            def _on_mousewheel(event):
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

            def _on_mousewheel_linux(event):
                if event.num == 4:
                    canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    canvas.yview_scroll(1, "units")

            def bind_mousewheel(event):
                canvas.bind_all("<MouseWheel>", _on_mousewheel)
                canvas.bind_all("<Button-4>", _on_mousewheel_linux)
                canvas.bind_all("<Button-5>", _on_mousewheel_linux)

            def unbind_mousewheel(event):
                canvas.unbind_all("<MouseWheel>")
                canvas.unbind_all("<Button-4>")
                canvas.unbind_all("<Button-5>")

            canvas.bind("<Enter>", bind_mousewheel)
            canvas.bind("<Leave>", unbind_mousewheel)

            note = tk.Label(
                frame,
                text="Chaque ModBusId est contrôlé uniquement dans son propre fichier.",
                font=("Arial", 10, "italic"),
                fg="#8a6d3b",
                bg="#fff8f0"
            )
            note.pack(pady=(12, 0))
            # --- Ajustement taille fenêtre ---
            top.update_idletasks()
            if len(erreurs_modbus) <= 1:
                top.geometry("")

        else:
            ok = tk.Label(
                frame,
                text="✅ Toutes les références sont bien arrivées à destination (dans chaque fichier) !",
                font=("Arial", 14, "bold"),
                fg="#228B22",
                bg="#fff8f0"
            )
            ok.pack(pady=16)
            # --- Ajustement taille fenêtre ---
            top.update_idletasks()
            top.geometry("")


    def _render_indexed_convoyage_recap(self, ref, recap):
        checkpoint_indices = list(CHECKPOINTS)
        checkpoints = CHECKPOINTS
        top = tk.Toplevel(self)
        apply_icon(top)
        top.title(f"Parcours détaillé du ModBusId {ref}")
        top.geometry("850x550")
        top.configure(bg="#f4f6fb")
        frame = tk.Frame(top, bg="#f4f6fb")
        frame.pack(fill='both', expand=True, padx=16, pady=10)

        tk.Label(frame, text=f"Parcours détaillé du ModBusId {ref} :", font=("Segoe UI", 13, "bold"),
                 bg="#f4f6fb", fg="#1a5276").pack(anchor="w", pady=(0, 10))

        # Entête : Horodatage d'abord
        entete = tk.Frame(frame, bg="#d1e3f5")
        entete.pack(fill="x", pady=(0, 1))
        for i, titre in enumerate(["Horodatage", "Checkpoint", "État", "Fichier"]):
            tk.Label(entete, text=titre, font=("Segoe UI", 10, "bold"), bg="#d1e3f5", anchor="w").grid(row=0, column=i,
                                                                                                       sticky="ew",
                                                                                                       padx=6, pady=3)

        for col in checkpoint_indices:
            etat = recap[col]["etat"]
            if not etat:
                continue
            row = tk.Frame(frame, bg="#e3f2fd")
            row.pack(fill="x", pady=1)
            tk.Label(row, text=recap[col]["horodatage"], bg="#e3f2fd", font=("Consolas", 10), anchor="w",
                     width=17).grid(row=0, column=0, sticky="ew", padx=6)
            tk.Label(row, text=checkpoints[col], bg="#e3f2fd", font=("Segoe UI", 10), anchor="w", width=38).grid(row=0,
                                                                                                                 column=1,
                                                                                                                 sticky="ew",
                                                                                                                 padx=6)
            tk.Label(row, text=etat, bg="#e3f2fd", font=("Segoe UI", 10, "bold"),
                     fg="#055300" if etat == "Présent" else "#d02b20", anchor="w", width=15).grid(row=0, column=2,
                                                                                                  sticky="ew", padx=6)
            tk.Label(row, text=recap[col]["file"], bg="#e3f2fd", font=("Consolas", 10), anchor="w", width=20).grid(
                row=0, column=3, sticky="ew", padx=6)

        tk.Button(frame, text="Fermer", command=top.destroy, bg="#dbeafe", font=("Segoe UI", 10, "bold")).pack(pady=12)


    def _render_indexed_comparison(self, refprod, pair):
        first, second = pair
        current_file, next_file = first['path'], second['path']
        idpk_line, idpk_line_next = first['text'], second['text']
        dims_current, dims_prev = json.loads(first['dims']), json.loads(second['dims'])
        # Dates déjà indexées : aucun parcours de trace pour cette fenêtre.
        def file_date(event):
            stamp = self._file_metadata.get(event['path'], {}).get('end', event['time'])
            return datetime.datetime.fromtimestamp(stamp).strftime("%d/%m/%Y %H:%M")
        date_current, date_prev = file_date(first), file_date(second)

        # 6. Fonction pour formatter joliment la ligne
        def format_idpk_line(line):
            parsed = parse_idpk(line)
            if not parsed:
                return line
            day, hour = timestamp_parts(line)
            values = ' '.join(f'{key}={value}' for key, value in parsed[1].items())
            return f"{day}|{hour[:5]}   Mesure IDPK ({parsed[0]}) : {values}"

        # 7. Popup élégante & compacte
        top = tk.Toplevel(self)
        apply_icon(top)
        top.title(f"Comparaison IDPK – {refprod}")
        # --- Positionner le popup à gauche de la fenêtre principale ---
        self.update_idletasks()
        main_x = self.winfo_rootx()
        main_y = self.winfo_rooty()
        largeur_popup = 425  # ou la largeur de ta popup
        top.geometry(f"{largeur_popup}x255+{main_x - largeur_popup - 18}+{main_y + 16}")

        top.configure(bg="#f4f6fb")
        top.resizable(False, False)



        frame = tk.Frame(top, bg="#f4f6fb")
        frame.pack(fill='both', expand=True, padx=16, pady=7)

        # Titre petit
        tk.Label(frame, text=f"Comparaison Résultat IDPK pour {refprod}",
                 font=("Segoe UI", 11, "bold"), fg="#00589b", bg="#f4f6fb").pack(anchor="center", pady=(0, 5))

        # Bloc dimensions actuelles
        tk.Label(
            frame,
            text=f"🟦 Fichier actuel : {trace_label(current_file)}   ({date_current})",
            anchor="w",
            font=("Segoe UI", 9, "italic"),
            fg="#226a9e",
            bg="#f4f6fb"
        ).pack(fill="x")
        tk.Label(
            frame,
            text=format_idpk_line(idpk_line),
            font=("Consolas", 9),
            bg="#e3f2fd",
            fg="#003366",
            anchor="w",
            justify="left"
        ).pack(fill="x", pady=(0, 2))
        dim_act = f"  l = {dims_current.get('l', '?')}    H = {dims_current.get('H', '?')}    L = {dims_current.get('L', '?')}"
        tk.Label(
            frame,
            text=dim_act,
            font=("Segoe UI", 9, "bold"),
            bg="#e3f2fd",
            fg="#003366"
        ).pack(fill="x", pady=(0, 6))

        # Bloc dimensions précédentes
        tk.Label(
            frame,
            text=f"🟪 Fichier précédent : {trace_label(next_file)}   ({date_prev})",
            anchor="w",
            font=("Segoe UI", 9, "italic"),
            fg="#75299e",
            bg="#f4f6fb"
        ).pack(fill="x")
        tk.Label(
            frame,
            text=format_idpk_line(idpk_line_next),
            font=("Consolas", 9),
            bg="#f6e3fd",
            fg="#3e275a",
            anchor="w",
            justify="left"
        ).pack(fill="x", pady=(0, 2))
        dim_prev = f"  l = {dims_prev.get('l', '?')}    H = {dims_prev.get('H', '?')}    L = {dims_prev.get('L', '?')}"
        tk.Label(
            frame,
            text=dim_prev,
            font=("Segoe UI", 9, "bold"),
            bg="#f6e3fd",
            fg="#3e275a"
        ).pack(fill="x", pady=(0, 4))

        # Affiche delta si dispo
        try:
            delta_l = int(dims_current['l']) - int(dims_prev['l'])
            delta_H = int(dims_current['H']) - int(dims_prev['H'])
            delta_L = int(dims_current['L']) - int(dims_prev['L'])
            delta_txt = f"Δ l = {delta_l:+}      Δ H = {delta_H:+}      Δ L = {delta_L:+}"
            emoji = "✅" if all(abs(x) <= 2 for x in [delta_l, delta_H, delta_L]) else "⚠️"
        except Exception:
            delta_txt = "Différence : N/A"
            emoji = "❓"
        tk.Label(
            frame,
            text=f"{emoji}  {delta_txt}",
            font=("Segoe UI", 10, "bold"),
            fg="#055300" if emoji == "✅" else "#b56010",
            bg="#f4f6fb"
        ).pack(anchor="center", pady=(2, 3))

        # Bouton Fermer petit
        tk.Button(
            frame,
            text="Fermer",
            font=("Segoe UI", 9, "bold"),
            command=top.destroy,
            bg="#dbeafe",
            activebackground="#bcdffb",
            relief="raised"
        ).pack(pady=(2, 2))


    def _render_indexed_adjacent(self, target_id, filepath, data):
        prev_items, next_items, rp_current, differences, current_number = data
        # 4) création / nettoyage de la popup
        if not (hasattr(self, 'adj_popup') and self.adj_popup and self.adj_popup.winfo_exists()):
            self.adj_popup = tk.Toplevel(self)
            apply_icon(self.adj_popup)
            self.adj_popup.title('IDs adjacents')
            self.adj_popup.resizable(True, True)
            self.prev_frame = ttk.LabelFrame(self.adj_popup, text='5 IDs précédents')
            self.current_frame = ttk.LabelFrame(self.adj_popup, text='ID actuel')
            self.next_frame = ttk.LabelFrame(self.adj_popup, text='5 IDs suivants')
            for f, pady in ((self.prev_frame, (10, 5)),
                            (self.current_frame, 5),
                            (self.next_frame, (5, 10))):
                f.pack(fill='both', padx=10, pady=pady, expand=True)
        else:
            for frame in (self.prev_frame, self.current_frame, self.next_frame):
                for w in frame.winfo_children():
                    w.destroy()
        self.adj_popup.title(f'IDs adjacents — {trace_label(filepath)}')

        def add_adjacent_row(frame, mb, rp, fpath, number):
            bad = differences.get((fpath, rp), False)
            marker = "⚠️ " if bad else ""
            lbl = ttk.Label(frame,
                            text=f"{marker}{mb} ({rp})",
                            cursor='hand2',
                            foreground='red' if bad else None)
            lbl.pack(anchor='w', padx=5, pady=2)
            lbl.bind('<Button-1>',
                     lambda e, m=mb, r=rp, fp=fpath, n=number: self._on_adjacent_ref_click(m, r, fp, n))

        # remplir les cadres
        if prev_items:
            for mb, rp, fp, number in reversed(prev_items):
                add_adjacent_row(self.prev_frame, mb, rp, fp, number)
        else:
            ttk.Label(self.prev_frame, text='—').pack(anchor='w', padx=5, pady=2)

        # ID courant
        bad_cur = differences.get((filepath, rp_current), False)
        marker_cur = "⚠️ " if bad_cur else ""
        lbl_cur = ttk.Label(self.current_frame,
                            text=f"{marker_cur}{target_id} ({rp_current})",
                            font=('Segoe UI', 10, 'bold'),
                            cursor='hand2',
                            foreground='red' if bad_cur else None)
        lbl_cur.pack(anchor='center', pady=4)
        lbl_cur.bind('<Button-1>',
                     lambda e, m=target_id, r=rp_current, fp=filepath, n=current_number:
                     self._on_adjacent_ref_click(m, r, fp, n))

        if next_items:
            for mb, rp, fp, number in next_items:
                add_adjacent_row(self.next_frame, mb, rp, fp, number)
        else:
            ttk.Label(self.next_frame, text='—').pack(anchor='w', padx=5, pady=2)

        # repositionnement
        self.adj_popup.update_idletasks()
        w = self.adj_popup.winfo_reqwidth()
        h = self.adj_popup.winfo_reqheight()
        x = self.winfo_rootx() + self.winfo_width() + 10
        y = self.winfo_rooty()
        self.adj_popup.geometry(f"{w}x{h}+{x}+{y}")
        self.adj_popup.lift()


    def _render_indexed_file_view(self, path, lines, highlight_line=None):
        from trace_view import show_trace_view
        return show_trace_view(self, path, lines, highlight_line)



if __name__ == "__main__":
    app = SearchApp()
    app.mainloop()
