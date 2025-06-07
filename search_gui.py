import os
import mmap
import threading
import tkinter as tk
from tkinter import ttk
import functools
from functools import partial




from concurrent.futures import ThreadPoolExecutor
from tkinter import messagebox, scrolledtext
import re
from tkinter import simpledialog


# Chemin fixe vers le dossier contenant les fichiers TracOmega
FIXED_DIRECTORY = r"C:/Apoteka/traces"
BUFFER = {}
current_file = None
BUFFER_MTIME = {}

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 40
        y = y + self.widget.winfo_rooty() + 40
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify="left",
                         background="#ffffe0", relief="solid", borderwidth=1,
                         font=("Segoe UI", 9))
        label.pack(ipadx=1)

    def hide_tip(self, event=None):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()


def reload_modified_files(directory_path):
    """
    Recharge dans le buffer tous les fichiers TracOmega ET TracMODBUS modifiés (ou nouveaux), mais garde les inchangés en mémoire (rapide).
    """
    global BUFFER, BUFFER_MTIME
    new_buffer = {}
    new_mtime = {}
    for root, _, files in os.walk(directory_path):
        for filename in files:
            if not (filename.lower().startswith('tracomega') or filename.lower().startswith('tracmodbus')):
                continue
            path = os.path.join(root, filename)
            try:
                mtime = os.path.getmtime(path)
                new_mtime[path] = mtime
                # Nouveau ou modifié : relire
                if (path not in BUFFER_MTIME) or (mtime != BUFFER_MTIME[path]):
                    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                        try:
                            import mmap
                            with mmap.mmap(f.fileno(), length=0, access=mmap.ACCESS_READ) as m:
                                text = m.read().decode('utf-8', 'ignore')
                        except Exception:
                            f.seek(0)
                            text = f.read()
                    new_buffer[path] = text
                else:
                    # Inchangé : reprendre de l'ancien buffer
                    new_buffer[path] = BUFFER[path]
            except Exception:
                continue
    BUFFER = new_buffer
    BUFFER_MTIME = new_mtime

def load_files_into_buffer(directory_path):
    buffer = {}
    for root, _, files in os.walk(directory_path):
        for filename in files:
            ext = os.path.splitext(filename)[1].lower()
            # if ext == '.sav':
            #     continue
            if not filename.lower().startswith('tracomega'):
                continue
            path = os.path.join(root, filename)
            try:
                with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                    try:
                        with mmap.mmap(f.fileno(), length=0, access=mmap.ACCESS_READ) as m:
                            text = m.read().decode('utf-8', 'ignore')
                    except Exception:
                        f.seek(0)
                        text = f.read()
                buffer[path] = text
            except:
                continue
    return buffer

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
            pattern = re.compile(rf"ModBusId={modbus_id}(\D|$)")
            return [line for line in lines if pattern.search(line)]
        else:
            return []


def search_file_for_keyword(path, keyword, exact_modbus=False):
    text = BUFFER.get(path, "")
    return find_lines_with_keyword(text, keyword, exact_modbus)


def extract_columns_from_line(line, indices, column_names=None):
    valeurs = [v.strip() for v in line.strip().split(';')]
    extraites = [valeurs[i] if i < len(valeurs) else '' for i in indices]
    if column_names:
        return {column_names[n]: extraites[n] for n in range(len(indices))}
    else:
        return extraites


class SearchApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Recherche TracOmega")
        self.geometry("850x740")
        self.configure(bg="#f7f9fa")
        self.resultats = []

        # --- Style ttk (inchangé)
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('TButton', font=("Segoe UI", 11, "bold"),
                        borderwidth=0, focusthickness=3, focuscolor='none',
                        padding=7, relief='flat', background='#2196f3', foreground='white')
        style.map('TButton', background=[('active', '#1565c0')], foreground=[('active', 'white')])
        style.configure('TCheckbutton', font=("Segoe UI", 10), background="#ffffff")
        style.configure('Nav.TButton', font=("Segoe UI", 10, "bold"),
                        padding=(8, 3), borderwidth=0, relief='flat', background='#2196f3', foreground='white')
        style.map('Nav.TButton', background=[('active', '#1565c0')], foreground=[('active', 'white')])
        style.configure('Arrow.TButton', font=("Segoe UI", 10, "bold"), padding=1, background="#2196f3",
                        foreground="#1A5276")
        style.map('Arrow.TButton', background=[('active', '#d0eaff')])

        # ==== MENU LATERAL GAUCHE ====
        self.sidebar = tk.Frame(self, bg="#182A44", width=170)
        self.sidebar.pack(side="left", fill="y")
        tk.Label(self.sidebar, text="MENU", font=("Segoe UI", 13, "bold"),
                 bg="#182A44", fg="#fff", pady=14).pack(anchor="w", padx=18)

        self.btn_convoyage_menu = ttk.Button(self.sidebar, text="🚚 Convoyage", command=self.show_convoyage_tools)
        self.btn_convoyage_menu.pack(fill="x", padx=16, pady=3)
        tk.Label(self.sidebar, text="", bg="#182A44").pack(pady=20)

        # ==== BARRE DE RECHERCHE FIXE ====
        self.search_bar = tk.Frame(self, bg="#ffffff", height=50)
        self.search_bar.pack(side="top", fill="x")

        tk.Label(self.search_bar, text="Mot clé :", font=("Segoe UI", 11), bg="#ffffff").pack(side='left', padx=(12, 5))
        self.keyword = ttk.Entry(self.search_bar, font=("Segoe UI", 11), width=18)
        self.keyword.pack(side='left', padx=3, ipadx=3)
        # self.keyword.focus_set()

        self.keyword.bind('<Return>', lambda event: self.perform_search())


        # Cases à cocher
        self.modbus = tk.BooleanVar(value=True)
        self.reference = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.search_bar, text="ModBusId", variable=self.modbus, command=self.on_modbus_check,
                        style='TCheckbutton').pack(side='left', padx=3)
        ttk.Checkbutton(self.search_bar, text="Référence", variable=self.reference, command=self.on_reference_check,
                        style='TCheckbutton').pack(side='left', padx=3)

        # Navigation
        self.btn_up = ttk.Button(self.search_bar, text="▲", width=2, command=self.on_up, style="Arrow.TButton")
        self.btn_down = ttk.Button(self.search_bar, text="▼", width=2, command=self.on_down, style="Arrow.TButton")
        self.btn_up.pack(side='left', padx=3)
        self.btn_down.pack(side='left', padx=3)

        # Bouton Comparer IDPK
        self.btn_compare_idpk = ttk.Button(
            self.search_bar,
            text="Comparer IDPK",
            width=14,
            command=self.compare_previous_idpk_popup
        )
        self.btn_compare_idpk.pack(side='left', padx=(16, 0), pady=(1, 1))

        # ==== PANEL PRINCIPAL (le seul qu'on vide et remplit selon l'écran) ====
        self.main_panel = tk.Frame(self, bg="#f7f9fa")
        self.main_panel.pack(side="top", fill="both", expand=True)

        self.show_search_panel()

        # Chargement du buffer TracOmega
        global BUFFER
        reload_modified_files(FIXED_DIRECTORY)
        self.lines = []
        self.line_index = 0
        self.file_list = sorted(BUFFER.keys(), key=lambda x: os.path.getmtime(x), reverse=True)
        self.file_index = 0

        # ATTENTION : PAS DE tag_bind sur self.txt ici !
        # Les tag_bind DOIVENT être faits UNIQUEMENT juste après la création de self.txt

    def convoyage_action(self):
        messagebox.showinfo("Convoyage", "Action Convoyage à définir ici !")

    def show_convoyage_tools(self):
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

        btn1 = ttk.Button(
            frame,
            text="🔍 Voir le parcours d’un ModBusId",
            command=self.convoyage_action,
            width=28
        )
        btn1.pack(fill="x", pady=(0, 8))
        # Explications sous chaque bouton
        tk.Label(frame, text="Affiche la dernière position d’un ModBusId dans TracMODBUS.",
                 font=("Segoe UI", 9), bg="#f7f9fa", fg="#555").pack(pady=(0, 14))


        btn2 = ttk.Button(
            frame,
            text="🚩 Lister les erreurs de convoyage",
            command=self.convoyage_errors_action,
            width=28
        )
        btn2.pack(fill="x", pady=(0, 8))

        # Explications sous chaque bouton
        tk.Label(frame, text="Liste tous les ModBusId jamais arrivés à destination dans les fichiers TracMODBUS.",
                 font=("Segoe UI", 9), bg="#f7f9fa", fg="#b31c1c").pack()

        # Bouton retour vers la recherche principale
        ttk.Button(self.main_panel, text="Retour à la recherche", command=self.show_search_panel).pack(pady=8)

        self.keyword.unbind('<Return>')  # Retire l'ancien comportement
        self.keyword.bind('<Return>', self._convoyage_entry_enter)

    def show_search_panel(self):
        # Efface le main_panel
        for widget in self.main_panel.winfo_children():
            widget.destroy()

        self.keyword.unbind('<Return>')
        self.keyword.bind('<Return>', lambda event: self.perform_search())

        # --- Header fichier centré avec navigation fichiers ---
        header_frame = tk.Frame(self.main_panel, bg="#f7f9fa")
        header_frame.pack(fill='x', pady=(13, 7))
        self.btn_prev_file = ttk.Button(header_frame, text="❮", width=1, command=self.on_left)
        self.btn_prev_file.pack(side='left', padx=(12, 0))
        self.file_label = tk.Label(
            header_frame, text="Fichier : ", font=("Segoe UI", 12, "bold"),
            bg="#f7f9fa", fg="#405060"
        )
        self.file_label.pack(side='left', expand=True, padx=8)
        self.btn_next_file = ttk.Button(header_frame, text="❯", width=1, command=self.on_right)
        self.btn_next_file.pack(side='left', padx=(0, 12))

        # --- Actions robot/convoyage en cards plus petits sous la recherche ---
        actions_frame = tk.Frame(self.main_panel, bg="#f7f9fa")
        actions_frame.pack(fill='x', padx=18, pady=(0, 7))

        robot_card = tk.LabelFrame(
            actions_frame, text="🤖 Robot", font=("Segoe UI", 9, "bold"),
            bg="#e3f2fd", fg="#1565c0", bd=1, relief="groove",
            padx=6, pady=4, labelanchor="n"
        )
        robot_card.pack(side='left', padx=(0, 12), ipadx=2, ipady=1, fill="both", expand=True)
        btn_frame = tk.Frame(robot_card, bg="#e3f2fd")
        btn_frame.pack(fill="x", pady=2)

        ttk.Button(btn_frame, text="Ouvrir trace", command=self.open_file_view).pack(side="left", expand=True, fill="x",
                                                                                     padx=(0, 4))
        btn_erreurs = ttk.Button(btn_frame, text="Erreurs Robot", command=self.check_errors)
        btn_erreurs.pack(side="left", expand=True, fill="x", padx=(4, 0))

        ToolTip(btn_erreurs,
                "Vérifie que chaque ModBusId arrivé sur T5 a bien été pris par le robot.\n"
                "Attention : cette vérification s’applique uniquement sur un Robot R1 ou R2"
                )

        # --- Résultats ---
        result_frame = tk.Frame(self.main_panel, bg="#ffffff", bd=1, relief="groove", highlightthickness=1,
                                highlightbackground="#cfd8dc")
        result_frame.pack(fill='both', expand=True, padx=24, pady=(5, 17))
        self.txt = scrolledtext.ScrolledText(result_frame, wrap='word', font=("Consolas", 11), bg="#fafbfc",
                                             borderwidth=0)
        self.txt.pack(fill='both', expand=True, padx=8, pady=8)

        # Re-binde les clics (à faire ICI après la création de self.txt)
        self.txt.tag_bind("clickable_get", "<Double-1>", self.on_double_click_get)
        self.txt.tag_bind("clickable_pute", "<Double-1>", self.on_double_click_pute)
        self.txt.tag_bind("clickable_refprod", "<Button-1>", self.on_click_refprod)

    def check_errors(self):
        reload_modified_files(FIXED_DIRECTORY)

        import threading

        def run_check(app_self):
            file_paths = app_self.file_list
            errors = []
            files_checked = []
            for file_path in file_paths:
                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        lines = f.readlines()
                    modbus_refs = {}
                    gripper_gets = set()
                    for line in lines:
                        if "GRIPPER GET" in line:
                            gripper_gets.add(line)
                        if "ModBusId=" in line and "RefProduit=" in line:
                            modbus_match = re.search(r"ModBusId=(\d+)", line)
                            ref_match = re.search(r"RefProduit=([\w\.]+)", line)
                            if modbus_match and ref_match:
                                modbus_id = modbus_match.group(1)
                                ref = ref_match.group(1)
                                modbus_refs[modbus_id] = ref
                    for modbus_id, ref in modbus_refs.items():
                        found = any(
                            ref in gr_line and gr_line.strip().endswith('$0101')
                            for gr_line in gripper_gets
                        )
                        if not found:
                            errors.append({
                                "file": os.path.basename(file_path),
                                "modbus_id": modbus_id,
                                "ref": ref
                            })
                    files_checked.append(os.path.basename(file_path))
                except Exception as e:
                    errors.append({
                        "file": os.path.basename(file_path),
                        "modbus_id": "-",
                        "ref": f"Erreur lecture fichier: {e}"
                    })
                    files_checked.append(os.path.basename(file_path))

            app_self.after(0, lambda: show_results(errors, files_checked))

        def show_results(errors, files_checked):
            import tkinter as tk

            WIDTH = 500  # Largeur de base de la popup

            top = tk.Toplevel()
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

                # Correct: nom du frame contenant les labels
                inner_frame = tk.Frame(scroll_canvas, bg="#ffeaea")
                window_id = scroll_canvas.create_window((0, 0), window=inner_frame, anchor="nw")

                label_refs = []

                # Ajout des labels dans inner_frame
                for err in errors:
                    err_label = tk.Label(
                        inner_frame,
                        text=f"❌ Fichier: {err['file']} | ModBusId={err['modbus_id']} | RefProduit={err['ref']} non prise par GRIPPER GET",
                        font=("Consolas", 11, "bold"),
                        fg="#b71c1c",
                        bg="#ffeaea",
                        anchor="w",
                        justify="left",
                        wraplength=WIDTH - 40
                    )
                    err_label.pack(fill="x", padx=8, pady=2)
                    label_refs.append(err_label)

                # Correction pour affichage direct !
                scroll_canvas.update_idletasks()
                scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))

                # Ajuste wrap + zone scroll au resize
                def on_frame_configure(event):
                    scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))
                    scroll_canvas.itemconfig(window_id, width=scroll_canvas.winfo_width())
                    for lbl in label_refs:
                        lbl.config(wraplength=max(scroll_canvas.winfo_width() - 32, 60))

                inner_frame.bind("<Configure>", on_frame_configure)

                # Molette souris intuitive (Windows/Mac/Linux)
                def _on_mousewheel(event):
                    # Windows / MacOS
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

        threading.Thread(target=run_check, args=(self,)).start()

    def perform_search(self):
        reload_modified_files(FIXED_DIRECTORY)

        raw = self.keyword.get().strip()
        if not raw:
            messagebox.showwarning("Attention", "Entrez un mot-clé ou un numéro valide !")
            self.btn_compare_idpk.config(state="disabled")
            return
        if self.modbus.get():
            key = f"ModBusId={raw}"
            exact = True
            self.btn_compare_idpk.config(state="disabled")  # Désactive pour ModBusId
        elif self.reference.get():
            key = f"RefProduit={raw}"
            exact = False
            self.btn_compare_idpk.config(state="normal")  # Active pour référence
        else:
            key = raw  # Cas par défaut ou rien n'est coché
            exact = False
            self.btn_compare_idpk.config(state="disabled")

        global current_file
        found = False
        for i, path in enumerate(self.file_list):
            lines = search_file_for_keyword(path, key, exact_modbus=exact)
            if lines:
                current_file = path
                self.file_index = i
                self.lines = lines
                self.line_index = 0
                self.display_line()
                found = True
                break
        if not found:
            current_file = None
            self.txt.delete(1.0, tk.END)
            self.txt.insert(tk.END, "Aucun fichier trouvé avec ce critère.\n")
            self.file_label.config(text="Fichier : Aucun")

    def on_modbus_check(self):
        if self.modbus.get():
            self.reference.set(False)
            self.btn_up.config(state="normal")
            self.btn_down.config(state="normal")
            self.btn_compare_idpk.config(state="disabled")

    def on_reference_check(self):
        if self.reference.get():
            self.modbus.set(False)
            self.btn_up.config(state="disabled")
            self.btn_down.config(state="disabled")
            self.btn_compare_idpk.config(state="normal")

    def find_gripper_line(self, content, refprod):
        # cherche la ligne GET dans tout le fichier
        for i, line in enumerate(content.splitlines()):
            if refprod in line and "GRIPPER GET 1x bote(s)" in line and "$0101" in line:
                return i, line  # retourne index et ligne
        return None, None

    def find_previous_get_and_idpk(self, content, refprod, current_get_index):
        """
        Parcours à l'envers jusqu'au GRIPPER GET $0101 précédent, puis cherche le Résultat IDPK qui suit.
        """
        lines = content.splitlines()
        for i in range(current_get_index - 1, -1, -1):
            line = lines[i]
            if (refprod in line and "GRIPPER GET 1x bote(s)" in line and "$0101" in line):
                # Trouvé le GET précédent
                # Cherche l'IDPK juste après (dans les 5 lignes suivantes, typiquement)
                for j in range(i + 1, min(i + 6, len(lines))):
                    next_line = lines[j]
                    if f"IDPK pour {refprod}" in next_line:
                        return j, next_line
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
            if refprod in line and "GRIPPER GET 1x bote(s)" in line and "$0101" in line:
                for j in range(i + 1, min(i + 7, len(lines))):
                    next_line = lines[j]
                    if f"IDPK pour {refprod}" in next_line:
                        return j, next_line, current_file

        # 2. Chercher dans les fichiers SUIVANTS (plus anciens)
        try:
            current_idx = self.file_list.index(current_file)
        except ValueError:
            return None, None, None  # Sécurité

        for prev_file in self.file_list[current_idx + 1:]:
            content_prev = BUFFER.get(prev_file, "")
            lines_prev = content_prev.splitlines()
            for i in range(len(lines_prev) - 1, -1, -1):
                line = lines_prev[i]
                if refprod in line and "GRIPPER GET 1x bote(s)" in line and "$0101" in line:
                    for j in range(i + 1, min(i + 7, len(lines_prev))):
                        next_line = lines_prev[j]
                        if f"IDPK pour {refprod}" in next_line:
                            return j, next_line, prev_file
        return None, None, None

    def find_result_idpk_line(self, content, refprod, get_index):
        """
        Recherche la première ligne Rsultat de la mesure IDPK pour refprod APRES la ligne GRIPPER GET.
        """
        lines = content.splitlines()
        for i in range(get_index + 1, len(lines)):
            l = lines[i]
            if f"IDPK pour {refprod}" in l:
                return i, l
        return None, None

    def find_gripper_pute_line(self, content, refprod, start_index):
        # cherche la ligne PUTE à partir de start_index + 1
        lines = content.splitlines()
        for i in range(start_index + 1, len(lines)):
            line = lines[i]
            if refprod in line and "GRIPPER PUTE 1x bote(s)" in line:
                return i, line
        return None, None

    def get_separator(self):
        # Largeur actuelle de la zone de texte (en caractères)
        # On prend la largeur du widget (en pixels) divisé par la taille moyenne d’un caractère (8-9 px pour Courier 12)
        width_px = self.txt.winfo_width()
        if width_px < 400:  # Valeur par défaut au lancement
            width_px = 700
        char_count = int(width_px / 8.5)
        return "-" * max(25, char_count - 2)

    def display_line(self):
        import datetime, os, re
        self.txt.config(state='normal')
        self.txt.delete(1.0, tk.END)
        line = self.lines[self.line_index]

        # STYLES
        self.txt.tag_configure("date", font=("Segoe UI", 10), foreground="#909090", spacing3=0, justify="center")
        self.txt.tag_configure("occ", font=("Segoe UI", 9), foreground="#adb5bd", spacing3=2, justify="center")
        self.txt.tag_configure("divider", foreground="#c7c7c7", font=("Segoe UI", 9), spacing3=1)
        self.txt.tag_configure("brute", font=("Consolas", 9), background="#f5f7fa", foreground="#1e293b",
                               spacing3=2, lmargin1=7, lmargin2=7)
        self.txt.tag_configure("getcard", font=("Segoe UI", 10, "bold"), foreground="#00589b", background="#e3f2fd",
                               spacing3=8, spacing1=4, lmargin1=10, lmargin2=10)
        self.txt.tag_configure("putecard", font=("Segoe UI", 10, "bold"), foreground="#b66800", background="#fff4e3",
                               spacing3=8, spacing1=4, lmargin1=10, lmargin2=10)
        self.txt.tag_configure("lastpute", font=("Segoe UI", 10, "bold"), foreground="#9f8700", background="#fffde3",
                               spacing3=8, spacing1=4, lmargin1=10, lmargin2=10)
        self.txt.tag_configure("none", font=("Segoe UI", 9, "italic"), foreground="#a3a3a3", spacing3=2,
                               justify="center")
        self.txt.tag_configure("refcard", font=("Segoe UI", 10, "bold"), foreground="#00876c", background="#e9fdf4",
                               spacing3=6, justify="center", spacing1=6, lmargin1=6, lmargin2=6)
        self.txt.tag_configure("clickable_refprod", foreground="#00876c", underline=1, font=("Segoe UI", 10, "bold"))
        self.txt.tag_configure("clickable_get", foreground="#00589b", underline=1)
        self.txt.tag_configure("clickable_pute", foreground="#b66800", underline=1)

        # DATE + OCCURRENCE
        file_mtime = os.path.getmtime(current_file)
        file_date = datetime.datetime.fromtimestamp(file_mtime)
        self.txt.insert(tk.END, f"{file_date.strftime('%d/%m/%Y')}\n", "date")
        self.txt.insert(tk.END, f"Occurrence {self.line_index + 1}/{len(self.lines)}\n", "occ")
        self.txt.insert(tk.END, "─" * 56 + "\n", "divider")

        # Ligne brute épurée
        def extract_fields_brut(line):
            m = re.match(r"(\d+)\|(\d{2}:\d{2}:\d{2}(?:\.\d+)?)", line)
            date_part = m.group(1) if m else ''
            heure_part = m.group(2) if m else ''
            refprod_match = re.search(r"RefProduit=([^\s/]+)", line)
            sn_match = re.search(r"SN=([^\s/]+)", line)
            codelot_match = re.search(r"CodeLot=([^\s/]+)", line)
            modbus_match = re.search(r"ModBusId=([^\s/]+)", line)
            parts = []
            if date_part and heure_part:
                parts.append(f"{date_part} | {heure_part}")
            if refprod_match:
                parts.append(f"RefProduit: {refprod_match.group(1)}")
            if sn_match:
                parts.append(f"SN: {sn_match.group(1)}")
            if codelot_match:
                parts.append(f"CodeLot: {codelot_match.group(1)}")
            if modbus_match:
                parts.append(f"ModBusId: {modbus_match.group(1)}")
            # On retourne à la fois le texte ET chaque info séparée pour clic ref
            return "    ".join(parts), {
                "refprod": refprod_match.group(1) if refprod_match else None,
                "modbusid": modbus_match.group(1) if modbus_match else None,
            }

        brut_text, line_fields = extract_fields_brut(line)
        self.txt.insert(tk.END, brut_text + "\n", "brute")
        self.txt.insert(tk.END, "─" * 56 + "\n", "divider")

        # RefProduit trouvé
        refprod = line_fields["refprod"]
        modbusid = line_fields["modbusid"]
        if refprod:
            start_idx = self.txt.index(tk.END)
            self.txt.insert(tk.END, f"RefProduit trouvé : {refprod}\n", ("refcard", "clickable_refprod"))
            end_idx = self.txt.index(tk.END)
            self.txt.tag_add("clickable_refprod", start_idx, end_idx)
            self.txt.insert(tk.END, "─" * 56 + "\n", "divider")

        # GRIPPER GET
        content = BUFFER.get(current_file, "")
        get_index, get_line = self.find_gripper_line(content, refprod)
        if get_line:
            start_idx = self.txt.index(tk.END)
            self.txt.insert(tk.END, f"GRIPPER GET :\n{get_line}\n", ("getcard", "clickable_get"))
            end_idx = self.txt.index(tk.END)
            self.txt.tag_add("clickable_get", start_idx, end_idx)
            self.txt.insert(tk.END, "─" * 56 + "\n", "divider")
            pute_index, pute_line = self.find_gripper_pute_line(content, refprod, get_index)
            if pute_line:
                start_idx = self.txt.index(tk.END)
                self.txt.insert(tk.END, f"GRIPPER PUTE :\n{pute_line}\n", ("putecard", "clickable_pute"))
                end_idx = self.txt.index(tk.END)
                self.txt.tag_add("clickable_pute", start_idx, end_idx)
                self.txt.insert(tk.END, "─" * 56 + "\n", "divider")
            else:
                self.txt.insert(tk.END, "Aucune ligne GRIPPER PUTE trouvée.\n", "none")
                self.txt.insert(tk.END, "─" * 56 + "\n", "divider")
        else:
            self.txt.insert(tk.END, "Aucune ligne GRIPPER GET trouvée.\n", "none")
            self.txt.insert(tk.END, "─" * 56 + "\n", "divider")

        # Dernier emplacement connu (référence)
        if self.reference.get() and refprod:
            content_lines = content.splitlines()
            last_pute_line = None
            for l in reversed(content_lines):
                if f"GRIPPER PUTE 1x bote(s): Code:{refprod}" in l:
                    last_pute_line = l
                    break
            if last_pute_line:
                self.txt.insert(tk.END, "Dernier emplacement connu (GRIPPER PUTE) :\n", "lastpute")
                self.txt.insert(tk.END, last_pute_line + "\n", "lastpute")
                self.txt.insert(tk.END, "─" * 56 + "\n", "divider")

        self.file_label.config(text=f"Fichier : {os.path.basename(current_file)}")

        # --- Correction comportement "intelligent" au clic sur la référence produit
        def on_click_refprod(event):
            # Si on est en mode Recherche ModbusId → switch en Recherche Référence
            if self.modbus.get():
                if refprod:
                    self.reference.set(True)
                    self.modbus.set(False)
                    self.btn_up.config(state="disabled")
                    self.btn_down.config(state="disabled")
                    self.keyword.delete(0, tk.END)
                    self.keyword.insert(0, refprod)
                    self.perform_search()
                else:
                    tk.messagebox.showinfo("Info", "Aucune référence produit trouvée.")
            else:
                # Sinon, switch vers Recherche ModbusId en utilisant modbusid déjà extrait
                if modbusid:
                    self.modbus.set(True)
                    self.reference.set(False)
                    self.btn_up.config(state="normal")
                    self.btn_down.config(state="normal")
                    self.keyword.delete(0, tk.END)
                    self.keyword.insert(0, modbusid)
                    self.perform_search()
                else:
                    tk.messagebox.showinfo("Info", "Aucun ModBusId trouvé dans la ligne brute.")

        self.txt.tag_bind("clickable_get", "<Double-1>", self.on_double_click_get)
        self.txt.tag_bind("clickable_pute", "<Double-1>", self.on_double_click_pute)
        self.txt.tag_bind("clickable_refprod", "<Button-1>", on_click_refprod)
        self.txt.config(state='normal')

    def convoyage_errors_action(self):
        reload_modified_files(FIXED_DIRECTORY)

        import tkinter as tk

        tracmodbus_files = [
            os.path.join(FIXED_DIRECTORY, f)
            for f in os.listdir(FIXED_DIRECTORY)
            if f.lower().startswith("tracmodbus")
        ]
        tracmodbus_files.sort(key=os.path.getmtime, reverse=True)

        if not tracmodbus_files:
            messagebox.showerror("Erreur", "Aucun fichier TracMODBUS trouvé dans le dossier.")
            return

        erreurs_modbus = []  # Liste des tuples (filename, modbusid)

        MODBUS_COLS = [1, 2, 3, 4]
        DEST1 = 6
        DEST2 = 7

        # TRAITEMENT FICHIER PAR FICHIER
        for tracmodbus_file in tracmodbus_files:
            try:
                with open(tracmodbus_file, "r", encoding="utf-8", errors="ignore") as fin:
                    lignes = fin.readlines()
            except Exception:
                continue

            modbus_naissants = set()
            arrived = set()

            for ligne in lignes:
                valeurs = [v.strip() for v in ligne.strip().split(';')]
                modbusid = None
                for col in MODBUS_COLS:
                    if col < len(valeurs):
                        v = valeurs[col]
                        if v and v != "0" and v.lstrip('-').isdigit():
                            modbusid = v.lstrip('-')
                            break
                if modbusid:
                    modbus_naissants.add(modbusid)

            for ligne in lignes:
                valeurs = [v.strip() for v in ligne.strip().split(';')]
                for dest_col in (DEST1, DEST2):
                    if dest_col < len(valeurs):
                        v = valeurs[dest_col]
                        if v and v != "0" and v.lstrip('-').isdigit() and not v.startswith('-'):
                            arrived.add(v)

            not_arrived = sorted([mb for mb in modbus_naissants if mb not in arrived])
            for mb in not_arrived:
                erreurs_modbus.append((os.path.basename(tracmodbus_file), mb))

        WIDTH = 500

        top = tk.Toplevel(self)
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
            for filename, mb in erreurs_modbus:
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

                def on_click_modbus(event, mbid=mb, fname=filename):
                    parent = event.widget.winfo_toplevel()
                    parent.update_idletasks()
                    parent_x = parent.winfo_rootx()
                    parent_y = parent.winfo_rooty()
                    parent_w = parent.winfo_width()
                    parent_h = parent.winfo_height()
                    popup_x = parent_x
                    popup_y = parent_y + parent_h + 16
                    self.convoyage_action(mbid, geometry_override=f"+{popup_x}+{popup_y}")

                lbl.bind("<Button-1>", on_click_modbus)

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

    def on_up(self):
        try:
            num = int(self.keyword.get()) - 1
        except ValueError:
            num = 0
        self.keyword.delete(0, tk.END)
        self.keyword.insert(0, str(num))
        if self.modbus.get() and hasattr(self,
                                         "convoyage_window") and self.convoyage_window and tk.Toplevel.winfo_exists(
                self.convoyage_window):
            self.convoyage_action(str(num))
        else:
            self.perform_search_in_current_file()

    def on_down(self):
        try:
            num = int(self.keyword.get()) + 1
        except ValueError:
            num = 0
        self.keyword.delete(0, tk.END)
        self.keyword.insert(0, str(num))
        if self.modbus.get() and hasattr(self,
                                         "convoyage_window") and self.convoyage_window and tk.Toplevel.winfo_exists(
                self.convoyage_window):
            self.convoyage_action(str(num))
        else:
            self.perform_search_in_current_file()

    def perform_search_in_current_file(self):
        reload_modified_files(FIXED_DIRECTORY)
        raw = self.keyword.get().strip()
        if not raw:
            messagebox.showwarning("Attention", "Entrez un mot-clé ou un numéro valide !")
            self.btn_compare_idpk.config(state="disabled")
            return
        key = f"ModBusId={raw}" if self.modbus.get() else raw

        # Met à jour le bouton
        if self.modbus.get():
            self.btn_compare_idpk.config(state="disabled")
        elif self.reference.get():
            self.btn_compare_idpk.config(state="normal")
        else:
            self.btn_compare_idpk.config(state="disabled")

        global current_file
        if not current_file:
            self.txt.delete(1.0, tk.END)
            self.txt.insert(tk.END, "Aucun fichier courant sélectionné.\n")
            self.file_label.config(text="Fichier : Aucun")
            return

        lines = search_file_for_keyword(current_file, key)
        if lines:
            self.lines = lines
            self.line_index = 0
            self.display_line()
        else:
            self.lines = []
            self.line_index = 0
            self.txt.delete(1.0, tk.END)
            self.txt.insert(tk.END, "Aucune occurrence trouvée dans ce fichier.\n")

    def convoyage_action(self, ref=None, geometry_override=None):
        import tkinter as tk
        import os

        def make_fully_clickable(frame, callback):
            """Rend cliquable tout le cadre ET tous ses enfants."""
            frame.bind("<Button-1>", callback)
            for child in frame.winfo_children():
                child.bind("<Button-1>", callback)

        if not hasattr(self, "convoyage_window"):
            self.convoyage_window = None
            self.convoyage_text_widget = None

        if ref is None:
            ref = self.keyword.get().strip()
        self.last_convoyage_ref = ref

        if not ref:
            tk.messagebox.showwarning("Attention", "Veuillez entrer une référence à rechercher.")
            return

        tracmodbus_files = [
            os.path.join(FIXED_DIRECTORY, f)
            for f in os.listdir(FIXED_DIRECTORY)
            if f.lower().startswith("tracmodbus")
        ]
        tracmodbus_files.sort(key=os.path.getmtime, reverse=True)  # du plus récent au plus ancien

        if not tracmodbus_files:
            tk.messagebox.showerror("Erreur", "Aucun fichier TracMODBUS trouvé dans le dossier.")
            return

        checkpoint_indices = [2, 31, 32, 33, 34, 35, 36, 37, 57, 58, 6, 114, 7, 115]
        indices = [0, 1, 2, 3, 4, 31, 32, 33, 34, 35, 36, 37, 57, 58, 6, 114, 7, 115]
        noms_colonnes = [
            'HORODATAGE', 'Boite entrée', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4',
            'En attente en bas de l’asc POSTE1', 'En attente en bas de l’asc POSTE2',
            'En attente en bas de l’asc POSTE3', 'En attente en bas de l’asc POSTE4',
            'Boite en train d’être convoyée vers l’ascenseur 1 ou 2',
            'Boite arrivée devant l’ascenseur 1',
            'Boite arrivée devant l’ascenseur 2',
            "Boite poussée depuis l’ascenseur 1 vers le T5/R1",
            "Boite poussée depuis l’ascenseur 2 vers le T5/R2",
            "Boite en attente sur T5 / R1", "Unnamed: 114", "Boite en attente sur T5 / R2", "Unnamed: 115"
        ]

        all_results = []

        for tracmodbus_file in tracmodbus_files:
            try:
                with open(tracmodbus_file, "r", encoding="utf-8", errors="ignore") as fin:
                    lignes = fin.readlines()
            except Exception:
                continue

            dernier_resultat = None
            last_positive = None

            for ligne in lignes:
                valeurs = [v.strip() for v in ligne.strip().split(';')]
                # 1. Chercher un positif dans les checkpoints
                positif_trouve = False
                for idx, col in enumerate(checkpoint_indices):
                    if col < len(valeurs) and valeurs[col] == ref:
                        checkpoint_nom = noms_colonnes[indices.index(col)] if col in indices else f"Colonne {col}"
                        etape = extract_columns_from_line(ligne, indices, noms_colonnes)
                        etape['_checkpoint_actuel'] = checkpoint_nom
                        etape['_col_idx'] = col
                        etape['_file'] = os.path.basename(tracmodbus_file)
                        dernier_resultat = etape
                        last_positive = {
                            "horodatage": etape.get('HORODATAGE', ''),
                            "checkpoint": checkpoint_nom,
                            "col_idx": col
                        }
                        positif_trouve = True
                        break
                if positif_trouve:
                    continue
                # 2. S'il n'y a PAS de positif sur cette ligne, regarder les négatifs
                for idx, col in enumerate(checkpoint_indices):
                    if col < len(valeurs) and valeurs[col] == f"-{ref}":
                        checkpoint_nom = noms_colonnes[indices.index(col)] if col in indices else f"Colonne {col}"
                        etape = extract_columns_from_line(ligne, indices, noms_colonnes)
                        etape['_checkpoint_attente'] = checkpoint_nom
                        etape['_col_idx'] = col
                        etape['_file'] = os.path.basename(tracmodbus_file)
                        if last_positive is not None:
                            etape['_last_positive'] = last_positive
                        dernier_resultat = etape
                        break

            if dernier_resultat:
                dernier_resultat['_filepath'] = tracmodbus_file
                all_results.append(dernier_resultat)

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

        def show_convoyage_steps_in_bloc(container, ref, filename, geometry_override=None):
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
            steps = self.get_convoyage_steps(ref, only_filename=filename)
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
        for res in all_results:
            filename = res.get('_file', '?')
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

            def show_details_in_this_bloc(event=None, container=bloc, ref=ref, filename=filename):
                # Récupère la position actuelle de la fenêtre
                top = self.convoyage_window
                geo = top.geometry()
                show_convoyage_steps_in_bloc(container, ref, filename, geometry_override=geo)

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
                    make_fully_clickable(cadre_positif, lambda event, container=bloc, ref=ref, filename=filename: show_details_in_this_bloc(event, container, ref, filename))


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
                make_fully_clickable(cadre_positif, lambda event, container=bloc, ref=ref, filename=filename: show_details_in_this_bloc(event, container, ref, filename))



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
                                                           filename=filename: show_details_in_this_bloc(event,
                                                                                                        container, ref,
                                                                                                        filename))

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

    def on_left(self):
        start_index = self.file_index - 1
        raw = self.keyword.get().strip()
        if not raw:
            return
        if self.modbus.get():
            key = f"ModBusId={raw}"
            exact = True
        elif self.reference.get():
            key = f"RefProduit={raw}"
            exact = False
        else:
            key = raw
            exact = False

        for i in range(start_index, -1, -1):
            path = self.file_list[i]
            lines = search_file_for_keyword(path, key, exact_modbus=exact)
            if lines:
                global current_file
                current_file = path
                self.file_index = i
                self.lines = lines
                self.line_index = 0
                self.display_line()
                break

    def on_right(self):
        start_index = self.file_index + 1
        raw = self.keyword.get().strip()
        if not raw:
            return
        if self.modbus.get():
            key = f"ModBusId={raw}"
            exact = True
        elif self.reference.get():
            key = f"RefProduit={raw}"
            exact = False
        else:
            key = raw
            exact = False

        for i in range(start_index, len(self.file_list)):
            path = self.file_list[i]
            lines = search_file_for_keyword(path, key, exact_modbus=exact)
            if lines:
                global current_file
                current_file = path
                self.file_index = i
                self.lines = lines
                self.line_index = 0
                self.display_line()
                break

    def on_double_click_get(self, event):
        self._open_exact_gripper_line(event, "GRIPPER GET")

    def on_double_click_pute(self, event):
        self._open_exact_gripper_line(event, "GRIPPER PUTE")

    def _open_exact_gripper_line(self, event, gripper_type):
        # Récupère la ligne exacte cliquée
        index = self.txt.index(f"@{event.x},{event.y}")
        line_text = self.txt.get(f"{index} linestart", f"{index} lineend").strip()
        if gripper_type not in line_text:
            # Ne fait rien si ce n'est pas la bonne ligne
            return

        # Ouvre la vue avec cette ligne précise
        self.open_gripper_line_view_exact(gripper_type, line_text)

    def open_gripper_line_view(self, gripper_type):
        if not current_file:
            messagebox.showinfo("Info", "Aucun fichier chargé.")
            return

        content = BUFFER.get(current_file, "")
        lines = content.splitlines()

        # Récupère RefProduit dans la ligne courante
        import re
        current_line = self.lines[self.line_index]
        m = re.search(r"RefProduit=([^/]+)", current_line)
        if not m:
            messagebox.showinfo("Info", "Référence produit introuvable dans la ligne courante.")
            return
        refprod = m.group(1)

        # Cherche la ligne correspondant au type et à la référence produit
        target_line_num = None
        for i, line in enumerate(lines):
            if gripper_type in line and f"Code:{refprod}" in line:
                target_line_num = i + 1
                break

        if target_line_num is None:
            messagebox.showinfo("Info", f"Aucune ligne '{gripper_type}' trouvée avec Code:{refprod}")
            return

        # Ouvre la fenêtre secondaire avec la ligne surlignée
        top = tk.Toplevel(self)
        top.title(f"Vue {gripper_type} dans {os.path.basename(current_file)}")
        top.geometry("800x600")

        text_widget = scrolledtext.ScrolledText(top, wrap='word')
        text_widget.pack(fill='both', expand=True)

        for line in lines:
            text_widget.insert(tk.END, line + "\n")

        start_idx = f"{target_line_num}.0"
        end_idx = f"{target_line_num}.end"
        text_widget.tag_add("highlight", start_idx, end_idx)
        text_widget.tag_config("highlight", background="yellow")
        text_widget.see(start_idx)

    def open_file_view(self):
        # === 1. Si le champ recherche est vide, ouvrir le fichier TracOmega le plus récent et afficher toutes ses lignes ===
        if not self.keyword.get().strip():
            trac_files = [
                f for f in os.listdir(FIXED_DIRECTORY)
                if f.lower().startswith("tracomega")
                   and not f.lower().startswith("tracomegaperm")
            ]
            if not trac_files:
                messagebox.showinfo("Info", "Aucun fichier TracOmega trouvé.")
                return
            trac_files.sort(key=lambda f: os.path.getmtime(os.path.join(FIXED_DIRECTORY, f)), reverse=True)
            fichier_recent = os.path.join(FIXED_DIRECTORY, trac_files[0])

            with open(fichier_recent, 'r', encoding='utf-8', errors='ignore') as f:
                contenu = f.read()
            lignes = contenu.splitlines()
            if not lignes:
                messagebox.showinfo("Info", "Le fichier est vide.")
                return

            global current_file
            current_file = fichier_recent
            if current_file not in self.file_list:
                self.file_list.insert(0, current_file)
            self.file_index = self.file_list.index(current_file)
            self.lines = lignes
            self.line_index = 0
            # Pas de return ici : la suite fonctionne normalement

        # === 2. Comportement habituel sinon ===
        if not current_file or not self.lines:
            messagebox.showinfo("Info", "Aucun fichier ou ligne sélectionnée.")
            return

        content = BUFFER.get(current_file, "")
        lines = content.splitlines()
        keyword = self.lines[self.line_index].strip()

        top = tk.Toplevel(self)
        top.title(f"Vue de {os.path.basename(current_file)}")
        top.geometry("900x650")
        top.configure(bg="#f7f9fa")

        # --- Barre recherche et navigation ---
        search_frame = tk.Frame(top, bg="#f7f9fa")
        search_frame.pack(fill='x', padx=18, pady=(3, 8))

        tk.Label(search_frame, text="Recherche dans ce fichier :", font=("Segoe UI", 10), bg="#f7f9fa",
                 fg="#37474f").pack(side='left', padx=(2, 6))
        local_search = ttk.Entry(search_frame, font=("Segoe UI", 11), width=30)
        local_search.pack(side='left', padx=(2, 8), ipadx=3)

        nav_btns = tk.Frame(search_frame, bg="#f7f9fa")
        nav_btns.pack(side='left', padx=(2, 3))
        btn_search_up = ttk.Button(nav_btns, text="↑", width=2)
        btn_search_up.pack(side='top', pady=(0, 1))
        btn_search_down = ttk.Button(nav_btns, text="↓", width=2)
        btn_search_down.pack(side='top', pady=(1, 0))

        # --- Boutons GET/PUTE à droite sur la même ligne ---
        get_grp = tk.Frame(search_frame, bg="#f7f9fa")
        get_grp.pack(side='left', padx=(18, 2))
        tk.Label(get_grp, text="GET", font=("Segoe UI", 9, "bold"), bg="#b3e5fc", fg="#1565c0", padx=6, pady=1,
                 relief="groove").pack(side='left', padx=(0, 3))
        btn_prev_get = ttk.Button(get_grp, text="⟨", width=2)
        btn_prev_get.pack(side='left', padx=1)
        btn_next_get = ttk.Button(get_grp, text="⟩", width=2)
        btn_next_get.pack(side='left', padx=1)

        pute_grp = tk.Frame(search_frame, bg="#f7f9fa")
        pute_grp.pack(side='left', padx=(8, 0))
        tk.Label(pute_grp, text="PUTE", font=("Segoe UI", 9, "bold"), bg="#ffcdd2", fg="#c62828", padx=6, pady=1,
                 relief="groove").pack(side='left', padx=(0, 3))
        btn_prev_pute = ttk.Button(pute_grp, text="⟨", width=2)
        btn_prev_pute.pack(side='left', padx=1)
        btn_next_pute = ttk.Button(pute_grp, text="⟩", width=2)
        btn_next_pute.pack(side='left', padx=1)

        # --- Zone texte principale ---
        text_widget = scrolledtext.ScrolledText(top, wrap='word', font=("Consolas", 11), bg="#fafbfc", borderwidth=0)
        text_widget.pack(fill='both', expand=True, padx=12, pady=10)

        # --- Tag pour surligner une ligne (GET/PUTE/navigation/occurrence) ---
        text_widget.tag_configure("mainline", background="#e3f2fd", font=("Consolas", 11, "bold"))
        text_widget.tag_configure("getline", background="#b3e5fc")
        text_widget.tag_configure("puteline", background="#ffcdd2")
        text_widget.tag_configure("search", background="#fff59d")

        # Affichage complet de la trace
        for line in lines:
            text_widget.insert(tk.END, line + "\n")

        # --- Fonction de surlignage unique (enlève tout avant) ---
        def clear_all_highlights():
            text_widget.tag_remove("mainline", "1.0", tk.END)
            text_widget.tag_remove("getline", "1.0", tk.END)
            text_widget.tag_remove("puteline", "1.0", tk.END)
            text_widget.tag_remove("search", "1.0", tk.END)
            text_widget.tag_remove("sel", "1.0", tk.END)

        def highlight_line(line_idx, tag):
            clear_all_highlights()
            idx = f"{line_idx + 1}.0"
            end_idx = f"{line_idx + 1}.end"
            text_widget.tag_add(tag, idx, end_idx)
            text_widget.see(idx)
            text_widget.mark_set(tk.INSERT, idx)
            text_widget.tag_add("sel", idx, end_idx)  # Optionnel, pour sélection visible aussi

        # Mise en surbrillance initiale sur la ligne trouvée (ligne brute)
        main_line_index = None
        for i, line in enumerate(lines):
            if line.strip() == keyword.strip():
                main_line_index = i
                break
        if main_line_index is not None:
            highlight_line(main_line_index, "mainline")

        # --- Recherche instantanée + navigation ---
        import re
        def search_all_positions(term):
            if not term:
                return []
            text = text_widget.get("1.0", tk.END)
            results = []
            idx = "1.0"
            while True:
                idx = text_widget.search(term, idx, nocase=1, stopindex=tk.END)
                if not idx:
                    break
                results.append(idx)
                idx = f"{idx}+{len(term)}c"
            return results

        search_state = {'positions': [], 'current': -1}

        def update_search_positions():
            term = local_search.get().strip()
            search_state['positions'] = search_all_positions(term)
            search_state['current'] = -1
            # Highlight all
            text_widget.tag_remove("search", "1.0", tk.END)
            for pos in search_state['positions']:
                text_widget.tag_add("search", pos, f"{pos}+{len(term)}c")
            text_widget.tag_config("search", background="#fff59d")

        def goto_search(offset):
            positions = search_state['positions']
            if not positions:
                return
            # Cherche la position actuelle du curseur dans la liste
            current = text_widget.index(tk.INSERT)
            closest = 0
            for i, pos in enumerate(positions):
                if text_widget.compare(pos, ">=", current):
                    closest = i
                    break
            else:
                closest = 0
            idx = (closest + offset) % len(positions)
            pos = positions[idx]
            text_widget.tag_remove("sel", "1.0", tk.END)
            highlight_line(int(float(pos.split('.')[0])) - 1, "mainline")
            text_widget.tag_add("search", pos, f"{pos}+{len(local_search.get().strip())}c")
            text_widget.tag_config("search", background="#fff59d")
            search_state['current'] = idx

        btn_search_down.config(command=lambda: goto_search(1))
        btn_search_up.config(command=lambda: goto_search(-1))
        local_search.bind('<Return>', lambda e: goto_search(1))
        local_search.bind('<KeyRelease>', lambda e: update_search_positions())
        update_search_positions()

        # --- Navigation GET/PUTE (prochaine/prev dans toute la trace, SURBRILLANCE COLOREE) ---
        def find_lines_containing(pattern):
            return [i for i, line in enumerate(lines) if pattern in line]

        def nav_get(offset):
            get_lines = find_lines_containing("GRIPPER GET")
            if not get_lines:
                return
            cur_idx = int(float(text_widget.index(tk.INSERT))) - 1
            sorted_gets = sorted(get_lines)
            if offset > 0:
                # Aller au GET suivant (strictement après)
                next_indices = [i for i in sorted_gets if i > cur_idx]
                if not next_indices:
                    next_idx = sorted_gets[0]  # Boucle au début
                else:
                    next_idx = next_indices[0]
                highlight_line(next_idx, "getline")
            else:
                # Aller au GET précédent (strictement avant)
                prev_indices = [i for i in sorted_gets if i < cur_idx]
                if not prev_indices:
                    prev_idx = sorted_gets[-1]  # Boucle à la fin
                else:
                    prev_idx = prev_indices[-1]
                highlight_line(prev_idx, "getline")

        def nav_pute(offset):
            pute_lines = find_lines_containing("GRIPPER PUTE")
            if not pute_lines:
                return
            cur_idx = int(float(text_widget.index(tk.INSERT))) - 1
            sorted_putes = sorted(pute_lines)
            if offset > 0:
                # Aller au PUTE suivant (strictement après)
                next_indices = [i for i in sorted_putes if i > cur_idx]
                if not next_indices:
                    next_idx = sorted_putes[0]
                else:
                    next_idx = next_indices[0]
                highlight_line(next_idx, "puteline")
            else:
                # Aller au PUTE précédent (strictement avant)
                prev_indices = [i for i in sorted_putes if i < cur_idx]
                if not prev_indices:
                    prev_idx = sorted_putes[-1]
                else:
                    prev_idx = prev_indices[-1]
                highlight_line(prev_idx, "puteline")

        btn_next_get.config(command=lambda: nav_get(1))
        btn_prev_get.config(command=lambda: nav_get(-1))
        btn_next_pute.config(command=lambda: nav_pute(1))
        btn_prev_pute.config(command=lambda: nav_pute(-1))

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
        import re
        index = self.txt.index(f"@{event.x},{event.y}")
        line_text = self.txt.get(f"{index} linestart", f"{index} lineend")

        if ":" not in line_text:
            return

        refprod = line_text.split(":")[-1].strip()
        if not refprod:
            return

        if self.reference.get():
            # Recherche ModBusId dans toutes les lignes de la zone texte
            modbus_id = None
            lines_total = int(self.txt.index('end').split('.')[0])
            for i in range(1, lines_total):
                check_line = self.txt.get(f"{i}.0", f"{i}.end")
                m = re.search(r"ModBusId=(\d+)", check_line)
                if m:
                    modbus_id = m.group(1)
                    break
            if modbus_id:
                self.modbus.set(True)
                self.reference.set(False)
                self.btn_up.config(state="normal")
                self.btn_down.config(state="normal")
                self.keyword.delete(0, tk.END)
                self.keyword.insert(0, modbus_id)
                self.perform_search()
            else:
                from tkinter import messagebox
                messagebox.showinfo("Info", "Aucun ModBusId trouvé dans la zone affichée.")
        else:
            self.reference.set(True)
            self.modbus.set(False)
            self.btn_up.config(state="disabled")
            self.btn_down.config(state="disabled")
            self.keyword.delete(0, tk.END)
            self.keyword.insert(0, refprod)
            self.perform_search()

    def open_gripper_line_view_exact(self, gripper_type, target_line):
        if not current_file:
            messagebox.showinfo("Info", "Aucun fichier chargé.")
            return

        content = BUFFER.get(current_file, "")
        lines = content.splitlines()

        # Cherche la LIGNE EXACTE
        target_line_num = None
        for i, line in enumerate(lines):
            if line.strip() == target_line.strip():
                target_line_num = i + 1  # +1 car Text widget commence à 1.0
                break

        if target_line_num is None:
            messagebox.showinfo("Info", f"Ligne exacte '{target_line}' introuvable dans {current_file}")
            return

        # Ouvre la fenêtre secondaire avec la ligne surlignée
        top = tk.Toplevel(self)
        top.title(f"Vue {gripper_type} dans {os.path.basename(current_file)}")
        top.geometry("800x600")

        text_widget = scrolledtext.ScrolledText(top, wrap='word')
        text_widget.pack(fill='both', expand=True)

        for line in lines:
            text_widget.insert(tk.END, line + "\n")

        start_idx = f"{target_line_num}.0"
        end_idx = f"{target_line_num}.end"
        text_widget.tag_add("highlight", start_idx, end_idx)
        text_widget.tag_config("highlight", background="yellow")
        text_widget.see(start_idx)

    def extract_idpk_dimensions(self, idpk_line):
        """
        Extrait les dimensions l, H, L d'une ligne Résultat IDPK.
        Exemple ligne : 'B 14|13:38:24.0 Rsultat de la mesure IDPK pour PLH000155.00 : l=64 H=41 L=91'
        """
        import re
        result = {}
        m = re.search(r"l=(\d+)", idpk_line)
        if m:
            result['l'] = m.group(1)
        m = re.search(r"H=(\d+)", idpk_line)
        if m:
            result['H'] = m.group(1)
        m = re.search(r"L=(\d+)", idpk_line)
        if m:
            result['L'] = m.group(1)
        return result

    def compare_previous_idpk_popup(self):
        import os
        import re
        import datetime
        import tkinter as tk

        # 1. Récupère la ref produit courante
        if not self.lines:
            tk.messagebox.showinfo("Info", "Aucune ligne sélectionnée.")
            return
        current_line = self.lines[self.line_index]
        m = re.search(r"RefProduit=([^\s/]+)", current_line)
        if not m:
            tk.messagebox.showinfo("Info", "Impossible d'extraire la référence produit.")
            return
        refprod = m.group(1)

        # 2. Cherche GRIPPER GET dans le fichier courant
        content = BUFFER.get(current_file, "")
        get_idx, get_line = self.find_gripper_line(content, refprod)
        if get_idx is None:
            tk.messagebox.showinfo("Info", "GRIPPER GET non trouvé dans le fichier courant.")
            return

        # 3. Cherche Résultat IDPK APRES ce GET
        idpk_idx, idpk_line = self.find_result_idpk_line(content, refprod, get_idx)
        if idpk_line is None:
            tk.messagebox.showinfo("Info", "Résultat IDPK non trouvé dans le fichier courant.")
            return
        dims_current = self.extract_idpk_dimensions(idpk_line)

        # 4. Trouve le fichier suivant (plus ancien) avec GRIPPER GET, puis IDPK
        file_idx = self.file_list.index(current_file)
        next_file = None
        idpk_line_next = None
        dims_prev = None
        for i in range(file_idx + 1, len(self.file_list)):
            path = self.file_list[i]
            content_next = BUFFER.get(path, "")
            get_idx_next, get_line_next = self.find_gripper_line(content_next, refprod)
            if get_idx_next is not None:
                idpk_idx_next, idpk_line_next = self.find_result_idpk_line(content_next, refprod, get_idx_next)
                if idpk_line_next is not None:
                    next_file = path
                    dims_prev = self.extract_idpk_dimensions(idpk_line_next)
                    break
        if not next_file or not idpk_line_next:
            tk.messagebox.showinfo("Comparaison IDPK",
                                   "Aucune occurrence précédente trouvée dans les fichiers plus anciens.")
            return

        # 5. Récupère les dates de modification des fichiers
        date_current = datetime.datetime.fromtimestamp(os.path.getmtime(current_file)).strftime("%d/%m/%Y %H:%M")
        date_prev = datetime.datetime.fromtimestamp(os.path.getmtime(next_file)).strftime("%d/%m/%Y %H:%M")

        # 6. Fonction pour formatter joliment la ligne
        def format_idpk_line(line):
            m = re.match(
                r"\w\s+(\d{2}\|\d{2}:\d{2})(?::\d+)?(?:\.\d+)?\s+Rsultat de la mesure IDPK pour ([^ :]+)\s*:\s*(.*)",
                line)
            if not m:
                m2 = re.match(
                    r"\w\s+(\d{2}\|\d{2}:\d{2})(?:\.\d+)?\s+Rsultat de la mesure IDPK pour ([^ :]+)\s*:\s*(.*)", line)
                if m2:
                    groups = m2.groups()
                else:
                    return line
            else:
                groups = m.groups()
            if groups:
                horo = groups[0]
                refprod = groups[1]
                valeurs = groups[2]
                return f"{horo}   Mesure IDPK ({refprod}) : {valeurs}"
            return line

        # 7. Popup élégante & compacte
        top = tk.Toplevel(self)
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
            text=f"🟦 Fichier actuel : {os.path.basename(current_file)}   ({date_current})",
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
            text=f"🟪 Fichier précédent : {os.path.basename(next_file)}   ({date_prev})",
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

    def _convoyage_entry_enter(self, event=None):
        modbusid = self.keyword.get().strip()
        if modbusid.isdigit():  # Simple vérif, adapte selon ton format d’id si besoin
            self.convoyage_action(modbusid)

    def show_convoyage_details(self, ref):
        import tkinter as tk

        tracmodbus_files = [
            os.path.join(FIXED_DIRECTORY, f)
            for f in os.listdir(FIXED_DIRECTORY)
            if f.lower().startswith("tracmodbus")
        ]
        tracmodbus_files.sort(key=os.path.getmtime, reverse=True)

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

        recap = {}
        for col in checkpoint_indices:
            recap[col] = {
                "etat": None,
                "horodatage": "",
                "file": ""
            }

        for file in tracmodbus_files:
            try:
                with open(file, "r", encoding="utf-8", errors="ignore") as fin:
                    lignes = fin.readlines()
            except Exception:
                continue

            for l in lignes:
                valeurs = [v.strip() for v in l.strip().split(';')]
                horo = valeurs[0] if len(valeurs) > 0 else ''
                for col in checkpoint_indices:
                    if col < len(valeurs):
                        v = valeurs[col]
                        if v == ref:
                            recap[col]["etat"] = "Présent"
                            recap[col]["horodatage"] = horo
                            recap[col]["file"] = os.path.basename(file)
                        elif v == f"-{ref}" and recap[col]["etat"] is None:
                            recap[col]["etat"] = "Jamais arrivé"
                            recap[col]["horodatage"] = horo
                            recap[col]["file"] = os.path.basename(file)

        top = tk.Toplevel(self)
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

    def _convoyage_details_click(self, ref, event):
        self.show_convoyage_details(ref)

    def get_convoyage_steps(self, ref, only_filename=None):
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
        tracmodbus_files = [
            os.path.join(FIXED_DIRECTORY, f)
            for f in os.listdir(FIXED_DIRECTORY)
            if f.lower().startswith("tracmodbus")
        ]
        tracmodbus_files.sort(key=os.path.getmtime, reverse=True)
        for file in tracmodbus_files:
            if only_filename and os.path.basename(file) != only_filename:
                continue
            try:
                with open(file, "r", encoding="utf-8", errors="ignore") as fin:
                    lignes = fin.readlines()
            except Exception:
                continue
            found = {col: False for col in checkpoint_indices}
            for l in lignes:
                valeurs = [v.strip() for v in l.strip().split(';')]
                horo = valeurs[0] if len(valeurs) > 0 else ''
                for col in checkpoint_indices:
                    if col < len(valeurs):
                        v = valeurs[col]
                        if not found[col] and v == ref:
                            recap.append({
                                "horodatage": horo,
                                "checkpoint": checkpoints[col],
                                "etat": "Présent",
                                "file": os.path.basename(file)
                            })
                            found[col] = True
                        elif not found[col] and v == f"-{ref}":
                            recap.append({
                                "horodatage": horo,
                                "checkpoint": checkpoints[col],
                                "etat": "Jamais arrivé",
                                "file": os.path.basename(file)
                            })
                            found[col] = True
        return recap


if __name__ == "__main__":
    app = SearchApp()
    app.mainloop()