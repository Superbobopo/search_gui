"""Main-thread source choice shown before opening a trace index."""

import tkinter as tk
import customtkinter as ctk

from app_assets import apply_icon


class SourceChoiceDialog(ctk.CTkToplevel):
    def __init__(self, parent, sources, accept, cancel):
        super().__init__(parent)
        self.title("Choisir les traces à charger")
        self.geometry("680x370")
        self.minsize(550, 330)
        self.transient(parent)
        apply_icon(self)
        self.protocol('WM_DELETE_WINDOW', cancel)
        ctk.CTkLabel(self, text="Plusieurs emplacements contiennent des traces.",
                     font=("Segoe UI", 17, "bold")).pack(anchor='w', padx=22, pady=(20, 6))
        ctk.CTkLabel(self, text="Sélectionnez les sources à utiliser pour cette session.").pack(
            anchor='w', padx=22, pady=(0, 12))
        self.choices = []
        self.checkboxes = []
        for name, directory in sources:
            variable = tk.BooleanVar(self, value=False)
            checkbox = ctk.CTkCheckBox(self, text=f"{name}\n{directory}", variable=variable,
                                      command=self._update_button)
            checkbox.pack(anchor='w', fill='x', padx=24, pady=8)
            self.choices.append((directory, variable))
            self.checkboxes.append(checkbox)
        actions = ctk.CTkFrame(self, fg_color='transparent')
        actions.pack(side='bottom', fill='x', padx=22, pady=18)
        ctk.CTkButton(actions, text="Quitter", command=cancel, width=100).pack(side='left')
        self.open_button = ctk.CTkButton(actions, text="Charger la sélection", state='disabled',
                                        command=lambda: accept(self.selected_directories()))
        self.open_button.pack(side='right')

    def selected_directories(self):
        return tuple(directory for directory, variable in self.choices if variable.get())

    def _update_button(self):
        self.open_button.configure(state='normal' if self.selected_directories() else 'disabled')

    def destroy(self):
        # Release Tk variables on the main thread, including checkbox references.
        for checkbox in self.checkboxes:
            checkbox.configure(variable=None)
        self.choices.clear()
        self.checkboxes.clear()
        super().destroy()
