# TracOmega

Application Windows de consultation et d'analyse des traces TracOmega et MODBUS, en Python 3.13 avec Tkinter et CustomTkinter.

Recherche par référence, ID et critères avancés, navigation dans les occurrences, diagnostics robot et convoyage, comparaison IDPK et tracking. L'interface propose les vues Classique et Synthèse ainsi que les thèmes Clair, Sombre et Système.

## Installation et lancement

Depuis PowerShell, avec Python 3.13 installé avec Tcl/Tk :

```powershell
git clone https://github.com/Superbobopo/search_gui.git
cd search_gui
py -3.13 -m venv .build-venv
.\.build-venv\Scripts\python.exe -m pip install -r requirements.txt
.\.build-venv\Scripts\python.exe .\search_gui.py
```

Le point d'entrée actif est `search_gui.py` à la racine.

## Traces et index

L'application découvre les traces dans `C:/Apoteka/traces`, dans `traces/` à côté du script (notamment `traces/trace r1` et `traces/trace r2`) et dans le dossier `traces/` du Bureau Windows :

```text
Bureau/
└── traces/
    ├── R1/
    └── R2/
```

Le Bureau est localisé par Windows, y compris lorsqu'il est déplacé dans OneDrive. Toutes les sources sont chargées ensemble ; les fichiers du Bureau portent les libellés `Bureau R1` et `Bureau R2`. Les fichiers homonymes restent distincts et les archives restent liées dans leur dossier robot. Les dossiers absents sont tolérés.

Les traces d'exploitation restent locales et sont exclues de Git. Les courts extraits dans `tests/fixtures/` servent aux tests de compatibilité. Les encodages UTF-8 et Windows-1252 et les formats historiques sont pris en charge.

L'index SQLite/FTS5 est créé automatiquement dans `%LOCALAPPDATA%/TracOmega/cache/<identifiant-du-projet>/trace_index.sqlite3`. Aucun serveur n'est nécessaire. La première indexation peut prendre plusieurs minutes ; les recherches deviennent disponibles progressivement. Voir [PERFORMANCE.md](PERFORMANCE.md) pour les mesures et leurs limites.

## Vérifications

```powershell
.\.build-venv\Scripts\python.exe -m unittest discover -s tests -v
.\.build-venv\Scripts\python.exe tools\smoke_interface.py
.\.build-venv\Scripts\python.exe tools\check_interactions.py
.\.build-venv\Scripts\python.exe tools\check_interactions.py --script
```

Les tests utilisent des données temporaires. Les contrôles d'interface exigent Windows avec Tcl/Tk fonctionnel. `tools/check_real_case.py` nécessite les traces locales et leur cache déjà construit.

Les consignes de maintenance et les commandes complémentaires sont dans [AGENTS.md](AGENTS.md). Les changements, vérifications et limites sont consignés dans [CHANGELOG.md](CHANGELOG.md). Les environnements Python, caches, builds et copies historiques locales sont exclus du dépôt.
