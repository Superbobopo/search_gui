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

Le Bureau est localisé par Windows, y compris lorsqu'il est déplacé dans OneDrive. Si plusieurs emplacements contiennent des traces au démarrage, une fenêtre demande lesquels charger : cochez une ou plusieurs sources, puis cliquez sur « Charger la sélection ». Avec un seul emplacement disponible, le chargement est automatique. Le choix vaut pour la session ; redémarrez l'application pour le modifier. R1 et R2 restent chargés ensemble dans les emplacements choisis. Les fichiers du Bureau portent les libellés `Bureau R1` et `Bureau R2`. Les fichiers homonymes restent distincts et les archives restent liées dans leur dossier robot. Les dossiers absents sont tolérés.

Les traces d'exploitation restent locales et sont exclues de Git. Les courts extraits dans `tests/fixtures/` servent aux tests de compatibilité. Les encodages UTF-8 et Windows-1252 et les formats historiques sont pris en charge.

L'index SQLite/FTS5 est créé automatiquement dans `%LOCALAPPDATA%/TracOmega/cache/<identifiant-du-projet>/`. Chaque combinaison de sources dispose d'un fichier `trace_index_sources_<identifiant>.sqlite3` pour ne pas mélanger les résultats ; l'ancien `trace_index.sqlite3` est conservé. La première ouverture d'une combinaison nécessite son indexation, puis son cache est réutilisé. Aucun serveur n'est nécessaire. La première indexation peut prendre plusieurs minutes ; les recherches deviennent disponibles progressivement. Voir [PERFORMANCE.md](PERFORMANCE.md) pour les mesures et leurs limites.

## Vérifications

```powershell
.\.build-venv\Scripts\python.exe -m unittest discover -s tests -v
.\.build-venv\Scripts\python.exe tools\smoke_interface.py
.\.build-venv\Scripts\python.exe tools\check_interactions.py
.\.build-venv\Scripts\python.exe tools\check_interactions.py --script
```

Choisir les vérifications utiles au changement ; ne pas répéter systématiquement les tests déjà réussis. Le contrôle ciblé du choix des sources est `tools/check_source_choice.py`. Les tests utilisent des données temporaires. Les contrôles d'interface exigent Windows avec Tcl/Tk fonctionnel et s'exécutent sur un bureau Windows isolé, sans prendre le focus du bureau utilisateur. Si l'isolation échoue, ils s'arrêtent. `tools/check_real_case.py` nécessite les traces locales et leur cache déjà construit.

Les consignes de maintenance et les commandes complémentaires sont dans [AGENTS.md](AGENTS.md). Les changements, vérifications et limites sont consignés dans [CHANGELOG.md](CHANGELOG.md). Les environnements Python, caches, builds et copies historiques locales sont exclus du dépôt.
