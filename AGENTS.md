# Consignes du projet

## Journal obligatoire

- Avant chaque modification, consulter ce fichier et `CHANGELOG.md` pour connaître les décisions, défauts corrigés, vérifications et limites.
- **Chaque intervention livrée doit mettre à jour le journal du projet dans `CHANGELOG.md`.** Ajouter une entrée datée avec la demande, les changements, la cause des défauts, les commandes réellement exécutées, leurs résultats et les limites restantes. Ne pas remplacer l'historique par la seule description de la dernière intervention.
- Distinguer les fonctionnalités livrées, les problèmes encore ouverts et les vérifications non exécutées. Ne jamais présenter une interaction non testée comme validée.
- Lors d'une correction, ajouter une vérification qui reproduit le défaut pour éviter sa réintroduction.

## Application et comportement à conserver

- Application Windows Python 3.13, Tkinter et CustomTkinter. Le point d'entrée actif est `search_gui.py` à la racine.
- Les dossiers `old/` et `PythonProject1/` sont des copies historiques ; ne pas les modifier pour une évolution de l'application active.
- Conserver les menus, boutons, recherches, navigation, vues de trace, contrôles robot/convoyage, comparaison IDPK et tracking existants, sauf demande explicite.
- Garder la compatibilité avec les anciennes traces et les exemples locaux. Adapter la lecture, pas les fichiers de traces.
- Les sources sont `C:/Apoteka/traces` et `traces/` à côté du script (ou du `.exe` en mode compilé). Charger R1 et R2 ensemble et garder leur provenance.
- Identifier les fichiers par leur chemin complet ; ne jamais résoudre un clic par le seul nom de fichier. Les liens d'archives et comparaisons chronologiques d'un robot restent dans leur dossier source.

## Formats et maintenance

- Centraliser la découverte des fichiers dans `trace_sources.py` et leur interprétation dans `trace_parser.py`.
- Les recherches de l'application passent par `TraceStore` (`trace_store.py`). Les anciens helpers en mémoire restent uniquement pour les tests de compatibilité ; ne pas les réintroduire dans les actions de l'interface.
- Les traces peuvent être en UTF-8 ou Windows-1252. Ne pas supprimer silencieusement les caractères non UTF-8.
- Les champs peuvent être séparés par `<Cr>`, espaces, `/` ou `;`. Comparer les références et identifiants complets, pas leurs préfixes.
- Accepter les libellés historiques sans accents et les libellés accentués. Conserver les lignes originales pour l'affichage et les liens.
- Préférer les dates inscrites dans les traces aux dates de copie des fichiers lorsqu'elles existent.
- Les lignes MODBUS n'ont pas d'en-tête : ne pas inventer de nouveaux noms ou décaler les registres d'après le seul nombre de colonnes. Conserver les positions existantes en l'absence de spécification vérifiable.
- Ignorer les `.sav`, tolérer les dossiers absents, et actualiser les caches/index après modification ou suppression.
- Ne pas modifier les traces fournies, les environnements virtuels ni les artefacts de build lors d'un changement de code. Ne reconstruire l'exécutable que si demandé.

## Indexation et interface réactive

- SQLite/FTS5 est intégré à Python ; aucun serveur n'est nécessaire. Le cache se trouve dans `%LOCALAPPDATA%/TracOmega/cache/<identifiant-du-projet>/trace_index.sqlite3`.
- L'index conserve les lignes et numéros originaux, les champs, événements robot, liens d'archives, lots, mesures IDPK et valeurs des registres MODBUS. Le texte intégral en mémoire est limité aux trois fichiers récemment ouverts.
- Synchroniser uniquement les fichiers ajoutés/modifiés/supprimés. Publier chaque fichier dans une transaction WAL. Pour un fichier actif, indexer un instantané borné et vérifier que le préfixe n'a pas été réécrit ; une simple croissance ne doit pas empêcher la publication.
- Conserver les recherches par sous-chaîne et sans distinction de casse, les préfixes de référence, ainsi que les correspondances exactes d'ID et de critères avancés.
- `indexed_ui.py` coordonne les actions et `background_tasks.py` élimine les réponses dépassées. Les workers ne doivent appeler aucune méthode Tkinter, y compris `after()` ; utiliser les files de messages drainées par le thread principal.
- La sélection courante appartient à l'instance de l'application : `OccurrenceSelection` dans `selection_state.py` conserve le chemin complet, le numéro original, la référence, le ModBusId et les événements. Les actions actives ne doivent ni importer `search_gui` pour partager un état global, ni analyser le texte affiché pour retrouver des identifiants. Le script lancé comme `__main__` et le module importé sont deux modules distincts.
- Les voisins et les clics transmettent le numéro de l'occurrence ; un ID réutilisé ne désigne pas forcément la première ligne de ce fichier. Compléter les cinq voisins distincts après déduplication, dans l'archive liée du même robot.
- Invalider les réponses d'une ancienne sélection, vider l'état après une absence et actualiser la popup d'IDs adjacents après navigation. Garder le double-clic historique GET/PUTE ; vérifier le texte et le surlignage dans la trace effectivement ouverte, y compris si le fichier a changé avant le prochain rafraîchissement.
- Préserver la présentation demandée de 2025 : date/occurrence centrées, couleurs de fond, polices, espacements et séparateurs. Une optimisation de données ne doit pas simplifier silencieusement le rendu. Les modèles historiques restent consultables dans `PythonProject1/search_gui.py`.
- Le cadre modernisé propose Clair/Sombre/Système et une vue Synthèse facultative. Classique reste la vue initiale et conserve son panneau historique clair, même en thème sombre. Changer de thème ou de rendu ne doit ni relancer la recherche ni déplacer la sélection.
- Les recherches sont disponibles progressivement. Indiquer les résultats partiels, et attendre la première indexation complète pour les diagnostics globaux. Alimenter les longues vues par lots.
- L'icône source est `assets/tracomega.svg`. Ses exports PNG transparent et ICO à sept tailles sont fournis. `app_assets.py` résout les ressources Python/PyInstaller ; garder `search_gui.spec` cohérent.

## Vérifications et commandes

- Lancement PowerShell : `.\.build-venv\Scripts\python.exe .\search_gui.py` (sans activation du venv).
- Tests : `.\.build-venv\Scripts\python.exe -m unittest discover -s tests -v`.
- Syntaxe : `.\.build-venv\Scripts\python.exe -m py_compile search_gui.py trace_sources.py trace_parser.py trace_store.py indexed_ui.py background_tasks.py app_assets.py selection_state.py style_config.py ui_theme.py trace_view.py`.
- Contrôle de l'interface avec traces et cache temporaires, fenêtres masquées : `.\.build-venv\Scripts\python.exe tools\smoke_interface.py`.
- Interactions réelles sur données temporaires, fenêtres hors écran : `.\.build-venv\Scripts\python.exe tools\check_interactions.py`, puis la variante `--script`. Cette dernière exécute réellement `python search_gui.py` depuis un autre dossier, avec un pilote de test installé temporairement avant `mainloop` ; elle vérifie aussi le cas des deux modules Python distincts.
- Ne pas se contenter d'appeler les méthodes Python : invoquer les boutons et générer les événements Tkinter de clic/double-clic/clavier. Capturer les exceptions des callbacks et faire échouer le contrôle. Vérifier le chemin/source, le contenu intégral, le numéro et les tags de surlignage ; la création d'une fenêtre seule ne valide pas son contenu.
- Cas réel de la capture fournie par l'utilisateur, sans synchronisation ni écriture du cache : `.\.build-venv\Scripts\python.exe tools\check_real_case.py`. Ce contrôle exige les traces locales et le cache existant ; ce n'est pas une première indexation. Les captures natives Windows de Tk/CustomTkinter peuvent être incomplètes : inspecter leur contenu avant de les conserver comme preuve.
- Mesure du cache courant : `.\.build-venv\Scripts\python.exe tools\benchmark_traces.py --application-cache`. Sans cette option, le benchmark construit un cache indépendant dans `.benchmark/`.
- Comparaison de versions : `.\.build-venv\Scripts\python.exe tools\benchmark_modernisation.py --store chemin\trace_store.py --output .benchmark\rapport-neuf.json`. Chaque exécution exige un rapport et une base absents ; elle copie les traces dans un dossier temporaire avant de mesurer la première indexation et un ajout. Ne pas confondre ce contrôle avec une mesure du cache existant.
- Si Tkinter signale `init.tcl` inutilisable depuis une session restreinte, vérifier le même lancement dans un contexte autorisé : la résolution Windows des chemins `AppData` peut être bloquée malgré un fichier lisible. Ne pas réinstaller Python ou modifier ses permissions sur ce seul message.
- Export reproductible de l'icône sous Windows : `powershell -NoProfile -ExecutionPolicy Bypass -File tools\create_icon.ps1` (uniquement pour ce processus).
- Build, si demandé : `.\.build-venv\Scripts\python.exe -m PyInstaller --clean --noconfirm search_gui.spec`.
- Pour une adaptation de format, ajouter des tests sur de courts extraits représentatifs de R1/R2 et sur le format historique, y compris les collisions de références et les fichiers homonymes.
- Les tests automatisés doivent fonctionner sans ouvrir l'interface ni dépendre des gros fichiers locaux. Vérifier également les traces réelles lorsqu'elles sont présentes.
- Comparer le temps de première indexation, de réouverture, des recherches présentes/absentes et de préparation du résultat. Consigner aussi la taille du cache et la mémoire ; ne pas présenter le temps d'un cache existant comme une première indexation.
