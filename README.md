# TracOmega

Application Windows de consultation et d'analyse des traces TracOmega et MODBUS, en Python 3.13 avec Tkinter et CustomTkinter.

Recherche par référence, ID et critères avancés, navigation dans les occurrences, diagnostics robot et convoyage, comparaison IDPK et tracking. Synthèse est la seule présentation des résultats, avec les thèmes Clair, Sombre et Système.

Synthèse est la seule présentation des résultats. Elle centre Référence, ID, SN et CodeLot et affiche les cartes GET, PUTE puis IDPK. Les deux mesures IDPK sont côte à côte ; les valeurs récentes deviennent rouges si une dimension diffère de plus de 2 unités, selon la comparaison utilisée par « Comparer IDPK ». Le détail des écarts reste accessible par ce bouton. Un double-clic sur les événements, les mesures ou les emplacements ouvre la trace et surligne la ligne exacte. Les emplacements par lot se lisent dans le défilement principal, sans zone défilante imbriquée, avec « Afficher plus » par groupes de 20.

Les textes de Synthèse et les données de la comparaison IDPK sont sélectionnables à la souris : Ctrl+C ou clic droit « Copier » copie la sélection, même partielle. Ils restent en lecture seule ; la molette conserve le défilement du panneau principal et le double-clic sur un événement ouvre toujours la trace.

Depuis une recherche ModBusId en Synthèse, un clic sur la référence en haut remplit la barre et lance la recherche Référence en conservant l'occurrence sélectionnée. Glisser sur le texte permet toujours de le sélectionner et le copier, sans lancer de recherche.

Le suivi référence présente les dépôts dans un tableau Date/Heure/Robot/Emplacement/Trace, du plus récent au plus ancien. Sélectionner une ligne affiche la période du fichier et le texte original dessous, sélectionnable et copiable. Ctrl+C copie les lignes choisies ; le clic droit permet aussi de copier une cellule. Double-clic ou Entrée ouvre la trace à la ligne du dépôt, y compris depuis l'historique.

« IDs adjacents » présente R1 et R2 côte à côte, avec Heure, ID, Référence, SN et CodeLot. Les boutons 5/5, 5/10 et 5/30 ou les champs Avant/Après et Appliquer choisissent de 0 à 1 000 IDs distincts de chaque côté, pour les deux robots. Le réglage reste pendant la session, puis revient à 5/5 au prochain lancement. L'occurrence sélectionnée fixe le repère horaire ; l'autre robot se place sur son occurrence la plus proche dans le même emplacement source, sans rapprocher les produits par numéro d'ID. L'écart reste visible, avec un avertissement au-delà de cinq minutes. Un clic sélectionne l'occurrence exacte et recale les deux colonnes, en conservant les tableaux pendant l'actualisation ; cliquer sur l'occurrence courante ne recharge rien. Les archives restent consultables après suppression des copies ; si plusieurs générations ont le même instant le plus proche, choisissez celle à consulter. Sans date exploitable ou sans autre robot, une explication remplace sa colonne. Le défilement est commun, les en-têtes restent visibles ; Ctrl+C et le clic droit permettent de copier les lignes, cellules et provenances.

Le rouge des voisins, de Synthèse et de Comparer IDPK utilise la même comparaison : **Référence, SN et CodeLot identiques**, mesures antérieures sur R1 ou R2 du même emplacement. Une différence de plus de 2 unités sur une dimension commune déclenche l'alerte, même si le passage immédiatement précédent était identique. Le détail montre la plus récente des mesures antérieures différentes, avec son libellé explicite ; sans différence, il montre la dernière mesure comparable. Survoler une ligne de voisins affiche ce détail sous les tableaux. Une mesure absente, une identité incomplète ou une association ambiguë est signalée sans emprunter les mesures d'une autre boîte. Une ligne sans rouge n'est donc pas toujours une comparaison réussie : consulter son statut. Les alertes attendent la vérification complète des sources.

Les panneaux déjà ouverts et les cartes Synthèse sont réutilisés lors de la navigation. Le retour à Recherche conserve le résultat, la sélection et le texte surligné ; si les traces ont changé entre-temps, la recherche est actualisée en arrière-plan. Changer Synthèse ou de thème ne relance pas les recherches déjà validées.

L'interface utilise un habillage commun pour les commandes, fenêtres, cartes et tableaux, avec les mêmes boutons principaux, secondaires et discrets en Clair/Sombre. La comparaison IDPK et les parcours de convoyage reprennent ce cadre. Les traces brutes conservent leurs polices et leurs couleurs métier ; leurs commandes sont harmonisées.

« Erreurs robot » présente les erreurs sous forme de cartes et un détail avec Référence, ModBusId, SN, CodeLot, date/heure, ticket original et provenance complète. Un clic consulte le détail ; un double-clic, Entrée ou « Afficher dans la recherche » rejoint l'occurrence et ses voisins. Les données sont sélectionnables et copiables, y compris depuis les archives dont les fichiers sources ont disparu.

La page **Convoyage** contient deux onglets, avec les résultats directement dans l'application. Dans **Parcours**, saisissez un ModBusId puis utilisez Entrée ou « Afficher le parcours » : les fichiers sont à gauche et le parcours sélectionné à droite. **Anomalies** charge les arrivées non retrouvées après vérification complète des sources ; « Actualiser » relance ce contrôle. Les résultats, la sélection et les défilements sont conservés lors des changements d'onglet ou de section. Les nouvelles publications de l'index actualisent l'onglet affiché ; l'autre se met à jour à son prochain affichage. Les erreurs et absences de résultat sont affichées dans la page, avec possibilité de réessayer. Les autres points d'entrée de convoyage conservent leurs fenêtres.

Les cartes indiquent l'ID, le badge **R1/R2**, le fichier et la provenance si plusieurs emplacements sont choisis. **Un MODBUS couvre les deux robots, même rangé dans `trace r1`** : le dossier n'indique pas sa destination et aucune copie séparée par robot n'est nécessaire. Les étapes nomment R1 ou R2 seulement lorsque leur checkpoint l'indique. Un clic ou les flèches clavier actualisent le détail ; la première ligne est sélectionnée à l'ouverture. La dernière position positive enregistrée est verte, l'attente enregistrée rouge et les informations absentes grises. Les parcours restent isolés par fichier exact et emplacement. Le détail est sélectionnable et copiable ; le clic droit copie l'ID, la ligne ou le chemin, également accessible par « Copier chemin » et « Afficher la provenance ». La séparation se règle à la souris, les défilements sont indépendants et la liste passe au-dessus du détail sous 900 px disponibles. La page utilise des textes plus lisibles de 14 px et des informations secondaires de 12 px, en Clair comme en Sombre.

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

L'index SQLite/FTS5 est créé automatiquement dans `%LOCALAPPDATA%/TracOmega/cache/<identifiant-du-projet>/`. Chaque combinaison de sources dispose d'un fichier `trace_index_sources_<identifiant>_history.sqlite3` pour ne pas mélanger les résultats. Les anciens `trace_index.sqlite3` et `trace_index_sources_<identifiant>.sqlite3` restent conservés et sont lus sans modification pour récupérer leurs données OMEGA. La première ouverture d'une combinaison nécessite son indexation, puis son cache est réutilisé. Aucun serveur n'est nécessaire. La première indexation peut prendre plusieurs minutes ; les recherches deviennent disponibles progressivement. Voir [PERFORMANCE.md](PERFORMANCE.md) pour les mesures historiques et leurs limites.

À la réouverture, les résultats déjà en cache sont recherchables immédiatement après le choix des sources. « Cache disponible », « Vérification des sources », « Indexation X/Y » et « Indexation incomplète » distinguent les étapes. La préparation de l'historique et les vérifications se poursuivent en arrière-plan ; les diagnostics globaux attendent leur achèvement. Si l'application a été fermée avant la fin, les fichiers déjà publiés sont réutilisés et seuls ceux qui manquent ou ont changé sont traités. Le fichier interrompu avant validation de sa transaction est retraité seul. Les caches existants sont conservés : l'installation de l'optimisation FTS ne demande pas une nouvelle indexation générale.

Les traces **OMEGA** disposent en plus d'un historique permanent dans `%LOCALAPPDATA%/TracOmega/history/<identifiant-du-projet>/omega_history.sqlite3`. Il contient leurs instantanés complets compressés et ne doit pas être supprimé comme un cache. Copier les nouvelles traces puis lancer l'application : les générations précédentes restent recherchables même lorsque le robot réutilise les noms numérotés ou que les copies sont supprimées. Pendant une session ouverte, les ajouts et changements de taille, date ou identité de fichier sont surveillés toutes les deux secondes. Au démarrage, les empreintes sont vérifiées même si la copie a conservé sa taille et sa date. Une copie identique, y compris renommée, n'ajoute pas de doublon dans un même dossier robot ; un ajout en fin de fichier complète sa génération. Un préfixe réécrit ou une troncature conserve l'ancienne génération et en crée une autre. Réimporter un instantané déjà observé conserve la génération la plus complète.

Les résultats portent le chemin d'origine, la période et un identifiant de génération. La mention **Historique** indique que cette génération complète n'est plus présente dans les copies sélectionnées ; l'ouverture et les doubles-clics GET/PUTE utilisent son texte sauvegardé, avec ses numéros de ligne originaux. Les sources disposant uniquement d'un historique restent proposées au démarrage. Aucune donnée d'une source non choisie n'est affichée. **MODBUS conserve le fonctionnement de miroir des fichiers présents**, sans historique cumulatif.

L'archivage est publié avant l'indexation : une reconstruction du cache de recherche peut repartir de l'historique. Une erreur ou une version incompatible de l'historique ne déclenche jamais sa suppression automatique. Il n'existe aucune purge automatique. Seules les traces effectivement enregistrées ou encore récupérables dans les anciens caches peuvent être conservées ; les périodes jamais copiées ne sont pas récupérables. Le stockage reste propre à ce poste et à l'emplacement du projet ; déplacer le projet ou changer de poste ne transfère pas automatiquement l'historique. Avant une migration de son schéma, utiliser la sauvegarde SQLite dédiée, qui inclut les transactions WAL validées.

## Vérifications

```powershell
.\.build-venv\Scripts\python.exe -m unittest discover -s tests -v
.\.build-venv\Scripts\python.exe tools\smoke_interface.py
.\.build-venv\Scripts\python.exe tools\check_interactions.py
.\.build-venv\Scripts\python.exe tools\check_interactions.py --script
```

Choisir les vérifications utiles au changement ; ne pas répéter systématiquement les tests déjà réussis. Le contrôle ciblé du choix des sources est `tools/check_source_choice.py`. Les tests utilisent des données temporaires. Les contrôles d'interface exigent Windows avec Tcl/Tk fonctionnel et s'exécutent sur un bureau Windows isolé, sans prendre le focus du bureau utilisateur. Si l'isolation échoue, ils s'arrêtent. `tools/check_real_case.py` nécessite les traces locales et leur cache déjà construit.

Historique OMEGA uniquement : `py -3.13 -B -m unittest tests.test_omega_history -v`, puis `py -3.13 -B tools/check_omega_history.py` pour les vrais doubles-clics sur un bureau privé. `py -3.13 -B tools/check_omega_history_samples.py` vérifie des extraits R1/R2 réels de 64 Kio maximum dans des bases temporaires, sans toucher aux sources ou au cache courant.

Démarrage et ingestion FTS : `py -3.13 -B -m unittest tests.test_omega_performance -v`, puis `py -3.13 -B tools/check_omega_startup.py` pour une recherche et des doubles-clics pendant une préparation volontairement bloquée. `tools/benchmark_omega_startup.py --output chemin-neuf.json` compare les quatre fichiers courants R1/R2 dans des bases temporaires ; les rapports et leurs limites sont détaillés dans `PERFORMANCE.md`.

Les consignes de maintenance et les commandes complémentaires sont dans [AGENTS.md](AGENTS.md). Les changements, vérifications et limites sont consignés dans [CHANGELOG.md](CHANGELOG.md). Les environnements Python, caches, builds et copies historiques locales sont exclus du dépôt.
