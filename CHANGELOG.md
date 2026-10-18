# Journal du projet TracOmega

Chaque intervention livrée ajoute une entrée ici. Consulter également `AGENTS.md` avant de modifier l'application. Les dates de travaux antérieurs n'étant pas établies, leur historique est reconstitué ci-dessous à la date de consignation ; cela ne leur attribue pas une date de réalisation inventée.

## 2026-10-18 — Préparation et publication GitHub

Demande : mettre ce projet sur GitHub.

Changements livrés : initialisation Git locale et rattachement à `Superbobopo/search_gui`, dépôt existant de la même application ; conservation de son historique `main`. Ajout de `.gitignore`, d'une notice `README.md` en français et de `requirements.txt` avec CustomTkinter 5.2.2, version installée. Les traces locales, environnements, caches, builds et copies historiques sont exclus de la publication. Les courts extraits de tests et les ressources graphiques sont conservés. Identité des commits configurée uniquement dans ce dépôt avec le pseudonyme GitHub et son adresse noreply.

Cause : le dossier livré ne contenait aucun dépôt Git, fichier d'exclusion ni instructions d'installation. Aucun défaut de l'application corrigé et aucun code métier modifié pendant cette intervention.

Commandes et contrôles réellement exécutés :

- `git status --short` et `git remote -v` : absence initiale de dépôt confirmée. Consultation du compte et des dépôts via le connecteur GitHub : dépôt public `Superbobopo/search_gui` retrouvé, avec une ancienne version de `search_gui.py`.
- `git init -b main`, `git remote add origin https://github.com/Superbobopo/search_gui.git`, `git fetch origin main`, `git ls-tree -r --name-only origin/main`, puis `git reset --mixed origin/main` : historique récupéré sans remplacer les fichiers locaux.
- `.\.build-venv\Scripts\python.exe -B -m unittest discover -s tests -v` : **63 tests réussis en 2,933 s**.
- `.\.build-venv\Scripts\python.exe -m py_compile search_gui.py trace_sources.py trace_parser.py trace_store.py indexed_ui.py background_tasks.py app_assets.py selection_state.py style_config.py ui_theme.py trace_view.py` : réussi.
- `git check-ignore` sur les traces, environnements, benchmarks, builds, copies historiques et `.env` : exclusions confirmées. Recherche `rg` de signatures de jetons GitHub/AWS et clés privées dans les sources et documents publiables : aucune correspondance ; ce contrôle ciblé ne constitue pas un audit exhaustif de secrets.
- `git -c credential.interactive=never push --dry-run origin main` : accès distant vérifié, branche encore identique à ce stade.

Limites : aucune installation dans un environnement neuf, interaction Tkinter ou mesure de performances réexécutée pour cette préparation documentaire. Aucun exécutable reconstruit ; traces, cache courant et copies historiques inchangés.

## 2026-10-18 — Audit du design et corrections des états visuels

Demande : trouver et corriger les coquilles de la nouvelle interface sans demander à l'utilisateur de les identifier.

Défauts, causes et corrections livrées :

- Les libellés SN et CodeLot créés après le choix du thème conservaient `#333` sur le fond sombre : contraste mesuré **1,22:1**. Ils utilisent désormais la palette Clair/Sombre et le panneau des filtres reçoit immédiatement le thème.
- Filtres et Copier chemin héritaient du survol des boutons d'action malgré leur fond secondaire et leur texte différent. Leur survol utilise maintenant la couleur de bordure de la palette. Les fonds et survols des segments Classique/Synthèse sont également explicités au lieu d'hériter des gris de CustomTkinter ; les couleurs du segment sélectionné restent distinctes.
- Le Canvas superposé des icônes de navigation conservait son fond bleu au survol et ne suivait pas un changement d'échelle après construction. `NavigationButton` synchronise l'icône lors des survols, sorties, redessins, clics et changements d'échelle ; les événements sur l'icône sont transmis au bouton.
- Les vérifications élargies ont révélé deux défauts de cycle de vie : les fenêtres de trace fermées conservaient leur `StringVar` dans le cycle fenêtre/contrôleur, parfois collecté par un worker (`RuntimeError: main thread is not in main loop`) ; les zones de texte de Synthèse conservaient des callbacks périodiques de scrollbar après leur destruction, avec réutilisation possible d'un nom de commande Tcl et appel sans événement d'un callback `<Configure>`. La variable est libérée dans le callback de destruction du thread principal ; `ManagedTextbox` annule ses propres callbacks avant destruction, sans modifier CustomTkinter installé.
- Régressions ajoutées dans `tools/check_interactions.py` : contraste natif normal/survol, filtres créés sous les trois apparences, fonds des icônes sur leurs deux surfaces, commandes accessibles à largeur logique 860 et échelles simulées 100/150/200 %, libération immédiate des variables des fenêtres fermées, absence des callbacks des zones de texte détruites dans `after info`. Les exceptions de finalisation Python sont maintenant capturées et font échouer le contrôle, comme les exceptions des callbacks Tkinter.

Commandes et vérifications réellement exécutées :

- Scripts ponctuels transmis à `.\.build-venv\Scripts\python.exe -B -` : préparation de traces/cache temporaires via `tools.check_interactions.prepare`, inspection des couleurs natives, géométries et survols. Reproduction du texte sombre des filtres et du rectangle des icônes ; échec initial du contrôle d'interactions sur le contraste SN. L'instrumentation temporaire de `tk.Variable` a identifié les variables de statut de `show_trace_view` lors des échecs ultérieurs par lancement réel.
- Captures ponctuelles par `System.Drawing.Graphics.CopyFromScreen`, enregistrées dans `%TEMP%` et inspectées : avant correction, puis résultats Classique clair, Classique dans le cadre sombre et Synthèse sombre avec filtres ouverts. La tentative préalable d'utiliser `PIL.ImageGrab` a échoué car Pillow n'est pas installé ; aucune dépendance ajoutée.
- `.\.build-venv\Scripts\python.exe -B -m unittest discover -s tests -v` : **63 tests réussis**, dernier passage **2,654 s**.
- `.\.build-venv\Scripts\python.exe -m py_compile search_gui.py trace_sources.py trace_parser.py trace_store.py indexed_ui.py background_tasks.py app_assets.py selection_state.py style_config.py ui_theme.py trace_view.py tools\check_interactions.py` : réussi après les corrections finales.
- `.\.build-venv\Scripts\python.exe -B tools\check_interactions.py` et variante `--script` : **parcours complets finaux réussis**. Les exécutions intermédiaires ont effectivement échoué sur le contraste, les finalisations Tk hors thread et le callback Synthèse sans événement, avant leurs corrections ; elles ne sont pas comptées comme validations. Vérification des vrais clics, doubles-clics, clavier, sources, lignes originales, surlignages, sélection conservée, grande trace, navigation, IDPK, convoyage, tracking et absence de traces.
- `.\.build-venv\Scripts\python.exe -B tools\smoke_interface.py` : réussi après les corrections finales, construction **0,134 s**, recherche disponible pendant l'indexation temporaire.
- `.\.build-venv\Scripts\python.exe -B tools\check_real_case.py` : réussi après les corrections finales, **66 fichiers**, cache ouvert en lecture seule ; Local R2, ticket ligne **65169**, GET **11:00:36.9**, PUTE **11:00:52.3** et doubles-clics avec surlignages exacts. Cache et traces vérifiés inchangés, aucune synchronisation. Ce contrôle utilise le cache existant et ne mesure pas une première indexation.

Limites : échelles simulées dans CustomTkinter, sans modification du réglage Windows ni test de plusieurs écrans physiques. Les captures couvrent les vues indiquées, sans comparaison pixel par pixel exhaustive. Pas de nouveau benchmark d'indexation ni d'audit exhaustif de chaque état de chaque popup. Aucune modification des traces fournies, du cache courant, des environnements virtuels ou des copies historiques ; aucun exécutable reconstruit.

## 2026-10-17 — Console modernisée, Synthèse et actualisation incrémentale

Demande : appliquer le plan de modernisation, conserver Classique par défaut et le cache existant, ajouter Clair/Sombre/Système et Synthèse, fluidifier les grandes traces et mesurer les gains. Inclut le diagnostic de l'échec Tcl/Tk observé dans la session de commandes.

Causes et décisions :

- Les commandes partageaient une ligne chargée ; les fenêtres mélangeaient styles, polices et couleurs. Le thème système coexistait avec des fonds fixes. Le tracking affichait une seconde saisie. La nouvelle présentation conserve les callbacks métier, les attributs utilisés par les contrôles et le panneau Classique historique.
- La recherche locale parcourait entièrement le widget à chaque touche ; les flèches GET/PUTE recalculaient toutes les positions à chaque clic. Les traitements par lots devaient aussi annuler leurs callbacks à la destruction des widgets.
- L'incrémentalité était limitée aux fichiers : une simple croissance déclenchait copie, parsing, suppression et réinsertion du fichier complet. La détection d'encodage relisait l'instantané. Un état de reprise additif évite de reconstruire le cache version 1.
- Tcl/Tk 8.6.15 et `init.tcl` sont installés et compatibles. Dans la session restreinte, `GetLongPathNameW` sous `AppData` échouait avec « accès refusé », tandis que la lecture directe fonctionnait. Le même Python crée correctement une fenêtre dans le contexte autorisé. Aucune réinstallation, modification des ACL ou des environnements virtuels n'a été faite.

Fonctionnalités livrées :

- Cadre bleu nuit/turquoise, commandes sur deux lignes avec adaptation compacte, navigation active, icônes Canvas sans dépendance supplémentaire, recherche unique par vue et mémorisation des saisies pendant la session. Thèmes Système/Clair/Sombre, polices harmonisées, clavier Tab/Entrée/Espace et focus visible.
- Source, fichier, occurrence, infobulle et copie du chemin complet. État d'indexation indépendant des messages de recherche, y compris actualisation incomplète. Fenêtres présentes et futures harmonisées ; Classique conserve ses fonds, polices, centrages, séparateurs et doubles-clics.
- Synthèse facultative avec référence/ID et événements GET/IDPK/PUTE horodatés, boutons ouvrant les lignes originales exactes, événements absents explicites, ligne brute et derniers emplacements. La bascule ne recherche pas à nouveau et conserve l'objet de sélection. La clé additive `idpk` de `TraceStore.details()` correspond au GET sélectionné.
- Recherche locale temporisée de 200 ms et découpée en tranches, annulation des anciennes frappes et des callbacks des fenêtres fermées. Navigation GET/PUTE par positions préparées et recherche binaire. Correction également du premier clic ↑ local qui sélectionnait l'avant-dernier résultat au lieu du dernier.
- Table de reprise `file_resume`, sans changement de `user_version=1` ni reconstruction globale. Vérification SHA-256 du préfixe, parsing de la dernière ligne physique et de l'ajout, reprise des dates/lots et publication WAL atomique. CR/LF/CRLF, lignes incomplètes, UTF-8/1252, réécriture, rotation, troncature et annulation couverts. Copie, empreinte et détection d'encodage combinées lors de la première lecture.
- Outil de comparaison de versions sur copies et bases neuves, mesures par fichier dans `sync()['indexing']`, documentation des performances et consignes de contrôle actualisées.

Vérifications effectivement exécutées :

- `.\.build-venv\Scripts\python.exe -B -m unittest discover -s tests -v` : **63 tests passent en 2,765 s**, dont 17 nouvelles régressions d'indexation. Comparaison incrémentale/indexation complète, conservation des identifiants SQL du préfixe, FTS, MODBUS, états invalides, UTF-8 coupé, en-têtes, lots/dates, homonymes, migration additive et rollback WAL.
- `.\.build-venv\Scripts\python.exe -m py_compile search_gui.py trace_sources.py trace_parser.py trace_store.py indexed_ui.py background_tasks.py app_assets.py selection_state.py style_config.py ui_theme.py trace_view.py tools\benchmark_modernisation.py tools\check_interactions.py` : réussi.
- `.\.build-venv\Scripts\python.exe -B tools\smoke_interface.py` : réussi, construction **0,115 s**, recherche pendant la première indexation puis fonctions historiques vérifiées.
- `.\.build-venv\Scripts\python.exe -B tools\check_interactions.py` et variante `--script` : parcours réels réussis par import et par lancement du point d'entrée depuis un autre dossier. Contrôles historiques conservés ; thèmes, contrastes natifs, activation clavier, styles Classique inchangés, identité de sélection conservée, boutons Synthèse GET/IDPK/PUTE et statuts séparés. Grande vue de 30 000 lignes supplémentaires : temporisation, recherches dépassées, annulation, Entrée, flèche ↑ initiale, réactivité et fermeture sans callback périmé.
- `.\.build-venv\Scripts\python.exe -B tools\check_real_case.py` : cache existant des **66 fichiers en lecture seule**, interface **0,2314 s**, ID **48,3 ms**, référence **58,9 ms**, absence **27,1 ms** affichage compris. Ticket Local R2 ligne 65169, GET 11:00:36.9, PUTE 11:00:52.3 et surlignages exacts ; cache et traces inchangés.
- Deux exécutions de `tools\benchmark_modernisation.py --store ... --output ...` : modules avant/après, sources temporaires, bases neuves distinctes, indexations successives. **429,107 s → 425,304 s** pour la première construction ; aucune amélioration significative revendiquée sur ce seul passage. Ajout de **219 octets à une trace de 7 958 806 octets : 11,708 s → 0,028 s**, ticket et événements vérifiés. Après : **4 lignes/254 octets reparsés**, préfixe intégral vérifié. Cache neuf **2 326 171 648 → 2 326 990 848 octets**, mémoire post-index **26,8 Mio** dans les deux versions. Détails, médianes de cinq recherches et limites dans `PERFORMANCE.md` et rapports `.benchmark/modernisation20261017/`.
- Scripts Python ponctuels sans écriture des traces : simulation CTk des échelles 100/150/200 % avant construction, contrôle des limites de l'écran, du champ de saisie et de tous les boutons. Fenêtres **1100×800**, **1650×960**, **1800×920** ; commandes accessibles et aucune exception de callback. Ce contrôle simule les facteurs de mise à l'échelle, sans changer les réglages Windows ou tester plusieurs écrans physiques.
- Captures Windows sur traces temporaires, inspectées : Classique clair, Synthèse sombre et Classique dans le cadre sombre. Elles ont permis de détecter et corriger des contrastes ainsi que le traitement incorrect de labels internes CTk. Les exports `PrintWindow` omettaient parfois des contrôles ; les aperçus ont été repris avec la fenêtre au premier plan et capture d'écran. Les paramètres réels des widgets et leurs contrastes sont également vérifiés par les interactions ; les captures seules ne valident pas tous les états.

Limites et problèmes encore ouverts :

- La vérification du préfixe reste proportionnelle à la taille du fichier. Un fichier ancien sans état de reprise sera réindexé complètement à sa prochaine modification avant de bénéficier des ajouts incrémentaux. Un début de fichier de moins de huit lignes physiques utilise aussi le traitement complet.
- L'ambiguïté du premier fichier sans BOM entre UTF-8 incomplet et Windows-1252 demeure ; un UTF-8 établi incomplet est réessayé. Sans date de création, la voie incrémentale conserve la date de secours initiale. Une réécriture préservant simultanément taille et mtime reste hors détection rapide.
- Aucun gain significatif de première indexation ou de taille du cache n'est revendiqué. `dbstat` est indisponible ; les tailles logiques des deux textes ne sont pas une ventilation physique des tables/index. Les chronos sont locaux et ne garantissent pas les mêmes résultats sur tous les disques.
- Les contrôles couvrent les scénarios énumérés, sans validation exhaustive de chaque état ni comparaison pixel par pixel de tous les thèmes. Aucune modification des traces fournies, du cache courant, des environnements virtuels, des copies historiques ou des artefacts de build ; aucun exécutable reconstruit.

## 2026-10-16 — Restauration des interactions et du panneau de résultats

Demande : retrouver les interactions et la présentation du logiciel de 2025, conserver les performances SQLite et garder une trace durable des corrections.

Causes identifiées :

- En lancement `python search_gui.py`, le constructeur utilisait les variables de `__main__` tandis que le contrôleur importait un second module `search_gui`. Le bouton IDs adjacents lisait donc un autre `current_file`. Les contrôles précédents par import ne reproduisaient pas cette situation.
- Le retour Référence → ID analysait le texte à l'écran pour chercher `ModBusId=`, alors que le rendu écrivait `ModBusId:`. Les données nécessaires existaient déjà dans l'index.
- Les voisins se plaçaient sur la première utilisation d'un ID ; une déduplication trop tardive pouvait aussi laisser moins de cinq suivants après passage à l'archive. Une absence de résultat conservait des informations de sélection périmées.
- Le rendu indexé avait perdu les styles du panneau historique. Les flèches GET/PUTE de la vue complète déplaçaient l'affichage sans mettre à jour le curseur utilisé au clic suivant.

Changements livrés :

- Sélection structurée appartenant à l'application (`selection_state.py`) ; les actions actives utilisent le chemin complet, le numéro original, la référence, l'ID et les événements, sans état global partagé entre modules ni lecture du texte affiché.
- Bouton IDs adjacents et retour ID → Référence → ID réparés. Le changement de mode conserve l'occurrence et la source sélectionnées. Les clics voisins et erreurs Robot transmettent le numéro de ligne.
- Voisinage ancré sur l'occurrence exacte, cinq IDs distincts par côté lorsque disponibles, complément des suivants dans le fichier lié du même robot après exclusion des doublons. Popup actualisée après navigation et vidée en cas d'absence.
- Résultats périmés écartés par génération de sélection et de vue. Recherches, diagnostics et lecture de trace restent en arrière-plan.
- Doubles-clics GET/PUTE conservés : vue complète du bon fichier et ligne exacte surlignée. Une modification avant rafraîchissement est détectée à l'ouverture ; une ligne déplacée unique est retrouvée, une ligne supprimée ou ambiguë demande de relancer la recherche. Les flèches et la recherche locale mettent aussi à jour le curseur.
- Styles historiques rétablis depuis le code de 2025 : date et occurrence centrées, fonds bleu/orange/vert/jaune, polices, marges, espacements et séparateurs. Provenance R1/R2 conservée.
- Règle permanente de journal dans `AGENTS.md`, ce fichier et outils reproductibles de contrôle d'interactions et du cas réel.

Vérifications effectivement exécutées :

- `python -m unittest discover -s tests -v` : les 41 tests existants passent. Après ajout de cinq régressions de service, `-q` : **46 tests passent**, dont ID réutilisé, cinq suivants après doublons entre archives du même robot, cinq précédents distincts, GET/PUTE de la seconde occurrence et fichier modifié/supprimé avant rafraîchissement.
- `python tools/check_interactions.py` et `python tools/check_interactions.py --script` : parcours Tkinter exécuté par import puis par la vraie commande `python search_gui.py`, depuis un dossier différent. Bouton IDs adjacents ; vrais clics Référence aller/retour ; vrais doubles-clics GET/PUTE avec contenu intégral, numéro original et surlignage vérifiés ; clic voisin ; clic erreur Robot ; ID réutilisé ; absence ; réponses de voisinage et d'ouverture volontairement retardées ; ligne déplacée/supprimée et fichier supprimé ; flèches GET/PUTE ; recherche locale ; comparaison IDPK ; tracking ; recherche avancée ; navigation ; recherches successives rapides et changement de vue. Le convoyage vérifie aussi le clic d'une carte jusqu'au détail avec les états affichés. Le cas sans aucune trace vérifie l'état vide et l'ouverture protégée. Aucune exception de callback Tkinter ignorée.
- `python tools/smoke_interface.py` : contrôle existant de l'interface et de la recherche pendant la première indexation temporaire réussi ; construction de l'interface environ 0,30 s.
- `python tools/check_real_case.py` : cache réel ouvert en lecture seule, **66 fichiers**, ID **683**, référence **A0061256AB**, **Local R2 / TracOMEGA.txt**, ticket ligne **65169** (numéro affichable commençant à 1). GET **11:00:36.9**, PUTE **11:00:52.3**. Les doubles-clics ouvrent le fichier complet et surlignent les lignes exactes. Taille et date du cache et des traces vérifiées inchangées. Plusieurs passages ont aussi utilisé une option expérimentale de capture, retirée après constat de son manque de fiabilité.
- Sur ce cache existant : lecture des métadonnées et vérification des fichiers **0,0046 s**, interface disponible **0,374 s**, ID **40,7 ms**, référence **29,4 ms**, référence absente **30,7 ms**, affichage compris. Autre passage après les corrections finales : interface **0,386 s**, ID **42,5 ms**, référence **61,8 ms**, absence **29,8 ms**. Pas de première indexation remesurée.
- Validation finale : **46 tests en 0,789 s**, compilation syntaxique des fichiers modifiés réussie, parcours complet `--script` réussi après rétablissement des états de boutons. Les essais par import et l'ancien contrôle d'indexation progressive ont aussi été exécutés avec succès.
- Présentation : le panneau de résultats a été inspecté sur une capture native partielle de la fenêtre (blocs de résultats visibles), comparé aux styles du code historique et ses paramètres principaux vérifiés par le contrôle Tkinter. L'export natif s'est ensuite montré instable (contrôles ou texte omis) et Windows n'a pas permis la capture au premier plan. Les exports incomplets et l'utilitaire expérimental ont été retirés. Aucune capture fiable de toute la fenêtre ni comparaison pixel par pixel avec une capture originale de 2025 n'est livrée.

Limites : ces scénarios couvrent les interactions énumérées ; ils ne prouvent pas tous les états possibles de tous les boutons. Pas de modification de schéma SQLite, de réindexation complète, de modification des traces fournies ni de reconstruction de l'exécutable. Les caches reconstruits et les modifications de traces des tests utilisent exclusivement des fichiers temporaires.

## 2026-10-16 — Historique consigné : index SQLite et icône

Demande antérieure : accélérer le chargement et les recherches et créer une icône adaptée.

Fonctionnalités livrées : `TraceStore`, SQLite persistant avec FTS5 trigram et index structurés, synchronisation incrémentale, transactions WAL, reconstruction d'un cache incompatible/corrompu, indexation progressive en arrière-plan, messages vers le thread Tkinter, élimination des anciennes recherches et cache mémoire limité aux trois traces ouvertes. Ressources loupe/robot SVG, PNG transparent et ICO sept tailles ; intégration Python/PyInstaller et fenêtres secondaires, sans rebuild de l'exécutable.

Vérifications antérieures consignées : 41 tests, contrôle d'interface par import et mesures documentées dans `PERFORMANCE.md`. Sur les 66 traces : première construction environ 6 min 31 s, cache environ 2,2 Go ; réouverture environ 0,03 s ; affichage des recherches testées 28–33 ms et mémoire résidente environ 47 Mo sans vue complète ouverte. Ces mesures sont historiques, elles n'ont pas été intégralement refaites pendant la restauration.

Régressions ensuite signalées : bouton IDs adjacents en lancement script, retour Référence → ID, rendu simplifié et état de sélection périmé. Le contrôle antérieur ne testait pas le véritable point d'entrée ni les clics jusqu'au contenu/surlignage. Ces défauts sont traités dans l'entrée de restauration ci-dessus ; la livraison initiale ne constituait pas une validation de chaque interaction.

## 2026-10-16 — Historique consigné : formats de traces actuels

Demande antérieure : adapter le logiciel aux traces locales récentes en conservant les fonctions historiques.

Fonctionnalités livrées : parseur commun `trace_parser.py`, UTF-8/Windows-1252, champs séparés par `<Cr>` ou séparateurs historiques, libellés accentués et anciens, reconnaissance GET/PUTE et IDPK, références exactes, dates des traces, diagnostics et suivi. Positions des registres MODBUS conservées, traces intactes. Consignes initiales `AGENTS.md` ajoutées.

Vérifications antérieures consignées : 24 tests de sources et de parseur sur extraits anciens/actuels, dont collisions de références et fichiers homonymes. Cette couverture de parsing ne validait pas tous les clics Tkinter.

## 2026-10-16 — Historique consigné : sources locales R1/R2

Demande antérieure : ajouter `traces/trace r1` et `traces/trace r2` en gardant `C:/Apoteka/traces`.

Fonctionnalités livrées : découverte centralisée dans `trace_sources.py`, chemins relatifs au script ou au `.exe`, dossiers absents tolérés, filtres TracOmega/TracMODBUS et exclusion `.sav`, identité par chemin complet, labels Local R1/Local R2/Apoteka, liens d'archives restant dans la source et actualisation des listes/index.

Vérifications antérieures consignées : 11 tests de sources, homonymes, filtres, liens, modification/suppression, répertoire de lancement et ressources Python/exécutable. Les traces locales restent externes ; aucun nouveau build à cette étape.

## 2025 — Application initiale (contexte fourni par l'auteur)

Projet réalisé par l'utilisateur en 2025. Les copies `PythonProject1/` et `old/` servent de référence historique et ne sont pas les points d'entrée à modifier. Aucun historique de tests daté de 2025 n'a été reconstitué.
