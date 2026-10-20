# Chargement et recherches

## Harmonisation de l'interface — 19 octobre 2026

`tools/check_ui_latency.py` sur bureau Windows privé, mêmes deux traces temporaires, avec contrôle de conservation des widgets, sélection/copie, bornage des panneaux, absence de SQL de recherche/comparaison lors des changements de thème et rafraîchissement après révision de source. Il mesure le callback puis le traitement des événements (avec une attente de drainage d'environ 25 ms). Ce n'est ni une première indexation ni une mesure sur les bases de production.

Médianes en ms, **callback / événements traités** :

| Transition | Avant | Dernier passage après |
|---|---:|---:|
| Synthèse → Classique | 13,04 / 48,56 | 14,01 / 50,41 |
| Classique → Synthèse | 30,68 / 105,62 | 36,69 / 122,34 |
| Recherche → Suivi | 17,90 / 286,17 | 19,54 / 250,16 |
| Suivi → Convoyage | 13,14 / 79,74 | 17,39 / 76,00 |
| Convoyage → Recherche | 48,80 / 488,52 | 34,18 / 476,60 |

Rapports conservés : `.benchmark/ui_unified_before_20261019.json`, `ui_unified_after_20261019.json`, `ui_unified_after_serial_20261019.json`, `ui_unified_after_final_20261019.json`. Le premier passage après a coïncidé avec un autre contrôle graphique privé et présente des durées nettement plus élevées (retour Recherche : 1 015,97 ms) ; il n'est pas utilisé comme comparaison isolée. Le passage isolé intermédiaire mesure ce retour à 445,58 ms et Classique → Synthèse à 97,24 ms. Le dernier passage suit le traitement des états/couleurs identiques des boutons ; tous les rapports restent disponibles, sans sélection du meilleur.

Les contrôles de réutilisation et de SQL réussissent. Les temps restent variables avec le rendu natif et la vérification périodique ; le dernier passage Classique → Synthèse prend 16,72 ms de plus que la référence. L'harmonisation n'est pas présentée comme une accélération générale. Mémoire RSS et tailles des bases non remesurées : le stockage n'a pas changé. Les listes de 10 000 anomalies sont vérifiées avec moins de 16 cartes instanciées aux dimensions testées.

## Voisins et comparaisons exactes — 19 octobre 2026

Mesure des vrais clics sur un bureau Windows privé : quatre fichiers OMEGA générés dans un dossier temporaire, 1 200 tickets, 300 identités sur deux robots et deux périodes. Douze clics R1 suivant le même parcours 1251→1252→1251→1250, répété trois fois. Temps complet jusqu'à la sélection, aux voisins et à la comparaison Synthèse disponibles ; compteur de toutes les instructions SQLite tracées, transactions comprises. Sources originales et bases permanentes non utilisées.

| Mesure | Avant | Après |
| --- | ---: | ---: |
| Navigation, médiane | 1 334,0 ms | 816,6 ms |
| Instructions SQLite par navigation, médiane | 270 | 18 |
| Première ouverture des voisins | 375,0 ms | 428,0 ms |
| Instructions SQLite à cette ouverture | 243 | 22 |
| Mémoire résidente du processus à la fin | 68,93 Mio | 70,96 Mio |
| Cache SQLite, fichier principal | 1 998 848 octets | 2 093 056 octets |
| Navigations conservant les deux widgets de tableaux | 0 / 12 | 12 / 12 |

Dernier gain médian de navigation : **38,8 %**, avec **93,3 %** d'instructions SQLite en moins ; environ **2,02 Mio** de mémoire résidente et **92 Kio** de cache supplémentaires. Les index sont additifs, préparés après disponibilité du cache. Les données de voisinage et comparaisons exactes sont mutualisées par révision publiée, avec caches bornés ; les tableaux et leurs styles sont mis à jour en place. L'historique et les traces restent inchangés.

Rapports : `.benchmark/adjacent_before_20261019.json` et `.benchmark/adjacent_after_20261019.json`. L'intermédiaire `adjacent_after_initial_20261019.json` mesurait 736,0 ms ; il n'est pas retenu à la place du dernier passage. La première ouverture n'est pas accélérée dans le dernier échantillon et la navigation reste entre 766,6 et 1 223,8 ms : le rendu de la sélection principale demeure coûteux. Aucune promesse de navigation instantanée, de résultat sur le gros corpus réel ou de temps disque froid. Une mesure de mémoire résidente finale ne constitue pas un pic mémoire ; les tailles WAL/historique ne sont pas remesurées ici. Les nouvelles comparaisons respectent Référence/SN/CodeLot et peuvent traverser les robots, contrairement à l'ancienne comparaison par référence : les temps comparent les parcours utilisateurs, pas deux règles métier identiques.

## Fluidité de l'interface — 19 octobre 2026

Mesure sur bureau Windows privé, deux petits fichiers OMEGA temporaires issus de `current_r2.txt`, vingt lots simulés à partir d'un événement indexé, trois échantillons par transition. Les temps ci-dessous sont les médianes ; « événements » comprend `update()` puis une courte phase de drainage d'au moins 25 ms. Il ne s'agit ni d'une mesure de fréquence d'affichage ni d'un temps d'indexation.

| Bascule | Callback avant | Callback après | Avec événements avant | Avec événements après |
| --- | ---: | ---: | ---: | ---: |
| Synthèse → Classique | 125,4 ms | 12,9 ms | 178,0 ms | 49,1 ms |
| Classique → Synthèse | 518,6 ms | 26,7 ms | 1 522,3 ms | 169,1 ms |

Les cartes et leur texte sélectionnable sont conservés, sans nouvelle recherche SQL pour ces bascules. Les changements de fonction réutilisent trois panneaux au maximum. Les callbacks Recherche→Suivi, Suivi→Convoyage et Convoyage→Recherche passent respectivement de 90,7/57,8/187,9 ms à 27,6/21,4/51,9 ms. Le drainage après navigation reste variable : environ 366/105/678 ms dans le dernier contrôle privé, qui conserve désormais le résultat avec ses vingt lots au retour, alors que la version précédente revenait à une recherche vide. Aucun gain équivalent n'est revendiqué sur tout le rafraîchissement visuel de ces changements de fonction.

Causes corrigées : destruction/reconstruction de Synthèse à chaque mode, reconstruction des panneaux, parcours global du thème sur chaque Map, reconfiguration identique des polices/couleurs/wraplength des textes. Le thème recolore désormais les champs existants. Les workers sont invalidés lors d'une navigation ; une révision d'index différente relance la recherche au retour. Le contrôle vérifie cette actualisation par ajout réel dans une trace temporaire, ainsi que la conservation des widgets, de la sélection copiée et l'absence de requêtes redondantes.

Rapports : `.benchmark/ui_latency_before_20261019.json` et `.benchmark/ui_latency_validated_20261019.json`. Les rapports intermédiaires `after_initial`, `after`, `final`, `final2`, `final3` et `delivered` documentent les essais successifs de réutilisation, ciblage du thème et placement des panneaux. La variation des temps Windows interdit de présenter le meilleur échantillon comme une garantie. Mémoire/tailles des bases non remesurées pour cette intervention d'interface ; rétention volontaire limitée aux trois panneaux et à la Synthèse courante. Aucun benchmark du corpus réel, aucune modification des traces ou bases permanentes, aucun exécutable reconstruit.

Depuis le 19 octobre 2026, OMEGA dispose d'un historique SQLite permanent compressé, séparé du cache de recherche par sélection. Les générations disparues restent indexables ; MODBUS conserve la synchronisation des fichiers présents. Les mesures du 17 octobre et antérieures précèdent cette conservation cumulative. La comparaison du 19 octobre ci-dessous porte sur quatre fichiers courants ; aucun benchmark global des 105 traces n'est revendiqué.

L'application utilise un index SQLite local. Le cache déjà publié est disponible avant la préparation de l'historique et la vérification des sources. OMEGA est vérifié par empreinte brute une fois au démarrage, sans réanalyse des contenus reconnus ; les passages suivants contrôlent taille/date/identité avant tout nouveau traitement. L'indexation initiale et la reprise d'un cache incomplet restent progressives : seuls les fichiers non publiés ou modifiés sont traités. Un fichier interrompu avant validation de sa transaction est retraité seul.

L'indexation, les recherches et les analyses tournent en arrière-plan. Les informations GET/PUTE, IDPK et les derniers emplacements par lot sont déjà indexés : afficher un résultat ne parcourt plus les archives. Le texte complet n'est chargé qu'à l'ouverture d'une trace, avec une limite de trois fichiers en mémoire.

## Mesures du 19 octobre 2026 — démarrage et FTS par lots

Comparaison avec la version précédente disposant déjà de l'historique OMEGA : **4 fichiers complets courants R1/R2 OMEGA/MODBUS, 11 661 460 octets**. Chaque exécution utilise ses propres copies temporaires et bases neuves. Copie des sources, vérification de leurs empreintes et mesures de stockage exclues du temps d'indexation ; cache Windows non purgé. Une exécution par version, sans prétendre mesurer un disque froid ou toute la collection locale.

| Opération | Avant | Après |
| --- | ---: | ---: |
| Première indexation, archivage inclus | 13,614 s | 5,273 s |
| Réouverture du service et métadonnées en cache | 24,2 ms | 26,7 ms |
| Vérification OMEGA à la réouverture | 1,028 s | 54,1 ms |
| Passage suivant sans changement | 24,3 ms | 19,8 ms |
| Ajout dans OMEGA R2, génération conservée | 1,209 s | 1,034 s |
| Cache de recherche neuf | 43 081 728 octets | 42 614 784 octets |
| Archive permanente neuve | 1 134 592 octets | 1 134 592 octets |
| Mémoire résidente après première indexation | 27,7 Mio | 28,7 Mio |
| Mémoire résidente finale | 28,8 Mio | 29,4 Mio |

La première indexation est réduite de **61,3 %** sur ce corpus ; la vérification au redémarrage est environ **19 fois plus rapide**. Ces facteurs ne sont pas extrapolés aux 105 fichiers. L'ajout conserve la voie d'indexation incrémentale ; la compression et le calcul canonique du fichier modifié restent proportionnels à son contenu.

Les empreintes des résultats et des événements GET/PUTE/IDPK des quatre recherches comparées sont identiques. Médianes de cinq passages, **service seul sans rendu Tk** : ID présent 2,10 → 2,20 ms et préparation des détails 10,64 → 10,91 ms ; référence présente 5,70 → 5,43 ms et détails 10,40 → 10,72 ms ; ID absent 1,91 → 1,99 ms ; référence absente 4,04 → 4,21 ms. Aucun gain significatif sur les recherches elles-mêmes n'est revendiqué : le changement vise les traitements répétés et l'alimentation de l'index.

Contrôle d'interface sur bureau privé avec préparation volontairement bloquée : cache disponible **0,623 s après le début de l'ouverture du service**, **0,978 s construction de l'interface comprise**. Recherche par bouton et vrais doubles-clics GET/PUTE exécutés avant libération de la préparation ; texte complet, lignes originales et surlignages vérifiés, sélection stable après la vérification. Le premier essai de ce contrôle incluait toute la construction Tk dans le seuil de disponibilité après choix et avait dépassé une seconde en parallèle d'un autre contrôle graphique ; les deux durées sont désormais distinguées et le contrôle chronométré est exécuté seul.

Le cache réel incomplet du Bureau et son historique ont aussi été copiés avec l'API de sauvegarde SQLite, **originaux ouverts uniquement en lecture seule** : 8 fichiers publiés dont 5 OMEGA. Sur cette copie, service et métadonnées disponibles en **19,1 ms**, recherche et détails en **13,6 ms**, avant toute préparation d'historique. Ce contrôle n'effectue ni synchronisation réelle ni indexation des 58 générations OMEGA restantes.

Rapports locaux : `.benchmark/omega_performance_20261019/before.json`, `after.json`, `cache_readiness.json`. Modules de référence conservés dans le sous-dossier `baseline/`. Commandes reproductibles (choisir de nouveaux rapports absents) :

```powershell
py -3.13 -B tools/benchmark_omega_startup.py --module-dir .benchmark/omega_performance_20261019/baseline --output .benchmark/omega_performance_20261019/before.json
py -3.13 -B tools/benchmark_omega_startup.py --output .benchmark/omega_performance_20261019/after.json
```

Les bases de mesure sont temporaires ; le script vérifie que les copies correspondent aux originaux, puis que tailles/dates de ceux-ci sont inchangées. Les temps du service ne sont pas des temps complets depuis le lancement de Python. Aucune archive réelle n'a été reconstruite ou effacée pour ces mesures.

## Mesures du 17 octobre 2026 — modernisation

Comparaison exécutée sur les **mêmes 66 traces, 612 258 393 octets et 7 299 737 lignes**. Chaque version utilise une base SQLite neuve distincte et des copies temporaires des sources. Les deux indexations complètes sont exécutées successivement. Le cache de fichiers de Windows n'est pas purgé : il s'agit de premières constructions SQLite, pas de mesures sur disque froid. La copie des traces et les analyses de stockage sont exclues des chronos d'indexation.

| Opération | Version avant cette intervention | Version modernisée |
| --- | ---: | ---: |
| Première indexation complète | 429,107 s | 425,304 s |
| Réouverture du service | 2,16 ms | 1,99 ms |
| Vérification des fichiers inchangés | 19,01 ms | 18,79 ms |
| ID présent, recherche + détails, médiane de 5 | 0,97 ms | 0,86 ms |
| Référence présente, recherche + détails, médiane de 5 | 12,95 ms | 12,65 ms |
| ID absent, médiane de 5 | 0,45 ms | 0,44 ms |
| Référence absente, médiane de 5 | 3,44 ms | 2,99 ms |
| Synchronisation après ajout de 219 octets à une trace de 7 958 806 octets | 11,708 s | 0,028 s |
| Cache neuf après première indexation | 2 326 171 648 octets | 2 326 990 848 octets |
| Mémoire résidente après première indexation, sans interface | 26,8 Mio | 26,8 Mio |
| Mémoire finale du benchmark, sans interface | 29,1 Mio | 29,6 Mio |

Les temps de recherche du tableau concernent **le service et la préparation des détails**, sans rendu Tkinter. Le contrôle réel du cache existant a également été exécuté : interface disponible en **0,231 s**, ID **48,3 ms**, référence **58,9 ms**, absence **27,1 ms**, affichage compris. Ce passage ouvre le cache en lecture seule ; il ne mesure pas une première indexation. Les doubles-clics GET/PUTE vérifient le chemin, le texte complet et la ligne surlignée.

L'amélioration importante mesurée concerne l'actualisation d'une trace qui grossit. Pour cet ajout, **254 octets et 4 lignes** sont reparsés, dont la dernière ligne précédente. Les **7 958 806 octets du préfixe restent lus et vérifiés par SHA-256** ; ce contrôle n'est donc pas proportionnel au seul ajout. Le ticket ajouté et ses événements GET/PUTE ont été vérifiés dans les deux versions. Ce résultat porte sur ce scénario précis, pas sur tous les volumes ou disques.

La première construction varie de moins de 1 % : aucun gain significatif n'est revendiqué sur cette seule comparaison. La lecture/copie/détection d'encodage combinée représente **0,490 s** ; la phase parsing et écritures représente **424,229 s**. Les mesures regroupent parsing, SQL et FTS ; elles ne permettent pas d'attribuer cette durée à chacun séparément.

Le stockage n'a pas été restructuré, conformément au choix de préserver le cache. L'état de reprise ajoute environ **0,78 Mio** sur les bases neuves comparées. Après l'ajout de test, le texte original et sa copie normalisée représentent chacun **602 849 312 octets logiques UTF-8**. Ce ne sont pas leurs tailles physiques SQLite. L'extension `dbstat` étant absente de cette compilation, la ventilation physique par table/index reste indisponible. Aucun index n'a été supprimé.

Rapports détaillés locaux : `.benchmark/modernisation20261017/baseline.json` et `after.json`. Le module antérieur utilisé est sauvegardé dans `.benchmark/modernisation_20261017_baseline/trace_store.py`. Les commandes effectivement exécutées sont :

```powershell
.\.build-venv\Scripts\python.exe -B tools\benchmark_modernisation.py --store .benchmark\modernisation_20261017_baseline\trace_store.py --output .benchmark\modernisation20261017\baseline.json
.\.build-venv\Scripts\python.exe -B tools\benchmark_modernisation.py --store trace_store.py --output .benchmark\modernisation20261017\after.json
```

Pour rejouer une mesure, choisir un nouveau chemin de rapport : l'outil refuse une base ou un rapport déjà présent. Les traces et le cache courant ne sont jamais ouverts en écriture par ce benchmark. Les manifests avant/après confirment que les sources fournies sont restées inchangées.

## Actualisation et affichage

L'état `file_resume` est une extension additive du schéma version 1. Un cache existant reste utilisable sans reconstruction : un fichier sans état de reprise est réindexé complètement à sa prochaine modification, puis bénéficie de la voie incrémentale. Les huit premières lignes physiques doivent être stabilisées. Une réécriture, une troncature, une rotation, un état invalide ou un changement d'encodage déclenche le traitement complet du seul fichier concerné. Une écriture UTF-8 encore coupée est réessayée sans remplacer l'instantané déjà publié.

La recherche locale attend **200 ms** après la dernière frappe, puis travaille par tranches courtes sur le thread principal. Les requêtes dépassées et les callbacks des fenêtres détruites sont annulés. Les positions GET/PUTE sont préparées au chargement et les flèches utilisent une recherche binaire. Le contrôle d'interactions couvre **30 000 lignes supplémentaires**, les recherches rapides, l'annulation en cours, la réactivité de la boucle Tk et la fermeture.

Classique reste le rendu initial et conserve ses styles historiques, même dans le cadre sombre. Synthèse utilise les mêmes données de sélection et donne accès aux lignes GET/IDPK/PUTE exactes. Le choix de thème et de rendu reste propre à la session.

## Mesures historiques — index initial

Mesures sur les 66 traces fournies, environ 617 Mo, sous Windows avec Python 3.13. Les temps de recherche incluent la préparation des détails du premier résultat ; ils excluent la cadence de 30 ms de la file de messages de l'interface.

| Opération | Avant | Avec l'index prêt |
| --- | ---: | ---: |
| Chargement / réouverture avec vérification des fichiers | ~7 s | ~0,03 s |
| ID présent, résultat et détails | Recherche puis parcours des archives pour l'affichage | ~2 ms |
| Référence présente, résultat et détails | Recherche puis parcours des archives pour l'affichage | ~14 ms |
| ID absent | ~2,65 s | ~1,5 ms |
| Référence absente | ~2,30 s | ~3,5 ms |
| Construction de l'interface | — | ~0,31 s |

La **première construction** du cache a demandé environ **6 min 31 s** pendant cette intervention. Elle ne doit pas être confondue avec une réouverture : elle enregistre aussi le texte intégral pour les recherches futures. Le cache des 66 fichiers occupe environ **2,2 Go**. Le processus de benchmark utilise environ **20 Mo de mémoire résidente** après les recherches, sans fenêtre de trace complète ouverte. Les performances dépendent du disque et du contenu des fichiers.

Le cache a été préparé pour les traces locales présentes lors de cette modification. Il est stocké sous `%LOCALAPPDATA%\TracOmega\cache\<identifiant-du-projet>`. Une modification des traces déclenche une actualisation en arrière-plan, vérifiée toutes les deux secondes. Les traces d'origine restent intactes. Un cache incompatible ou endommagé est recréé automatiquement.

Le contrôle de l'application complète avec ce cache a mesuré **0,34 s** avant disponibilité des 66 fichiers, puis **28 à 33 ms** de la demande jusqu'à l'affichage pour les recherches testées. Sa mémoire résidente était d'environ **47 Mo**, sans fenêtre de trace complète ouverte.

## Vérifications

Après la restauration des interactions du 16/10/2026, les contrôles en lecture seule du cache des 66 traces donnent environ **0,37–0,39 s** avant disponibilité de l'interface et **29–62 ms** jusqu'à l'affichage pour l'ID 683, sa référence A0061256AB et une référence absente. Aucun changement de schéma ni réindexation complète n'est nécessaire. Les mesures détaillées et leurs limites sont consignées dans `CHANGELOG.md`.

```powershell
.\.build-venv\Scripts\python.exe -m unittest discover -s tests -v
.\.build-venv\Scripts\python.exe tools\smoke_interface.py
.\.build-venv\Scripts\python.exe tools\check_interactions.py
.\.build-venv\Scripts\python.exe tools\check_interactions.py --script
.\.build-venv\Scripts\python.exe tools\check_real_case.py
.\.build-venv\Scripts\python.exe tools\benchmark_traces.py --application-cache
```

Pour mesurer une première construction indépendante, utiliser le benchmark sans `--application-cache` : il crée un cache de mesure dans `.benchmark/`. Il faut prévoir l'espace disque et le temps indiqués ci-dessus.

## Icône

Les fichiers `assets/tracomega.svg`, `assets/tracomega.png` et `assets/tracomega.ico` représentent une loupe et un robot, en bleu foncé et turquoise. L'ICO contient les tailles 16, 24, 32, 48, 64, 128 et 256 pixels. Les exports sont reproductibles avec `tools/create_icon.ps1`, sans dépendance Python supplémentaire.

L'icône est appliquée aux fenêtres et configurée dans le fichier PyInstaller. L'exécutable existant n'a pas été reconstruit.
