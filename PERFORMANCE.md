# Chargement et recherches

L'application utilise un index SQLite local. Le premier passage lit et indexe les traces ; les lancements suivants vérifient leur taille et leur date de modification, puis ne traitent que les changements. Les fichiers courants sont prioritaires. Pendant cette première étape, les recherches fonctionnent sur les fichiers déjà indexés et affichent la mention « résultats partiels ».

L'indexation, les recherches et les analyses tournent en arrière-plan. Les informations GET/PUTE, IDPK et les derniers emplacements par lot sont déjà indexés : afficher un résultat ne parcourt plus les archives. Le texte complet n'est chargé qu'à l'ouverture d'une trace, avec une limite de trois fichiers en mémoire.

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
