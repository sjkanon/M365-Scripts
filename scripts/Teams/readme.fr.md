[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **Teams**

# Teams

Outillage d'export et d'archivage Microsoft Teams / SharePoint.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Invoke-TeamsArchive.ps1`](Invoke-TeamsArchive.ps1) ([docs](#invoke-teamsarchiveps1)) | Exécute un flux d'export et d'archivage Teams/SharePoint pour une liste d'équipes et de canaux lue dans Excel |

---

### Invoke-TeamsArchive.ps1

Archiveur Teams avec un flux d'export Graph, Teams et SharePoint. PowerShell 7+ requis, à exécuter en tant que Global Admin.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-Step10Action` | `interactive` (par défaut), `archive`, `undo` ou `skip` |
| `-Step10Only` | Exécuter uniquement l'étape 10 (archivage/désarchivage), de manière non interactive |
| `-ChannelAction` | `none` (par défaut), `archive` ou `undo` — mode rapide par canal |
| `-ChannelArchiveTag` | Texte de marqueur utilisé pour le repli par renommage (par défaut : `[ARCHIEF]`) |
| `-ChannelFallbackToRename` | Se rabattre sur un marqueur de renommage si l'appel à l'API Graph d'archivage/désarchivage échoue |
| `-DryRun` | Simulation — conserve toute l'authentification/l'amorçage et valide les étapes 6 à 9 en sondant les comptages, sans écrire d'exports ni modifier l'état d'archivage |
| `-WorksheetName` | Feuille contenant la liste des équipes. Par défaut : la première feuille du fichier Excel |

Le fichier Excel doit contenir les colonnes `TeamName`, `ChannelName` et `Archive` (les lignes avec `Archive` = `Archive` sont traitées). Rien dans le script n'est lié à un client : l'ID du tenant, l'URL SharePoint, le fichier Excel et le dossier d'archive sont demandés par l'assistant de configuration (par défaut `C:\Temp\Teams_Channels.xlsx` et `C:\Temp\Teams_Archive`).

Comportement actuel (v8.19) :
- Crée une inscription d'application Entra temporaire unique pour l'exécution.
- N'accorde que les autorisations déléguées de configuration nécessaires pendant l'amorçage.
- Applique le consentement délégué Graph/SharePoint à cette application temporaire.
- Supprime l'application temporaire et le service principal lors du nettoyage (et en cas d'échec d'une étape clé de la configuration).
- Enregistre un hook de nettoyage à la sortie afin que l'application temporaire soit aussi supprimée à la fermeture de PowerShell/Ctrl+C.
- Vérifie d'abord l'accès aux dossiers Teams/SharePoint pendant l'export des fichiers, et n'accorde des droits Graph plus élevés que si l'accès est refusé.
- Résout l'emplacement des fichiers des canaux via Graph filesFolder pour tous les types de canaux (standard/privé/partagé), avec mise en cache des canaux et recherche de repli.
- Normalise les valeurs TeamName/ChannelName provenant d'Excel (trim) pour éviter les échecs de recherche dus à des espaces en fin de chaîne.
- Réutilise le même résolveur de canaux mis en cache, avec repli Graph, dans l'export des conversations, ce qui améliore la cohérence de la détection des canaux en dry-run comme en exécution normale.
- Applique une correspondance normalisée des noms de canaux (trim + réduction des espaces + minuscules) dans la recherche en cache et la recherche de repli Graph, afin de réduire les faux cas « Kanaal niet gevonden ».
- Traite un NotFound SharePoint pendant l'export comme un saut contrôlé plutôt que comme des échecs bloquants bruyants.
- Télécharge les fichiers avec des nouvelles tentatives par fichier, un repli par reconnexion et une validation du nombre après téléchargement pour garantir l'exhaustivité.
- Enregistre la sortie selon une structure par canal : `Teams > Team > Channel > Files, Chat, Members`.
- N'archive pas les équipes par défaut ; l'archivage exige désormais une confirmation explicite pendant l'étape 10.
- L'étape 10 permet aussi d'annuler l'archivage (`unarchive`), avec logique de nouvelle tentative.
- L'étape 10 prend en charge un mode rapide non interactif : `-Step10Only -Step10Action undo|archive|skip`.
- Important : dans Microsoft Teams, l'archivage/le désarchivage est une action au niveau de l'équipe, pas du canal.
- L'étape 10 prend en charge un véritable archivage/désarchivage par canal via Microsoft Graph (`/channels/{id}/archive|unarchive`).
- Mode rapide par canal : `-Step10Only -ChannelAction archive|undo`.
- Repli facultatif sur un marqueur de renommage en cas d'échec de l'API : `-ChannelFallbackToRename` (marqueur via `-ChannelArchiveTag`).
- Le mode dry-run conserve toute l'authentification/l'amorçage et valide les étapes 6 à 9 en sondant l'existence et les comptages dans Teams/SharePoint/Graph, sans écrire sur disque les exports de membres/conversations/fichiers.
- En dry-run, le rapport de l'étape 11 utilise les comptages sondés (fichiers/messages détectés) au lieu des fichiers exportés localement.
- Les modifications d'archivage/désarchivage de l'étape 10 restent simulées, avec une sortie `[DRYRUN]`.
- Le redémarrage dans une session propre après le nettoyage des modules transmet tous les paramètres (y compris `-DryRun`) et renvoie le code de sortie de l'exécution redémarrée.
