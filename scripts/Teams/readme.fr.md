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

Archiveur Teams : exporte les membres, les fichiers et les conversations des canaux listés, puis (en option) archive des équipes ou des canaux. Tout passe par Microsoft Graph. PowerShell 7+ requis, à exécuter en tant que Global Admin du client (ou avec des droits GDAP sur ce client).

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-Step10Action` | `interactive` (par défaut), `archive`, `undo` ou `skip` |
| `-Step10Only` | Exécuter uniquement l'étape 10 (archivage/désarchivage), de manière non interactive |
| `-ChannelAction` | `none` (par défaut), `archive` ou `undo` — mode rapide par canal |
| `-ChannelArchiveTag` | Texte de marqueur utilisé pour le repli par renommage (par défaut : `[ARCHIEF]`) |
| `-ChannelFallbackToRename` | Se rabattre sur un marqueur de renommage si l'appel à l'API Graph d'archivage/désarchivage échoue |
| `-DryRun` | Simulation — conserve toute la connexion et valide les étapes 6 à 9 en sondant les comptages, sans écrire d'exports ni modifier l'état d'archivage |
| `-WorksheetName` | Feuille contenant la liste des équipes. Par défaut : la première feuille du fichier Excel |
| `-TenantId` | ID du tenant client. Proposé par défaut dans l'assistant : le client GDAP issu de `Connect-Tenant` |
| `-ClientId` / `-CertificateThumbprint` | App-only au lieu du mode délégué, avec cette inscription d'application et ce certificat |
| `-AppOnly` | App-only avec l'inscription d'application du tenant dans `graph.appid.json` |
| `-PnPClientId` | Application PnP pour le repli administrateur de site. Par défaut : l'entrée du tenant dans `pnp.appid.json` |

**Exemples**

```powershell
# Délégué (par défaut) : code d'appareil selon load.config.ps1, d'abord un dry run
.\Invoke-TeamsArchive.ps1 -DryRun

# Archiver uniquement les canaux listés, en app-only
.\Invoke-TeamsArchive.ps1 -Step10Only -ChannelAction archive -AppOnly -TenantId <tenant-guid>
```

**Remarques**

Le fichier Excel doit contenir les colonnes `TeamName`, `ChannelName` et `Archive` (les lignes avec `Archive` = `Archive` sont traitées). Rien dans le script n'est lié à un client : l'ID du tenant, l'URL SharePoint, le fichier Excel et le dossier d'archive sont demandés par l'assistant de configuration (par défaut `C:\Temp\Teams_Channels.xlsx` et `C:\Temp\Teams_Archive`).

Connexion (v9.0) :
- Passe par [`Connect-M365.ps1`](../Startup/readme.fr.md#connect-m365ps1). **Délégué par défaut** : vous vous connectez en tant qu'administrateur, avec un code d'appareil quand `useDeviceCodeAuth` est activé dans `load.config.ps1` (et, sans `load.config.ps1`, avec un code d'appareil comme avant). Le redémarrage dans une session propre transmet ces réglages et le client GDAP à la nouvelle session.
- **Plus d'inscription d'application temporaire.** Les versions précédentes en créaient une à chaque exécution, y donnaient le consentement puis la supprimaient ; les étendues déléguées sont désormais demandées à la connexion sur l'application Microsoft Graph Command Line Tools.
- **App-only** avec `-ClientId` + `-CertificateThumbprint`, ou `-AppOnly`. L'application a alors besoin des autorisations d'application `Group.Read.All`, `Sites.Read.All`, `TeamMember.Read.All`, `ChannelMessage.Read.All` (une API protégée que Microsoft doit approuver) et `TeamSettings.ReadWrite.All` / `ChannelSettings.ReadWrite.All` pour l'archivage.
- La session Graph est fermée à la fin ; elle appartient à la session redémarrée.

Ce qui passe par où :
- **Graph** pour tout ce qui le permet : trouver les équipes (`/groups`, filtré sur `resourceProvisioningOptions` = `Team`), les membres (`/teams/{id}/members`), les canaux (`/teams/{id}/channels`), l'emplacement des fichiers du canal (`filesFolder`), la liste et le téléchargement des fichiers (`/drives/{id}/items/{id}/children` et `/content`), les conversations (`/messages`, `/replies`) et l'archivage (`/teams/{id}/archive`, `/channels/{id}/archive`). Le module MicrosoftTeams n'est plus utilisé ni installé.
- **PnP** uniquement pour le repli administrateur de site : quand les fichiers d'un canal répondent « access denied » à une connexion déléguée, l'administrateur connecté devient une fois par site administrateur de la collection de sites (`Set-PnPSite -Owners`) — Graph n'a pas d'API pour cela. Il utilise l'application PnP de `pnp.appid.json` (ou `-PnPClientId`) ; sans application, le script indique comment le faire à la main. L'app-only n'en a jamais besoin.

Comportement :
- Résout l'emplacement des fichiers des canaux via Graph `filesFolder` pour tous les types de canaux (standard/privé/partagé), avec la liste des canaux mise en cache par équipe.
- Normalise les valeurs TeamName/ChannelName provenant d'Excel (trim) et compare les noms de canaux normalisés (trim + réduction des espaces + minuscules) pour éviter les faux cas « Kanaal niet gevonden ».
- Traite un NotFound SharePoint pendant l'export comme un saut contrôlé.
- Télécharge les fichiers avec des nouvelles tentatives par fichier et une vérification du nombre après téléchargement. L'arborescence du canal est désormais conservée sous `Files` (elle était auparavant aplatie : deux fichiers de même nom dans des sous-dossiers différents s'écrasaient et le comptage échouait).
- Enregistre la sortie par canal : `Teams > Team > Channel > Files, Chat, Members`. `members.csv` garde les colonnes `Name`, `User`, `Role`.
- N'archive pas les équipes par défaut ; l'archivage exige une confirmation explicite à l'étape 10, qui prend aussi en charge `unarchive`, un mode rapide non interactif (`-Step10Only -Step10Action undo|archive|skip`) et un véritable archivage/désarchivage par canal (`-Step10Only -ChannelAction archive|undo`), avec un repli facultatif sur un marqueur de renommage (`-ChannelFallbackToRename`, `-ChannelArchiveTag`).
- Dans Microsoft Teams, archiver/désarchiver une équipe est une action au niveau de l'équipe.
- Le dry-run conserve toute la connexion, valide les étapes 6 à 9 en sondant les comptages sans écrire sur disque, utilise ces comptages dans le rapport de l'étape 11 et ne fait que simuler l'étape 10 (`[DRYRUN]`).
- Le redémarrage dans une session propre après le nettoyage des modules transmet tous les paramètres (y compris `-DryRun`), renvoie le code de sortie de l'exécution redémarrée et ne laisse plus sa variable de marqueur dans la session appelante (une deuxième exécution dans la même session sautait auparavant le nettoyage).
