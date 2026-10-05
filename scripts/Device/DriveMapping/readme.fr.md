[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Device](../readme.fr.md) › **DriveMapping**

# DriveMapping

Mappe des bibliothèques de documents SharePoint Online / OneDrive sur des lettres de lecteur fixes via le redirecteur WebDAV — à utiliser comme script d'ouverture de session par utilisateur (application Win32 Intune ou tâche planifiée à l'ouverture de session), et non via [`menu.ps1`](../../../menu.ps1).

## Scripts

| Script | Description |
|--------|-------------|
| [`New-CloudDriveMapping.ps1`](New-CloudDriveMapping.ps1) ([docs](#new-clouddrivemappingps1)) | Mapper des bibliothèques de documents SharePoint Online / OneDrive sur des lettres de lecteur via WebDAV — essai à blanc sauf avec `-Apply` |

---

### New-CloudDriveMapping.ps1

Convertit chaque URL `https://` de bibliothèque de documents en sa forme UNC WebDAV (`\\<host>@SSL\DavWWWRoot\<path>`) et la mappe avec `net use`. S'exécute par défaut en mode essai à blanc — aucun lecteur n'est mappé sans `-Apply`.

Suppose que l'utilisateur connecté dispose déjà d'une session/SSO valide vers le tenant (de la même manière qu'un navigateur atteint SharePoint via WebDAV) — il ne s'agit pas d'une authentification Graph app-only. Nécessite le service Windows **WebClient** (redirecteur WebDAV) ; présent par défaut sur Windows 10/11, il requiert la fonctionnalité **Desktop Experience** sur Windows Server.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-MappingsCsv` | * | CSV avec les colonnes `DriveLetter`, `Url` et, en option, `Label` — une ligne par mappage |
| `-DriveLetter` | * | Mappage ponctuel unique : lettre de lecteur (à utiliser avec `-Url`) |
| `-Url` | * | Mappage ponctuel unique : URL de la bibliothèque de documents (à utiliser avec `-DriveLetter`) |
| `-Label` | Non | Nom convivial du mappage ponctuel unique |
| `-RemoveExisting` | Non | Supprime d'abord tout mappage existant sur la ou les lettres cibles (`net use <letter>: /delete /y`) |
| `-Persist` | Non | Rend le mappage persistant après redémarrage (`/persistent:yes`). Désactivé par défaut — un script d'ouverture de session remappe normalement à chaque connexion |
| `-Apply` | Non | Effectue réellement le mappage (par défaut : essai à blanc) |
| `-OutputPath` | Non | Dossier de journal (par défaut : `$env:TEMP` — les scripts d'ouverture de session s'exécutent généralement dans le contexte utilisateur, pas en administrateur) |

*Soit `-MappingsCsv`, soit `-DriveLetter` + `-Url` est obligatoire.

**Exemple de CSV de mappages**

```csv
DriveLetter,Url,Label
S,https://contoso.sharepoint.com/sites/Finance/Shared Documents,Finance Docs
O,https://contoso-my.sharepoint.com/personal/j_doe_contoso_com/Documents,My OneDrive
```

**Exemples**

```powershell
# Essai à blanc à partir d'un CSV de mappages
.\New-CloudDriveMapping.ps1 -MappingsCsv .\mappings.csv

# Appliquer les mappages du CSV, en supprimant d'abord tout mappage existant sur ces lettres
.\New-CloudDriveMapping.ps1 -MappingsCsv .\mappings.csv -RemoveExisting -Apply

# Mappage ponctuel unique
.\New-CloudDriveMapping.ps1 -DriveLetter Z -Url "https://contoso.sharepoint.com/sites/Finance/Shared Documents" -Apply
```

**Déploiement comme script d'ouverture de session**
- Empaquetez-le comme application Win32 Intune (ou stratégie de script PowerShell) exécutée dans le contexte **utilisateur**, déclenchée à l'ouverture de session, qui appelle `New-CloudDriveMapping.ps1 -MappingsCsv <path> -RemoveExisting -Apply`
- Livrez `mappings.csv` à côté du script (listes de mappages par client/par groupe)

**Remarques**
- Aucun identifiant n'est stocké ni demandé — la réussite du mappage dépend entièrement de la session tenant existante de l'utilisateur connecté
- Un fichier journal listant les chemins WebDAV résolus et l'état de chaque mappage est écrit après chaque exécution
