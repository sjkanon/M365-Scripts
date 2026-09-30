[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **Azure**

# Scripts d'infrastructure Azure

Scripts de gestion directe des ressources Azure IaaS (et non du tenant M365) — distincts de toutes les autres catégories de ce dépôt, qui ciblent Microsoft 365 / Entra ID via Graph ou Exchange Online. Nécessite le module PowerShell `Az` et une session `Connect-AzAccount` authentifiée. Non intégré à [`menu.ps1`](../../menu.ps1) — à exécuter directement sur l'abonnement cible.

---

## Dossiers

| Dossier | Description |
|--------|-------------|
| [`VM/`](VM/readme.fr.md) | Convertir le type de contrôleur de disque d'une VM Azure entre SCSI et NVMe |

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Search-AADDSUserActivity.ps1`](Search-AADDSUserActivity.ps1) ([docs](#search-aaddsuseractivityps1)) | Rechercher un utilisateur unique dans toutes les tables d'audit Azure AD Domain Services de Log Analytics, en une seule requête |

---

### Search-AADDSUserActivity.ps1

Recherche un nom d'utilisateur donné dans les tables de diagnostic AAD DS d'un espace de travail Log Analytics, en une seule requête `union`, pour que vous n'ayez pas à savoir à l'avance dans quelle table (ou colonne) un événement a atterri. Nécessite des paramètres de diagnostic sur le domaine managé AAD DS qui envoient les journaux vers l'espace de travail cible.

**Tables interrogées par défaut**
- `AADDomainServicesAccountManagement`
- `AADDomainServicesAccountLogon`
- `AADDomainServicesLogonLogoff`
- `AADDomainServicesDirectoryServiceAccess`

Remplacez-les avec `-Table` pour en ajouter d'autres (par ex. `AADDomainServicesDNSAuditsGeneral`) ou pour restreindre la recherche.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Username` | Oui | Nom d'utilisateur/SamAccountName/UPN à rechercher — comparé avec `has` sur toutes les colonnes |
| `-WorkspaceId` | * | ID de l'espace de travail Log Analytics (GUID `CustomerId`) |
| `-WorkspaceName` | * | Nom de l'espace de travail — résolu automatiquement ; à combiner avec `-ResourceGroupName` en cas d'ambiguïté |
| `-ResourceGroupName` | Non | Groupe de ressources de l'espace de travail, pour lever l'ambiguïté sur `-WorkspaceName` |
| `-HoursBack` | Non | Nombre d'heures à remonter à partir de maintenant (par défaut : `2`). Ignoré si `-StartTime` est défini |
| `-StartTime` / `-EndTime` | Non | Fenêtre de recherche explicite (heure locale), prévaut sur `-HoursBack` |
| `-Table` | Non | Tables à inclure dans l'union (voir les valeurs par défaut ci-dessus) |
| `-MaxRows` | Non | Nombre maximal de lignes renvoyées, les plus récentes en premier (par défaut : `5000`) |
| `-ExportPath` | Non | Dossier du rapport CSV (par défaut : `C:\Temp`) |

*Si `-WorkspaceId` est omis, le script tente de résoudre automatiquement un espace de travail unique via `-WorkspaceName`/`-ResourceGroupName`, ou l'unique espace de travail de l'abonnement s'il n'y en a qu'un.

**Exemples**

```powershell
# ID d'espace de travail direct, fenêtre par défaut de 2 h
.\Search-AADDSUserActivity.ps1 -Username "jdoe" -WorkspaceId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

# Résoudre l'espace de travail par son nom, remonter 24 heures en arrière
.\Search-AADDSUserActivity.ps1 -Username "jdoe" -WorkspaceName "log-aadds-prod" -HoursBack 24

# Fenêtre de temps explicite
.\Search-AADDSUserActivity.ps1 -Username "jdoe" -WorkspaceId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -StartTime "2026-07-25 00:00" -EndTime "2026-07-27 00:00"
```

**Remarques**
- Se connecte automatiquement avec `Connect-AzAccount` si aucune session Az n'est active
- Le jeu de résultats est plafonné à `-MaxRows` (5000 par défaut) — le script avertit si ce plafond a été atteint, pour que vous sachiez qu'il faut réduire la fenêtre ou relever le plafond
- Le CSV est exporté dans `-ExportPath` sous le nom `AADDSUserActivity_<username>_<timestamp>.csv`

**Modules requis**
```powershell
Install-Module Az.Accounts, Az.OperationalInsights -Scope CurrentUser
```
