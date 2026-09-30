[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [LegacyUtilities](../readme.fr.md) › **Teams**

# Legacy Utilities — Teams

Utilitaires de provisionnement d'équipes et de Planner via Microsoft Graph (et Microsoft Teams
PowerShell pour la création en masse d'équipes et de canaux). Ils se connectent automatiquement
si aucune session n'est active, et réutilisent la session existante si vous êtes déjà connecté.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Copy-Team.ps1`](Copy-Team.ps1) ([docs](#copy-teamps1)) | Cloner une équipe existante (applications/onglets/paramètres/canaux/membres) |
| [`Copy-PlannerPlan.ps1`](Copy-PlannerPlan.ps1) ([docs](#copy-plannerplanps1)) | Copier les compartiments/tâches/listes de contrôle d'un plan Planner vers un nouveau plan |
| [`New-ProjectTeam.ps1`](New-ProjectTeam.ps1) ([docs](#new-projectteamps1)) | Créer des équipes en masse avec un modèle de canaux standard piloté par CSV |

---

### Copy-Team.ps1

Soumet le clonage d'une équipe via Graph (`POST /teams/{id}/clone`, asynchrone) et interroge
jusqu'à ce que la nouvelle équipe apparaisse. Essai à blanc par défaut.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-SourceTeamId` | Oui | ID d'objet ou nom d'affichage de l'équipe à cloner |
| `-NewTeamName` | Oui | Nom d'affichage du clone |
| `-NewTeamDescription` / `-NewMailNickname` | Non | Reprennent `-NewTeamName` s'ils sont omis |
| `-Visibility` | Non | `Private` (par défaut) ou `Public` |
| `-PartsToClone` | Non | Un ou plusieurs éléments parmi Apps, Tabs, Settings, Channels, Members (par défaut : tous) |
| `-Apply` | Non | Soumettre réellement le clonage (par défaut : aperçu) |

```powershell
.\Copy-Team.ps1 -SourceTeamId "Project Template" -NewTeamName "Project 1234" -Apply
```

---

### Copy-PlannerPlan.ps1

Copie chaque compartiment et chaque tâche (avec descriptions et listes de contrôle) d'un plan
Planner source vers un plan nouvellement créé dans un autre groupe. Regroupe deux anciennes
variantes (Graph REST brut et PnP.PowerShell) en un seul script utilisant
Microsoft.Graph.Planner, sans dépendance à PnP. Essai à blanc par défaut.

```powershell
.\Copy-PlannerPlan.ps1 -SourcePlanId "xqQg5FS2LkCp935s-FIFm2QAFkHM" -DestinationGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Apply
```

---

### New-ProjectTeam.ps1

Crée des équipes en masse à partir d'un CSV (une ligne par équipe) et applique à chacune le même
modèle de canaux (un second CSV). Remplacement généralisé d'un ancien script qui codait en dur
la structure fixe des canaux de projet d'une entreprise. Essai à blanc par défaut.

**Fichiers CSV en entrée**

```csv
# teams.csv
TeamName,MailNickname,Owner,Visibility
"Project 1001","project-1001","pm@contoso.com","Private"

# channels.csv
ChannelName,Description
"Documents","Signed customer documents"
"Internal",""
```

```powershell
.\New-ProjectTeam.ps1 -TeamsCsvPath .\teams.csv -ChannelsCsvPath .\channels.csv -Apply
```

---

**Modules requis**

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module MicrosoftTeams -Scope CurrentUser
```
