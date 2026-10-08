[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [LegacyUtilities](../readme.fr.md) › **Teams**

# Legacy Utilities — Teams

Utilitaires de provisionnement d'équipes et de Planner, entièrement via Microsoft Graph — le module MicrosoftTeams n'est plus nécessaire. Ils se connectent via [`Connect-M365.ps1`](../../Startup/readme.fr.md) : en délégué en tant qu'administrateur par défaut (navigateur, ou code d'appareil / client GDAP selon `load.config.ps1`), en app-only avec `-ClientId` + `-CertificateThumbprint` ou `-AppOnly` (application de `graph.appid.json`). Une session adaptée pour le bon tenant est réutilisée et reste connectée ; seule une session ouverte par le script est fermée. Chaque script accepte `-TenantId`, `-ClientId`, `-CertificateThumbprint` et `-AppOnly`.

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

**Remarques**
- Interroge l'opération de clonage indiquée dans l'en-tête `Location` et affiche l'ID de la nouvelle équipe, ou l'erreur si le clonage a échoué ; auparavant, il attendait n'importe quel groupe portant le nouveau nom, condition qu'un groupe existant du même nom remplissait aussi
- Les apostrophes dans un nom d'affichage `-SourceTeamId` sont échappées pour le filtre

---

### Copy-PlannerPlan.ps1

Copie chaque compartiment et chaque tâche (avec descriptions et listes de contrôle) d'un plan
Planner source vers un plan nouvellement créé dans un autre groupe. Regroupe deux anciennes
variantes (Graph REST brut et PnP.PowerShell) en un seul script utilisant
Microsoft.Graph.Planner, sans dépendance à PnP. Essai à blanc par défaut.

```powershell
.\Copy-PlannerPlan.ps1 -SourcePlanId "xqQg5FS2LkCp935s-FIFm2QAFkHM" -DestinationGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Apply
```

**Remarques**
- Lit tous les compartiments et toutes les tâches page par page (un plan dépassant une page perdait le reste)
- Les listes de contrôle sont à nouveau copiées — les éléments étaient lus dans les mauvaises propriétés — y compris pour les tâches avec une liste de contrôle mais sans description
- En délégué, l'administrateur doit être membre des deux groupes ; l'app-only fonctionne avec l'autorisation d'application `Tasks.ReadWrite.All`

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
```

**Fonctionnement**

1. `POST /groups` crée le groupe Microsoft 365 avec le `MailNickname` du CSV, le propriétaire étant propriétaire et membre
2. `POST /teams` avec `group@odata.bind` et le modèle `standard` en fait une équipe (les 404 sont retentés pendant la réplication du nouveau groupe)
3. L'opération asynchrone de l'en-tête `Location` est interrogée jusqu'à ce que l'équipe soit provisionnée
4. `POST /teams/{id}/channels` ajoute chaque canal

Étendues déléguées : `Group.ReadWrite.All`, `User.Read.All`, `Team.Create`, `Channel.Create`. App-only : `Group.ReadWrite.All`, `User.Read.All` (autorisations d'application). Un CSV sans `TeamName` ou `ChannelName` est désormais refusé (ce contrôle ne se déclenchait jamais auparavant).
