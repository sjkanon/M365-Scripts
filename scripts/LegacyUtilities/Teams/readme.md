# Legacy Utilities — Teams

Team and Planner provisioning utilities via Microsoft Graph (and Microsoft
Teams PowerShell for bulk Team/channel creation). Connect automatically if no
session is active; reuse an existing session if already connected.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Copy-Team.ps1`](#copy-teamps1) | Clone an existing Team (apps/tabs/settings/channels/members) |
| [`Copy-PlannerPlan.ps1`](#copy-plannerplanps1) | Copy a Planner plan's buckets/tasks/checklists to a new plan |
| [`New-ProjectTeam.ps1`](#new-projectteamps1) | Bulk-create Teams with a standard CSV-driven channel template |

---

### Copy-Team.ps1

Submits a Team clone via Graph (`POST /teams/{id}/clone`, asynchronous) and
polls until the new Team appears. Dry-run by default.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-SourceTeamId` | Yes | Object ID or display name of the Team to clone |
| `-NewTeamName` | Yes | Display name for the clone |
| `-NewTeamDescription` / `-NewMailNickname` | No | Default to `-NewTeamName` if omitted |
| `-Visibility` | No | `Private` (default) or `Public` |
| `-PartsToClone` | No | Any of Apps, Tabs, Settings, Channels, Members (default: all) |
| `-Apply` | No | Actually submit the clone (default: preview) |

```powershell
.\Copy-Team.ps1 -SourceTeamId "Project Template" -NewTeamName "Project 1234" -Apply
```

---

### Copy-PlannerPlan.ps1

Copies every bucket and task (with descriptions and checklists) from a source
Planner plan into a newly-created plan on a different group. Consolidates two
old variants (raw Graph REST and PnP.PowerShell) into one script using
Microsoft.Graph.Planner — no PnP dependency needed. Dry-run by default.

```powershell
.\Copy-PlannerPlan.ps1 -SourcePlanId "xqQg5FS2LkCp935s-FIFm2QAFkHM" -DestinationGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Apply
```

---

### New-ProjectTeam.ps1

Bulk-creates Teams from a CSV (one row per Team) and applies the same channel
template (a second CSV) to each. Generalized replacement for an old script
that hardcoded one company's fixed project-channel layout. Dry-run by default.

**CSV inputs**

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

**Required modules**

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module MicrosoftTeams -Scope CurrentUser
```
