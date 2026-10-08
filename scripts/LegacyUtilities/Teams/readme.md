**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [LegacyUtilities](../readme.md) › **Teams**

# Legacy Utilities — Teams

Team and Planner provisioning utilities, all on Microsoft Graph — the MicrosoftTeams module is no longer needed. They sign in through [`Connect-M365.ps1`](../../Startup/readme.md): delegated as the admin by default (browser, or device code / GDAP customer per `load.config.ps1`), app-only with `-ClientId` + `-CertificateThumbprint` or `-AppOnly` (app from `graph.appid.json`). A session for the right tenant that already fits is reused and left connected; only a session the script opened is disconnected. Every script accepts `-TenantId`, `-ClientId`, `-CertificateThumbprint` and `-AppOnly`.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Copy-Team.ps1`](Copy-Team.ps1) ([docs](#copy-teamps1)) | Clone an existing Team (apps/tabs/settings/channels/members) |
| [`Copy-PlannerPlan.ps1`](Copy-PlannerPlan.ps1) ([docs](#copy-plannerplanps1)) | Copy a Planner plan's buckets/tasks/checklists to a new plan |
| [`New-ProjectTeam.ps1`](New-ProjectTeam.ps1) ([docs](#new-projectteamps1)) | Bulk-create Teams with a standard CSV-driven channel template |

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

**Notes**
- Polls the clone operation from the `Location` header and reports the new Team's ID, or the error when the clone failed; before, it waited for any group with the new name, which an existing group with that name also satisfied
- Quotes in a `-SourceTeamId` display name are escaped for the filter

---

### Copy-PlannerPlan.ps1

Copies every bucket and task (with descriptions and checklists) from a source
Planner plan into a newly-created plan on a different group. Consolidates two
old variants (raw Graph REST and PnP.PowerShell) into one script using
Microsoft.Graph.Planner — no PnP dependency needed. Dry-run by default.

```powershell
.\Copy-PlannerPlan.ps1 -SourcePlanId "xqQg5FS2LkCp935s-FIFm2QAFkHM" -DestinationGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Apply
```

**Notes**
- Reads all buckets and tasks page by page (a plan larger than one page lost the rest)
- Checklists are copied again — the items were read from the wrong properties — also for tasks with a checklist but no description
- Delegated, the admin must be a member of both groups; app-only works with the `Tasks.ReadWrite.All` application permission

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
```

**How it works**

1. `POST /groups` creates the Microsoft 365 group with the CSV's `MailNickname`, the owner as owner and member
2. `POST /teams` with `group@odata.bind` and the `standard` template turns it into a Team (404s are retried while the new group replicates)
3. The async operation from the `Location` header is polled until the Team is provisioned
4. `POST /teams/{id}/channels` adds each channel

Delegated scopes: `Group.ReadWrite.All`, `User.Read.All`, `Team.Create`, `Channel.Create`. App-only: `Group.ReadWrite.All`, `User.Read.All` (application permissions). A CSV without `TeamName` or `ChannelName` is now rejected (that check never fired before).
