# Patron Toolkit — Teams

Microsoft Teams tenant governance and inventory reporting.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-TeamsConfigReport.ps1`](#get-teamsconfigreportps1) | Report Teams tenant governance settings and team inventory |

---

### Get-TeamsConfigReport.ps1

Reports the tenant-wide Teams policies that matter most for a governance/security
review: external access (federation), guest access, global meeting policy (anonymous
join, recording, presenter role), global messaging policy, and app setup policy
(sideloading). Also lists every team with visibility, archived status, and owner/member
counts — flagging any team with zero owners.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-IncludeTeamsInventory` | No | Also list every team with member/owner counts (default: on) |
| `-OutputPath` | No | Folder for the CSV report(s) (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Get-TeamsConfigReport.ps1

.\Get-TeamsConfigReport.ps1 -IncludeTeamsInventory:$false
```

**Notes**
- Read-only

**Required module**
```powershell
Install-Module MicrosoftTeams -Scope CurrentUser
```
