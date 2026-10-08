**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [PatronToolkit](../readme.md) › **Teams**

# Patron Toolkit — Teams

Microsoft Teams tenant governance and inventory reporting.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-TeamsConfigReport.ps1`](Get-TeamsConfigReport.ps1) ([docs](#get-teamsconfigreportps1)) | Report Teams tenant governance settings and team inventory |

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
| `-IncludeTeamsInventory` | No | Also list every team with member/owner counts, via Graph (default: on) |
| `-OutputPath` | No | Folder for the CSV report(s) (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Tenant ID or domain. Defaults to the GDAP customer (`load.config.ps1`) or your own tenant; required for app-only |
| `-ClientId` | No | App registration for app-only sign-in (with `-CertificateThumbprint`). Without it the script signs in delegated, as you |
| `-CertificateThumbprint` | No | Certificate thumbprint for app-only sign-in with `-ClientId` |
| `-AppOnly` | No | App-only sign-in with the ClientId and thumbprint for the tenant from `graph.appid.json` |

**Examples**

```powershell
.\Get-TeamsConfigReport.ps1

.\Get-TeamsConfigReport.ps1 -IncludeTeamsInventory:$false

# App-only, Teams and Graph with the app from graph.appid.json
.\Get-TeamsConfigReport.ps1 -TenantId contoso.onmicrosoft.com -AppOnly
```

**Notes**
- Read-only
- The `Cs*` tenant policies stay on Teams PowerShell (`Connect-M365Teams`) — Microsoft
  Graph has no API for them. The team inventory moved to Graph: `/groups` filtered on
  Team provisioning, `/teams/{id}` (`isArchived`) and `/teams/{id}/members` (owner role),
  replacing `Get-Team` / `Get-TeamUser`
- Guest access now reads `AllowGuestUser` (`Get-CsTeamsClientConfiguration`) and
  `DisableAnonymousJoin` (`Get-CsTeamsMeetingConfiguration`); the earlier version read
  `AllowAnonymousUsersToJoinMeeting` from the guest meeting configuration, which has no
  such property, and always reported an empty value
- Sign-in via [`Connect-M365.ps1`](../../Startup/readme.md#connect-m365ps1), delegated by
  default (Teams Administrator or Global Reader role; Graph scopes `Group.Read.All`,
  `TeamMember.Read.All`, `TeamSettings.Read.All`); app-only with `-ClientId` +
  `-CertificateThumbprint` or `-AppOnly` for both. With `-IncludeTeamsInventory:$false`
  there is no Graph sign-in

**Required modules**
```powershell
Install-Module MicrosoftTeams -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```
