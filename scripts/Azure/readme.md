# Azure Infrastructure Scripts

Scripts for managing Azure IaaS resources directly (not the M365 tenant) — separate from every other category in this repo, which targets Microsoft 365 / Entra ID via Graph or Exchange Online. Requires the `Az` PowerShell module and an authenticated `Connect-AzAccount` session. Not wired into [`menu.ps1`](../../menu.ps1) — run directly against the target subscription.

---

## Folders

| Folder | Description |
|--------|-------------|
| [`VM/`](VM/readme.md) | Convert Azure VM disk controller type between SCSI and NVMe |

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Search-AADDSUserActivity.ps1`](#search-aaddsuseractivityps1) | Search all Azure AD Domain Services audit tables in Log Analytics for a single user in one query |

---

### Search-AADDSUserActivity.ps1

Searches across the AAD DS diagnostic tables in a Log Analytics workspace for a given username in one `union` query, so you don't need to know beforehand which table (or column) an event landed in. Requires diagnostic settings on the AAD DS managed domain sending logs to the target workspace.

**Default tables searched**
- `AADDomainServicesAccountManagement`
- `AADDomainServicesAccountLogon`
- `AADDomainServicesLogonLogoff`
- `AADDomainServicesDirectoryServiceAccess`

Override with `-Table` to add others (e.g. `AADDomainServicesDNSAuditsGeneral`) or narrow the search down.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Username` | Yes | Username/SamAccountName/UPN to search for — matched with `has` across every column |
| `-WorkspaceId` | * | Log Analytics workspace ID (`CustomerId` GUID) |
| `-WorkspaceName` | * | Workspace name — resolved automatically; combine with `-ResourceGroupName` if ambiguous |
| `-ResourceGroupName` | No | Resource group of the workspace, to disambiguate `-WorkspaceName` |
| `-HoursBack` | No | Hours to look back from now (default: `2`). Ignored if `-StartTime` is set |
| `-StartTime` / `-EndTime` | No | Explicit search window (local time), overrides `-HoursBack` |
| `-Table` | No | Tables to include in the union (see defaults above) |
| `-MaxRows` | No | Max rows returned, most recent first (default: `5000`) |
| `-ExportPath` | No | Folder for the CSV report (default: `C:\Temp`) |

*If `-WorkspaceId` is omitted, the script tries to auto-resolve a single workspace via `-WorkspaceName`/`-ResourceGroupName`, or the only workspace in the subscription if there's just one.

**Examples**

```powershell
# Direct workspace ID, default 2h window
.\Search-AADDSUserActivity.ps1 -Username "jdoe" -WorkspaceId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

# Resolve workspace by name, look back 24 hours
.\Search-AADDSUserActivity.ps1 -Username "jdoe" -WorkspaceName "log-aadds-prod" -HoursBack 24

# Explicit time window
.\Search-AADDSUserActivity.ps1 -Username "jdoe" -WorkspaceId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -StartTime "2026-07-25 00:00" -EndTime "2026-07-27 00:00"
```

**Notes**
- Connects with `Connect-AzAccount` automatically if no Az session is active
- Result set is capped at `-MaxRows` (default 5000) — the script warns if the cap was hit, so you know to narrow the window or raise it
- CSV is exported to `-ExportPath` as `AADDSUserActivity_<username>_<timestamp>.csv`

**Required modules**
```powershell
Install-Module Az.Accounts, Az.OperationalInsights -Scope CurrentUser
```
