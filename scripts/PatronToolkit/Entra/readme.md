**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [PatronToolkit](../readme.md) › **Entra**

# Patron Toolkit — Entra

MFA/SSPR registration reporting and Conditional Access policy backup via Microsoft Graph.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-MfaRegistrationReport.ps1`](Get-MfaRegistrationReport.ps1) ([docs](#get-mfaregistrationreportps1)) | Report MFA/SSPR registration status for all or selected users |
| [`Export-ConditionalAccessPolicies.ps1`](Export-ConditionalAccessPolicies.ps1) ([docs](#export-conditionalaccesspoliciesps1)) | Back up all Conditional Access policies and named locations to JSON/CSV |

---

### Get-MfaRegistrationReport.ps1

Reports MFA and SSPR registration status for users via Microsoft Graph's authentication
methods registration details report. Flags users who are not MFA-registered, calling out
admin accounts separately since an unregistered admin account is the highest-priority
finding.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-UserList` | No | One or more UPNs to report on. If omitted, all users are reported |
| `-AdminsOnly` | No | Only report on directory role holders |
| `-NotRegisteredOnly` | No | Only include users who are not MFA-registered |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Tenant ID or domain. Defaults to the GDAP customer (`load.config.ps1`) or your own tenant; required for app-only |
| `-ClientId` | No | App registration for app-only sign-in (with `-CertificateThumbprint`). Without it the script signs in delegated, as you |
| `-CertificateThumbprint` | No | Certificate thumbprint for app-only sign-in with `-ClientId` |
| `-AppOnly` | No | App-only sign-in with the ClientId and thumbprint for the tenant from `graph.appid.json` |

**Examples**

```powershell
# Full tenant report
.\Get-MfaRegistrationReport.ps1

# Only users who still need to register
.\Get-MfaRegistrationReport.ps1 -NotRegisteredOnly

# Admins who aren't MFA-registered — the highest priority finding
.\Get-MfaRegistrationReport.ps1 -AdminsOnly -NotRegisteredOnly
```

**Notes**
- Requires an Entra ID P1/P2 license (the underlying report is a Premium feature)
- Sign-in via [`Connect-M365.ps1`](../../Startup/readme.md#connect-m365ps1): Microsoft
  Graph, delegated by default (scopes `AuditLog.Read.All`, `Reports.Read.All` plus a
  Reports Reader / Security Reader / Global Reader role); app-only with `-ClientId` +
  `-CertificateThumbprint` or `-AppOnly` (application permission `AuditLog.Read.All`)
- `-UserList` matching is now case-insensitive and also works when only one user matches

---

### Export-ConditionalAccessPolicies.ps1

Backs up every Conditional Access policy and named location currently configured in the
tenant to JSON (one file per policy, plus a combined snapshot) and a flattened CSV
summary — a point-in-time backup/change-tracking tool, distinct from
[`Import-ConditionalAccessBaseline.ps1`](../../Entra/readme.md) which imports a specific
community baseline. Read-only.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-OutputPath` | No | Backup folder (default: `.\CAPolicyBackup_<timestamp>\` under `C:\Temp\` / `~/Downloads\`) |
| `-IncludeNamedLocations` | No | Also export named locations (default: on) |
| `-TenantId` | No | Tenant ID or domain. Defaults to the GDAP customer (`load.config.ps1`) or your own tenant; required for app-only |
| `-ClientId` | No | App registration for app-only sign-in (with `-CertificateThumbprint`). Without it the script signs in delegated, as you |
| `-CertificateThumbprint` | No | Certificate thumbprint for app-only sign-in with `-ClientId` |
| `-AppOnly` | No | App-only sign-in with the ClientId and thumbprint for the tenant from `graph.appid.json` |

**Examples**

```powershell
.\Export-ConditionalAccessPolicies.ps1

.\Export-ConditionalAccessPolicies.ps1 -OutputPath "C:\Backups\ContosoCA"

# App-only, with the app registration from graph.appid.json
.\Export-ConditionalAccessPolicies.ps1 -TenantId contoso.onmicrosoft.com -AppOnly
```

**Notes**
- Generic CA policy JSON re-import was intentionally not built — treat the JSON as a
  backup/diff artifact, not an importable format; use
  `scripts/Entra/Import-ConditionalAccessBaseline.ps1` for a maintained import flow
- Sign-in via [`Connect-M365.ps1`](../../Startup/readme.md#connect-m365ps1): Microsoft
  Graph, delegated by default (scope `Policy.Read.All`); app-only with `-ClientId` +
  `-CertificateThumbprint` or `-AppOnly` (application permission `Policy.Read.All`). A
  fitting Graph session that is already open is reused and left connected

**Required module**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```
