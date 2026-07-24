# Patron Toolkit — Entra

MFA/SSPR registration reporting and Conditional Access policy backup via Microsoft Graph.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-MfaRegistrationReport.ps1`](#get-mfaregistrationreportps1) | Report MFA/SSPR registration status for all or selected users |
| [`Export-ConditionalAccessPolicies.ps1`](#export-conditionalaccesspoliciesps1) | Back up all Conditional Access policies and named locations to JSON/CSV |

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
| `-TenantId` | No | Entra ID tenant ID or domain |

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
- Required scope: `Reports.Read.All` (or `AuditLog.Read.All`)

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
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Export-ConditionalAccessPolicies.ps1

.\Export-ConditionalAccessPolicies.ps1 -OutputPath "C:\Backups\ContosoCA"
```

**Notes**
- Generic CA policy JSON re-import was intentionally not built — treat the JSON as a
  backup/diff artifact, not an importable format; use
  `scripts/Entra/Import-ConditionalAccessBaseline.ps1` for a maintained import flow
- Required scope: `Policy.Read.All`

**Required module**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```
