# Entra ID Scripts

Scripts for managing users and resources in Microsoft Entra ID (formerly Azure AD) via Microsoft Graph.

---

## Scripts

### Distributionlist.ps1

Resolves members from a Dynamic Distribution Group (Exchange Online) and copies them into a regular Distribution Group.
Also exports the resolved recipient list to CSV.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-DynamicGroupIdentity` | Yes | Source dynamic distribution group identity |
| `-TargetGroupIdentity` | Yes | Target regular distribution group identity |
| `-TargetDisplayName` | No | Display name for new target group |
| `-TargetAlias` | No | Alias for new target group |
| `-TargetPrimarySmtpAddress` | No | SMTP address for new target group |
| `-ClearTargetMembers` | No | Remove existing target members first |
| `-ExportCsvPath` | No | CSV output path for resolved members (default: `C:\Temp\DynamicGroupMembers_<timestamp>.csv` on Windows, `~/Downloads` on Linux/macOS) |
| `-SkipMemberAdd` | No | Only export members, do not update target group |
| `-RenameDynamicGroupTo` | No | Rename source dynamic distribution group after processing |

**Examples**

```powershell
# Resolve dynamic group and populate regular group
.\Distributionlist.ps1 -DynamicGroupIdentity "All Sales" -TargetGroupIdentity "All Sales Static"

# Full refresh of target group members
.\Distributionlist.ps1 -DynamicGroupIdentity sales@contoso.com -TargetGroupIdentity sales-static@contoso.com -ClearTargetMembers

# Dry run
.\Distributionlist.ps1 -DynamicGroupIdentity "All Staff" -TargetGroupIdentity "All Staff Static" -WhatIf

# Convert and rename the original dynamic group
.\Distributionlist.ps1 -DynamicGroupIdentity "All Sales" -TargetGroupIdentity "All Sales Static" -RenameDynamicGroupTo "All Sales (Legacy Dynamic)"

# Keep the original name on the static group:
# 1) dynamic group is renamed first, 2) static group is created with the original display name
.\Distributionlist.ps1 -DynamicGroupIdentity "All Sales" -TargetGroupIdentity "All Sales" -RenameDynamicGroupTo "All Sales (Legacy Dynamic)"
```

**Notes**
- Requires Exchange Online PowerShell module and active EXO session (`Connect-ExchangeOnline`)
- Dynamic Distribution Groups are Exchange objects; this script uses Exchange cmdlets

---

### Remove-M365Users.ps1

Bulk-deletes M365 user accounts from a tenant. Revokes sessions and removes licenses before deletion. Defaults to dry-run mode — pass `-Apply` to perform actual deletions.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-UserList` | * | Array of UPNs to delete |
| `-CsvPath` | * | Path to CSV (`UserPrincipalName` column) or TXT (one UPN per line) |
| `-Apply` | No | Actually perform deletions (default: dry run) |
| `-SkipLicenseRemoval` | No | Skip removing licenses before deletion |
| `-SkipSessionRevoke` | No | Skip revoking active sessions |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

*Either `-UserList` or `-CsvPath` is required.

**Examples**

```powershell
# Dry run — no changes made
.\Remove-M365Users.ps1 -UserList "user1@contoso.com","user2@contoso.com"

# Actual deletion from CSV
.\Remove-M365Users.ps1 -CsvPath .\users.csv -Apply

# Pipeline input
"user1@contoso.com","user2@contoso.com" | .\Remove-M365Users.ps1 -Apply
```

**Notes**
- Deletion is a soft-delete — accounts land in Deleted Users and are recoverable for 30 days
- A CSV report is always written, even in dry-run mode
- If already connected to Graph (e.g. via `functies.ps1`), the existing connection is reused

**Required module**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

---

### New-M365User.ps1

Creates a single M365 user via Microsoft Graph. Generates a random 16-character password if none is supplied. Optionally assigns a license after creation.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-UserPrincipalName` | Yes | UPN of the new user |
| `-DisplayName` | Yes | Display name |
| `-GivenName` | No | First name |
| `-Surname` | No | Last name |
| `-Password` | No | Initial password (auto-generated if omitted) |
| `-UsageLocation` | No | Two-letter ISO country code (default: `NL`) |
| `-Department` | No | Department |
| `-JobTitle` | No | Job title |
| `-MobilePhone` | No | Mobile phone number |
| `-LicenseSkuId` | No | License SKU part number to assign (e.g. `ENTERPRISEPACK`) |
| `-NoPasswordReset` | No | Do not force password change on first sign-in |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
# Minimal — auto-generated password
.\New-M365User.ps1 -UserPrincipalName "j.doe@contoso.com" -DisplayName "Jane Doe"

# Full details with license
.\New-M365User.ps1 -UserPrincipalName "j.doe@contoso.com" -DisplayName "Jane Doe" `
    -GivenName "Jane" -Surname "Doe" -Department "Finance" -LicenseSkuId "ENTERPRISEPACK"
```

---

### Import-M365Users.ps1

Bulk-creates M365 users from a CSV file via Microsoft Graph. Defaults to dry-run mode — pass `-Apply` to create accounts. Generates a unique password per user if no Password column is present. Passwords are written to the results CSV.

**Required CSV columns:** `UserPrincipalName`, `DisplayName`

**Optional CSV columns:** `GivenName`, `Surname` (or `LastName`), `Department`, `JobTitle`, `MobilePhone`, `UsageLocation`, `Password`, `LicenseSkuId`

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-CsvPath` | Yes | Path to the input CSV |
| `-Apply` | No | Actually create accounts (default: dry run) |
| `-UsageLocation` | No | Default country code for all users (default: `NL`) |
| `-LicenseSkuId` | No | Assign this license to all users (overrides CSV column) |
| `-NoPasswordReset` | No | Do not force password change on first sign-in |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
# Dry run first
.\Import-M365Users.ps1 -CsvPath .\users.csv

# Create accounts
.\Import-M365Users.ps1 -CsvPath .\users.csv -Apply

# Create with a license for everyone
.\Import-M365Users.ps1 -CsvPath .\users.csv -LicenseSkuId "ENTERPRISEPACK" -Apply
```

**Notes**
- Dry-run always writes a results CSV — check it before running with `-Apply`
- Generated passwords are in the results CSV — share securely
- A license requires `UsageLocation` to be set; the script handles this automatically

---

### Get-M365UserLicenses.ps1

Checks assigned licenses for a list of users via Microsoft Graph and exports a CSV report. Supports inline list, CSV, TXT, and pipeline input.

**Input options**
- `-UserList` with UPNs/emails
- `-CsvPath` to `.csv` (column: `UserPrincipalName`, `UPN`, or `Mail`)
- `-CsvPath` to `.txt` (one UPN/email per line)

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-UserList` | * | Array of UPNs/emails |
| `-CsvPath` | * | Path to CSV/TXT with users |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\UserLicenseReport_<timestamp>.csv`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

*Either `-UserList` or `-CsvPath` is required.

**Examples**

```powershell
# Check specific users
.\Get-M365UserLicenses.ps1 -UserList "user1@contoso.com","user2@contoso.com"

# Check users from CSV
.\Get-M365UserLicenses.ps1 -CsvPath .\users.csv

# Pipeline input
"user1@contoso.com","user2@contoso.com" | .\Get-M365UserLicenses.ps1
```

**Report output**
- One row per user-license assignment
- Unlicensed users are included with `LicenseStatus = Unlicensed`
- Not found users are included with `LicenseStatus = NotFound`

---

### Import-ConditionalAccessBaseline.ps1

Imports the latest [ConditionalAccessBaseline](https://github.com/j0eyv/ConditionalAccessBaseline) into your tenant using Microsoft Graph.

What it does:
- Downloads the latest baseline (or uses `-SourcePath`)
- Creates/reuses required CA exclusion groups
- Creates/reuses named locations
- Remaps old baseline IDs to your tenant IDs
- Imports Conditional Access policies for all personas/platforms
- Imports policies as **Off** by default (`state = disabled`)

It also supports a follow-up action to switch imported policies to report-only or enabled.

**Parameters (most used)**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Action` | No | `Import` (default) or `SetState` |
| `-PolicyStateOnImport` | No | `disabled` (default) or `enabledForReportingButNotEnforced` |
| `-TargetState` | No | For `SetState`: `disabled`, `enabledForReportingButNotEnforced`, or `enabled` |
| `-SourcePath` | No | Local baseline folder containing `Config\...` |
| `-TenantId` | No | Tenant ID or domain |
| `-UpdateExisting` | No | Update existing policies with matching display names |

**Examples**

```powershell
# Import latest baseline and keep all CA policies OFF
.\Import-ConditionalAccessBaseline.ps1

# Import baseline in report-only mode
.\Import-ConditionalAccessBaseline.ps1 -PolicyStateOnImport enabledForReportingButNotEnforced

# Later: enable imported baseline policies
.\Import-ConditionalAccessBaseline.ps1 -Action SetState -TargetState enabled
```

**Notes**
- Keep at least one break-glass account excluded before enabling policies
- Review exclusion groups and named locations after import
