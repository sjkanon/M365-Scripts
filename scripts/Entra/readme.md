# Entra ID Scripts

Scripts for managing users and resources in Microsoft Entra ID (formerly Azure AD) via Microsoft Graph.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Set-UserManager.ps1`](#set-usermanagerps1) | Report and optionally bulk-set the manager for a set of Entra ID users |
| [`Remove-M365Users.ps1`](#remove-m365usersps1) | Bulk-delete M365 user accounts (dry-run by default) |
| [`New-M365User.ps1`](#new-m365userps1) | Create a single M365 user, optional license |
| [`Import-M365Users.ps1`](#import-m365usersps1) | Bulk-create M365 users from CSV (dry-run by default) |
| [`Get-M365UserLicenses.ps1`](#get-m365userlicensesps1) | Report assigned licenses for a list of users |
| [`Import-ConditionalAccessBaseline.ps1`](#import-conditionalaccessbaselineps1) | Import the community Conditional Access baseline |
| [`Test-M365GroupMembership.ps1`](#test-m365groupmembershipps1) | Audit M365 Group / Teams owners and members |
| [`Copy-GroupMember.ps1`](#copy-groupmemberps1) | Copy members from one Entra ID group into another (dry-run by default) |
| [`New-TemporaryConditionalAccessPolicy.ps1`](#new-temporaryconditionalaccesspolicyps1) | Create a temporary CA policy for one user or group |
| [`Remove-TemporaryConditionalAccessPolicies.ps1`](#remove-temporaryconditionalaccesspoliciesps1) | Remove expired/all temporary CA policies |
| [`New-UserTemporaryAccessPass.ps1`](#new-usertemporaryaccesspassps1) | Create a TAP code for a user |

> Dynamic-to-static distribution group conversion (`Set-Distributionlist-dynamic-static.ps1`) lives in [`scripts/Exchange/`](../Exchange/readme.md) — it uses Exchange Online cmdlets, not Graph.

---

### Set-UserManager.ps1

Reports and optionally bulk-sets the manager for a set of Entra ID users. Users can be selected by group, department, current manager, or an explicit UPN list; when `-NewManager` is omitted the script only reports each user's current manager.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-GroupId` | * | Object ID of an Entra ID group whose members should be processed |
| `-GroupName` | * | Display name of an Entra ID group (resolved to an ID automatically; errors if ambiguous) |
| `-Department` | * | Department string to filter users by (exact match via OData) |
| `-CurrentManager` | * | UPN or object ID of a manager — processes all their direct reports tenant-wide |
| `-UserList` | * | Explicit array of UPNs or object IDs |
| `-NewManager` | No | UPN or object ID to set as manager for all resolved users. Omit to only report current managers |
| `-OutputPath` | No | CSV export path |
| `-TenantId` | No | Entra ID tenant ID or domain for `Connect-MgGraph` |

*Exactly one of `-GroupId` / `-GroupName` / `-Department` / `-CurrentManager` / `-UserList` selects the user source.

**Examples**

```powershell
# Show managers for all members of a group
.\Set-UserManager.ps1 -GroupName "Sales Team"

# Bulk-set manager for a group
.\Set-UserManager.ps1 -GroupName "Sales Team" -NewManager "jane.doe@contoso.com"

# Bulk-set manager for a department, export report
.\Set-UserManager.ps1 -Department "Logistics" -NewManager "jane.doe@contoso.com" -OutputPath C:\Temp\ManagerReport.csv

# Find all direct reports of a manager
.\Set-UserManager.ps1 -CurrentManager "old.boss@contoso.com"

# Re-assign all direct reports of one manager to another
.\Set-UserManager.ps1 -CurrentManager "old.boss@contoso.com" -NewManager "new.boss@contoso.com"
```

Supports `-WhatIf` (`SupportsShouldProcess`).

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

---

### New-TemporaryConditionalAccessPolicy.ps1

Creates a temporary Conditional Access policy for one user or group.

- Marks the policy name with prefix `TEMP-CA -`
- Writes an expiry timestamp in policy description (`Expires=<UTC timestamp>`)
- Supports either duration-based windows or exact local start/end date-time windows
- By default keeps the script session open and removes the policy immediately when expiry is reached

Important:
- Immediate cleanup at expiry requires the script session to stay open
- If you close the session early, run the cleanup script later
- This script does not create a Windows Scheduled Task; the wait/cleanup loop runs in the current session

**Examples**

```powershell
# Temporary MFA requirement for 2 hours, auto-remove at expiry
.\New-TemporaryConditionalAccessPolicy.ps1 -TargetType User -TargetId "<object-id>" -DisplayName "Temporary MFA" -DurationHours 2

# Temporary policy with exact local start/end date-time
.\New-TemporaryConditionalAccessPolicy.ps1 -TargetType User -TargetId "<object-id>" -DisplayName "Install Window" -StartDateTimeLocal "2026-07-23 19:00" -EndDateTimeLocal "2026-07-23 22:00"

# Temporary block policy, no auto cleanup wait loop
.\New-TemporaryConditionalAccessPolicy.ps1 -TargetType Group -TargetId "<object-id>" -DisplayName "Temporary Block" -Action Block -NoAutoCleanup
```

---

### Remove-TemporaryConditionalAccessPolicies.ps1

Removes temporary policies created with prefix `TEMP-CA -`.

Modes:
- default: remove only expired TEMP-CA policies
- `-PolicyId`: remove one specific policy
- `-RemoveAllTempPolicies`: remove all TEMP-CA policies

**Examples**

```powershell
# Remove only expired temporary policies
.\Remove-TemporaryConditionalAccessPolicies.ps1

# Remove one specific policy
.\Remove-TemporaryConditionalAccessPolicies.ps1 -PolicyId "<policy-id>"
```

---

### New-UserTemporaryAccessPass.ps1

Creates a Temporary Access Pass (TAP) for one user.

**Example**

```powershell
.\New-UserTemporaryAccessPass.ps1 -UserId "user@contoso.com" -LifetimeMinutes 60 -IsUsableOnce
```

**Notes**
- Prefer `-IsUsableOnce` for support/install scenarios
- Share the TAP code through a secure channel and expire it quickly

---

### Test-M365GroupMembership.ps1

Lists all owners and members of Microsoft 365 Groups (including Teams-backed groups). Results are exported to CSV with one row per owner/member entry. Connects to Graph automatically if no session is active; reuses an existing session if already connected.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Group` | No | Display name or Object ID of a single group. If omitted, all M365 groups are audited |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
# Audit all M365 groups
.\Test-M365GroupMembership.ps1

# Single group by display name
.\Test-M365GroupMembership.ps1 -Group "Team Finance"

# Single group by Object ID
.\Test-M365GroupMembership.ps1 -Group "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```

**Required scopes**
- `Group.Read.All`
- `Directory.Read.All`

**Required module**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

---

### Copy-GroupMember.ps1

Copies the members of one Entra ID group into another group. Members already present in the target are skipped, so the script is safe to re-run. Defaults to dry-run — pass `-Apply` to write changes.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-SourceGroup` | Yes | Display name or Object ID of the group to copy FROM |
| `-TargetGroup` | Yes | Display name or Object ID of the group to copy TO |
| `-MemberType` | No | `All` (default), `User`, `Group`, `Device` or `ServicePrincipal` |
| `-Flatten` | No | Expand nested groups and copy their effective members instead of the nested group object |
| `-Mirror` | No | Also remove members from the target that are not in the source (exact copy instead of union) |
| `-Apply` | No | Actually add/remove members (default: dry run) |
| `-Disconnect` | No | Sign out of Graph when finished (off by default — disconnecting clears the token cache and forces a new browser prompt next run) |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\GroupMemberCopy_<timestamp>.csv`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
# Dry run — show what would be copied
.\Copy-GroupMember.ps1 -SourceGroup "All Staff" -TargetGroup "MFA Rollout"

# Actually copy the members
.\Copy-GroupMember.ps1 -SourceGroup "All Staff" -TargetGroup "MFA Rollout" -Apply

# Copy only users, expanding nested groups
.\Copy-GroupMember.ps1 -SourceGroup "Sales" -TargetGroup "Sales Mail" -MemberType User -Flatten -Apply

# Make the target an exact copy of the source (adds and removes)
.\Copy-GroupMember.ps1 -SourceGroup "Pilot" -TargetGroup "Pilot Copy" -Mirror -Apply
```

**Notes**
- Display names are resolved via Graph; an ambiguous name is a hard error — use the Object ID instead
- A dynamic-membership target group is rejected: its membership is rule-driven and cannot be edited
- Mail-enabled security and distribution groups are not writable through Graph — use Exchange Online cmdlets for those
- Supports `-WhatIf` (`SupportsShouldProcess`)

**Required scopes**
- `Group.Read.All`
- `GroupMember.ReadWrite.All`

**Required module**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```
