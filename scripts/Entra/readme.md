# Entra ID Scripts

Scripts for managing users and resources in Microsoft Entra ID (formerly Azure AD) via Microsoft Graph.

---

## Scripts

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
| `-OutputPath` | No | CSV report path (default: `.\DeletedAccounts_<timestamp>.csv`) |
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
| `-OutputPath` | No | CSV report path (default: `.\CreatedAccounts_<timestamp>.csv`) |
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
