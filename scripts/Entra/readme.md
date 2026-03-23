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
