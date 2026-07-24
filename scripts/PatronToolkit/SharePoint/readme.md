# Patron Toolkit — SharePoint

SharePoint Online tenant sharing configuration and external user auditing.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Test-SharePointSharingConfig.ps1`](#test-sharepointsharingconfigps1) | Report tenant sharing settings, per-site overrides, and external users |

---

### Test-SharePointSharingConfig.ps1

Reports the tenant-wide external sharing settings that matter most for a security review
(sharing capability, default link type, anonymous link expiry/permission, re-sharing by
external users, legacy auth). Optionally also enumerates every external (guest) user
across all site collections and flags individual sites whose sharing capability is
broader than the tenant default.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-TenantName` | * | SharePoint tenant name, e.g. `contoso` for `https://contoso-admin.sharepoint.com` |
| `-AdminUrl` | * | Full SharePoint admin center URL (alternative to `-TenantName`) |
| `-IncludeExternalUsers` | No | Also enumerate external users across all sites (slower) |
| `-IncludeSiteOverrides` | No | Also flag sites with broader-than-default sharing |
| `-OutputPath` | No | Folder for the CSV report(s) (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Passed through to `Connect-SPOService` if supported |

*Required unless already connected via `Connect-SPOService`.

**Examples**

```powershell
.\Test-SharePointSharingConfig.ps1 -TenantName "contoso"

.\Test-SharePointSharingConfig.ps1 -TenantName "contoso" -IncludeExternalUsers -IncludeSiteOverrides
```

**Notes**
- For SharePoint storage/version reporting, see
  [`Reporting/Get-SharePointStorageReport.ps1`](../../Reporting/readme.md) — a separate,
  Graph-based capability already in this repo

**Required module**
```powershell
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
```
