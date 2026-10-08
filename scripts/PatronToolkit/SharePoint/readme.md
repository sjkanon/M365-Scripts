**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [PatronToolkit](../readme.md) › **SharePoint**

# Patron Toolkit — SharePoint

SharePoint Online tenant sharing configuration and external user auditing.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Test-SharePointSharingConfig.ps1`](Test-SharePointSharingConfig.ps1) ([docs](#test-sharepointsharingconfigps1)) | Report tenant sharing settings, per-site overrides, and external users |

---

### Test-SharePointSharingConfig.ps1

Reports the tenant-wide external sharing settings that matter most for a security review,
read from Microsoft Graph (`GET /admin/sharepoint/settings`): sharing capability, re-sharing
by external users, whether the accepting account must match the invited one, the sharing
domain allow/block mode, legacy auth and idle session sign-out. Optionally, through
PnP.PowerShell, also the default link type / link permission / anonymous link expiry,
sites whose sharing capability is broader than the tenant default, and every external
(guest) user across all site collections.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-TenantName` | No | SharePoint tenant name, e.g. `contoso` for `https://contoso-admin.sharepoint.com`. Only for the PnP part; also turns on `-IncludeLinkSettings` |
| `-AdminUrl` | No | Full SharePoint admin center URL (alternative to `-TenantName`). Without either, the URL is looked up via Graph (`/sites/root`) |
| `-IncludeLinkSettings` | No | Also report default link type, default link permission and anonymous link expiry (PnP) |
| `-IncludeExternalUsers` | No | Also enumerate external users across all sites (PnP, slower) |
| `-IncludeSiteOverrides` | No | Also flag sites with broader-than-default sharing (PnP) |
| `-OutputPath` | No | Folder for the CSV report(s) (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Tenant ID or domain. Defaults to the GDAP customer (`load.config.ps1`) or your own tenant; required for app-only |
| `-ClientId` | No | App registration for app-only sign-in (with `-CertificateThumbprint`), used for Graph and PnP. Without it the script signs in delegated, as you |
| `-CertificateThumbprint` | No | Certificate thumbprint for app-only sign-in with `-ClientId` |
| `-AppOnly` | No | App-only sign-in (Graph and PnP) with the ClientId and thumbprint for the tenant from `graph.appid.json` |
| `-PnPClientId` | No | PnP app registration for the delegated PnP sign-in (default: the tenant's entry in `pnp.appid.json`) |

**Examples**

```powershell
# Graph only: tenant sharing settings
.\Test-SharePointSharingConfig.ps1

# As before: tenant settings plus link settings (PnP)
.\Test-SharePointSharingConfig.ps1 -TenantName "contoso"

.\Test-SharePointSharingConfig.ps1 -IncludeLinkSettings -IncludeExternalUsers -IncludeSiteOverrides

# App-only, Graph and PnP with the app from graph.appid.json
.\Test-SharePointSharingConfig.ps1 -TenantId contoso.onmicrosoft.com -AppOnly -IncludeSiteOverrides
```

**Notes**
- The legacy SharePoint Online Management Shell (`Connect-SPOService`, `Get-SPOTenant`,
  `Get-SPOSite`, `Get-SPOExternalUser`) is no longer used. The earlier `-TenantId` was
  documented but never passed on
- Graph has no API for the link defaults, per-site sharing capability or external users
  per site, so those stay on PnP (`Get-PnPTenant`, `Get-PnPTenantSite`,
  `Get-PnPExternalUser`) and only run when asked for. If the PnP sign-in fails, those parts
  are skipped with a warning and the Graph part is still reported
- Graph reports `sharingCapability` in camelCase (`externalUserAndGuestSharing`) and the
  inverse of the old setting: `ResharingByExternalUsersEnabled` instead of
  `PreventExternalUsersFromResharing`
- Sign-in via [`Connect-M365.ps1`](../../Startup/readme.md#connect-m365ps1): Graph
  delegated by default (scope `SharePointTenantSettings.Read.All`, plus `Sites.Read.All`
  only to look up the admin URL; SharePoint Administrator role); app-only with `-ClientId`
  + `-CertificateThumbprint` or `-AppOnly`. PnP needs its own app registration since
  September 2024: `-PnPClientId` or `pnp.appid.json` for delegated, the same app as Graph
  for app-only (SharePoint `Sites.FullControl.All`)
- For SharePoint storage/version reporting, see
  [`Reporting/Get-SharePointStorageReport.ps1`](../../Reporting/readme.md) — a separate,
  Graph-based capability already in this repo

**Required modules**
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module PnP.PowerShell -Scope CurrentUser   # only for the PnP part
```
