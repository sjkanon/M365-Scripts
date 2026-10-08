**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [Office365Toolkit](../readme.md) › **Intune**

# Office365Toolkit / Intune

Tenant-wide Intune / Endpoint Manager policy inventory.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-IntunePolicyInventory.ps1`](Get-IntunePolicyInventory.ps1) ([docs](#get-intunepolicyinventoryps1)) | Inventory all compliance/configuration/app protection/Endpoint Security policies |

---

### Get-IntunePolicyInventory.ps1

Lists every policy across the main Intune policy surfaces — device compliance
policies, device configuration profiles, Settings Catalog configuration
policies, app protection policies, and Endpoint Security ("intents") policies —
with assignment counts. A quick point-in-time inventory/checklist, not a
baseline diff. Read-only.

Signs in through [`Connect-M365.ps1`](../../Startup/readme.md#connect-m365ps1):
delegated (you sign in as the admin) by default, with device code and the GDAP
customer taken from `load.config.ps1`; app-only with `-ClientId` +
`-CertificateThumbprint`, or `-AppOnly` (`graph.appid.json`). A Graph session
that already fits is reused and left connected.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-OutputPath` | No | CSV report file path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Tenant ID or domain; defaults to the GDAP customer from `load.config.ps1` |
| `-ClientId` | No | App registration for app-only sign-in (with `-CertificateThumbprint`) |
| `-CertificateThumbprint` | No | Certificate thumbprint for app-only sign-in |
| `-AppOnly` | No | App-only with ClientId and thumbprint from `graph.appid.json` |

**Examples**

```powershell
.\Get-IntunePolicyInventory.ps1

.\Get-IntunePolicyInventory.ps1 -OutputPath C:\Reports\intune.csv

.\Get-IntunePolicyInventory.ps1 -TenantId contoso.onmicrosoft.com -AppOnly
```

**Notes**
- Settings Catalog (`configurationPolicies`) and Endpoint Security (`intents`)
  exist only on the Graph beta endpoint. The v1.0 SDK has no cmdlets for them, so
  earlier versions caught the error and silently left both out. All five
  surfaces are now read with `Invoke-MgGraphRequest`, following
  `@odata.nextLink`.
- A policy with no assignments now shows `0` instead of an empty count. App
  protection policies have no assignments on the base `managedAppPolicy` type,
  so their count stays empty.
- For comparing a customer tenant's Intune configuration against an MSP
  reference baseline (drift detection), use
  [`scripts/Intune/Compare-IntuneConfig.ps1`](../../Intune/readme.md#compare-intuneconfigps1)
  instead — that script does a full backup-based diff; this one is a quick
  inventory of what currently exists.

**Required scopes:** `DeviceManagementConfiguration.Read.All`, `DeviceManagementApps.Read.All`
**Required module:** `Microsoft.Graph.Authentication`
