# Office365Toolkit / Intune

Tenant-wide Intune / Endpoint Manager policy inventory.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-IntunePolicyInventory.ps1`](#get-intunepolicyinventoryps1) | Inventory all compliance/configuration/app protection/Endpoint Security policies |

---

### Get-IntunePolicyInventory.ps1

Lists every policy across the main Intune policy surfaces — device compliance
policies, device configuration profiles, Settings Catalog configuration
policies, app protection policies, and Endpoint Security ("intents") policies —
with assignment counts. A quick point-in-time inventory/checklist, not a
baseline diff. Read-only.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Get-IntunePolicyInventory.ps1

.\Get-IntunePolicyInventory.ps1 -OutputPath C:\Reports
```

**Notes**
- For comparing a customer tenant's Intune configuration against an MSP
  reference baseline (drift detection), use
  [`scripts/Intune/Compare-IntuneConfig.ps1`](../../Intune/readme.md#compare-intuneconfigps1)
  instead — that script does a full backup-based diff; this one is a quick
  inventory of what currently exists.

**Required scopes:** `DeviceManagementConfiguration.Read.All`, `DeviceManagementApps.Read.All`
**Required module:** `Microsoft.Graph.DeviceManagement`
