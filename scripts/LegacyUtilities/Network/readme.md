# Legacy Utilities — Network

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Connect-AzureFileShareDrive.ps1`](#connect-azurefilesharedriveps1) | Mount an Azure Files SMB share as a persistent drive letter |

---

### Connect-AzureFileShareDrive.ps1

Tests SMB (port 445) connectivity to the storage account, saves its access key
via `cmdkey`, and maps the share with `New-PSDrive`. Generalized replacement
for a script that hardcoded a specific storage account name and access key —
this version takes them as parameters and never persists the key beyond what
`cmdkey` itself stores. Dry-run by default (connectivity test only, no mount).

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-StorageAccountName` | Yes | Azure Storage account name |
| `-ShareName` | Yes | File share name |
| `-StorageAccountKey` | Yes | Storage account access key |
| `-DriveLetter` | No | Drive letter (default: `S`) |
| `-Persist` | No | Persist the mapping across reboots |
| `-Apply` | No | Actually save the credential and mount (default: preview/test only) |

```powershell
.\Connect-AzureFileShareDrive.ps1 -StorageAccountName "contosofiles" -ShareName "documents" -StorageAccountKey $key -DriveLetter S -Persist -Apply
```

**Notes**
- Requires outbound TCP 445 to `*.file.core.windows.net` — blocked by some
  ISPs/firewalls. Use an Azure P2S/S2S VPN or ExpressRoute to tunnel SMB
  traffic over a different port if 445 is unavailable.
- For SharePoint/OneDrive document library mapping (not raw Azure Files),
  see [`scripts/Device/DriveMapping/`](../../Device/DriveMapping/readme.md) instead.
