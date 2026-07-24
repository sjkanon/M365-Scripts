# DriveMapping

Maps SharePoint Online / OneDrive document libraries to persistent drive letters via the WebDAV redirector — for use as a per-user logon script (Intune Win32 app or a scheduled task at logon), not via [`menu.ps1`](../../../menu.ps1).

---

## New-CloudDriveMapping.ps1

Converts each `https://` document library URL into its WebDAV UNC form (`\\<host>@SSL\DavWWWRoot\<path>`) and maps it with `net use`. Runs as a dry-run by default — no drive is mapped without `-Apply`.

Relies on the signed-in user already having a valid session/SSO to the tenant (the same way a browser reaches SharePoint over WebDAV) — this is not app-only Graph auth. Requires the Windows **WebClient** service (WebDAV redirector); present by default on Windows 10/11, needs the **Desktop Experience** feature on Windows Server.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-MappingsCsv` | * | CSV with columns `DriveLetter`, `Url`, optional `Label` — one row per mapping |
| `-DriveLetter` | * | Single ad hoc mapping: drive letter (use with `-Url`) |
| `-Url` | * | Single ad hoc mapping: document library URL (use with `-DriveLetter`) |
| `-Label` | No | Friendly name for the single ad hoc mapping |
| `-RemoveExisting` | No | Remove any existing mapping on the target letter(s) first (`net use <letter>: /delete /y`) |
| `-Persist` | No | Make the mapping persistent across reboots (`/persistent:yes`). Off by default — a logon script normally remaps every sign-in |
| `-Apply` | No | Actually perform the mapping (default: dry run) |
| `-OutputPath` | No | Log folder (default: `$env:TEMP` — logon scripts usually run in user context, not admin) |

*Either `-MappingsCsv` or `-DriveLetter` + `-Url` is required.

**Mappings CSV example**

```csv
DriveLetter,Url,Label
S,https://contoso.sharepoint.com/sites/Finance/Shared Documents,Finance Docs
O,https://contoso-my.sharepoint.com/personal/j_doe_contoso_com/Documents,My OneDrive
```

**Examples**

```powershell
# Dry run from a CSV of mappings
.\New-CloudDriveMapping.ps1 -MappingsCsv .\mappings.csv

# Apply mappings from CSV, clearing any existing mapping on those letters first
.\New-CloudDriveMapping.ps1 -MappingsCsv .\mappings.csv -RemoveExisting -Apply

# Single ad hoc mapping
.\New-CloudDriveMapping.ps1 -DriveLetter Z -Url "https://contoso.sharepoint.com/sites/Finance/Shared Documents" -Apply
```

**Deployment as a logon script**
- Package as an Intune Win32 app (or PowerShell script policy) running in **user** context, triggered at logon, calling `New-CloudDriveMapping.ps1 -MappingsCsv <path> -RemoveExisting -Apply`
- Ship `mappings.csv` alongside the script (per-customer/per-group mapping lists)

**Notes**
- No credentials are stored or prompted — mapping success depends entirely on the signed-in user's existing tenant session
- A log file listing resolved WebDAV paths and mapping status is written after each run
