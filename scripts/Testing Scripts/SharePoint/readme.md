# Testing — SharePoint

Audit and reporting scripts for SharePoint Online via Microsoft Graph.

---

## Scripts

### Get-SharePointStorageReport.ps1

Reports storage usage across all SharePoint sites in a tenant (or a single site), including current file sizes and version history storage.

**Two modes**

| Mode | How | What it does |
|------|-----|-------------|
| Quick (default) | No `-Apply` | Reads quota data from the Graph API — fast, one row per library |
| Full scan | `-Apply` | Recursively enumerates all files + version history — accurate but slow for large tenants |

**Output files (saved to `C:\Temp\` / `~/Downloads\`)**

| File | Content |
|------|---------|
| `SharePoint_Summary_<timestamp>.csv` | One row per document library with totals |
| `SharePoint_Detail_<timestamp>.csv` | One row per file (full scan only) |

**Summary CSV columns (full scan)**

| Column | Description |
|--------|-------------|
| `SiteName` | Site display name |
| `SiteUrl` | Full site URL |
| `Library` | Document library name |
| `FileCount` | Number of files in the library |
| `VersionSizeMB` | Total storage used by version history |
| `TotalSizeMB` | Current files + version history combined |

**Detail CSV columns (full scan)**

| Column | Description |
|--------|-------------|
| `Path` | File path within the library |
| `SizeMB` | Current file size |
| `VersionCount` | Number of stored versions |
| `VersionSizeMB` | Storage used by version history for this file |
| `TotalSizeMB` | Current + version history |
| `Modified` | Last modified date |

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Apply` | No | Perform full recursive file scan (default: quota data only) |
| `-SiteUrl` | No | Scan a single site. If omitted, all sites in the tenant are scanned |
| `-SkipVersions` | No | Skip version history analysis (faster full scan) |
| `-OutputPath` | No | Override default output folder |
| `-TenantId` | No | Entra ID tenant ID or domain. Optional if already connected |

**Examples**

```powershell
# Quick summary — site quotas only
.\Get-SharePointStorageReport.ps1

# Full scan — all sites, all files, with version history
.\Get-SharePointStorageReport.ps1 -Apply

# Full scan — single site
.\Get-SharePointStorageReport.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Finance" -Apply

# Full scan — skip version history (faster)
.\Get-SharePointStorageReport.ps1 -Apply -SkipVersions
```

**Required scopes**
- `Sites.Read.All`
- `Files.Read.All`

**Required module**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

**Notes**
- Full scan with version history can take a long time on large tenants — run per site first to estimate duration
- Version history columns are empty in quick mode (quota data only)
- Re-running is safe — no changes are made to SharePoint
