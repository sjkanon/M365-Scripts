# Intune

Autopilot enrollment, compliance policy automation, configuration drift detection, and customer desktop deployment (wallpaper, lockscreen, taskbar lock shortcut).

> Office theme/color deployment lives in [`Custom Scripts/Intune/Desktop/`](../Custom%20Scripts/Intune/readme.md) — those scripts hardcode their download URL to that path.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Compare-IntuneConfig.ps1`](#compare-intuneconfigps1) | Compare a customer tenant's Intune configuration against an MSP baseline backup |

## Folders

| Folder | Description |
|--------|-------------|
| [`Get-Autopilot/`](Get-Autopilot/readme.md) | Windows Autopilot hardware hash collection |
| [`iOS-Compliance-Updater/`](iOS-Compliance-Updater/readme.md) | Auto-updates the minimum iOS version in an Intune compliance policy |
| [`Desktop/`](Desktop/readme.md) | Corporate wallpaper + lockscreen, taskbar lock shortcut |
| [`DiskCleanup/`](DiskCleanup/readme.md) | Win32-app wrapper that runs the C:\ disk cleanup script and reboots the device |

---

### Compare-IntuneConfig.ps1

Wraps the community `IntuneBackupAndRestore` module to detect Intune configuration drift — compliance policies, configuration profiles, and other exported objects that differ from (or are missing/added relative to) a reference "baseline" backup. Read-only — never changes configuration, only reports differences.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-BaselinePath` | Yes | MSP reference backup folder (from a previous `Start-IntuneBackup -Path <path>` run) |
| `-CustomerBackupPath` | No | Existing backup of the customer tenant. If omitted, the script backs up the currently connected tenant first |
| `-TenantId` | No | Tenant ID or domain to connect to (only used when `-CustomerBackupPath` is omitted) |
| `-OutputPath` | No | Folder for the auto-backup and diff report (default: `C:\Temp\` / `~/Downloads`) |

**Examples**

```powershell
# Compare an existing customer backup against the MSP baseline
.\Compare-IntuneConfig.ps1 -BaselinePath C:\IntuneBaseline -CustomerBackupPath C:\Temp\CustomerBackup

# Back up the currently connected (GDAP) customer tenant and compare it live
.\Compare-IntuneConfig.ps1 -BaselinePath C:\IntuneBaseline
```

**Notes**
- GDAP-aware: resolves `-TenantId` automatically from the selected customer tenant (`$global:cid`) if omitted, same as `Move-InboxToArchive.ps1` / `Get-SharePointStorageReport.ps1`
- Not wired into `menu.ps1` — works with backup folder paths and a tenant-wide export, run it directly

**Required module**

```powershell
Install-Module IntuneBackupAndRestore -Scope CurrentUser
```
