# OneDriveManagement

Device-side OneDrive for Business maintenance scripts: a restart/reset watchdog, per-library sync teardown, and Known Folder Move redirection. Distinct from [`scripts/Device/DriveMapping/`](../../Device/DriveMapping/readme.md) (SharePoint/OneDrive drive-letter mapping via WebDAV), which is a different capability.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Register-OneDriveWatchdog.ps1`](#register-onedrivewatchdogps1) | Keep OneDrive running via a scheduled task; also supports a one-time `/reset` |
| [`Stop-OneDriveLibrarySync.ps1`](#stop-onedrivelibrarysyncps1) | Stop syncing one specific library without touching any other |
| [`Set-OneDriveKnownFolderRedirect.ps1`](#set-onedriveknownfolderredirectps1) | Redirect Desktop/Documents/Pictures/Downloads into OneDrive |

---

### Register-OneDriveWatchdog.ps1

Registers a scheduled task (runs as the current user, default every 59 minutes) that relaunches OneDrive if it isn't running, or optionally force-restarts it every run. `-Reset` instead performs a one-time `onedrive.exe /reset` for a stuck sync profile.

| Parameter | Description |
|-----------|-------------|
| `-IntervalMinutes` | Watchdog interval (default: `59`) |
| `-ForceRestartOnEachRun` | Stop and relaunch OneDrive on every trigger, not just when not running |
| `-Reset` | Run a one-time `/reset` instead of registering the watchdog |
| `-Apply` | Actually register/reset (default: preview only) |

```powershell
.\Register-OneDriveWatchdog.ps1 -Apply
.\Register-OneDriveWatchdog.ps1 -ForceRestartOnEachRun -IntervalMinutes 30 -Apply
.\Register-OneDriveWatchdog.ps1 -Reset -Apply
```

---

### Stop-OneDriveLibrarySync.ps1

Cleanly stops syncing one specific SharePoint/OneDrive library — shuts down OneDrive, removes that library's policy file and mount-point cache entry, edits the sync database ini, removes the local folder, and relaunches OneDrive, all without disturbing any other synced library.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-MatchPattern` | Yes | Distinctive string identifying the library's policy file |
| `-LocalFolderPath` | Yes | Local synced folder to remove |
| `-Apply` | No | Actually perform the teardown (default: preview only) |

```powershell
.\Stop-OneDriveLibrarySync.ps1 -MatchPattern "ProjectX" -LocalFolderPath "$env:USERPROFILE\Contoso\ProjectX - Documents" -Apply
```

> Operates on OneDrive's undocumented internal state files — test on one machine before wider rollout.

---

### Set-OneDriveKnownFolderRedirect.ps1

Redirects known folders into the OneDrive for Business sync root using `SHSetKnownFolderPath`, migrating existing content with Robocopy and hiding (not deleting) the original folder. Optionally enables OneDrive's `Timerautomount` flag so previously-synced libraries auto-mount on a new device.

| Parameter | Description |
|-----------|-------------|
| `-Folders` | Which folders to redirect (default: Desktop, Documents, Pictures, Downloads) |
| `-EnableAutoMountSharedLibraries` | Also enable `Timerautomount` |
| `-Apply` | Actually redirect (default: preview only) |

```powershell
.\Set-OneDriveKnownFolderRedirect.ps1 -Apply
.\Set-OneDriveKnownFolderRedirect.ps1 -Folders Desktop,Documents -EnableAutoMountSharedLibraries -Apply
```

> Run in the signed-in user's context, not SYSTEM — it reads that user's OneDrive account registry key.

---

All scripts default to a dry run; pass `-Apply` to make changes, per house style.
