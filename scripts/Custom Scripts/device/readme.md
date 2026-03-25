# Device Management Scripts

Scripts for managing and maintaining Windows endpoints. All scripts require administrator privileges.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Invoke-WindowsActivation.ps1`](#invoke-windowsactivationps1) | Activate Windows, manage product keys and KMS settings |
| [`Invoke-WindowsCleanup.ps1`](#invoke-windowscleanupps1) | Scan and remove reclaimable disk space |
| [`Time sync/Restart-Time-Sync.ps1`](#time-syncrestart-time-syncps1) | Fix Windows time sync by restarting W32tm and registering a scheduled task |
| [`audio/`](audio/readme.md) | Detect and disable the internal microphone on laptops |

---

## Invoke-WindowsActivation.ps1

Activate Windows or manage licensing settings from the command line.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-Status` | Show current activation status (WMI + `slmgr /dli`) |
| `-ProductKey` | Install a product key (retail or KMS generic) |
| `-KmsServer` | Set the KMS activation server address |
| `-KmsPort` | KMS server port (default: 1688) |
| `-Activate` | Trigger activation against Microsoft or configured KMS server |
| `-RemoveKey` | Remove the installed product key (before reimage / license transfer) |
| `-ReArm` | Reset the grace-period counter (max ~3-5x per Windows install) |
| `-Force` | Skip confirmation prompts |

**Examples**

```powershell
# Check current activation status
.\Invoke-WindowsActivation.ps1 -Status

# Install a retail key and activate online
.\Invoke-WindowsActivation.ps1 -ProductKey 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX' -Activate

# Point to a corporate KMS server and activate
.\Invoke-WindowsActivation.ps1 -KmsServer 'kms.company.local' -Activate

# Remove key before reimaging
.\Invoke-WindowsActivation.ps1 -RemoveKey -Force
```

---

## Invoke-WindowsCleanup.ps1

Scan and optionally remove reclaimable disk space. Runs as a dry-run by default — no files are deleted without `-Apply`.

**What it cleans**

| Category | Details |
|----------|---------|
| Temp files | User temp (all profiles) + `C:\Windows\Temp` |
| Windows Update | `SoftwareDistribution\Download` + DeliveryOptimization |
| Prefetch | `C:\Windows\Prefetch` |
| Memory dumps | System minidumps + per-user CrashDumps |
| WER | Windows Error Reporting queues (system + per-user) |
| Cache | Thumbnail cache, DirectX shader cache, font cache |
| Recycle Bin | All drives |
| Browser cache | Edge, Chrome (multi-profile), Firefox — all user profiles |
| Event logs | All Windows event logs |
| App & system logs | Dynamic scan of entire C:\ for `logs`/`log`/`logging` folders |
| DISM | Component store cleanup (`/StartComponentCleanup /ResetBase`) |

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-Apply` | Perform the actual cleanup (default: dry-run only) |
| `-SkipBrowserCache` | Skip browser cache cleanup |
| `-SkipAppLogs` | Skip application and system log cleanup |
| `-SkipEventLogs` | Skip Windows event log clearing |
| `-SkipDism` | Skip DISM component store cleanup (slow) |
| `-SkipRecycleBin` | Skip emptying the Recycle Bin |
| `-OutputPath` | Override report output folder (default: `C:\Temp\`) |

**Examples**

```powershell
# Dry run — see how much space can be freed
.\Invoke-WindowsCleanup.ps1

# Full cleanup
.\Invoke-WindowsCleanup.ps1 -Apply

# Cleanup, skip browser cache and DISM
.\Invoke-WindowsCleanup.ps1 -Apply -SkipBrowserCache -SkipDism
```

A CSV report with per-category results is saved to `C:\Temp\` after each run.

---

## Time sync/Restart-Time-Sync.ps1

Fixes Windows time synchronisation issues by:
1. Setting `W32tm` service startup type to Automatic and starting it
2. Configuring Dutch NTP pool servers (`0.nl.pool.ntp.org`, `1.nl.pool.ntp.org`)
3. Forcing an immediate resync
4. Registering a scheduled task that reruns the sync every 59 minutes

```powershell
.\Restart-Time-Sync.ps1
```

> Run once per device. The scheduled task ensures time stays in sync going forward.
