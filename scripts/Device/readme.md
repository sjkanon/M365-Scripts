# Device Management Scripts

Scripts for managing and maintaining Windows endpoints. All scripts require administrator privileges.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Clear-TempFiles.ps1`](#clear-tempfilesps1) | Clear the shared script temp folder (`C:\Temp` on Windows, `/tmp` on Linux/macOS) |
| [`Invoke-WindowsActivation.ps1`](#invoke-windowsactivationps1) | Activate Windows, manage product keys and KMS settings |
| [`Invoke-WindowsCleanup.ps1`](#invoke-windowscleanupps1) | Scan and remove reclaimable disk space |
| [`Test-OpenVpnDiagnostics.ps1`](#test-openvpndiagnosticsps1) | Diagnose OpenVPN Connect issues |
| [`Time sync/`](Time%20sync/readme.md) | Fix Windows time sync by restarting W32tm and registering a scheduled task |
| [`audio/`](audio/readme.md) | Detect and disable the internal microphone on laptops |

---

### Clear-TempFiles.ps1

Clears only the shared script temp folder.
Default target path is `C:\Temp` on Windows and `/tmp` on Linux/macOS.
Runs in dry-run mode by default and only removes files when `-Apply` is provided.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-Apply` | Perform actual deletion (default is dry-run) |
| `-TempPath` | Override target temp path |
| `-OlderThanDays` | Only target items older than N days (default: `1`) |

**Examples**

```powershell
# Dry run
.\Clear-TempFiles.ps1

# Delete only temp files older than 7 days
.\Clear-TempFiles.ps1 -Apply -OlderThanDays 7

# Override temp path (example Linux/macOS)
.\Clear-TempFiles.ps1 -Apply -TempPath /var/tmp
```

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

## Test-OpenVpnDiagnostics.ps1

Collects and evaluates diagnostic information for OpenVPN Connect issues on a Windows machine. Checks each relevant layer from driver to network and reports any problems found.

**Checks performed**

| Section | What is checked |
|---------|----------------|
| Wintun / TAP adapters | PnP device status — flags anything not `OK` |
| Virtual network adapters | Adapter visibility — warns if none found while VPN should be active |
| Network profiles | Flags VPN adapters set to `Public` (should be `Private`) |
| Installed VPN software | Lists all VPN-related apps — warns about potential conflicts with OpenVPN Connect |
| Hyper-V / WSL / virtualisation | Lists enabled features — warns if Hyper-V is active (can conflict with Wintun) |
| OpenVPN service | Service status and start type — flags if not running |
| Active routes | Routes via VPN adapter — warns if adapter exists but no routes are present |
| DNS configuration | DNS servers per active adapter |
| Event Log | Last 20 OpenVPN entries from the Application log |

Results are printed to screen with a summary of all issues at the end.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-ExportTxt` | No | Save the full report to a txt file |
| `-OutputPath` | No | Custom path for the report (implies `-ExportTxt`). Default: `C:\Temp\OpenVpnDiagnostics_<timestamp>.txt` |

**Examples**

```powershell
# Run diagnostics, output to screen only
.\Test-OpenVpnDiagnostics.ps1

# Run and save report to C:\Temp\
.\Test-OpenVpnDiagnostics.ps1 -ExportTxt

# Save to a custom path
.\Test-OpenVpnDiagnostics.ps1 -OutputPath "C:\Support\vpn-report.txt"
```

---

## Time sync/

Fixes Windows time synchronisation issues by restarting `W32tm` against Dutch NTP pool servers and registering a scheduled task that reruns the sync every 59 minutes. See [`Time sync/readme.md`](Time%20sync/readme.md) for full details.
