# Device Management Scripts

Scripts for managing and maintaining Windows endpoints. All scripts require administrator privileges.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Clear-TempFiles.ps1`](#clear-tempfilesps1) | Clear the shared script temp folder (`C:\Temp` on Windows, `/tmp` on Linux/macOS) |
| [`Invoke-WindowsActivation.ps1`](#invoke-windowsactivationps1) | Activate Windows, manage product keys and KMS settings |
| [`Invoke-WindowsCleanup.ps1`](#invoke-windowscleanupps1) | Scan and remove reclaimable disk space |
| [`Remove-OemBloatware.ps1`](#remove-oembloatwareps1) | Remove OEM (HP/Lenovo/Dell) and generic Microsoft Store bloatware |
| [`Test-OpenVpnDiagnostics.ps1`](#test-openvpndiagnosticsps1) | Diagnose OpenVPN Connect issues |
| [`Update-TeamsClient.ps1`](#update-teamsclientps1) | Update new Teams + Outlook meeting add-in, only when Microsoft published a newer build ([how it works](Update-TeamsClient.md)) |
| [`Time sync/`](Time%20sync/readme.md) | Fix Windows time sync by restarting W32tm and registering a scheduled task |
| [`audio/`](audio/readme.md) | Detect and disable the internal microphone on laptops |
| [`DriveMapping/`](DriveMapping/readme.md) | Map SharePoint/OneDrive document libraries to drive letters at logon |

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

## Remove-OemBloatware.ps1

Detects the device manufacturer and removes known OEM bloatware via `winget`, plus a generic list of consumer Microsoft Store apps (Xbox, Solitaire, Bing News/Weather, Cortana, Clipchamp, etc.) via `Remove-AppxPackage`. Runs as a dry-run by default — no app is removed without `-Apply`.

> The bloatware lists are a starting point, not exhaustive — package IDs vary by OEM preload image and change over time. Run `winget list` / `Get-AppxPackage | Select Name` on a representative device first and adjust the lists in the script if needed.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-Apply` | Actually remove matched apps (default: dry-run only) |
| `-SkipOem` | Skip manufacturer-specific removal, only process the generic Microsoft Store list |
| `-SkipAppx` | Skip the generic Microsoft Store list, only process manufacturer-specific apps |
| `-Manufacturer` | Override auto-detection (`HP`, `Lenovo`, `Dell`) |
| `-OutputPath` | Override report output folder (default: `C:\Temp\`) |

**Examples**

```powershell
# Dry run — see what would be removed on this device
.\Remove-OemBloatware.ps1

# Actually remove OEM + generic bloatware
.\Remove-OemBloatware.ps1 -Apply

# Only remove generic Microsoft Store junk, leave OEM apps alone
.\Remove-OemBloatware.ps1 -Apply -SkipOem
```

A CSV report (found/removed per app) is saved to `C:\Temp\` after each run.

**Requires:** `winget` (App Installer from the Microsoft Store) for OEM package removal; Administrator privileges.

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

## Update-TeamsClient.ps1

> Full reference: [Update-TeamsClient.md](Update-TeamsClient.md) — decision flow, version check, design decisions and troubleshooting.

Keeps the new Teams client and the Outlook meeting add-in current on an endpoint or AVD session host. It asks the Teams config service which build Microsoft publishes for this architecture and **only acts when that build is newer than what is installed** — an up-to-date device is left completely alone. When an update is due it downloads and signature-checks `teamsbootstrapper.exe`, uninstalls the meeting add-in, removes and deprovisions the `MSTeams` AppX package, provisions the new build for all users and reinstalls the add-in MSI that ships inside it.

Every state-changing step goes through `ShouldProcess`, so `-WhatIf` walks the full flow without touching the machine. Runs by hand (it elevates itself via UAC and asks for confirmation once) and unattended from an RMM such as NinjaOne.

**Steps**

| # | Step | Honours `-WhatIf` |
|---|------|-------------------|
| 1 | Preflight — installed package, add-in, running Teams/Outlook | read-only |
| 2 | Version check — published build vs installed build | read-only |
| 3 | Create working folder, download bootstrapper, verify Microsoft signature | yes |
| 4 | Uninstall add-in, remove `MSTeams` AppX for all users, deprovision it | yes |
| 5 | Provision new Teams (`teamsbootstrapper.exe -p`) | yes |
| 6 | Install Teams Meeting Add-in MSI (`ALLUSERS=1`) | yes |
| 7 | Verify add-in registration + provisioned package | reported as skipped under `-WhatIf` |

If the client is current but only the meeting add-in is missing, steps 3–5 are skipped and just the add-in is installed.

**Why the order matters:** nothing is touched until a newer build is confirmed, and the installer is fetched and verified *before* the first uninstall — so a failed download or a blocked URL can never leave the device without a Teams client.

**Version check**

`https://config.teams.microsoft.com/config/v1/MicrosoftTeams/...` is the feed the Teams client itself uses to decide it is out of date. It returns the current build per architecture (`BuildSettings.WebView2PreAuth.<arch>.latestVersion`). An installed build equal to or newer than that means there is nothing to do. If the service cannot be reached the run stops instead of reinstalling blindly — `-Force` overrides that. `-Ring` selects a different update ring (default `general`).

> The add-in uninstall and verification read both the 64-bit and the `WOW6432Node` uninstall hive — the add-in installs 32-bit, so the 64-bit hive alone misses it. The add-in MSI version comes from the MSI property table (`WindowsInstaller.Installer` COM), not from `Get-AppLockerFileInformation`, which is missing on some editions and breaks under PowerShell 7.

**Safety**

- `msiexec` and the bootstrapper run with a timeout (`-TimeoutSeconds`, default 900) and are killed if they hang, so an RMM job cannot block the agent.
- MSI exit code `1618` (another install in progress) is retried twice; `3010` counts as success and flags a pending reboot in the summary.
- Unexpected errors abort the run instead of continuing half-way.
- A run that actually changes something writes a transcript to `C:\Temp\Update-TeamsClient_<timestamp>.log`; a check that finds nothing to do leaves no log litter behind.

**Exit codes**

| Code | Meaning |
|------|---------|
| `0` | Success, or already up to date |
| `1` | Failure |
| `2` | `-CheckOnly` only: a newer build is available |

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-WhatIf` | Show what an update would do without performing it |
| `-Quiet` | Print nothing unless there is news: a newer build, an action, or a failure |
| `-CheckOnly` | Only report whether a newer build exists (exit code 2), change nothing |
| `-Confirm:$false` | Never ask for confirmation (use this for unattended runs) |
| `-Ring` | Update ring queried at the config service (default: `general`) |
| `-WorkingDir` | Bootstrapper download folder (default: `C:\IT\AVD\Teams`) |
| `-LogPath` | Transcript folder (default: `C:\Temp`) |
| `-BootstrapperUrl` | Override the `teamsbootstrapper.exe` download URL (https only) |
| `-SkipMeetingAddIn` | Leave the meeting add-in alone, and do not treat a missing add-in as work |
| `-SkipSignatureCheck` | Accept an installer not signed by Microsoft (internal mirror) |
| `-TimeoutSeconds` | Per-process timeout for msiexec/bootstrapper (default: `900`) |
| `-Force` | Reinstall even when Teams is current, and continue without Teams or version info |

**Examples**

```powershell
# Dry run — check for a newer build and show what an update would do
.\Update-TeamsClient.ps1 -WhatIf

# Update only if Microsoft published a newer build
.\Update-TeamsClient.ps1

# Scheduled RMM run: silent unless there is a newer build or a problem
.\Update-TeamsClient.ps1 -Quiet -Confirm:$false

# Detection only: exit code 2 when an update is available
.\Update-TeamsClient.ps1 -CheckOnly -Quiet

# Repair: reinstall the current build regardless of the version check
.\Update-TeamsClient.ps1 -Force
```

**Running it from NinjaOne**

1. Add the script (Language: PowerShell, Operating System: Windows, Architecture: **All**, Run As: **System**).
2. Preview a device first: run it with `-WhatIf -Confirm:$false` in the *Parameters* field — the job output shows the version comparison and every step an update would perform, and the device stays untouched.
3. Schedule the real run with `-Quiet -Confirm:$false`. On an up-to-date device it prints nothing and exits `0`, so the activity feed only shows the devices where it actually did something.
4. For a detection/condition job use `-CheckOnly -Quiet`: silent and `0` when current, output and exit code `2` when a newer build is published.
5. Optional script variables (checkboxes `whatIf`, `quiet`, `checkOnly`, `force`, `skipMeetingAddIn`, `skipSignatureCheck`; text fields `workingDir`, `logPath`, `ring`) are picked up from the environment when the matching parameter is not passed, so a technician can tick *whatIf* instead of typing parameters.

If the agent starts PowerShell 32-bit, the script relaunches itself 64-bit via `SysNative` first — without that, the registry reads are redirected to `WOW6432Node` and `$env:ProgramFiles` points at the x86 folder, so neither the AppX package nor the add-in MSI is found.

---

## Time sync/

Fixes Windows time synchronisation issues by restarting `W32tm` against Dutch NTP pool servers and registering a scheduled task that reruns the sync every 59 minutes. See [`Time sync/readme.md`](Time%20sync/readme.md) for full details.


