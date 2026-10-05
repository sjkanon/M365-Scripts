**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [Intune](../readme.md) › **DiskCleanup**

# DiskCleanup

Intune Win32-app deployment that runs [`Invoke-WindowsCleanup.ps1`](../../Device/readme.md#invoke-windowscleanupps1) (`scripts/Device/`) on the C:\ drive as SYSTEM and reboots the device afterwards. Built as a thin wrapper so all cleanup logic stays in one place — nothing here duplicates it.

## Contents

| Script | Role in Intune |
|---|---|
| [`Invoke-DiskCleanupIntune.ps1`](Invoke-DiskCleanupIntune.ps1) ([docs](#invoke-diskcleanupintuneps1)) | Install command content script — calls the shared `Invoke-WindowsCleanup.ps1 -Apply`, then `Restart-Computer -Force` |
| [`Detect-DiskCleanupIntune.ps1`](Detect-DiskCleanupIntune.ps1) ([docs](#detect-diskcleanupintuneps1)) | Custom detection script |

### Invoke-DiskCleanupIntune.ps1

Win32-app install command. Runs `Invoke-WindowsCleanup.ps1 -Apply` from the same content folder, stamps `HKLM:\SOFTWARE\DiskCleanupDeploy\LastRunUtc` on success and then forces a restart.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-SkipDism` | No | Skip the DISM component store cleanup (`/StartComponentCleanup /ResetBase`) for a shorter, predictable runtime |
| `-NoRestart` | No | Clean up but do not restart — for manual testing outside Intune only, leave it out of the real install command |

**Examples**

```powershell
# Intune install command
%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File Invoke-DiskCleanupIntune.ps1

# Manual test run without DISM and without the reboot
.\Invoke-DiskCleanupIntune.ps1 -SkipDism -NoRestart
```

### Detect-DiskCleanupIntune.ps1

Win32-app custom detection script. Reports "installed" (exit 0) while the `LastRunUtc` stamp is at most `$MaxAgeDays` (30) days old, and "not installed" (exit 1) once it is older or missing, so Intune runs the cleanup again. Intune passes no parameters to detection scripts: edit `$MaxAgeDays` in the script before packaging to change the cycle.

## Behavior

- **Reboot**: immediate and forced (`Restart-Computer -Force`) right after cleanup — no user warning/countdown. Make sure the assignment/notification informs end users in advance.
- **DISM component store cleanup**: included by default (biggest space win, but can take tens of minutes). Pass `-SkipDism` in the install command if you need a predictable short runtime, and raise the Win32 app's install timeout (default 60 min) if you keep it.
- **Recurring by design**: the install script stamps `HKLM:\SOFTWARE\DiskCleanupDeploy\LastRunUtc` on success. The detection script reports "not installed" once that stamp is older than `$MaxAgeDays` (30 by default, edit the constant in `Detect-DiskCleanupIntune.ps1` before packaging to change it), so Intune re-runs the cleanup on its own every cycle — no monthly content bump needed like the ClaudeDesktop app.
- Logs to `%ProgramData%\DiskCleanupDeploy\cleanup.log`, including the full output and CSV report from `Invoke-WindowsCleanup.ps1`.

## Packaging as a Win32 app

Both files here need `Invoke-WindowsCleanup.ps1` copied alongside them before packaging — Win32-app content is a flat folder, so the wrapper resolves it via `$PSScriptRoot` at install time.

```powershell
Install-Module IntuneWin32App -Scope CurrentUser   # if not already installed
Connect-MgGraph -Scopes "DeviceManagementApps.ReadWrite.All"

$source = "C:\Temp\DiskCleanupContent"
New-Item -ItemType Directory -Path $source -Force | Out-Null
Copy-Item ".\Invoke-DiskCleanupIntune.ps1" $source
Copy-Item "..\..\Device\Invoke-WindowsCleanup.ps1" $source

$package = New-IntuneWin32AppPackage -SourceFolder $source -SetupFile "Invoke-DiskCleanupIntune.ps1" -OutputFolder "C:\Temp\DiskCleanupOutput"

$detection = New-IntuneWin32AppDetectionRuleScript -ScriptFile ".\Detect-DiskCleanupIntune.ps1"

New-IntuneWin32App -FilePath $package.Path `
    -DisplayName "Disk Cleanup (C:)" `
    -Publisher "IT" `
    -InstallCommandLine "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File Invoke-DiskCleanupIntune.ps1" `
    -UninstallCommandLine "cmd.exe /c echo not applicable" `
    -InstallExperience "system" `
    -RestartBehavior "suppress" `
    -DetectionRule $detection
```

Assign it **Required** to the target device group. `-RestartBehavior "suppress"` tells Intune not to add its own reboot prompt on top — the script already forces one.

### Prerequisites

- `Microsoft.Graph.Authentication`, `IntuneWin32App` PowerShell modules
- Run from Windows (packaging is Windows-only)
