#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Intune Remediation fix script — clears Win32 apps stuck behind Intune's GRS retry cooldown
    on this device.

.DESCRIPTION
    Companion remediation script for Detect-StuckWin32AppEnforcement.ps1 — deploy both together
    as an Intune "Scripts and remediations > Remediations" package so a device that gets stuck
    (3 failed Win32 app install attempts -> 24h local GRS cooldown, independent of whatever gets
    fixed afterwards in Intune itself) clears and resyncs itself automatically on the next
    scheduled run, with no manual per-device PowerShell session needed.

    Same clearing logic as the standalone Repair-StuckWin32AppEnforcement.ps1 run with
    -Apply -ForceSync, minus the interactive dry-run/report options that only make sense for a
    manually-run, one-off troubleshooting session — this script always clears everything it
    finds and always forces an immediate resync, since that is unconditionally the right thing
    to do once Intune's Remediations engine has already decided (via the paired detection
    script) that this device is non-compliant.

    For each Win32 app with a cached EnforcementStateMessage carrying a real error code (not 0 =
    success, not 3010 = soft reboot pending), removes:
        HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\<UserSID>\<AppId>
        HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\Reporting\<UserSID>\<AppId>
        HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\<UserSID>\GRS\<LastHashValue>
    then restarts the IntuneManagementExtension service and triggers an immediate MDM check-in
    (same PushLaunch scheduled task Company Portal's "Sync" button uses), so every Win32 app is
    re-evaluated from scratch without waiting for the next automatic check-in interval.

.NOTES
    Deploy as the "Remediation script" half of an Intune Remediation, paired with
    Detect-StuckWin32AppEnforcement.ps1. Run using SYSTEM (not logged-on credentials), 64-bit
    PowerShell, no signature check.

    Registry structure and remediation approach based on community documentation of Intune's GRS
    (retry schedule) mechanism — Microsoft does not publicly document this internal IME state:
      https://www.anoopcnair.com/override-grs-trigger-ime-to-retry-failed-win32/
      https://call4cloud.nl/retry-failed-win32app-installation/
      https://msnugget.com/retry-failed-win32-apps-on-demand-with-intune-remediations/
#>

$ErrorActionPreference = 'Stop'

function Get-FailedWin32AppStates {
    $win32AppsKeyPath = 'HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps'
    if (-not (Test-Path $win32AppsKeyPath)) { return @() }

    $results = [System.Collections.Generic.List[object]]::new()
    $subKeys = Get-ChildItem -Path $win32AppsKeyPath -Recurse -ErrorAction SilentlyContinue

    foreach ($subKey in $subKeys) {
        $prop = Get-ItemProperty -Path $subKey.PSPath -Name EnforcementStateMessage -ErrorAction SilentlyContinue
        if (-not $prop) { continue }
        if ($prop.EnforcementStateMessage -notmatch '"ErrorCode":(-?\d+|null)') { continue }
        $errorCodeRaw = $Matches[1]
        if ($errorCodeRaw -eq 'null') { continue }
        $errorCode = [int]$errorCodeRaw
        if ($errorCode -eq 0 -or $errorCode -eq 3010) { continue }

        $relativePath = $subKey.PSPath -replace '^Microsoft\.PowerShell\.Core\\Registry::', ''
        if (-not ($relativePath -match 'Win32Apps\\([^\\]+)\\([^\\]+)')) { continue }
        $foundAppId = $Matches[2]
        if ($foundAppId -eq 'Reporting') { continue }

        $results.Add([PSCustomObject]@{
            UserObjectId = $Matches[1]
            AppId        = $foundAppId
            ErrorCode    = $errorCode
        })
    }

    return @($results | Sort-Object UserObjectId, AppId -Unique)
}

function Get-LastHashValue {
    param([string]$UserObjectId, [string]$TargetAppId)
    $reportingKeyPath = "HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\Reporting\$UserObjectId\$TargetAppId\ReportCache\$UserObjectId"
    if (-not (Test-Path $reportingKeyPath)) { return $null }
    $reportingKey = Get-ItemProperty -Path $reportingKeyPath -Name LastHashValue -ErrorAction SilentlyContinue
    return $reportingKey.LastHashValue
}

function Remove-StuckAppState {
    param([string]$UserObjectId, [string]$TargetAppId, [string]$LastHashValue)
    $pathsToRemove = @(
        "HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\$UserObjectId\$TargetAppId",
        "HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\Reporting\$UserObjectId\$TargetAppId"
    )
    if ($LastHashValue) {
        $pathsToRemove += "HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\$UserObjectId\GRS\$LastHashValue"
    }

    foreach ($path in $pathsToRemove) {
        if (Test-Path $path) {
            Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
            Write-Host "Removed: $path"
        }
    }
}

function Invoke-ForceMdmSync {
    # Zelfde mechanisme als de "Sync"-knop in Company Portal / Instellingen > Werk of schoolaccount.
    $pushLaunchTasks = Get-ScheduledTask -TaskName 'PushLaunch' -ErrorAction SilentlyContinue |
        Where-Object { $_.TaskPath -like '*Microsoft\Windows\EnterpriseMgmt*' }
    if (-not $pushLaunchTasks) {
        Write-Host "WARN: no PushLaunch scheduled task found -- device may not be MDM-enrolled."
        return
    }
    foreach ($task in $pushLaunchTasks) {
        Start-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath
        Write-Host "Sync triggered via task: $($task.TaskPath)$($task.TaskName)"
    }
}

try {
    $failedStates = Get-FailedWin32AppStates

    if ($failedStates.Count -eq 0) {
        Write-Host "Nothing to remediate -- no Win32 apps with a stuck/failed enforcement state found."
        exit 0
    }

    foreach ($state in $failedStates) {
        Write-Host "Clearing app $($state.AppId) (user: $($state.UserObjectId)) -- last ErrorCode: $($state.ErrorCode)"
        $lastHashValue = Get-LastHashValue -UserObjectId $state.UserObjectId -TargetAppId $state.AppId
        Remove-StuckAppState -UserObjectId $state.UserObjectId -TargetAppId $state.AppId -LastHashValue $lastHashValue
    }

    Write-Host "Restarting IntuneManagementExtension service..."
    Restart-Service -Name 'IntuneManagementExtension' -Force
    Write-Host "Forcing immediate MDM sync..."
    Invoke-ForceMdmSync

    Write-Host "Remediation complete: $($failedStates.Count) app(s) cleared and re-queued for evaluation."
    exit 0
} catch {
    Write-Host "ERROR: remediation failed: $_"
    exit 1
}
