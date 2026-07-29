#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Detect and clear Win32 apps stuck behind Intune's GRS retry cooldown on this device.

.DESCRIPTION
    The Intune Management Extension (IME) retries a failing Win32 app install 3 times, 5 minutes
    apart. After the 3rd failure it locks the app into a 24-hour "GRS" (retry schedule) cooldown —
    visible in AppActionProcessor.log as "... to install is in GRS. The app will not be enforced."
    or "... has previously been selected for enforcement. The app will not be enforced." During
    that cooldown IME will NOT retry the app again, even after you fix the underlying problem
    (corrected requirement rule, new package, fixed install command) — the cooldown is local
    per-device state, unrelated to what's configured in Intune itself.

    This script finds every locally cached Win32 app whose last recorded EnforcementStateMessage
    has a real error code (not 0 = success, not 3010 = soft reboot pending) under:
        HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\<UserSID>\<AppId>
    For each match it can remove that app's cached enforcement state, its reporting cache, and the
    matching GRS cooldown key (resolved via the reporting cache's LastHashValue) under:
        HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\Reporting\<UserSID>\<AppId>
        HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\<UserSID>\GRS\<LastHashValue>
    then restarts the IntuneManagementExtension service so it re-evaluates every app from scratch
    on the next check-in — no more waiting out the 24h window.

    Run without -Apply for a dry run — reports which apps are found stuck and would be cleared.
    Run with -Apply to actually clear them.

.PARAMETER Apply
    Actually remove the stuck registry state and restart the IntuneManagementExtension service.
    Without this switch, only reports what would be cleared.

.PARAMETER AppId
    Optional. Limit to one specific Win32 app's GUID (as shown in the Intune admin center app URL,
    or in AppActionProcessor.log). Default: every app found with a non-zero/non-3010 error code.

.PARAMETER ForceSync
    After clearing state, also trigger an immediate MDM check-in via the same scheduled task
    Company Portal's "Sync" button uses (PushLaunch), instead of waiting for the next automatic
    check-in interval. Only meaningful together with -Apply.

.PARAMETER OutputPath
    CSV report of found (and, with -Apply, cleared) app states. Default: C:\Temp\.

.EXAMPLE
    # Dry run — see what's currently stuck on this device
    .\Repair-StuckWin32AppEnforcement.ps1

.EXAMPLE
    # Clear every stuck app and force an immediate resync
    .\Repair-StuckWin32AppEnforcement.ps1 -Apply -ForceSync

.EXAMPLE
    # Only clear one specific app (Claude Desktop's app id, for example)
    .\Repair-StuckWin32AppEnforcement.ps1 -Apply -AppId "96ff358d-0e16-4224-b946-eb61bc930fca"

.NOTES
    Registry structure and remediation approach based on community documentation of Intune's GRS
    (retry schedule) mechanism — Microsoft does not publicly document this internal IME state:
      https://www.anoopcnair.com/override-grs-trigger-ime-to-retry-failed-win32/
      https://call4cloud.nl/retry-failed-win32app-installation/
      https://msnugget.com/retry-failed-win32-apps-on-demand-with-intune-remediations/
    Report saved to the output path; run as Administrator (reads/writes HKLM and restarts a service).
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch] $Apply,
    [string] $AppId,
    [switch] $ForceSync,
    [string] $OutputPath = 'C:\Temp'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$reportCsv = Join-Path $OutputPath "StuckWin32AppEnforcement_$ts.csv"

Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host '   Repair-StuckWin32AppEnforcement' -ForegroundColor Cyan
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host ''

if (-not $Apply) {
    Write-Host '  ================================================' -ForegroundColor Yellow
    Write-Host '   DRY RUN — no registry keys will be removed, no service restart.' -ForegroundColor Yellow
    Write-Host '   Add -Apply to actually clear the GRS cooldown.' -ForegroundColor Yellow
    Write-Host '  ================================================' -ForegroundColor Yellow
    Write-Host ''
}

# ── Helpers ────────────────────────────────────────────────────────────────────
function Get-UsernameFromSid {
    param([string]$Sid)
    try {
        return ([System.Security.Principal.SecurityIdentifier]$Sid).Translate([System.Security.Principal.NTAccount]).Value
    } catch {
        return $null
    }
}

function Get-FailedWin32AppStates {
    # Alleen "echte" fouten (niet 0 = geslaagd, niet 3010 = herstart in afwachting) — dat zijn
    # precies de apps die na 3 mislukte pogingen 24u in GRS-cooldown zitten.
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

        # SID en App-ID zijn de twee padsegmenten direct na "Win32Apps\" — regex i.p.v. een vaste
        # index, zodat dit standhoudt ook als IME een extra subkey-niveau toevoegt/verwijdert.
        $relativePath = $subKey.PSPath -replace '^Microsoft\.PowerShell\.Core\\Registry::', ''
        if ($relativePath -notmatch 'Win32Apps\\([^\\]+)\\([^\\]+)') { continue }
        $userObjectId = $Matches[1]
        $foundAppId   = $Matches[2]
        if ($foundAppId -eq 'Reporting') { continue }
        if ($AppId -and $foundAppId -ne $AppId) { continue }

        $results.Add([PSCustomObject]@{
            UserObjectId = $userObjectId
            UserName     = Get-UsernameFromSid -Sid $userObjectId
            AppId        = $foundAppId
            ErrorCode    = $errorCode
            SubKeyPath   = $subKey.PSPath
        })
    }

    # Eén entry per (user, app) — meerdere subkeys onder dezelfde app kunnen matchen.
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
            if ($PSCmdlet.ShouldProcess($path, 'Remove registry key')) {
                Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                Write-Host "    [DONE] Verwijderd: $path" -ForegroundColor Green
            }
        } else {
            Write-Host "    [SKIP] Niet gevonden (al schoon): $path" -ForegroundColor DarkGray
        }
    }
}

function Invoke-ForceMdmSync {
    # Zelfde mechanisme als de "Sync"-knop in Company Portal / Instellingen > Werk of schoolaccount.
    $pushLaunchTasks = Get-ScheduledTask -TaskName 'PushLaunch' -ErrorAction SilentlyContinue |
        Where-Object { $_.TaskPath -like '*Microsoft\Windows\EnterpriseMgmt*' }
    if (-not $pushLaunchTasks) {
        Write-Host '    [WARN] Geen PushLaunch scheduled task gevonden — apparaat is mogelijk niet MDM-ingeschreven.' -ForegroundColor Yellow
        return
    }
    foreach ($task in $pushLaunchTasks) {
        Start-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath
        Write-Host "    [DONE] Sync getriggerd via taak: $($task.TaskPath)$($task.TaskName)" -ForegroundColor Green
    }
}

# ── Scan ───────────────────────────────────────────────────────────────────────
Write-Host '  Gecachte Win32-app enforcement-status doorzoeken op mislukte pogingen...' -ForegroundColor Cyan
$failedStates = Get-FailedWin32AppStates

if ($failedStates.Count -eq 0) {
    Write-Host ''
    Write-Host '  Geen apps met een vastgelegde installatiefout gevonden — niets om op te ruimen.' -ForegroundColor Green
    Write-Host ''
    exit 0
}

Write-Host ("  Gevonden: {0} app(s) met een mislukte laatste poging:" -f $failedStates.Count) -ForegroundColor Yellow
Write-Host ''

$reportRows = [System.Collections.Generic.List[object]]::new()

foreach ($state in $failedStates) {
    $userLabel = if ($state.UserName) { $state.UserName } else { $state.UserObjectId }
    Write-Host ("  App {0} (user: {1}) — laatste ErrorCode: {2}" -f $state.AppId, $userLabel, $state.ErrorCode) -ForegroundColor Yellow

    $lastHashValue = Get-LastHashValue -UserObjectId $state.UserObjectId -TargetAppId $state.AppId

    if ($Apply) {
        Remove-StuckAppState -UserObjectId $state.UserObjectId -TargetAppId $state.AppId -LastHashValue $lastHashValue
    }

    $reportRows.Add([PSCustomObject]@{
        AppId         = $state.AppId
        UserSid       = $state.UserObjectId
        UserName      = $state.UserName
        ErrorCode     = $state.ErrorCode
        GrsHash       = $lastHashValue
        Status        = if ($Apply) { 'Cleared' } else { 'Found (dry run)' }
    })
}

$reportRows | Export-Csv -Path $reportCsv -NoTypeInformation -Encoding UTF8
Write-Host ''
Write-Host ("  Report : {0}" -f $reportCsv) -ForegroundColor Green

if ($Apply) {
    Write-Host ''
    Write-Host '  Intune Management Extension-service herstarten...' -ForegroundColor Cyan
    Restart-Service -Name 'IntuneManagementExtension' -Force
    Write-Host '  [OK]   Service herstart — apps worden opnieuw geëvalueerd bij de volgende check-in.' -ForegroundColor Green

    if ($ForceSync) {
        Write-Host ''
        Write-Host '  Directe MDM-sync forceren...' -ForegroundColor Cyan
        Invoke-ForceMdmSync
    }
} else {
    Write-Host ''
    Write-Host '  Run met -Apply om deze apps daadwerkelijk uit hun GRS-cooldown te halen.' -ForegroundColor Yellow
}

Write-Host ''
