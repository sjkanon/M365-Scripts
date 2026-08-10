#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Intune Remediation detection script - finds Win32 apps stuck behind Intune's GRS retry
    cooldown on this device.

.DESCRIPTION
    Companion detection script for Remediate-StuckWin32AppEnforcement.ps1, meant to be deployed
    together as an Intune "Scripts and remediations > Remediations" package - so a device that
    ends up stuck (e.g. behind a since-fixed requirement rule, like Claude Desktop's
    RequiredOSArchitecture bug - see ../Desktop/ClaudeDesktop/readme.md "Known issue") clears
    itself automatically on the next scheduled remediation run, with no manual per-device
    PowerShell session needed. Same detection logic as the standalone
    Repair-StuckWin32AppEnforcement.ps1's dry-run mode, split out so Intune's own Remediations
    scheduler drives it end to end.

    Reports non-compliant (exit 1) if any Win32 app under
    HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps has a cached
    EnforcementStateMessage with a real error code (not 0 = success, not 3010 = soft reboot
    pending) - exactly the apps that get 24h-locked into GRS after 3 failed attempts.

    Exits 0 (compliant, no remediation triggered) when nothing is stuck - including the normal
    case of a device that has never had any Win32 app installation failure.

.NOTES
    Deploy as the "Detection script" half of an Intune Remediation, paired with
    Remediate-StuckWin32AppEnforcement.ps1 as the "Remediation script" half. Run using SYSTEM
    (not logged-on credentials), 64-bit PowerShell, no signature check - same settings as the
    Claude Desktop custom detection script.
#>

$ErrorActionPreference = 'Stop'

function Get-FailedWin32AppStates {
    # Alleen "echte" fouten (niet 0 = geslaagd, niet 3010 = herstart in afwachting) - dat zijn
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

        # SID en App-ID zijn de twee padsegmenten direct na "Win32Apps\" - regex i.p.v. een vaste
        # index, zodat dit standhoudt ook als IME een extra subkey-niveau toevoegt/verwijdert.
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

    # Een entry per (user, app) - meerdere subkeys onder dezelfde app kunnen matchen.
    return @($results | Sort-Object UserObjectId, AppId -Unique)
}

try {
    # @() dwingt array-vorm af ongeacht hoeveel items de functie teruggeeft - zonder dit "pakt"
    # PowerShell een 1-item resultaat bij het verlaten van de functie uit tot een los object, dat
    # onder Set-StrictMode geen .Count-property heeft ("property 'Count' cannot be found").
    $failedStates = @(Get-FailedWin32AppStates)

    if ($failedStates.Count -eq 0) {
        Write-Host "Compliant: no Win32 apps with a stuck/failed enforcement state found."
        exit 0
    }

    $summary = ($failedStates | ForEach-Object { "$($_.AppId)=ErrorCode:$($_.ErrorCode)" }) -join '; '
    Write-Host "Non-compliant: $($failedStates.Count) Win32 app(s) with a real last error code (possible GRS cooldown): $summary"
    exit 1
} catch {
    Write-Host "ERROR: detection failed: $_"
    exit 1
}
