#Requires -Version 5.1
<#
.SYNOPSIS
    List and optionally trigger firmware upgrades for UniFi devices across one or all sites.

.DESCRIPTION
    Logs in to a UniFi Network Controller / UniFi OS console, enumerates devices per site,
    and reports which ones have a firmware upgrade available (the controller's own
    "upgradable" flag — this script does not check for new firmware itself, it relies on
    the controller having already synced release info from Ubiquiti).

    Run without -Apply for a dry run — lists upgradable devices per site without changing
    anything. Run with -Apply to trigger the upgrade for each listed device (one at a time,
    with per-device confirmation unless -Force). Devices reboot during a firmware upgrade —
    this causes a brief network/service interruption for anything connected through them.

.PARAMETER Controller
    Base URL of the controller, e.g. https://unifi.contoso.local:8443 or https://192.168.1.1.

.PARAMETER Credential
    Controller admin credentials. Prompted with Get-Credential if omitted.

.PARAMETER Site
    Limit to a single site (by name). If omitted, all sites are processed.

.PARAMETER SkipCertificateCheck
    Accept self-signed/untrusted certificates.

.PARAMETER Apply
    Actually trigger the upgrade for each upgradable device. Without this switch, only a
    dry-run listing is shown.

.PARAMETER Force
    Skip the per-device confirmation prompt when used with -Apply.

.PARAMETER OutputPath
    Report output folder (default: C:\Temp\ on Windows, ~/Downloads on macOS/Linux).

.EXAMPLE
    # Dry run — see what would be upgraded across all sites
    .\Update-UnifiFirmware.ps1 -Controller "https://192.168.1.1" -SkipCertificateCheck

.EXAMPLE
    # Upgrade everything upgradable on one site, with per-device confirmation
    .\Update-UnifiFirmware.ps1 -Controller "https://192.168.1.1" -Site "Head Office" -Apply

.EXAMPLE
    # Unattended upgrade across all sites — use with care, causes reboots
    .\Update-UnifiFirmware.ps1 -Controller "https://192.168.1.1" -Apply -Force

.NOTES
    Author  : Sjoerd Kanon
    Requires: UnifiApi.ps1 in the same folder (dot-sourced automatically)
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param (
    [Parameter(Mandatory)] [string] $Controller,
    [pscredential] $Credential,
    [string] $Site,
    [switch] $SkipCertificateCheck,
    [switch] $Apply,
    [switch] $Force,
    [string] $OutputPath = $(if ($IsWindows -or -not $PSVersionTable.PSVersion) { 'C:\Temp' } else { "$HOME/Downloads" })
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'UnifiApi.ps1')

if (-not $Credential) {
    $Credential = Get-Credential -Message "UniFi controller admin credentials for $Controller"
}

if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
$ts        = Get-Date -Format 'yyyyMMdd_HHmmss'
$reportCsv = Join-Path $OutputPath "UnifiFirmwareUpgrade_$ts.csv"
$results   = [System.Collections.Generic.List[object]]::new()

Write-Host ''
Write-Host "  Mode: $(if ($Apply) { 'APPLY — upgrades will be triggered (devices will reboot)' } else { 'DRY RUN — no changes will be made' })" -ForegroundColor Cyan
Write-Host "  Connecting to $Controller ..." -ForegroundColor Cyan
$session = Connect-UnifiController -Controller $Controller -Credential $Credential -SkipCertificateCheck:$SkipCertificateCheck

try {
    $sites = @(Get-UnifiSite -Session $session)
    if ($Site) {
        $sites = $sites | Where-Object { $_.name -eq $Site -or $_.desc -eq $Site }
        if (-not $sites) { throw "Site '$Site' not found on this controller." }
    }

    foreach ($s in $sites) {
        $devices = @(Get-UnifiDevice -Session $session -Site $s.name)
        $upgradable = $devices | Where-Object { $_.upgradable }

        if (-not $upgradable) {
            Write-Host "  Site: $($s.desc) — up to date" -ForegroundColor DarkGray
            continue
        }

        Write-Host ''
        Write-Host "  Site: $($s.desc) — $($upgradable.Count) device(s) upgradable" -ForegroundColor Yellow

        foreach ($d in $upgradable) {
            $name = if ($d.name) { $d.name } else { $d.model }
            $upgradeTarget = if ($d.upgrade_to_firmware) { $d.upgrade_to_firmware } else { '(controller-selected)' }
            Write-Host "    $name [$($d.mac)] $($d.version) -> $upgradeTarget" -ForegroundColor Yellow

            $status = 'Skipped (dry run)'
            $shouldGo = $Apply -and ($Force -or $PSCmdlet.ShouldProcess("$name [$($d.mac)] on site $($s.desc)", 'Trigger firmware upgrade'))
            if ($shouldGo) {
                try {
                    Invoke-UnifiApi -Session $session -Site $s.name -Path 'cmd/devmgr' -Method Post `
                        -Body @{ cmd = 'upgrade'; mac = $d.mac } | Out-Null
                    $status = 'Upgrade triggered'
                    Write-Host "    [OK] Upgrade triggered for $name" -ForegroundColor Green
                } catch {
                    $status = "Failed: $($_.Exception.Message)"
                    Write-Host "    [FAIL] $name : $($_.Exception.Message)" -ForegroundColor Red
                }
            }

            $results.Add([PSCustomObject]@{
                Site           = $s.desc
                Device         = $name
                Mac            = $d.mac
                CurrentVersion = $d.version
                TargetVersion  = $upgradeTarget
                Status         = $status
            })
        }
    }
} finally {
    Disconnect-UnifiController -Session $session
}

$results | Export-Csv -Path $reportCsv -NoTypeInformation -Encoding UTF8
Write-Host ''
Write-Host "  Report: $reportCsv" -ForegroundColor DarkGray
if (-not $Apply -and $results.Count -gt 0) {
    Write-Host "  Dry run only — rerun with -Apply to trigger these upgrades." -ForegroundColor Yellow
}
