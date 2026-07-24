#Requires -Version 5.1
<#
.SYNOPSIS
    Generate an HTML network documentation report from a UniFi Controller / UniFi OS console.

.DESCRIPTION
    Logs in to a UniFi Network Controller or UniFi OS console (UDM/UDM-Pro/UDR), enumerates
    every site (or a single site with -Site), lists all adopted devices per site, and writes
    an HTML report with device name, model, MAC, IP, firmware version, adoption state, and
    uptime. Read-only — this script never changes controller configuration.

    Credentials are requested via -Credential (a PSCredential) or, if omitted, prompted
    interactively with Get-Credential. Never hardcode a controller password in a script or
    scheduled task — use a saved credential (e.g. Windows Credential Manager) if this needs
    to run unattended.

.PARAMETER Controller
    Base URL of the controller, e.g. https://unifi.contoso.local:8443 or https://192.168.1.1
    for a UniFi OS console.

.PARAMETER Credential
    Controller admin credentials. Prompted with Get-Credential if omitted.

.PARAMETER Site
    Limit the report to a single site (by name). If omitted, all sites are included.

.PARAMETER SkipCertificateCheck
    Accept self-signed/untrusted certificates — common for on-prem controllers without a
    public CA certificate.

.PARAMETER OutputPath
    Report output folder (default: C:\Temp\ on Windows, ~/Downloads on macOS/Linux).

.EXAMPLE
    .\Get-UnifiNetworkReport.ps1 -Controller "https://192.168.1.1" -SkipCertificateCheck

.EXAMPLE
    .\Get-UnifiNetworkReport.ps1 -Controller "https://unifi.contoso.local:8443" -Site "Head Office" -Credential $cred

.NOTES
    Author  : Sjoerd Kanon
    Requires: UnifiApi.ps1 in the same folder (dot-sourced automatically)
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory)] [string] $Controller,
    [pscredential] $Credential,
    [string] $Site,
    [switch] $SkipCertificateCheck,
    [string] $OutputPath = $(if ($IsWindows -or -not $PSVersionTable.PSVersion) { 'C:\Temp' } else { "$HOME/Downloads" })
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'UnifiApi.ps1')

if (-not $Credential) {
    $Credential = Get-Credential -Message "UniFi controller admin credentials for $Controller"
}

if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
$ts       = Get-Date -Format 'yyyyMMdd_HHmmss'
$reportHtml = Join-Path $OutputPath "UnifiNetworkReport_$ts.html"

Write-Host ''
Write-Host "  Connecting to $Controller ..." -ForegroundColor Cyan
$session = Connect-UnifiController -Controller $Controller -Credential $Credential -SkipCertificateCheck:$SkipCertificateCheck

try {
    $sites = @(Get-UnifiSite -Session $session)
    if ($Site) {
        $sites = $sites | Where-Object { $_.name -eq $Site -or $_.desc -eq $Site }
        if (-not $sites) { throw "Site '$Site' not found on this controller." }
    }

    $sections = foreach ($s in $sites) {
        Write-Host "  Site: $($s.desc)" -ForegroundColor Yellow
        $devices = @(Get-UnifiDevice -Session $session -Site $s.name)

        $rows = foreach ($d in $devices) {
            $uptimeSpan = if ($d.uptime) { [TimeSpan]::FromSeconds([double]$d.uptime) } else { $null }
            [PSCustomObject]@{
                Name       = if ($d.name) { $d.name } else { $d.model }
                Model      = $d.model
                Mac        = $d.mac
                Ip         = $d.ip
                Firmware   = $d.version
                Upgradable = [bool]$d.upgradable
                State      = switch ([int]$d.state) { 1 { 'Connected' } 0 { 'Disconnected' } default { "State $($d.state)" } }
                Uptime     = if ($uptimeSpan) { "{0}d {1}h {2}m" -f $uptimeSpan.Days, $uptimeSpan.Hours, $uptimeSpan.Minutes } else { '-' }
            }
        }

        $rowsHtml = ($rows | Sort-Object Name | ForEach-Object {
            $upgradeBadge = if ($_.Upgradable) { '<span class="badge">update available</span>' } else { '' }
            "<tr><td>$($_.Name)</td><td>$($_.Model)</td><td>$($_.Mac)</td><td>$($_.Ip)</td><td>$($_.Firmware) $upgradeBadge</td><td>$($_.State)</td><td>$($_.Uptime)</td></tr>"
        }) -join "`n"

        @"
<h2>$($s.desc) <span class="muted">($($devices.Count) devices)</span></h2>
<table>
<thead><tr><th>Name</th><th>Model</th><th>MAC</th><th>IP</th><th>Firmware</th><th>State</th><th>Uptime</th></tr></thead>
<tbody>
$rowsHtml
</tbody>
</table>
"@
    }

    $html = @"
<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>UniFi Network Report - $Controller</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; margin: 2rem; color: #1a1a1a; }
  h1 { font-size: 1.4rem; } h2 { font-size: 1.1rem; margin-top: 2rem; }
  .muted { color: #777; font-weight: normal; font-size: 0.9rem; }
  table { border-collapse: collapse; width: 100%; margin-top: 0.5rem; }
  th, td { text-align: left; padding: 6px 10px; border-bottom: 1px solid #ddd; font-size: 0.9rem; }
  th { background: #f2f2f2; }
  .badge { background: #ffe08a; color: #664d00; border-radius: 4px; padding: 1px 6px; font-size: 0.75rem; }
</style></head>
<body>
<h1>UniFi Network Report — $Controller</h1>
<p class="muted">Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
$($sections -join "`n")
</body></html>
"@

    $html | Out-File -FilePath $reportHtml -Encoding UTF8
    Write-Host ''
    Write-Host "  Report: $reportHtml" -ForegroundColor Green
} finally {
    Disconnect-UnifiController -Session $session
}
