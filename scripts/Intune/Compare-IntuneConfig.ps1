#Requires -Version 5.1
<#
.SYNOPSIS
    Compare a customer tenant's Intune configuration against an MSP baseline backup.

.DESCRIPTION
    Wraps the community `IntuneBackupAndRestore` module to detect configuration drift:
    compliance policies, configuration profiles, and other Intune objects that differ
    from (or are missing/added relative to) a reference "baseline" backup.

    Read-only — this script never changes Intune configuration, it only reports
    differences. Use `Import-ConditionalAccessBaseline.ps1` in `scripts/Entra/` for
    the Conditional Access side of a baseline; this script covers everything else
    IntuneBackupAndRestore exports (compliance/configuration profiles, app protection, etc.).

    Two ways to get the "customer" side of the comparison:
      - Pass -CustomerBackupPath to an existing backup folder (from a previous
        Start-IntuneBackup run), or
      - Omit it and the script runs Start-IntuneBackup itself against the currently
        connected tenant (reuses an existing Graph session; GDAP-aware like the other
        scripts in this repo — resolves the customer tenant from $global:cid if set).

.PARAMETER BaselinePath
    Path to the MSP reference backup folder (created previously with
    `Start-IntuneBackup -Path <path>` against your reference/template tenant).

.PARAMETER CustomerBackupPath
    Path to an existing backup of the customer tenant. If omitted, the script creates
    a fresh backup of the currently connected tenant into -OutputPath first.

.PARAMETER TenantId
    Entra ID tenant ID or domain to connect to before backing up the customer tenant.
    Only used when -CustomerBackupPath is omitted. Optional if already connected or
    resolvable from a GDAP customer tenant context.

.PARAMETER OutputPath
    Folder for the auto-backup (when -CustomerBackupPath is omitted) and the diff
    report (default: C:\Temp\ / ~/Downloads).

.EXAMPLE
    # Compare an existing customer backup against the MSP baseline
    .\Compare-IntuneConfig.ps1 -BaselinePath C:\IntuneBaseline -CustomerBackupPath C:\Temp\CustomerBackup

.EXAMPLE
    # Back up the currently connected (GDAP) customer tenant and compare it live
    .\Compare-IntuneConfig.ps1 -BaselinePath C:\IntuneBaseline

.NOTES
    Author  : Sjoerd Kanon
    Required module: IntuneBackupAndRestore
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory)] [string] $BaselinePath,
    [string] $CustomerBackupPath,
    [string] $TenantId,
    [string] $OutputPath = $(if ($IsWindows -or -not $PSVersionTable.PSVersion) { 'C:\Temp' } else { "$HOME/Downloads" })
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $BaselinePath)) {
    throw "Baseline path not found: $BaselinePath"
}

if (-not (Get-Module -ListAvailable -Name IntuneBackupAndRestore)) {
    throw "The 'IntuneBackupAndRestore' module is required. Install it with: Install-Module IntuneBackupAndRestore -Scope CurrentUser"
}
Import-Module IntuneBackupAndRestore -ErrorAction Stop

if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
$ts        = Get-Date -Format 'yyyyMMdd_HHmmss'
$reportCsv = Join-Path $OutputPath "IntuneConfigDrift_$ts.csv"

if (-not $CustomerBackupPath) {
    # Reuse an existing Graph session if one is active; otherwise connect, GDAP-aware.
    if (-not (Get-MgContext)) {
        $connectParams = @{}
        $resolvedTenant = $TenantId
        if (-not $resolvedTenant -and $global:authMode -eq 'GDAP' -and $global:cid) {
            $resolvedTenant = $global:cid
            Write-Host "  Resolved tenant from GDAP customer context: $resolvedTenant" -ForegroundColor DarkGray
        }
        if ($resolvedTenant) { $connectParams['TenantId'] = $resolvedTenant }
        Connect-MgGraph @connectParams -Scopes 'DeviceManagementConfiguration.Read.All', 'DeviceManagementApps.Read.All' | Out-Null
    }

    $CustomerBackupPath = Join-Path $OutputPath "IntuneBackup_$ts"
    New-Item -ItemType Directory -Path $CustomerBackupPath | Out-Null
    Write-Host "  Backing up connected tenant to $CustomerBackupPath ..." -ForegroundColor Cyan
    Start-IntuneBackup -Path $CustomerBackupPath | Out-Null
}

Write-Host ''
Write-Host "  Baseline : $BaselinePath" -ForegroundColor Cyan
Write-Host "  Customer : $CustomerBackupPath" -ForegroundColor Cyan
Write-Host ''

$diff = Compare-IntuneBackupDirectories -ReferenceDirectory $BaselinePath -DifferenceDirectory $CustomerBackupPath

if (-not $diff) {
    Write-Host '  No drift detected — customer tenant matches the baseline.' -ForegroundColor Green
} else {
    $diff | Format-Table -AutoSize
    $diff | Export-Csv -Path $reportCsv -NoTypeInformation -Encoding UTF8
    Write-Host ''
    Write-Host "  $($diff.Count) drifted/added/removed object(s). Report: $reportCsv" -ForegroundColor Yellow
}
