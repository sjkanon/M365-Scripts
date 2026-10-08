#Requires -Version 7.0
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
      - Omit it and the script signs in and runs Start-IntuneBackup itself.

    Sign-in (only when -CustomerBackupPath is omitted) goes through
    scripts\Startup\Connect-M365.ps1: delegated by default (you sign in as an Intune
    admin; device code when $global:useDeviceCodeAuth is set; under GDAP the customer
    tenant from $global:cid), app-only with -ClientId and -CertificateThumbprint, or
    -AppOnly. The backup itself stays with the IntuneBackupAndRestore module, which
    reads Intune through Microsoft Graph (Invoke-MgGraphRequest) on this session.
    Start-IntuneBackup checks for five ReadWrite scopes and, when one is missing,
    calls Connect-MgGraph on its own - without a tenant, so a GDAP session would land
    in your own tenant. The script therefore asks for exactly those scopes, although
    it only reads.

.PARAMETER BaselinePath
    Path to the MSP reference backup folder (created previously with
    `Start-IntuneBackup -Path <path>` against your reference/template tenant).

.PARAMETER CustomerBackupPath
    Path to an existing backup of the customer tenant. If omitted, the script creates
    a fresh backup of the currently connected tenant into -OutputPath first.

.PARAMETER TenantId
    Entra ID tenant ID or domain to back up. Only used when -CustomerBackupPath is
    omitted. Defaults to the GDAP customer tenant, else the tenant you sign in to.

.PARAMETER ClientId
    App registration for app-only sign-in, with -CertificateThumbprint. The app needs
    the five DeviceManagement*.ReadWrite.All application permissions listed below.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.PARAMETER OutputPath
    Folder for the auto-backup (when -CustomerBackupPath is omitted) and the diff
    report (default: C:\Temp\ / ~/Downloads).

.EXAMPLE
    # Compare an existing customer backup against the MSP baseline
    .\Compare-IntuneConfig.ps1 -BaselinePath C:\IntuneBaseline -CustomerBackupPath C:\Temp\CustomerBackup

.EXAMPLE
    # Back up the (GDAP) customer tenant and compare it live; delegated sign-in
    .\Compare-IntuneConfig.ps1 -BaselinePath C:\IntuneBaseline

.NOTES
    Author  : Sjoerd Kanon
    Required module: IntuneBackupAndRestore (4.x, Microsoft.Graph based), Microsoft.Graph.Authentication
    Scopes (live backup): DeviceManagementApps.ReadWrite.All, DeviceManagementConfiguration.ReadWrite.All,
                          DeviceManagementServiceConfig.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All,
                          DeviceManagementScripts.ReadWrite.All
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory)] [string] $BaselinePath,
    [string] $CustomerBackupPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly,
    [string] $OutputPath = $(if ($IsWindows) { 'C:\Temp' } else { "$HOME/Downloads" })
)

. (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')

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
    # Exactly the scopes Start-IntuneBackup checks for; with fewer it reconnects on its
    # own, without -TenantId (see .DESCRIPTION).
    $scopes = 'DeviceManagementApps.ReadWrite.All', 'DeviceManagementConfiguration.ReadWrite.All',
              'DeviceManagementServiceConfig.ReadWrite.All', 'DeviceManagementManagedDevices.ReadWrite.All',
              'DeviceManagementScripts.ReadWrite.All'
    $graph = Connect-M365Graph -Scopes $scopes -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly
    Write-Host "  Connected to tenant $($graph.TenantId) ($($graph.AuthType))." -ForegroundColor DarkGray

    $CustomerBackupPath = Join-Path $OutputPath "IntuneBackup_$ts"
    New-Item -ItemType Directory -Path $CustomerBackupPath | Out-Null
    Write-Host "  Backing up connected tenant to $CustomerBackupPath ..." -ForegroundColor Cyan
    try {
        Start-IntuneBackup -Path $CustomerBackupPath | Out-Null
    } finally {
        Disconnect-M365Graph $graph
    }
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
