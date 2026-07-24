#Requires -Version 5.1
<#
.SYNOPSIS
    Report registered Windows Autopilot devices and deployment profiles for a tenant.

.DESCRIPTION
    Connects to Microsoft Graph and lists every Windows Autopilot device identity
    registered in the tenant (serial number, model, manufacturer, group tag, enrollment
    state, deployment profile assignment status), alongside a summary of the deployment
    profiles that exist. Read-only tenant-side inventory report — this is not the same as
    Intune/Get-Autopilot/Get-WindowsAutoPilotInfo.ps1 (which collects a hardware hash from
    a physical device to register it); this script reports on devices already registered.

.PARAMETER GroupTag
    Only report devices with this group tag.

.PARAMETER OutputPath
    CSV report path. Defaults to .\AutopilotDevices_<timestamp>.csv.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Get-AutopilotDevices.ps1

.EXAMPLE
    .\Get-AutopilotDevices.ps1 -GroupTag "Finance-Laptops"

.NOTES
    Inspired by the capability list of the retired directorcia/patron toolkit
    (autopilot-device-get.ps1, autopilot-profile-get.ps1), rewritten from scratch against
    Microsoft.Graph.DeviceManagement.Enrollment — the originals used the deprecated
    Microsoft.Graph.Intune module. The mutating counterparts (autopilot-device-import.ps1
    / -assign.ps1 / -del.ps1 and autopilot-profile-import.ps1 / -assign.ps1 / -set.ps1 /
    -del.ps1) were intentionally not ported — device registration/profile assignment is
    typically driven from a CSV per deployment batch and is better handled with
    Get-WindowsAutoPilotInfo.ps1 (already in this repo) plus the Intune admin center or
    Microsoft's own Get-AutopilotDeployment community module, rather than a bulk scripted
    delete/reassign against this list.

    Required scopes: DeviceManagementServiceConfig.Read.All
#>
[CmdletBinding()]
param(
    [string] $GroupTag,
    [string] $OutputPath,
    [string] $TenantId
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    $null = Get-MgContext -ErrorAction Stop
    if (-not (Get-MgContext)) { throw }
} catch {
    $connectParams = @{ Scopes = @('DeviceManagementServiceConfig.Read.All') }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Windows Autopilot Device Inventory" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# ── Deployment profiles ──────────────────────────────────────────────────────
Write-Host "  Retrieving deployment profiles..." -ForegroundColor DarkGray
$profiles = @(Get-MgDeviceManagementWindowsAutopilotDeploymentProfile -All -ErrorAction SilentlyContinue)
Write-Host "  Found $($profiles.Count) deployment profile(s):" -ForegroundColor DarkGray
foreach ($p in $profiles) {
    Write-Host "    - $($p.DisplayName)  (Id: $($p.Id))" -ForegroundColor DarkGray
}
Write-Host ""

# ── Devices ───────────────────────────────────────────────────────────────────
Write-Host "  Retrieving Autopilot devices..." -ForegroundColor DarkGray
$devices = @(Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -All -ErrorAction Stop)
if ($GroupTag) {
    $devices = $devices | Where-Object { $_.GroupTag -eq $GroupTag }
}
Write-Host "  Found $($devices.Count) device(s)." -ForegroundColor DarkGray
Write-Host ""

$results = foreach ($d in $devices) {
    [PSCustomObject]@{
        SerialNumber                    = $d.SerialNumber
        Model                            = $d.Model
        Manufacturer                     = $d.Manufacturer
        GroupTag                         = $d.GroupTag
        EnrollmentState                  = $d.EnrollmentState
        DeploymentProfileAssignmentStatus = $d.DeploymentProfileAssignmentStatus
        AzureAdDeviceId                  = $d.AzureAdDeviceId
        ManagedDeviceId                  = $d.ManagedDeviceId
        LastContactedDateTime            = $d.LastContactedDateTime
        AddressableUserName              = $d.AddressableUserName
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
if (-not $results -or @($results).Count -eq 0) {
    Write-Host "  No Autopilot devices found." -ForegroundColor DarkGray
} else {
    $results | Format-Table SerialNumber, Model, GroupTag, EnrollmentState, DeploymentProfileAssignmentStatus -AutoSize

    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "AutopilotDevices_$ts.csv"
    }
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

$unassigned = @($results | Where-Object { $_.DeploymentProfileAssignmentStatus -ne 'assigned' }).Count
Write-Host ""
Write-Host ("  {0} device(s) — {1} without an assigned deployment profile" -f @($results).Count, $unassigned) -ForegroundColor $(if ($unassigned -gt 0) { 'Yellow' } else { 'Cyan' })
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
