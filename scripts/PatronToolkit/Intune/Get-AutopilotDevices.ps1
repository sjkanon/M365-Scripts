#Requires -Version 7.0
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
    Entra ID tenant ID or domain. Defaults to the GDAP customer (load.config.ps1) or
    your own tenant. Required for app-only sign-in.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint). Without it the
    script signs in delegated, as you.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

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

    Devices come from Graph v1.0 (windowsAutopilotDeviceIdentities); deployment profiles
    only exist in Graph beta (windowsAutopilotDeploymentProfiles) and are read with
    Invoke-MgGraphRequest — the Microsoft.Graph v2 SDK has no
    Get-MgDeviceManagementWindowsAutopilotDeploymentProfile cmdlet, which is why the
    earlier version always listed zero profiles.

    Sign-in: Microsoft Graph through scripts\Startup\Connect-M365.ps1 - delegated by
    default (scope DeviceManagementServiceConfig.Read.All plus an Intune role; device
    code / GDAP customer per load.config.ps1), app-only with
    -ClientId/-CertificateThumbprint or -AppOnly (the same application permission). An
    existing fitting Graph session is reused and left connected.
#>
[CmdletBinding()]
param(
    [string] $GroupTag,
    [string] $OutputPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
$graph = Connect-M365Graph -Scopes 'DeviceManagementServiceConfig.Read.All' -TenantId $TenantId `
    -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

function Get-GraphPaged {
    param([string] $Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    while ($Uri) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType Hashtable -ErrorAction Stop
        foreach ($v in @($resp.value)) { if ($null -ne $v) { $items.Add($v) } }
        $Uri = $resp.'@odata.nextLink'
    }
    $items
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Windows Autopilot Device Inventory" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# ── Deployment profiles ──────────────────────────────────────────────────────
Write-Host "  Retrieving deployment profiles..." -ForegroundColor DarkGray
try {
    $profiles = @(Get-GraphPaged -Uri 'https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles')
    Write-Host "  Found $($profiles.Count) deployment profile(s):" -ForegroundColor DarkGray
    foreach ($p in $profiles) {
        Write-Host "    - $($p.displayName)  (Id: $($p.id))" -ForegroundColor DarkGray
    }
} catch {
    Write-Host "  [WARN] Could not list deployment profiles: $($_.Exception.Message)" -ForegroundColor Yellow
}
Write-Host ""

# ── Devices ───────────────────────────────────────────────────────────────────
Write-Host "  Retrieving Autopilot devices..." -ForegroundColor DarkGray
try {
    $devices = @(Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -All -ErrorAction Stop)
} catch {
    Write-Host "  [ERROR] Could not list Autopilot devices: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-M365Graph $graph
    exit 1
}
if ($GroupTag) {
    $devices = @($devices | Where-Object { $_.GroupTag -eq $GroupTag })
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

# Graph reports assignedInSync / assignedOutOfSync / assignedUnkownSyncState for an
# assigned profile - there is no plain 'assigned' value, so the old -ne 'assigned'
# counted every device as unassigned.
$unassigned = @($results | Where-Object { [string]$_.DeploymentProfileAssignmentStatus -notlike 'assigned*' }).Count
Write-Host ""
Write-Host ("  {0} device(s) — {1} without an assigned deployment profile" -f @($results).Count, $unassigned) -ForegroundColor $(if ($unassigned -gt 0) { 'Yellow' } else { 'Cyan' })
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
Disconnect-M365Graph $graph
