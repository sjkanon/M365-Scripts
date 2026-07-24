#Requires -Version 5.1
<#
.SYNOPSIS
    Report which Entra ID groups are assigned to which Intune policies, profiles, and apps.

.DESCRIPTION
    Connects to Microsoft Graph and enumerates Intune device configuration profiles,
    compliance policies, settings catalog (configuration) policies, and mobile apps,
    resolving each object's assignments to readable group names (or "All users" /
    "All devices" for the built-in virtual targets), and whether the assignment is an
    Include or Exclude target. Read-only — one CSV row per policy/assignment pair.

    Useful for answering "what's this group actually getting pushed?" or "where is this
    policy assigned?" across a tenant without clicking through every blade in the Intune
    admin center.

.PARAMETER PolicyType
    Which object types to report on. Default: all of them.
    (DeviceConfiguration, CompliancePolicy, SettingsCatalog, MobileApp)

.PARAMETER OutputPath
    CSV report path. Defaults to .\IntunePolicyAssignments_<timestamp>.csv.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Get-IntunePolicyAssignments.ps1

.EXAMPLE
    .\Get-IntunePolicyAssignments.ps1 -PolicyType CompliancePolicy,SettingsCatalog

.NOTES
    Inspired by the capability list of the retired directorcia/patron toolkit
    (endpoint-assign-get.ps1), rewritten from scratch against the
    Microsoft.Graph.DeviceManagement module — the original used the deprecated
    Microsoft.Graph.Intune module and a custom raw-REST assignment-target switch
    statement. The mutating counterparts (endpoint-assign-set.ps1 / -del.ps1, which
    bulk-add/remove assignments) were intentionally not ported — bulk assignment changes
    are high-blast-radius and better reviewed one policy at a time in the admin center, or
    scripted per-engagement against this report's output.

    Required modules: Microsoft.Graph.DeviceManagement, Microsoft.Graph.Groups
    Required scopes: DeviceManagementConfiguration.Read.All, DeviceManagementApps.Read.All,
                      Group.Read.All
#>
[CmdletBinding()]
param(
    [ValidateSet('DeviceConfiguration', 'CompliancePolicy', 'SettingsCatalog', 'MobileApp')]
    [string[]] $PolicyType = @('DeviceConfiguration', 'CompliancePolicy', 'SettingsCatalog', 'MobileApp'),
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
    $connectParams = @{ Scopes = @('DeviceManagementConfiguration.Read.All', 'DeviceManagementApps.Read.All', 'Group.Read.All') }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Intune Policy Assignment Report" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

$groupCache = @{}
function Get-CachedGroupName {
    param([string] $GroupId)
    if ($groupCache.ContainsKey($GroupId)) { return $groupCache[$GroupId] }
    try {
        $g = Get-MgGroup -GroupId $GroupId -Property DisplayName -ErrorAction Stop
        $groupCache[$GroupId] = $g.DisplayName
    } catch {
        $groupCache[$GroupId] = $GroupId
    }
    return $groupCache[$GroupId]
}

function Resolve-AssignmentTarget {
    param($Target)
    switch ($Target.AdditionalProperties.'@odata.type') {
        '#microsoft.graph.allDevicesAssignmentTarget' { return @{ Name = 'All devices'; Type = 'Include' } }
        '#microsoft.graph.allLicensedUsersAssignmentTarget' { return @{ Name = 'All users'; Type = 'Include' } }
        '#microsoft.graph.exclusionGroupAssignmentTarget' { return @{ Name = (Get-CachedGroupName $Target.AdditionalProperties.groupId); Type = 'Exclude' } }
        '#microsoft.graph.groupAssignmentTarget' { return @{ Name = (Get-CachedGroupName $Target.AdditionalProperties.groupId); Type = 'Include' } }
        default { return @{ Name = 'Unknown'; Type = 'Unknown' } }
    }
}

$results = [System.Collections.Generic.List[PSObject]]::new()

function Add-AssignmentRows {
    param([string] $Category, [string] $Name, [string] $Id, $Assignments)
    if (-not $Assignments -or @($Assignments).Count -eq 0) {
        $script:results.Add([PSCustomObject]@{ PolicyType = $Category; PolicyName = $Name; PolicyId = $Id; AssignmentType = 'None'; Target = '(no assignments)' })
        return
    }
    foreach ($a in $Assignments) {
        $resolved = Resolve-AssignmentTarget -Target $a.Target
        $script:results.Add([PSCustomObject]@{ PolicyType = $Category; PolicyName = $Name; PolicyId = $Id; AssignmentType = $resolved.Type; Target = $resolved.Name })
    }
}

if ($PolicyType -contains 'DeviceConfiguration') {
    Write-Host "  Device configuration profiles..." -ForegroundColor DarkGray
    foreach ($p in Get-MgDeviceManagementDeviceConfiguration -All -ErrorAction SilentlyContinue) {
        $assignments = Get-MgDeviceManagementDeviceConfigurationAssignment -DeviceConfigurationId $p.Id -ErrorAction SilentlyContinue
        Add-AssignmentRows -Category 'Device Configuration' -Name $p.DisplayName -Id $p.Id -Assignments $assignments
    }
}

if ($PolicyType -contains 'CompliancePolicy') {
    Write-Host "  Compliance policies..." -ForegroundColor DarkGray
    foreach ($p in Get-MgDeviceManagementDeviceCompliancePolicy -All -ErrorAction SilentlyContinue) {
        $assignments = Get-MgDeviceManagementDeviceCompliancePolicyAssignment -DeviceCompliancePolicyId $p.Id -ErrorAction SilentlyContinue
        Add-AssignmentRows -Category 'Compliance Policy' -Name $p.DisplayName -Id $p.Id -Assignments $assignments
    }
}

if ($PolicyType -contains 'SettingsCatalog') {
    Write-Host "  Settings catalog policies..." -ForegroundColor DarkGray
    foreach ($p in Get-MgDeviceManagementConfigurationPolicy -All -ErrorAction SilentlyContinue) {
        $assignments = Get-MgDeviceManagementConfigurationPolicyAssignment -DeviceManagementConfigurationPolicyId $p.Id -ErrorAction SilentlyContinue
        Add-AssignmentRows -Category 'Settings Catalog' -Name $p.Name -Id $p.Id -Assignments $assignments
    }
}

if ($PolicyType -contains 'MobileApp') {
    Write-Host "  Mobile apps..." -ForegroundColor DarkGray
    foreach ($p in Get-MgDeviceAppManagementMobileApp -All -ErrorAction SilentlyContinue) {
        $assignments = Get-MgDeviceAppManagementMobileAppAssignment -MobileAppId $p.Id -ErrorAction SilentlyContinue
        Add-AssignmentRows -Category 'Mobile App' -Name $p.DisplayName -Id $p.Id -Assignments $assignments
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
if ($results.Count -eq 0) {
    Write-Host "  No policies found for the selected type(s)." -ForegroundColor DarkGray
} else {
    $results | Format-Table PolicyType, PolicyName, AssignmentType, Target -AutoSize

    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "IntunePolicyAssignments_$ts.csv"
    }
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

$policyCount = ($results | Select-Object -Unique PolicyType, PolicyId).Count
Write-Host ""
Write-Host ("  {0} polic(y/ies) — {1} assignment row(s)" -f $policyCount, $results.Count) -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
