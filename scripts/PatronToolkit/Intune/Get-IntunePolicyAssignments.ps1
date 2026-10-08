#Requires -Version 7.0
<#
.SYNOPSIS
    Report which Entra ID groups are assigned to which Intune policies, profiles, and apps.

.DESCRIPTION
    Connects to Microsoft Graph and enumerates Intune device configuration profiles,
    compliance policies, settings catalog (configuration) policies, and mobile apps,
    resolving each object's assignments to readable group names (or "All users" /
    "All devices" for the built-in virtual targets), and whether the assignment is an
    Include or Exclude target. Mobile app rows also carry the assignment intent
    (required / available / uninstall). Read-only — one CSV row per policy/assignment pair.

    Useful for answering "what's this group actually getting pushed?" or "where is this
    policy assigned?" across a tenant without clicking through every blade in the Intune
    admin center.

.PARAMETER PolicyType
    Which object types to report on. Default: all of them.
    (DeviceConfiguration, CompliancePolicy, SettingsCatalog, MobileApp)

.PARAMETER OutputPath
    CSV report path. Defaults to .\IntunePolicyAssignments_<timestamp>.csv.

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
    .\Get-IntunePolicyAssignments.ps1

.EXAMPLE
    .\Get-IntunePolicyAssignments.ps1 -PolicyType CompliancePolicy,SettingsCatalog

.EXAMPLE
    .\Get-IntunePolicyAssignments.ps1 -TenantId contoso.onmicrosoft.com -AppOnly

.NOTES
    Inspired by the capability list of the retired directorcia/patron toolkit
    (endpoint-assign-get.ps1), rewritten from scratch against Microsoft Graph — the
    original used the deprecated Microsoft.Graph.Intune module and a custom raw-REST
    assignment-target switch statement. The mutating counterparts
    (endpoint-assign-set.ps1 / -del.ps1, which bulk-add/remove assignments) were
    intentionally not ported — bulk assignment changes are high-blast-radius and better
    reviewed one policy at a time in the admin center, or scripted per-engagement against
    this report's output.

    All calls are Invoke-MgGraphRequest with @odata.nextLink paging and
    $expand=assignments. Device configuration, compliance policies and mobile apps use
    Graph v1.0; settings catalog policies (configurationPolicies) only exist in Graph
    beta — the Microsoft.Graph v2 SDK has no Get-MgDeviceManagementConfigurationPolicy
    cmdlet, which is why the earlier version silently reported none.

    Sign-in: Microsoft Graph through scripts\Startup\Connect-M365.ps1 - delegated by
    default (scopes DeviceManagementConfiguration.Read.All, DeviceManagementApps.Read.All,
    Group.Read.All plus an Intune role; device code / GDAP customer per load.config.ps1),
    app-only with -ClientId/-CertificateThumbprint or -AppOnly (the same application
    permissions). An existing fitting Graph session is reused and left connected.
#>
[CmdletBinding()]
param(
    [ValidateSet('DeviceConfiguration', 'CompliancePolicy', 'SettingsCatalog', 'MobileApp')]
    [string[]] $PolicyType = @('DeviceConfiguration', 'CompliancePolicy', 'SettingsCatalog', 'MobileApp'),
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
$graph = Connect-M365Graph -Scopes 'DeviceManagementConfiguration.Read.All', 'DeviceManagementApps.Read.All', 'Group.Read.All' `
    -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Intune Policy Assignment Report" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

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

$groupCache = @{}
function Get-CachedGroupName {
    param([string] $GroupId)
    if (-not $GroupId) { return '(no group id)' }
    if ($groupCache.ContainsKey($GroupId)) { return $groupCache[$GroupId] }
    try {
        $g = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId`?`$select=displayName" -OutputType Hashtable -ErrorAction Stop
        $groupCache[$GroupId] = $g.displayName
    } catch {
        $groupCache[$GroupId] = $GroupId   # deleted group or no access: show the id
    }
    return $groupCache[$GroupId]
}

function Resolve-AssignmentTarget {
    param($Target)
    switch ([string]$Target.'@odata.type') {
        '#microsoft.graph.allDevicesAssignmentTarget'       { return @{ Name = 'All devices'; Type = 'Include' } }
        '#microsoft.graph.allLicensedUsersAssignmentTarget' { return @{ Name = 'All users'; Type = 'Include' } }
        '#microsoft.graph.exclusionGroupAssignmentTarget'   { return @{ Name = (Get-CachedGroupName $Target.groupId); Type = 'Exclude' } }
        '#microsoft.graph.groupAssignmentTarget'            { return @{ Name = (Get-CachedGroupName $Target.groupId); Type = 'Include' } }
        default { return @{ Name = "Unknown ($($Target.'@odata.type'))"; Type = 'Unknown' } }
    }
}

$results = [System.Collections.Generic.List[PSObject]]::new()

function Add-AssignmentRows {
    param([string] $Category, [string] $Name, [string] $Id, $Assignments)
    if (-not $Assignments -or @($Assignments).Count -eq 0) {
        $script:results.Add([PSCustomObject]@{ PolicyType = $Category; PolicyName = $Name; PolicyId = $Id; AssignmentType = 'None'; Target = '(no assignments)'; Intent = '' })
        return
    }
    foreach ($a in $Assignments) {
        $resolved = Resolve-AssignmentTarget -Target $a.target
        $script:results.Add([PSCustomObject]@{ PolicyType = $Category; PolicyName = $Name; PolicyId = $Id; AssignmentType = $resolved.Type; Target = $resolved.Name; Intent = [string]$a.intent })
    }
}

# One entry per object type: the collection, which Graph version carries it, and the
# property that holds the display name.
$collections = [ordered]@{
    DeviceConfiguration = @{ Label = 'Device Configuration'; Base = 'https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations';     NameProp = 'displayName' }
    CompliancePolicy    = @{ Label = 'Compliance Policy';    Base = 'https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies'; NameProp = 'displayName' }
    SettingsCatalog     = @{ Label = 'Settings Catalog';     Base = 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies';    NameProp = 'name' }
    MobileApp           = @{ Label = 'Mobile App';           Base = 'https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps';            NameProp = 'displayName' }
}

$failed = 0
foreach ($type in $collections.Keys) {
    if ($PolicyType -notcontains $type) { continue }
    $c = $collections[$type]
    Write-Host "  $($c.Label)..." -ForegroundColor DarkGray
    try {
        $objects = @(Get-GraphPaged -Uri "$($c.Base)?`$expand=assignments")
    } catch {
        $failed++
        Write-Host "  [WARN] Could not list $($c.Label): $($_.Exception.Message)" -ForegroundColor Yellow
        continue
    }
    foreach ($o in $objects) {
        $assignments = $o['assignments']
        if (-not $o.ContainsKey('assignments')) {
            # Expand not honoured for this object: ask for its assignments directly.
            try { $assignments = @(Get-GraphPaged -Uri "$($c.Base)/$($o.id)/assignments") }
            catch { Write-Host "  [WARN] Assignments of '$($o[$c.NameProp])': $($_.Exception.Message)" -ForegroundColor Yellow; $assignments = $null }
        }
        Add-AssignmentRows -Category $c.Label -Name $o[$c.NameProp] -Id $o.id -Assignments $assignments
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
if ($results.Count -eq 0) {
    Write-Host "  No policies found for the selected type(s)." -ForegroundColor DarkGray
} else {
    $results | Format-Table PolicyType, PolicyName, AssignmentType, Target, Intent -AutoSize

    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "IntunePolicyAssignments_$ts.csv"
    }
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

$policyCount = @($results | Select-Object -Unique PolicyType, PolicyId).Count
Write-Host ""
Write-Host ("  {0} polic(y/ies) — {1} assignment row(s)" -f $policyCount, $results.Count) -ForegroundColor Cyan
if ($failed) { Write-Host "  $failed object type(s) could not be read — see the warnings above." -ForegroundColor Yellow }
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
Disconnect-M365Graph $graph
