#Requires -Version 7.0
<#
.SYNOPSIS
    Assign a tenant's baseline Intune policies (matching a name filter) to a target group.

.DESCRIPTION
    Bulk-assigns Intune configuration profiles, compliance policies, administrative
    templates, platform (PowerShell) scripts, and security baselines whose display
    name matches -NameFilter to a single target group — typically the "all users
    except break-glass" group created by New-TenantBaselineGroups.ps1, right after a
    new tenant's baseline policies have been imported/configured.

    The group is ADDED to each policy's existing assignments: Intune's /assign action
    replaces the whole assignment list, so the script reads the current assignments
    first and posts them back together with the new one. Policies already assigned to
    the group are skipped.

    Uses Microsoft Graph's beta endpoint (Intune policy assignment is not fully
    exposed on v1.0 for all policy types) and follows @odata.nextLink, so tenants with
    more policies than one page are covered. Default behavior is a dry run — pass
    -Apply to actually create the assignments.

    Sign-in goes through scripts\Startup\Connect-M365.ps1: delegated by default (you
    sign in as an Intune admin; device code when $global:useDeviceCodeAuth is set, the
    GDAP customer tenant from $global:cid), app-only with -ClientId and
    -CertificateThumbprint, or -AppOnly. An existing Graph session is reused only when
    it is for the right tenant and already holds the scopes below.

.PARAMETER TargetGroupId
    Object ID of the Entra ID group to assign matching policies to.

.PARAMETER NameFilter
    Wildcard filter (PowerShell -like syntax) applied to each policy's display name.
    Default: "*Default*".

.PARAMETER PolicyTypes
    Which Intune object types to include. Default: all of them.
    Valid values: ConfigurationProfiles, ComplianceProfiles, AdminTemplates, Scripts, SecurityBaselines.
    Settings catalog policies (configurationPolicies) are not covered.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Defaults to the GDAP customer tenant ($global:cid /
    $env:M365_CUSTOMER_TENANTID), else the tenant you sign in to.

.PARAMETER ClientId
    App registration for app-only sign-in, with -CertificateThumbprint.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.PARAMETER Apply
    Actually create the assignments. Without this switch, the script only reports
    which policies would be assigned.

.EXAMPLE
    # Preview what would be assigned to the baseline group
    .\Set-IntuneBaselinePolicyAssignment.ps1 -TargetGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

.EXAMPLE
    .\Set-IntuneBaselinePolicyAssignment.ps1 -TargetGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Apply

.EXAMPLE
    # Only compliance policies and security baselines named like "Baseline*"
    .\Set-IntuneBaselinePolicyAssignment.ps1 -TargetGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -NameFilter "Baseline*" -PolicyTypes ComplianceProfiles, SecurityBaselines -Apply

.NOTES
    Required scopes : DeviceManagementConfiguration.ReadWrite.All, DeviceManagementServiceConfig.ReadWrite.All
    Required module : Microsoft.Graph.Authentication
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $TargetGroupId,

    [string] $NameFilter = '*Default*',

    [ValidateSet('ConfigurationProfiles', 'ComplianceProfiles', 'AdminTemplates', 'Scripts', 'SecurityBaselines')]
    [string[]] $PolicyTypes = @('ConfigurationProfiles', 'ComplianceProfiles', 'AdminTemplates', 'Scripts', 'SecurityBaselines'),

    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly,
    [switch] $Apply
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

try {
    $graph = Connect-M365Graph -Scopes 'DeviceManagementConfiguration.ReadWrite.All', 'DeviceManagementServiceConfig.ReadWrite.All' `
        -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly
} catch {
    Write-Host "  [ERROR] Could not connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

function Get-GraphCollection {
    # GET a collection and follow @odata.nextLink until the last page.
    param([string] $Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    while ($Uri) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
        foreach ($v in @($resp.value)) { if ($null -ne $v) { $items.Add($v) } }
        $Uri = $resp.'@odata.nextLink'
    }
    return , $items
}

# Resource path and the key of the assignment list in the /assign body, per policy type.
$typeMap = @{
    ConfigurationProfiles = @{ Resource = 'deviceManagement/deviceConfigurations';     BodyKey = 'assignments' }
    ComplianceProfiles    = @{ Resource = 'deviceManagement/deviceCompliancePolicies'; BodyKey = 'assignments' }
    AdminTemplates        = @{ Resource = 'deviceManagement/groupPolicyConfigurations'; BodyKey = 'assignments' }
    SecurityBaselines     = @{ Resource = 'deviceManagement/intents';                  BodyKey = 'assignments' }
    Scripts               = @{ Resource = 'deviceManagement/deviceManagementScripts';  BodyKey = 'deviceManagementScriptAssignments' }
}

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Intune Baseline Policy Assignment" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "  Target group : $TargetGroupId"
Write-Host "  Name filter  : $NameFilter"
Write-Host "  Mode         : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

$assignedCount = 0

foreach ($type in $PolicyTypes) {
    $cfg = $typeMap[$type]
    Write-Host "  $type" -ForegroundColor Cyan

    try {
        $policies = Get-GraphCollection -Uri "https://graph.microsoft.com/beta/$($cfg.Resource)"
    } catch {
        Write-Host "    [WARN] Could not list $type : $($_.Exception.Message)" -ForegroundColor Yellow
        continue
    }

    $selected = @($policies | Where-Object { $_.displayName -like $NameFilter })
    if ($selected.Count -eq 0) {
        Write-Host "    No policies matched '$NameFilter'." -ForegroundColor DarkGray
        continue
    }

    foreach ($policy in $selected) {
        $base = "https://graph.microsoft.com/beta/$($cfg.Resource)/$($policy.id)"
        try {
            $current = Get-GraphCollection -Uri "$base/assignments"
        } catch {
            Write-Host "    [WARN] Could not read assignments of '$($policy.displayName)': $($_.Exception.Message)" -ForegroundColor Yellow
            continue
        }

        $already = $current | Where-Object {
            $_.target.groupId -eq $TargetGroupId -and $_.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget'
        }
        if ($already) {
            Write-Host "    [SKIP] '$($policy.displayName)' is already assigned to the group." -ForegroundColor DarkGray
            continue
        }

        if (-not $Apply) {
            Write-Host "    [PREVIEW] Would assign '$($policy.displayName)' (keeping its $($current.Count) existing assignment(s))" -ForegroundColor Yellow
            continue
        }
        if (-not $PSCmdlet.ShouldProcess($policy.displayName, "Assign to group $TargetGroupId")) { continue }

        # /assign replaces the full list: send the existing targets back with the new one.
        $list = @($current | ForEach-Object { @{ target = $_.target } })
        $list += @{ target = @{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = $TargetGroupId } }
        $body = @{ $cfg.BodyKey = $list }

        try {
            Invoke-MgGraphRequest -Method POST -Uri "$base/assign" -Body ($body | ConvertTo-Json -Depth 10) -ContentType 'application/json' -ErrorAction Stop | Out-Null
            Write-Host "    [OK]   Assigned '$($policy.displayName)'" -ForegroundColor Green
            $assignedCount++
        } catch {
            Write-Host "    [WARN] Failed to assign '$($policy.displayName)': $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

Write-Host ""
if ($Apply) { Write-Host "  Assigned $assignedCount polic(y/ies) to $TargetGroupId." -ForegroundColor Green }
else { Write-Host "  Re-run with -Apply to perform the assignments shown above." -ForegroundColor Yellow }
Write-Host ""

Disconnect-M365Graph $graph
