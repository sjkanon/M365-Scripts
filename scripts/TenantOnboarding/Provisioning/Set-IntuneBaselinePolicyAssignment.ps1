#Requires -Version 5.1
<#
.SYNOPSIS
    Assign a tenant's baseline Intune policies (matching a name filter) to a target group.

.DESCRIPTION
    Bulk-assigns Intune configuration profiles, compliance policies, administrative
    templates, PowerShell/remediation scripts, and security baselines whose display
    name matches -NameFilter to a single target group — typically the "all users
    except break-glass" group created by New-TenantBaselineGroups.ps1, right after a
    new tenant's baseline policies have been imported/configured.

    Uses Microsoft Graph's beta endpoint (Intune policy assignment is not fully
    exposed on v1.0 for all policy types). Default behavior is a dry run — pass
    -Apply to actually create the assignments. Connects to Microsoft Graph
    automatically if no session is active; reuses an existing session if already
    connected.

.PARAMETER TargetGroupId
    Object ID of the Entra ID group to assign matching policies to.

.PARAMETER NameFilter
    Wildcard filter (PowerShell -like syntax) applied to each policy's display name.
    Default: "*Default*".

.PARAMETER PolicyTypes
    Which Intune object types to include. Default: all of them.
    Valid values: ConfigurationProfiles, ComplianceProfiles, AdminTemplates, Scripts, SecurityBaselines.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected, or resolvable from a
    GDAP customer tenant context.

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
    [switch] $Apply
)

$effectiveTenantId = $TenantId
if (-not $effectiveTenantId) {
    try {
        if ($global:authMode -eq 'GDAP' -and $global:cid) { $effectiveTenantId = [string]$global:cid }
        elseif ($env:M365_CUSTOMER_TENANTID) { $effectiveTenantId = [string]$env:M365_CUSTOMER_TENANTID }
    } catch {}
}

$script:ConnectedHere = $false
try {
    if (-not (Get-MgContext)) {
        $connectParams = @{ Scopes = @('DeviceManagementConfiguration.ReadWrite.All', 'DeviceManagementServiceConfig.ReadWrite.All') }
        if ($effectiveTenantId) { $connectParams['TenantId'] = $effectiveTenantId }
        Connect-MgGraph @connectParams -NoWelcome -ErrorAction Stop
        $script:ConnectedHere = $true
    }
} catch {
    Write-Host "  [ERROR] Could not connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Resource path, assignment endpoint suffix, and the assignment body shape per policy type.
$typeMap = @{
    ConfigurationProfiles = @{ Resource = 'deviceManagement/deviceConfigurations';          AssignSuffix = 'assign'; BodyKey = 'assignments' }
    ComplianceProfiles    = @{ Resource = 'deviceManagement/deviceCompliancePolicies';       AssignSuffix = 'assign'; BodyKey = 'assignments' }
    AdminTemplates        = @{ Resource = 'deviceManagement/groupPolicyConfigurations';      AssignSuffix = 'assign'; BodyKey = 'assignments' }
    SecurityBaselines     = @{ Resource = 'deviceManagement/intents';                        AssignSuffix = 'assign'; BodyKey = 'assignments' }
    Scripts               = @{ Resource = 'deviceManagement/deviceManagementScripts';        AssignSuffix = 'assign'; BodyKey = 'deviceManagementScriptGroupAssignments' }
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
        $resp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/$($cfg.Resource)" -ErrorAction Stop
    } catch {
        Write-Host "    [WARN] Could not list $type : $($_.Exception.Message)" -ForegroundColor Yellow
        continue
    }

    $matches = @($resp.value | Where-Object { $_.displayName -like $NameFilter })
    if ($matches.Count -eq 0) {
        Write-Host "    No policies matched '$NameFilter'." -ForegroundColor DarkGray
        continue
    }

    foreach ($policy in $matches) {
        if (-not $Apply) {
            Write-Host "    [PREVIEW] Would assign '$($policy.displayName)'" -ForegroundColor Yellow
            continue
        }
        if (-not $PSCmdlet.ShouldProcess($policy.displayName, "Assign to group $TargetGroupId")) { continue }

        $uri = "https://graph.microsoft.com/beta/$($cfg.Resource)/$($policy.id)/$($cfg.AssignSuffix)"
        $body = if ($cfg.BodyKey -eq 'deviceManagementScriptGroupAssignments') {
            @{ deviceManagementScriptGroupAssignments = @(@{
                '@odata.type'  = '#microsoft.graph.deviceManagementScriptGroupAssignment'
                targetGroupId  = $TargetGroupId
                id             = $policy.id
            }) }
        } else {
            @{ assignments = @(@{ id = ''; target = @{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = $TargetGroupId } }) }
        }

        try {
            Invoke-MgGraphRequest -Method POST -Uri $uri -Body ($body | ConvertTo-Json -Depth 6) -ContentType 'application/json' -ErrorAction Stop | Out-Null
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

if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
