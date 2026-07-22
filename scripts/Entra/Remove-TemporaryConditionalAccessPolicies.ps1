#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.SignIns
<#
.SYNOPSIS
    Remove temporary Conditional Access policies created by this toolkit.

.DESCRIPTION
    Removes one policy by ID, or scans all TEMP-CA policies and removes the
    ones that are expired based on the Expires= timestamp in the description.

.PARAMETER PolicyId
    Remove one specific Conditional Access policy by ID.

.PARAMETER RemoveAllTempPolicies
    Remove all policies with display name prefix TEMP-CA -.

.PARAMETER IncludeNotYetExpired
    With scan mode, also remove temporary policies that are not expired yet.

.PARAMETER TenantId
    Optional tenant ID/domain for Connect-MgGraph.

.EXAMPLE
    .\Remove-TemporaryConditionalAccessPolicies.ps1

.EXAMPLE
    .\Remove-TemporaryConditionalAccessPolicies.ps1 -PolicyId "<policyId>"

.EXAMPLE
    .\Remove-TemporaryConditionalAccessPolicies.ps1 -RemoveAllTempPolicies
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string]$PolicyId,
    [switch]$RemoveAllTempPolicies,
    [switch]$IncludeNotYetExpired,
    [string]$TenantId
)

$ErrorActionPreference = 'Stop'

function Write-Step { param([string]$Message) Write-Host "`n=== $Message ===" -ForegroundColor Cyan }
function Write-Ok   { param([string]$Message) Write-Host "[OK]   $Message" -ForegroundColor Green }
function Write-Warn { param([string]$Message) Write-Host "[WARN] $Message" -ForegroundColor Yellow }

function Connect-GraphForConditionalAccess {
    param([string]$Tenant)

    $requiredScopes = @(
        'Policy.ReadWrite.ConditionalAccess',
        'Policy.Read.All'
    )

    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    $missingScope = $true

    if ($ctx -and $ctx.Scopes) {
        $missingScope = ($requiredScopes | Where-Object { $_ -notin $ctx.Scopes }).Count -gt 0
    }

    if (-not $ctx -or $missingScope) {
        $params = @{
            Scopes       = $requiredScopes
            ContextScope = 'Process'
            NoWelcome    = $true
        }
        if ($Tenant) {
            $params['TenantId'] = $Tenant
        }
        Connect-MgGraph @params | Out-Null
    }
}

function Get-ExpiresUtcFromDescription {
    param([string]$Description)

    if (-not $Description) { return $null }

    $match = [Regex]::Match($Description, 'Expires=([^;\s]+)')
    if (-not $match.Success) { return $null }

    try {
        return [DateTime]::Parse($match.Groups[1].Value).ToUniversalTime()
    } catch {
        return $null
    }
}

Connect-GraphForConditionalAccess -Tenant $TenantId

if ($PolicyId) {
    Write-Step 'Removing policy by ID'
    $policy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $PolicyId -ErrorAction Stop
    if ($PSCmdlet.ShouldProcess($policy.DisplayName, 'Remove temporary CA policy')) {
        Remove-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $PolicyId -ErrorAction Stop
        Write-Ok "Removed: $($policy.DisplayName) ($PolicyId)"
    }
    return
}

Write-Step 'Scanning temporary Conditional Access policies'
$all = Get-MgIdentityConditionalAccessPolicy -All
$tempPolicies = $all | Where-Object { $_.DisplayName -like 'TEMP-CA -*' }

if (-not $tempPolicies -or $tempPolicies.Count -eq 0) {
    Write-Warn 'No temporary CA policies found.'
    return
}

$nowUtc = [DateTime]::UtcNow
$toRemove = @()

foreach ($policy in $tempPolicies) {
    $expiresUtc = Get-ExpiresUtcFromDescription -Description $policy.Description
    if ($RemoveAllTempPolicies) {
        $toRemove += $policy
        continue
    }

    if ($IncludeNotYetExpired) {
        $toRemove += $policy
        continue
    }

    if ($expiresUtc -and $expiresUtc -le $nowUtc) {
        $toRemove += $policy
    }
}

if (-not $toRemove -or $toRemove.Count -eq 0) {
    Write-Warn 'No matching temporary CA policies to remove.'
    return
}

Write-Host "Found $($toRemove.Count) policy/policies to remove."

foreach ($policy in $toRemove) {
    if ($PSCmdlet.ShouldProcess($policy.DisplayName, 'Remove temporary CA policy')) {
        Remove-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id -ErrorAction Stop
        Write-Ok "Removed: $($policy.DisplayName) ($($policy.Id))"
    }
}
