#Requires -Version 7.0
<#
.SYNOPSIS
    Create the baseline Entra ID security groups used to bootstrap a newly onboarded tenant.

.DESCRIPTION
    Creates a small set of security groups that a new-tenant onboarding checklist
    typically needs before Conditional Access / Intune baselines can be applied:

      - A dynamic exclusion group ("all users except break-glass accounts") — used to
        scope Conditional Access policies and Intune baseline assignments so the
        break-glass account(s) created by New-BreakGlassAdminAccount.ps1 are never
        accidentally locked out.
      - Any number of additional static "feature toggle" groups, supplied via
        -AdditionalGroupNames — e.g. groups used to gate a specific add-on (password
        manager, a line-of-business app, Windows 365) to a subset of users via
        Conditional Access or Intune assignment filters.

    Default behavior is a dry run — pass -Apply to actually create the groups.

    Sign-in goes through scripts\Startup\Connect-M365.ps1: delegated by default (you
    sign in as an admin; device code when $global:useDeviceCodeAuth is set, the GDAP
    customer tenant from $global:cid), app-only with -ClientId and
    -CertificateThumbprint, or -AppOnly. An existing Graph session is reused only when
    it is for the right tenant and already holds Group.ReadWrite.All.

.PARAMETER BreakGlassUpnPattern
    A substring/pattern matched against userPrincipalName to build the dynamic
    exclusion group's membership rule (e.g. "breakglass-admin" or your break-glass
    account's naming convention). Required unless -SkipExclusionGroup is used.

.PARAMETER ExclusionGroupName
    Display name for the dynamic exclusion group. Default: "SG - All Users Except Break Glass".

.PARAMETER SkipExclusionGroup
    Skip creating the dynamic break-glass exclusion group.

.PARAMETER AdditionalGroupNames
    Array of display names for additional static security groups to create (empty
    membership — populate later via Add-UserToFeatureGroup.ps1 or a dynamic rule of
    your own). Example: "SG - Enable Password Manager", "SG - Enable Windows 365".

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
    Actually create the groups. Without this switch, the script only reports what it
    would create.

.EXAMPLE
    # Preview the baseline exclusion group plus two feature-toggle groups
    .\New-TenantBaselineGroups.ps1 -BreakGlassUpnPattern "breakglass-admin" `
        -AdditionalGroupNames "SG - Enable Password Manager", "SG - Enable Windows 365"

.EXAMPLE
    .\New-TenantBaselineGroups.ps1 -BreakGlassUpnPattern "breakglass-admin" -Apply

.NOTES
    Required scopes : Group.ReadWrite.All
    Required module : Microsoft.Graph.Groups
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $BreakGlassUpnPattern,
    [string] $ExclusionGroupName = 'SG - All Users Except Break Glass',
    [switch] $SkipExclusionGroup,
    [string[]] $AdditionalGroupNames = @(),
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly,
    [switch] $Apply
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

if (-not $SkipExclusionGroup -and -not $BreakGlassUpnPattern) {
    throw "-BreakGlassUpnPattern is required unless -SkipExclusionGroup is specified."
}

try {
    $graph = Connect-M365Graph -Scopes 'Group.ReadWrite.All' -TenantId $TenantId `
        -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly
} catch {
    Write-Host "  [ERROR] Could not connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   New Tenant Baseline Groups" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "  Mode : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

function New-BaselineGroup {
    param(
        [string] $DisplayName,
        [string] $MembershipRule
    )

    $mailNickname = ($DisplayName -replace '[^a-zA-Z0-9]', '')
    $existing = Get-MgGroup -Filter "displayName eq '$($DisplayName -replace "'", "''")'" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($existing) {
        Write-Host "  [SKIP] '$DisplayName' already exists (Id: $($existing.Id))." -ForegroundColor DarkGray
        return $existing
    }

    if (-not $Apply) {
        Write-Host "  [PREVIEW] Would create group '$DisplayName'$(if ($MembershipRule) { " (dynamic: $MembershipRule)" })." -ForegroundColor Yellow
        return $null
    }

    if (-not $PSCmdlet.ShouldProcess($DisplayName, "Create Entra ID group")) { return $null }

    $params = @{
        DisplayName     = $DisplayName
        MailEnabled     = $false
        MailNickname    = $mailNickname
        SecurityEnabled = $true
    }
    if ($MembershipRule) {
        $params['GroupTypes']                     = @('DynamicMembership')
        $params['MembershipRule']                 = $MembershipRule
        $params['MembershipRuleProcessingState']  = 'On'
    }

    $group = New-MgGroup @params -ErrorAction Stop
    Write-Host "  [OK]   Created '$DisplayName' (Id: $($group.Id))." -ForegroundColor Green
    return $group
}

if (-not $SkipExclusionGroup) {
    $rule = "(NOT (user.userPrincipalName -contains `"$BreakGlassUpnPattern`"))"
    New-BaselineGroup -DisplayName $ExclusionGroupName -MembershipRule $rule | Out-Null
}

foreach ($name in $AdditionalGroupNames) {
    New-BaselineGroup -DisplayName $name | Out-Null
}

Write-Host ""
if (-not $Apply) { Write-Host "  Re-run with -Apply to create these groups." -ForegroundColor Yellow }
Write-Host ""

Disconnect-M365Graph $graph
