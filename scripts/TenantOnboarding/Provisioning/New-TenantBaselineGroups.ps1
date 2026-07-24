#Requires -Version 5.1
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
    Connects to Microsoft Graph automatically if no session is active; reuses an
    existing session if already connected.

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
    Entra ID tenant ID or domain. Optional if already connected, or resolvable from a
    GDAP customer tenant context.

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
    [switch] $Apply
)

if (-not $SkipExclusionGroup -and -not $BreakGlassUpnPattern) {
    throw "-BreakGlassUpnPattern is required unless -SkipExclusionGroup is specified."
}

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
        $connectParams = @{ Scopes = @('Group.ReadWrite.All') }
        if ($effectiveTenantId) { $connectParams['TenantId'] = $effectiveTenantId }
        Connect-MgGraph @connectParams -NoWelcome -ErrorAction Stop
        $script:ConnectedHere = $true
    }
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
    $existing = Get-MgGroup -Filter "displayName eq '$DisplayName'" -ErrorAction SilentlyContinue
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

if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
