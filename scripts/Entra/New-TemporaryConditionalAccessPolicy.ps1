#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.SignIns
<#
.SYNOPSIS
    Create a temporary Conditional Access policy for a user or group.

.DESCRIPTION
    Creates a Conditional Access policy with an expiry marker in the description.
    Optionally keeps the current session open and removes the policy immediately
    after the configured end time.

.PARAMETER TargetType
    Target object type: User or Group.

.PARAMETER TargetId
    Object ID of the target user or group.

.PARAMETER DisplayName
    Friendly policy name (without prefix).

.PARAMETER DurationHours
    Duration in hours from start time. Default: 4.

.PARAMETER StartDelayMinutes
    Optional start delay in minutes (when no explicit start/end are supplied).

.PARAMETER StartDateTimeLocal
    Optional local start date/time (e.g. 2026-07-22 19:30).

.PARAMETER EndDateTimeLocal
    Optional local end date/time (e.g. 2026-07-22 22:00).

.PARAMETER Action
    Policy behavior:
      RequireMfa  -> requires MFA
      Block       -> blocks sign-in

.PARAMETER State
    Policy state.

.PARAMETER TenantId
    Optional tenant ID/domain for Connect-MgGraph.

.PARAMETER NoAutoCleanup
    If set, do not wait and auto-remove at expiry.

.EXAMPLE
    .\New-TemporaryConditionalAccessPolicy.ps1 -TargetType User -TargetId "<objectId>" -DisplayName "Breakglass temp" -DurationHours 2 -Action RequireMfa

.EXAMPLE
    .\New-TemporaryConditionalAccessPolicy.ps1 -TargetType User -TargetId "<objectId>" -DisplayName "Install window" -StartDateTimeLocal "2026-07-23 19:00" -EndDateTimeLocal "2026-07-23 22:00"

.EXAMPLE
    .\New-TemporaryConditionalAccessPolicy.ps1 -TargetType Group -TargetId "<objectId>" -DisplayName "Temp strict policy" -Action Block -NoAutoCleanup
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('User', 'Group')]
    [string]$TargetType,

    [Parameter(Mandatory = $true)]
    [string]$TargetId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$DisplayName,

    [ValidateRange(1, 168)]
    [int]$DurationHours = 4,

    [ValidateRange(0, 1440)]
    [int]$StartDelayMinutes = 0,

    [datetime]$StartDateTimeLocal,

    [datetime]$EndDateTimeLocal,

    [ValidateSet('RequireMfa', 'Block')]
    [string]$Action = 'RequireMfa',

    [ValidateSet('enabled', 'enabledForReportingButNotEnforced', 'disabled')]
    [string]$State = 'enabled',

    [string]$TenantId,

    [switch]$NoAutoCleanup
)

$ErrorActionPreference = 'Stop'

function Write-Step { param([string]$Message) Write-Host "`n=== $Message ===" -ForegroundColor Cyan }
function Write-Ok   { param([string]$Message) Write-Host "[OK]   $Message" -ForegroundColor Green }
function Write-Warn { param([string]$Message) Write-Host "[WARN] $Message" -ForegroundColor Yellow }

function Wait-UntilUtc {
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$UtcTime,
        [Parameter(Mandatory = $true)]
        [string]$Phase
    )

    while ($true) {
        $remaining = [int][Math]::Ceiling(($UtcTime - [DateTime]::UtcNow).TotalSeconds)
        if ($remaining -le 0) { break }

        Write-Host "  [$Phase] Remaining: $remaining second(s)"
        $sleepFor = if ($remaining -gt 60) { 60 } else { $remaining }
        Start-Sleep -Seconds $sleepFor
    }
}

function Connect-GraphForConditionalAccess {
    param([string]$Tenant)

    $requiredScopes = @(
        'Policy.ReadWrite.ConditionalAccess',
        'Policy.Read.All',
        'Directory.Read.All'
    )

    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    $missingScope = $true

    if ($ctx -and $ctx.Scopes) {
        $missingScope = ($requiredScopes | Where-Object { $_ -notin $ctx.Scopes }).Count -gt 0
    }

    if (-not $ctx -or $missingScope) {
        Write-Step 'Connecting to Microsoft Graph for Conditional Access'
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

function Remove-TemporaryPolicyById {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PolicyId,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    Remove-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $PolicyId -ErrorAction Stop
    Write-Ok "Temporary policy removed: $Name ($PolicyId)"
}

Connect-GraphForConditionalAccess -Tenant $TenantId

if (($PSBoundParameters.ContainsKey('StartDateTimeLocal') -and -not $PSBoundParameters.ContainsKey('EndDateTimeLocal')) -or
    (-not $PSBoundParameters.ContainsKey('StartDateTimeLocal') -and $PSBoundParameters.ContainsKey('EndDateTimeLocal'))) {
    throw 'Provide both -StartDateTimeLocal and -EndDateTimeLocal, or neither.'
}

if ($PSBoundParameters.ContainsKey('StartDateTimeLocal')) {
    $startUtc = $StartDateTimeLocal.ToUniversalTime()
    $endUtc   = $EndDateTimeLocal.ToUniversalTime()
} else {
    $startUtc = [DateTime]::UtcNow.AddMinutes($StartDelayMinutes)
    $endUtc   = $startUtc.AddHours($DurationHours)
}

if ($endUtc -le $startUtc) {
    throw 'End time must be later than start time.'
}

$startIso = $startUtc.ToString('o')
$endIso   = $endUtc.ToString('o')
$policyName = "TEMP-CA - $DisplayName"
$description = "TEMP-CA; Start=$startIso; Expires=$endIso; DesiredState=$State; AutoCleanup=$(-not $NoAutoCleanup)"

$desiredState = $State
$initialState = $State
if ($startUtc -gt [DateTime]::UtcNow -and $desiredState -ne 'disabled') {
    $initialState = 'disabled'
}

$usersCondition = @{}
if ($TargetType -eq 'User') {
    $usersCondition['includeUsers'] = @($TargetId)
    $usersCondition['includeGroups'] = @()
} else {
    $usersCondition['includeUsers'] = @()
    $usersCondition['includeGroups'] = @($TargetId)
}
$usersCondition['excludeUsers'] = @()
$usersCondition['excludeGroups'] = @()
$usersCondition['includeRoles'] = @()
$usersCondition['excludeRoles'] = @()

$grantControls = if ($Action -eq 'Block') {
    @{
        operator = 'OR'
        builtInControls = @('block')
    }
} else {
    @{
        operator = 'OR'
        builtInControls = @('mfa')
    }
}

$body = @{
    displayName = $policyName
    state       = $initialState
    conditions  = @{
        users        = $usersCondition
        applications = @{ includeApplications = @('All'); excludeApplications = @() }
        clientAppTypes = @('all')
    }
    grantControls = $grantControls
}

Write-Step 'Creating temporary Conditional Access policy'
if ($PSCmdlet.ShouldProcess($policyName, 'Create temporary conditional access policy')) {
    $policy = New-MgIdentityConditionalAccessPolicy -BodyParameter $body -ErrorAction Stop
    Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id -BodyParameter @{ description = $description } | Out-Null
    Write-Ok "Created: $policyName"
    Write-Host "      PolicyId : $($policy.Id)"
    Write-Host "      Initial  : $initialState"
    Write-Host "      Desired  : $desiredState"
    Write-Host "      Start UTC: $startIso"
    Write-Host "      End UTC  : $endIso"
}

if ($NoAutoCleanup) {
    if ($startUtc -gt [DateTime]::UtcNow -and $desiredState -ne 'disabled') {
        Write-Warn 'Policy is currently disabled and scheduled for future start. Enable it manually at start time if needed.'
    }
    Write-Warn 'Auto cleanup disabled. Remove the policy manually or run Remove-TemporaryConditionalAccessPolicies.ps1.'
    return
}

if ($startUtc -gt [DateTime]::UtcNow -and $desiredState -ne 'disabled') {
    Write-Step 'Waiting for scheduled start time'
    Wait-UntilUtc -UtcTime $startUtc -Phase 'Start'

    Write-Step 'Start time reached, enabling policy'
    Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id -BodyParameter @{ state = $desiredState } | Out-Null
    Write-Ok "Policy enabled with state: $desiredState"
}

$waitSecondsToEnd = [int][Math]::Ceiling(($endUtc - [DateTime]::UtcNow).TotalSeconds)
if ($waitSecondsToEnd -le 0) {
    Write-Warn 'End time already reached; running cleanup now.'
    Remove-TemporaryPolicyById -PolicyId $policy.Id -Name $policyName
    return
}

Write-Step 'Waiting for expiry to run immediate cleanup'
Write-Host "Policy will be removed automatically at end time (UTC): $endIso"
Wait-UntilUtc -UtcTime $endUtc -Phase 'End'

Write-Step 'Expiry reached, removing temporary policy'
Remove-TemporaryPolicyById -PolicyId $policy.Id -Name $policyName
