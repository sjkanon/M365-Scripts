#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.SignIns
<#
.SYNOPSIS
    Create a Temporary Access Pass (TAP) for a user.

.DESCRIPTION
    Creates a TAP via Microsoft Graph and returns the pass code and validity.

.PARAMETER UserId
    User object ID or UPN.

.PARAMETER LifetimeMinutes
    TAP lifetime in minutes.

.PARAMETER IsUsableOnce
    If set, TAP can be used once (recommended).

.PARAMETER TenantId
    Optional tenant ID/domain for Connect-MgGraph.

.EXAMPLE
    .\New-UserTemporaryAccessPass.ps1 -UserId "user@contoso.com" -LifetimeMinutes 60 -IsUsableOnce
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$UserId,

    [ValidateRange(10, 43200)]
    [int]$LifetimeMinutes = 60,

    [switch]$IsUsableOnce,

    [string]$TenantId
)

$ErrorActionPreference = 'Stop'

function Write-Step { param([string]$Message) Write-Host "`n=== $Message ===" -ForegroundColor Cyan }
function Write-Ok   { param([string]$Message) Write-Host "[OK]   $Message" -ForegroundColor Green }

$requiredScopes = @(
    'UserAuthenticationMethod.ReadWrite.All',
    'User.Read.All'
)

$ctx = Get-MgContext -ErrorAction SilentlyContinue
$missingScope = $true
if ($ctx -and $ctx.Scopes) {
    $missingScope = ($requiredScopes | Where-Object { $_ -notin $ctx.Scopes }).Count -gt 0
}

if (-not $ctx -or $missingScope) {
    Write-Step 'Connecting to Microsoft Graph for TAP creation'
    $connectParams = @{
        Scopes       = $requiredScopes
        ContextScope = 'Process'
        NoWelcome    = $true
    }
    if ($TenantId) {
        $connectParams['TenantId'] = $TenantId
    }
    Connect-MgGraph @connectParams | Out-Null
}

Write-Step 'Creating Temporary Access Pass'
$body = @{
    lifetimeInMinutes = $LifetimeMinutes
    isUsableOnce      = $IsUsableOnce.IsPresent
}

$tap = New-MgUserAuthenticationTemporaryAccessPassMethod -UserId $UserId -BodyParameter $body -ErrorAction Stop

$startUtc = if ($tap.StartDateTime) { [DateTime]::Parse($tap.StartDateTime).ToUniversalTime() } else { [DateTime]::UtcNow }
$endUtc   = $startUtc.AddMinutes($LifetimeMinutes)

Write-Ok "Temporary Access Pass created for user: $UserId"
Write-Host "      TAP Code  : $($tap.TemporaryAccessPass)"
Write-Host "      Start UTC : $($startUtc.ToString('o'))"
Write-Host "      End UTC   : $($endUtc.ToString('o'))"
Write-Host "      One-time  : $($IsUsableOnce.IsPresent)"

[PSCustomObject]@{
    UserId             = $UserId
    TemporaryAccessPass = $tap.TemporaryAccessPass
    StartUtc           = $startUtc
    EndUtc             = $endUtc
    IsUsableOnce       = $IsUsableOnce.IsPresent
}
