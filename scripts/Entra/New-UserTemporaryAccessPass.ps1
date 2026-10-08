#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.SignIns
<#
.SYNOPSIS
    Create a Temporary Access Pass (TAP) for a user.

.DESCRIPTION
    Creates a TAP via Microsoft Graph and returns the pass code and validity.

    Sign-in goes through scripts\Startup\Connect-M365.ps1: delegated as the admin by
    default (device code / GDAP customer per load.config.ps1), app-only with -ClientId
    and -CertificateThumbprint or -AppOnly. A fitting Graph session is reused and left
    open, as before. Delegated scopes: UserAuthenticationMethod.ReadWrite.All, User.Read.All.

.PARAMETER UserId
    User object ID or UPN.

.PARAMETER LifetimeMinutes
    TAP lifetime in minutes.

.PARAMETER IsUsableOnce
    If set, TAP can be used once (recommended).

.PARAMETER TenantId
    Optional tenant ID/domain. Defaults to the GDAP customer when authMode is GDAP.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint and -TenantId).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

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

    [string]$TenantId,

    [string]$ClientId,

    [string]$CertificateThumbprint,

    [switch]$AppOnly
)

. (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')

$ErrorActionPreference = 'Stop'

function Write-Step { param([string]$Message) Write-Host "`n=== $Message ===" -ForegroundColor Cyan }
function Write-Ok   { param([string]$Message) Write-Host "[OK]   $Message" -ForegroundColor Green }

# Reuses a session for the right tenant that has the scopes; connects otherwise.
$null = Connect-M365Graph -Scopes 'UserAuthenticationMethod.ReadWrite.All', 'User.Read.All' `
    -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

Write-Step 'Creating Temporary Access Pass'
$body = @{
    lifetimeInMinutes = $LifetimeMinutes
    isUsableOnce      = $IsUsableOnce.IsPresent
}

$tap = New-MgUserAuthenticationTemporaryAccessPassMethod -UserId $UserId -BodyParameter $body -ErrorAction Stop

# StartDateTime is already a [datetime] in SDK v2; re-parsing its string form
# depended on the current culture.
$startUtc = if ($tap.StartDateTime) { ([datetime]$tap.StartDateTime).ToUniversalTime() } else { [DateTime]::UtcNow }
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
