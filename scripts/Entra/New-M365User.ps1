#Requires -Version 5.1
<#
.SYNOPSIS
    Create a new Microsoft 365 user via Microsoft Graph.

.DESCRIPTION
    Creates a single M365 user account in Entra ID. If no password is supplied,
    a random 16-character password is generated. A usage location is required
    before a license can be assigned.

    The new account is enabled by default with ForceChangePasswordNextSignIn set
    unless -NoPasswordReset is specified.

.PARAMETER UserPrincipalName
    UPN for the new user (e.g. j.doe@contoso.com).

.PARAMETER DisplayName
    Display name shown in the GAL and Azure portal.

.PARAMETER GivenName
    First name.

.PARAMETER Surname
    Last name.

.PARAMETER Password
    Initial password. If omitted, a 16-character random password is generated.

.PARAMETER UsageLocation
    Two-letter ISO country code for license assignment. Default: NL.

.PARAMETER Department
    Department field.

.PARAMETER JobTitle
    Job title field.

.PARAMETER MobilePhone
    Mobile phone number.

.PARAMETER LicenseSkuId
    SKU ID of the license to assign after creation (e.g. O365_BUSINESS_PREMIUM).
    Run Get-MgSubscribedSku to list available SKU IDs.

.PARAMETER NoPasswordReset
    Do not force a password change on first sign-in.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\New-M365User.ps1 -UserPrincipalName "j.doe@contoso.com" -DisplayName "Jane Doe"

.EXAMPLE
    .\New-M365User.ps1 -UserPrincipalName "j.doe@contoso.com" -DisplayName "Jane Doe" `
        -GivenName "Jane" -Surname "Doe" -Department "Finance" -JobTitle "Controller" `
        -LicenseSkuId "ENTERPRISEPACK"
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)]
    [string] $UserPrincipalName,

    [Parameter(Mandatory)]
    [string] $DisplayName,

    [string] $GivenName,
    [string] $Surname,
    [string] $Password,
    [string] $UsageLocation = 'NL',
    [string] $Department,
    [string] $JobTitle,
    [string] $MobilePhone,
    [string] $LicenseSkuId,
    [switch] $NoPasswordReset,
    [string] $TenantId
)

# ── Password generator ────────────────────────────────────────────────────────
function New-RandomPassword {
    $upper   = 'ABCDEFGHJKLMNPQRSTUVWXYZ'.ToCharArray()
    $lower   = 'abcdefghjkmnpqrstuvwxyz'.ToCharArray()
    $digits  = '23456789'.ToCharArray()
    $special = '!@#$%&*'.ToCharArray()
    $all     = $upper + $lower + $digits + $special

    # Guarantee at least one of each category
    $chars  = @(
        $upper  | Get-Random
        $lower  | Get-Random
        $digits | Get-Random
        $special| Get-Random
    )
    # Fill remaining 12 from full set
    $chars += 1..12 | ForEach-Object { $all | Get-Random }

    # Shuffle
    ($chars | Sort-Object { Get-Random }) -join ''
}

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
$ctx = Get-MgContext -ErrorAction SilentlyContinue
if (-not $ctx) {
    $connectParams = @{
        Scopes    = @('User.ReadWrite.All', 'Directory.ReadWrite.All')
        NoWelcome = $true
    }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
}

# ── Resolve password ──────────────────────────────────────────────────────────
$generated = $false
if (-not $Password) {
    $Password  = New-RandomPassword
    $generated = $true
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   New M365 User" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  UPN             : $UserPrincipalName"
Write-Host "  Display name    : $DisplayName"
if ($GivenName)     { Write-Host "  First name      : $GivenName" }
if ($Surname)       { Write-Host "  Last name       : $Surname" }
if ($Department)    { Write-Host "  Department      : $Department" }
if ($JobTitle)      { Write-Host "  Job title       : $JobTitle" }
if ($MobilePhone)   { Write-Host "  Mobile          : $MobilePhone" }
Write-Host "  Usage location  : $UsageLocation"
if ($LicenseSkuId)  { Write-Host "  License         : $LicenseSkuId" }
Write-Host ""

# ── Build user params ─────────────────────────────────────────────────────────
$userParams = @{
    UserPrincipalName         = $UserPrincipalName
    DisplayName               = $DisplayName
    AccountEnabled            = $true
    UsageLocation             = $UsageLocation
    PasswordProfile           = @{
        Password                      = $Password
        ForceChangePasswordNextSignIn = -not $NoPasswordReset
    }
}
if ($GivenName)   { $userParams['GivenName']   = $GivenName }
if ($Surname)     { $userParams['Surname']      = $Surname }
if ($Department)  { $userParams['Department']   = $Department }
if ($JobTitle)    { $userParams['JobTitle']     = $JobTitle }
if ($MobilePhone) { $userParams['MobilePhone']  = $MobilePhone }

# ── Create user ───────────────────────────────────────────────────────────────
try {
    $newUser = New-MgUser @userParams -ErrorAction Stop
    Write-Host "  User created    : $($newUser.UserPrincipalName)" -ForegroundColor Green
    Write-Host "  Object ID       : $($newUser.Id)" -ForegroundColor DarkGray
    if ($generated) {
        Write-Host "  Password        : $Password" -ForegroundColor Yellow
        Write-Host "  (generated — note this down, it will not be shown again)" -ForegroundColor DarkYellow
    }
} catch {
    Write-Host "  [ERROR] Failed to create user: $($_.Exception.Message)" -ForegroundColor Red
    if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
    exit 1
}

# ── Assign license ────────────────────────────────────────────────────────────
if ($LicenseSkuId) {
    Write-Host ""
    Write-Host "  Assigning license '$LicenseSkuId'..." -ForegroundColor DarkGray

    try {
        $sku = Get-MgSubscribedSku -ErrorAction Stop |
               Where-Object { $_.SkuPartNumber -eq $LicenseSkuId -or $_.SkuId -eq $LicenseSkuId } |
               Select-Object -First 1

        if (-not $sku) {
            Write-Host "  [WARN] SKU '$LicenseSkuId' not found in this tenant. License not assigned." -ForegroundColor Yellow
            Write-Host "  Available SKUs: $(((Get-MgSubscribedSku).SkuPartNumber) -join ', ')" -ForegroundColor DarkGray
        } else {
            Set-MgUserLicense -UserId $newUser.Id -BodyParameter @{
                AddLicenses    = @(@{ SkuId = $sku.SkuId })
                RemoveLicenses = @()
            } | Out-Null
            Write-Host "  License assigned: $($sku.SkuPartNumber)" -ForegroundColor Green
        }
    } catch {
        Write-Host "  [WARN] License assignment failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Done" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
