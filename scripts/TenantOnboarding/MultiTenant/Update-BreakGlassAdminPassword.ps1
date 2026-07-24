#Requires -Version 5.1
<#
.SYNOPSIS
    Rotate the password of a named break-glass admin account, in one tenant or across all GDAP customers.

.DESCRIPTION
    Generates a new strong random password for a break-glass/emergency-access account
    (identified by UPN in the current tenant, or by local-part pattern across every
    GDAP customer tenant) and sets it via Microsoft Graph. The modern, Graph-based
    replacement for the legacy MSOnline "loop every partner tenant and call
    Set-MsolUserPassword" pattern (MSOnline/AzureAD are retired).

    New passwords are printed once per tenant so they can be captured into your
    password manager / sealed-envelope process — never written to a file.

    Default behavior is a dry run — pass -Apply to actually change any password.

.PARAMETER UserPrincipalNameLocalPart
    The local part (before @) of the break-glass account's UPN, e.g. "breakglass-admin".
    Matched against each tenant's *.onmicrosoft.com domain when -AllCustomers is used.

.PARAMETER UserPrincipalName
    Full UPN of the break-glass account in a single, already-connected tenant. Use this
    instead of -AllCustomers for a one-off rotation.

.PARAMETER AllCustomers
    Rotate the password in every GDAP customer tenant returned by
    tenantRelationships/delegatedAdminCustomers, using -ClientId app-only auth.

.PARAMETER ClientId
    App registration (multi-tenant, GDAP-enabled) Client ID. Required with -AllCustomers.

.PARAMETER ClientSecret
    Client secret for -ClientId.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for -ClientId (preferred over a client secret).

.PARAMETER PasswordLength
    Length of the generated password. Default: 24.

.PARAMETER Apply
    Actually change the password(s). Without this switch, the script only reports
    which accounts would be updated.

.EXAMPLE
    # Single tenant, already connected
    .\Update-BreakGlassAdminPassword.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com" -Apply

.EXAMPLE
    # Every GDAP customer tenant
    .\Update-BreakGlassAdminPassword.ps1 -UserPrincipalNameLocalPart "breakglass-admin" -AllCustomers `
        -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" -Apply

.NOTES
    Required scopes (single tenant) : User.ReadWrite.All (or UserAuthenticationMethod.ReadWrite.All)
    Required module : Microsoft.Graph.Authentication
#>
[CmdletBinding()]
param(
    [string] $UserPrincipalNameLocalPart,
    [string] $UserPrincipalName,
    [switch] $AllCustomers,
    [string] $ClientId,
    [string] $ClientSecret,
    [string] $CertificateThumbprint,
    [int] $PasswordLength = 24,
    [switch] $Apply
)

if ($AllCustomers -and -not $UserPrincipalNameLocalPart) {
    throw "-UserPrincipalNameLocalPart is required with -AllCustomers."
}
if (-not $AllCustomers -and -not $UserPrincipalName) {
    throw "-UserPrincipalName is required unless -AllCustomers is specified."
}
if ($AllCustomers -and -not $ClientId) {
    throw "-ClientId is required with -AllCustomers."
}

function New-RandomPassword {
    param([int] $Length = 24)
    $all   = 'abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789!@#$%^&*-_=+'
    $bytes = [System.Security.Cryptography.RandomNumberGenerator]::GetBytes($Length)
    -join (0..($Length - 1) | ForEach-Object { $all[$bytes[$_] % $all.Length] })
}

function Set-BreakGlassPassword {
    param([string] $Upn)

    $user = Get-MgUser -Filter "userPrincipalName eq '$Upn'" -ErrorAction SilentlyContinue
    if (-not $user) {
        Write-Host "    [WARN] User '$Upn' not found." -ForegroundColor Yellow
        return
    }

    $newPassword = New-RandomPassword -Length $PasswordLength
    if (-not $Apply) {
        Write-Host "    [PREVIEW] Would rotate password for '$Upn'." -ForegroundColor Yellow
        return
    }

    Update-MgUser -UserId $user.Id -PasswordProfile @{
        Password                      = $newPassword
        ForceChangePasswordNextSignIn = $false
    } -ErrorAction Stop

    Write-Host "    [OK]   Password rotated for '$Upn':" -ForegroundColor Green
    Write-Host "           $newPassword" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Break-Glass Admin Password Rotation" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "  Mode : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

if (-not $AllCustomers) {
    $script:ConnectedHere = $false
    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes 'User.ReadWrite.All' -NoWelcome -ErrorAction Stop
        $script:ConnectedHere = $true
    }
    Set-BreakGlassPassword -Upn $UserPrincipalName
    if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
    Write-Host ""
    exit 0
}

# ── -AllCustomers: enumerate GDAP customers from the home/partner tenant ───────
$homeConnectParams = @{ ClientId = $ClientId; NoWelcome = $true }
if ($CertificateThumbprint) { $homeConnectParams['CertificateThumbprint'] = $CertificateThumbprint }
elseif ($ClientSecret) {
    $secure = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    $homeConnectParams['ClientSecretCredential'] = [System.Management.Automation.PSCredential]::new($ClientId, $secure)
} else {
    throw "Provide -ClientSecret or -CertificateThumbprint with -AllCustomers."
}

Connect-MgGraph @homeConnectParams -ErrorAction Stop
$customers = @()
$uri = 'https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers'
while ($uri) {
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    $customers += $resp.value
    $uri = $resp.'@odata.nextLink'
}
Disconnect-MgGraph | Out-Null

Write-Host "  Found $($customers.Count) customer tenant(s)." -ForegroundColor DarkGray
Write-Host ""

foreach ($customer in $customers) {
    Write-Host "  $($customer.displayName) ($($customer.tenantId))" -ForegroundColor Cyan
    try {
        $custParams = @{ ClientId = $ClientId; TenantId = $customer.tenantId; NoWelcome = $true }
        if ($CertificateThumbprint) { $custParams['CertificateThumbprint'] = $CertificateThumbprint }
        else {
            $secure = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
            $custParams['ClientSecretCredential'] = [System.Management.Automation.PSCredential]::new($ClientId, $secure)
        }
        Connect-MgGraph @custParams -ErrorAction Stop

        $domains = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/domains' -ErrorAction Stop
        $onMicrosoftDomain = ($domains.value | Where-Object { $_.id -like '*.onmicrosoft.com' } | Select-Object -First 1).id
        if (-not $onMicrosoftDomain) {
            Write-Host "    [WARN] No *.onmicrosoft.com domain found — skipping." -ForegroundColor Yellow
            Disconnect-MgGraph | Out-Null
            continue
        }

        Set-BreakGlassPassword -Upn "$UserPrincipalNameLocalPart@$onMicrosoftDomain"
        Disconnect-MgGraph | Out-Null
    } catch {
        Write-Host "    [WARN] Skipped: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Host ""
if (-not $Apply) { Write-Host "  Re-run with -Apply to perform the rotations shown above." -ForegroundColor Yellow }
Write-Host ""
