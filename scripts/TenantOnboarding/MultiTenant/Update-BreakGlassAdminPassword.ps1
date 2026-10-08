#Requires -Version 7.0
<#
.SYNOPSIS
    Rotate the password of a named break-glass admin account, in one tenant or across all GDAP customers.

.DESCRIPTION
    Generates a new strong random password for a break-glass/emergency-access account
    (identified by UPN in one tenant, or by local-part on the initial
    *.onmicrosoft.com domain across every GDAP customer tenant) and sets it via
    Microsoft Graph. The modern, Graph-based replacement for the legacy MSOnline "loop
    every partner tenant and call Set-MsolUserPassword" pattern (MSOnline/AzureAD are
    retired).

    New passwords are printed once per tenant so they can be captured into your
    password manager / sealed-envelope process — never written to a file.

    Default behavior is a dry run — pass -Apply to actually change any password.

    Sign-in goes through scripts\Startup\Connect-M365.ps1:

    Delegated (default)
        Single tenant: you sign in as an admin of that tenant (-TenantId, else the GDAP
        customer from $global:cid, else your own tenant). An existing Graph session is
        reused only when it is for that tenant and holds User.ReadWrite.All.
        -AllCustomers: you sign in to your partner tenant to list the GDAP customers,
        then to each customer as yourself through GDAP. Resetting an administrator's
        password needs Privileged Authentication Administrator (or Global
        Administrator) in the GDAP relationship. Each customer is a separate sign-in;
        with device code ($global:useDeviceCodeAuth) you enter a code per customer.

    App-only (-ClientId with -CertificateThumbprint or -ClientSecret, or -AppOnly)
        A multi-tenant app registration of your own, consented in every customer
        tenant, with User.ReadWrite.All and Domain.Read.All as application permissions
        and the Privileged Authentication Administrator role assigned to its service
        principal (an app cannot reset an admin's password otherwise). GDAP alone does
        NOT give an app access: GDAP grants delegated rights to your users only.

.PARAMETER UserPrincipalNameLocalPart
    The local part (before @) of the break-glass account's UPN, e.g. "breakglass-admin".
    Combined with each tenant's initial *.onmicrosoft.com domain when -AllCustomers is used.

.PARAMETER UserPrincipalName
    Full UPN of the break-glass account in a single tenant. Use this instead of
    -AllCustomers for a one-off rotation.

.PARAMETER AllCustomers
    Rotate the password in every GDAP customer tenant returned by
    tenantRelationships/delegatedAdminCustomers (or /contracts when that cannot be read).

.PARAMETER TenantId
    Single tenant: the tenant to connect to (defaults to the GDAP customer, else the
    tenant you sign in to). With -AllCustomers: your partner (home) tenant, required for
    app-only; delegated it defaults to the tenant of the account you sign in with.

.PARAMETER ClientId
    App registration for app-only sign-in. Omit for delegated sign-in.

.PARAMETER ClientSecret
    Client secret for -ClientId.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for -ClientId (preferred over a client secret).

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.PARAMETER PasswordLength
    Length of the generated password. Default: 24.

.PARAMETER Apply
    Actually change the password(s). Without this switch, the script only reports
    which accounts would be updated.

.EXAMPLE
    # Single tenant, delegated (reuses a fitting session)
    .\Update-BreakGlassAdminPassword.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com" -Apply

.EXAMPLE
    # Every GDAP customer tenant, delegated as the partner admin
    .\Update-BreakGlassAdminPassword.ps1 -UserPrincipalNameLocalPart "breakglass-admin" -AllCustomers -Apply

.EXAMPLE
    # Every GDAP customer tenant, app-only
    .\Update-BreakGlassAdminPassword.ps1 -UserPrincipalNameLocalPart "breakglass-admin" -AllCustomers `
        -TenantId "partner.onmicrosoft.com" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" -Apply

.NOTES
    Required scopes (tenant)  : User.ReadWrite.All, Domain.Read.All (-AllCustomers)
    Required scopes (partner) : DelegatedAdminRelationship.Read.All, Directory.Read.All (-AllCustomers)
    Required module : Microsoft.Graph.Authentication, Microsoft.Graph.Users
#>
[CmdletBinding()]
param(
    [string] $UserPrincipalNameLocalPart,
    [string] $UserPrincipalName,
    [switch] $AllCustomers,
    [string] $TenantId,
    [string] $ClientId,
    [string] $ClientSecret,
    [string] $CertificateThumbprint,
    [switch] $AppOnly,
    [int] $PasswordLength = 24,
    [switch] $Apply
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

if ($AllCustomers -and -not $UserPrincipalNameLocalPart) {
    throw "-UserPrincipalNameLocalPart is required with -AllCustomers."
}
if (-not $AllCustomers -and -not $UserPrincipalName) {
    throw "-UserPrincipalName is required unless -AllCustomers is specified."
}

if ($AppOnly -and -not $ClientId -and $AllCustomers) {
    # One multi-tenant app for every customer: take the partner tenant's entry.
    $reg = Get-M365AppRegistration -TenantId $TenantId
    $ClientId = $reg.ClientId
    $CertificateThumbprint = $reg.CertificateThumbprint
    if (-not $TenantId) { $TenantId = $reg.Tenant }
}
$secret = if ($ClientSecret) { ConvertTo-SecureString $ClientSecret -AsPlainText -Force } else { $null }
if ($ClientId -and -not $CertificateThumbprint -and -not $secret) {
    throw "Provide -CertificateThumbprint or -ClientSecret with -ClientId."
}
$auth = @{ ClientId = $ClientId; CertificateThumbprint = $CertificateThumbprint; ClientSecret = $secret }

function New-RandomPassword {
    param([int] $Length = 24)
    $all   = 'abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789!@#$%^&*-_=+'
    $bytes = [System.Security.Cryptography.RandomNumberGenerator]::GetBytes($Length)
    -join (0..($Length - 1) | ForEach-Object { $all[$bytes[$_] % $all.Length] })
}

function Set-BreakGlassPassword {
    param([string] $Upn)

    $user = Get-MgUser -Filter "userPrincipalName eq '$($Upn -replace "'", "''")'" -ErrorAction SilentlyContinue
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
Write-Host "  Mode    : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host "  Sign-in : $(if ($ClientId -or $AppOnly) { 'app-only' } else { 'delegated' })" -ForegroundColor DarkGray
Write-Host ""

if (-not $AllCustomers) {
    try {
        $graph = Connect-M365Graph -Scopes 'User.ReadWrite.All' -TenantId $TenantId -AppOnly:$AppOnly @auth
    } catch {
        Write-Host "  [ERROR] Could not connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
    try {
        Set-BreakGlassPassword -Upn $UserPrincipalName
    } catch {
        Write-Host "    [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        Disconnect-M365Graph $graph
    }
    Write-Host ""
    exit 0
}

# ── -AllCustomers: enumerate GDAP customers from the home/partner tenant ───────
if ($ClientId -and -not $TenantId) { throw "App-only sign-in with -AllCustomers needs -TenantId (your partner tenant)." }
# Without -TenantId, 'organizations' signs in to the tenant of the account you use -
# otherwise the helper would pick the GDAP customer from $global:cid.
$homeTenant = if ($TenantId) { $TenantId } elseif (Resolve-M365TenantId) { 'organizations' } else { $null }

try {
    $homeConn = Connect-M365Graph -Scopes 'DelegatedAdminRelationship.Read.All', 'Directory.Read.All' -TenantId $homeTenant @auth
} catch {
    Write-Host "  [ERROR] Could not connect to the partner (home) tenant: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

$customers = [System.Collections.Generic.List[object]]::new()
try {
    $uri = 'https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers'
    while ($uri) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        foreach ($c in @($resp.value)) { $customers.Add([pscustomobject]@{ TenantId = $c.tenantId; DisplayName = $c.displayName }) }
        $uri = $resp.'@odata.nextLink'
    }
} catch {
    Write-Host "  [WARN] Could not read delegatedAdminCustomers ($($_.Exception.Message)); trying /contracts." -ForegroundColor Yellow
    $customers.Clear()
    try {
        $uri = 'https://graph.microsoft.com/v1.0/contracts?$top=999'
        while ($uri) {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            foreach ($c in @($resp.value)) { $customers.Add([pscustomobject]@{ TenantId = $c.customerId; DisplayName = $c.displayName }) }
            $uri = $resp.'@odata.nextLink'
        }
    } catch {
        Write-Host "  [ERROR] Could not list customers: $($_.Exception.Message)" -ForegroundColor Red
        Disconnect-M365Graph $homeConn
        exit 1
    }
}
Disconnect-M365Graph $homeConn

Write-Host "  Found $($customers.Count) customer tenant(s)." -ForegroundColor DarkGray
Write-Host ""

foreach ($customer in $customers) {
    Write-Host "  $($customer.DisplayName) ($($customer.TenantId))" -ForegroundColor Cyan
    $conn = $null
    try {
        $conn = Connect-M365Graph -Scopes 'User.ReadWrite.All', 'Domain.Read.All' -TenantId $customer.TenantId @auth

        # The initial domain (contoso.onmicrosoft.com) - not the first *.onmicrosoft.com
        # match, which can be contoso.mail.onmicrosoft.com.
        $domains = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/domains' -ErrorAction Stop
        $initialDomain = ($domains.value | Where-Object { $_.isInitial } | Select-Object -First 1).id
        if (-not $initialDomain) {
            Write-Host "    [WARN] No initial *.onmicrosoft.com domain found — skipping." -ForegroundColor Yellow
            continue
        }

        Set-BreakGlassPassword -Upn "$UserPrincipalNameLocalPart@$initialDomain"
    } catch {
        Write-Host "    [WARN] Skipped: $($_.Exception.Message)" -ForegroundColor Yellow
    } finally {
        Disconnect-M365Graph $conn
    }
}

Write-Host ""
if (-not $Apply) { Write-Host "  Re-run with -Apply to perform the rotations shown above." -ForegroundColor Yellow }
Write-Host ""
