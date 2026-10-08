#Requires -Version 7.0
<#
.SYNOPSIS
    Report licensed users across all (or selected) GDAP/delegated-admin customer tenants.

.DESCRIPTION
    Iterates every customer tenant delegated to this partner via Granular Delegated
    Admin Privileges (GDAP) and reports licensed users and their assigned SKUs into a
    single CSV — the modern, Microsoft Graph-based replacement for the legacy
    MSOnline-based "loop over Get-MsolPartnerContract -All" pattern (MSOnline/AzureAD
    are retired and no longer function).

    Customers come from tenantRelationships/delegatedAdminCustomers in your partner
    tenant; when that list cannot be read, the script falls back to /contracts.

    Two ways to sign in (both through scripts\Startup\Connect-M365.ps1):

    Delegated (default)
        You sign in once as a partner admin in your own tenant to list the customers,
        then the script connects to each customer tenant as you, through GDAP. Your
        GDAP relationship must include a role that can read users (Global Reader,
        Directory Readers or User Administrator). Each customer is a separate sign-in:
        with the browser it usually completes from the cached account, with device code
        ($global:useDeviceCodeAuth) you enter a code per customer. The first time,
        Microsoft Graph Command Line Tools may need consent in the customer tenant.

    App-only (-ClientId with -CertificateThumbprint or -ClientSecret, or -AppOnly)
        A multi-tenant app registration of your own that is consented in every
        customer tenant (with User.Read.All as an application permission). GDAP alone
        does NOT give an app access: GDAP grants delegated rights to your users only.
        Needs -TenantId (your partner tenant).

.PARAMETER TenantId
    Your partner (home) tenant ID or domain. Required for app-only. For delegated it
    defaults to the tenant of the account you sign in with.

.PARAMETER ClientId
    Multi-tenant app registration for app-only sign-in. Omit for delegated sign-in.

.PARAMETER ClientSecret
    Client secret for -ClientId. Use -CertificateThumbprint instead where possible.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for -ClientId (preferred over a client secret).

.PARAMETER AppOnly
    App-only sign-in with the ClientId and CertificateThumbprint from graph.appid.json
    (the entry for -TenantId, your partner tenant), used for every customer.

.PARAMETER CustomerTenantId
    Restrict the report to one or more specific customer tenant IDs. If omitted, all
    customers returned by Microsoft Graph are processed.

.PARAMETER OutputPath
    CSV report path. Default: C:\Temp\MultiTenantLicenseReport_<timestamp>.csv

.EXAMPLE
    # Delegated: sign in as a partner admin, report every GDAP customer
    .\Get-MultiTenantLicenseReport.ps1

.EXAMPLE
    # App-only with a certificate
    .\Get-MultiTenantLicenseReport.ps1 -TenantId "partner.onmicrosoft.com" `
        -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"

.EXAMPLE
    # Only two specific customer tenants, delegated
    .\Get-MultiTenantLicenseReport.ps1 `
        -CustomerTenantId "cccccccc-cccc-cccc-cccc-cccccccccccc","dddddddd-dddd-dddd-dddd-dddddddddddd"

.NOTES
    Partner tenant, delegated scopes  : DelegatedAdminRelationship.Read.All, Directory.Read.All (for the /contracts fallback)
    Partner tenant, app permission    : DelegatedAdminRelationship.Read.All (or Directory.Read.All for /contracts)
    Customer tenant, delegated scope  : User.Read.All (plus a GDAP role that can read users)
    Customer tenant, app permission   : User.Read.All, consented in that tenant
    Required module : Microsoft.Graph.Authentication
#>
[CmdletBinding()]
param(
    [string] $TenantId,
    [string] $ClientId,
    [string] $ClientSecret,
    [string] $CertificateThumbprint,
    [switch] $AppOnly,
    [string[]] $CustomerTenantId,
    [string] $OutputPath
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

# ── Sign-in mode ────────────────────────────────────────────────────────────────
if ($AppOnly -and -not $ClientId) {
    $reg = Get-M365AppRegistration -TenantId $TenantId
    $ClientId = $reg.ClientId
    $CertificateThumbprint = $reg.CertificateThumbprint
    if (-not $TenantId) { $TenantId = $reg.Tenant }
}
$secret = if ($ClientSecret) { ConvertTo-SecureString $ClientSecret -AsPlainText -Force } else { $null }
if ($ClientId -and -not $CertificateThumbprint -and -not $secret) { throw "Provide -CertificateThumbprint or -ClientSecret with -ClientId." }
if ($ClientId -and -not $TenantId) { throw "App-only sign-in needs -TenantId (your partner tenant)." }

# The partner tenant. Without -TenantId, 'organizations' signs in to the tenant of the
# account you use - otherwise the helper would pick the GDAP customer from $global:cid.
$homeTenant = if ($TenantId) { $TenantId } elseif (Resolve-M365TenantId) { 'organizations' } else { $null }
$auth = @{ ClientId = $ClientId; CertificateThumbprint = $CertificateThumbprint; ClientSecret = $secret }

function Get-PartnerCustomer {
    # GDAP customers; /contracts (DAP/reseller relationships) when that list cannot be read.
    $list = [System.Collections.Generic.List[object]]::new()
    try {
        $uri = 'https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers'
        while ($uri) {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            foreach ($c in @($resp.value)) { $list.Add([pscustomobject]@{ TenantId = $c.tenantId; DisplayName = $c.displayName }) }
            $uri = $resp.'@odata.nextLink'
        }
        return $list
    } catch {
        Write-Host "  [WARN] Could not read delegatedAdminCustomers ($($_.Exception.Message)); trying /contracts." -ForegroundColor Yellow
    }
    $list.Clear()
    $uri = 'https://graph.microsoft.com/v1.0/contracts?$top=999'
    while ($uri) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        foreach ($c in @($resp.value)) { $list.Add([pscustomobject]@{ TenantId = $c.customerId; DisplayName = $c.displayName }) }
        $uri = $resp.'@odata.nextLink'
    }
    return $list
}

$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
if (-not $OutputPath) {
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path $outputDir "MultiTenantLicenseReport_$ts.csv"
}

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Multi-Tenant License Report" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "  Sign-in : $(if ($ClientId) { "app-only ($ClientId)" } else { 'delegated (GDAP)' })" -ForegroundColor DarkGray
Write-Host ""

# ── List the customers from the partner tenant ──────────────────────────────────
try {
    $homeConn = Connect-M365Graph -Scopes 'DelegatedAdminRelationship.Read.All', 'Directory.Read.All' -TenantId $homeTenant @auth
} catch {
    Write-Host "  [ERROR] Could not connect to the partner (home) tenant: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

try {
    $customers = @(Get-PartnerCustomer)
} catch {
    Write-Host "  [ERROR] Could not list customers: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-M365Graph $homeConn
    exit 1
}
Disconnect-M365Graph $homeConn

if ($CustomerTenantId) {
    $customers = @($customers | Where-Object { $_.TenantId -in $CustomerTenantId })
}

Write-Host "  Found $($customers.Count) customer tenant(s) to process." -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($customer in $customers) {
    Write-Host "  Processing $($customer.DisplayName) ($($customer.TenantId))..." -ForegroundColor Cyan

    $conn = $null
    try {
        $conn = Connect-M365Graph -Scopes 'User.Read.All' -TenantId $customer.TenantId @auth

        $users = [System.Collections.Generic.List[object]]::new()
        $uri = 'https://graph.microsoft.com/v1.0/users?$select=displayName,userPrincipalName,assignedLicenses,accountEnabled&$top=999'
        while ($uri) {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            foreach ($u in @($resp.value)) { $users.Add($u) }
            $uri = $resp.'@odata.nextLink'
        }

        $licensed = @($users | Where-Object { $_.assignedLicenses.Count -gt 0 })
        foreach ($user in $licensed) {
            $results.Add([PSCustomObject]@{
                CustomerName      = $customer.DisplayName
                TenantId          = $customer.TenantId
                DisplayName       = $user.displayName
                UserPrincipalName = $user.userPrincipalName
                AccountEnabled    = $user.accountEnabled
                LicenseCount      = $user.assignedLicenses.Count
                LicenseSkuIds     = ($user.assignedLicenses.skuId -join ', ')
            })
        }
        Write-Host "    [OK]   $($users.Count) user(s), $($licensed.Count) licensed." -ForegroundColor DarkGray
    } catch {
        Write-Host "    [WARN] Skipped: $($_.Exception.Message)" -ForegroundColor Yellow
    } finally {
        Disconnect-M365Graph $conn
    }
}

$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "  Report saved: $OutputPath ($($results.Count) row(s))" -ForegroundColor Green
Write-Host ""
