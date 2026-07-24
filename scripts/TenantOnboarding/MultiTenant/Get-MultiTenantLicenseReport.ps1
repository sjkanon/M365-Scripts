#Requires -Version 5.1
<#
.SYNOPSIS
    Report licensed users across all (or selected) GDAP/delegated-admin customer tenants.

.DESCRIPTION
    Iterates every customer tenant delegated to this partner via Granular Delegated
    Admin Privileges (GDAP) and reports licensed users and their assigned SKUs into a
    single CSV — the modern, Microsoft Graph-based replacement for the legacy
    MSOnline-based "loop over Get-MsolPartnerContract -All" pattern (MSOnline/AzureAD
    are retired and no longer function).

    For each customer tenant, the script connects app-only using the partner's own
    multi-tenant app registration (the same one used for GDAP admin-on-behalf-of
    access) and reads licensed users via Microsoft Graph.

.PARAMETER ClientId
    App registration (multi-tenant, GDAP-enabled) Client ID used to connect to each
    customer tenant on your behalf.

.PARAMETER ClientSecret
    Client secret for -ClientId. Use -CertificateThumbprint instead where possible.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for -ClientId (preferred over a client secret).

.PARAMETER CustomerTenantId
    Restrict the report to one or more specific customer tenant IDs. If omitted, all
    delegatedAdminCustomers returned by Microsoft Graph are processed.

.PARAMETER OutputPath
    CSV report path. Default: C:\Temp\MultiTenantLicenseReport_<timestamp>.csv

.EXAMPLE
    .\Get-MultiTenantLicenseReport.ps1 -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"

.EXAMPLE
    # Only two specific customer tenants
    .\Get-MultiTenantLicenseReport.ps1 -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ClientSecret "your-client-secret" `
        -CustomerTenantId "cccccccc-cccc-cccc-cccc-cccccccccccc","dddddddd-dddd-dddd-dddd-dddddddddddd"

.NOTES
    Required Graph permission (on the partner tenant's app) : DelegatedAdminRelationship.Read.All
    Required Graph permission (application, per customer via GDAP) : User.Read.All, Organization.Read.All
    Required module : Microsoft.Graph.Authentication
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $ClientId,

    [string] $ClientSecret,
    [string] $CertificateThumbprint,
    [string[]] $CustomerTenantId,
    [string] $OutputPath
)

if (-not $ClientSecret -and -not $CertificateThumbprint) {
    throw "Provide either -ClientSecret or -CertificateThumbprint."
}

$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
if (-not $OutputPath) {
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path $outputDir "MultiTenantLicenseReport_$ts.csv"
}

# ── Connect to the HOME/partner tenant to enumerate delegated customers ────────
$homeConnectParams = @{ ClientId = $ClientId; NoWelcome = $true }
if ($CertificateThumbprint) { $homeConnectParams['CertificateThumbprint'] = $CertificateThumbprint }
else {
    $secure = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    $homeConnectParams['ClientSecretCredential'] = [System.Management.Automation.PSCredential]::new($ClientId, $secure)
}

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Multi-Tenant License Report" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

try {
    Connect-MgGraph @homeConnectParams -ErrorAction Stop
} catch {
    Write-Host "  [ERROR] Could not connect to the partner (home) tenant: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

try {
    $customers = @()
    $uri = 'https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers'
    while ($uri) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        $customers += $resp.value
        $uri = $resp.'@odata.nextLink'
    }
} catch {
    Write-Host "  [ERROR] Could not list delegated admin customers: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit 1
}

if ($CustomerTenantId) {
    $customers = $customers | Where-Object { $_.tenantId -in $CustomerTenantId }
}

Write-Host "  Found $($customers.Count) customer tenant(s) to process." -ForegroundColor DarkGray
Write-Host ""

Disconnect-MgGraph | Out-Null

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($customer in $customers) {
    Write-Host "  Processing $($customer.displayName) ($($customer.tenantId))..." -ForegroundColor Cyan

    try {
        $custParams = @{ ClientId = $ClientId; TenantId = $customer.tenantId; NoWelcome = $true }
        if ($CertificateThumbprint) { $custParams['CertificateThumbprint'] = $CertificateThumbprint }
        else {
            $secure = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
            $custParams['ClientSecretCredential'] = [System.Management.Automation.PSCredential]::new($ClientId, $secure)
        }
        Connect-MgGraph @custParams -ErrorAction Stop

        $users = @()
        $uri = 'https://graph.microsoft.com/v1.0/users?$select=displayName,userPrincipalName,assignedLicenses,accountEnabled&$top=999'
        while ($uri) {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            $users += $resp.value
            $uri = $resp.'@odata.nextLink'
        }

        foreach ($user in ($users | Where-Object { $_.assignedLicenses.Count -gt 0 })) {
            $results.Add([PSCustomObject]@{
                CustomerName      = $customer.displayName
                TenantId          = $customer.tenantId
                DisplayName       = $user.displayName
                UserPrincipalName = $user.userPrincipalName
                AccountEnabled    = $user.accountEnabled
                LicenseCount      = $user.assignedLicenses.Count
                LicenseSkuIds     = ($user.assignedLicenses.skuId -join ', ')
            })
        }
        Write-Host "    [OK]   $($users.Count) user(s), $((($users | Where-Object { $_.assignedLicenses.Count -gt 0 })).Count) licensed." -ForegroundColor DarkGray
        Disconnect-MgGraph | Out-Null
    } catch {
        Write-Host "    [WARN] Skipped: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "  Report saved: $OutputPath ($($results.Count) row(s))" -ForegroundColor Green
Write-Host ""
