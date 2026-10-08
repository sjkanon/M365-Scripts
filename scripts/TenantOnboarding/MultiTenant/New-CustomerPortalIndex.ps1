#Requires -Version 7.0
<#
.SYNOPSIS
    Generate an HTML index of admin-portal quick links for every GDAP customer tenant.

.DESCRIPTION
    Builds a single searchable HTML page listing each delegated-admin customer tenant
    with deep links into the Microsoft 365 admin center, Entra ID, Exchange admin
    center, Teams admin center, and Intune — the modern Microsoft Graph
    (tenantRelationships/delegatedAdminCustomers, GDAP) replacement for the legacy
    MSOnline "Get-MsolPartnerContract -All" partner-portal generator scripts (MSOnline
    is retired; the old direct delegated-admin "BeginClientSession.aspx" portal URLs
    are being phased out in favor of GDAP + tenant-scoped admin center URLs).
    When delegatedAdminCustomers cannot be read, the script falls back to /contracts.

    No branding/company name is hardcoded — set -Title to your own.

    Only the partner (home) tenant is read, so one sign-in is enough. Sign-in goes
    through scripts\Startup\Connect-M365.ps1: delegated by default (you sign in as a
    partner admin; device code when $global:useDeviceCodeAuth is set), app-only with
    -ClientId and -CertificateThumbprint or -ClientSecret (needs -TenantId), or -AppOnly.

.PARAMETER TenantId
    Your partner (home) tenant ID or domain. Required for app-only. For delegated it
    defaults to the tenant of the account you sign in with.

.PARAMETER ClientId
    App registration in your partner tenant for app-only sign-in. Omit for delegated.

.PARAMETER ClientSecret
    Client secret for -ClientId. Use -CertificateThumbprint instead where possible.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for -ClientId (preferred over a client secret).

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.PARAMETER Title
    Page title / heading. Default: "Customer Portal Index".

.PARAMETER OutputPath
    HTML output path. Default: C:\Temp\CustomerPortalIndex.html

.EXAMPLE
    # Delegated: sign in as a partner admin
    .\New-CustomerPortalIndex.ps1 -Title "Our Customers"

.EXAMPLE
    .\New-CustomerPortalIndex.ps1 -TenantId "partner.onmicrosoft.com" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" -Title "Our Customers"

.NOTES
    Required scopes (delegated) : DelegatedAdminRelationship.Read.All, Directory.Read.All (for the /contracts fallback)
    Required permission (app)   : DelegatedAdminRelationship.Read.All (or Directory.Read.All for /contracts)
    Required module : Microsoft.Graph.Authentication
#>
[CmdletBinding()]
param(
    [string] $TenantId,
    [string] $ClientId,
    [string] $ClientSecret,
    [string] $CertificateThumbprint,
    [switch] $AppOnly,
    [string] $Title = 'Customer Portal Index',
    [string] $OutputPath
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

if ($AppOnly -and -not $ClientId) {
    $reg = Get-M365AppRegistration -TenantId $TenantId
    $ClientId = $reg.ClientId
    $CertificateThumbprint = $reg.CertificateThumbprint
    if (-not $TenantId) { $TenantId = $reg.Tenant }
}
if ($ClientId -and -not $ClientSecret -and -not $CertificateThumbprint) {
    throw "Provide -CertificateThumbprint or -ClientSecret with -ClientId."
}
if ($ClientId -and -not $TenantId) { throw "App-only sign-in needs -TenantId (your partner tenant)." }
$secret = if ($ClientSecret) { ConvertTo-SecureString $ClientSecret -AsPlainText -Force } else { $null }

# The partner tenant. Without -TenantId, 'organizations' signs in to the tenant of the
# account you use - otherwise the helper would pick the GDAP customer from $global:cid.
$homeTenant = if ($TenantId) { $TenantId } elseif (Resolve-M365TenantId) { 'organizations' } else { $null }

$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
if (-not $OutputPath) { $OutputPath = Join-Path $outputDir 'CustomerPortalIndex.html' }

Write-Host ""
Write-Host "  Connecting to partner tenant..." -ForegroundColor Cyan
try {
    $graph = Connect-M365Graph -Scopes 'DelegatedAdminRelationship.Read.All', 'Directory.Read.All' -TenantId $homeTenant `
        -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ClientSecret $secret
} catch {
    Write-Host "  [ERROR] Could not connect to the partner tenant: $($_.Exception.Message)" -ForegroundColor Red
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
        Disconnect-M365Graph $graph
        exit 1
    }
}
Disconnect-M365Graph $graph

Write-Host "  Found $($customers.Count) customer tenant(s)." -ForegroundColor DarkGray

# Build the table by hand: ConvertTo-Html HTML-encodes cell values, which turned the
# links into literal "<a href=...>" text.
function ConvertTo-HtmlText([string] $Text) { [System.Net.WebUtility]::HtmlEncode($Text) }
function New-Link([string] $Url, [string] $Text) { "<a target=`"_blank`" href=`"$(ConvertTo-HtmlText $Url)`">$Text</a>" }

$rows = foreach ($c in ($customers | Sort-Object DisplayName)) {
    $t = [string]$c.TenantId
    $cells = @(
        (ConvertTo-HtmlText $c.DisplayName)
        (New-Link "https://admin.microsoft.com/Adminportal/Home?delegatedOrg=$t#/homepage" 'Admin Center')
        (New-Link "https://entra.microsoft.com/$t" 'Entra ID')
        (New-Link "https://admin.exchange.microsoft.com/?delegatedOrg=$t" 'Exchange')
        (New-Link "https://admin.teams.microsoft.com/?delegatedOrg=$t" 'Teams')
        (New-Link "https://intune.microsoft.com/$t" 'Intune')
        (ConvertTo-HtmlText $t)
    )
    '<tr>' + (($cells | ForEach-Object { "<td>$_</td>" }) -join '') + '</tr>'
}

$titleHtml = ConvertTo-HtmlText $Title
$html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>$titleHtml</title>
<script>
function filterRows() {
    const filter = document.querySelector('#searchBox').value.toUpperCase();
    document.querySelectorAll('table tr:not(.header)').forEach(function (tr) {
        const match = [...tr.children].some(td => td.innerText.toUpperCase().includes(filter));
        tr.style.display = match ? '' : 'none';
    });
}
</script>
<style>
body { font-family: Segoe UI, Arial, sans-serif; font-size: 10pt; background: #f5f5f5; }
table { border-collapse: collapse; width: 95%; margin: 10px auto; background: #fff; }
th, td { border: 1px solid #ddd; padding: 6px 10px; text-align: left; }
th { background: #222; color: #fff; }
tr:nth-child(even) { background: #f0f0f0; }
#searchBox { width: 50%; font-size: 14px; padding: 8px; margin: 10px auto; display: block; }
h1 { text-align: center; }
</style>
</head>
<body>
<h1>$titleHtml</h1>
<input id="searchBox" type="text" onkeyup="filterRows()" placeholder="Search customers...">
<table>
<tr class="header"><th>Customer</th><th>M365 Admin Center</th><th>Entra ID</th><th>Exchange Admin</th><th>Teams Admin</th><th>Intune</th><th>Tenant ID</th></tr>
$($rows -join "`n")
</table>
</body>
</html>
"@

$html | Out-File -FilePath $OutputPath -Encoding UTF8

Write-Host ""
Write-Host "  Index saved: $OutputPath" -ForegroundColor Green
Write-Host ""
