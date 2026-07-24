#Requires -Version 5.1
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

    No branding/company name is hardcoded — set -Title to your own.

.PARAMETER ClientId
    App registration (multi-tenant, GDAP-enabled) Client ID used to enumerate
    delegated customers from the home/partner tenant.

.PARAMETER ClientSecret
    Client secret for -ClientId.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for -ClientId (preferred over a client secret).

.PARAMETER Title
    Page title / heading. Default: "Customer Portal Index".

.PARAMETER OutputPath
    HTML output path. Default: C:\Temp\CustomerPortalIndex.html

.EXAMPLE
    .\New-CustomerPortalIndex.ps1 -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" -Title "Our Customers"

.NOTES
    Required Graph permission (on the partner tenant's app) : DelegatedAdminRelationship.Read.All
    Required module : Microsoft.Graph.Authentication
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $ClientId,

    [string] $ClientSecret,
    [string] $CertificateThumbprint,
    [string] $Title = 'Customer Portal Index',
    [string] $OutputPath
)

if (-not $ClientSecret -and -not $CertificateThumbprint) {
    throw "Provide either -ClientSecret or -CertificateThumbprint."
}

$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
if (-not $OutputPath) { $OutputPath = Join-Path $outputDir 'CustomerPortalIndex.html' }

$connectParams = @{ ClientId = $ClientId; NoWelcome = $true }
if ($CertificateThumbprint) { $connectParams['CertificateThumbprint'] = $CertificateThumbprint }
else {
    $secure = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    $connectParams['ClientSecretCredential'] = [System.Management.Automation.PSCredential]::new($ClientId, $secure)
}

Write-Host ""
Write-Host "  Connecting to partner tenant..." -ForegroundColor Cyan
Connect-MgGraph @connectParams -ErrorAction Stop

$customers = @()
$uri = 'https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers'
while ($uri) {
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    $customers += $resp.value
    $uri = $resp.'@odata.nextLink'
}
Disconnect-MgGraph | Out-Null

Write-Host "  Found $($customers.Count) customer tenant(s)." -ForegroundColor DarkGray

$rows = foreach ($c in ($customers | Sort-Object displayName)) {
    [pscustomobject]@{
        'Customer'          = $c.displayName
        'M365 Admin Center' = "<a target=`"_blank`" href=`"https://admin.microsoft.com/Adminportal/Home?delegatedOrg=$($c.tenantId)#/homepage`">Admin Center</a>"
        'Entra ID'          = "<a target=`"_blank`" href=`"https://entra.microsoft.com/$($c.tenantId)`">Entra ID</a>"
        'Exchange Admin'    = "<a target=`"_blank`" href=`"https://admin.exchange.microsoft.com/?delegatedOrg=$($c.tenantId)`">Exchange</a>"
        'Teams Admin'       = "<a target=`"_blank`" href=`"https://admin.teams.microsoft.com/?delegatedOrg=$($c.tenantId)`">Teams</a>"
        'Intune'            = "<a target=`"_blank`" href=`"https://intune.microsoft.com/$($c.tenantId)`">Intune</a>"
        'Tenant ID'         = $c.tenantId
    }
}

$head = @"
<script>
function filterRows() {
    const filter = document.querySelector('#searchBox').value.toUpperCase();
    document.querySelectorAll('table tr:not(.header)').forEach(function (tr) {
        const match = [...tr.children].some(td => td.innerText.toUpperCase().includes(filter));
        tr.style.display = match ? '' : 'none';
    });
}
</script>
<title>$Title</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; font-size: 10pt; background: #f5f5f5; }
table { border-collapse: collapse; width: 95%; margin: 10px auto; background: #fff; }
th, td { border: 1px solid #ddd; padding: 6px 10px; text-align: left; }
th { background: #222; color: #fff; }
tr:nth-child(even) { background: #f0f0f0; }
#searchBox { width: 50%; font-size: 14px; padding: 8px; margin: 10px auto; display: block; }
h1 { text-align: center; }
</style>
"@

$preContent = "<h1>$Title</h1><input id=`"searchBox`" type=`"text`" onkeyup=`"filterRows()`" placeholder=`"Search customers...`">"

$rows | ConvertTo-Html -Head $head -PreContent $preContent -As Table |
    Out-File -FilePath $OutputPath -Encoding UTF8

Write-Host ""
Write-Host "  Index saved: $OutputPath" -ForegroundColor Green
Write-Host ""
