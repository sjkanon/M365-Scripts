#Requires -Version 7.0
<#
.SYNOPSIS
    Create a Dynamic Distribution Group from a job title (or other recipient) filter.

.DESCRIPTION
    Previews the recipients matched by an Exchange Online recipient filter (by job
    title by default, or a custom -RecipientFilter) and, after confirmation, creates
    a Dynamic Distribution Group using that filter.

    Exchange Online only: Microsoft Graph has no API for Dynamic Distribution Groups or
    for previewing a recipient filter, so this script uses ExchangeOnlineManagement.
    Sign-in goes through scripts\Startup\Connect-M365.ps1 (Connect-M365Exchange):
    delegated by default (you sign in as an Exchange admin; device code when
    $global:useDeviceCodeAuth is set, the GDAP customer through -DelegatedOrganization),
    app-only with -ClientId and -CertificateThumbprint (and -TenantId as the
    *.onmicrosoft.com domain), or -AppOnly. An existing Exchange session for the same
    tenant is reused and left open.

.PARAMETER JobTitle
    Job title to filter on. Builds the filter:
    "((Title -eq '<JobTitle>') -and (ExchangeUserAccountControl -ne 'AccountDisabled'))"
    Use -RecipientFilter instead for a custom filter (e.g. Department-based).

.PARAMETER RecipientFilter
    A raw Exchange recipient filter string, used instead of -JobTitle for more
    complex criteria.

.PARAMETER Name
    Name for the new group. Default: the job title (or required when using
    -RecipientFilter).

.PARAMETER PrimarySmtpAddress
    Primary SMTP address for the new group. Default: Exchange derives it from the
    alias and the tenant's default accepted domain.

.PARAMETER TenantId
    Tenant domain (contoso.onmicrosoft.com) or ID. Defaults to the GDAP customer
    tenant, else the tenant you sign in to. App-only needs the domain form.

.PARAMETER ClientId
    App registration for app-only sign-in, with -CertificateThumbprint.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.PARAMETER Apply
    Actually create the group. Without this switch, the script only previews the
    matching recipients.

.EXAMPLE
    # Preview who matches
    .\New-DynamicDistributionGroupByFilter.ps1 -JobTitle "Sales Manager"

.EXAMPLE
    .\New-DynamicDistributionGroupByFilter.ps1 -JobTitle "Sales Manager" -Apply

.EXAMPLE
    .\New-DynamicDistributionGroupByFilter.ps1 -Name "Finance Dept" `
        -RecipientFilter "((Department -eq 'Finance') -and (ExchangeUserAccountControl -ne 'AccountDisabled'))" -Apply

.NOTES
    Required role   : Exchange Administrator (or Recipient Management)
    Required module : ExchangeOnlineManagement
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $JobTitle,
    [string] $RecipientFilter,
    [string] $Name,
    [string] $PrimarySmtpAddress,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly,
    [switch] $Apply
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

if (-not $JobTitle -and -not $RecipientFilter) { throw "Provide either -JobTitle or -RecipientFilter." }
if ($RecipientFilter -and -not $Name) { throw "-Name is required when using -RecipientFilter." }

# A quote in the job title would end the OPATH string early; double it.
$filter = if ($RecipientFilter) { $RecipientFilter } else { "((Title -eq '$($JobTitle -replace "'", "''")') -and (ExchangeUserAccountControl -ne 'AccountDisabled'))" }
if (-not $Name) { $Name = $JobTitle }

Write-Host ""
Write-Host "  New-DynamicDistributionGroupByFilter : $Name" -ForegroundColor Cyan
Write-Host "  Filter : $filter"
Write-Host "  Mode   : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

try {
    $exo = Connect-M365Exchange -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly
} catch {
    Write-Host "  [ERROR] Could not connect to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

$recipients = @(Get-Recipient -RecipientPreviewFilter $filter -ResultSize Unlimited)
Write-Host "  $($recipients.Count) matching recipient(s):" -ForegroundColor DarkGray
$recipients | Select-Object DisplayName, Title, PrimarySmtpAddress | Format-Table -AutoSize | Out-Host

if (-not $Apply) {
    Write-Host "  Would create Dynamic Distribution Group '$Name' with the filter above." -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to create it." -ForegroundColor Yellow
    Write-Host ""
    Disconnect-M365Exchange $exo
    exit 0
}

if (-not $PSCmdlet.ShouldProcess($Name, "Create Dynamic Distribution Group")) { Disconnect-M365Exchange $exo; exit 0 }

$newGroupParams = @{
    Name            = $Name
    DisplayName     = "$Name group"
    Alias           = ($Name -replace '[^a-zA-Z0-9._-]', '')
    RecipientFilter = $filter
}
if ($PrimarySmtpAddress) { $newGroupParams['PrimarySmtpAddress'] = $PrimarySmtpAddress }

try {
    New-DynamicDistributionGroup @newGroupParams -ErrorAction Stop | Out-Null
    Write-Host "  [OK]   Dynamic Distribution Group '$Name' created." -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
}
Write-Host ""

Disconnect-M365Exchange $exo
