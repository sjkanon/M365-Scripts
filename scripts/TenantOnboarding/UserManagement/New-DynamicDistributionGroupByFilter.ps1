#Requires -Version 5.1
<#
.SYNOPSIS
    Create a Dynamic Distribution Group from a job title (or other recipient) filter.

.DESCRIPTION
    Previews the recipients matched by an Exchange Online recipient filter (by job
    title by default, or a custom -RecipientFilter) and, after confirmation, creates
    a Dynamic Distribution Group using that filter. Requires an active Exchange
    Online session (Connect-ExchangeOnline); Dynamic Distribution Groups are an
    Exchange feature, not a Graph one.

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
    Primary SMTP address for the new group. Default: derived from -Name and the
    tenant's default accepted domain.

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
    Required module : ExchangeOnlineManagement
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $JobTitle,
    [string] $RecipientFilter,
    [string] $Name,
    [string] $PrimarySmtpAddress,
    [switch] $Apply
)

if (-not $JobTitle -and -not $RecipientFilter) { throw "Provide either -JobTitle or -RecipientFilter." }
if ($RecipientFilter -and -not $Name) { throw "-Name is required when using -RecipientFilter." }

$filter = if ($RecipientFilter) { $RecipientFilter } else { "((Title -eq '$JobTitle') -and (ExchangeUserAccountControl -ne 'AccountDisabled'))" }
if (-not $Name) { $Name = $JobTitle }

Write-Host ""
Write-Host "  New-DynamicDistributionGroupByFilter : $Name" -ForegroundColor Cyan
Write-Host "  Filter : $filter"
Write-Host "  Mode   : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Write-Host "  [ERROR] No active Exchange Online session. Run Connect-ExchangeOnline first." -ForegroundColor Red
    exit 1
}

$matches = Get-Recipient -RecipientPreviewFilter $filter -ResultSize Unlimited
Write-Host "  $($matches.Count) matching recipient(s):" -ForegroundColor DarkGray
$matches | Select-Object DisplayName, Title, PrimarySmtpAddress | Format-Table -AutoSize | Out-Host

if (-not $Apply) {
    Write-Host "  Would create Dynamic Distribution Group '$Name' with the filter above." -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to create it." -ForegroundColor Yellow
    Write-Host ""
    exit 0
}

if (-not $PSCmdlet.ShouldProcess($Name, "Create Dynamic Distribution Group")) { exit 0 }

$newGroupParams = @{
    Name            = $Name
    DisplayName     = "$Name group"
    Alias           = ($Name -replace '\s', '')
    RecipientFilter = $filter
}
if ($PrimarySmtpAddress) { $newGroupParams['PrimarySmtpAddress'] = $PrimarySmtpAddress }

New-DynamicDistributionGroup @newGroupParams -ErrorAction Stop | Out-Null
Write-Host "  [OK]   Dynamic Distribution Group '$Name' created." -ForegroundColor Green
Write-Host ""
