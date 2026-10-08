#Requires -Version 7.0
<#
.SYNOPSIS
    Bulk-create Mail Contacts from a CSV file, optionally adding each to a
    distribution group.

.DESCRIPTION
    Reads a CSV of external contacts (Name + ExternalEmailAddress) and creates a
    Mail Contact for each one via New-MailContact. Optionally adds every created
    contact to a distribution group in the same pass. Defaults to a safe
    preview — pass -Apply to actually create contacts.

    Stays on Exchange Online PowerShell: mail contacts (Exchange recipients) and
    distribution-group membership are managed in Exchange; Graph orgContacts are
    read-only and Graph cannot write distribution-group members.
    Sign-in goes through scripts\Startup\Connect-M365.ps1 (Connect-M365Exchange):
    delegated as the admin by default (device code / GDAP customer via
    -DelegatedOrganization per load.config.ps1), app-only with -ClientId and
    -CertificateThumbprint or -AppOnly. An Exchange session for the tenant is reused
    and left connected; only a session this script opened is disconnected.

.PARAMETER CsvPath
    Path to a CSV with columns: Name, ExternalEmailAddress.

.PARAMETER DistributionGroup
    Optional identity (name, alias, or SMTP address) of a distribution group to
    add every created contact to.

.PARAMETER Apply
    Actually create the contacts (and group memberships). Without this switch,
    the script only lists what would be created.

.PARAMETER OutputPath
    CSV report path. Defaults to `C:\Temp\BulkMailContacts_<timestamp>.csv`
    (`~/Downloads` on Linux/macOS).

.PARAMETER TenantId
    Tenant domain (contoso.onmicrosoft.com) or ID. Defaults to the GDAP customer when
    authMode is GDAP. App-only needs the domain form.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint and -TenantId).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.EXAMPLE
    # Preview
    .\New-BulkMailContacts.ps1 -CsvPath .\contacts.csv

.EXAMPLE
    .\New-BulkMailContacts.ps1 -CsvPath .\contacts.csv -Apply

.EXAMPLE
    # Also add each new contact to a distribution group
    .\New-BulkMailContacts.ps1 -CsvPath .\contacts.csv -DistributionGroup "everyone@contoso.com" -Apply

.NOTES
    CSV example:
        Name,ExternalEmailAddress
        "Jane Doe",jane.doe@example.com

    Required module: ExchangeOnlineManagement
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $CsvPath,

    [string] $DistributionGroup,
    [switch] $Apply,
    [string] $OutputPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
if (-not $OutputPath) { $OutputPath = Join-Path $outputDir "BulkMailContacts_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" }

$rows = Import-Csv -Path $CsvPath
if (-not $rows -or -not $rows[0].PSObject.Properties.Name -contains 'Name' -or -not $rows[0].PSObject.Properties.Name -contains 'ExternalEmailAddress') {
    Write-Error "CSV must have 'Name' and 'ExternalEmailAddress' columns. Found: $($rows[0].PSObject.Properties.Name -join ', ')"
    exit 1
}

# ── Connection ────────────────────────────────────────────────────────────────
# -Organization (used here before) only applies to app-only sign-in, so a GDAP
# customer was never reached; the helper uses -DelegatedOrganization for that.
$exo = Connect-M365Exchange -TenantId $TenantId -ClientId $ClientId `
    -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   New-BulkMailContacts" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ("  Mode : {0}" -f $(if ($Apply) { 'Apply' } else { 'Preview only' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host "  Rows : $($rows.Count)"
if ($DistributionGroup) { Write-Host "  Distribution group : $DistributionGroup" }
Write-Host ""

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($row in $rows) {
    $name  = $row.Name
    $email = $row.ExternalEmailAddress

    if (-not $name -or -not $email) {
        Write-Host "  [SKIP] Row missing Name or ExternalEmailAddress." -ForegroundColor Yellow
        continue
    }

    if (-not $Apply) {
        Write-Host "  [PREVIEW] $name <$email>" -ForegroundColor DarkGray
        $results.Add([PSCustomObject]@{ Name = $name; ExternalEmailAddress = $email; Status = 'Preview' })
        continue
    }

    if (-not $PSCmdlet.ShouldProcess($email, "Create mail contact '$name'")) { continue }

    try {
        $existing = Get-MailContact -Identity $email -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Host "  [SKIP] $email already exists as a contact." -ForegroundColor Yellow
            $results.Add([PSCustomObject]@{ Name = $name; ExternalEmailAddress = $email; Status = 'AlreadyExists' })
            continue
        }

        $contact = New-MailContact -Name $name -ExternalEmailAddress $email -ErrorAction Stop
        Write-Host "  [OK]   Created $name <$email>" -ForegroundColor Green
        $status = 'Created'

        if ($DistributionGroup) {
            try {
                Add-DistributionGroupMember -Identity $DistributionGroup -Member $contact.Identity -ErrorAction Stop
                $status = 'Created; AddedToGroup'
            } catch {
                $status = "Created; GroupError: $($_.Exception.Message)"
            }
        }

        $results.Add([PSCustomObject]@{ Name = $name; ExternalEmailAddress = $email; Status = $status })
    } catch {
        Write-Host "  [WARN] $email : $($_.Exception.Message)" -ForegroundColor Yellow
        $results.Add([PSCustomObject]@{ Name = $name; ExternalEmailAddress = $email; Status = "Error: $($_.Exception.Message)" })
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
if (-not $Apply) { Write-Host "  Re-run with -Apply to create these contacts." -ForegroundColor Yellow }
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
Disconnect-M365Exchange $exo
