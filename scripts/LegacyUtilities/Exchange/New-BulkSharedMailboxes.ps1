#Requires -Version 7.0
<#
.SYNOPSIS
    Bulk-create shared mailboxes from a CSV file.

.DESCRIPTION
    Reads a CSV of shared mailboxes to create (Name + PrimarySmtpAddress, optional
    Alias) and creates each one via New-Mailbox -Shared. Defaults to a safe
    preview — pass -Apply to actually create the mailboxes.

    Stays on Exchange Online PowerShell: Graph cannot create shared mailboxes;
    New-Mailbox -Shared is Exchange-only.
    Sign-in goes through scripts\Startup\Connect-M365.ps1 (Connect-M365Exchange):
    delegated as the admin by default (device code / GDAP customer via
    -DelegatedOrganization per load.config.ps1), app-only with -ClientId and
    -CertificateThumbprint or -AppOnly. An Exchange session for the tenant is reused
    and left connected; only a session this script opened is disconnected.

.PARAMETER CsvPath
    Path to a CSV with columns: Name, PrimarySmtpAddress, and optionally Alias
    (derived from the SMTP address's local part if omitted).

.PARAMETER Apply
    Actually create the mailboxes. Without this switch, the script only lists
    what would be created.

.PARAMETER OutputPath
    CSV report path. Defaults to `C:\Temp\BulkSharedMailboxes_<timestamp>.csv`
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
    .\New-BulkSharedMailboxes.ps1 -CsvPath .\sharedmailboxes.csv

.EXAMPLE
    .\New-BulkSharedMailboxes.ps1 -CsvPath .\sharedmailboxes.csv -Apply

.NOTES
    CSV example:
        Name,PrimarySmtpAddress,Alias
        "Sales Team",sales@contoso.com,sales
        "Support Desk",support@contoso.com,

    Required module: ExchangeOnlineManagement
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $CsvPath,

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
if (-not $OutputPath) { $OutputPath = Join-Path $outputDir "BulkSharedMailboxes_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" }

$rows = Import-Csv -Path $CsvPath
if (-not $rows -or -not $rows[0].PSObject.Properties.Name -contains 'Name' -or -not $rows[0].PSObject.Properties.Name -contains 'PrimarySmtpAddress') {
    Write-Error "CSV must have at least 'Name' and 'PrimarySmtpAddress' columns. Found: $($rows[0].PSObject.Properties.Name -join ', ')"
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
Write-Host "   New-BulkSharedMailboxes" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ("  Mode : {0}" -f $(if ($Apply) { 'Apply' } else { 'Preview only' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host "  Rows : $($rows.Count)"
Write-Host ""

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($row in $rows) {
    $name  = $row.Name
    $smtp  = $row.PrimarySmtpAddress
    $alias = if ($row.PSObject.Properties.Name -contains 'Alias' -and $row.Alias) { $row.Alias } else { ($smtp -split '@')[0] }

    if (-not $name -or -not $smtp) {
        Write-Host "  [SKIP] Row missing Name or PrimarySmtpAddress." -ForegroundColor Yellow
        continue
    }

    if (-not $Apply) {
        Write-Host "  [PREVIEW] $name <$smtp> (alias: $alias)" -ForegroundColor DarkGray
        $results.Add([PSCustomObject]@{ Name = $name; PrimarySmtpAddress = $smtp; Alias = $alias; Status = 'Preview' })
        continue
    }

    if (-not $PSCmdlet.ShouldProcess($smtp, "Create shared mailbox '$name'")) { continue }

    try {
        $existing = Get-Mailbox -Identity $smtp -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Host "  [SKIP] $smtp already exists." -ForegroundColor Yellow
            $results.Add([PSCustomObject]@{ Name = $name; PrimarySmtpAddress = $smtp; Alias = $alias; Status = 'AlreadyExists' })
            continue
        }

        New-Mailbox -Shared -Name $name -DisplayName $name -PrimarySmtpAddress $smtp -Alias $alias -ErrorAction Stop | Out-Null
        Write-Host "  [OK]   Created $name <$smtp>" -ForegroundColor Green
        $results.Add([PSCustomObject]@{ Name = $name; PrimarySmtpAddress = $smtp; Alias = $alias; Status = 'Created' })
    } catch {
        Write-Host "  [WARN] $smtp : $($_.Exception.Message)" -ForegroundColor Yellow
        $results.Add([PSCustomObject]@{ Name = $name; PrimarySmtpAddress = $smtp; Alias = $alias; Status = "Error: $($_.Exception.Message)" })
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
if (-not $Apply) { Write-Host "  Re-run with -Apply to create these mailboxes." -ForegroundColor Yellow }
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
Disconnect-M365Exchange $exo
