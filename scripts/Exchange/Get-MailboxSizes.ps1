#Requires -Version 7.0
<#
.SYNOPSIS
    Report mailbox sizes for all or a single mailbox.

.DESCRIPTION
    Connects to Exchange Online and retrieves mailbox statistics:
    total item size (MB/GB), item count, and quota status.
    Results are sorted by size descending and exported to CSV.

.PARAMETER Mailbox
    UPN of a single mailbox. If omitted, all user and shared mailboxes are reported.

.PARAMETER OutputPath
    CSV report path. Defaults to .\MailboxSizes_<timestamp>.csv.

.PARAMETER TenantId
    Tenant ID or domain. Defaults to the GDAP customer when load.config.ps1 sets
    authMode GDAP; otherwise you land in your own tenant. App-only needs a domain.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint). Without it
    you sign in delegated as yourself (device code per load.config.ps1).

.PARAMETER CertificateThumbprint
    Certificate for -ClientId.

.PARAMETER AppOnly
    App-only with ClientId and CertificateThumbprint for the tenant from
    graph.appid.json in the repo root.

.EXAMPLE
    .\Get-MailboxSizes.ps1

.EXAMPLE
    .\Get-MailboxSizes.ps1 -Mailbox "user@contoso.com"

.EXAMPLE
    .\Get-MailboxSizes.ps1 -OutputPath "C:\Reports\mailboxsizes.csv"
#>
[CmdletBinding()]
param(
    [string] $Mailbox,
    [string] $OutputPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
# Delegated by default (device code and GDAP customer per load.config.ps1),
# app-only with -ClientId/-CertificateThumbprint or -AppOnly. Exchange Online
# PowerShell: Graph's reports/getMailboxUsageDetail is aggregated, a day or more
# behind, and shows concealed names when the tenant hides user details in reports.
. (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')
$exo = Connect-M365Exchange -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Mailbox Size Report" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# ── Get mailboxes ─────────────────────────────────────────────────────────────
if ($Mailbox) {
    $mailboxes = @(Get-EXOMailbox -Identity $Mailbox -ErrorAction Stop)
} else {
    Write-Host "  Retrieving mailboxes..." -ForegroundColor DarkGray
    $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox)
}

Write-Host "  Getting statistics for $($mailboxes.Count) mailbox(es)..." -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($mbx in $mailboxes) {
    try {
        $stats = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -ErrorAction SilentlyContinue
        if (-not $stats) { continue }

        # Parse size — TotalItemSize format: "1.234 GB (1,325,023,744 bytes)"
        $bytes = $null
        if ($stats.TotalItemSize -match '\(([0-9,]+)\s+bytes\)') {
            $bytes = [long]($Matches[1] -replace ',', '')
        }

        $sizeMB = if ($bytes) { [math]::Round($bytes / 1MB, 2) } else { 0 }
        $sizeGB = if ($bytes) { [math]::Round($bytes / 1GB, 3) } else { 0 }

        $results.Add([PSCustomObject]@{
            DisplayName       = $mbx.DisplayName
            UserPrincipalName = $mbx.UserPrincipalName
            MailboxType       = $mbx.RecipientTypeDetails
            'SizeMB'          = $sizeMB
            'SizeGB'          = $sizeGB
            ItemCount         = $stats.ItemCount
            QuotaStatus       = $stats.StorageLimitStatus
        })
    } catch {
        Write-Host "  [WARN] $($mbx.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
$sorted = $results | Sort-Object SizeMB -Descending

if ($sorted.Count -eq 0) {
    Write-Host "  No mailbox statistics found." -ForegroundColor DarkGray
} else {
    $sorted | Format-Table DisplayName, UserPrincipalName, SizeMB, SizeGB, ItemCount, QuotaStatus -AutoSize

    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "MailboxSizes_$ts.csv"
    }

    $sorted | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

$totalGB = [math]::Round(($results | Measure-Object SizeGB -Sum).Sum, 3)
Write-Host ""
Write-Host ("  {0} mailbox(es) — total storage: {1} GB" -f $results.Count, $totalGB) -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
Disconnect-M365Exchange $exo
