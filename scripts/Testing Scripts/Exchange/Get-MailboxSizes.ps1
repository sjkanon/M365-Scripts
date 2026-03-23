#Requires -Version 5.1
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
    Entra ID tenant ID or domain. Optional if already connected.

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
    [string] $TenantId
)

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
} catch {
    $connectParams = @{ ShowBanner = $false }
    if ($TenantId) { $connectParams['Organization'] = $TenantId }
    Connect-ExchangeOnline @connectParams
    $script:ConnectedHere = $true
}

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
        $OutputPath = Join-Path (Get-Location) "MailboxSizes_$ts.csv"
    }

    $sorted | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

$totalGB = [math]::Round(($results | Measure-Object SizeGB -Sum).Sum, 3)
Write-Host ""
Write-Host ("  {0} mailbox(es) — total storage: {1} GB" -f $results.Count, $totalGB) -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
