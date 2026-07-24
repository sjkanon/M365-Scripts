#Requires -Version 5.1
<#
.SYNOPSIS
    Report mail flow (message trace) for a tenant over a recent time window.

.DESCRIPTION
    Connects to Exchange Online and retrieves message trace data — who sent
    what to whom, when, and what happened to it (delivered, quarantined,
    failed, etc.) — for the requested window and optional sender/recipient/
    status filters. Uses the newer Get-MessageTraceV2 cmdlet when available
    (falls back to the classic Get-MessageTrace on older module versions).
    Results are exported to CSV.

    Message trace only covers the last ~10 days; for older mail flow history,
    use the historical search in the Exchange admin center instead.

.PARAMETER Hours
    Number of hours back from now to search. Ignored if -StartDate is given.
    Default 48.

.PARAMETER StartDate
    Explicit start of the search window. Overrides -Hours.

.PARAMETER EndDate
    Explicit end of the search window. Default: now.

.PARAMETER SenderAddress
    Filter to a specific sender email address.

.PARAMETER RecipientAddress
    Filter to a specific recipient email address.

.PARAMETER Status
    Filter to a specific delivery status (e.g. Delivered, Failed, Quarantined,
    Pending, FilteredAsSpam).

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\ (Windows) or ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Get-MessageTraceReport.ps1

.EXAMPLE
    .\Get-MessageTraceReport.ps1 -Hours 24 -RecipientAddress "user@contoso.com"

.EXAMPLE
    .\Get-MessageTraceReport.ps1 -SenderAddress "billing@vendor.com" -Status Failed

.NOTES
    Capability inspired by o365-msgtrace.ps1 from the retired
    directorcia/Office365 (CIAOPS) toolkit, which piped raw output straight
    to Out-GridView with no filtering or export. This rewrite adds
    sender/recipient/status filters, a proper date-range parameter set, CSV
    export, and prefers the newer message trace v2 cmdlets.

    Required module: ExchangeOnlineManagement
#>
[CmdletBinding()]
param(
    [int]      $Hours = 48,
    [datetime] $StartDate,
    [datetime] $EndDate = (Get-Date),
    [string]   $SenderAddress,
    [string]   $RecipientAddress,
    [string]   $Status,
    [string]   $OutputPath,
    [string]   $TenantId
)

if (-not $StartDate) { $StartDate = $EndDate.AddHours(-$Hours) }
if ($StartDate -ge $EndDate) { throw "-StartDate must be earlier than -EndDate." }

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

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
Write-Host "   Message Trace Report" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Window : $($StartDate.ToString('yyyy-MM-dd HH:mm')) → $($EndDate.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor DarkGray
Write-Host ""

$useV2 = [bool](Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue)

$traceParams = @{ StartDate = $StartDate; EndDate = $EndDate }
if ($SenderAddress)    { $traceParams['SenderAddress']    = $SenderAddress }
if ($RecipientAddress) { $traceParams['RecipientAddress'] = $RecipientAddress }
if ($Status)            { $traceParams['Status']          = $Status }

Write-Host "  Retrieving message trace ($(if ($useV2) { 'v2' } else { 'classic' }) cmdlet)..." -ForegroundColor DarkGray
try {
    if ($useV2) {
        $raw = Get-MessageTraceV2 @traceParams -ResultSize 5000 -ErrorAction Stop
    } else {
        $raw = Get-MessageTrace @traceParams -ErrorAction Stop
    }
} catch {
    Write-Host "  [ERROR] Message trace failed: $($_.Exception.Message)" -ForegroundColor Red
    if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
    exit 1
}

$results = $raw | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, Size, MessageId, MessageTraceId |
    Sort-Object Received -Descending

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
if ($results.Count -eq 0) {
    Write-Host "  No messages found matching the given criteria." -ForegroundColor DarkGray
} else {
    $results | Select-Object Received, SenderAddress, RecipientAddress, Status, Subject | Format-Table -AutoSize

    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "MessageTrace_$ts.csv"
    }
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "  $($results.Count) message(s) found." -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ──────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
