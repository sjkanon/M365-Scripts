#Requires -Version 5.1
<#
.SYNOPSIS
    Export a message trace (mail flow) report from Exchange Online.

.DESCRIPTION
    Connects to Exchange Online and runs a message trace for a given time window,
    optionally filtered by sender, recipient, and delivery status. Exports a summary CSV
    and, with -IncludeDetail, a per-message detail CSV (delivery hop-by-hop events) —
    useful for "did this email arrive / where did it go" mail flow diagnostics.

    Note: message trace only covers the last 10 days (a Microsoft-side limit of the
    Get-MessageTrace cmdlet). For older mail, use the Exchange admin center's historical
    search instead.

.PARAMETER Hours
    How many hours back to trace, from now. Default: 48. Ignored if -StartDate is given.

.PARAMETER StartDate
    Explicit start of the trace window.

.PARAMETER EndDate
    Explicit end of the trace window. Default: now.

.PARAMETER SenderAddress
    Filter to a specific sender address.

.PARAMETER RecipientAddress
    Filter to a specific recipient address.

.PARAMETER Status
    Filter by delivery status (e.g. Delivered, Failed, Pending, Quarantined).

.PARAMETER IncludeDetail
    Also export per-message delivery detail (one Get-MessageTraceDetail call per message
    in the summary — can be slow for a large result set).

.PARAMETER OutputPath
    Folder for the CSV report(s). Defaults to C:\Temp (Windows) / ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Get-MessageTraceReport.ps1

.EXAMPLE
    .\Get-MessageTraceReport.ps1 -SenderAddress "user@contoso.com" -Hours 24

.EXAMPLE
    .\Get-MessageTraceReport.ps1 -RecipientAddress "user@contoso.com" -Status Failed -IncludeDetail

.NOTES
    Inspired by the capability list of the retired directorcia/patron toolkit
    (o365-msgtrace-csv.ps1), rewritten from scratch with configurable filters and a
    parameterized output path — the original hardcoded a 48-hour window and a fixed
    c:\downloads output location with no filtering.
#>
[CmdletBinding()]
param(
    [int] $Hours = 48,
    [datetime] $StartDate,
    [datetime] $EndDate = (Get-Date),
    [string] $SenderAddress,
    [string] $RecipientAddress,
    [string] $Status,
    [switch] $IncludeDetail,
    [string] $OutputPath,
    [string] $TenantId
)

if (-not $StartDate) { $StartDate = $EndDate.AddHours(-$Hours) }
if ($StartDate -lt (Get-Date).AddDays(-10)) {
    Write-Host "  [WARN] -StartDate is more than 10 days ago — Get-MessageTrace only covers the last 10 days. Results will be truncated to that window." -ForegroundColor Yellow
}

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($OutputPath) { $OutputPath } elseif ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
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
Write-Host "  Window : $($StartDate.ToString('yyyy-MM-dd HH:mm')) to $($EndDate.ToString('yyyy-MM-dd HH:mm'))"
if ($SenderAddress)    { Write-Host "  Sender    : $SenderAddress" }
if ($RecipientAddress) { Write-Host "  Recipient : $RecipientAddress" }
if ($Status)           { Write-Host "  Status    : $Status" }
Write-Host ""

# ── Run trace ─────────────────────────────────────────────────────────────────
$traceParams = @{ StartDate = $StartDate; EndDate = $EndDate; ResultSize = 5000 }
if ($SenderAddress)    { $traceParams['SenderAddress'] = $SenderAddress }
if ($RecipientAddress) { $traceParams['RecipientAddress'] = $RecipientAddress }
if ($Status)           { $traceParams['Status'] = $Status }

Write-Host "  Running message trace..." -ForegroundColor DarkGray
try {
    $results = @(Get-MessageTrace @traceParams -ErrorAction Stop | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, Size, MessageId, MessageTraceId)
} catch {
    Write-Host "  [ERROR] Message trace failed: $($_.Exception.Message)" -ForegroundColor Red
    if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
    exit 1
}

Write-Host "  $($results.Count) message(s) found." -ForegroundColor DarkGray
Write-Host ""

# ── Output ────────────────────────────────────────────────────────────────────
if ($results.Count -eq 0) {
    Write-Host "  No messages matched the given window/filters." -ForegroundColor DarkGray
} else {
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $summaryPath = Join-Path $outputDir "MessageTrace_$ts.csv"
    $results | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Summary saved: $summaryPath" -ForegroundColor Green

    if ($IncludeDetail) {
        Write-Host "  Retrieving delivery detail for $($results.Count) message(s)..." -ForegroundColor DarkGray
        $detailPath = Join-Path $outputDir "MessageTrace_Detail_$ts.csv"
        $detail = [System.Collections.Generic.List[PSObject]]::new()
        $i = 0
        foreach ($msg in $results) {
            $i++
            Write-Progress -Activity "Retrieving message trace detail" -Status "$i / $($results.Count)" -PercentComplete ([int]($i / $results.Count * 100))
            try {
                Get-MessageTraceDetail -MessageTraceId $msg.MessageTraceId -RecipientAddress $msg.RecipientAddress -ErrorAction Stop | ForEach-Object { $detail.Add($_) }
            } catch {
                Write-Host "  [WARN] Could not retrieve detail for message $($msg.MessageTraceId): $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        Write-Progress -Activity "Retrieving message trace detail" -Completed
        $detail | Export-Csv -Path $detailPath -NoTypeInformation -Encoding UTF8
        Write-Host "  Detail saved: $detailPath" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host ("  {0} message(s) in trace window" -f $results.Count) -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
