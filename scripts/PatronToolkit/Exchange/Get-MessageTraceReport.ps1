#Requires -Version 7.0
<#
.SYNOPSIS
    Export a message trace (mail flow) report from Exchange Online.

.DESCRIPTION
    Connects to Exchange Online and runs a message trace for a given time window,
    optionally filtered by sender, recipient, and delivery status, with
    Get-MessageTraceV2 (the classic Get-MessageTrace / Get-MessageTraceDetail cmdlets are
    retired). Exports a summary CSV and, with -IncludeDetail, a per-message detail CSV
    (Get-MessageTraceDetailV2, delivery hop-by-hop events) — useful for "did this email
    arrive / where did it go" mail flow diagnostics.

    Message trace V2 reaches back 90 days. One call covers at most 10 days and returns at
    most 5000 rows, so the script splits the window into 10-day slices and pages through
    each slice (EndDate + StartingRecipientAddress from the last row) until it has
    everything, capped by -MaxResults.

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
    Also export per-message delivery detail (one Get-MessageTraceDetailV2 call per
    message in the summary — can be slow and throttled for a large result set).

.PARAMETER MaxResults
    Stop paging after this many trace rows. Default: 50000.

.PARAMETER OutputPath
    Folder for the CSV report(s). Defaults to C:\Temp (Windows) / ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Tenant domain (contoso.onmicrosoft.com) or ID. Defaults to the GDAP customer
    (load.config.ps1) or your own tenant. App-only sign-in needs the domain.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint). Without it the
    script signs in delegated, as you.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

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

    Sign-in: Exchange Online through scripts\Startup\Connect-M365.ps1 - delegated by
    default (device code and the GDAP customer via -DelegatedOrganization per
    load.config.ps1), app-only with -ClientId/-CertificateThumbprint or -AppOnly
    (Exchange.ManageAsApp plus an Exchange role on the app). Stays on Exchange Online:
    Microsoft Graph has no message trace API.
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
    [ValidateRange(1, 1000000)]
    [int] $MaxResults = 50000,
    [string] $OutputPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

if (-not $PSBoundParameters.ContainsKey('StartDate')) { $StartDate = $EndDate.AddHours(-$Hours) }
if ($StartDate -ge $EndDate) {
    Write-Host "  [ERROR] -StartDate must be earlier than -EndDate." -ForegroundColor Red
    exit 1
}
$oldest = (Get-Date).AddDays(-90)
if ($StartDate -lt $oldest) {
    Write-Host "  [WARN] Message trace keeps 90 days — the window is cut to start at $($oldest.ToString('yyyy-MM-dd HH:mm'))." -ForegroundColor Yellow
    $StartDate = $oldest
}

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($OutputPath) { $OutputPath } elseif ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
$exo = Connect-M365Exchange -TenantId $TenantId -ClientId $ClientId `
    -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

# The V2 cmdlets arrive with the Exchange session, so check after connecting.
foreach ($cmd in 'Get-MessageTraceV2', 'Get-MessageTraceDetailV2') {
    if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
        Write-Host "  [ERROR] $cmd is not available in this session. Update ExchangeOnlineManagement (scripts\Startup\Install-Modules.ps1)." -ForegroundColor Red
        Disconnect-M365Exchange $exo
        exit 1
    }
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
function ConvertTo-TraceUtc([datetime] $dt) {
    # Exchange returns Received in UTC without always tagging the Kind; parameters are
    # converted from local time.
    switch ($dt.Kind) {
        'Utc'         { $dt }
        'Unspecified' { [datetime]::SpecifyKind($dt, [DateTimeKind]::Utc) }
        default       { $dt.ToUniversalTime() }
    }
}

$filter = @{}
if ($SenderAddress)    { $filter['SenderAddress'] = $SenderAddress }
if ($RecipientAddress) { $filter['RecipientAddress'] = $RecipientAddress }
if ($Status)           { $filter['Status'] = $Status }

$pageSize  = 5000           # V2 maximum per call
$sliceDays = 10             # V2 maximum window per call
$rows = [System.Collections.Generic.List[PSObject]]::new()
$seen = [System.Collections.Generic.HashSet[string]]::new()
$truncated = $false

Write-Host "  Running message trace (Get-MessageTraceV2)..." -ForegroundColor DarkGray
try {
    $sliceFrom = $StartDate.ToUniversalTime()
    $windowEnd = $EndDate.ToUniversalTime()
    while ($sliceFrom -lt $windowEnd -and -not $truncated) {
        $sliceTo = $sliceFrom.AddDays($sliceDays)
        if ($sliceTo -gt $windowEnd) { $sliceTo = $windowEnd }

        # V2 returns newest first; page backwards with the last row's Received time as
        # the next EndDate and its recipient as StartingRecipientAddress.
        $cursorEnd = $sliceTo
        $cursorRecipient = $null
        while ($true) {
            $p = $filter.Clone()
            $p['StartDate'] = $sliceFrom
            $p['EndDate'] = $cursorEnd
            $p['ResultSize'] = $pageSize
            if ($cursorRecipient) { $p['StartingRecipientAddress'] = $cursorRecipient }

            $batch = @(Get-MessageTraceV2 @p -ErrorAction Stop)
            $added = 0
            foreach ($row in $batch) {
                # The continuation row can come back again; keep each message/recipient once.
                if ($seen.Add("$($row.MessageTraceId)|$($row.RecipientAddress)")) { $rows.Add($row); $added++ }
            }
            Write-Host ("  Retrieved {0,6} row(s)  [{1} -> {2} UTC]" -f $rows.Count,
                $sliceFrom.ToString('yyyy-MM-dd HH:mm'), $sliceTo.ToString('yyyy-MM-dd HH:mm')) -ForegroundColor DarkGray

            if ($rows.Count -ge $MaxResults) { $truncated = $true; break }
            if ($batch.Count -lt $pageSize) { break }
            if ($added -eq 0) {
                # A full page with nothing new would loop forever; stop this slice.
                Write-Host "  [WARN] Paging made no progress in this slice; it may be incomplete." -ForegroundColor Yellow
                break
            }
            $last = $batch[-1]
            $cursorEnd = ConvertTo-TraceUtc ([datetime]$last.Received)
            $cursorRecipient = $last.RecipientAddress
        }
        $sliceFrom = $sliceTo
    }
} catch {
    Write-Host "  [ERROR] Message trace failed: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-M365Exchange $exo
    exit 1
}

if ($truncated) {
    Write-Host "  [WARN] Stopped at -MaxResults $MaxResults rows — narrow the window or filters, or raise -MaxResults." -ForegroundColor Yellow
}
$results = @($rows | Select-Object -First $MaxResults |
    Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, Size, MessageId, MessageTraceId)

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
                Get-MessageTraceDetailV2 -MessageTraceId $msg.MessageTraceId -RecipientAddress $msg.RecipientAddress -ErrorAction Stop | ForEach-Object { $detail.Add($_) }
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
Disconnect-M365Exchange $exo
