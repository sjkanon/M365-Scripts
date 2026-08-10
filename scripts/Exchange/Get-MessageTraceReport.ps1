#Requires -Version 5.1
<#
.SYNOPSIS
    Trace exactly where mail went: date/time, sender, recipient, status and the
    address it was forwarded/redirected to.

.DESCRIPTION
    Runs a message trace against Exchange Online and answers the question
    "who received this, when exactly, and where did it go next?".

    Reports per message:
      - Exact timestamp in both local time and UTC (message trace stores UTC)
      - SenderAddress and RecipientAddress
      - Subject, Status, message size, originating/delivering IP
      - MessageId and MessageTraceId (for follow-up lookups)
      - ForwardedTo — the address(es) the message ended up at, detected three ways:
          1. Other recipients sharing the same MessageId (SMTP forward / redirect)
          2. Redirect / transport-rule hops from Get-MessageTraceDetailV2 (-IncludeDetails)
          3. A follow-up message sent BY the recipient with the same subject
             (client-side "FW:" forward, which gets a brand new MessageId)

    Also reports the *configured* forwarding for every internal mailbox involved
    (ForwardingSMTPAddress / ForwardingAddress plus inbox rules with
    ForwardTo / RedirectTo / ForwardAsAttachmentTo), so a forward that has not
    fired yet in the traced window is still visible.

    Uses Get-MessageTraceV2 when available and falls back to the retired
    Get-MessageTrace on older module versions. Long ranges are split into
    10-day chunks automatically (the V2 limit) and every chunk is paginated
    until exhausted.

.PARAMETER Mailbox
    Trace both directions for this address: everything it sent AND everything it
    received. Also pulls that mailbox's forwarding configuration and inbox rules.

.PARAMETER Sender
    Filter on sender address. Combine with -Recipient to trace one specific flow.

.PARAMETER Recipient
    Filter on recipient address.

.PARAMETER Subject
    Client-side subject filter, wildcards allowed (e.g. "*invoice*"). Message
    trace itself cannot filter on subject, so this is applied after retrieval.

.PARAMETER MessageId
    Internet MessageId to trace, with or without angle brackets.

.PARAMETER Days
    Number of days back from -EndDate to search. Ignored if -StartDate is given.
    Default 2. Message trace retains 90 days.

.PARAMETER StartDate
    Explicit start of the search window (local time). Overrides -Days.

.PARAMETER EndDate
    Explicit end of the search window (local time). Default: now.

.PARAMETER Status
    Filter on delivery status: Delivered, Failed, Pending, Expanded,
    Quarantined, FilteredAsSpam, GettingStatus, None.

.PARAMETER IncludeDetails
    Retrieve the per-hop delivery detail (Get-MessageTraceDetailV2) for each
    traced message. This is what exposes Redirect / transport-rule forwarding
    targets. Slow and throttled — capped by -MaxDetailLookups.

.PARAMETER MaxDetailLookups
    Maximum number of messages to pull hop details for. Default 50. If more
    messages match, the truncation is reported explicitly.

.PARAMETER SkipForwardingConfig
    Do not inspect mailbox forwarding settings and inbox rules.

.PARAMETER OutputPath
    Main CSV report path. Defaults to C:\Temp\ (Windows) or ~/Downloads.
    The detail and forwarding-config reports are written next to it with
    _Details / _ForwardingConfig suffixes.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    # Everything one mailbox sent and received in the last 2 days, incl. forwards
    .\Get-MessageTraceReport.ps1 -Mailbox "user@contoso.com"

.EXAMPLE
    # One specific flow over the last 30 days, with per-hop detail
    .\Get-MessageTraceReport.ps1 -Sender "boss@contoso.com" -Recipient "user@contoso.com" -Days 30 -IncludeDetails

.EXAMPLE
    # Where did this specific message end up?
    .\Get-MessageTraceReport.ps1 -MessageId "<abc123@contoso.com>" -Days 10 -IncludeDetails

.EXAMPLE
    # All failed mail from one sender in an explicit window
    .\Get-MessageTraceReport.ps1 -Sender "noreply@contoso.com" -Status Failed `
        -StartDate (Get-Date "2026-08-01") -EndDate (Get-Date "2026-08-08")

.NOTES
    Message trace covers the last 90 days. Timestamps returned by Exchange are
    UTC; both UTC and converted local time are reported.

    A client-side forward (Outlook "Forward" button or a client-only rule)
    creates a NEW message with a new MessageId — those are matched on subject,
    which is a heuristic, and are labelled as such in the report.

    Required module: ExchangeOnlineManagement
#>
[CmdletBinding()]
param(
    [string]   $Mailbox,
    [string]   $Sender,
    [string]   $Recipient,
    [string]   $Subject,
    [string]   $MessageId,
    [int]      $Days = 2,
    [datetime] $StartDate,
    [datetime] $EndDate = (Get-Date),
    [ValidateSet('None','GettingStatus','Failed','Pending','Delivered','Expanded','Quarantined','FilteredAsSpam')]
    [string]   $Status,
    [switch]   $IncludeDetails,
    [int]      $MaxDetailLookups = 50,
    [switch]   $SkipForwardingConfig,
    [string]   $OutputPath,
    [string]   $TenantId
)

if (-not $StartDate) { $StartDate = $EndDate.AddDays(-$Days) }
if ($StartDate -ge $EndDate) { throw "-StartDate must be earlier than -EndDate." }
if ($StartDate -lt (Get-Date).AddDays(-90)) {
    Write-Warning "Message trace only retains 90 days — results before $((Get-Date).AddDays(-90).ToString('yyyy-MM-dd')) will be empty."
}

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

# ── Cmdlet selection (V2 preferred, V1 fallback) ──────────────────────────────
$useV2 = [bool](Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue)
if (-not $useV2 -and -not (Get-Command Get-MessageTrace -ErrorAction SilentlyContinue)) {
    throw "Neither Get-MessageTraceV2 nor Get-MessageTrace is available. Update the ExchangeOnlineManagement module."
}
$traceCmd  = if ($useV2) { 'Get-MessageTraceV2' } else { 'Get-MessageTrace' }
$detailCmd = if (Get-Command Get-MessageTraceDetailV2 -ErrorAction SilentlyContinue) { 'Get-MessageTraceDetailV2' }
             elseif (Get-Command Get-MessageTraceDetail -ErrorAction SilentlyContinue) { 'Get-MessageTraceDetail' }
             else { $null }

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Message Trace Report" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Window    : $($StartDate.ToString('yyyy-MM-dd HH:mm')) → $($EndDate.ToString('yyyy-MM-dd HH:mm')) (local)" -ForegroundColor DarkGray
Write-Host "  Cmdlet    : $traceCmd" -ForegroundColor DarkGray
if ($Mailbox)   { Write-Host "  Mailbox   : $Mailbox (sent + received)" -ForegroundColor DarkGray }
if ($Sender)    { Write-Host "  Sender    : $Sender" -ForegroundColor DarkGray }
if ($Recipient) { Write-Host "  Recipient : $Recipient" -ForegroundColor DarkGray }
if ($Subject)   { Write-Host "  Subject   : $Subject" -ForegroundColor DarkGray }
if ($MessageId) { Write-Host "  MessageId : $MessageId" -ForegroundColor DarkGray }
if ($Status)    { Write-Host "  Status    : $Status" -ForegroundColor DarkGray }
Write-Host ""

# ── Helpers ───────────────────────────────────────────────────────────────────
function ConvertTo-Utc([datetime] $dt) {
    if ($dt.Kind -eq [DateTimeKind]::Utc) { $dt } else { $dt.ToUniversalTime() }
}

function ConvertFrom-UtcStamp($value) {
    # Exchange returns UTC; Kind is not always tagged, so force it before converting.
    if (-not $value) { return $null }
    $dt = [datetime]$value
    [datetime]::SpecifyKind($dt, [DateTimeKind]::Utc).ToLocalTime()
}

function Get-NormalizedSubject([string] $s) {
    if (-not $s) { return '' }
    $n = $s.Trim()
    # Strip repeated reply/forward prefixes in EN / NL / FR / DE.
    while ($n -match '^(re|fw|fwd|antw|doorst|tr|aw|wg)\s*(\[\d+\])?\s*:\s*') {
        $n = $n -replace '^(re|fw|fwd|antw|doorst|tr|aw|wg)\s*(\[\d+\])?\s*:\s*', ''
        $n = $n.Trim()
    }
    $n.ToLowerInvariant()
}

function Get-EmailsFromText([string] $text) {
    if (-not $text) { return @() }
    ([regex]::Matches($text, "[\w\.\-\+']+@[\w\.\-]+\.\w{2,}") | ForEach-Object { $_.Value.ToLowerInvariant() }) | Sort-Object -Unique
}

# ── Forwarding configuration ──────────────────────────────────────────────────
# Shared by the pre-trace seed lookup (which turns discovered forward targets
# into extra trace filters) and the post-trace sweep over every address seen.
$forwardConfig    = [System.Collections.Generic.List[PSObject]]::new()
$checkedMailboxes = [System.Collections.Generic.HashSet[string]]::new()
$acceptedDomains  = if ($SkipForwardingConfig) { @() } else { @((Get-AcceptedDomain).DomainName) }

function Test-InternalAddress([string] $addr) {
    if (-not $addr) { return $false }
    $acceptedDomains -contains ($addr.ToLowerInvariant() -split '@')[-1]
}

function Read-ForwardingConfig {
    <#
        Records the mailbox forwarding setting and any forwarding inbox rules for
        one address, and returns the target addresses it found. Each address is
        inspected once; repeat calls return an empty list.
    #>
    param([string] $Address)

    if ($SkipForwardingConfig -or -not $Address) { return @() }
    $addr = $Address.ToLowerInvariant()
    if (-not (Test-InternalAddress $addr)) { return @() }
    if (-not $checkedMailboxes.Add($addr)) { return @() }

    try {
        $mbx = Get-EXOMailbox -Identity $addr -Properties ForwardingSMTPAddress, ForwardingAddress, DeliverToMailboxAndForward -ErrorAction Stop
    } catch {
        return @()   # not a mailbox (distribution group, contact, guest, …)
    }

    $targets = [System.Collections.Generic.List[string]]::new()

    if ($mbx.ForwardingSMTPAddress -or $mbx.ForwardingAddress) {
        $to = @(
            @($mbx.ForwardingSMTPAddress, $mbx.ForwardingAddress) |
                Where-Object { $_ } |
                ForEach-Object { Get-EmailsFromText ($_ -split 'smtp:')[-1] }
        ) | Sort-Object -Unique

        foreach ($t in $to) { if ($targets -notcontains $t) { $targets.Add($t) } }

        $forwardConfig.Add([PSCustomObject]@{
            Mailbox                    = $mbx.PrimarySmtpAddress
            Source                     = 'Mailbox setting'
            RuleName                   = ''
            ForwardTo                  = ($to -join '; ')
            DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
            Enabled                    = $true
        })
    }

    try {
        $rules = @(Get-InboxRule -Mailbox $addr -ErrorAction Stop)
    } catch {
        Write-Verbose "Could not read inbox rules for $addr : $($_.Exception.Message)"
        return $targets
    }

    foreach ($rule in $rules) {
        $raw = @(@($rule.ForwardTo) + @($rule.RedirectTo) + @($rule.ForwardAsAttachmentTo)) | Where-Object { $_ }
        if (-not $raw) { continue }

        $to = @($raw | ForEach-Object { Get-EmailsFromText $_ }) | Sort-Object -Unique
        foreach ($t in $to) { if ($targets -notcontains $t) { $targets.Add($t) } }

        $forwardConfig.Add([PSCustomObject]@{
            Mailbox                    = $mbx.PrimarySmtpAddress
            Source                     = 'Inbox rule'
            RuleName                   = $rule.Name
            ForwardTo                  = ($to -join '; ')
            DeliverToMailboxAndForward = $null
            Enabled                    = $rule.Enabled
        })
    }

    $targets
}

# ── Trace retrieval ───────────────────────────────────────────────────────────
function Invoke-Trace {
    param(
        [hashtable] $Filter,
        [datetime]  $From,
        [datetime]  $To
    )

    $collected = [System.Collections.Generic.List[PSObject]]::new()
    $pageSize  = 5000

    # V2 caps a single call at a 10-day window; split anything longer.
    $chunkDays = if ($useV2) { 10 } else { 30 }
    $chunkFrom = $From

    while ($chunkFrom -lt $To) {
        $chunkTo = $chunkFrom.AddDays($chunkDays)
        if ($chunkTo -gt $To) { $chunkTo = $To }

        if ($useV2) {
            # V2 returns newest-first; page backwards using the last row as the
            # next EndDate + StartingRecipientAddress continuation token.
            $cursorEnd       = ConvertTo-Utc $chunkTo
            $cursorRecipient = $null
            $chunkStartUtc   = ConvertTo-Utc $chunkFrom

            do {
                $p = $Filter.Clone()
                $p['StartDate']  = $chunkStartUtc
                $p['EndDate']    = $cursorEnd
                $p['ResultSize'] = $pageSize
                if ($cursorRecipient) { $p['StartingRecipientAddress'] = $cursorRecipient }

                $batch = @(Get-MessageTraceV2 @p -ErrorAction Stop)
                foreach ($row in $batch) { $collected.Add($row) }
                Write-Host ("  Retrieved {0,6} row(s)  [{1} → {2}]" -f $collected.Count,
                    $chunkFrom.ToString('yyyy-MM-dd'), $chunkTo.ToString('yyyy-MM-dd')) -ForegroundColor DarkGray

                if ($batch.Count -lt $pageSize) { break }
                $last            = $batch[-1]
                $cursorEnd       = ConvertTo-Utc ([datetime]$last.Received)
                $cursorRecipient = $last.RecipientAddress
            } while ($true)
        } else {
            $page = 1
            do {
                $p = $Filter.Clone()
                $p['StartDate'] = ConvertTo-Utc $chunkFrom
                $p['EndDate']   = ConvertTo-Utc $chunkTo
                $p['PageSize']  = $pageSize
                $p['Page']      = $page

                $batch = @(Get-MessageTrace @p -ErrorAction Stop)
                foreach ($row in $batch) { $collected.Add($row) }
                Write-Host ("  Retrieved {0,6} row(s)  [page {1}]" -f $collected.Count, $page) -ForegroundColor DarkGray
                $page++
            } while ($batch.Count -eq $pageSize -and $page -le 1000)
        }

        $chunkFrom = $chunkTo
    }

    $collected
}

# Build one filter set per direction we need to query.
$filters = [System.Collections.Generic.List[hashtable]]::new()
if ($Mailbox) {
    $filters.Add(@{ SenderAddress    = $Mailbox })
    $filters.Add(@{ RecipientAddress = $Mailbox })
} else {
    $base = @{}
    if ($Sender)    { $base['SenderAddress']    = $Sender }
    if ($Recipient) { $base['RecipientAddress'] = $Recipient }
    $filters.Add($base)
}

# A mailbox forward or a redirect rule keeps the ORIGINAL sender on the
# forwarded copy, so the delivery to the forward target mentions neither the
# traced mailbox as sender nor as recipient — filtering on the mailbox alone
# would never return it. Resolve the configured forward targets first and trace
# those addresses as recipients too, so the actual hand-off shows up with its
# own exact timestamp.
$seedTargets = [System.Collections.Generic.List[string]]::new()
foreach ($seed in @($Mailbox, $Recipient, $Sender | Where-Object { $_ })) {
    foreach ($t in (Read-ForwardingConfig -Address $seed)) {
        if ($seedTargets -notcontains $t) { $seedTargets.Add($t) }
    }
}
foreach ($t in $seedTargets) {
    Write-Host "  Configured forward found → also tracing deliveries to $t" -ForegroundColor Yellow
    $filters.Add(@{ RecipientAddress = $t })
}
if ($seedTargets.Count -gt 0) { Write-Host "" }

foreach ($f in $filters) {
    if ($MessageId) { $f['MessageId'] = $MessageId }
    if ($Status)    { $f['Status']    = $Status }
}

$raw = [System.Collections.Generic.List[PSObject]]::new()
foreach ($f in $filters) {
    foreach ($row in (Invoke-Trace -Filter $f -From $StartDate -To $EndDate)) { $raw.Add($row) }
}

# Both directions of a -Mailbox trace overlap for internal mail — de-duplicate.
$rows = $raw |
    Sort-Object @{ Expression = { "$($_.MessageTraceId)|$($_.RecipientAddress)|$($_.Received)" } } -Unique

if ($Subject) { $rows = @($rows | Where-Object { $_.Subject -like $Subject }) }
$rows = @($rows | Sort-Object Received -Descending)

Write-Host ""
Write-Host "  $($rows.Count) message row(s) matched." -ForegroundColor Cyan
Write-Host ""

if ($rows.Count -eq 0) {
    Write-Host "  Nothing to report." -ForegroundColor Yellow
    Write-Host ""
    if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
    return
}

# ── Forward detection 1: same MessageId, multiple recipients ─────────────────
$byMessageId = @{}
foreach ($r in $rows) {
    if (-not $r.MessageId) { continue }
    $key = $r.MessageId.ToString()
    if (-not $byMessageId.ContainsKey($key)) { $byMessageId[$key] = [System.Collections.Generic.List[PSObject]]::new() }
    $byMessageId[$key].Add($r)
}

# ── Forward detection 2: hop details (Redirect / transport rule) ─────────────
$detailRows    = [System.Collections.Generic.List[PSObject]]::new()
$redirectByKey = @{}   # "MessageTraceId|Recipient" → list of target addresses

if ($IncludeDetails) {
    if (-not $detailCmd) {
        Write-Warning "-IncludeDetails requested but no Get-MessageTraceDetail(V2) cmdlet is available. Skipping hop details."
    } else {
        $targets = @($rows | Select-Object -First $MaxDetailLookups)
        if ($rows.Count -gt $MaxDetailLookups) {
            Write-Host "  NOTE: hop details limited to the $MaxDetailLookups most recent of $($rows.Count) rows (-MaxDetailLookups)." -ForegroundColor Yellow
        }
        Write-Host "  Retrieving hop details for $($targets.Count) message(s)..." -ForegroundColor DarkGray

        $i = 0
        foreach ($r in $targets) {
            $i++
            if ($i % 10 -eq 0) { Write-Host "    $i / $($targets.Count)..." -ForegroundColor DarkGray }

            $dp = @{
                MessageTraceId   = $r.MessageTraceId
                RecipientAddress = $r.RecipientAddress
                StartDate        = ConvertTo-Utc $StartDate
                EndDate          = ConvertTo-Utc $EndDate
            }
            try {
                $hops = @(& $detailCmd @dp -ErrorAction Stop)
            } catch {
                Write-Verbose "Detail lookup failed for $($r.MessageTraceId) / $($r.RecipientAddress): $($_.Exception.Message)"
                continue
            }

            foreach ($h in $hops) {
                $detailRows.Add([PSCustomObject]@{
                    TimeLocal      = ConvertFrom-UtcStamp $h.Date
                    TimeUtc        = $h.Date
                    Sender         = $r.SenderAddress
                    Recipient      = $r.RecipientAddress
                    Event          = $h.Event
                    Action         = $h.Action
                    Detail         = $h.Detail
                    Subject        = $r.Subject
                    MessageId      = $r.MessageId
                    MessageTraceId = $r.MessageTraceId
                })

                # Redirect / forward / transport-rule hops name the target address
                # inside the free-text Detail field.
                if ("$($h.Event) $($h.Action)" -match 'redirect|forward|transport rule|expand') {
                    $found = @(Get-EmailsFromText $h.Detail) |
                        Where-Object { $_ -ne $r.RecipientAddress.ToString().ToLowerInvariant() -and
                                       $_ -ne $r.SenderAddress.ToString().ToLowerInvariant() }
                    if ($found) {
                        $key = "$($r.MessageTraceId)|$($r.RecipientAddress)"
                        if (-not $redirectByKey.ContainsKey($key)) { $redirectByKey[$key] = [System.Collections.Generic.List[string]]::new() }
                        foreach ($f in $found) { if ($redirectByKey[$key] -notcontains $f) { $redirectByKey[$key].Add($f) } }
                    }
                }
            }
        }
        Write-Host ""
    }
}

# ── Forward detection 3: client-side forward (new MessageId, same subject) ───
# Index every row by "sender + normalized subject" so a message received by X
# can be matched to a later message sent BY X carrying the same subject.
$sentIndex = @{}
foreach ($r in $rows) {
    $key = "$($r.SenderAddress.ToString().ToLowerInvariant())|$(Get-NormalizedSubject $r.Subject)"
    if (-not $sentIndex.ContainsKey($key)) { $sentIndex[$key] = [System.Collections.Generic.List[PSObject]]::new() }
    $sentIndex[$key].Add($r)
}

function Get-ClientForwards($row) {
    $key = "$($row.RecipientAddress.ToString().ToLowerInvariant())|$(Get-NormalizedSubject $row.Subject)"
    if (-not $sentIndex.ContainsKey($key)) { return @() }
    $origin = [datetime]$row.Received
    $origSender = $row.SenderAddress.ToString().ToLowerInvariant()
    $sentIndex[$key] |
        Where-Object {
            [datetime]$_.Received -ge $origin -and
            $_.MessageId -ne $row.MessageId -and
            # Mail back to the original sender is a reply, not a forward —
            # normalizing the subject strips "Re:" as well as "Fw:".
            $_.RecipientAddress.ToString().ToLowerInvariant() -ne $origSender
        } |
        ForEach-Object {
            [PSCustomObject]@{
                Address = $_.RecipientAddress.ToString().ToLowerInvariant()
                Time    = ConvertFrom-UtcStamp $_.Received
            }
        }
}

# ── Build the report ──────────────────────────────────────────────────────────
$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($r in $rows) {
    $recipient = $r.RecipientAddress.ToString().ToLowerInvariant()
    $forwards  = [System.Collections.Generic.List[string]]::new()
    $methods   = [System.Collections.Generic.List[string]]::new()

    # 1. Same MessageId delivered to other recipients (SMTP forward / redirect / DL expansion)
    if ($r.MessageId -and $byMessageId.ContainsKey($r.MessageId.ToString())) {
        $siblings = @($byMessageId[$r.MessageId.ToString()] |
            Where-Object { $_.RecipientAddress.ToString().ToLowerInvariant() -ne $recipient })
        foreach ($s in $siblings) {
            $addr = $s.RecipientAddress.ToString().ToLowerInvariant()
            if ($forwards -notcontains $addr) { $forwards.Add($addr) }
        }
        if ($siblings.Count -gt 0 -and $methods -notcontains 'SameMessageId') { $methods.Add('SameMessageId') }
    }

    # 2. Redirect / transport-rule hop detail
    $key = "$($r.MessageTraceId)|$($r.RecipientAddress)"
    if ($redirectByKey.ContainsKey($key)) {
        foreach ($addr in $redirectByKey[$key]) {
            if ($forwards -notcontains $addr) { $forwards.Add($addr) }
        }
        if ($methods -notcontains 'RedirectHop') { $methods.Add('RedirectHop') }
    }

    # 3. Client-side forward — heuristic, subject-matched
    $client = @(Get-ClientForwards $r)
    foreach ($c in $client) {
        if ($forwards -notcontains $c.Address) { $forwards.Add($c.Address) }
    }
    if ($client.Count -gt 0 -and $methods -notcontains 'ClientForward(subject match)') { $methods.Add('ClientForward(subject match)') }

    $localTime = ConvertFrom-UtcStamp $r.Received

    $results.Add([PSCustomObject]@{
        DateLocal        = $localTime.ToString('yyyy-MM-dd')
        TimeLocal        = $localTime.ToString('HH:mm:ss')
        ReceivedLocal    = $localTime.ToString('yyyy-MM-dd HH:mm:ss')
        ReceivedUtc      = ([datetime]$r.Received).ToString('yyyy-MM-dd HH:mm:ss')
        Sender           = $r.SenderAddress
        Recipient        = $r.RecipientAddress
        Subject          = $r.Subject
        Status           = $r.Status
        ForwardedTo      = ($forwards -join '; ')
        ForwardDetection = ($methods -join '; ')
        SizeKB           = if ($r.Size) { [math]::Round($r.Size / 1KB, 1) } else { $null }
        FromIP           = $r.FromIP
        ToIP             = $r.ToIP
        MessageId        = $r.MessageId
        MessageTraceId   = $r.MessageTraceId
    })
}

# ── Configured forwarding for the internal mailboxes involved ────────────────
$forwardConfig = [System.Collections.Generic.List[PSObject]]::new()

if (-not $SkipForwardingConfig) {
    $acceptedDomains = @((Get-AcceptedDomain).DomainName)

    # @() around each side — with a single result these are scalars and a bare
    # + would concatenate the two strings instead of building a list.
    $addresses = @(
        @($results.Sender) + @($results.Recipient) |
            Where-Object { $_ } |
            ForEach-Object { $_.ToString().ToLowerInvariant() } |
            Sort-Object -Unique |
            Where-Object { $acceptedDomains -contains ($_ -split '@')[-1] }
    )

    if ($addresses.Count -gt 0) {
        Write-Host "  Checking forwarding configuration for $($addresses.Count) internal address(es)..." -ForegroundColor DarkGray

        foreach ($addr in $addresses) {
            try {
                $mbx = Get-EXOMailbox -Identity $addr -Properties ForwardingSMTPAddress, ForwardingAddress, DeliverToMailboxAndForward -ErrorAction Stop
            } catch {
                continue   # not a mailbox (distribution group, external contact, guest, …)
            }

            if ($mbx.ForwardingSMTPAddress -or $mbx.ForwardingAddress) {
                $forwardConfig.Add([PSCustomObject]@{
                    Mailbox                    = $mbx.PrimarySmtpAddress
                    Source                     = 'Mailbox setting'
                    RuleName                   = ''
                    ForwardTo                  = (@($mbx.ForwardingSMTPAddress, $mbx.ForwardingAddress) |
                                                    Where-Object { $_ } |
                                                    ForEach-Object { ($_ -split 'smtp:')[-1].Trim() }) -join '; '
                    DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
                    Enabled                    = $true
                })
            }

            try {
                $rules = @(Get-InboxRule -Mailbox $addr -ErrorAction Stop)
            } catch {
                Write-Verbose "Could not read inbox rules for $addr : $($_.Exception.Message)"
                continue
            }

            foreach ($rule in $rules) {
                $targets = @(@($rule.ForwardTo) + @($rule.RedirectTo) + @($rule.ForwardAsAttachmentTo)) | Where-Object { $_ }
                if (-not $targets) { continue }

                $forwardConfig.Add([PSCustomObject]@{
                    Mailbox                    = $mbx.PrimarySmtpAddress
                    Source                     = 'Inbox rule'
                    RuleName                   = $rule.Name
                    ForwardTo                  = (@($targets | ForEach-Object { Get-EmailsFromText $_ }) | Sort-Object -Unique) -join '; '
                    DeliverToMailboxAndForward = $null
                    Enabled                    = $rule.Enabled
                })
            }
        }
        Write-Host ""
    }
}

# ── Display ───────────────────────────────────────────────────────────────────
$fmt = "  {0,-19} {1,-32} {2,-32} {3,-10} {4}"
Write-Host ($fmt -f 'Date/time (local)', 'Sender', 'Recipient', 'Status', 'Subject') -ForegroundColor White
Write-Host ($fmt -f '-----------------', '------', '---------', '------', '-------') -ForegroundColor White

foreach ($r in $results) {
    $color = switch -Regex ("$($r.Status)") {
        'Delivered|Expanded'          { 'Green';  break }
        'Failed|Quarantined|Spam'     { 'Red';    break }
        default                       { 'Yellow' }
    }
    Write-Host ($fmt -f $r.ReceivedLocal,
                        ("$($r.Sender)"    -replace '^(.{31}).+$', '$1…'),
                        ("$($r.Recipient)" -replace '^(.{31}).+$', '$1…'),
                        $r.Status,
                        ("$($r.Subject)"   -replace '^(.{50}).+$', '$1…')) -ForegroundColor $color

    if ($r.ForwardedTo) {
        Write-Host ("      └─ forwarded to: {0}   [{1}]" -f $r.ForwardedTo, $r.ForwardDetection) -ForegroundColor Magenta
    }
}

if ($forwardConfig.Count -gt 0) {
    Write-Host ""
    Write-Host "  ── Configured forwarding on involved mailboxes ──" -ForegroundColor Cyan
    foreach ($f in $forwardConfig) {
        $tag = if ($f.Source -eq 'Inbox rule') { "rule '$($f.RuleName)'" } else { 'mailbox setting' }
        $state = if ($f.Enabled -eq $false) { ' (disabled)' } else { '' }
        Write-Host ("  {0,-40} → {1}   [{2}{3}]" -f $f.Mailbox, $f.ForwardTo, $tag, $state) -ForegroundColor Yellow
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
if (-not $OutputPath) {
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path $outputDir "MessageTrace_$ts.csv"
}

$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "  Report saved: $OutputPath" -ForegroundColor Green

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($OutputPath)
$baseDir  = Split-Path -Parent $OutputPath
if (-not $baseDir) { $baseDir = $outputDir }

if ($detailRows.Count -gt 0) {
    $detailPath = Join-Path $baseDir "$baseName`_Details.csv"
    $detailRows | Export-Csv -Path $detailPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Hop details : $detailPath" -ForegroundColor Green
}

if ($forwardConfig.Count -gt 0) {
    $configPath = Join-Path $baseDir "$baseName`_ForwardingConfig.csv"
    $forwardConfig | Export-Csv -Path $configPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Forwarding  : $configPath" -ForegroundColor Green
}

$forwardedCount = @($results | Where-Object { $_.ForwardedTo }).Count
Write-Host ""
Write-Host "  $($results.Count) message(s) traced, $forwardedCount with a detected forward target." -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
