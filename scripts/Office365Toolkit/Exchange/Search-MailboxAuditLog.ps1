#Requires -Version 5.1
<#
.SYNOPSIS
    Search the Microsoft 365 Unified Audit Log for sign-in and mailbox access
    events.

.DESCRIPTION
    Connects to Exchange Online and searches the Unified Audit Log
    (Search-UnifiedAuditLog) for a given record type / operation / date range,
    paginating through results (5000 per page) until everything in range has
    been retrieved. Defaults to the last 2 days, covering both interactive
    sign-ins (successful and failed) and mailbox logins — the two most common
    "who accessed what, from where" questions during an incident.

    Requires the Unified Audit Log to already be enabled for the tenant, and
    that the events being searched for occurred after it was enabled.

    Results are displayed (successful in green, failed in red) and exported
    to CSV.

.PARAMETER Days
    Number of days back from now to search. Ignored if -StartDate is given.
    Default 2.

.PARAMETER StartDate
    Explicit start of the search window (local time). Overrides -Days.

.PARAMETER EndDate
    Explicit end of the search window (local time). Default: now.

.PARAMETER RecordType
    Unified Audit Log record type(s) to search. Default:
    'AzureActiveDirectoryStsLogon', 'ExchangeItem' (covers interactive sign-in
    and mailbox item access respectively). See Microsoft's audit log schema
    reference for the full list of valid values.

.PARAMETER Operations
    Operation name(s) to filter to. Default: 'UserLoggedIn', 'UserLoginFailed',
    'MailboxLogin'.

.PARAMETER UserIds
    Optional. Restrict the search to specific user(s) (UPN or email).

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\ (Windows) or ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    # Last 2 days, sign-ins + mailbox logins
    .\Search-MailboxAuditLog.ps1

.EXAMPLE
    # Last 30 days, failed sign-ins only, one user
    .\Search-MailboxAuditLog.ps1 -Days 30 -RecordType AzureActiveDirectoryStsLogon -Operations UserLoginFailed -UserIds "user@contoso.com"

.EXAMPLE
    # Explicit window
    .\Search-MailboxAuditLog.ps1 -StartDate (Get-Date "2026-07-01") -EndDate (Get-Date "2026-07-15")

.NOTES
    Capability inspired by o365-login-audit.ps1 and o365-mblogin-audit.ps1
    from the retired directorcia/Office365 (CIAOPS) toolkit, which were two
    near-identical single-purpose scripts hardcoded to specific record
    types/operations. This rewrite merges them into one generic, fully
    parameterized Unified Audit Log search wrapper.

    The Unified Audit Log is not immediate — allow up to 30-60 minutes for
    recent activity to appear. It generally records interactive sign-ins,
    not silent token refreshes, so counts may differ from Entra ID sign-in
    logs.

    Required module: ExchangeOnlineManagement
#>
[CmdletBinding()]
param(
    [int]      $Days = 2,
    [datetime] $StartDate,
    [datetime] $EndDate = (Get-Date),
    [string[]] $RecordType = @('AzureActiveDirectoryStsLogon', 'ExchangeItem'),
    [string[]] $Operations = @('UserLoggedIn', 'UserLoginFailed', 'MailboxLogin'),
    [string[]] $UserIds,
    [string]   $OutputPath,
    [string]   $TenantId
)

if (-not $StartDate) { $StartDate = $EndDate.AddDays(-$Days) }
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
Write-Host "   Unified Audit Log Search" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Window     : $($StartDate.ToString('yyyy-MM-dd HH:mm')) → $($EndDate.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor DarkGray
Write-Host "  RecordType : $($RecordType -join ', ')" -ForegroundColor DarkGray
Write-Host "  Operations : $($Operations -join ', ')" -ForegroundColor DarkGray
if ($UserIds) { Write-Host "  UserIds    : $($UserIds -join ', ')" -ForegroundColor DarkGray }
Write-Host ""

# ── Search (paginated) ────────────────────────────────────────────────────────
$sessionId = "SearchMailboxAuditLog-$(Get-Date -Format 'yyyyMMddHHmmss')"
$all = [System.Collections.Generic.List[PSObject]]::new()
$page = 1

do {
    Write-Host "  Retrieving page $page..." -ForegroundColor DarkGray
    $searchParams = @{
        StartDate      = $StartDate
        EndDate        = $EndDate
        RecordType     = $RecordType
        Operations     = $Operations
        SessionId      = $sessionId
        SessionCommand = 'ReturnLargeSet'
        ResultSize     = 5000
    }
    if ($UserIds) { $searchParams['UserIds'] = $UserIds }

    $batch = Search-UnifiedAuditLog @searchParams -ErrorAction Stop
    foreach ($item in $batch) { $all.Add($item) }
    $page++
} until (-not $batch -or $batch.Count -eq 0)

Write-Host "  Retrieved $($all.Count) raw record(s). Parsing..." -ForegroundColor DarkGray
Write-Host ""

# ── Parse ─────────────────────────────────────────────────────────────────────
$results = foreach ($entry in $all) {
    $data = $entry.AuditData | ConvertFrom-Json
    [PSCustomObject]@{
        CreationTime = $data.CreationTime
        UserId       = $data.UserId
        Operation    = $data.Operation
        ClientIP     = $(if ($data.ClientIP) { $data.ClientIP } else { $data.ClientIPAddress })
        RecordType   = $entry.RecordType
        ResultStatus = $data.ResultStatus
    }
}
$results = $results | Sort-Object CreationTime -Descending

# ── Display ───────────────────────────────────────────────────────────────────
Write-Host ("  {0,-20} {1,-16} {2,-30} {3}" -f 'Time', 'ClientIP', 'Operation', 'UserId') -ForegroundColor White
Write-Host ("  {0,-20} {1,-16} {2,-30} {3}" -f '----', '--------', '---------', '------') -ForegroundColor White
foreach ($r in $results) {
    $color = if ($r.Operation -eq 'UserLoginFailed') { 'Red' } else { 'Green' }
    Write-Host ("  {0,-20} {1,-16} {2,-30} {3}" -f $r.CreationTime, $r.ClientIP, $r.Operation, $r.UserId) -ForegroundColor $color
}

# ── Output ────────────────────────────────────────────────────────────────────
if (-not $OutputPath) {
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path $outputDir "AuditLogSearch_$ts.csv"
}
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
Write-Host ""
Write-Host "  $($results.Count) event(s) found." -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ──────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
