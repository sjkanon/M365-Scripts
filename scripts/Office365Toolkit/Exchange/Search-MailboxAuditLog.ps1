#Requires -Version 7.0
<#
.SYNOPSIS
    Search the Microsoft 365 Unified Audit Log for sign-in and mailbox access
    events.

.DESCRIPTION
    Searches the Unified Audit Log for a given record type / operation / date
    range and returns everything in range. Defaults to the last 2 days,
    covering both interactive sign-ins (successful and failed) and mailbox
    logins — the two most common "who accessed what, from where" questions
    during an incident.

    By default the search runs through the Microsoft Graph Audit Log Query API
    (/security/auditLog/queries): the script creates a query, polls it until
    it has succeeded, then pages through its records. The query runs
    asynchronously on the service side and can take several minutes.
    -UseExchange runs the old Search-UnifiedAuditLog path in Exchange Online
    instead (synchronous, at most 50,000 records per record type).

    Requires the Unified Audit Log to already be enabled for the tenant, and
    that the events being searched for occurred after it was enabled.

    Results are displayed (successful in green, failed in red) and exported
    to CSV.

    Sign-in: delegated (you sign in as the admin) by default, through
    scripts\Startup\Connect-M365.ps1 — device code and the GDAP customer come
    from load.config.ps1. App-only with -ClientId + -CertificateThumbprint, or
    -AppOnly (graph.appid.json). An existing session that fits is reused and
    left connected.

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
    and mailbox item access respectively). Use the audit log names
    (PascalCase); for Graph they are converted to the API's camelCase form.
    See Microsoft's audit log schema reference for the full list.

.PARAMETER Operations
    Operation name(s) to filter to. Default: 'UserLoggedIn', 'UserLoginFailed',
    'MailboxLogin'.

.PARAMETER UserIds
    Optional. Restrict the search to specific user(s) (UPN).

.PARAMETER UseExchange
    Search with Search-UnifiedAuditLog in Exchange Online instead of the Graph
    Audit Log Query API.

.PARAMETER TimeoutMinutes
    How long to wait for the Graph audit log query to finish. Default 60. The
    query keeps running in the service after a timeout; the script prints its
    id.

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\ (Windows) or ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Defaults to the GDAP customer (load.config.ps1),
    else the tenant you sign in to. App-only Exchange needs the domain form.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.EXAMPLE
    # Last 2 days, sign-ins + mailbox logins
    .\Search-MailboxAuditLog.ps1

.EXAMPLE
    # Last 30 days, failed sign-ins only, one user
    .\Search-MailboxAuditLog.ps1 -Days 30 -RecordType AzureActiveDirectoryStsLogon -Operations UserLoginFailed -UserIds "user@contoso.com"

.EXAMPLE
    # Explicit window
    .\Search-MailboxAuditLog.ps1 -StartDate (Get-Date "2026-07-01") -EndDate (Get-Date "2026-07-15")

.EXAMPLE
    # Same search through Exchange Online (Search-UnifiedAuditLog)
    .\Search-MailboxAuditLog.ps1 -UseExchange

.NOTES
    Capability inspired by o365-login-audit.ps1 and o365-mblogin-audit.ps1
    from the retired directorcia/Office365 (CIAOPS) toolkit, which were two
    near-identical single-purpose scripts hardcoded to specific record
    types/operations. This rewrite merges them into one generic, fully
    parameterized Unified Audit Log search wrapper.

    Search-UnifiedAuditLog takes one record type per call; earlier versions
    passed the two default record types in one call. The -UseExchange path
    now runs one paged search per record type. The Graph query takes all
    record types at once.

    The Unified Audit Log is not immediate — allow up to 30-60 minutes for
    recent activity to appear. It generally records interactive sign-ins,
    not silent token refreshes, so counts may differ from Entra ID sign-in
    logs.

    Graph: delegated scope AuditLogsQuery.Read.All (app-only: the same as an
    application permission). The signed-in admin also needs a Purview role
    that may search the audit log (e.g. Audit Logs or View-Only Audit Logs).
    Required modules: Microsoft.Graph.Authentication; ExchangeOnlineManagement
    for -UseExchange.
#>
[CmdletBinding()]
param(
    [int]      $Days = 2,
    [datetime] $StartDate,
    [datetime] $EndDate = (Get-Date),
    [string[]] $RecordType = @('AzureActiveDirectoryStsLogon', 'ExchangeItem'),
    [string[]] $Operations = @('UserLoggedIn', 'UserLoginFailed', 'MailboxLogin'),
    [string[]] $UserIds,
    [switch]   $UseExchange,
    [ValidateRange(1, 1440)]
    [int]      $TimeoutMinutes = 60,
    [string]   $OutputPath,
    [string]   $TenantId,
    [string]   $ClientId,
    [string]   $CertificateThumbprint,
    [switch]   $AppOnly
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

if (-not $StartDate) { $StartDate = $EndDate.AddDays(-$Days) }
if ($StartDate -ge $EndDate) { throw "-StartDate must be earlier than -EndDate." }

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
$auth = @{ TenantId = $TenantId; ClientId = $ClientId; CertificateThumbprint = $CertificateThumbprint; AppOnly = $AppOnly }
$graph = $null; $exo = $null
if ($UseExchange) {
    $exo = Connect-M365Exchange @auth
} else {
    $graph = Connect-M365Graph @auth -Scopes 'AuditLogsQuery.Read.All'
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Unified Audit Log Search" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Source     : $(if ($UseExchange) { 'Exchange Online (Search-UnifiedAuditLog)' } else { 'Microsoft Graph (Audit Log Query API)' })" -ForegroundColor DarkGray
Write-Host "  Window     : $($StartDate.ToString('yyyy-MM-dd HH:mm')) → $($EndDate.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor DarkGray
Write-Host "  RecordType : $($RecordType -join ', ')" -ForegroundColor DarkGray
Write-Host "  Operations : $($Operations -join ', ')" -ForegroundColor DarkGray
if ($UserIds) { Write-Host "  UserIds    : $($UserIds -join ', ')" -ForegroundColor DarkGray }
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

if (-not $UseExchange) {
    # ── Graph: create the query, wait for it, page the records ──────────────────
    $base = 'https://graph.microsoft.com/v1.0/security/auditLog/queries'
    $body = @{
        displayName         = "Search-MailboxAuditLog $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        filterStartDateTime = $StartDate.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        filterEndDateTime   = $EndDate.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    }
    # The API's auditLogRecordType enum is camelCase (azureActiveDirectoryStsLogon).
    if ($RecordType) { $body['recordTypeFilters'] = @($RecordType | ForEach-Object { $_.Substring(0, 1).ToLowerInvariant() + $_.Substring(1) }) }
    if ($Operations) { $body['operationFilters'] = @($Operations) }
    if ($UserIds)    { $body['userPrincipalNameFilters'] = @($UserIds) }

    try {
        $query = Invoke-MgGraphRequest -Method POST -Uri $base -Body ($body | ConvertTo-Json -Depth 5) `
            -ContentType 'application/json' -OutputType Hashtable -ErrorAction Stop
    } catch {
        Write-Host "  [ERROR] Could not create the audit log query: $($_.Exception.Message)" -ForegroundColor Red
        Disconnect-M365Graph $graph
        exit 1
    }
    $queryId = $query['id']
    Write-Host "  Query $queryId created. Waiting for it to finish (this can take several minutes)..." -ForegroundColor DarkGray

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $status = [string]$query['status']
    while ($status -notin 'succeeded', 'failed', 'cancelled') {
        if ((Get-Date) -gt $deadline) {
            Write-Host "  [ERROR] Query $queryId still '$status' after $TimeoutMinutes minute(s). It keeps running in the service;" -ForegroundColor Red
            Write-Host "          read it later with GET $base/$queryId/records" -ForegroundColor Red
            Disconnect-M365Graph $graph
            exit 1
        }
        Start-Sleep -Seconds 30
        try {
            $status = [string](Invoke-MgGraphRequest -Method GET -Uri "$base/$queryId" -OutputType Hashtable -ErrorAction Stop)['status']
        } catch {
            Write-Host "  [WARN] Status check failed, retrying: $($_.Exception.Message)" -ForegroundColor Yellow
            continue
        }
        Write-Host "  Status: $status" -ForegroundColor DarkGray
    }
    if ($status -ne 'succeeded') {
        Write-Host "  [ERROR] Audit log query $queryId ended with status '$status'." -ForegroundColor Red
        Disconnect-M365Graph $graph
        exit 1
    }

    $uri = "$base/$queryId/records"
    $page = 1
    while ($uri) {
        Write-Host "  Retrieving page $page..." -ForegroundColor DarkGray
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType Hashtable -ErrorAction Stop
        foreach ($rec in $resp['value']) {
            $data = $rec['auditData']
            $results.Add([PSCustomObject]@{
                CreationTime = $rec['createdDateTime']
                UserId       = $(if ($rec['userPrincipalName']) { $rec['userPrincipalName'] } else { $rec['userId'] })
                Operation    = $rec['operation']
                ClientIP     = $(if ($rec['clientIp']) { $rec['clientIp'] } elseif ($data) { $data['ClientIP'] ?? $data['ClientIPAddress'] })
                RecordType   = $rec['auditLogRecordType']
                ResultStatus = $(if ($data) { $data['ResultStatus'] })
            })
        }
        $uri = $resp['@odata.nextLink']
        $page++
    }
} else {
    # ── Exchange Online: Search-UnifiedAuditLog, one paged session per record type ──
    foreach ($type in $RecordType) {
        $sessionId = "SearchMailboxAuditLog-$type-$(Get-Date -Format 'yyyyMMddHHmmss')"
        $page = 1
        do {
            Write-Host "  [$type] Retrieving page $page..." -ForegroundColor DarkGray
            $searchParams = @{
                StartDate      = $StartDate
                EndDate        = $EndDate
                RecordType     = $type
                Operations     = $Operations
                SessionId      = $sessionId
                SessionCommand = 'ReturnLargeSet'
                ResultSize     = 5000
            }
            if ($UserIds) { $searchParams['UserIds'] = $UserIds }

            $batch = @(Search-UnifiedAuditLog @searchParams -ErrorAction Stop)
            foreach ($entry in $batch) {
                $data = $entry.AuditData | ConvertFrom-Json
                $results.Add([PSCustomObject]@{
                    CreationTime = $data.CreationTime
                    UserId       = $data.UserId
                    Operation    = $data.Operation
                    ClientIP     = $(if ($data.ClientIP) { $data.ClientIP } else { $data.ClientIPAddress })
                    RecordType   = $entry.RecordType
                    ResultStatus = $data.ResultStatus
                })
            }
            $page++
            # ReturnLargeSet: ResultCount on each record is the total for the session.
            $total = if ($batch.Count) { [int]$batch[0].ResultCount } else { 0 }
            $seen  = if ($batch.Count) { [int]$batch[-1].ResultIndex } else { 0 }
        } until ($batch.Count -eq 0 -or ($total -gt 0 -and $seen -ge $total))
    }
}

Write-Host "  Retrieved $($results.Count) record(s)." -ForegroundColor DarkGray
Write-Host ""

$results = @($results | Sort-Object { [datetime]$_.CreationTime } -Descending)

# ── Display ───────────────────────────────────────────────────────────────────
Write-Host ("  {0,-20} {1,-16} {2,-30} {3}" -f 'Time', 'ClientIP', 'Operation', 'UserId') -ForegroundColor White
Write-Host ("  {0,-20} {1,-16} {2,-30} {3}" -f '----', '--------', '---------', '------') -ForegroundColor White
foreach ($r in $results) {
    $color = if ($r.Operation -eq 'UserLoginFailed' -or $r.ResultStatus -in 'Failed', 'Failure') { 'Red' } else { 'Green' }
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
if ($graph) { Disconnect-M365Graph $graph }
if ($exo)   { Disconnect-M365Exchange $exo }
