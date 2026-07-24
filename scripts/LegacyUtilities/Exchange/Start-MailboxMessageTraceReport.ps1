#Requires -Version 5.1
<#
.SYNOPSIS
    Request a per-mailbox historical message trace report, emailed as a compressed
    export.

.DESCRIPTION
    Submits one Start-HistoricalSearch request per mailbox (message trace, sender
    filter) for a given date range, delivered to a notification address as a
    compressed export — the standard way to pull message trace data older than
    the 10-day window covered by Get-MessageTrace. Modernized replacement for an
    old script that used the retired MSOnline module (Get-MsolUser) to build the
    mailbox list; this one uses Get-EXOMailbox instead.

    Connects to Exchange Online automatically if no session is active; reuses an
    existing session if already connected. Defaults to a safe preview — pass
    -Apply to actually submit the search requests.

.PARAMETER NotifyAddress
    Email address the completed report(s) will be sent to.

.PARAMETER UserList
    Explicit array of UPNs to report on. If omitted (and -AllMailboxes is not
    used), an error is raised — you must pick a scope.

.PARAMETER AllMailboxes
    Report on every licensed user mailbox in the tenant instead of an explicit
    list.

.PARAMETER StartDate
    Start of the reporting window. Default: 30 days ago.

.PARAMETER EndDate
    End of the reporting window. Default: now.

.PARAMETER ReportTitlePrefix
    Prefix for each search's ReportTitle (search name shows as
    "<prefix>-<UPN>-MonthlyReport"). Default: "MailboxTrace".

.PARAMETER Apply
    Actually submit the search requests. Without this switch, the script only
    lists what would be submitted.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    # Preview for two mailboxes
    .\Start-MailboxMessageTraceReport.ps1 -NotifyAddress "admin@contoso.com" -UserList "user1@contoso.com","user2@contoso.com"

.EXAMPLE
    .\Start-MailboxMessageTraceReport.ps1 -NotifyAddress "admin@contoso.com" -UserList "user1@contoso.com" -Apply

.EXAMPLE
    # All licensed mailboxes, custom date range
    .\Start-MailboxMessageTraceReport.ps1 -NotifyAddress "admin@contoso.com" -AllMailboxes -StartDate (Get-Date "2025-01-01") -EndDate (Get-Date "2025-02-01") -Apply

.NOTES
    Each Start-HistoricalSearch request is processed asynchronously by Microsoft
    and the report arrives by email — this script only submits the requests, it
    does not wait for or download results.

    Required module: ExchangeOnlineManagement
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'List')]
param(
    [Parameter(Mandatory)]
    [string] $NotifyAddress,

    [Parameter(ParameterSetName = 'List')]
    [string[]] $UserList,

    [Parameter(ParameterSetName = 'All', Mandatory)]
    [switch] $AllMailboxes,

    [datetime] $StartDate = (Get-Date).AddDays(-30),
    [datetime] $EndDate = (Get-Date),
    [string] $ReportTitlePrefix = 'MailboxTrace',
    [switch] $Apply,
    [string] $TenantId
)

if ($PSCmdlet.ParameterSetName -eq 'List' -and -not $UserList) {
    Write-Error "Specify -UserList or -AllMailboxes."
    exit 1
}

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
Write-Host "   Start-MailboxMessageTraceReport" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Notify address : $NotifyAddress"
Write-Host "  Date range     : $($StartDate.ToString('yyyy-MM-dd')) -> $($EndDate.ToString('yyyy-MM-dd'))"
Write-Host ("  Mode           : {0}" -f $(if ($Apply) { 'Apply' } else { 'Preview only' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

# ── Resolve mailboxes ──────────────────────────────────────────────────────────
if ($AllMailboxes) {
    Write-Host "  Retrieving licensed mailboxes..." -ForegroundColor DarkGray
    $users = (Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox).UserPrincipalName
} else {
    $users = $UserList
}

Write-Host "  Mailbox(es): $($users.Count)" -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($user in $users) {
    $reportTitle = "$ReportTitlePrefix-$user-MonthlyReport"

    if (-not $Apply) {
        Write-Host "  [PREVIEW] $reportTitle (sender: $user)" -ForegroundColor DarkGray
        $results.Add([PSCustomObject]@{ Mailbox = $user; ReportTitle = $reportTitle; Status = 'Preview' })
        continue
    }

    if (-not $PSCmdlet.ShouldProcess($user, "Submit historical message trace search")) { continue }

    try {
        Start-HistoricalSearch -ReportTitle $reportTitle -StartDate $StartDate -EndDate $EndDate `
            -SenderAddress $user -ReportType MessageTrace -NotifyAddress $NotifyAddress -CompressFile $true -ErrorAction Stop | Out-Null
        Write-Host "  [OK]   Submitted: $reportTitle" -ForegroundColor Green
        $results.Add([PSCustomObject]@{ Mailbox = $user; ReportTitle = $reportTitle; Status = 'Submitted' })
    } catch {
        Write-Host "  [WARN] $user : $($_.Exception.Message)" -ForegroundColor Yellow
        $results.Add([PSCustomObject]@{ Mailbox = $user; ReportTitle = $reportTitle; Status = "Error: $($_.Exception.Message)" })
    }
}

Write-Host ""
Write-Host "  Processed $($results.Count) mailbox(es)." -ForegroundColor Cyan
if (-not $Apply) { Write-Host "  Re-run with -Apply to submit these searches." -ForegroundColor Yellow }
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
