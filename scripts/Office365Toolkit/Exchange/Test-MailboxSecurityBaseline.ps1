#Requires -Version 5.1
<#
.SYNOPSIS
    Audit mailboxes against a set of Exchange Online security/hygiene baseline
    settings.

.DESCRIPTION
    Connects to Exchange Online and checks every mailbox (or a single mailbox)
    against a set of best-practice hygiene settings:
      - Mailbox audit logging enabled, and log age limit at/above threshold
      - Deleted item retention at/above threshold
      - Litigation hold enabled
      - Archive mailbox active
      - No mailbox-level forwarding configured
      - POP3 / IMAP disabled (legacy, unauthenticated-friendly protocols)
      - Not hidden from address lists (flagged for awareness, not a fail)

    Each mailbox gets a Pass/Fail per check plus an overall Status. Results are
    exported to CSV. Read-only — this script never changes mailbox settings.

.PARAMETER Mailbox
    UPN of a single mailbox to check. If omitted, all user and shared mailboxes
    are checked.

.PARAMETER MinAuditLogAgeLimitDays
    Minimum acceptable mailbox audit log age limit, in days. Default 90.

.PARAMETER MinRetainDeletedItemsDays
    Minimum acceptable deleted item retention, in days. Default 30.

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\ (Windows) or ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Test-MailboxSecurityBaseline.ps1

.EXAMPLE
    .\Test-MailboxSecurityBaseline.ps1 -Mailbox "user@contoso.com"

.EXAMPLE
    .\Test-MailboxSecurityBaseline.ps1 -MinAuditLogAgeLimitDays 180 -MinRetainDeletedItemsDays 30

.NOTES
    Capability inspired by o365-mx-check.ps1 and o365-mx-usr-all.ps1 from the
    retired directorcia/Office365 (CIAOPS) toolkit, which printed pass/fail
    checks to the console only. This rewrite consolidates both into a single
    parameterized report with CSV export, consistent with this repo's other
    Exchange audit scripts.

    Mailbox-level external forwarding is flagged here as a simple
    present/absent check; for a domain-aware external-vs-internal breakdown,
    see Get-ExternalForwards.ps1. For inbox-rule and Sweep-rule based
    forwarding (a common BEC indicator not visible on the mailbox object
    itself), see Test-MailboxForwardingRisk.ps1 in this folder.

    Required module: ExchangeOnlineManagement
#>
[CmdletBinding()]
param(
    [string] $Mailbox,
    [int]    $MinAuditLogAgeLimitDays = 90,
    [int]    $MinRetainDeletedItemsDays = 30,
    [string] $OutputPath,
    [string] $TenantId
)

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
Write-Host "   Mailbox Security Baseline Audit" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# ── Get mailboxes ─────────────────────────────────────────────────────────────
if ($Mailbox) {
    $mailboxes = @(Get-Mailbox -Identity $Mailbox -ErrorAction Stop)
} else {
    Write-Host "  Retrieving mailboxes..." -ForegroundColor DarkGray
    $mailboxes = @(Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox)
}

Write-Host "  Auditing $($mailboxes.Count) mailbox(es)..." -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($mbx in $mailboxes) {
    $casInfo = $null
    try { $casInfo = Get-CASMailbox -Identity $mbx.UserPrincipalName -ErrorAction Stop } catch {}

    $auditLogDays = $null
    try { $auditLogDays = [timespan]::Parse($mbx.AuditLogAgeLimit).Days } catch {}
    $retainDays = $null
    try { $retainDays = [timespan]::Parse($mbx.RetainDeletedItemsFor).Days } catch {}

    $checkAudit    = [bool]$mbx.AuditEnabled
    $checkAuditAge = ($null -ne $auditLogDays) -and ($auditLogDays -ge $MinAuditLogAgeLimitDays)
    $checkRetain   = ($null -ne $retainDays) -and ($retainDays -ge $MinRetainDeletedItemsDays)
    $checkLitHold  = [bool]$mbx.LitigationHoldEnabled
    $checkArchive  = $mbx.ArchiveStatus -eq 'Active'
    $checkNoFwd    = [string]::IsNullOrEmpty($mbx.ForwardingAddress) -and [string]::IsNullOrEmpty($mbx.ForwardingSmtpAddress)
    $checkNoPop    = if ($casInfo) { -not $casInfo.PopEnabled } else { $null }
    $checkNoImap   = if ($casInfo) { -not $casInfo.ImapEnabled } else { $null }

    $failCount = @($checkAudit, $checkAuditAge, $checkRetain, $checkLitHold, $checkArchive, $checkNoFwd, $checkNoPop, $checkNoImap) |
        Where-Object { $_ -eq $false } | Measure-Object | Select-Object -ExpandProperty Count

    $results.Add([PSCustomObject]@{
        DisplayName           = $mbx.DisplayName
        UserPrincipalName     = $mbx.UserPrincipalName
        MailboxType           = $mbx.RecipientTypeDetails
        AuditEnabled          = $checkAudit
        AuditLogAgeLimitDays  = $auditLogDays
        AuditLogAgeLimitOk    = $checkAuditAge
        RetainDeletedItemsFor = $retainDays
        RetainDeletedItemsOk  = $checkRetain
        LitigationHoldEnabled = $checkLitHold
        ArchiveActive         = $checkArchive
        NoMailboxForwarding   = $checkNoFwd
        PopDisabled           = $checkNoPop
        ImapDisabled          = $checkNoImap
        HiddenFromAddressLists = [bool]$mbx.HiddenFromAddressListsEnabled
        FailedChecks          = $failCount
        Status                = if ($failCount -eq 0) { 'Pass' } else { 'Review' }
    })
}

# ── Output ────────────────────────────────────────────────────────────────────
$sorted = $results | Sort-Object FailedChecks -Descending
$sorted | Format-Table DisplayName, UserPrincipalName, Status, FailedChecks -AutoSize

if (-not $OutputPath) {
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path $outputDir "MailboxSecurityBaseline_$ts.csv"
}
$sorted | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

$reviewCount = @($results | Where-Object { $_.Status -eq 'Review' }).Count
Write-Host ""
Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
Write-Host ""
Write-Host ("  {0} mailbox(es) audited — {1} need review." -f $results.Count, $reviewCount) -ForegroundColor $(if ($reviewCount -eq 0) { 'Green' } else { 'Yellow' })
Write-Host ""

# ── Disconnect if we connected ──────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
