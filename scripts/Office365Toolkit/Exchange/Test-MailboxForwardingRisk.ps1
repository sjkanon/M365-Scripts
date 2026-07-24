#Requires -Version 5.1
<#
.SYNOPSIS
    Audit Outlook inbox rules and Sweep rules across mailboxes for
    forwarding/redirect/exfiltration patterns.

.DESCRIPTION
    Connects to Exchange Online and inspects every mailbox's inbox rules and
    Sweep rules for actions commonly abused after a mailbox compromise:
    ForwardTo, RedirectTo, ForwardAsAttachmentTo, CopyToFolder combined with
    DeleteMessage, or a Sweep rule that moves/deletes incoming mail. These are
    a classic business email compromise (BEC) indicator — an attacker with
    temporary access creates a rule that silently forwards or hides mail (e.g.
    invoice/payment threads) without ever needing standing access afterwards.

    Read-only — this script only reports, it never removes rules.

.PARAMETER Mailbox
    UPN of a single mailbox to check. If omitted, all mailboxes are checked.

.PARAMETER IncludeDisabledRules
    Also report disabled rules that match the risky patterns (they're not
    currently active, but are worth reviewing — e.g. staged for later, or
    disabled instead of deleted).

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\ (Windows) or ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Test-MailboxForwardingRisk.ps1

.EXAMPLE
    .\Test-MailboxForwardingRisk.ps1 -Mailbox "user@contoso.com" -IncludeDisabledRules

.NOTES
    Capability inspired by o365-exo-fwd-chk.ps1 from the retired
    directorcia/Office365 (CIAOPS) toolkit. That script also checked
    mailbox-level ForwardingSmtpAddress — this repo already covers that,
    with external-domain awareness, in Get-ExternalForwards.ps1, so this
    script focuses on the part that isn't covered elsewhere: inbox rules and
    Sweep rules, which don't show up on the mailbox object itself.

    Rule recipients (ForwardTo/RedirectTo/etc.) are flagged External when
    they resolve to an SMTP address outside the tenant's accepted domains,
    and Internal/Unknown otherwise (some rule actions target a folder or
    non-SMTP recipient and can't be classified this way).

    Required module: ExchangeOnlineManagement
#>
[CmdletBinding()]
param(
    [string] $Mailbox,
    [switch] $IncludeDisabledRules,
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
Write-Host "   Mailbox Rule Forwarding Risk Audit" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

$acceptedDomains = (Get-AcceptedDomain).DomainName
Write-Host "  Internal domains: $($acceptedDomains -join ', ')" -ForegroundColor DarkGray
Write-Host ""

# ── Get mailboxes ─────────────────────────────────────────────────────────────
if ($Mailbox) {
    $mailboxes = @(Get-EXOMailbox -Identity $Mailbox -ErrorAction Stop)
} else {
    Write-Host "  Retrieving mailboxes..." -ForegroundColor DarkGray
    $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited)
}

Write-Host "  Checking $($mailboxes.Count) mailbox(es) for risky rules..." -ForegroundColor DarkGray
Write-Host ""

function Get-RuleRecipientClassification {
    param([string[]] $Recipients, [string[]] $AcceptedDomains)
    if (-not $Recipients) { return $null }
    $addresses = $Recipients | ForEach-Object {
        if ($_ -match '\[SMTP:([^\]]+)\]') { $Matches[1] } else { $_ }
    }
    $domains = $addresses | ForEach-Object { ($_ -split '@')[-1] } | Where-Object { $_ }
    if (-not $domains) { return 'Unknown' }
    if ($domains | Where-Object { $AcceptedDomains -notcontains $_ }) { return 'External' }
    return 'Internal'
}

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($mbx in $mailboxes) {
    # ── Inbox rules ──────────────────────────────────────────────────────────
    $rules = $null
    try { $rules = Get-InboxRule -Mailbox $mbx.UserPrincipalName -ErrorAction Stop } catch {
        Write-Host "  [WARN] Could not read inbox rules for $($mbx.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Yellow
        continue
    }

    foreach ($rule in $rules) {
        if (-not $rule.Enabled -and -not $IncludeDisabledRules) { continue }

        $isRisky = $rule.ForwardTo -or $rule.RedirectTo -or $rule.ForwardAsAttachmentTo -or
                   ($rule.CopyToFolder -and $rule.DeleteMessage) -or
                   ($rule.MoveToFolder -and $rule.DeleteMessage)
        if (-not $isRisky) { continue }

        $recipients = @($rule.ForwardTo) + @($rule.RedirectTo) + @($rule.ForwardAsAttachmentTo)
        $classification = Get-RuleRecipientClassification -Recipients $recipients -AcceptedDomains $acceptedDomains

        $entry = [PSCustomObject]@{
            Mailbox        = $mbx.UserPrincipalName
            RuleType       = 'InboxRule'
            RuleName       = $rule.Name
            Enabled        = $rule.Enabled
            ForwardTo      = ($rule.ForwardTo -join '; ')
            RedirectTo     = ($rule.RedirectTo -join '; ')
            ForwardAsAttachmentTo = ($rule.ForwardAsAttachmentTo -join '; ')
            DeleteMessage  = [bool]$rule.DeleteMessage
            MoveOrCopyToFolder = $(if ($rule.MoveToFolder) { $rule.MoveToFolder } else { $rule.CopyToFolder })
            RecipientScope = $classification
        }
        $results.Add($entry)

        $color = if ($classification -eq 'External') { 'Red' } else { 'Yellow' }
        Write-Host ("  [{0}] {1} — rule '{2}' ({3})" -f $classification, $mbx.UserPrincipalName, $rule.Name, $(if ($rule.Enabled) { 'enabled' } else { 'disabled' })) -ForegroundColor $color
    }

    # ── Sweep rules ──────────────────────────────────────────────────────────
    $sweeps = $null
    try { $sweeps = Get-SweepRule -Mailbox $mbx.UserPrincipalName -ErrorAction SilentlyContinue } catch {}

    foreach ($sweep in $sweeps) {
        if (-not $sweep.Enabled -and -not $IncludeDisabledRules) { continue }

        $results.Add([PSCustomObject]@{
            Mailbox        = $mbx.UserPrincipalName
            RuleType       = 'SweepRule'
            RuleName       = $sweep.Name
            Enabled        = $sweep.Enabled
            ForwardTo      = $null
            RedirectTo     = $null
            ForwardAsAttachmentTo = $null
            DeleteMessage  = $sweep.DestinationFolder -eq 'DeletedItems'
            MoveOrCopyToFolder = $sweep.DestinationFolder
            RecipientScope = 'N/A'
        })

        Write-Host ("  [SWEEP] {0} — rule '{1}' moves mail to '{2}'" -f $mbx.UserPrincipalName, $sweep.Name, $sweep.DestinationFolder) -ForegroundColor Yellow
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
if ($results.Count -eq 0) {
    Write-Host "  No risky inbox or Sweep rules found." -ForegroundColor Green
} else {
    Write-Host "  $($results.Count) risky rule(s) found." -ForegroundColor Yellow

    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "MailboxForwardingRisk_$ts.csv"
    }
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "  Checked $($mailboxes.Count) mailbox(es)." -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ──────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
