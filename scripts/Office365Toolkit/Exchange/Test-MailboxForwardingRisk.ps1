#Requires -Version 7.0
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

    Sign-in: delegated (you sign in as the admin) by default, through
    scripts\Startup\Connect-M365.ps1 — device code and the GDAP customer come
    from load.config.ps1; under GDAP the customer is reached with
    -DelegatedOrganization (earlier versions passed -Organization, which only
    applies to app-only sign-in, so they landed in the partner's own tenant).
    App-only with -ClientId + -CertificateThumbprint, or -AppOnly
    (graph.appid.json). An existing Exchange session for the tenant is reused
    and left connected.

.PARAMETER Mailbox
    UPN of a single mailbox to check. If omitted, all mailboxes are checked.

.PARAMETER IncludeDisabledRules
    Also report disabled rules that match the risky patterns (they're not
    currently active, but are worth reviewing — e.g. staged for later, or
    disabled instead of deleted).

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\ (Windows) or ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Defaults to the GDAP customer (load.config.ps1),
    else the tenant you sign in to. App-only needs the domain form
    (contoso.onmicrosoft.com).

.PARAMETER ClientId
    App registration for app-only Exchange Online sign-in (with
    -CertificateThumbprint).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

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

    Exchange Online only: Graph can read inbox rules (messageRules) of other
    users only with an app-only Mail permission, and has no API for Sweep
    rules, so the delegated default stays on Get-InboxRule / Get-SweepRule.

    Required module: ExchangeOnlineManagement
#>
[CmdletBinding()]
param(
    [string] $Mailbox,
    [switch] $IncludeDisabledRules,
    [string] $OutputPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
$exo = Connect-M365Exchange -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

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
    $Recipients = @($Recipients | Where-Object { $_ })
    if (-not $Recipients) { return 'N/A' }   # move/copy + delete rule, no recipient
    # "Name" [EX:/o=...] is a recipient inside this Exchange organization; it has no
    # domain to split, and splitting it on '@' used to flag it as External.
    $hasExchangeRecipient = [bool]($Recipients | Where-Object { $_ -match '\[EX:' })
    $addresses = $Recipients | Where-Object { $_ -notmatch '\[EX:' } | ForEach-Object {
        if ($_ -match '\[SMTP:([^\]]+)\]') { $Matches[1] } else { $_ }
    }
    $domains = $addresses | Where-Object { $_ -like '*@*' } | ForEach-Object { ($_ -split '@')[-1].Trim().TrimEnd(']') } | Where-Object { $_ }
    if (-not $domains) { return $(if ($hasExchangeRecipient) { 'Internal' } else { 'Unknown' }) }
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
Disconnect-M365Exchange $exo
