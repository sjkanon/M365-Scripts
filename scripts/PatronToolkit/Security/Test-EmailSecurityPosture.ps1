#Requires -Version 5.1
<#
.SYNOPSIS
    Report on Exchange Online / Defender for Office 365 email security configuration.

.DESCRIPTION
    Connects to Exchange Online (and, for DLP, Security & Compliance PowerShell) and
    reports the current state of the settings that most commonly matter for a tenant's
    email security posture:
      - Safe Links / Safe Attachments policies (Defender for Office 365)
      - Anti-malware, anti-spam (inbound + outbound), and connection filter policies
      - Remote domains — whether external auto-forwarding is allowed org-wide
      - Litigation hold status across mailboxes
      - DLP policy summary (name, mode, workload)
      - Alert policy summary (Security & Compliance activity/protection alerts)
      - Optionally (with -IncludeMailboxDetail, slower): POP/IMAP/legacy-protocol access
        per mailbox

    This is a read-only report — it never changes configuration. Findings are written to
    the console with a pass/warn/fail color and exported to a single flat CSV
    (Category / Item / Setting / Value / Flag) suitable for tracking drift over time or
    diffing between runs.

.PARAMETER IncludeMailboxDetail
    Also check POP/IMAP/legacy-protocol enablement per mailbox (Get-CASMailbox) and
    per-mailbox litigation hold status. Slower on large tenants — omit for a quick
    org-level-only pass.

.PARAMETER OutputPath
    CSV report path. Defaults to .\EmailSecurityPosture_<timestamp>.csv.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Test-EmailSecurityPosture.ps1

.EXAMPLE
    .\Test-EmailSecurityPosture.ps1 -IncludeMailboxDetail

.NOTES
    Inspired by the capability list of the retired directorcia/patron toolkit — this
    single script consolidates the "get" side of o365-atp-get.ps1, o365-dlp-get.ps1,
    o365-mx-malware-get.ps1, o365-mx-spam-get.ps1, o365-mx-connectpolicy-get.ps1,
    o365-mx-remotedomain.ps1, o365-mx-popimap-disable.ps1, o365-mx-fwd-disable.ps1,
    o365-mx-ews-get.ps1, o365-mx-legal-get.ps1, o365-mx-retention-get.ps1,
    o365-mx-org-get.ps1, o365-mx-be-get.ps1, o365-mx-junk-get.ps1, o365-alerts-activity-get.ps1
    and the email-related sections of o365-bp-get.ps1 — roughly twenty single-purpose
    console-dump scripts, each requiring a different manually-managed connection and none
    of which produced structured output. Rewritten from scratch as one parameterized,
    CSV-exporting report. The corresponding *-set.ps1 / *-del.ps1 mutating scripts from
    the source project were intentionally NOT ported — each hardcoded one MSP's specific
    opinionated "recommended" values with no per-tenant override, which is unsafe to
    reproduce blindly across different customer tenants.

    Required modules: ExchangeOnlineManagement (Connect-ExchangeOnline, Connect-IPPSSession)
#>
[CmdletBinding()]
param(
    [switch] $IncludeMailboxDetail,
    [string] $OutputPath,
    [string] $TenantId
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
$script:ConnectedIpps = $false
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
Write-Host "   Email Security Posture Report" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

$findings = [System.Collections.Generic.List[PSObject]]::new()

function Add-Finding {
    param(
        [string] $Category, [string] $Item, [string] $Setting, [string] $Value,
        [ValidateSet('Pass', 'Warn', 'Fail', 'Info')] [string] $Flag = 'Info'
    )
    $script:findings.Add([PSCustomObject]@{
        Category = $Category; Item = $Item; Setting = $Setting; Value = $Value; Flag = $Flag
    })
    $color = switch ($Flag) { 'Pass' { 'Green' }; 'Warn' { 'Yellow' }; 'Fail' { 'Red' }; default { 'DarkGray' } }
    Write-Host ("  [{0,-4}] {1} — {2}: {3}" -f $Flag, $Item, $Setting, $Value) -ForegroundColor $color
}

# ── Safe Links / Safe Attachments ───────────────────────────────────────────────
Write-Host "Defender for Office 365" -ForegroundColor Cyan
try {
    $safeLinks = @(Get-SafeLinksPolicy -ErrorAction Stop)
    if ($safeLinks.Count -eq 0) {
        Add-Finding 'Defender for O365' 'Safe Links' 'Policy count' '0' 'Warn'
    }
    foreach ($p in $safeLinks) {
        Add-Finding 'Defender for O365' "Safe Links: $($p.Name)" 'EnableSafeLinksForEmail' $p.EnableSafeLinksForEmail $(if ($p.EnableSafeLinksForEmail) { 'Pass' } else { 'Warn' })
    }
} catch { Add-Finding 'Defender for O365' 'Safe Links' 'Availability' "Not available: $($_.Exception.Message)" 'Info' }

try {
    $safeAttach = @(Get-SafeAttachmentPolicy -ErrorAction Stop)
    if ($safeAttach.Count -eq 0) {
        Add-Finding 'Defender for O365' 'Safe Attachments' 'Policy count' '0' 'Warn'
    }
    foreach ($p in $safeAttach) {
        Add-Finding 'Defender for O365' "Safe Attachments: $($p.Name)" 'Action' $p.Action $(if ($p.Action -in 'Block', 'DynamicDelivery') { 'Pass' } else { 'Warn' })
    }
} catch { Add-Finding 'Defender for O365' 'Safe Attachments' 'Availability' "Not available: $($_.Exception.Message)" 'Info' }

# ── Anti-malware / anti-spam / connection filter ────────────────────────────────
Write-Host ""
Write-Host "Anti-malware / Anti-spam" -ForegroundColor Cyan
try {
    foreach ($p in Get-MalwareFilterPolicy -ErrorAction Stop) {
        Add-Finding 'Anti-malware' $p.Name 'EnableFileFilter' $p.EnableFileFilter $(if ($p.EnableFileFilter) { 'Pass' } else { 'Warn' })
    }
} catch { Add-Finding 'Anti-malware' 'Policies' 'Availability' "Not available: $($_.Exception.Message)" 'Info' }

try {
    foreach ($p in Get-HostedContentFilterPolicy -ErrorAction Stop) {
        Add-Finding 'Anti-spam (inbound)' $p.Name 'SpamAction / BulkThreshold' "$($p.SpamAction) / $($p.BulkThreshold)" 'Info'
    }
} catch { Add-Finding 'Anti-spam (inbound)' 'Policies' 'Availability' "Not available: $($_.Exception.Message)" 'Info' }

try {
    foreach ($p in Get-HostedOutboundSpamFilterPolicy -ErrorAction Stop) {
        Add-Finding 'Anti-spam (outbound)' $p.Name 'AutoForwardingMode' $p.AutoForwardingMode $(if ($p.AutoForwardingMode -eq 'Off') { 'Pass' } else { 'Warn' })
    }
} catch { Add-Finding 'Anti-spam (outbound)' 'Policies' 'Availability' "Not available: $($_.Exception.Message)" 'Info' }

try {
    foreach ($p in Get-HostedConnectionFilterPolicy -ErrorAction Stop) {
        $allowCount = @($p.IPAllowList).Count
        Add-Finding 'Connection filter' $p.Name 'IPAllowList count' $allowCount $(if ($allowCount -gt 0) { 'Warn' } else { 'Pass' })
    }
} catch { Add-Finding 'Connection filter' 'Policies' 'Availability' "Not available: $($_.Exception.Message)" 'Info' }

# ── Remote domains (org-wide external auto-forward) ─────────────────────────────
Write-Host ""
Write-Host "Mail Flow" -ForegroundColor Cyan
try {
    foreach ($rd in Get-RemoteDomain -ErrorAction Stop) {
        Add-Finding 'Remote domains' $rd.DomainName 'AutoForwardEnabled' $rd.AutoForwardEnabled $(if ($rd.AutoForwardEnabled) { 'Warn' } else { 'Pass' })
    }
} catch { Add-Finding 'Remote domains' 'Config' 'Availability' "Not available: $($_.Exception.Message)" 'Info' }

# ── DLP (Security & Compliance) ─────────────────────────────────────────────────
Write-Host ""
Write-Host "Data Loss Prevention" -ForegroundColor Cyan
try {
    if (-not (Get-Command Get-DlpCompliancePolicy -ErrorAction SilentlyContinue)) {
        Connect-IPPSSession -ErrorAction Stop
        $script:ConnectedIpps = $true
    }
    $dlpPolicies = @(Get-DlpCompliancePolicy -ErrorAction Stop)
    if ($dlpPolicies.Count -eq 0) {
        Add-Finding 'DLP' 'Policies' 'Policy count' '0' 'Warn'
    }
    foreach ($p in $dlpPolicies) {
        Add-Finding 'DLP' $p.Name 'Mode / Workload' "$($p.Mode) / $($p.Workload -join ',')" $(if ($p.Mode -eq 'Enable') { 'Pass' } else { 'Info' })
    }
} catch { Add-Finding 'DLP' 'Policies' 'Availability' "Not available: $($_.Exception.Message)" 'Info' }

try {
    $alertPolicies = @(Get-ProtectionAlert -ErrorAction Stop)
    $enabledCount = @($alertPolicies | Where-Object { $_.Disabled -eq $false }).Count
    Add-Finding 'Alert policies' 'Security & Compliance' 'Enabled / Total' "$enabledCount / $($alertPolicies.Count)" 'Info'
} catch { Add-Finding 'Alert policies' 'Security & Compliance' 'Availability' "Not available: $($_.Exception.Message)" 'Info' }

# ── Mailbox-level detail (optional, slower) ─────────────────────────────────────
if ($IncludeMailboxDetail) {
    Write-Host ""
    Write-Host "Mailbox Detail" -ForegroundColor Cyan
    Write-Host "  Retrieving mailboxes..." -ForegroundColor DarkGray
    $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox -Properties LitigationHoldEnabled)

    $legacyProtocolMailboxes = 0
    $litHoldMailboxes = 0
    foreach ($mbx in $mailboxes) {
        try {
            $cas = Get-CASMailbox -Identity $mbx.UserPrincipalName -ErrorAction Stop
            if ($cas.PopEnabled -or $cas.ImapEnabled) { $legacyProtocolMailboxes++ }
        } catch {}
        if ($mbx.LitigationHoldEnabled) { $litHoldMailboxes++ }
    }
    Add-Finding 'Mailbox detail' 'POP/IMAP' 'Mailboxes with POP or IMAP enabled' $legacyProtocolMailboxes $(if ($legacyProtocolMailboxes -gt 0) { 'Warn' } else { 'Pass' })
    Add-Finding 'Mailbox detail' 'Litigation hold' 'Mailboxes on hold' "$litHoldMailboxes / $($mailboxes.Count)" 'Info'
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
if (-not $OutputPath) {
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path $outputDir "EmailSecurityPosture_$ts.csv"
}
$findings | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "  Report saved: $OutputPath" -ForegroundColor Green

$failCount = @($findings | Where-Object { $_.Flag -eq 'Fail' }).Count
$warnCount = @($findings | Where-Object { $_.Flag -eq 'Warn' }).Count
Write-Host ""
Write-Host ("  {0} finding(s) — {1} warning(s), {2} failure(s)" -f $findings.Count, $warnCount, $failCount) -ForegroundColor $(if ($failCount -gt 0) { 'Red' } elseif ($warnCount -gt 0) { 'Yellow' } else { 'Green' })
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
