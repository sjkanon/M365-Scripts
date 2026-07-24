#Requires -Version 5.1
<#
.SYNOPSIS
    Detect suspicious mailbox inbox rules (a common Business Email Compromise indicator).

.DESCRIPTION
    Connects to Exchange Online and inspects inbox rules for every mailbox (or one
    mailbox), flagging rules that match common post-compromise attacker patterns:
      - Forward/redirect/forward-as-attachment to an external (non-tenant) address
      - Silently deletes messages (DeleteMessage)
      - Moves messages to a rarely-checked folder (RSS Feeds, Conversation History,
        Junk Email, Archive, or a custom folder) while also matching on keywords
        commonly used in BEC/finance-fraud scenarios (invoice, payment, wire, password,
        security alert, etc.)
      - Is disabled-looking but still enabled with no visible name (blank/whitespace name)

    Default behavior is a read-only report. Pass -Apply to disable (not delete) rules
    flagged as high-confidence matches — disabling is reversible and non-destructive,
    unlike deletion. Supports -WhatIf.

.PARAMETER Mailbox
    UPN of a single mailbox. If omitted, all mailboxes are checked.

.PARAMETER Apply
    Disable rules flagged as high-confidence suspicious matches (external forward/redirect
    combined with delete, or delete-only rules). Without this switch, only reports.

.PARAMETER OutputPath
    CSV report path. Defaults to .\SuspiciousInboxRules_<timestamp>.csv.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Get-SuspiciousInboxRules.ps1

.EXAMPLE
    .\Get-SuspiciousInboxRules.ps1 -Mailbox "user@contoso.com"

.EXAMPLE
    # Preview which rules would be disabled
    .\Get-SuspiciousInboxRules.ps1 -Apply -WhatIf

.EXAMPLE
    .\Get-SuspiciousInboxRules.ps1 -Apply

.NOTES
    Inspired by the capability list of the retired directorcia/patron toolkit
    (o365-mx-inboxrules-get.ps1 / o365-mx-inboxrules-del.ps1), rewritten from scratch with
    BEC-pattern detection heuristics added — the original only dumped raw rule properties
    to CSV with no risk scoring.

    Required role: View-Only Recipients (report) / Recipient Management or Organization
    Management (to disable rules with -Apply)
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $Mailbox,
    [switch] $Apply,
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
Write-Host "   Suspicious Inbox Rule Detection" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# ── Accepted domains (to determine what counts as "external") ─────────────────
$acceptedDomains = @()
try {
    $acceptedDomains = @(Get-AcceptedDomain -ErrorAction Stop | Select-Object -ExpandProperty DomainName)
} catch {
    Write-Host "  [WARN] Could not retrieve accepted domains — external-forward detection may be less accurate." -ForegroundColor Yellow
}

$becKeywords = @('invoice', 'payment', 'wire', 'password', 'bank', 'security alert', 'urgent', 'w-2', 'w2', 'remittance')
$hiddenFolders = @('RSS Feeds', 'Conversation History', 'Junk Email', 'Deleted Items', 'Archive')

function Test-ExternalAddress {
    param([string[]] $Addresses)
    foreach ($addr in $Addresses) {
        if (-not $addr) { continue }
        $domain = ($addr -split '@')[-1] -replace '[\[\]]', ''
        if ($domain -and ($acceptedDomains -notcontains $domain)) { return $true }
    }
    return $false
}

# ── Get mailboxes ─────────────────────────────────────────────────────────────
if ($Mailbox) {
    $mailboxes = @(Get-EXOMailbox -Identity $Mailbox -ErrorAction Stop)
} else {
    Write-Host "  Retrieving mailboxes..." -ForegroundColor DarkGray
    $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox)
}

Write-Host "  Checking inbox rules for $($mailboxes.Count) mailbox(es)..." -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($mbx in $mailboxes) {
    try {
        $rules = @(Get-InboxRule -Mailbox $mbx.UserPrincipalName -IncludeHidden -ErrorAction Stop)
    } catch {
        Write-Host "  [WARN] $($mbx.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Yellow
        continue
    }

    foreach ($rule in $rules) {
        if (-not $rule.Enabled) { continue }

        $reasons = [System.Collections.Generic.List[string]]::new()
        $externalTargets = @($rule.ForwardTo) + @($rule.RedirectTo) + @($rule.ForwardAsAttachmentTo) | Where-Object { $_ }

        if ($externalTargets.Count -gt 0 -and (Test-ExternalAddress -Addresses $externalTargets)) {
            $reasons.Add('Forwards/redirects to an external address')
        }
        if ($rule.DeleteMessage) {
            $reasons.Add('Silently deletes matching messages')
        }
        $movesToHiddenFolder = $rule.MoveToFolder -and ($hiddenFolders | Where-Object { $rule.MoveToFolder -like "*$_*" })
        $matchesKeyword = $false
        foreach ($kw in $becKeywords) {
            if (($rule.SubjectContainsWords -join ' ') -match [regex]::Escape($kw) -or
                ($rule.BodyContainsWords -join ' ') -match [regex]::Escape($kw)) {
                $matchesKeyword = $true
                break
            }
        }
        if ($movesToHiddenFolder -and $matchesKeyword) {
            $reasons.Add('Moves keyword-matching messages to a rarely-checked folder')
        }

        if ($reasons.Count -eq 0) { continue }

        $highConfidence = ($reasons -match 'external|deletes').Count -gt 0

        $results.Add([PSCustomObject]@{
            Mailbox            = $mbx.UserPrincipalName
            RuleName           = $rule.Name
            Enabled            = $rule.Enabled
            ForwardTo          = ($rule.ForwardTo -join '; ')
            RedirectTo         = ($rule.RedirectTo -join '; ')
            DeleteMessage      = $rule.DeleteMessage
            MoveToFolder       = $rule.MoveToFolder
            SubjectContains    = ($rule.SubjectContainsWords -join '; ')
            Reasons            = ($reasons -join '; ')
            HighConfidence     = $highConfidence
            RuleIdentity       = $rule.Identity
        })

        $color = if ($highConfidence) { 'Red' } else { 'Yellow' }
        Write-Host "  [FLAG] $($mbx.UserPrincipalName) — '$($rule.Name)': $($reasons -join '; ')" -ForegroundColor $color
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
if ($results.Count -eq 0) {
    Write-Host "  No suspicious inbox rules found." -ForegroundColor DarkGray
} else {
    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "SuspiciousInboxRules_$ts.csv"
    }
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

# ── Remediation ───────────────────────────────────────────────────────────────
if ($Apply) {
    $toDisable = @($results | Where-Object { $_.HighConfidence })
    Write-Host ""
    Write-Host "  Disabling $($toDisable.Count) high-confidence rule(s)..." -ForegroundColor Cyan
    foreach ($item in $toDisable) {
        if ($PSCmdlet.ShouldProcess("$($item.Mailbox) — '$($item.RuleName)'", "Disable inbox rule")) {
            try {
                Disable-InboxRule -Identity $item.RuleIdentity -Mailbox $item.Mailbox -Confirm:$false -ErrorAction Stop
                Write-Host "  [OK] Disabled '$($item.RuleName)' on $($item.Mailbox)" -ForegroundColor Green
            } catch {
                Write-Host "  [WARN] Could not disable '$($item.RuleName)' on $($item.Mailbox): $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
} elseif ($results.Count -gt 0) {
    $highConfidenceCount = @($results | Where-Object { $_.HighConfidence }).Count
    Write-Host ""
    Write-Host "  $highConfidenceCount high-confidence rule(s) would be disabled. Re-run with -Apply to disable them." -ForegroundColor Yellow
}

Write-Host ""
Write-Host ("  {0} mailbox(es) checked — {1} suspicious rule(s) found" -f $mailboxes.Count, $results.Count) -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
