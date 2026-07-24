#Requires -Version 5.1
<#
.SYNOPSIS
    Create or update a baseline Exchange Online Protection anti-spam and
    anti-malware policy, scoped to the tenant's accepted domains.

.DESCRIPTION
    Connects to Exchange Online and reports the tenant's current default (and,
    if present, custom-named) hosted content filter (anti-spam) and malware
    filter policies alongside a recommended baseline configuration. With
    -Apply, creates the baseline policy + rule pair for whichever protection
    types are requested (default: both) — or updates them in place if
    -UpdateExisting is also specified and a policy with the same name already
    exists.

    Default behavior is a safe report-only preview — no changes are made
    without -Apply.

.PARAMETER Domains
    Recipient domain(s) the baseline rule(s) should apply to. Defaults to all
    of the tenant's accepted domains (Get-AcceptedDomain) if omitted.

.PARAMETER Protection
    Which policy type(s) to create/update: Spam, Malware, or Both (default).

.PARAMETER SpamPolicyName
    Name for the anti-spam content filter policy + rule. Default:
    "MSP Baseline Anti-Spam".

.PARAMETER MalwarePolicyName
    Name for the anti-malware filter policy + rule. Default:
    "MSP Baseline Anti-Malware".

.PARAMETER UpdateExisting
    If a policy with the target name already exists, update it in place
    instead of skipping it.

.PARAMETER Apply
    Actually create/update the policies and rules. Without this switch, the
    script only reports the recommended baseline vs. what currently exists.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    # Preview only
    .\New-EOPProtectionBaseline.ps1

.EXAMPLE
    # Create both baseline policies for all accepted domains
    .\New-EOPProtectionBaseline.ps1 -Apply

.EXAMPLE
    # Anti-spam only, specific domains, update if it already exists
    .\New-EOPProtectionBaseline.ps1 -Protection Spam -Domains "contoso.com" -UpdateExisting -Apply

.NOTES
    Capability inspired by o365-spam-policy.ps1 and o365-malware-policy.ps1
    from the retired directorcia/Office365 (CIAOPS) toolkit, which hardcoded
    a specific tenant's onmicrosoft.com domain. This rewrite generalizes the
    domain scope to a parameter (defaulting to the connected tenant's accepted
    domains) and merges both policies into one baseline script with a
    dry-run/-Apply gate, consistent with this repo's other mutating scripts.

    These are baseline recommendations, not a full EOP hardening pass — review
    the settings against your own tenant's needs (e.g. Standard vs. Strict
    preset security policies may already cover this) before applying broadly.

    Supports -WhatIf (SupportsShouldProcess).

    Required module: ExchangeOnlineManagement
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string[]] $Domains,

    [ValidateSet('Spam', 'Malware', 'Both')]
    [string] $Protection = 'Both',

    [string] $SpamPolicyName = 'MSP Baseline Anti-Spam',
    [string] $MalwarePolicyName = 'MSP Baseline Anti-Malware',
    [switch] $UpdateExisting,
    [switch] $Apply,
    [string] $TenantId
)

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    $null = Get-AcceptedDomain -ResultSize 1 -ErrorAction Stop
} catch {
    $connectParams = @{ ShowBanner = $false }
    if ($TenantId) { $connectParams['Organization'] = $TenantId }
    Connect-ExchangeOnline @connectParams
    $script:ConnectedHere = $true
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   EOP Protection Baseline" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host ("  Mode : {0}" -f $(if ($Apply) { 'Apply (policies will be created/updated)' } else { 'Preview only (no changes)' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

if (-not $Domains) {
    $Domains = (Get-AcceptedDomain).DomainName
    Write-Host "  -Domains not specified — defaulting to accepted domains: $($Domains -join ', ')" -ForegroundColor DarkGray
    Write-Host ""
}

# ── Baseline definitions ─────────────────────────────────────────────────────
$spamBaseline = @{
    BulkSpamAction                       = 'MoveToJmf'
    BulkThreshold                        = 7
    HighConfidenceSpamAction             = 'MoveToJmf'
    InlineSafetyTipsEnabled              = $true
    MarkAsSpamBulkMail                   = 'On'
    IncreaseScoreWithImageLinks          = 'Off'
    IncreaseScoreWithNumericIps          = 'On'
    IncreaseScoreWithRedirectToOtherPort = 'On'
    IncreaseScoreWithBizOrInfoUrls       = 'On'
    MarkAsSpamEmptyMessages              = 'On'
    MarkAsSpamJavaScriptInHtml           = 'On'
    MarkAsSpamFramesInHtml               = 'On'
    MarkAsSpamObjectTagsInHtml           = 'On'
    MarkAsSpamEmbedTagsInHtml            = 'On'
    MarkAsSpamFormTagsInHtml             = 'On'
    MarkAsSpamWebBugsInHtml              = 'On'
    MarkAsSpamSensitiveWordList          = 'On'
    MarkAsSpamSpfRecordHardFail          = 'On'
    MarkAsSpamFromAddressAuthFail        = 'On'
    MarkAsSpamNdrBackscatter             = 'On'
    PhishSpamAction                      = 'MoveToJmf'
    SpamAction                           = 'MoveToJmf'
    ZapEnabled                           = $true
}

$malwareBaseline = @{
    Action                             = 'DeleteMessage'
    EnableFileFilter                   = $true
    EnableInternalSenderNotifications  = $true
    ZapEnabled                         = $true
}

# ── Spam policy ───────────────────────────────────────────────────────────────
if ($Protection -in 'Spam', 'Both') {
    Write-Host "  --- Anti-Spam: $SpamPolicyName ---" -ForegroundColor White
    $existing = Get-HostedContentFilterPolicy -Identity $SpamPolicyName -ErrorAction SilentlyContinue

    if ($existing) {
        Write-Host "  [EXISTS] Policy already present." -ForegroundColor DarkGray
        if (-not $UpdateExisting) {
            Write-Host "  Skipping (pass -UpdateExisting to update it in place)." -ForegroundColor DarkGray
        } elseif ($Apply -and $PSCmdlet.ShouldProcess($SpamPolicyName, "Update hosted content filter policy")) {
            Set-HostedContentFilterPolicy -Identity $SpamPolicyName @spamBaseline -ErrorAction Stop
            Write-Host "  [OK] Policy updated." -ForegroundColor Green
        } elseif (-not $Apply) {
            Write-Host "  Would update this policy with the baseline settings (re-run with -Apply -UpdateExisting)." -ForegroundColor Yellow
        }
    } else {
        Write-Host "  [MISSING] Policy does not exist yet." -ForegroundColor Yellow
        if ($Apply -and $PSCmdlet.ShouldProcess($SpamPolicyName, "Create hosted content filter policy + rule")) {
            New-HostedContentFilterPolicy -Name $SpamPolicyName @spamBaseline -ErrorAction Stop | Out-Null
            New-HostedContentFilterRule -Name $SpamPolicyName -HostedContentFilterPolicy $SpamPolicyName -RecipientDomainIs $Domains -Enabled $true -ErrorAction Stop | Out-Null
            Write-Host "  [OK] Policy + rule created, scoped to: $($Domains -join ', ')" -ForegroundColor Green
        } elseif (-not $Apply) {
            Write-Host "  Would create this policy + rule scoped to: $($Domains -join ', ') (re-run with -Apply)." -ForegroundColor Yellow
        }
    }
    Write-Host ""
}

# ── Malware policy ────────────────────────────────────────────────────────────
if ($Protection -in 'Malware', 'Both') {
    Write-Host "  --- Anti-Malware: $MalwarePolicyName ---" -ForegroundColor White
    $existing = Get-MalwareFilterPolicy -Identity $MalwarePolicyName -ErrorAction SilentlyContinue

    if ($existing) {
        Write-Host "  [EXISTS] Policy already present." -ForegroundColor DarkGray
        if (-not $UpdateExisting) {
            Write-Host "  Skipping (pass -UpdateExisting to update it in place)." -ForegroundColor DarkGray
        } elseif ($Apply -and $PSCmdlet.ShouldProcess($MalwarePolicyName, "Update malware filter policy")) {
            Set-MalwareFilterPolicy -Identity $MalwarePolicyName @malwareBaseline -ErrorAction Stop
            Write-Host "  [OK] Policy updated." -ForegroundColor Green
        } elseif (-not $Apply) {
            Write-Host "  Would update this policy with the baseline settings (re-run with -Apply -UpdateExisting)." -ForegroundColor Yellow
        }
    } else {
        Write-Host "  [MISSING] Policy does not exist yet." -ForegroundColor Yellow
        if ($Apply -and $PSCmdlet.ShouldProcess($MalwarePolicyName, "Create malware filter policy + rule")) {
            New-MalwareFilterPolicy -Name $MalwarePolicyName @malwareBaseline -ErrorAction Stop | Out-Null
            New-MalwareFilterRule -Name $MalwarePolicyName -MalwareFilterPolicy $MalwarePolicyName -RecipientDomainIs $Domains -Priority 0 -Enabled $true -ErrorAction Stop | Out-Null
            Write-Host "  [OK] Policy + rule created, scoped to: $($Domains -join ', ')" -ForegroundColor Green
        } elseif (-not $Apply) {
            Write-Host "  Would create this policy + rule scoped to: $($Domains -join ', ') (re-run with -Apply)." -ForegroundColor Yellow
        }
    }
    Write-Host ""
}

# ── Disconnect if we connected ──────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
