#Requires -Version 5.1
<#
.SYNOPSIS
    Report (and optionally fix) unified audit log and per-mailbox audit logging gaps.

.DESCRIPTION
    Connects to Exchange Online and checks two things:
      - Whether the tenant-wide Unified Audit Log is enabled (Get-AdminAuditLogConfig)
      - For every mailbox: whether mailbox auditing is enabled (AuditEnabled) and, if so,
        whether the audit log retention (AuditLogAgeLimit) meets a minimum threshold

    Default behavior is a read-only report. Pass -Apply to enable the unified audit log
    if disabled, and to enable mailbox auditing / raise AuditLogAgeLimit to the minimum
    on flagged mailboxes. Supports -WhatIf.

.PARAMETER MinimumAuditLogAgeDays
    Minimum acceptable AuditLogAgeLimit in days. Default: 180.

.PARAMETER Apply
    Enable the unified audit log (if disabled) and fix flagged mailboxes. Without this
    switch, only reports.

.PARAMETER OutputPath
    CSV report path. Defaults to .\MailboxAuditingConfig_<timestamp>.csv.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Test-MailboxAuditingConfig.ps1

.EXAMPLE
    .\Test-MailboxAuditingConfig.ps1 -Apply -WhatIf

.EXAMPLE
    .\Test-MailboxAuditingConfig.ps1 -MinimumAuditLogAgeDays 365 -Apply

.NOTES
    Inspired by the capability list of the retired directorcia/patron toolkit
    (o365-audit-get.ps1, o365-audit-enable.ps1, o365-auditlog-retent.ps1,
    o365-mx-audit-get.ps1, o365-mx-audit-set.ps1), consolidated and rewritten from
    scratch — the originals were five separate console-dump scripts with no remediation
    and no CSV export.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [int] $MinimumAuditLogAgeDays = 180,
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
Write-Host "   Audit Logging Configuration" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# ── Unified Audit Log ────────────────────────────────────────────────────────
$auditConfig = Get-AdminAuditLogConfig
$uatEnabled = $auditConfig.UnifiedAuditLogIngestionEnabled
$color = if ($uatEnabled) { 'Green' } else { 'Red' }
Write-Host "  Unified Audit Log ingestion enabled: $uatEnabled" -ForegroundColor $color

if (-not $uatEnabled -and $Apply) {
    if ($PSCmdlet.ShouldProcess('Tenant', 'Enable Unified Audit Log ingestion')) {
        try {
            Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true -ErrorAction Stop
            Write-Host "  [OK] Unified Audit Log ingestion enabled." -ForegroundColor Green
        } catch {
            Write-Host "  [WARN] Could not enable Unified Audit Log: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
} elseif (-not $uatEnabled) {
    Write-Host "  Re-run with -Apply to enable it." -ForegroundColor Yellow
}
Write-Host ""

# ── Per-mailbox auditing ─────────────────────────────────────────────────────
Write-Host "  Retrieving mailboxes..." -ForegroundColor DarkGray
$mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox -Properties AuditEnabled, AuditLogAgeLimit)
Write-Host "  Checking $($mailboxes.Count) mailbox(es)..." -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($mbx in $mailboxes) {
    $ageLimitDays = $null
    if ($mbx.AuditLogAgeLimit) {
        try { $ageLimitDays = [timespan]::Parse([string]$mbx.AuditLogAgeLimit).Days } catch { $ageLimitDays = $null }
    }

    $needsEnable = -not $mbx.AuditEnabled
    $needsAgeBump = ($null -ne $ageLimitDays) -and ($ageLimitDays -lt $MinimumAuditLogAgeDays)
    $flagged = $needsEnable -or $needsAgeBump

    $results.Add([PSCustomObject]@{
        DisplayName       = $mbx.DisplayName
        UserPrincipalName = $mbx.UserPrincipalName
        AuditEnabled      = $mbx.AuditEnabled
        AuditLogAgeLimitDays = $ageLimitDays
        Flagged           = $flagged
        Issue             = if ($needsEnable) { 'Auditing disabled' } elseif ($needsAgeBump) { "Retention below $MinimumAuditLogAgeDays days" } else { '' }
    })

    if ($flagged) {
        Write-Host "  [FLAG] $($mbx.UserPrincipalName): $(if ($needsEnable) { 'auditing disabled' } else { "retention $ageLimitDays day(s)" })" -ForegroundColor Yellow

        if ($Apply) {
            $setParams = @{ Identity = $mbx.UserPrincipalName; AuditEnabled = $true }
            if ($needsAgeBump) { $setParams['AuditLogAgeLimit'] = "$MinimumAuditLogAgeDays.00:00:00" }
            if ($PSCmdlet.ShouldProcess($mbx.UserPrincipalName, 'Enable/raise mailbox auditing')) {
                try {
                    Set-Mailbox @setParams -ErrorAction Stop
                    Write-Host "  [OK] Fixed $($mbx.UserPrincipalName)" -ForegroundColor Green
                } catch {
                    Write-Host "  [WARN] Could not fix $($mbx.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
if (-not $OutputPath) {
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path $outputDir "MailboxAuditingConfig_$ts.csv"
}
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "  Report saved: $OutputPath" -ForegroundColor Green

$flaggedCount = @($results | Where-Object { $_.Flagged }).Count
Write-Host ""
Write-Host ("  {0} mailbox(es) checked — {1} flagged" -f $results.Count, $flaggedCount) -ForegroundColor $(if ($flaggedCount -gt 0) { 'Yellow' } else { 'Cyan' })
if ($flaggedCount -gt 0 -and -not $Apply) {
    Write-Host "  Re-run with -Apply to fix flagged mailboxes." -ForegroundColor Yellow
}
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
