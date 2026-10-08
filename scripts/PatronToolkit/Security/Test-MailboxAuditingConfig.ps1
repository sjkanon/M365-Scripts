#Requires -Version 7.0
<#
.SYNOPSIS
    Report (and optionally fix) unified audit log and per-mailbox audit logging gaps.

.DESCRIPTION
    Connects to Exchange Online and checks two things:
      - Whether the tenant-wide Unified Audit Log is enabled (Get-AdminAuditLogConfig)
      - Whether mailbox auditing is switched off org-wide (Get-OrganizationConfig
        AuditDisabled) - reported only, never changed
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
    Tenant domain (contoso.onmicrosoft.com) or ID. Defaults to the GDAP customer
    (load.config.ps1) or your own tenant. App-only Exchange sign-in needs the domain.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint). Without it the
    script signs in delegated, as you.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

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

    Sign-in: Exchange Online through scripts\Startup\Connect-M365.ps1 - delegated by
    default (an Exchange admin / Compliance admin role; device code and the GDAP customer
    via -DelegatedOrganization per load.config.ps1), app-only with
    -ClientId/-CertificateThumbprint or -AppOnly (Exchange.ManageAsApp plus an Exchange
    role on the app). Stays on Exchange Online: Microsoft Graph has no API for the
    unified audit log switch or the per-mailbox audit settings.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [int] $MinimumAuditLogAgeDays = 180,
    [switch] $Apply,
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
$exo = Connect-M365Exchange -TenantId $TenantId -ClientId $ClientId `
    -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Audit Logging Configuration" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# ── Unified Audit Log ────────────────────────────────────────────────────────
try {
    $auditConfig = Get-AdminAuditLogConfig -ErrorAction Stop
} catch {
    Write-Host "  [ERROR] Could not read the audit log configuration: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-M365Exchange $exo
    exit 1
}
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

# Mailbox auditing on by default is an org switch: with AuditDisabled = True no mailbox is
# audited, whatever its own AuditEnabled says. Reported only - turning it back on is a
# deliberate org decision (Set-OrganizationConfig -AuditDisabled $false).
try {
    $orgAuditDisabled = (Get-OrganizationConfig -ErrorAction Stop).AuditDisabled
    Write-Host "  Org-wide mailbox auditing disabled (AuditDisabled): $orgAuditDisabled" -ForegroundColor $(if ($orgAuditDisabled) { 'Red' } else { 'Green' })
    if ($orgAuditDisabled) {
        Write-Host "  Mailbox auditing is off for the whole organization; per-mailbox AuditEnabled has no effect until Set-OrganizationConfig -AuditDisabled `$false." -ForegroundColor Yellow
    }
} catch {
    Write-Host "  [WARN] Could not read Get-OrganizationConfig: $($_.Exception.Message)" -ForegroundColor Yellow
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
Disconnect-M365Exchange $exo
