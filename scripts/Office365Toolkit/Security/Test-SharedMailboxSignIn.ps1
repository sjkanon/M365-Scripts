#Requires -Version 5.1
<#
.SYNOPSIS
    Report (and optionally block) direct interactive sign-in to shared mailboxes.

.DESCRIPTION
    Shared mailboxes should normally only be accessed via delegated Full Access /
    Send As permissions on a licensed user's own account — never signed into
    directly. A shared mailbox with a directly usable, enabled account (and often
    no MFA, since it's rarely licensed) is a common soft target for account
    takeover. This script cross-references Exchange Online shared mailboxes with
    their Entra ID account state (Microsoft Graph) and reports which ones still
    allow direct sign-in.

    Default behavior is a safe report-only preview. Pass -Apply to disable direct
    sign-in (AccountEnabled = $false) for every shared mailbox account found
    enabled. This does not affect delegated access (Full Access / Send As) — only
    the mailbox's own account can no longer sign in directly.

.PARAMETER Mailbox
    UPN of a single shared mailbox to check. If omitted, all shared mailboxes are
    checked.

.PARAMETER Apply
    Actually disable sign-in for shared mailbox accounts found enabled. Without
    this switch, the script only reports which ones are enabled.

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\ (Windows) or ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Test-SharedMailboxSignIn.ps1

.EXAMPLE
    .\Test-SharedMailboxSignIn.ps1 -Apply

.EXAMPLE
    .\Test-SharedMailboxSignIn.ps1 -Mailbox "helpdesk@contoso.com" -Apply

.NOTES
    Capability inspired by o365-exo-sharedblock.ps1 from the retired
    directorcia/Office365 (CIAOPS) toolkit, which used the retired AzureAD
    module (Get-AzureADUser / Set-AzureADUser). This rewrite uses Microsoft
    Graph (Get-MgUser / Update-MgUser) for the account state, alongside
    Exchange Online for the shared mailbox list.

    Supports -WhatIf (SupportsShouldProcess).

    Required Graph scope: User.ReadWrite.All
    Required modules: ExchangeOnlineManagement, Microsoft.Graph.Users
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

# ── Connection (Exchange Online) ────────────────────────────────────────────
$script:ConnectedExoHere = $false
try {
    $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
} catch {
    $connectParams = @{ ShowBanner = $false }
    if ($TenantId) { $connectParams['Organization'] = $TenantId }
    Connect-ExchangeOnline @connectParams
    $script:ConnectedExoHere = $true
}

# ── Connection (Microsoft Graph) ────────────────────────────────────────────
$script:ConnectedGraphHere = $false
try {
    $null = Get-MgContext -ErrorAction Stop
    if (-not (Get-MgContext)) { throw }
} catch {
    $connectParams = @{ Scopes = @('User.ReadWrite.All') }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedGraphHere = $true
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Shared Mailbox Direct Sign-In Audit" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host ("  Mode : {0}" -f $(if ($Apply) { 'Apply (sign-in will be disabled)' } else { 'Preview only (no changes)' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

# ── Get shared mailboxes ─────────────────────────────────────────────────────
if ($Mailbox) {
    $mailboxes = @(Get-EXOMailbox -Identity $Mailbox -RecipientTypeDetails SharedMailbox -ErrorAction Stop)
} else {
    Write-Host "  Retrieving shared mailboxes..." -ForegroundColor DarkGray
    $mailboxes = @(Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited)
}

Write-Host "  Checking $($mailboxes.Count) shared mailbox(es)..." -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($mbx in $mailboxes) {
    try {
        $user = Get-MgUser -UserId $mbx.ExternalDirectoryObjectId -Property Id, DisplayName, UserPrincipalName, AccountEnabled -ErrorAction Stop
    } catch {
        Write-Host "  [WARN] Could not resolve Entra ID account for $($mbx.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Yellow
        continue
    }

    $entry = [PSCustomObject]@{
        DisplayName       = $mbx.DisplayName
        UserPrincipalName = $mbx.UserPrincipalName
        AccountEnabled    = $user.AccountEnabled
        Action            = 'None'
    }

    if ($user.AccountEnabled) {
        Write-Host ("  [RISK] {0,-40} — direct sign-in ENABLED" -f $mbx.UserPrincipalName) -ForegroundColor Red

        if ($Apply -and $PSCmdlet.ShouldProcess($mbx.UserPrincipalName, "Disable direct sign-in (AccountEnabled = false)")) {
            try {
                Update-MgUser -UserId $user.Id -AccountEnabled:$false -ErrorAction Stop
                $entry.Action = 'Disabled'
                Write-Host "         [OK] Sign-in disabled" -ForegroundColor Green
            } catch {
                $entry.Action = "Failed: $($_.Exception.Message)"
                Write-Host "         [WARN] Failed to disable: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host ("  [OK]   {0,-40} — direct sign-in already disabled" -f $mbx.UserPrincipalName) -ForegroundColor Green
    }

    $results.Add($entry)
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
$atRisk = @($results | Where-Object { $_.AccountEnabled })
if ($atRisk.Count -eq 0) {
    Write-Host "  No shared mailboxes with direct sign-in enabled." -ForegroundColor Green
} else {
    Write-Host "  $($atRisk.Count) of $($results.Count) shared mailbox(es) have direct sign-in enabled." -ForegroundColor $(if ($Apply) { 'Green' } else { 'Yellow' })
    if (-not $Apply) {
        Write-Host "  Re-run with -Apply to disable sign-in for those accounts." -ForegroundColor Yellow
    }
}

if (-not $OutputPath) {
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path $outputDir "SharedMailboxSignIn_$ts.csv"
}
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
Write-Host ""

# ── Disconnect if we connected ──────────────────────────────────────────────
if ($script:ConnectedExoHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
if ($script:ConnectedGraphHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
