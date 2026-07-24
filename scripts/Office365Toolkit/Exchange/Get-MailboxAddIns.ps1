#Requires -Version 5.1
<#
.SYNOPSIS
    Report Outlook add-ins installed per mailbox.

.DESCRIPTION
    Connects to Exchange Online and lists the Outlook add-ins (Get-App) present
    on each mailbox — both centrally deployed and user/sideloaded add-ins.
    Useful for spotting unexpected or unapproved add-ins, which are a known
    phishing/consent-grant vector (a malicious add-in can read mail the same
    way a malicious OAuth app can). Results are exported to CSV.

.PARAMETER Mailbox
    UPN of a single mailbox to check. If omitted, all user and shared mailboxes
    are checked.

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\ (Windows) or ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Get-MailboxAddIns.ps1

.EXAMPLE
    .\Get-MailboxAddIns.ps1 -Mailbox "user@contoso.com"

.NOTES
    Capability inspired by o365-exo-addins.ps1 from the retired
    directorcia/Office365 (CIAOPS) toolkit, which printed a per-mailbox table
    to the console only. This rewrite adds CSV export and tenant-wide
    parameterization, consistent with this repo's other Exchange audit
    scripts.

    Get-App reports add-ins visible to the currently connected admin/mailbox
    context; some organizations restrict what it can enumerate for other
    users' mailboxes depending on RBAC scope.

    Required module: ExchangeOnlineManagement
#>
[CmdletBinding()]
param(
    [string] $Mailbox,
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
Write-Host "   Mailbox Add-In Report" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# ── Get mailboxes ─────────────────────────────────────────────────────────────
if ($Mailbox) {
    $mailboxes = @(Get-EXOMailbox -Identity $Mailbox -ErrorAction Stop)
} else {
    Write-Host "  Retrieving mailboxes..." -ForegroundColor DarkGray
    $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox)
}

Write-Host "  Checking add-ins for $($mailboxes.Count) mailbox(es)..." -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($mbx in $mailboxes) {
    try {
        $apps = Get-App -Mailbox $mbx.UserPrincipalName -ErrorAction Stop
    } catch {
        Write-Host "  [WARN] $($mbx.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Yellow
        continue
    }

    if (-not $apps) { continue }

    foreach ($app in $apps) {
        $results.Add([PSCustomObject]@{
            Mailbox      = $mbx.UserPrincipalName
            DisplayName  = $app.DisplayName
            ProviderName = $app.ProviderName
            Enabled      = $app.Enabled
            AppVersion   = $app.AppVersion
            AppId        = $app.AppId
        })
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
if ($results.Count -eq 0) {
    Write-Host "  No add-ins found on any checked mailbox." -ForegroundColor DarkGray
} else {
    $results | Sort-Object Mailbox, DisplayName | Format-Table Mailbox, DisplayName, ProviderName, Enabled, AppVersion -AutoSize

    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "MailboxAddIns_$ts.csv"
    }
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

Write-Host ""
Write-Host ("  {0} mailbox(es) checked — {1} add-in installation(s) found." -f $mailboxes.Count, $results.Count) -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ──────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
