#Requires -Version 5.1
<#
.SYNOPSIS
    Audit mailboxes with external mail forwarding configured.

.DESCRIPTION
    Connects to Exchange Online and checks every mailbox for a ForwardingSMTPAddress
    that points to a domain outside the tenant's accepted domains.
    External forwards are a common security/compliance risk and should be reviewed
    regularly.

    Results are displayed and exported to CSV.

.PARAMETER Mailbox
    UPN of a single mailbox to check. If omitted, all mailboxes are checked.

.PARAMETER OutputPath
    CSV report path. Defaults to .\ExternalForwards_<timestamp>.csv.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Get-ExternalForwards.ps1

.EXAMPLE
    .\Get-ExternalForwards.ps1 -Mailbox "user@contoso.com"

.EXAMPLE
    .\Get-ExternalForwards.ps1 -OutputPath "C:\Reports\forwards.csv"
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
Write-Host "   External Forwarding Audit" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# ── Get accepted domains (internal) ──────────────────────────────────────────
$acceptedDomains = (Get-AcceptedDomain).DomainName
Write-Host "  Internal domains: $($acceptedDomains -join ', ')" -ForegroundColor DarkGray
Write-Host ""

# ── Get mailboxes ─────────────────────────────────────────────────────────────
if ($Mailbox) {
    $mailboxes = @(Get-EXOMailbox -Identity $Mailbox -Properties ForwardingSMTPAddress, ForwardingAddress, DeliverToMailboxAndForward -ErrorAction Stop)
} else {
    Write-Host "  Retrieving mailboxes..." -ForegroundColor DarkGray
    $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties ForwardingSMTPAddress, ForwardingAddress, DeliverToMailboxAndForward)
}

Write-Host "  Checking $($mailboxes.Count) mailbox(es) for external forwards..." -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($mbx in $mailboxes) {
    $forwardAddr = $mbx.ForwardingSMTPAddress

    if (-not $forwardAddr) { continue }

    # Strip SMTP: prefix
    $email  = ($forwardAddr -split 'SMTP:')[-1].Trim()
    $domain = ($email -split '@')[-1].Trim()

    if ($acceptedDomains -notcontains $domain) {
        $results.Add([PSCustomObject]@{
            Mailbox                   = $mbx.UserPrincipalName
            DisplayName               = $mbx.DisplayName
            ForwardingAddress         = $email
            ForwardingDomain          = $domain
            DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
        })

        Write-Host ("  [FORWARD] {0,-40} → {1}" -f $mbx.UserPrincipalName, $email) -ForegroundColor Yellow
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
if ($results.Count -eq 0) {
    Write-Host "  No external forwards found." -ForegroundColor Green
} else {
    Write-Host "  $($results.Count) external forward(s) found." -ForegroundColor Yellow

    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "ExternalForwards_$ts.csv"
    }

    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "  Checked $($mailboxes.Count) mailbox(es)." -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
