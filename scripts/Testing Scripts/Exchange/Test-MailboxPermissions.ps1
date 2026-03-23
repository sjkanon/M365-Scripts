#Requires -Version 5.1
<#
.SYNOPSIS
    Audit mailbox access permissions — Full Access, Send As, and Send on Behalf.

.DESCRIPTION
    Connects to Exchange Online and checks all three delegation types for
    a single mailbox or all mailboxes in the tenant:
      - Full Access    (Get-MailboxPermission)
      - Send As        (Get-RecipientPermission)
      - Send on Behalf (GrantSendOnBehalfTo from Get-EXOMailbox)

    Results are displayed in a table and optionally exported to CSV.

.PARAMETER Mailbox
    UPN of the mailbox to audit. If omitted, all user and shared mailboxes are checked.

.PARAMETER OutputPath
    Path to write the CSV report. Defaults to .\MailboxPermissions_<timestamp>.csv.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Test-MailboxPermissions.ps1

.EXAMPLE
    .\Test-MailboxPermissions.ps1 -Mailbox "shared@contoso.com"

.EXAMPLE
    .\Test-MailboxPermissions.ps1 -OutputPath "C:\Reports\permissions.csv"
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
Write-Host "   Mailbox Permissions Audit" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# ── Get mailboxes ─────────────────────────────────────────────────────────────
if ($Mailbox) {
    $mailboxes = @(Get-EXOMailbox -Identity $Mailbox -Properties GrantSendOnBehalfTo -ErrorAction Stop)
} else {
    Write-Host "  Retrieving mailboxes..." -ForegroundColor DarkGray
    $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox -Properties GrantSendOnBehalfTo)
}

Write-Host "  Checking permissions for $($mailboxes.Count) mailbox(es)..." -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($mbx in $mailboxes) {
    $mbxUpn = $mbx.UserPrincipalName

    # Full Access
    try {
        Get-MailboxPermission -Identity $mbxUpn -ErrorAction SilentlyContinue |
            Where-Object { $_.IsInherited -eq $false -and $_.User -notlike '*SELF*' } |
            ForEach-Object {
                $results.Add([PSCustomObject]@{
                    Mailbox        = $mbxUpn
                    PermissionType = 'FullAccess'
                    GrantedTo      = $_.User
                    Rights         = ($_.AccessRights -join ', ')
                    Deny           = $_.Deny
                })
            }
    } catch {
        Write-Host "  [WARN] FullAccess — $mbxUpn : $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Send As
    try {
        Get-RecipientPermission -Identity $mbxUpn -ErrorAction SilentlyContinue |
            Where-Object { $_.Trustee -notlike '*SELF*' -and $_.Trustee -notlike 'NT AUTHORITY*' } |
            ForEach-Object {
                $results.Add([PSCustomObject]@{
                    Mailbox        = $mbxUpn
                    PermissionType = 'SendAs'
                    GrantedTo      = $_.Trustee
                    Rights         = ($_.AccessRights -join ', ')
                    Deny           = $false
                })
            }
    } catch {
        Write-Host "  [WARN] SendAs — $mbxUpn : $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Send on Behalf
    if ($mbx.GrantSendOnBehalfTo) {
        foreach ($delegate in $mbx.GrantSendOnBehalfTo) {
            $results.Add([PSCustomObject]@{
                Mailbox        = $mbxUpn
                PermissionType = 'SendOnBehalf'
                GrantedTo      = $delegate
                Rights         = 'SendOnBehalf'
                Deny           = $false
            })
        }
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
if ($results.Count -eq 0) {
    Write-Host "  No delegated permissions found." -ForegroundColor DarkGray
} else {
    $results | Format-Table -AutoSize

    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "MailboxPermissions_$ts.csv"
    }

    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "  Checked $($mailboxes.Count) mailbox(es) — $($results.Count) permission entries found." -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
