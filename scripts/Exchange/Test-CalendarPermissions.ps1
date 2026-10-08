#Requires -Version 7.0
<#
.SYNOPSIS
    Audit calendar folder permissions across one or all mailboxes.

.DESCRIPTION
    Connects to Exchange Online and retrieves calendar folder permissions
    for a single mailbox or all mailboxes in the tenant.
    Results are displayed in a table and optionally exported to CSV.

.PARAMETER Mailbox
    UPN of the mailbox to audit. If omitted, all user and shared mailboxes are checked.

.PARAMETER OutputPath
    Path to write the CSV report. Defaults to .\CalendarPermissions_<timestamp>.csv.

.PARAMETER TenantId
    Tenant ID or domain. Defaults to the GDAP customer when load.config.ps1 sets
    authMode GDAP; otherwise you land in your own tenant. App-only needs a domain.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint). Without it
    you sign in delegated as yourself (device code per load.config.ps1).

.PARAMETER CertificateThumbprint
    Certificate for -ClientId.

.PARAMETER AppOnly
    App-only with ClientId and CertificateThumbprint for the tenant from
    graph.appid.json in the repo root.

.EXAMPLE
    .\Test-CalendarPermissions.ps1

.EXAMPLE
    .\Test-CalendarPermissions.ps1 -Mailbox "user@contoso.com"

.EXAMPLE
    .\Test-CalendarPermissions.ps1 -OutputPath "C:\Reports\calendar.csv"
#>
[CmdletBinding()]
param(
    [string] $Mailbox,
    [string] $OutputPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
# Delegated by default (device code and GDAP customer per load.config.ps1),
# app-only with -ClientId/-CertificateThumbprint or -AppOnly. Exchange Online
# PowerShell: Graph's calendarPermissions only reads another user's calendar
# app-only or when the admin already has rights on it, and has no Exchange roles
# such as PublishingEditor or the delegate flags.
. (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')
$exo = Connect-M365Exchange -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Calendar Permissions Audit" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# ── Get mailboxes ─────────────────────────────────────────────────────────────
if ($Mailbox) {
    $mailboxes = @(Get-EXOMailbox -Identity $Mailbox -ErrorAction Stop)
} else {
    Write-Host "  Retrieving mailboxes..." -ForegroundColor DarkGray
    $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox)
}

Write-Host "  Checking calendar permissions for $($mailboxes.Count) mailbox(es)..." -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($mbx in $mailboxes) {
    try {
        # Locale-independent: filter by FolderType instead of folder name
        $calFolder = Get-EXOMailboxFolderStatistics -Identity $mbx.UserPrincipalName -FolderScope Calendar |
            Where-Object { $_.FolderType -eq 'Calendar' } |
            Select-Object -First 1

        if (-not $calFolder) { continue }

        $folderPath = "$($mbx.UserPrincipalName):\$($calFolder.Name)"
        $perms = Get-EXOMailboxFolderPermission -Identity $folderPath -ErrorAction SilentlyContinue

        foreach ($perm in $perms) {
            $results.Add([PSCustomObject]@{
                Mailbox      = $mbx.UserPrincipalName
                CalendarName = $calFolder.Name
                User         = $perm.User.DisplayName
                AccessRights = ($perm.AccessRights -join ', ')
                SharingFlags = ($perm.SharingPermissionFlags -join ', ')
            })
        }
    } catch {
        Write-Host "  [WARN] $($mbx.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
if ($results.Count -eq 0) {
    Write-Host "  No calendar permissions found." -ForegroundColor DarkGray
} else {
    $results | Format-Table -AutoSize

    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "CalendarPermissions_$ts.csv"
    }

    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "  Checked $($mailboxes.Count) mailbox(es) — $($results.Count) permission entries found." -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
Disconnect-M365Exchange $exo
