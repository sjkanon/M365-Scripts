#Requires -Version 5.1
<#
.SYNOPSIS
    Grant Full Access and/or Send As delegate rights on one mailbox, a CSV list of
    mailboxes, or every mailbox in the tenant.

.DESCRIPTION
    Consolidates three near-identical old ad hoc scripts (grant Full Access +
    Send As to one shared mailbox for one user; the same for a CSV list of
    "shared" mailboxes; and a blanket grant across every mailbox in the org) into
    one parameterized, dry-run-by-default script.

    Connects to Exchange Online automatically if no session is active; reuses an
    existing session if already connected.

.PARAMETER User
    UPN of the user (or group) to grant delegate access to.

.PARAMETER Mailbox
    Identity (UPN/alias/SMTP) of a single target mailbox.

.PARAMETER CsvPath
    Path to a CSV with a column identifying mailboxes — accepts
    PrimarySmtpAddress, EmailAddress, Mailbox, or Identity (first match wins) —
    or a plain TXT file with one identity per line.

.PARAMETER AllMailboxes
    Grant access to every user and shared mailbox in the tenant. Use with care —
    this is a broad, high-impact grant; strongly recommended to run without
    -Apply first to review the full scope.

.PARAMETER AccessRights
    Which rights to grant: FullAccess, SendAs, or Both (default: Both).

.PARAMETER AutoMapping
    Enable Outlook auto-mapping for Full Access grants (mailbox appears
    automatically in the delegate's Outlook profile). Default: off — matches the
    safer behavior recommended for shared/delegate access at scale, since
    auto-mapping every mailbox into one person's Outlook profile can make it slow
    to start.

.PARAMETER Apply
    Actually grant the permissions. Without this switch, the script only lists
    what would be granted.

.PARAMETER OutputPath
    CSV report path. Defaults to `C:\Temp\MailboxDelegateAccess_<timestamp>.csv`
    (`~/Downloads` on Linux/macOS).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    # Preview — single mailbox
    .\Add-MailboxDelegateAccess.ps1 -Mailbox "sales@contoso.com" -User "j.doe@contoso.com"

.EXAMPLE
    .\Add-MailboxDelegateAccess.ps1 -Mailbox "sales@contoso.com" -User "j.doe@contoso.com" -Apply

.EXAMPLE
    # Bulk — CSV of shared mailboxes, one delegate for all of them
    .\Add-MailboxDelegateAccess.ps1 -CsvPath .\mailboxes.csv -User "j.doe@contoso.com" -Apply

.EXAMPLE
    # Every mailbox in the tenant — review scope first
    .\Add-MailboxDelegateAccess.ps1 -AllMailboxes -User "helpdesk@contoso.com"
    .\Add-MailboxDelegateAccess.ps1 -AllMailboxes -User "helpdesk@contoso.com" -Apply

.EXAMPLE
    # Full Access only, with auto-mapping
    .\Add-MailboxDelegateAccess.ps1 -Mailbox "sales@contoso.com" -User "j.doe@contoso.com" -AccessRights FullAccess -AutoMapping -Apply

.NOTES
    Required module: ExchangeOnlineManagement
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Single')]
param(
    [Parameter(Mandatory)]
    [string] $User,

    [Parameter(ParameterSetName = 'Single', Mandatory)]
    [string] $Mailbox,

    [Parameter(ParameterSetName = 'Csv', Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $CsvPath,

    [Parameter(ParameterSetName = 'All', Mandatory)]
    [switch] $AllMailboxes,

    [ValidateSet('FullAccess', 'SendAs', 'Both')]
    [string] $AccessRights = 'Both',

    [switch] $AutoMapping,
    [switch] $Apply,
    [string] $OutputPath,
    [string] $TenantId
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
if (-not $OutputPath) { $OutputPath = Join-Path $outputDir "MailboxDelegateAccess_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" }

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
Write-Host "   Add-MailboxDelegateAccess" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  User         : $User"
Write-Host "  AccessRights : $AccessRights"
Write-Host ("  Mode         : {0}" -f $(if ($Apply) { 'Apply' } else { 'Preview only' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

# ── Resolve target mailboxes ──────────────────────────────────────────────────
$targets = [System.Collections.Generic.List[string]]::new()

switch ($PSCmdlet.ParameterSetName) {
    'Single' { $targets.Add($Mailbox) }
    'Csv' {
        $ext = [System.IO.Path]::GetExtension($CsvPath).ToLower()
        if ($ext -eq '.csv') {
            $raw = Import-Csv -Path $CsvPath
            $col = $raw[0].PSObject.Properties.Name |
                Where-Object { $_ -match '^(PrimarySmtpAddress|EmailAddress|Mailbox|Identity)$' } |
                Select-Object -First 1
            if (-not $col) {
                Write-Error "CSV must have a PrimarySmtpAddress, EmailAddress, Mailbox, or Identity column. Found: $($raw[0].PSObject.Properties.Name -join ', ')"
                if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
                exit 1
            }
            $raw.$col | Where-Object { $_ } | ForEach-Object { $targets.Add($_.Trim()) }
        } else {
            Get-Content -Path $CsvPath | Where-Object { $_.Trim() -and $_ -notmatch '^\s*#' } | ForEach-Object { $targets.Add($_.Trim()) }
        }
    }
    'All' {
        Write-Host "  Retrieving all mailboxes..." -ForegroundColor DarkGray
        (Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox).PrimarySmtpAddress |
            ForEach-Object { $targets.Add($_) }
    }
}

$targets = $targets | Sort-Object -Unique
Write-Host "  Target mailbox(es): $($targets.Count)" -ForegroundColor DarkGray
Write-Host ""

if (-not $Apply -and $targets.Count -gt 0) {
    Write-Host "  Preview — no changes will be made. Re-run with -Apply to grant access." -ForegroundColor Yellow
    Write-Host ""
}

# ── Grant access ───────────────────────────────────────────────────────────────
$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($mbx in $targets) {
    $status = [ordered]@{ Mailbox = $mbx; User = $User; FullAccess = 'n/a'; SendAs = 'n/a' }

    if (-not $Apply) {
        if ($AccessRights -in 'FullAccess', 'Both') { $status.FullAccess = 'Preview' }
        if ($AccessRights -in 'SendAs', 'Both') { $status.SendAs = 'Preview' }
        Write-Host "  [PREVIEW] $mbx" -ForegroundColor DarkGray
        $results.Add([PSCustomObject]$status)
        continue
    }

    if (-not $PSCmdlet.ShouldProcess($mbx, "Grant $AccessRights to $User")) { continue }

    if ($AccessRights -in 'FullAccess', 'Both') {
        try {
            Add-MailboxPermission -Identity $mbx -User $User -AccessRights FullAccess `
                -InheritanceType All -AutoMapping:$AutoMapping.IsPresent -Confirm:$false -ErrorAction Stop | Out-Null
            $status.FullAccess = 'Granted'
        } catch {
            $status.FullAccess = "Error: $($_.Exception.Message)"
        }
    }

    if ($AccessRights -in 'SendAs', 'Both') {
        try {
            Add-RecipientPermission -Identity $mbx -Trustee $User -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
            $status.SendAs = 'Granted'
        } catch {
            $status.SendAs = "Error: $($_.Exception.Message)"
        }
    }

    $color = if ($status.FullAccess -match '^Error' -or $status.SendAs -match '^Error') { 'Yellow' } else { 'Green' }
    Write-Host ("  [{0}] {1}  FullAccess={2}  SendAs={3}" -f $(if ($color -eq 'Green') { 'OK' } else { 'WARN' }), $mbx, $status.FullAccess, $status.SendAs) -ForegroundColor $color
    $results.Add([PSCustomObject]$status)
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
if ($results.Count -gt 0) {
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
} else {
    Write-Host "  No target mailboxes resolved." -ForegroundColor DarkGray
}
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
