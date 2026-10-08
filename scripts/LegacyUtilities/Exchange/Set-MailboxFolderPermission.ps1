#Requires -Version 7.0
<#
.SYNOPSIS
    Grant a user access rights on every folder in a mailbox (not just Calendar/Inbox).

.DESCRIPTION
    Applies the same folder permission to every folder in a target mailbox, skipping
    non-content system folders (Sync Issues and its children, Recoverable Items,
    Purges, Versions, Deletions). Useful for full-mailbox delegate access (e.g.
    covering an absent employee, or a shared "everything" mailbox) where
    Add-MailboxPermission -AccessRights FullAccess isn't enough because the
    delegate also needs to see private-item-suppressed or non-default folders in
    Outlook, or where only specific folder-level rights (not full mailbox access)
    should be granted.

    Stays on Exchange Online PowerShell: Graph has no API for mailbox folder
    permissions (Add/Set-MailboxFolderPermission); only the calendar has
    calendarPermission in Graph.
    Sign-in goes through scripts\Startup\Connect-M365.ps1 (Connect-M365Exchange):
    delegated as the admin by default (device code / GDAP customer via
    -DelegatedOrganization per load.config.ps1), app-only with -ClientId and
    -CertificateThumbprint or -AppOnly. An Exchange session for the tenant is reused
    and left connected; only a session this script opened is disconnected.

    Defaults to a safe preview — pass -Apply to actually grant the permission.

.PARAMETER Mailbox
    UPN or identity of the target mailbox whose folders will get the permission.

.PARAMETER User
    UPN or identity of the user to grant access to.

.PARAMETER AccessRights
    Folder permission level to grant. One of: Owner, PublishingEditor, Editor,
    PublishingAuthor, Author, NonEditingAuthor, Reviewer, Contributor,
    AvailabilityOnly, LimitedDetails, None.

.PARAMETER Apply
    Actually grant the permission. Without this switch, the script only lists the
    folders that would be affected.

.PARAMETER TenantId
    Tenant domain (contoso.onmicrosoft.com) or ID. Defaults to the GDAP customer when
    authMode is GDAP. App-only needs the domain form.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint and -TenantId).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.EXAMPLE
    # Preview — list every folder that would receive the permission
    .\Set-MailboxFolderPermission.ps1 -Mailbox "shared@contoso.com" -User "j.doe@contoso.com" -AccessRights Reviewer

.EXAMPLE
    .\Set-MailboxFolderPermission.ps1 -Mailbox "shared@contoso.com" -User "j.doe@contoso.com" -AccessRights Editor -Apply

.NOTES
    Skips: /Sync Issues, /Sync Issues/Conflicts, /Sync Issues/Local Failures,
    /Sync Issues/Server Failures, /Recoverable Items, /Deletions, /Purges,
    /Versions, /Top of Information Store (the folder root itself).

.NOTES
    Required module: ExchangeOnlineManagement
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $Mailbox,

    [Parameter(Mandatory)]
    [string] $User,

    [Parameter(Mandatory)]
    [ValidateSet('Owner', 'PublishingEditor', 'Editor', 'PublishingAuthor', 'Author', 'NonEditingAuthor', 'Reviewer', 'Contributor', 'AvailabilityOnly', 'LimitedDetails', 'None')]
    [string] $AccessRights,

    [switch] $Apply,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

$excludedFolders = @(
    '/Sync Issues',
    '/Sync Issues/Conflicts',
    '/Sync Issues/Local Failures',
    '/Sync Issues/Server Failures',
    '/Recoverable Items',
    '/Deletions',
    '/Purges',
    '/Versions'
)

# ── Connection ────────────────────────────────────────────────────────────────
# -Organization (used here before) only applies to app-only sign-in, so a GDAP
# customer was never reached; the helper uses -DelegatedOrganization for that.
$exo = Connect-M365Exchange -TenantId $TenantId -ClientId $ClientId `
    -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Set-MailboxFolderPermission" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Mailbox      : $Mailbox"
Write-Host "  User         : $User"
Write-Host "  AccessRights : $AccessRights"
Write-Host ("  Mode         : {0}" -f $(if ($Apply) { 'Apply' } else { 'Preview only' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

# ── Enumerate folders ─────────────────────────────────────────────────────────
$folders = @(Get-MailboxFolderStatistics -Identity $Mailbox -ErrorAction Stop |
    Where-Object { $excludedFolders -notcontains $_.FolderPath })

Write-Host "  Found $($folders.Count) folder(s) to process." -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($folder in $folders) {
    $folderPath = $folder.FolderPath.Replace('/', '\')
    if ($folderPath -match 'Top of Information Store') {
        $folderPath = $folderPath.Replace('\Top of Information Store', '\')
    }
    $identity = "$($Mailbox):$folderPath"

    if (-not $Apply) {
        Write-Host "  [PREVIEW] $identity" -ForegroundColor DarkGray
        $results.Add([PSCustomObject]@{ Folder = $identity; Status = 'Preview' })
        continue
    }

    if (-not $PSCmdlet.ShouldProcess($identity, "Grant $AccessRights to $User")) { continue }

    try {
        # Add-MailboxFolderPermission errors if a permission already exists for the
        # user on that folder — fall back to Set- in that case rather than failing.
        Add-MailboxFolderPermission -Identity $identity -User $User -AccessRights $AccessRights -ErrorAction Stop | Out-Null
        Write-Host "  [OK]   $identity" -ForegroundColor Green
        $results.Add([PSCustomObject]@{ Folder = $identity; Status = 'Granted' })
    } catch {
        if ($_.Exception.Message -match 'already has permission|already exists') {
            try {
                Set-MailboxFolderPermission -Identity $identity -User $User -AccessRights $AccessRights -ErrorAction Stop | Out-Null
                Write-Host "  [OK]   $identity (updated)" -ForegroundColor Green
                $results.Add([PSCustomObject]@{ Folder = $identity; Status = 'Updated' })
            } catch {
                Write-Host "  [WARN] $identity : $($_.Exception.Message)" -ForegroundColor Yellow
                $results.Add([PSCustomObject]@{ Folder = $identity; Status = "Error: $($_.Exception.Message)" })
            }
        } else {
            Write-Host "  [WARN] $identity : $($_.Exception.Message)" -ForegroundColor Yellow
            $results.Add([PSCustomObject]@{ Folder = $identity; Status = "Error: $($_.Exception.Message)" })
        }
    }
}

Write-Host ""
Write-Host "  Processed $($results.Count) folder(s)." -ForegroundColor Cyan
if (-not $Apply) { Write-Host "  Re-run with -Apply to grant the permission." -ForegroundColor Yellow }
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
Disconnect-M365Exchange $exo
