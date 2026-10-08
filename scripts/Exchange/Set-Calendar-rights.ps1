#Requires -Version 7.0
#Requires -Modules ExchangeOnlineManagement
<#
.SYNOPSIS
    Give a user access rights on another user's calendar.

.DESCRIPTION
    Adds (or updates) a mailbox folder permission on the default calendar of the
    target mailbox. The calendar is found by folder type, not by name, so it works
    for Dutch (\Agenda), French (\Calendrier), English (\Calendar) and any other
    mailbox language.

    -User and -TargetMailbox take a full UPN, or a name without domain; the tenant's
    default accepted domain is then appended.

    Exchange Online PowerShell, not Graph: Graph's calendarPermissions can only change
    another user's calendar with an app-only Calendars.ReadWrite permission, or when
    the signed-in admin already has rights on that calendar, and it has no equivalent
    of roles such as PublishingEditor or Contributor. Exchange accepts the admin's
    Exchange role for every mailbox, so a delegated sign-in is enough.

.PARAMETER User
    The user who receives the rights: UPN, or name without domain (e.g. Sjoerd.Kanon).

.PARAMETER TargetMailbox
    The mailbox whose calendar is shared: UPN, or name without domain (e.g. Jan.Jansen).

.PARAMETER AccessRights
    The access level to grant:
    Owner, PublishingEditor, Editor, PublishingAuthor, Author,
    NonEditingAuthor, Reviewer, Contributor, AvailabilityOnly, LimitedDetails

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
    .\Set-Calendar-rights.ps1 -User Sjoerd.Kanon -TargetMailbox Jan.Jansen -AccessRights Reviewer

.EXAMPLE
    .\Set-Calendar-rights.ps1 -User sjoerd@contoso.com -TargetMailbox jan@contoso.com -AccessRights Editor -WhatIf

.NOTES
    Connects to Exchange Online itself (delegated by default); an existing session for
    the same tenant is reused and left open.
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)]
    [string]$User,

    [Parameter(Mandatory)]
    [string]$TargetMailbox,

    [Parameter(Mandatory)]
    [ValidateSet(
        'Owner', 'PublishingEditor', 'Editor', 'PublishingAuthor', 'Author',
        'NonEditingAuthor', 'Reviewer', 'Contributor', 'AvailabilityOnly', 'LimitedDetails'
    )]
    [string]$AccessRights,

    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$AppOnly
)

# Delegated by default (device code and GDAP customer per load.config.ps1), app-only
# with -ClientId/-CertificateThumbprint or -AppOnly.
. (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')
$exo = Connect-M365Exchange -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

try {
    # Names without a domain get the tenant's default domain (replaces Get-MsolDomain).
    $defaultDomain = $null
    if ($User -notmatch '@' -or $TargetMailbox -notmatch '@') {
        $defaultDomain = (Get-AcceptedDomain | Where-Object { $_.Default }).DomainName
        if (-not $defaultDomain) { throw 'Could not read the default accepted domain.' }
    }
    $userUPN   = if ($User -match '@')          { $User }          else { "$User@$defaultDomain" }
    $targetUPN = if ($TargetMailbox -match '@') { $TargetMailbox } else { "$TargetMailbox@$defaultDomain" }

    # The default calendar by folder type, so the mailbox language does not matter.
    $calendar = Get-EXOMailboxFolderStatistics -Identity $targetUPN -FolderScope Calendar -ErrorAction Stop |
        Where-Object { $_.FolderType -eq 'Calendar' } |
        Select-Object -First 1
    if (-not $calendar) { throw "No default calendar found in mailbox $targetUPN." }
    $path = "${targetUPN}:\$($calendar.Name)"

    Write-Verbose "User     : $userUPN"
    Write-Verbose "Calendar : $path"
    Write-Verbose "Rights   : $AccessRights"

    # Add when the user has no entry yet, otherwise change the existing one.
    $existing = Get-MailboxFolderPermission -Identity $path -User $userUPN -ErrorAction SilentlyContinue
    if ($existing) {
        if ($PSCmdlet.ShouldProcess($path, "Set-MailboxFolderPermission ($AccessRights) for $userUPN (was: $($existing.AccessRights -join ', '))")) {
            Set-MailboxFolderPermission -Identity $path -User $userUPN -AccessRights $AccessRights -ErrorAction Stop | Out-Null
            Write-Host "Rights changed on $path for $userUPN : $AccessRights" -ForegroundColor Green
        }
    } else {
        if ($PSCmdlet.ShouldProcess($path, "Add-MailboxFolderPermission ($AccessRights) for $userUPN")) {
            Add-MailboxFolderPermission -Identity $path -User $userUPN -AccessRights $AccessRights -ErrorAction Stop | Out-Null
            Write-Host "Rights added on $path for $userUPN : $AccessRights" -ForegroundColor Green
        }
    }

    # Show the current rights as verification.
    Write-Host "`nCurrent calendar rights on ${path}:" -ForegroundColor Cyan
    Get-MailboxFolderPermission -Identity $path -ErrorAction SilentlyContinue |
        Where-Object { $_.User.DisplayName -notin @('Default', 'Anonymous') } |
        Select-Object Identity, User, AccessRights |
        Format-Table -AutoSize
}
finally {
    Disconnect-M365Exchange $exo
}
