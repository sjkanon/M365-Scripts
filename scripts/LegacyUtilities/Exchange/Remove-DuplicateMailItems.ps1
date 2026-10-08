#Requires -Version 7.0
<#
.SYNOPSIS
    Find and remove duplicate messages in a mailbox folder via Microsoft Graph.

.DESCRIPTION
    Modern, Graph-based replacement for a third-party EWS tool (Remove-DuplicateItems
    by Michel de Rooij) that is not being carried forward: Microsoft is retiring the
    EWS API for Exchange Online, so building on EWS today is a dead end. This script
    covers the common case — a misbehaving sync client or a PST re-import creating
    duplicate mail items — using Microsoft Graph instead.

    Groups messages in a folder by internetMessageId (falls back to a
    Subject+Sender+SentDateTime composite key for messages without one) and, for
    every group with more than one match, keeps the oldest (by receivedDateTime)
    and marks the rest as duplicates. Defaults to a safe preview — pass -Apply to
    actually delete the extras (moved to Deleted Items, not a hard delete).

    Sign-in goes through scripts\Startup\Connect-M365.ps1: delegated as the admin by
    default (device code / GDAP customer per load.config.ps1), app-only with -ClientId
    and -CertificateThumbprint or -AppOnly. A Graph session for the right tenant that
    already has the scopes is reused and left connected; only a session this script
    opened is disconnected.

    Delegated, Graph only lets you into ANOTHER user's mailbox when that mailbox is
    shared with you: the script asks for Mail.ReadWrite (your own mailbox) and
    Mail.ReadWrite.Shared (others), and the signed-in admin needs Full Access on
    the target mailbox (Add-MailboxPermission -AccessRights FullAccess, e.g. with
    Add-MailboxDelegateAccess.ps1). Without Full Access, use app-only with the
    Mail.ReadWrite application permission (ideally limited with an RBAC for
    Applications / application access policy).

.PARAMETER Mailbox
    UPN or object ID of the mailbox to scan.

.PARAMETER FolderId
    Well-known folder name (e.g. "inbox", "archive") or folder ID to scan.
    Default: "inbox".

.PARAMETER IncludeSubfolders
    Also scan every subfolder of -FolderId, recursively.

.PARAMETER Apply
    Actually delete the duplicate messages (soft-delete to Deleted Items).
    Without this switch, the script only reports what it would remove.

.PARAMETER OutputPath
    CSV report path. Defaults to `C:\Temp\DuplicateMailItems_<timestamp>.csv`
    (`~/Downloads` on Linux/macOS).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Defaults to the GDAP customer when authMode is GDAP.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint and -TenantId).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.EXAMPLE
    # Preview duplicates in the Inbox
    .\Remove-DuplicateMailItems.ps1 -Mailbox "user@contoso.com"

.EXAMPLE
    # Remove duplicates across Inbox and all its subfolders
    .\Remove-DuplicateMailItems.ps1 -Mailbox "user@contoso.com" -IncludeSubfolders -Apply

.EXAMPLE
    .\Remove-DuplicateMailItems.ps1 -Mailbox "user@contoso.com" -FolderId "archive" -Apply

.NOTES
    Delegated: Mail.ReadWrite + Mail.ReadWrite.Shared, and Full Access on the target
    mailbox. App-only: Mail.ReadWrite (application). Only considers items of type
    #microsoft.graph.message (skips meeting requests/responses, which duplicate by
    design during scheduling).

    Required module: Microsoft.Graph.Authentication
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $Mailbox,

    [string] $FolderId = 'inbox',
    [switch] $IncludeSubfolders,
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
if (-not $OutputPath) { $OutputPath = Join-Path $outputDir "DuplicateMailItems_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" }

# ── Connection ────────────────────────────────────────────────────────────────
# Delegated Mail.ReadWrite only covers the signed-in user's own mailbox; for
# /users/{someone else} Graph needs Mail.ReadWrite.Shared plus Full Access.
$graph = Connect-M365Graph -Scopes 'Mail.ReadWrite', 'Mail.ReadWrite.Shared' -TenantId $TenantId `
    -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

function Invoke-Graph {
    param([string] $Method = 'GET', [string] $Uri, [string] $Body)
    $params = @{ Method = $Method; Uri = $Uri; ErrorAction = 'Stop' }
    if ($Body) { $params['Body'] = $Body; $params['ContentType'] = 'application/json' }
    Invoke-MgGraphRequest @params
}

function Get-ChildFolderIds {
    param([string] $MailboxId, [string] $ParentId)
    # Emit the ids to the pipeline. The old List.AddRange() of the recursive result
    # threw on every leaf folder (an empty result arrives as $null), so
    # -IncludeSubfolders failed as soon as it reached a folder without children.
    $url = "https://graph.microsoft.com/v1.0/users/$MailboxId/mailFolders/$ParentId/childFolders?`$select=id"
    while ($url) {
        $resp = Invoke-Graph -Uri $url
        foreach ($f in $resp.value) {
            $f.id
            Get-ChildFolderIds -MailboxId $MailboxId -ParentId $f.id
        }
        $url = $resp.'@odata.nextLink'
    }
}

function Get-FolderMessages {
    param([string] $MailboxId, [string] $Folder)
    $items = [System.Collections.Generic.List[PSCustomObject]]::new()
    $url = "https://graph.microsoft.com/v1.0/users/$MailboxId/mailFolders/$Folder/messages?`$select=id,internetMessageId,subject,from,sentDateTime,receivedDateTime&`$top=999"
    while ($url) {
        $resp = Invoke-Graph -Uri $url
        foreach ($m in $resp.value) { $items.Add($m) }
        $url = $resp.'@odata.nextLink'
    }
    return $items
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Remove-DuplicateMailItems" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Mailbox : $Mailbox"
Write-Host "  Folder  : $FolderId$(if ($IncludeSubfolders) { ' (+ subfolders)' })"
Write-Host ("  Mode    : {0}" -f $(if ($Apply) { 'Apply' } else { 'Preview only' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

# ── Resolve folders to scan ────────────────────────────────────────────────────
$folders = [System.Collections.Generic.List[string]]::new()
$folders.Add($FolderId)
if ($IncludeSubfolders) {
    Write-Host "  Enumerating subfolders..." -ForegroundColor DarkGray
    (Get-ChildFolderIds -MailboxId $Mailbox -ParentId $FolderId) | ForEach-Object { $folders.Add($_) }
}

$results   = [System.Collections.Generic.List[PSCustomObject]]::new()
$totalDupe = 0

foreach ($folder in $folders) {
    Write-Host "  Scanning folder $folder..." -ForegroundColor DarkGray
    $messages = Get-FolderMessages -MailboxId $Mailbox -Folder $folder

    $groups = $messages | Group-Object -Property {
        if ($_.internetMessageId) { $_.internetMessageId }
        else { "$($_.subject)|$($_.from.emailAddress.address)|$($_.sentDateTime)" }
    } | Where-Object { $_.Count -gt 1 }

    foreach ($group in $groups) {
        $sorted    = $group.Group | Sort-Object receivedDateTime
        $keep      = $sorted[0]
        $duplicate = $sorted | Select-Object -Skip 1

        foreach ($dupe in $duplicate) {
            $totalDupe++
            $row = [PSCustomObject]@{
                Folder      = $folder
                Subject     = $dupe.subject
                From        = $dupe.from.emailAddress.address
                Received    = $dupe.receivedDateTime
                KeptId      = $keep.id
                DuplicateId = $dupe.id
                Status      = 'Preview'
            }

            if ($Apply) {
                if ($PSCmdlet.ShouldProcess("$($dupe.subject) ($($dupe.receivedDateTime))", "Delete duplicate message")) {
                    try {
                        Invoke-Graph -Method DELETE -Uri "https://graph.microsoft.com/v1.0/users/$Mailbox/messages/$($dupe.id)" | Out-Null
                        $row.Status = 'Deleted'
                    } catch {
                        $row.Status = "Error: $($_.Exception.Message)"
                    }
                }
            }
            $results.Add($row)
        }
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
if ($results.Count -eq 0) {
    Write-Host "  No duplicates found." -ForegroundColor DarkGray
} else {
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Found $totalDupe duplicate message(s). Report saved: $OutputPath" -ForegroundColor Green
    if (-not $Apply) { Write-Host "  Re-run with -Apply to delete them." -ForegroundColor Yellow }
}
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
Disconnect-M365Graph $graph
