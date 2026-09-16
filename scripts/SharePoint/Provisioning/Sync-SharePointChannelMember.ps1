#Requires -Version 7.0
<#
.SYNOPSIS
    Make an Entra ID security group the source of truth for who is in a private Teams
    channel. Reads the group, writes the channel roster. Supports -WhatIf.

.DESCRIPTION
    A private channel cannot be given rights through a group: Teams tracks its
    membership one person at a time, and Graph only accepts individual users there.
    Adding the group to the channel's SharePoint site instead is not the answer - the
    channel roster is synced back over it, and anyone who gains access that way can
    reach the files while the channel stays invisible to them in Teams.

    So this closes the gap the other way round. You manage the group; this puts the
    people in it into the channel.

        1. Resolve the team from the configuration's team.mailNickname.
        2. Per private channel, work out who should be in it from the configured
           groups - nested groups included, service principals and other groups
           skipped, only real users end up in a channel.
        3. Make sure each of them is a member of the parent team first. Teams refuses
           a private-channel member who is not on the team, and the error it gives
           does not say so.
        4. Add the ones who are missing. With -Prune, remove the ones who are in the
           channel but not in any of its groups - owners are never removed.

    There is no read-only role
    --------------------------
    A private channel has owners and members, and members may post, edit and delete
    files. A group named -RO therefore cannot mean "may look" here; everyone this
    script adds can write. The configuration decides which groups feed the channel,
    and the run says per group how many people it brought in, so an -RO group feeding
    a private channel is visible rather than assumed.

    If read-only genuinely matters for a pillar, a private channel is the wrong shape
    for it - use a document library with its own permissions, where Read is a real
    role.

    Which groups feed which channel
    -------------------------------
    From the container's channelMembers section:

        "channelMembers": [ { "group": "SG-X-MGMT-RW" }, { "group": "SG-X-MGMT-RO" } ]

    A configuration written before that existed has no such section, so as a fallback
    every configured group whose name contains the container's key is used, and the
    run reports that it fell back rather than doing it quietly.

.PARAMETER ConfigPath
    Path to the structure configuration JSON. Default: the single *.config.json next
    to this script that no longer contains the CHANGEME placeholders.

.PARAMETER Container
    Only sync these containers (keys from the containers section). Default: every
    private channel.

.PARAMETER Prune
    Also remove channel members who are not in any of the channel's groups. Owners are
    never removed, and neither is the account running the script.

.PARAMETER Tenant
    Tenant name or ID. Defaults to the tenant in the configuration file.

.PARAMETER ClientId
    Client ID of an app registration to sign in to Graph with. Without it the
    Microsoft Graph PowerShell app is used and the scopes are requested on the spot.

.PARAMETER ReportPath
    CSV of every add and removal. Default:
    C:\Temp\SharePointStructure_ChannelMembers_<timestamp>.csv

.PARAMETER Quiet
    Print only the summary and anything that needs attention.

.EXAMPLE
    # What would change in the private channels?
    .\Sync-SharePointChannelMember.ps1 -WhatIf

.EXAMPLE
    # Bring the groups' members into their channels
    .\Sync-SharePointChannelMember.ps1

.EXAMPLE
    # Make the group authoritative both ways - anyone not in it loses the channel
    .\Sync-SharePointChannelMember.ps1 -Prune

.NOTES
    Author  : Sjoerd Kanon
    Requires: Microsoft.Graph.Authentication
    Rights  : delegated Group.Read.All, TeamMember.ReadWrite.All,
              ChannelMember.ReadWrite.All
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $ConfigPath,
    [string[]] $Container,
    [switch] $Prune,
    [string] $Tenant,
    [string] $ClientId,
    [string] $ReportPath,
    [switch] $Quiet
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

. (Join-Path $PSScriptRoot 'SharePointStructure.Common.ps1')

if (-not $ConfigPath) {
    $candidates = @(Get-ChildItem -Path $PSScriptRoot -Filter '*.config.json' -ErrorAction SilentlyContinue |
                    Where-Object { (Get-Content $_.FullName -Raw) -notmatch 'CHANGEME' })
    if ($candidates.Count -eq 1) { $ConfigPath = $candidates[0].FullName }
    elseif ($candidates.Count -eq 0) { throw 'No configuration found - run New-StructureConfig.ps1 first.' }
    else { throw ("More than one configuration here - pass -ConfigPath. Found: {0}" -f (($candidates | ForEach-Object { $_.Name }) -join ', ')) }
}

$config = Import-StructureConfig -Path $ConfigPath
if (-not $Tenant) { $Tenant = $config.tenant }

$team = Get-ConfigValue $config 'team'
if (-not $team) { throw "The configuration has no team section ($ConfigPath)." }

$simulate = [bool] $WhatIfPreference

$private = @($config.containers | Where-Object { (Get-ConfigValue $_ 'channelType') -eq 'Private' })
if ($Container) {
    $known = @($config.containers | ForEach-Object { $_.key })
    foreach ($key in $Container) {
        if ($key -notin $known) { throw "Unknown container '$key'. Known: $($known -join ', ')" }
    }
    $private = @($private | Where-Object { $_.key -in $Container })
}

$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not $ReportPath) {
    $ReportPath = Join-Path $outputDir "SharePointStructure_ChannelMembers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
}

$report  = [System.Collections.Generic.List[object]]::new()
$added   = 0
$removed = 0
$warned  = 0

function Invoke-Graph {
    <# Graph with the reason Graph gives, not just the status code. #>
    param([Parameter(Mandatory)] [string] $Url, [string] $Method = 'GET', $Body)

    $splat = @{ Uri = "https://graph.microsoft.com/$Url"; Method = $Method; ErrorAction = 'Stop' }
    if ($Body) {
        $splat['Body']        = ($Body | ConvertTo-Json -Depth 8)
        $splat['ContentType'] = 'application/json'
    }
    try {
        return Invoke-MgGraphRequest @splat
    } catch {
        $detail = ''
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            try   { $detail = ($_.ErrorDetails.Message | ConvertFrom-Json).error.message }
            catch { $detail = $_.ErrorDetails.Message }
        }
        if (-not $detail) { $detail = $_.Exception.Message }
        throw "$Method $Url -> $detail"
    }
}

function Get-GraphPage {
    <# Every page of a collection, followed through @odata.nextLink. #>
    param([Parameter(Mandatory)] [string] $Url)

    $items = [System.Collections.Generic.List[object]]::new()
    $next  = $Url
    while ($next) {
        $page = Invoke-Graph -Url $next
        foreach ($item in @($page.value)) { $items.Add($item) }
        $link = Get-ConfigValue $page '@odata.nextLink'
        # nextLink comes back absolute; Invoke-Graph prefixes the host itself.
        $next = if ($link) { $link -replace '^https://graph\.microsoft\.com/', '' } else { $null }
    }
    return ,$items
}

function Get-ChannelGroup {
    <#
        The groups that decide who is in this channel. Explicit channelMembers wins;
        otherwise fall back to configured groups named after the container, and say so.
    #>
    param($Definition)

    $explicit = @(Get-ConfigValue $Definition 'channelMembers' @() | ForEach-Object { $_.group })
    if ($explicit.Count -gt 0) { return ,$explicit }

    $guessed = @(Get-ConfigValue $config 'groups' @() |
                 Where-Object { $_.displayName -like "*$($Definition.key)*" } |
                 ForEach-Object { $_.displayName })
    if ($guessed.Count -gt 0) {
        Write-Warn "'$($Definition.key)' has no channelMembers in the configuration - falling back to groups named after it: $($guessed -join ', ')"
        $script:warned++
    }
    return ,$guessed
}

function Get-GroupUser {
    <#
        The real users in a group, nested groups included. Anything that is not a user
        is dropped: a channel roster holds people.
    #>
    param([Parameter(Mandatory)] [string] $GroupName)

    $escaped = $GroupName -replace "'", "''"
    $found   = Invoke-Graph -Url "v1.0/groups?`$filter=displayName eq '$escaped'&`$select=id,displayName"
    $group   = @($found.value) | Select-Object -First 1
    if (-not $group) {
        Write-Warn "group '$GroupName' does not exist - skipped"
        $script:warned++
        return ,@()
    }

    $members = Get-GraphPage -Url "v1.0/groups/$($group.id)/transitiveMembers?`$select=id,displayName,userPrincipalName"
    $users   = @($members | Where-Object { (Get-ConfigValue $_ '@odata.type') -eq '#microsoft.graph.user' })
    return ,$users
}

try {
    if (-not $Quiet) {
        Write-Host ''
        Write-Host "  Client : $(Get-ConfigValue $config 'client' '(unnamed)')" -ForegroundColor Cyan
        Write-Host "  Team   : $($team.displayName)" -ForegroundColor Cyan
        Write-Host "  Scope  : $(if ($private.Count) { ($private | ForEach-Object { $_.title }) -join ', ' } else { 'no private channels in the configuration' })" -ForegroundColor Cyan
        Write-Host "  Mode   : $(if ($simulate) { '-WhatIf - nothing will be changed' } elseif ($Prune) { 'APPLY + PRUNE' } else { 'APPLY - only adds' })" -ForegroundColor Cyan
    }

    if ($private.Count -eq 0) {
        Write-Head 'Summary'
        Write-Skip 'No private channels to sync.'
        Write-Host ''
        exit 0
    }

    Connect-StructureGraph -Tenant $Tenant -ClientId $ClientId -AuthenticationOnly -Scopes @(
        'Group.Read.All'
        'Directory.Read.All'
        'TeamMember.ReadWrite.All'
        'ChannelMember.ReadWrite.All'
    )

    $escaped = $team.mailNickname -replace "'", "''"
    $found   = Invoke-Graph -Url "v1.0/groups?`$filter=mailNickname eq '$escaped'&`$select=id,displayName"
    $teamId  = (@($found.value) | Select-Object -First 1).id
    if (-not $teamId) { throw "Team '$($team.displayName)' (alias $($team.mailNickname)) not found." }

    $channels    = @((Invoke-Graph -Url "v1.0/teams/$teamId/channels").value)
    $teamMembers = Get-GraphPage -Url "v1.0/teams/$teamId/members"
    $teamUserIds = @($teamMembers | ForEach-Object { Get-ConfigValue $_ 'userId' } | Where-Object { $_ })

    foreach ($entry in $private) {
        Write-Head "Channel '$($entry.title)'"

        $channel = $channels | Where-Object { $_.displayName -eq $entry.title } | Select-Object -First 1
        if (-not $channel) {
            Write-Bad "channel '$($entry.title)' not found - run New-SharePointTeam.ps1 first"
            $warned++
            continue
        }

        # -- who should be in it
        $wanted = @{}
        foreach ($groupName in (Get-ChannelGroup -Definition $entry)) {
            $users = Get-GroupUser -GroupName $groupName
            foreach ($user in $users) { $wanted[$user.id] = $user }
            Write-Ok "group '$groupName': $($users.Count) user(s)"
        }
        if ($wanted.Count -eq 0) {
            Write-Skip 'no users in the configured groups - nothing to do'
            continue
        }

        # -- who is in it
        $current   = Get-GraphPage -Url "v1.0/teams/$teamId/channels/$($channel.id)/members"
        $currentIds = @($current | ForEach-Object { Get-ConfigValue $_ 'userId' } | Where-Object { $_ })

        foreach ($id in $wanted.Keys) {
            if ($id -in $currentIds) { continue }
            $user = $wanted[$id]
            $who  = Get-ConfigValue $user 'userPrincipalName' (Get-ConfigValue $user 'displayName' $id)

            if (-not $PSCmdlet.ShouldProcess("$($entry.title) / $who", 'Add to private channel')) { continue }

            try {
                # Teams refuses a private-channel member who is not on the team, and
                # the message it returns does not mention that - so put them on the
                # team first.
                if ($id -notin $teamUserIds) {
                    Invoke-Graph -Url "v1.0/teams/$teamId/members" -Method POST -Body @{
                        '@odata.type'     = '#microsoft.graph.aadUserConversationMember'
                        'user@odata.bind' = "https://graph.microsoft.com/v1.0/users('$id')"
                        roles             = @()
                    } | Out-Null
                    $teamUserIds += $id
                    Write-Change "$who added to the team"
                }

                Invoke-Graph -Url "v1.0/teams/$teamId/channels/$($channel.id)/members" -Method POST -Body @{
                    '@odata.type'     = '#microsoft.graph.aadUserConversationMember'
                    'user@odata.bind' = "https://graph.microsoft.com/v1.0/users('$id')"
                    roles             = @()
                } | Out-Null

                Write-Change "$who added to '$($entry.title)'"
                $report.Add([PSCustomObject]@{ Channel = $entry.title; Action = 'added'; User = $who })
                $added++
            } catch {
                Write-Bad "could not add $who : $($_.Exception.Message)"
                $warned++
            }
        }

        # -- who should not be
        if ($Prune) {
            foreach ($member in $current) {
                $id = Get-ConfigValue $member 'userId'
                if (-not $id -or $id -in $wanted.Keys) { continue }
                # An owner keeps the channel manageable; removing the last one orphans it.
                if ('owner' -in @(Get-ConfigValue $member 'roles' @())) {
                    Write-Skip "$(Get-ConfigValue $member 'displayName' $id) is a channel owner - left alone"
                    continue
                }
                $who = Get-ConfigValue $member 'displayName' $id
                if (-not $PSCmdlet.ShouldProcess("$($entry.title) / $who", 'Remove from private channel')) { continue }

                try {
                    Invoke-Graph -Url "v1.0/teams/$teamId/channels/$($channel.id)/members/$($member.id)" -Method DELETE | Out-Null
                    Write-Change "$who removed from '$($entry.title)'"
                    $report.Add([PSCustomObject]@{ Channel = $entry.title; Action = 'removed'; User = $who })
                    $removed++
                } catch {
                    Write-Bad "could not remove $who : $($_.Exception.Message)"
                    $warned++
                }
            }
        }
    }

    if ($report.Count -gt 0) {
        if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }
        $report | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    }

    Write-Head 'Summary'
    Write-Host "    Added   : $added" -ForegroundColor Cyan
    Write-Host "    Removed : $removed$(if (-not $Prune) { '  (-Prune removes people the groups no longer list)' })" -ForegroundColor Cyan
    if ($warned -gt 0) { Write-Warn "$warned thing(s) need attention - scroll back" }
    if ($report.Count -gt 0) { Write-Host "    Report  : $ReportPath" -ForegroundColor Cyan }
    if ($added -eq 0 -and $removed -eq 0) { Write-Ok 'Channel membership already matches the groups.' }
    Write-Host ''
} catch {
    Write-Host ''
    Write-Bad "Aborted: $($_.Exception.Message)"
    Write-Host ''
    exit 1
}

exit 0
