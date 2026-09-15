#Requires -Version 7.0
<#
.SYNOPSIS
    Create the Microsoft 365 Team and its channels - including the private MGMT
    channel - from the structure configuration, and write the resulting site URLs
    back into it. Idempotent, supports -WhatIf.

.DESCRIPTION
    The step before all the others. Everything in this folder assumes a team with a
    channel per pillar already exists; this is what makes that true, so the whole
    structure can be built from an empty tenant.

        1. Team     the Microsoft 365 group and its team, named from the team section
                    of the configuration. The mailNickname decides the site URL:
                    .../sites/<mailNickname>.
        2. Channels one per container with a channelType, standard or private, named
                    from the container's title.
        3. Sites    the team site URL, and the private channel's own site URL, read
                    back from Graph and written into the configuration's sites
                    section - that URL cannot be known in advance, SharePoint makes
                    it up from the team and channel name.
        4. Folders  opening a channel's Files tab is what makes SharePoint create the
                    channel folder. This does it from here instead, so the next step
                    does not fail on a folder nobody has clicked on yet.

    Naming
    ------
    Everything this creates takes its name from the configuration. Nothing is derived
    from anything else, so a second client is a second config file:

        team.displayName        the team as it appears in Teams
        team.mailNickname       the mailbox alias AND the site URL
        containers[].title      the channel name
        containers[].folder     the folder inside the channel's library
        containers[].list       the library the folder lives in

    Private channel sites take a moment
    -----------------------------------
    SharePoint provisions a private channel's site collection asynchronously, and the
    URL does not exist the instant the channel does. This polls for it - a couple of
    minutes is normal, five is not unusual on a busy tenant. Without that wait the
    metadata step would connect to a site that is not there yet.

    What it will not do
    -------------------
    It never deletes a channel and never renames an existing one. A channel whose name
    does not match the configuration is reported, not corrected: renaming a channel
    moves its folder and breaks every link anyone has shared.

.PARAMETER ConfigPath
    Path to the structure configuration JSON.
    Default: petsolutions.config.json next to this script.

.PARAMETER Owner
    UPN of the team owner, and the owner of the private channel. Defaults to the first
    entry in the configuration's team.owners, or the signed-in account.

.PARAMETER SkipConfigUpdate
    Do not write the discovered site URLs back into the configuration file. The run
    then prints them instead, for you to paste in yourself.

.PARAMETER TimeoutMinutes
    How long to wait for a private channel's site collection to appear (default 10).

.PARAMETER Interactive
    Sign in interactively (the default - creating a team app-only needs an owner and
    is rarely what you want).

.PARAMETER ClientId
    Client ID of the Entra app registration used to sign in.

.PARAMETER Tenant
    Tenant name or ID. Defaults to the tenant in the configuration file.

.PARAMETER Disconnect
    Sign out of PnP when finished.

.EXAMPLE
    # What would be created, and under which names?
    .\New-SharePointTeam.ps1 -Interactive -ClientId <app-id> -WhatIf

.EXAMPLE
    # Create the team, the six channels and the private MGMT channel
    .\New-SharePointTeam.ps1 -Interactive -ClientId <app-id> -Owner jan@petsolutions.be

.NOTES
    Author  : Sjoerd Kanon
    Requires: PnP.PowerShell 2.x and Microsoft.Graph.Authentication
    Rights  : Teams administrator or a user allowed to create teams, and the app needs
              the delegated Group.ReadWrite.All and Channel.Create scopes.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $ConfigPath,
    [string] $Owner,
    [switch] $SkipConfigUpdate,

    [ValidateRange(1, 60)]
    [int] $TimeoutMinutes = 10,

    [switch] $Interactive,
    [string] $ClientId,
    [string] $Tenant,
    [switch] $Disconnect
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

. (Join-Path $PSScriptRoot 'SharePointStructure.Common.ps1')

trap {
    Write-Host ''
    Write-Host "  FAILED at line $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    Write-Host "    $($_.Exception.GetType().Name): $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ''
    break
}

Assert-PnPModule

if (-not $ConfigPath) { $ConfigPath = Join-Path $PSScriptRoot 'petsolutions.config.json' }
# The site URLs are what this script discovers, so placeholders there are expected.
$config = Import-StructureConfig -Path $ConfigPath -AllowPlaceholders
if (-not $Tenant) { $Tenant = $config.tenant }
if ($Tenant -match 'CHANGEME') { throw "Fill in the tenant in $ConfigPath first." }

$team = Get-ConfigValue $config 'team'
if (-not $team) { throw "The configuration has no team section - nothing to create ($ConfigPath)." }

if (-not $Owner) {
    $Owner = @(Get-ConfigValue $team 'owners' @()) | Where-Object { $_ -notmatch 'CHANGEME' } | Select-Object -First 1
}

$simulate  = [bool] $WhatIfPreference

$discovered = @{}
$changeCount = 0

Write-Host ''
Write-Host "  Client : $(Get-ConfigValue $config 'client' '(unnamed)')" -ForegroundColor Cyan
Write-Host "  Team   : $($team.displayName)  (alias $($team.mailNickname))" -ForegroundColor Cyan
Write-Host "  Owner  : $(if ($Owner) { $Owner } else { 'the signed-in account' })" -ForegroundColor Cyan
Write-Host "  Mode   : $(if ($simulate) { '-WhatIf - nothing will be created' } else { 'APPLY' })" `
    -ForegroundColor $(if ($simulate) { 'Yellow' } else { 'Cyan' })

function Invoke-Graph {
    <#
        A Graph call on the Graph sign-in this script makes itself.

        Deliberately not routed through the PnP connection: that would use whichever
        app registration PnP happens to be signed in with, and an app cached by
        another script in this repo carries the SharePoint scopes but not
        Group.ReadWrite.All or Channel.Create. The call then comes back 403 halfway
        through creating a team, which is a miserable place to find out.
    #>
    param([Parameter(Mandatory)] [string] $Url, [string] $Method = 'GET', $Body)

    $splat = @{ Uri = "https://graph.microsoft.com/$Url"; Method = $Method; ErrorAction = 'Stop' }
    if ($Body) {
        $splat['Body']        = ($Body | ConvertTo-Json -Depth 8)
        $splat['ContentType'] = 'application/json'
    }

    try {
        return Invoke-MgGraphRequest @splat
    } catch {
        # "BadRequest" on its own is useless. Graph puts the reason in the response
        # body, and that is the difference between guessing and knowing.
        $detail = ''
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            try {
                $parsed = $_.ErrorDetails.Message | ConvertFrom-Json
                if ($parsed.error) { $detail = "$($parsed.error.code): $($parsed.error.message)" }
            } catch {
                $detail = $_.ErrorDetails.Message
            }
        }
        if (-not $detail) { $detail = $_.Exception.Message }
        throw "$Method $Url -> $detail"
    }
}

function Enable-Team {
    <#
        Turn a Microsoft 365 group into a team, retrying while it replicates.

        PUT /groups/{id}/team, not POST /teams: the POST form binds a group *and* a
        template and is fussy about both, while the PUT is the documented way to
        team-enable a group that already exists. A group created seconds ago answers
        404, and sometimes 400, for a minute or two.
    #>
    param([Parameter(Mandatory)] [string] $GroupId, [int] $Attempts = 12)

    $lastError = ''
    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            Invoke-Graph -Url "v1.0/groups/$GroupId/team" -Method PUT -Body @{
                memberSettings    = @{ allowCreatePrivateChannels = $true; allowCreateUpdateChannels = $true }
                messagingSettings = @{ allowUserEditMessages = $true; allowUserDeleteMessages = $true }
            } | Out-Null
            return
        } catch {
            $lastError = $_.Exception.Message
            if ($attempt -eq $Attempts) { break }
            Write-Skip "not ready yet (attempt $attempt): $lastError"
            Start-Sleep -Seconds 15
        }
    }
    throw ("The group could not be turned into a team. Last answer from Graph:`n    $lastError`n" +
           "    The group exists as $GroupId - rerun this script and it picks up from there, " +
           'or team-enable it once by hand in the Teams admin centre.')
}

function Test-IsTeam {
    <# Whether a group has already been team-enabled. #>
    param([Parameter(Mandatory)] [string] $GroupId)

    try {
        Invoke-Graph -Url "v1.0/teams/$GroupId" | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Wait-ChannelSite {
    <#
        A private channel's site collection is provisioned asynchronously - the channel
        exists before its site does. Poll the channel's files folder until Graph hands
        back a URL, then derive the site collection from it.
    #>
    param(
        [Parameter(Mandatory)] [string] $TeamId,
        [Parameter(Mandatory)] [string] $ChannelId,
        [Parameter(Mandatory)] [string] $ChannelName
    )

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $attempt  = 0
    while ((Get-Date) -lt $deadline) {
        $attempt++
        try {
            $folder = Invoke-Graph -Url "v1.0/teams/$TeamId/channels/$ChannelId/filesFolder"
            $webUrl = [string] $folder.webUrl
            if ($webUrl) {
                # .../sites/Petsolutions-MGMT/Shared%20Documents -> .../sites/Petsolutions-MGMT
                if ($webUrl -match '^(https://[^/]+/sites/[^/]+)') { return $Matches[1] }
                return $webUrl
            }
        } catch {
            # 404 while the site is still being created is the normal case here.
        }
        Write-Skip "waiting for the site behind '$ChannelName' (attempt $attempt)..."
        Start-Sleep -Seconds 20
    }
    throw "The site for channel '$ChannelName' did not appear within $TimeoutMinutes minutes. Rerun this script - it will pick up where it left off."
}

try {
    # -- 1. The team -----------------------------------------------------------
    Write-Head '1. Microsoft 365 group and team'

    # Sign in to Graph for the team and channel work, asking for exactly the scopes
    # this needs. Creating a team and a private channel is Graph, not SharePoint.
    Connect-StructureGraph -Tenant $Tenant -AuthenticationOnly -Scopes @(
        'Group.ReadWrite.All'
        'Directory.Read.All'
        'Team.Create'
        'Channel.Create'
        'ChannelSettings.ReadWrite.All'
        'Sites.Read.All'
    )

    $existing = $null
    try {
        $escaped  = $team.mailNickname -replace "'", "''"
        $found    = Invoke-Graph -Url "v1.0/groups?`$filter=mailNickname eq '$escaped'"
        $existing = @($found.value) | Select-Object -First 1
    } catch {
        throw "Could not query Microsoft 365 groups: $($_.Exception.Message)"
    }

    $teamId = $null
    if ($existing) {
        $teamId = $existing.id
        if ($existing.displayName -ne $team.displayName) {
            Write-Warn "Its display name is '$($existing.displayName)', the config says '$($team.displayName)' - left as is, renaming a team is yours to decide."
        }

        # A group can exist without being a team - a run that got this far and then
        # failed leaves exactly that behind, and it has to be finishable.
        if (Test-IsTeam -GroupId $teamId) {
            Write-Ok "Team '$($existing.displayName)' already exists ($teamId)"
        } elseif ($PSCmdlet.ShouldProcess($existing.displayName, 'Turn the existing group into a team')) {
            Write-Warn "Group '$($existing.displayName)' exists but is not a team yet - finishing that now."
            Enable-Team -GroupId $teamId
            Write-Change "Group '$($existing.displayName)' is now a team"
            $changeCount++
            Write-Step 'Waiting 30s for the team site to settle...'
            Start-Sleep -Seconds 30
        }
    } elseif ($PSCmdlet.ShouldProcess($team.displayName, 'Create Microsoft 365 team')) {
        # Two steps on purpose: the group first, then team-enable it. Creating a team
        # straight from a template gives no control over the mailNickname, and that
        # nickname is what decides the site URL the rest of this set connects to.
        if (-not $Owner) { throw 'A new team needs an owner. Pass -Owner, or fill in team.owners in the configuration.' }

        $ownerObject = Invoke-Graph -Url "v1.0/users/$([uri]::EscapeDataString($Owner))"
        $group = Invoke-Graph -Url 'v1.0/groups' -Method POST -Body @{
            displayName          = $team.displayName
            mailNickname         = $team.mailNickname
            description          = (Get-ConfigValue $team 'description' '')
            visibility           = (Get-ConfigValue $team 'visibility' 'Private')
            groupTypes           = @('Unified')
            mailEnabled          = $true
            securityEnabled      = $false
            'owners@odata.bind'  = @("https://graph.microsoft.com/v1.0/users/$($ownerObject.id)")
            'members@odata.bind' = @("https://graph.microsoft.com/v1.0/users/$($ownerObject.id)")
        }
        $teamId = $group.id
        Write-Change "Microsoft 365 group '$($team.displayName)' created ($teamId)"

        # The group has to exist everywhere before it can be team-enabled; a fresh one
        # answers 404, and sometimes 400, for a minute or two.
        #
        # PUT /groups/{id}/team, not POST /teams: the POST form binds a group *and* a
        # template and is fussy about both, while the PUT is the documented way to
        # team-enable a group that already exists.
        Write-Step 'Waiting for the group to replicate before turning it into a team...'
        Start-Sleep -Seconds 15
        Enable-Team -GroupId $teamId
        Write-Change "Team '$($team.displayName)' created"
        $changeCount++
        Write-Step 'Waiting 30s for the team site to settle...'
        Start-Sleep -Seconds 30
    }

    if (-not $teamId) {
        Write-Skip 'No team to work with (-WhatIf) - the channels below are reported from the configuration only.'
        foreach ($entry in ($config.containers | Where-Object { Get-ConfigValue $_ 'channelType' })) {
            Write-Change "would create $((Get-ConfigValue $entry 'channelType').ToLower()) channel '$($entry.title)'"
        }
        Write-Host ''
        exit 0
    }

    # -- 2. The team site URL --------------------------------------------------
    $site = Invoke-Graph -Url "v1.0/groups/$teamId/sites/root"
    $teamSiteUrl = ([string] $site.webUrl).TrimEnd('/')
    $discovered['team'] = $teamSiteUrl
    Write-Ok "Team site: $teamSiteUrl"

    # -- 3. Channels -----------------------------------------------------------
    Write-Head '2. Channels'
    $channels = @((Invoke-Graph -Url "v1.0/teams/$teamId/channels").value)

    foreach ($entry in $config.containers) {
        $channelType = Get-ConfigValue $entry 'channelType'
        if (-not $channelType) {
            Write-Skip "'$($entry.title)' is a library, not a channel - handled by Set-SharePointLibraries.ps1"
            continue
        }

        $live = $channels | Where-Object { $_.displayName -eq $entry.title } | Select-Object -First 1
        if ($live) {
            Write-Ok "channel '$($entry.title)' ($($live.membershipType))"
        } elseif ($PSCmdlet.ShouldProcess($entry.title, "Create $channelType channel")) {
            $body = @{
                displayName    = $entry.title
                description    = "Pijler $($entry.key)"
                membershipType = $(if ($channelType -eq 'Private') { 'private' } else { 'standard' })
            }
            if ($channelType -eq 'Private') {
                if (-not $Owner) { throw 'A private channel needs an owner. Pass -Owner, or fill in team.owners.' }
                # A private channel is created with its owner in one call - Graph
                # refuses to create one with no members at all.
                $ownerObject = Invoke-Graph -Url "v1.0/users/$([uri]::EscapeDataString($Owner))"
                $body['members'] = @(@{
                    '@odata.type'     = '#microsoft.graph.aadUserConversationMember'
                    'user@odata.bind' = "https://graph.microsoft.com/v1.0/users('$($ownerObject.id)')"
                    roles             = @('owner')
                })
            }
            $live = Invoke-Graph -Url "v1.0/teams/$teamId/channels" -Method POST -Body $body
            Write-Change "$($channelType.ToLower()) channel '$($entry.title)' created"
            $changeCount++
        }
        if (-not $live) { continue }

        # A private channel has its own site collection, and only Graph knows its URL.
        if ($channelType -eq 'Private') {
            $siteKey = Get-ConfigValue $entry 'site'
            $url     = Wait-ChannelSite -TeamId $teamId -ChannelId $live.id -ChannelName $entry.title
            $discovered[$siteKey] = $url
            Write-Ok "site for '$($entry.title)': $url"
        }
    }

    # -- 4. Channel folders ----------------------------------------------------
    # SharePoint creates a standard channel's folder the first time somebody opens the
    # Files tab. Doing it here means the next step does not fail on a folder nobody
    # has clicked on yet.
    Write-Head '3. Channel folders'
    $teamConnection = Connect-Structure -Url $teamSiteUrl -Tenant $Tenant -ClientId $ClientId -Interactive:$Interactive
    foreach ($entry in ($config.containers | Where-Object { (Get-ConfigValue $_ 'channelType') -eq 'Standard' })) {
        $listTitle  = Get-ConfigValue $entry 'list' 'Documents'
        $folderName = Get-ConfigValue $entry 'folder' $entry.title

        $list = Get-PnPList -Identity $listTitle -Connection $teamConnection -ErrorAction SilentlyContinue
        if (-not $list) { Write-Warn "library '$listTitle' not found yet - rerun in a minute"; continue }

        $context = Get-PnPContext -Connection $teamConnection
        $context.Load($list.RootFolder)
        $context.ExecuteQuery()
        $folderUrl = "$($list.RootFolder.ServerRelativeUrl.TrimEnd('/'))/$folderName"

        if (Get-PnPFolder -Url $folderUrl -Connection $teamConnection -ErrorAction SilentlyContinue) {
            Write-Ok "folder '$folderName'"
        } elseif ($PSCmdlet.ShouldProcess($folderUrl, 'Create channel folder')) {
            Add-PnPFolder -Name $folderName -Folder $list.RootFolder.ServerRelativeUrl -Connection $teamConnection | Out-Null
            Write-Change "folder '$folderName' created"
            $changeCount++
        }
    }

    # -- 5. Write the URLs back ------------------------------------------------
    Write-Head '4. Site URLs'
    foreach ($key in $discovered.Keys) { Write-Ok "$key -> $($discovered[$key])" }

    if ($SkipConfigUpdate) {
        Write-Skip 'Configuration not updated (-SkipConfigUpdate) - paste the URLs above into the sites section yourself.'
    } elseif ($simulate) {
        Write-Skip 'Configuration not updated (-WhatIf).'
    } else {
        $raw = Get-Content -Path $ConfigPath -Raw -Encoding UTF8
        $json = $raw | ConvertFrom-Json
        $changed = $false
        foreach ($key in $discovered.Keys) {
            if ($json.sites.PSObject.Properties.Name -notcontains $key) { continue }
            if ($json.sites.$key -eq $discovered[$key]) { continue }
            # Replace the value textually so the file keeps its formatting and comments
            # stay where the author put them.
            $raw = $raw.Replace("`"$($json.sites.$key)`"", "`"$($discovered[$key])`"")
            $changed = $true
        }
        if ($changed) {
            Set-Content -Path $ConfigPath -Value $raw -NoNewline -Encoding UTF8
            Write-Change "Site URLs written back to $ConfigPath"
            $changeCount++
        } else {
            Write-Ok 'Configuration already had the right site URLs.'
        }
    }

    Write-Head 'Summary'
    if ($simulate) {
        Write-Host "    $changeCount change(s) would be made - rerun without -WhatIf." -ForegroundColor Yellow
    } elseif ($changeCount -eq 0) {
        Write-Ok 'Team and channels already match the configuration.'
    } else {
        Write-Change "$changeCount change(s) applied."
        Write-Host '    Next: New-SharePointMetadata.ps1, or Install-SharePointStructure.ps1 for the rest in one go.' -ForegroundColor DarkGray
    }
    Write-Host ''
} finally {
    if ($Disconnect) { Disconnect-Structure }
}
