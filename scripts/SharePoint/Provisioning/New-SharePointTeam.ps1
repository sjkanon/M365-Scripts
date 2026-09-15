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
    Requires: PnP.PowerShell 2.x
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
$tenantName = ($Tenant -split '\.')[0]
$rootUrl    = "https://$tenantName.sharepoint.com"

$discovered = @{}
$changeCount = 0

Write-Host ''
Write-Host "  Client : $(Get-ConfigValue $config 'client' '(unnamed)')" -ForegroundColor Cyan
Write-Host "  Team   : $($team.displayName)  (alias $($team.mailNickname))" -ForegroundColor Cyan
Write-Host "  Owner  : $(if ($Owner) { $Owner } else { 'the signed-in account' })" -ForegroundColor Cyan
Write-Host "  Mode   : $(if ($simulate) { '-WhatIf - nothing will be created' } else { 'APPLY' })" `
    -ForegroundColor $(if ($simulate) { 'Yellow' } else { 'Cyan' })

function Invoke-Graph {
    <# Graph through the PnP connection, so the whole run stays on one app. #>
    param([Parameter(Mandatory)] [string] $Url, [string] $Method = 'GET', $Body, $Connection)

    $splat = @{ Url = $Url; Method = $Method; Connection = $Connection; ErrorAction = 'Stop' }
    if ($Body) { $splat['Content'] = ($Body | ConvertTo-Json -Depth 8) }
    return Invoke-PnPGraphMethod @splat
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
        [Parameter(Mandatory)] [string] $ChannelName,
        [Parameter(Mandatory)] $Connection
    )

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $attempt  = 0
    while ((Get-Date) -lt $deadline) {
        $attempt++
        try {
            $folder = Invoke-Graph -Url "v1.0/teams/$TeamId/channels/$ChannelId/filesFolder" -Connection $Connection
            $webUrl = Get-ConfigValue $folder 'webUrl'
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

    # Any site will do as a connection for the Graph calls; the tenant root always
    # exists, and the team site does not yet.
    $connection = Connect-Structure -Url $rootUrl -Tenant $Tenant -ClientId $ClientId -Interactive:$Interactive

    $existing = $null
    try {
        $escaped  = $team.mailNickname -replace "'", "''"
        $found    = Invoke-Graph -Url "v1.0/groups?`$filter=mailNickname eq '$escaped'" -Connection $connection
        $existing = @(Get-ConfigValue $found 'value' @()) | Select-Object -First 1
    } catch {
        throw "Could not query Microsoft 365 groups: $($_.Exception.Message)"
    }

    $teamId = $null
    if ($existing) {
        $teamId = $existing.id
        Write-Ok "Team '$($existing.displayName)' already exists ($teamId)"
        if ($existing.displayName -ne $team.displayName) {
            Write-Warn "Its display name is '$($existing.displayName)', the config says '$($team.displayName)' - left as is, renaming a team is yours to decide."
        }
    } elseif ($PSCmdlet.ShouldProcess($team.displayName, 'Create Microsoft 365 team')) {
        $newTeam = New-PnPTeamsTeam -DisplayName $team.displayName `
            -MailNickName $team.mailNickname `
            -Description (Get-ConfigValue $team 'description' '') `
            -Visibility (Get-ConfigValue $team 'visibility' 'Private') `
            -Owners $(if ($Owner) { @($Owner) } else { @() }) `
            -Connection $connection
        $teamId = $newTeam.GroupId
        Write-Change "Team '$($team.displayName)' created ($teamId)"
        $changeCount++
        Write-Step 'Waiting 30s for the team and its site to settle...'
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
    $site = Invoke-Graph -Url "v1.0/groups/$teamId/sites/root" -Connection $connection
    $teamSiteUrl = (Get-ConfigValue $site 'webUrl').TrimEnd('/')
    $discovered['team'] = $teamSiteUrl
    Write-Ok "Team site: $teamSiteUrl"

    # -- 3. Channels -----------------------------------------------------------
    Write-Head '2. Channels'
    $channels = @(Invoke-Graph -Url "v1.0/teams/$teamId/channels" -Connection $connection |
                  ForEach-Object { Get-ConfigValue $_ 'value' @() })

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
            $addSplat = @{
                Team        = $teamId
                DisplayName = $entry.title
                Connection  = $connection
            }
            if ($channelType -eq 'Private') {
                if (-not $Owner) { throw "A private channel needs an owner. Pass -Owner, or fill in team.owners." }
                $addSplat['ChannelType'] = 'Private'
                $addSplat['OwnerUPN']    = $Owner
            }
            $live = Add-PnPTeamsChannel @addSplat
            Write-Change "$($channelType.ToLower()) channel '$($entry.title)' created"
            $changeCount++
        }
        if (-not $live) { continue }

        # A private channel has its own site collection, and only Graph knows its URL.
        if ($channelType -eq 'Private') {
            $siteKey = Get-ConfigValue $entry 'site'
            $url     = Wait-ChannelSite -TeamId $teamId -ChannelId $live.id -ChannelName $entry.title -Connection $connection
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
