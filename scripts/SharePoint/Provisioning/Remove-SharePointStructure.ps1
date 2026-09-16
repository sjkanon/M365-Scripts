#Requires -Version 7.0
<#
.SYNOPSIS
    Remove what the structure scripts created, in the order that works. Reports only
    unless you pass -Apply.

.DESCRIPTION
    The undo for this folder: a botched run, a test tenant, a client that never went
    live. It walks the configuration and removes what it describes, deepest first, so
    nothing is left holding on to something else.

        Tabs          the library tabs added to channels
        Channels      the channels named by the configuration - and with them the
                      files in their folders
        Libraries     the libraries a container owns (kind: Library)
        ContentTypes  unbound from the lists that use them, then removed
        Columns       the site columns, on every site in the configuration
        TermSet       the term set, the term group and the terms in it
        Groups        the Entra ID security groups
        Team          the Microsoft 365 group behind the team - which takes the site,
                      every library, every file and every chat message with it

    Deliberately the reverse default
    --------------------------------
    Without -Apply this changes nothing at all: it prints what it would remove and
    what it would refuse to. Forgetting -WhatIf on a destructive script is the
    dangerous direction, so the safe state is the one you get by default.

    What it refuses to do
    ---------------------
      - A library or channel folder that still holds files is skipped unless you pass
        -IncludeContent. The count is reported either way.
      - The General channel is never removed; Teams does not allow it and neither
        does this.
      - The team's own Documents library is never removed.
      - A content type still in use somewhere is reported, not forced.
      - -Scope Team asks you to type the team's name. Nothing else in this repo does
        that, and nothing else in this repo deletes a client's files.

    What it cannot undo
    -------------------
    Removing the term set orphans the Leverancier value on every document that carried
    one - the field keeps a GUID that no longer resolves. Removing a column takes its
    data with it. Neither is recoverable from the recycle bin, so both are reported
    with what they will cost before they run.

    A deleted Microsoft 365 group is soft-deleted for 30 days and can be restored in
    the Entra portal; a deleted channel sits in its own 30-day recycle. Files in a
    removed library go to the site recycle bin.

.PARAMETER ConfigPath
    Path to the structure configuration JSON. Default: the single *.config.json next
    to this script that no longer contains the CHANGEME placeholders.

.PARAMETER Scope
    What to remove. Repeatable. Default: everything except Team.
    Tabs, Channels, Libraries, ContentTypes, Columns, TermSet, Groups, Team, All.
    All still excludes Team - that one is only ever removed by naming it.

.PARAMETER Apply
    Actually remove. Without it nothing is changed.

.PARAMETER IncludeContent
    Also remove libraries and channel folders that still hold files.

.PARAMETER Interactive
    Sign in interactively (the default).

.PARAMETER ClientId
    Client ID of the Entra app registration used for the SharePoint work.

.PARAMETER Tenant
    Tenant name or ID. Defaults to the tenant in the configuration file.

.PARAMETER ReportPath
    CSV of everything removed, skipped or refused.
    Default: C:\Temp\SharePointStructure_Cleanup_<timestamp>.csv

.EXAMPLE
    # What would go? Changes nothing.
    .\Remove-SharePointStructure.ps1

.EXAMPLE
    # Undo a botched run, leaving the team and its files alone
    .\Remove-SharePointStructure.ps1 -Scope Channels,ContentTypes,Columns,TermSet,Groups -Apply

.EXAMPLE
    # Start over on a test tenant: everything, files included, team and all
    .\Remove-SharePointStructure.ps1 -Scope All,Team -IncludeContent -Apply

.NOTES
    Author  : Sjoerd Kanon
    Requires: PnP.PowerShell 2.x, Microsoft.Graph.Authentication
#>
[CmdletBinding()]
param(
    [string] $ConfigPath,

    [ValidateSet('All', 'Tabs', 'Channels', 'Libraries', 'ContentTypes', 'Columns', 'TermSet', 'Groups', 'Team')]
    [string[]] $Scope = @('All'),

    [switch] $Apply,
    [switch] $IncludeContent,

    [switch] $Interactive,
    [string] $ClientId,
    [string] $Tenant,
    [string] $ReportPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

. (Join-Path $PSScriptRoot 'SharePointStructure.Common.ps1')

Assert-PnPModule

if (-not $ConfigPath) {
    $candidates = @(Get-ChildItem -Path $PSScriptRoot -Filter '*.config.json' -ErrorAction SilentlyContinue |
                    Where-Object { (Get-Content $_.FullName -Raw) -notmatch 'CHANGEME' })
    if ($candidates.Count -eq 1) { $ConfigPath = $candidates[0].FullName }
    elseif ($candidates.Count -eq 0) { throw 'No configuration found - nothing to clean up.' }
    else { throw ("More than one configuration here - pass -ConfigPath. Found: {0}" -f (($candidates | ForEach-Object { $_.Name }) -join ', ')) }
}

$config = Import-StructureConfig -Path $ConfigPath -AllowPlaceholders
if (-not $Tenant) { $Tenant = $config.tenant }

# The app registration this set provisions is cached per tenant in pnp.appid.json, and
# Install-SharePointStructure puts it there. Cleaning up is the same job in reverse
# against the same sites, so asking for -ClientId again - and failing every site
# connect when it is left out - made the removal harder to run than the install.
#
# Reused, not created: if there is no cached app there is nothing this script can do
# about it, and the message already says what to pass.
if (-not $ClientId) {
    $ClientId = Get-CachedStructureClientId -Tenant $Tenant
    if ($ClientId) {
        Write-Ok "Using the cached app registration for ${Tenant}: $ClientId"
        # A cached app plus no certificate means the sign-in is a browser one; saying
        # so beats Connect-Structure refusing with a message about flags.
        if (-not $Interactive) { $Interactive = [switch]::Present }
    }
}

# 'All' is everything except the team: deleting a client's whole team is never
# something you should get by asking for "all".
$allScopes = @('Tabs', 'Channels', 'Libraries', 'ContentTypes', 'Columns', 'TermSet', 'Groups')
$wanted    = @()
foreach ($s in $Scope) { if ($s -eq 'All') { $wanted += $allScopes } else { $wanted += $s } }
$wanted = @($wanted | Select-Object -Unique)

$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not $ReportPath) {
    $ReportPath = Join-Path $outputDir "SharePointStructure_Cleanup_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
}

$report   = [System.Collections.Generic.List[object]]::new()
$removed  = 0
$skipped  = 0
$failed   = 0
$script:graphReady = $false

function Add-Row {
    param([string] $What, [string] $Name, [string] $Result, [string] $Detail = '')

    $report.Add([PSCustomObject]@{ What = $What; Name = $Name; Result = $Result; Detail = $Detail })
    switch ($Result) {
        'removed' { Write-Change "$What '$Name'$(if ($Detail) { " - $Detail" })"; $script:removed++ }
        'would'   { Write-Change "would remove $What '$Name'$(if ($Detail) { " - $Detail" })" }
        'skipped' { Write-Skip  "$What '$Name' - $Detail"; $script:skipped++ }
        'failed'  { Write-Bad   "$What '$Name' - $Detail"; $script:failed++ }
        default   { Write-Ok    "$What '$Name' - $Detail" }
    }
}

function Connect-Graph {
    if ($script:graphReady) { return }
    Connect-StructureGraph -Tenant $Tenant -ClientId $ClientId -AuthenticationOnly -Scopes @(
        'Group.ReadWrite.All', 'Directory.Read.All', 'Channel.ReadBasic.All'
        'ChannelSettings.ReadWrite.All', 'TeamsTab.ReadWrite.All'
    )
    $script:graphReady = $true
}


function Get-TeamId {
    <#
        The team's identity, resolved once, as plain strings - or $null when there is
        no team to remove anything from.

        Handing back the raw Graph object was the bug this replaces. Six call sites
        each dug the id out again with Get-ConfigValue on the way into a URL, so a
        response that did not have the shape those calls assumed produced
        "v1.0/teams//channels" - which Graph rejects with "teamId needs to be a valid
        GUID", blamed on whichever step happened to ask first rather than on the
        lookup that came back short. Resolved once and validated here, a URL can only
        ever be built from something already known to be a GUID.
    #>
    Connect-Graph
    $team = Get-ConfigValue $config 'team'
    if (-not $team) { return $null }
    $escaped = $team.mailNickname -replace "'", "''"
    $found   = Invoke-StructureGraph -Url "v1.0/groups?`$filter=mailNickname eq '$escaped'&`$select=id,displayName"
    $group   = @(Get-ConfigValue $found 'value' @()) | Select-Object -First 1
    if (-not $group) {
        Write-Warn "no group has mailNickname '$($team.mailNickname)' - Graph answered $(Get-ValueShape $found)"
        return $null
    }

    $id   = [string] (Get-ConfigValue $group 'id')
    $guid = [guid]::Empty
    if (-not [guid]::TryParse($id, [ref] $guid)) {
        # Found something, but nothing that can be removed by id. Said out loud rather
        # than skipped quietly: "no team found" would be a different answer, and both
        # the alias and the shape that produced this are only in hand right here.
        Write-Warn "group '$($team.mailNickname)' has no usable id - it came back as $(Get-ValueShape $group)"
        return $null
    }

    return [pscustomobject]@{
        Id          = $id
        DisplayName = [string] (Get-ConfigValue $group 'displayName' $team.mailNickname)
    }
}


function Resolve-Team {
    <#
        Get-TeamId's answer, and nothing else that may have reached the output stream
        on the way here.

        Three runs failed on a $teamGroup that was truthy and had no Id, which none of
        Get-TeamId's return paths can produce - every one of them is either $null or an
        object that has one. That only happens if something besides the return value
        arrived as well. Rather than guess which line wrote it, the lookup is read as
        what it is - a stream - the one usable object is taken from it, and anything
        else is named rather than silently used as if it were a team.
    #>
    $team = $null
    foreach ($item in @(Get-TeamId)) {
        if ($null -eq $item) { continue }
        if (-not $team -and (Get-ConfigValue $item 'Id')) { $team = $item; continue }

        # The Graph SDK's welcome banner arrives here as a bare string - despite
        # -NoWelcome on every Connect-MgGraph, and despite the sign-in helper now
        # discarding its own output stream. It is written outside any pipeline this
        # code controls, so it is caught rather than prevented. A string can never be
        # a team, so it goes quietly; anything else is worth saying out loud, because
        # anything else would mean something new is wrong.
        if ($item -is [string]) { continue }
        Write-Warn "the team lookup also wrote $(Get-ValueShape $item) to its output - ignored"
    }
    return $team
}

Write-Host ''
Write-Host '  ┌──────────────────────────────────────────────────────────────┐' -ForegroundColor $(if ($Apply) { 'Red' } else { 'Cyan' })
Write-Host '  │  Remove the SharePoint structure                              │' -ForegroundColor $(if ($Apply) { 'Red' } else { 'Cyan' })
Write-Host '  └──────────────────────────────────────────────────────────────┘' -ForegroundColor $(if ($Apply) { 'Red' } else { 'Cyan' })
Write-Host "  Client : $(Get-ConfigValue $config 'client' '(unnamed)')" -ForegroundColor Cyan
Write-Host "  Config : $ConfigPath" -ForegroundColor Cyan
Write-Host "  Scope  : $($wanted -join ', ')" -ForegroundColor Cyan
Write-Host "  Mode   : $(if ($Apply) { 'APPLY - things will be removed' } else { 'REPORT ONLY - pass -Apply to actually remove' })" `
    -ForegroundColor $(if ($Apply) { 'Red' } else { 'Yellow' })
if ($Apply -and -not $IncludeContent) {
    Write-Host '  Note   : libraries and folders that still hold files are skipped (-IncludeContent overrides)' -ForegroundColor DarkGray
}

try {
    # -- Tabs ------------------------------------------------------------------
    if ('Tabs' -in $wanted -or 'Channels' -in $wanted) {
        Write-Head 'Tabs'
        $teamGroup = Resolve-Team
        if (-not $teamGroup) {
            Write-Skip 'no team found - nothing to do'
        } else {
            $channels = @(Get-StructureGraphCollection -Url "v1.0/teams/$($teamGroup.Id)/channels")
            foreach ($entry in $config.containers) {
                $tab = Get-ConfigValue $entry 'tab'
                if (-not $tab) { continue }
                $channel = $channels | Where-Object { (Get-ConfigValue $_ 'displayName') -eq $entry.title } | Select-Object -First 1
                if (-not $channel) { continue }

                $tabName = Get-ConfigValue $tab 'name' $entry.title
                $live = @(Get-StructureGraphCollection -Url "v1.0/teams/$($teamGroup.Id)/channels/$(Get-ConfigValue $channel 'id')/tabs") |
                        Where-Object { (Get-ConfigValue $_ 'displayName') -eq $tabName } | Select-Object -First 1
                if (-not $live) { continue }

                if (-not $Apply) { Add-Row -What 'tab' -Name "$($entry.title)/$tabName" -Result 'would'; continue }
                try {
                    Invoke-StructureGraph -Url "v1.0/teams/$($teamGroup.Id)/channels/$(Get-ConfigValue $channel 'id')/tabs/$(Get-ConfigValue $live 'id')" -Method DELETE | Out-Null
                    Add-Row -What 'tab' -Name "$($entry.title)/$tabName" -Result 'removed'
                } catch {
                    Add-Row -What 'tab' -Name "$($entry.title)/$tabName" -Result 'failed' -Detail $_.Exception.Message
                }
            }
        }
    }

    # -- Channels --------------------------------------------------------------
    if ('Channels' -in $wanted) {
        Write-Head 'Channels'
        $teamGroup = Resolve-Team
        if (-not $teamGroup) {
            Write-Skip 'no team found - nothing to do'
        } else {
            $channels = @(Get-StructureGraphCollection -Url "v1.0/teams/$($teamGroup.Id)/channels")
            foreach ($entry in $config.containers) {
                if (-not (Get-ConfigValue $entry 'channelType')) { continue }
                $channel = $channels | Where-Object { (Get-ConfigValue $_ 'displayName') -eq $entry.title } | Select-Object -First 1
                if (-not $channel) { Add-Row -What 'channel' -Name $entry.title -Result 'gone' -Detail 'not there'; continue }

                # The default channel cannot be removed - Teams refuses, and so does
                # this. Its name follows the team's language, hence both spellings.
                if ((Get-ConfigValue $channel 'displayName') -in @('General', 'Algemeen')) {
                    Add-Row -What 'channel' -Name $entry.title -Result 'skipped' -Detail 'the General channel cannot be removed'
                    continue
                }

                if (-not $Apply) {
                    Add-Row -What 'channel' -Name $entry.title -Result 'would' -Detail 'with the files in its folder'
                    continue
                }
                try {
                    Invoke-StructureGraph -Url "v1.0/teams/$($teamGroup.Id)/channels/$(Get-ConfigValue $channel 'id')" -Method DELETE | Out-Null
                    Add-Row -What 'channel' -Name $entry.title -Result 'removed' -Detail 'recoverable for 30 days'
                } catch {
                    Add-Row -What 'channel' -Name $entry.title -Result 'failed' -Detail $_.Exception.Message
                }
            }
        }
    }

    # -- SharePoint: libraries, content types, columns, term set ---------------
    $needsPnP = @('Libraries', 'ContentTypes', 'Columns', 'TermSet') | Where-Object { $_ -in $wanted }
    if ($needsPnP) {
        foreach ($siteKey in $config.sites.PSObject.Properties.Name) {
            $url = Get-StructureSiteUrl -Config $config -SiteKey $siteKey
            Write-Head "Site '$siteKey' - $url"

            $connection = $null
            try {
                $connection = Connect-Structure -Url $url -Tenant $Tenant -ClientId $ClientId -Interactive:$Interactive
            } catch {
                Write-Bad "could not connect: $($_.Exception.Message)"
                $failed++
                continue
            }

            # -- libraries
            if ('Libraries' -in $wanted) {
                Write-Step 'Libraries'
                foreach ($entry in ($config.containers | Where-Object { (Get-ConfigValue $_ 'kind' 'Library') -eq 'Library' -and $_.site -eq $siteKey })) {
                    $title = Get-ConfigValue $entry 'list' $entry.title
                    # The team's own library is not ours to remove - everything else in
                    # the team lives in it.
                    if ($title -in @('Documents', 'Documenten', 'Shared Documents')) {
                        Add-Row -What 'library' -Name $title -Result 'skipped' -Detail 'the team library is never removed'
                        continue
                    }
                    $list = Get-PnPList -Identity $title -Connection $connection -Includes ItemCount -ErrorAction SilentlyContinue
                    if (-not $list) { Add-Row -What 'library' -Name $title -Result 'gone' -Detail 'not there'; continue }

                    $count = [int] $list.ItemCount
                    if ($count -gt 0 -and -not $IncludeContent) {
                        Add-Row -What 'library' -Name $title -Result 'skipped' -Detail "$count item(s) in it - pass -IncludeContent"
                        continue
                    }
                    if (-not $Apply) { Add-Row -What 'library' -Name $title -Result 'would' -Detail "$count item(s)"; continue }
                    try {
                        Remove-PnPList -Identity $title -Force -Connection $connection
                        Add-Row -What 'library' -Name $title -Result 'removed' -Detail "$count item(s) to the recycle bin"
                    } catch {
                        Add-Row -What 'library' -Name $title -Result 'failed' -Detail $_.Exception.Message
                    }
                }
            }

            # -- content types, before the columns they hold
            if ('ContentTypes' -in $wanted) {
                Write-Step 'Content types'
                foreach ($definition in $config.contentTypes) {
                    $contentType = Get-PnPContentType -Identity $definition.name -Connection $connection -ErrorAction SilentlyContinue
                    if (-not $contentType) { Add-Row -What 'content type' -Name $definition.name -Result 'gone' -Detail 'not there'; continue }
                    if (-not $Apply) { Add-Row -What 'content type' -Name $definition.name -Result 'would'; continue }

                    # Unbind first: a content type still attached to a list cannot go.
                    foreach ($entry in ($config.containers | Where-Object { $_.site -eq $siteKey })) {
                        $listTitle = Get-ConfigValue $entry 'list' $entry.title
                        if ($definition.name -notin @(Get-ConfigValue $entry 'contentTypes' @())) { continue }
                        try {
                            Remove-PnPContentTypeFromList -List $listTitle -ContentType $definition.name -Connection $connection -ErrorAction Stop
                        } catch {
                            # The list may be gone already, which is fine.
                        }
                    }

                    try {
                        Remove-PnPContentType -Identity $definition.name -Force -Connection $connection -ErrorAction Stop
                        Add-Row -What 'content type' -Name $definition.name -Result 'removed'
                    } catch {
                        Add-Row -What 'content type' -Name $definition.name -Result 'failed' -Detail $_.Exception.Message
                    }
                }
            }

            # -- columns
            if ('Columns' -in $wanted) {
                Write-Step 'Site columns'
                foreach ($definition in $config.columns) {
                    $field = Get-PnPField -Identity $definition.internalName -Connection $connection -ErrorAction SilentlyContinue
                    if (-not $field) { Add-Row -What 'column' -Name $definition.internalName -Result 'gone' -Detail 'not there'; continue }
                    if (-not $Apply) {
                        Add-Row -What 'column' -Name $definition.internalName -Result 'would' -Detail 'takes its data with it'
                        continue
                    }
                    try {
                        Remove-PnPField -Identity $definition.internalName -Force -Connection $connection -ErrorAction Stop
                        Add-Row -What 'column' -Name $definition.internalName -Result 'removed'
                    } catch {
                        Add-Row -What 'column' -Name $definition.internalName -Result 'failed' -Detail $_.Exception.Message
                    }
                }
            }

            # -- term set, once: the term store is tenant-wide
            if ('TermSet' -in $wanted -and $siteKey -eq $config.sites.PSObject.Properties.Name[0]) {
                Write-Step 'Term store'
                $termStore = Get-ConfigValue $config 'termStore'
                if (-not $termStore) {
                    Write-Skip 'no termStore section'
                } else {
                    $group = Get-PnPTermGroup -Identity $termStore.group -Connection $connection -ErrorAction SilentlyContinue
                    if (-not $group) {
                        Add-Row -What 'term group' -Name $termStore.group -Result 'gone' -Detail 'not there'
                    } elseif (-not $Apply) {
                        Add-Row -What 'term group' -Name $termStore.group -Result 'would' `
                            -Detail 'every Leverancier tag on a document stops resolving'
                    } else {
                        try {
                            Remove-PnPTermGroup -Identity $termStore.group -Force -Connection $connection -ErrorAction Stop
                            Add-Row -What 'term group' -Name $termStore.group -Result 'removed' -Detail 'tags on documents no longer resolve'
                        } catch {
                            Add-Row -What 'term group' -Name $termStore.group -Result 'failed' -Detail $_.Exception.Message
                        }
                    }
                }
            }
        }
    }

    # -- Entra ID groups -------------------------------------------------------
    if ('Groups' -in $wanted) {
        Write-Head 'Entra ID security groups'
        Connect-Graph
        foreach ($definition in (Get-ConfigValue $config 'groups' @())) {
            $escaped = $definition.displayName -replace "'", "''"
            $found   = Invoke-StructureGraph -Url "v1.0/groups?`$filter=displayName eq '$escaped'&`$select=id,displayName"
            $group   = @(Get-ConfigValue $found 'value' @()) | Select-Object -First 1
            if (-not $group) { Add-Row -What 'group' -Name $definition.displayName -Result 'gone' -Detail 'not there'; continue }
            if (-not $Apply) { Add-Row -What 'group' -Name $definition.displayName -Result 'would'; continue }
            try {
                Invoke-StructureGraph -Url "v1.0/groups/$(Get-ConfigValue $group 'id')" -Method DELETE | Out-Null
                Add-Row -What 'group' -Name $definition.displayName -Result 'removed'
            } catch {
                Add-Row -What 'group' -Name $definition.displayName -Result 'failed' -Detail $_.Exception.Message
            }
        }
    }

    # -- The team --------------------------------------------------------------
    if ('Team' -in $wanted) {
        Write-Head 'The team itself'
        $teamGroup = Resolve-Team
        if (-not $teamGroup) {
            Add-Row -What 'team' -Name (Get-ConfigValue $config.team 'displayName' '?') -Result 'gone' -Detail 'not there'
        } elseif (-not $Apply) {
            Add-Row -What 'team' -Name ($teamGroup.DisplayName) -Result 'would' `
                -Detail 'the site, every library, every file and every chat message'
        } else {
            Write-Host ''
            Write-Warn "This removes the team '$($teamGroup.DisplayName)', its SharePoint site, every library,"
            Write-Warn 'every file in them and every chat message. Soft-deleted for 30 days.'
            Write-Host "  Type the team name to confirm: " -NoNewline -ForegroundColor Red
            $typed = Read-Host
            if ($typed -ne ($teamGroup.DisplayName)) {
                Add-Row -What 'team' -Name ($teamGroup.DisplayName) -Result 'skipped' -Detail 'name not confirmed'
            } else {
                try {
                    Invoke-StructureGraph -Url "v1.0/groups/$($teamGroup.Id)" -Method DELETE | Out-Null
                    Add-Row -What 'team' -Name ($teamGroup.DisplayName) -Result 'removed' -Detail 'restorable for 30 days in Entra ID'
                } catch {
                    Add-Row -What 'team' -Name ($teamGroup.DisplayName) -Result 'failed' -Detail $_.Exception.Message
                }
            }
        }
    }

    # -- Summary ---------------------------------------------------------------
    if ($report.Count -gt 0) {
        if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }
        $report | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    }

    Write-Head 'Summary'
    if ($Apply) {
        Write-Host "    Removed : $removed" -ForegroundColor Cyan
        Write-Host "    Skipped : $skipped" -ForegroundColor Cyan
        if ($failed -gt 0) { Write-Bad "$failed could not be removed - see the report" }
    } else {
        $would = @($report | Where-Object { $_.Result -eq 'would' }).Count
        Write-Host "    Would remove : $would" -ForegroundColor Yellow
        Write-Host "    Skipped      : $skipped" -ForegroundColor Yellow
        Write-Host ''
        Write-Host '    Nothing was changed. Rerun with -Apply to actually remove.' -ForegroundColor Yellow
    }
    if ($report.Count -gt 0) { Write-Host "    Report  : $ReportPath" -ForegroundColor Cyan }
    Write-Host ''
} catch {
    # Name the line and the statement. "Aborted: the property 'id' cannot be found"
    # with nothing else attached is a riddle, and this script is the one you reach for
    # when something already went wrong.
    Write-Host ''
    Write-Bad "Aborted at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    if ($_.InvocationInfo.Line) {
        Write-Host "    $($_.InvocationInfo.Line.Trim())" -ForegroundColor DarkRed
    }
    Write-Host ''
    exit 1
}

exit $(if ($failed -gt 0) { 1 } else { 0 })
