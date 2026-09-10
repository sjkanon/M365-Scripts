#Requires -Version 7.0
<#
.SYNOPSIS
    Provision the libraries and channel folders from a structure configuration:
    content types bound, default metadata set, and permissions granted to the Entra
    ID security group per pillar. Idempotent, supports -WhatIf.

.DESCRIPTION
    Step two of the structure set. Assumes New-SharePointMetadata.ps1 has already
    created the columns and content types.

    Per entry in the configuration's containers section:

      1. Container    a document library (kind: Library) is created when missing; a
                      Teams channel folder (kind: ChannelFolder) is resolved inside
                      the channel's library and created when missing.
      2. Content types the pillar's content types are bound to the library, and the
                      folder is given its own content type order so the New menu in
                      the Leveranciers channel offers Leveranciersdocument and not
                      the five types belonging to the other pillars.
      3. Defaults     default column values on the folder or library root, so Pijler
                      and Status are filled in without anyone typing them.
      4. View         a grouped view per pillar, added but never made the default -
                      changing the default view of a Teams library surprises everyone
                      in the channel at once.
      5. Permissions  inheritance broken where the config says so, and the configured
                      role granted to each Entra ID security group.

    Standard channels and permissions
    ---------------------------------
    Read this before running with the shipped configuration.

    A Teams standard channel is visible to every member of the team by design.
    Tightening the SharePoint permissions on the channel's folder does hide the
    files, but Teams keeps showing the channel: members who lost access get an error
    on the Files tab instead of a closed door, and Microsoft does not support the
    combination. It works, it is just not pretty and it is not blessed.

    The supported alternative is a private channel (its own site collection, its own
    membership - which is what MGMT already uses here) or a shared channel. If a
    pillar genuinely has to be closed off, move it to one of those and set
    uniquePermissions to false for its container.

    The script warns on every standard-channel folder it breaks inheritance on.
    -SkipChannelFolderPermissions leaves those folders inheriting and only does the
    rest, which is the conservative way to run this.

    Entra ID groups
    ---------------
    Groups are matched by display name. -EnsureGroups creates the ones that do not
    exist yet (Microsoft.Graph.Groups required); without it a missing group is
    reported and its grant skipped, so a typo in the config cannot silently leave a
    library wide open.

    Nothing is removed unless -RemoveOtherPermissions is given. That switch strips
    role assignments the config does not mention, which is what you want on the
    external FUTECH library and almost never want on an internal one.

.PARAMETER ConfigPath
    Path to the structure configuration JSON.
    Default: petsolutions.config.json next to this script.

.PARAMETER Container
    Only handle these containers (keys from the containers section, e.g. Leveranciers,
    FUTECH). Default: all of them.

.PARAMETER EnsureGroups
    Create the Entra ID security groups from the configuration that do not exist yet.
    Needs Group.ReadWrite.All.

.PARAMETER SkipChannelFolderPermissions
    Do not touch permissions on Teams channel folders - only libraries created by
    this script get unique permissions. The conservative option, see the note on
    standard channels above.

.PARAMETER RemoveOtherPermissions
    Also remove role assignments that the configuration does not list. Site owners,
    sharing-link groups and the everyone-claim are never removed.

.PARAMETER RemoveStockContentType
    Remove the built-in Document content type from a library once the pillar's own
    content types are bound, so nobody can file a document without metadata.

.PARAMETER Interactive
    Sign in interactively (default when no certificate is supplied).

.PARAMETER ClientId
    Client ID of the Entra app registration used to sign in.

.PARAMETER Thumbprint
    Certificate thumbprint for app-only sign-in. Implies app-only.

.PARAMETER CertificatePath
    PFX file for app-only sign-in, as an alternative to -Thumbprint.

.PARAMETER CertificatePassword
    Password for -CertificatePath.

.PARAMETER Tenant
    Tenant name or ID. Defaults to the tenant in the configuration file.

.PARAMETER ReportPath
    Write the resulting permissions per container to this CSV.
    Default: C:\Temp\SharePointStructure_Permissions_<timestamp>.csv

.PARAMETER Disconnect
    Sign out of PnP when finished.

.EXAMPLE
    # Dry run - what would this do to the libraries and their permissions?
    .\Set-SharePointLibraries.ps1 -Interactive -ClientId <app-id> -WhatIf

.EXAMPLE
    # Full provisioning, creating the security groups on the way
    .\Set-SharePointLibraries.ps1 -Interactive -ClientId <app-id> -EnsureGroups

.EXAMPLE
    # Only the external customer library, and strip anything the config does not list
    .\Set-SharePointLibraries.ps1 -Interactive -ClientId <app-id> `
        -Container FUTECH -RemoveOtherPermissions

.EXAMPLE
    # Everything except the unsupported standard-channel folder permissions
    .\Set-SharePointLibraries.ps1 -Interactive -ClientId <app-id> `
        -SkipChannelFolderPermissions

.NOTES
    Author  : Sjoerd Kanon
    Requires: PnP.PowerShell 2.x, Microsoft.Graph.Groups for -EnsureGroups
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $ConfigPath,
    [string[]] $Container,

    [switch] $EnsureGroups,
    [switch] $SkipChannelFolderPermissions,
    [switch] $RemoveOtherPermissions,
    [switch] $RemoveStockContentType,

    [switch] $Interactive,
    [string] $ClientId,
    [string] $Thumbprint,
    [string] $CertificatePath,
    [securestring] $CertificatePassword,
    [string] $Tenant,

    [string] $ReportPath,
    [switch] $Disconnect
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

. (Join-Path $PSScriptRoot 'SharePointStructure.Common.ps1')

trap {
    Write-Host ''
    Write-Host "  FAILED at line $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    Write-Host "    $($_.InvocationInfo.Line.Trim())" -ForegroundColor DarkRed
    Write-Host "    $($_.Exception.GetType().Name): $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ''
    break
}

Assert-PnPModule

if (-not $ConfigPath) { $ConfigPath = Join-Path $PSScriptRoot 'petsolutions.config.json' }
$config = Import-StructureConfig -Path $ConfigPath
if (-not $Tenant) { $Tenant = $config.tenant }

$simulate   = [bool] $WhatIfPreference
$containers = @($config.containers)
if ($Container) {
    $known = @($containers | ForEach-Object { $_.key })
    foreach ($key in $Container) {
        if ($key -notin $known) { throw "Unknown container '$key'. Known: $($known -join ', ')" }
    }
    $containers = @($containers | Where-Object { $_.key -in $Container })
}

$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not $ReportPath) {
    $ReportPath = Join-Path $outputDir "SharePointStructure_Permissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
}

$connectSplat = @{
    Tenant              = $Tenant
    ClientId            = $ClientId
    Thumbprint          = $Thumbprint
    CertificatePath     = $CertificatePath
    CertificatePassword = $CertificatePassword
    Interactive         = $Interactive
}

$changeCount   = 0
$warningCount  = 0
$report        = [System.Collections.Generic.List[object]]::new()
$groupObjectId = @{}          # display name -> Entra object id, resolved once
$principalMap  = @{}          # "<site key>|<display name>" -> SharePoint principal

Write-Host ''
Write-Host "  Client : $(Get-ConfigValue $config 'client' '(unnamed)')" -ForegroundColor Cyan
Write-Host "  Config : $ConfigPath" -ForegroundColor Cyan
Write-Host "  Scope  : $($containers.Count) container(s) - $(($containers | ForEach-Object { $_.key }) -join ', ')" -ForegroundColor Cyan
Write-Host "  Mode   : $(if ($simulate) { '-WhatIf - nothing will be changed' } else { 'APPLY' })" -ForegroundColor $(if ($simulate) { 'Yellow' } else { 'Cyan' })

# -- Content type order on a folder --------------------------------------------
function Set-FolderContentTypeOrder {
    <#
        Give a channel folder its own content type order, so the New menu inside the
        Leveranciers channel offers Leveranciersdocument first and does not list the
        five types that belong to the other pillars. The list still carries them all -
        it has to, they share one library.

        Returns $true when something changed.
    #>
    param(
        [Parameter(Mandatory)] $Folder,
        [Parameter(Mandatory)] $List,
        [Parameter(Mandatory)] [string[]] $ContentTypeNames,
        [Parameter(Mandatory)] $Connection,
        [switch] $WhatIfMode
    )

    $context = Get-PnPContext -Connection $Connection
    $context.Load($List.ContentTypes)
    $context.Load($Folder, 'ContentTypeOrder', 'UniqueContentTypeOrder')
    $context.ExecuteQuery()

    $wanted = @()
    foreach ($name in $ContentTypeNames) {
        $match = $List.ContentTypes | Where-Object { $_.Name -eq $name } | Select-Object -First 1
        if ($match) { $wanted += $match.Id }
    }
    if ($wanted.Count -eq 0) { return $false }

    $current = @()
    if ($Folder.UniqueContentTypeOrder) { $current = @($Folder.UniqueContentTypeOrder | ForEach-Object { $_.StringValue }) }
    $wantedStrings = @($wanted | ForEach-Object { $_.StringValue })

    if (($current -join '|') -eq ($wantedStrings -join '|')) { return $false }
    if ($WhatIfMode) { return $true }

    $collection = New-Object 'System.Collections.Generic.List[Microsoft.SharePoint.Client.ContentTypeId]'
    foreach ($id in $wanted) { $collection.Add($id) }
    $Folder.UniqueContentTypeOrder = $collection
    $Folder.Update()
    $context.ExecuteQuery()
    return $true
}

# -- Entra ID groups -----------------------------------------------------------
function Initialize-StructureGroup {
    <#
        Resolve every group the selected containers grant to, once for the whole run.
        A group that cannot be resolved is reported here rather than per library, so
        the operator sees the whole problem in one place before anything is changed.
    #>
    $needed = @()
    foreach ($entry in $containers) {
        foreach ($permission in (Get-ConfigValue $entry 'permissions' @())) { $needed += $permission.group }
    }
    $needed = @($needed | Select-Object -Unique)
    if ($needed.Count -eq 0) { Write-Skip 'No group grants in scope'; return }

    Connect-StructureGraph -Tenant $Tenant -ClientId $ClientId -Thumbprint $Thumbprint

    foreach ($name in $needed) {
        $definition = @(Get-ConfigValue $config 'groups' @() | Where-Object { $_.displayName -eq $name }) | Select-Object -First 1
        $group = Resolve-StructureGroup -Definition $definition -Create:$EnsureGroups -WhatIfMode:$simulate

        if ($group) {
            $groupObjectId[$name] = $group.Id
            Write-Ok "group '$name' ($($group.Id))"
            if ($EnsureGroups -and -not $simulate) { $script:changeCount++ }
        } else {
            $script:warningCount++
            if ($EnsureGroups) {
                Write-Warn "group '$name' would be created - its grants are skipped in this -WhatIf run"
            } else {
                Write-Warn "group '$name' does not exist - its grants are skipped. Rerun with -EnsureGroups to create it."
            }
        }
    }
}

function Get-StructurePrincipal {
    <#
        SharePoint principal for a configured group, per site collection. Cached:
        EnsureUser is a write on the site's user information list, so calling it once
        per library would be six writes for nothing.
    #>
    param([string] $GroupName, [string] $SiteKey, $Connection)

    $key = "$SiteKey|$GroupName"
    if ($principalMap.ContainsKey($key)) { return $principalMap[$key] }
    if (-not $groupObjectId.ContainsKey($GroupName)) { return $null }

    try {
        $principal = Get-SecurablePrincipal -ObjectId $groupObjectId[$GroupName] -Connection $Connection
    } catch {
        Write-Warn "could not resolve '$GroupName' on this site: $($_.Exception.Message)"
        $script:warningCount++
        return $null
    }
    $principalMap[$key] = $principal
    return $principal
}

# -- Containers ----------------------------------------------------------------
function Set-StructureContainer {
    <# One library or channel folder, end to end. #>
    param($Definition, $Connection)

    $kind      = Get-ConfigValue $Definition 'kind' 'Library'
    $listTitle = Get-ConfigValue $Definition 'list' $Definition.title

    Write-Step "Container '$($Definition.key)' ($kind - $listTitle)"

    # -- the list itself
    $list = Get-PnPList -Identity $listTitle -Connection $Connection -ErrorAction SilentlyContinue
    if (-not $list) {
        if ($kind -eq 'ChannelFolder') {
            Write-Bad "library '$listTitle' not found - is the Teams channel created and its Files tab opened once?"
            $script:warningCount++
            return
        }
        if (-not $PSCmdlet.ShouldProcess($listTitle, 'Create document library')) { return }
        $list = New-PnPList -Title $listTitle -Template DocumentLibrary -EnableContentTypes -OnQuickLaunch -Connection $Connection
        Write-Change "library '$listTitle' created"
        $script:changeCount++
    } else {
        Write-Ok "library '$listTitle'"
    }
    if (-not $list) { return }

    if (-not $list.ContentTypesEnabled) {
        if ($PSCmdlet.ShouldProcess($listTitle, 'Enable content type management')) {
            Set-PnPList -Identity $listTitle -EnableContentTypes $true -Connection $Connection | Out-Null
            Write-Change "content types enabled on '$listTitle'"
            $script:changeCount++
        }
    }

    # -- the folder, for a channel container
    $folder = $null
    if ($kind -eq 'ChannelFolder') {
        $folderName  = Get-ConfigValue $Definition 'folder' $Definition.title
        $context     = Get-PnPContext -Connection $Connection
        $context.Load($list.RootFolder)
        $context.ExecuteQuery()
        $folderUrl = "$($list.RootFolder.ServerRelativeUrl.TrimEnd('/'))/$folderName"

        $folder = Get-PnPFolder -Url $folderUrl -Connection $Connection -ErrorAction SilentlyContinue
        if (-not $folder) {
            if ($PSCmdlet.ShouldProcess($folderUrl, 'Create channel folder')) {
                # Teams normally creates this when the channel is opened; doing it
                # here keeps a fresh team from blocking the whole run.
                $folder = Add-PnPFolder -Name $folderName -Folder $list.RootFolder.ServerRelativeUrl -Connection $Connection
                Write-Change "folder '$folderName' created"
                $script:changeCount++
            }
        } else {
            Write-Ok "folder '$folderName'"
        }
    }

    # -- content types on the list
    $context = Get-PnPContext -Connection $Connection
    $context.Load($list.ContentTypes)
    $context.ExecuteQuery()
    $bound = @($list.ContentTypes | ForEach-Object { $_.Name })

    $wantedTypes = @(Get-ConfigValue $Definition 'contentTypes' @())
    foreach ($name in $wantedTypes) {
        if ($name -in $bound) { Write-Ok "content type '$name' bound"; continue }
        if ($PSCmdlet.ShouldProcess("$listTitle / $name", 'Bind content type')) {
            Add-PnPContentTypeToList -List $listTitle -ContentType $name -Connection $Connection
            Write-Change "content type '$name' bound to '$listTitle'"
            $script:changeCount++
        }
    }

    $defaultType = Get-ConfigValue $Definition 'defaultContentType'
    if ($defaultType -and $kind -eq 'Library') {
        if ($PSCmdlet.ShouldProcess("$listTitle / $defaultType", 'Set as default content type')) {
            Set-PnPDefaultContentTypeToList -List $listTitle -ContentType $defaultType -Connection $Connection
            Write-Ok "default content type '$defaultType'"
        }
    }

    if ($RemoveStockContentType -and 'Document' -in $bound -and $wantedTypes.Count -gt 0) {
        if ($PSCmdlet.ShouldProcess("$listTitle / Document", 'Remove built-in content type')) {
            try {
                Remove-PnPContentTypeFromList -List $listTitle -ContentType 'Document' -Connection $Connection
                Write-Change "built-in 'Document' content type removed from '$listTitle'"
                $script:changeCount++
            } catch {
                Write-Warn "could not remove the built-in Document content type: $($_.Exception.Message)"
                $script:warningCount++
            }
        }
    }

    # -- per-folder content type order
    if ($folder -and $wantedTypes.Count -gt 0) {
        $ordered = @($defaultType) + @($wantedTypes | Where-Object { $_ -ne $defaultType })
        $ordered = @($ordered | Where-Object { $_ })
        $changed = Set-FolderContentTypeOrder -Folder $folder -List $list -ContentTypeNames $ordered `
            -Connection $Connection -WhatIfMode:$simulate
        if ($changed) {
            Write-Change "folder content type order set to: $($ordered -join ', ')"
            $script:changeCount++
        }
    }

    # -- default column values
    $defaults = Get-ConfigValue $Definition 'defaultColumnValues'
    if ($defaults) {
        foreach ($property in $defaults.PSObject.Properties) {
            $target = if ($folder) { Get-ConfigValue $Definition 'folder' $Definition.title } else { '/' }
            if ($PSCmdlet.ShouldProcess("$listTitle$(if ($folder) { "/$target" })", "Default $($property.Name) = $($property.Value)")) {
                try {
                    Set-PnPDefaultColumnValues -List $listTitle -Field $property.Name -Value $property.Value `
                        -Folder $target -Connection $Connection
                    Write-Ok "default $($property.Name) = $($property.Value)"
                } catch {
                    Write-Warn "could not set the default for $($property.Name): $($_.Exception.Message)"
                    $script:warningCount++
                }
            }
        }
    }

    # -- view
    $view = Get-ConfigValue $Definition 'view'
    if ($view) {
        # On a shared channel library the views are list-wide, so the pillar name goes
        # in the title - six views called "Op merk" would be unusable.
        $viewTitle = if ($kind -eq 'ChannelFolder') { "$($Definition.title) - $($view.title)" } else { $view.title }
        $existing  = Get-PnPView -List $listTitle -Identity $viewTitle -Connection $Connection -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Ok "view '$viewTitle'"
        } elseif ($PSCmdlet.ShouldProcess("$listTitle / $viewTitle", 'Create view')) {
            $query = ''
            $groupBy = Get-ConfigValue $view 'groupBy'
            if ($groupBy) {
                $query = "<GroupBy Collapse=`"TRUE`" GroupLimit=`"100`"><FieldRef Name=`"$groupBy`" /></GroupBy>"
            }
            $viewSplat = @{
                List       = $listTitle
                Title      = $viewTitle
                Fields     = @(Get-ConfigValue $view 'fields' @())
                Connection = $Connection
            }
            if ($query) { $viewSplat['Query'] = $query }
            # Never -SetAsDefault: the default view of a Teams library is what every
            # member of the channel sees the moment they open Files.
            Add-PnPView @viewSplat | Out-Null
            Write-Change "view '$viewTitle' created"
            $script:changeCount++
        }
    }

    # -- permissions
    Set-StructurePermission -Definition $Definition -List $list -Folder $folder -Connection $Connection
}

function Set-StructurePermission {
    <# Break inheritance where configured and grant the roles from the config. #>
    param($Definition, $List, $Folder, $Connection)

    $wanted = @(Get-ConfigValue $Definition 'permissions' @())
    if (-not (Get-ConfigValue $Definition 'uniquePermissions' $false)) {
        $note = Get-ConfigValue $Definition 'note'
        Write-Skip "permissions inherited$(if ($note) { " - $note" })"
        return
    }
    if ($wanted.Count -eq 0) {
        Write-Skip 'uniquePermissions is set but no grants are configured - left inheriting'
        return
    }

    $kind = Get-ConfigValue $Definition 'kind' 'Library'
    if ($kind -eq 'ChannelFolder') {
        if ($SkipChannelFolderPermissions) {
            Write-Skip 'channel folder permissions skipped (-SkipChannelFolderPermissions)'
            return
        }
        # Said once per folder, on purpose: this is the part of the model Microsoft
        # does not support and the part that generates the helpdesk calls.
        Write-Warn "'$($Definition.title)' is a standard-channel folder - members who lose access keep seeing the channel in Teams and get an error on the Files tab. A private or shared channel is the supported way to close a pillar off."
        $script:warningCount++
    }

    $map = @{}
    foreach ($entry in $wanted) {
        $principal = Get-StructurePrincipal -GroupName $entry.group -SiteKey $Definition.site -Connection $Connection
        if ($principal) { $map[$entry.group] = $principal }
    }

    $securable = $null
    if ($Folder) {
        $context = Get-PnPContext -Connection $Connection
        $context.Load($Folder.ListItemAllFields)
        $context.ExecuteQuery()
        $securable = $Folder.ListItemAllFields
    } else {
        $securable = $List
    }
    if (-not $securable) { Write-Skip 'nothing to secure yet (-WhatIf on a container that does not exist)'; return }

    $changes = Set-SecurableRole -Securable $securable -Desired $wanted -Connection $Connection `
        -PrincipalMap $map `
        -KeepExisting:([bool](Get-ConfigValue $Definition 'keepExistingPermissions' $true)) `
        -RemoveOther:$RemoveOtherPermissions `
        -WhatIfMode:$simulate

    if ($changes.Count -eq 0) {
        Write-Ok 'permissions already match the configuration'
    } else {
        foreach ($change in $changes) { Write-Change "permissions: $change" }
        $script:changeCount += $changes.Count
    }

    if (-not $simulate) {
        foreach ($row in (Get-StructureRoleReport -Securable $securable -Connection $Connection)) {
            $report.Add([PSCustomObject]@{
                Container = $Definition.key
                Site      = $Definition.site
                Target    = if ($Folder) { "$($Definition.list)/$($Definition.folder)" } else { $Definition.list }
                Principal = $row.Principal
                LoginName = $row.LoginName
                Roles     = $row.Roles
            })
        }
    }
}

# -- Run -----------------------------------------------------------------------
try {
    Write-Head 'Entra ID groups'
    Initialize-StructureGroup

    foreach ($siteKey in @($containers | ForEach-Object { $_.site } | Select-Object -Unique)) {
        $url = Get-StructureSiteUrl -Config $config -SiteKey $siteKey
        Write-Head "Site '$siteKey' - $url"
        $connection = Connect-Structure -Url $url @connectSplat

        foreach ($entry in ($containers | Where-Object { $_.site -eq $siteKey })) {
            Set-StructureContainer -Definition $entry -Connection $connection
        }
    }

    Write-Head 'Summary'
    if ($report.Count -gt 0) {
        if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }
        $report | Sort-Object Container, Principal | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
        Write-Host "    Permission report: $ReportPath" -ForegroundColor Cyan
    }
    if ($simulate) {
        Write-Host "    $changeCount change(s) would be made - rerun without -WhatIf to apply." -ForegroundColor Yellow
    } elseif ($changeCount -eq 0) {
        Write-Ok 'Libraries and permissions already match the configuration - nothing changed.'
    } else {
        Write-Change "$changeCount change(s) applied."
    }
    if ($warningCount -gt 0) {
        Write-Host "    $warningCount warning(s) - scroll back, or rerun with -SkipChannelFolderPermissions." -ForegroundColor Yellow
    }
    Write-Host ''
} finally {
    if ($Disconnect) {
        Disconnect-Structure
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
        Write-Host '  Disconnected.' -ForegroundColor DarkGray
    }
}
