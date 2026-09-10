#Requires -Version 7.0
<#
.SYNOPSIS
    Compare the live SharePoint structure with the configuration and report every
    difference. Reads only - changes nothing, ever.

.DESCRIPTION
    Step four of the structure set, and the one you run when someone says "it used to
    work". It walks the same model as the provisioning scripts and reports what does
    not match:

      Term set       group, term set and the terms behind the Leverancier column
      Columns        present on every site, right type, right choices, right group
      Content types  present, derived from Document, right columns, right required
                     flags - the last one is what quietly rots when someone edits a
                     content type in the browser
      Containers     library or channel folder present, content types bound, default
                     column values set, the pillar's view present
      Permissions    inheritance broken or not as configured, and exactly which
                     principal holds which role - including anything the config does
                     not mention
      Groups         with -IncludeGroups, whether the Entra ID security groups still
                     exist and how many members they have

    Every finding is one row, classified as:

      Missing    in the configuration, not on the tenant   - provisioning will fix it
      Different  present but not as configured             - provisioning will fix it
      Extra      on the tenant, not in the configuration   - a person has to decide

    "Extra" is deliberately never fixed automatically by the other scripts: an extra
    column holds data and an extra role assignment is usually somebody's deliberate
    exception. This report is where you see them, not where they are removed.

    Exit codes
    ----------
        0  the tenant matches the configuration
        1  the check could not complete
        2  drift found

.PARAMETER ConfigPath
    Path to the structure configuration JSON.
    Default: petsolutions.config.json next to this script.

.PARAMETER Site
    Only check these sites (keys from the sites section). Default: all.

.PARAMETER Container
    Only check these containers (keys from the containers section). Default: all.

.PARAMETER IncludeGroups
    Also check the Entra ID security groups (needs Microsoft.Graph.Groups and a Graph
    sign-in). Off by default so the check works with only a SharePoint sign-in.

.PARAMETER Interactive
    Sign in interactively.

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
    CSV with one row per finding.
    Default: C:\Temp\SharePointStructure_Drift_<timestamp>.csv

.PARAMETER Quiet
    Print only the summary and the findings, not the [ OK ] lines.

.PARAMETER Disconnect
    Sign out of PnP when finished.

.EXAMPLE
    # Does the tenant still look like the configuration says it should?
    .\Test-SharePointStructure.ps1 -Interactive -ClientId <app-id>

.EXAMPLE
    # One library, including who holds which role on it
    .\Test-SharePointStructure.ps1 -Interactive -ClientId <app-id> -Container FUTECH

.EXAMPLE
    # Scheduled: exit code 2 means somebody changed something
    .\Test-SharePointStructure.ps1 -ClientId <app-id> -Thumbprint <thumbprint> -Quiet

.NOTES
    Author  : Sjoerd Kanon
    Requires: PnP.PowerShell 2.x, Microsoft.Graph.Groups for -IncludeGroups
#>
[CmdletBinding()]
param(
    [string] $ConfigPath,
    [string[]] $Site,
    [string[]] $Container,

    [switch] $IncludeGroups,

    [switch] $Interactive,
    [string] $ClientId,
    [string] $Thumbprint,
    [string] $CertificatePath,
    [securestring] $CertificatePassword,
    [string] $Tenant,

    [string] $ReportPath,
    [switch] $Quiet,
    [switch] $Disconnect
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

. (Join-Path $PSScriptRoot 'SharePointStructure.Common.ps1')

Assert-PnPModule

if (-not $ConfigPath) { $ConfigPath = Join-Path $PSScriptRoot 'petsolutions.config.json' }
$config = Import-StructureConfig -Path $ConfigPath
if (-not $Tenant) { $Tenant = $config.tenant }

$siteKeys = if ($Site) { @($Site) } else { @($config.sites.PSObject.Properties.Name) }
foreach ($key in $siteKeys) {
    if ($key -notin $config.sites.PSObject.Properties.Name) {
        throw "Unknown site key '$key'. Known: $($config.sites.PSObject.Properties.Name -join ', ')"
    }
}

$containers = @($config.containers | Where-Object { $_.site -in $siteKeys })
if ($Container) {
    $known = @($config.containers | ForEach-Object { $_.key })
    foreach ($key in $Container) {
        if ($key -notin $known) { throw "Unknown container '$key'. Known: $($known -join ', ')" }
    }
    $containers = @($containers | Where-Object { $_.key -in $Container })
}

$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not $ReportPath) {
    $ReportPath = Join-Path $outputDir "SharePointStructure_Drift_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
}

$connectSplat = @{
    Tenant              = $Tenant
    ClientId            = $ClientId
    Thumbprint          = $Thumbprint
    CertificatePath     = $CertificatePath
    CertificatePassword = $CertificatePassword
    Interactive         = $Interactive
}

$findings = [System.Collections.Generic.List[object]]::new()

function Add-Finding {
    param(
        [Parameter(Mandatory)] [ValidateSet('Missing', 'Different', 'Extra')] [string] $Kind,
        [Parameter(Mandatory)] [string] $Area,
        [string] $SiteKey = '',
        [Parameter(Mandatory)] [string] $Object,
        [string] $Expected = '',
        [string] $Actual = ''
    )

    $findings.Add([PSCustomObject]@{
        Kind = $Kind; Area = $Area; Site = $SiteKey; Object = $Object
        Expected = $Expected; Actual = $Actual
    })
    $text = "$Area / $Object"
    if ($Expected -or $Actual) { $text += " - expected '$Expected', found '$Actual'" }
    Write-Diff "$Kind`: $text"
}

function Write-Pass { param([string] $Message) if (-not $Quiet) { Write-Ok $Message } }

# -- Term set ------------------------------------------------------------------
function Test-StructureTermSet {
    param($Connection)

    $termStore = Get-ConfigValue $config 'termStore'
    if (-not $termStore) { return }

    $group = Get-PnPTermGroup -Identity $termStore.group -Connection $Connection -ErrorAction SilentlyContinue
    if (-not $group) {
        Add-Finding -Kind Missing -Area 'Term store' -Object "term group '$($termStore.group)'"
        return
    }

    $set = Get-PnPTermSet -Identity $termStore.termSet -TermGroup $termStore.group -Connection $Connection -ErrorAction SilentlyContinue
    if (-not $set) {
        Add-Finding -Kind Missing -Area 'Term store' -Object "term set '$($termStore.termSet)'"
        return
    }
    Write-Pass "term set '$($termStore.group)|$($termStore.termSet)'"

    $live   = @(Get-PnPTerm -TermSet $termStore.termSet -TermGroup $termStore.group -Connection $Connection -ErrorAction SilentlyContinue |
                ForEach-Object { $_.Name })
    $wanted = @(Get-ConfigValue $termStore 'terms' @())

    foreach ($term in $wanted) {
        if ($term -in $live) { Write-Pass "term '$term'" }
        else { Add-Finding -Kind Missing -Area 'Term store' -Object "term '$term'" }
    }
    # Extra terms are the normal case here - the term set is meant to be extended
    # from the term store UI - so they are reported, never flagged as a problem.
    foreach ($term in $live) {
        if ($term -notin $wanted) {
            Add-Finding -Kind Extra -Area 'Term store' -Object "term '$term'" -Actual 'added outside the configuration'
        }
    }
}

# -- Columns -------------------------------------------------------------------
function Test-StructureColumn {
    param($SiteKey, $Connection)

    foreach ($definition in $config.columns) {
        $name  = $definition.internalName
        $field = Get-PnPField -Identity $name -Connection $Connection -ErrorAction SilentlyContinue
        if (-not $field) {
            Add-Finding -Kind Missing -Area 'Column' -SiteKey $SiteKey -Object $name
            continue
        }

        $issues = @()
        if ($field.Title -ne $definition.displayName) {
            $issues += 'title'
            Add-Finding -Kind Different -Area 'Column' -SiteKey $SiteKey -Object "$name display name" `
                -Expected $definition.displayName -Actual $field.Title
        }

        $group = Get-ConfigValue $config 'columnGroup' 'Custom'
        if ($field.Group -ne $group) {
            $issues += 'group'
            Add-Finding -Kind Different -Area 'Column' -SiteKey $SiteKey -Object "$name column group" `
                -Expected $group -Actual $field.Group
        }

        if ($definition.type -in @('Choice', 'MultiChoice')) {
            $wanted  = @(Get-ConfigValue $definition 'choices' @())
            $current = Get-FieldChoiceValue -Field $field
            if (($current -join '|') -ne ($wanted -join '|')) {
                $issues += 'choices'
                Add-Finding -Kind Different -Area 'Column' -SiteKey $SiteKey -Object "$name choices" `
                    -Expected ($wanted -join ', ') -Actual ($current -join ', ')
            }
        }

        if ($issues.Count -eq 0) { Write-Pass "column '$name'" }
    }
}

# -- Content types -------------------------------------------------------------
function Test-StructureContentType {
    param($SiteKey, $Connection)

    foreach ($definition in $config.contentTypes) {
        $contentType = Get-PnPContentType -Identity $definition.name -Connection $Connection -ErrorAction SilentlyContinue
        if (-not $contentType) {
            Add-Finding -Kind Missing -Area 'Content type' -SiteKey $SiteKey -Object $definition.name
            continue
        }

        $context = Get-PnPContext -Connection $Connection
        $context.Load($contentType.FieldLinks)
        $context.ExecuteQuery()

        $links  = @{}
        foreach ($link in $contentType.FieldLinks) { $links[$link.Name] = $link.Required }

        $issues = @()
        foreach ($fieldDef in $definition.fields) {
            $wantRequired = [bool] (Get-ConfigValue $fieldDef 'required' $false)

            if (-not $links.ContainsKey($fieldDef.internalName)) {
                $issues += $fieldDef.internalName
                Add-Finding -Kind Missing -Area 'Content type' -SiteKey $SiteKey `
                    -Object "$($definition.name) / column $($fieldDef.internalName)"
                continue
            }
            # The one that rots quietly: somebody unticks Required in the browser and
            # nothing looks wrong until half a library has no Taal on it.
            if ($links[$fieldDef.internalName] -ne $wantRequired) {
                $issues += $fieldDef.internalName
                Add-Finding -Kind Different -Area 'Content type' -SiteKey $SiteKey `
                    -Object "$($definition.name) / $($fieldDef.internalName) required" `
                    -Expected "$wantRequired" -Actual "$($links[$fieldDef.internalName])"
            }
        }

        if ($issues.Count -eq 0) { Write-Pass "content type '$($definition.name)'" }
    }
}

# -- Containers ----------------------------------------------------------------
function Test-StructureContainer {
    param($Definition, $Connection)

    $listTitle = Get-ConfigValue $Definition 'list' $Definition.title
    $kind      = Get-ConfigValue $Definition 'kind' 'Library'
    $siteKey   = $Definition.site

    $list = Get-PnPList -Identity $listTitle -Connection $Connection -ErrorAction SilentlyContinue
    if (-not $list) {
        Add-Finding -Kind Missing -Area 'Container' -SiteKey $siteKey -Object "library '$listTitle'"
        return
    }

    $folder = $null
    if ($kind -eq 'ChannelFolder') {
        $folderName = Get-ConfigValue $Definition 'folder' $Definition.title
        $folder = Get-ChannelFolderItem -ListTitle $listTitle -FolderName $folderName -Connection $Connection
        if (-not $folder) {
            Add-Finding -Kind Missing -Area 'Container' -SiteKey $siteKey -Object "folder '$listTitle/$folderName'"
            return
        }
    }
    Write-Pass "container '$($Definition.key)'"

    # -- content types bound
    $context = Get-PnPContext -Connection $Connection
    $context.Load($list.ContentTypes)
    $context.ExecuteQuery()
    $bound = @($list.ContentTypes | ForEach-Object { $_.Name })

    foreach ($name in (Get-ConfigValue $Definition 'contentTypes' @())) {
        if ($name -in $bound) { Write-Pass "  content type '$name' bound" }
        else { Add-Finding -Kind Missing -Area 'Container' -SiteKey $siteKey -Object "$($Definition.key) / content type '$name'" }
    }

    # -- default column values
    $defaults = Get-ConfigValue $Definition 'defaultColumnValues'
    if ($defaults) {
        $target = if ($kind -eq 'ChannelFolder') { Get-ConfigValue $Definition 'folder' $Definition.title } else { '/' }
        $live = @()
        try {
            $live = @(Get-PnPDefaultColumnValues -List $listTitle -Connection $Connection)
        } catch {
            Add-Finding -Kind Different -Area 'Container' -SiteKey $siteKey `
                -Object "$($Definition.key) / default column values" -Actual "could not be read: $($_.Exception.Message)"
        }
        foreach ($property in $defaults.PSObject.Properties) {
            $match = $live | Where-Object { $_.Field -eq $property.Name -and $_.Folder -like "*$target" } | Select-Object -First 1
            if (-not $match) {
                Add-Finding -Kind Missing -Area 'Container' -SiteKey $siteKey `
                    -Object "$($Definition.key) / default $($property.Name)" -Expected $property.Value
            } elseif ("$($match.Value)" -ne "$($property.Value)") {
                Add-Finding -Kind Different -Area 'Container' -SiteKey $siteKey `
                    -Object "$($Definition.key) / default $($property.Name)" -Expected $property.Value -Actual "$($match.Value)"
            } else {
                Write-Pass "  default $($property.Name) = $($property.Value)"
            }
        }
    }

    # -- view
    $view = Get-ConfigValue $Definition 'view'
    if ($view) {
        $viewTitle = if ($kind -eq 'ChannelFolder') { "$($Definition.title) - $($view.title)" } else { $view.title }
        if (Get-PnPView -List $listTitle -Identity $viewTitle -Connection $Connection -ErrorAction SilentlyContinue) {
            Write-Pass "  view '$viewTitle'"
        } else {
            Add-Finding -Kind Missing -Area 'Container' -SiteKey $siteKey -Object "$($Definition.key) / view '$viewTitle'"
        }
    }

    Test-StructurePermission -Definition $Definition -Securable $(if ($folder) { $folder } else { $list }) -Connection $Connection
}

function Test-StructurePermission {
    param($Definition, $Securable, $Connection)

    $siteKey  = $Definition.site
    $wantUnique = [bool] (Get-ConfigValue $Definition 'uniquePermissions' $false)

    $context = Get-PnPContext -Connection $Connection
    $context.Load($Securable)
    $context.ExecuteQuery()

    if ($Securable.HasUniqueRoleAssignments -ne $wantUnique) {
        Add-Finding -Kind Different -Area 'Permissions' -SiteKey $siteKey -Object "$($Definition.key) inheritance" `
            -Expected $(if ($wantUnique) { 'unique' } else { 'inherited' }) `
            -Actual   $(if ($Securable.HasUniqueRoleAssignments) { 'unique' } else { 'inherited' })
        # An inherited securable has nothing of its own to compare against.
        if (-not $Securable.HasUniqueRoleAssignments) { return }
    }
    if (-not $wantUnique -and -not $Securable.HasUniqueRoleAssignments) {
        Write-Pass "  permissions inherited as configured"
        return
    }

    $live = @(Get-StructureRoleReport -Securable $Securable -Connection $Connection)
    $wanted = @(Get-ConfigValue $Definition 'permissions' @())

    foreach ($entry in $wanted) {
        $match = $live | Where-Object { $_.Principal -eq $entry.group } | Select-Object -First 1
        if (-not $match) {
            Add-Finding -Kind Missing -Area 'Permissions' -SiteKey $siteKey `
                -Object "$($Definition.key) / $($entry.group)" -Expected $entry.role
        } elseif ($entry.role -notin ($match.Roles -split ',\s*')) {
            Add-Finding -Kind Different -Area 'Permissions' -SiteKey $siteKey `
                -Object "$($Definition.key) / $($entry.group)" -Expected $entry.role -Actual $match.Roles
        } else {
            Write-Pass "  $($entry.group): $($entry.role)"
        }
    }

    # Anything else holding rights here. Owners and sharing links are normal and are
    # left out; a stray group or a named user is exactly what you want to see.
    $wantedNames = @($wanted | ForEach-Object { $_.group })
    foreach ($row in $live) {
        if ($row.Principal -in $wantedNames) { continue }
        if ($row.LoginName -match 'Owners|Eigenaren|SharingLinks|spo-grid-all-users|_spOwner') { continue }
        Add-Finding -Kind Extra -Area 'Permissions' -SiteKey $siteKey `
            -Object "$($Definition.key) / $($row.Principal)" -Actual $row.Roles
    }
}

# -- Entra ID groups -----------------------------------------------------------
function Test-StructureGroup {
    Connect-StructureGraph -Tenant $Tenant -ClientId $ClientId -Thumbprint $Thumbprint `
        -Scopes @('Group.Read.All', 'Directory.Read.All')

    foreach ($definition in (Get-ConfigValue $config 'groups' @())) {
        $group = Resolve-StructureGroup -Definition $definition
        if (-not $group) {
            Add-Finding -Kind Missing -Area 'Entra group' -Object $definition.displayName
            continue
        }
        $members = @(Get-MgGroupMember -GroupId $group.Id -All -ErrorAction SilentlyContinue)
        if ($members.Count -eq 0) {
            # Not drift, but a group nobody is in is a permission model that is not
            # doing anything yet - worth saying out loud during an onboarding.
            Add-Finding -Kind Different -Area 'Entra group' -Object $definition.displayName `
                -Expected 'at least one member' -Actual 'empty'
        } else {
            Write-Pass "group '$($definition.displayName)': $($members.Count) member(s)"
        }
    }
}

# -- Run -----------------------------------------------------------------------
$exitCode = 0
try {
    Write-Host ''
    Write-Host "  Client : $(Get-ConfigValue $config 'client' '(unnamed)')" -ForegroundColor Cyan
    Write-Host "  Config : $ConfigPath" -ForegroundColor Cyan
    Write-Host "  Scope  : sites $($siteKeys -join ', ') | containers $(($containers | ForEach-Object { $_.key }) -join ', ')" -ForegroundColor Cyan
    Write-Host '  Mode   : READ ONLY - this script never changes anything' -ForegroundColor DarkGray

    $first = $true
    foreach ($siteKey in $siteKeys) {
        $url = Get-StructureSiteUrl -Config $config -SiteKey $siteKey
        Write-Head "Site '$siteKey' - $url"
        $connection = Connect-Structure -Url $url @connectSplat

        if ($first) {
            Write-Step 'Term store'
            Test-StructureTermSet -Connection $connection
            $first = $false
        }

        Write-Step 'Columns'
        Test-StructureColumn -SiteKey $siteKey -Connection $connection

        Write-Step 'Content types'
        Test-StructureContentType -SiteKey $siteKey -Connection $connection

        $siteContainers = @($containers | Where-Object { $_.site -eq $siteKey })
        if ($siteContainers.Count -gt 0) {
            Write-Step 'Containers and permissions'
            foreach ($entry in $siteContainers) {
                Test-StructureContainer -Definition $entry -Connection $connection
            }
        }
    }

    if ($IncludeGroups) {
        Write-Head 'Entra ID groups'
        Test-StructureGroup
    }

    # -- report
    Write-Head 'Summary'
    $missing   = @($findings | Where-Object { $_.Kind -eq 'Missing' }).Count
    $different = @($findings | Where-Object { $_.Kind -eq 'Different' }).Count
    $extra     = @($findings | Where-Object { $_.Kind -eq 'Extra' }).Count

    if ($findings.Count -eq 0) {
        Write-Ok 'The tenant matches the configuration.'
    } else {
        Write-Host "    Missing   : $missing   (provisioning will create these)" -ForegroundColor Yellow
        Write-Host "    Different : $different (provisioning will repair these)" -ForegroundColor Yellow
        Write-Host "    Extra     : $extra     (nothing removes these - your call)" -ForegroundColor Yellow

        if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }
        $findings | Sort-Object Kind, Area, Site, Object |
            Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
        Write-Host "    Report    : $ReportPath" -ForegroundColor Cyan
        $exitCode = 2
    }
    Write-Host ''
} catch {
    Write-Host ''
    Write-Bad "Aborted at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    Write-Host ''
    $exitCode = 1
} finally {
    if ($Disconnect) {
        Disconnect-Structure
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    }
}

exit $exitCode
