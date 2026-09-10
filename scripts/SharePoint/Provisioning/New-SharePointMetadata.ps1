#Requires -Version 7.0
<#
.SYNOPSIS
    Provision the metadata model - managed metadata term set, site columns and
    content types - from a structure configuration file. Idempotent, supports -WhatIf.

.DESCRIPTION
    Step one of the structure set: everything that has to exist before a library can
    use it.

      1. Term set     the managed metadata group, term set and terms behind the
                      Leverancier column. Terms already in the store are left alone,
                      so the list stays extendable from the term store UI without the
                      next run undoing that.
      2. Site columns one per entry in the columns section, created where missing and
                      repaired where they drifted - choices, description, default
                      value and column group.
      3. Content types one per pillar, derived from Document (0x0101), with the right
                      columns attached and the right ones marked required.

    Runs against every site in the configuration's sites section, not just the team
    site. That matters: a Teams private channel lives in its own site collection, so
    the MGMT channel needs its own copy of the columns and content types - a site
    column is not tenant-wide.

    Nothing is deleted. A column that is in the site but not in the config is left
    where it is and reported; removing it would take its data with it. Use
    Test-SharePointStructure.ps1 to see that kind of drift in one list.

    Second and later runs
    ---------------------
    Every step checks before it writes, so a repeat run reports [ OK ] across the
    board and changes nothing. Making a column required after the fact works too -
    the Required flag on an existing field link is updated in place, and pushed down
    to the lists that already use the content type.

    Sign-in
    -------
    Interactive by default (-Interactive, as an admin who may edit the site).
    App-only with a certificate for an unattended run: -ClientId + -Thumbprint.
    The app needs SharePoint Sites.FullControl.All for the content type work, and a
    term store administrator for the term set - app-only can create terms only when
    the app's service principal is a term store admin.

.PARAMETER ConfigPath
    Path to the structure configuration JSON.
    Default: petsolutions.config.json next to this script.

.PARAMETER Site
    Only provision these sites (keys from the configuration's sites section, e.g.
    team, mgmt). Default: all of them.

.PARAMETER Only
    Limit the run to one part: TermSet, Columns or ContentTypes. Default All.
    Handy when a term store admin has to run the term set part separately.

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

.PARAMETER AllowPlaceholders
    Skip the check that the CHANGEME placeholders in the configuration were filled
    in. Only useful for validating a config without connecting.

.PARAMETER Disconnect
    Sign out of PnP when finished.

.EXAMPLE
    # Dry run - what would the metadata model change on this tenant?
    .\New-SharePointMetadata.ps1 -Interactive -ClientId <app-id> -WhatIf

.EXAMPLE
    # Provision columns and content types on both sites
    .\New-SharePointMetadata.ps1 -Interactive -ClientId <app-id>

.EXAMPLE
    # Only the private channel site, after MGMT was moved to its own channel
    .\New-SharePointMetadata.ps1 -Interactive -ClientId <app-id> -Site mgmt

.EXAMPLE
    # Unattended, app-only with a certificate
    .\New-SharePointMetadata.ps1 -ClientId <app-id> -Thumbprint <thumbprint>

.NOTES
    Author  : Sjoerd Kanon
    Requires: PnP.PowerShell 2.x
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $ConfigPath,
    [string[]] $Site,

    [ValidateSet('All', 'TermSet', 'Columns', 'ContentTypes')]
    [string] $Only = 'All',

    [switch] $Interactive,
    [string] $ClientId,
    [string] $Thumbprint,
    [string] $CertificatePath,
    [securestring] $CertificatePassword,
    [string] $Tenant,

    [switch] $AllowPlaceholders,
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
$config = Import-StructureConfig -Path $ConfigPath -AllowPlaceholders:$AllowPlaceholders
if (-not $Tenant) { $Tenant = $config.tenant }

$simulate = [bool] $WhatIfPreference
$siteKeys = if ($Site) { @($Site) } else { @($config.sites.PSObject.Properties.Name) }
foreach ($key in $siteKeys) {
    if ($key -notin $config.sites.PSObject.Properties.Name) {
        throw "Unknown site key '$key'. Known: $($config.sites.PSObject.Properties.Name -join ', ')"
    }
}

$connectSplat = @{
    Tenant              = $Tenant
    ClientId            = $ClientId
    Thumbprint          = $Thumbprint
    CertificatePath     = $CertificatePath
    CertificatePassword = $CertificatePassword
    Interactive         = $Interactive
}

$changeCount = 0

Write-Host ''
Write-Host "  Client : $(Get-ConfigValue $config 'client' '(unnamed)')" -ForegroundColor Cyan
Write-Host "  Config : $ConfigPath" -ForegroundColor Cyan
Write-Host "  Sites  : $($siteKeys -join ', ')" -ForegroundColor Cyan
Write-Host "  Scope  : $Only" -ForegroundColor Cyan
Write-Host "  Mode   : $(if ($simulate) { '-WhatIf - nothing will be changed' } else { 'APPLY' })" -ForegroundColor $(if ($simulate) { 'Yellow' } else { 'Cyan' })

# -- 1. Term set ---------------------------------------------------------------
function Set-StructureTermSet {
    <#
        Term group, term set and terms. Only ever adds - a term that is in the store
        but not in the config stays, because it is almost certainly a leverancier
        someone added on purpose and its GUID is referenced by tagged documents.
    #>
    param($Connection)

    $termStore = Get-ConfigValue $config 'termStore'
    if (-not $termStore) { Write-Skip 'No termStore section in the configuration'; return }

    $group = Get-PnPTermGroup -Identity $termStore.group -Connection $Connection -ErrorAction SilentlyContinue
    if (-not $group) {
        if ($PSCmdlet.ShouldProcess($termStore.group, 'Create term group')) {
            $group = New-PnPTermGroup -Name $termStore.group -Connection $Connection
            Write-Change "term group '$($termStore.group)' created"
            $script:changeCount++
        }
    } else {
        Write-Ok "term group '$($termStore.group)'"
    }
    if (-not $group) { Write-Skip 'Term set and terms need the group - skipped in -WhatIf'; return }

    $set = Get-PnPTermSet -Identity $termStore.termSet -TermGroup $termStore.group -Connection $Connection -ErrorAction SilentlyContinue
    if (-not $set) {
        if ($PSCmdlet.ShouldProcess("$($termStore.group)|$($termStore.termSet)", 'Create term set')) {
            $set = New-PnPTermSet -Name $termStore.termSet -TermGroup $termStore.group `
                -Description (Get-ConfigValue $termStore 'description' '') -Connection $Connection
            Write-Change "term set '$($termStore.termSet)' created"
            $script:changeCount++
        }
    } else {
        Write-Ok "term set '$($termStore.termSet)'"
    }
    if (-not $set) { Write-Skip 'Terms need the term set - skipped in -WhatIf'; return }

    $existing = @(Get-PnPTerm -TermSet $termStore.termSet -TermGroup $termStore.group -Connection $Connection -ErrorAction SilentlyContinue |
                  ForEach-Object { $_.Name })
    foreach ($term in (Get-ConfigValue $termStore 'terms' @())) {
        if ($term -in $existing) { Write-Ok "term '$term'"; continue }
        if ($PSCmdlet.ShouldProcess($term, 'Create term')) {
            New-PnPTerm -Name $term -TermSet $termStore.termSet -TermGroup $termStore.group -Connection $Connection | Out-Null
            Write-Change "term '$term' created"
            $script:changeCount++
        }
    }

    $extra = @($existing | Where-Object { $_ -notin (Get-ConfigValue $termStore 'terms' @()) })
    if ($extra.Count -gt 0) {
        Write-Skip "$($extra.Count) term(s) in the store that the config does not list - left alone: $($extra -join ', ')"
    }
}

# -- 2. Site columns -----------------------------------------------------------
function Set-StructureColumn {
    <#
        One site column, created or repaired. Choices are written through the schema
        XML and pushed down, so a library that already uses the column picks the new
        choices up instead of keeping a stale copy of the list.
    #>
    param($Definition, $Connection)

    $name  = $Definition.internalName
    $field = Get-PnPField -Identity $name -Connection $Connection -ErrorAction SilentlyContinue

    if (-not $field) {
        if (-not $PSCmdlet.ShouldProcess($name, "Create site column ($($Definition.type))")) { return }

        if ($Definition.type -eq 'Taxonomy') {
            $termStore = Get-ConfigValue $config 'termStore'
            $field = Add-PnPTaxonomyField -DisplayName $Definition.displayName -InternalName $name `
                -TermSetPath "$($termStore.group)|$($Definition.termSet)" `
                -Group (Get-ConfigValue $config 'columnGroup' 'Custom') `
                -MultiValue:([bool](Get-ConfigValue $Definition 'multiValue' $false)) `
                -Connection $Connection
        } else {
            $field = Add-PnPField -DisplayName $Definition.displayName -InternalName $name `
                -Type $Definition.type `
                -Group (Get-ConfigValue $config 'columnGroup' 'Custom') `
                -Choices (Get-ConfigValue $Definition 'choices' @()) `
                -Connection $Connection
        }
        Write-Change "column '$($Definition.displayName)' ($name) created"
        $script:changeCount++
    }

    if (-not $field) { return }   # -WhatIf on a column that does not exist yet

    # -- repairs on an existing column
    $updates = @{}
    if ($field.Title -ne $Definition.displayName) { $updates['Title'] = $Definition.displayName }

    $description = Get-ConfigValue $Definition 'description' ''
    if ($field.Description -ne $description) { $updates['Description'] = $description }

    $group = Get-ConfigValue $config 'columnGroup' 'Custom'
    if ($field.Group -ne $group) { $updates['Group'] = $group }

    $default = Get-ConfigValue $Definition 'defaultValue'
    if ($default -and $field.DefaultValue -ne $default) { $updates['DefaultValue'] = $default }

    if ($updates.Count -gt 0) {
        if ($PSCmdlet.ShouldProcess($name, "Update $($updates.Keys -join ', ')")) {
            Set-PnPField -Identity $name -Values $updates -UpdateExistingLists -Connection $Connection | Out-Null
            Write-Change "column '$name' updated: $($updates.Keys -join ', ')"
            $script:changeCount++
        }
    }

    if ($Definition.type -in @('Choice', 'MultiChoice')) {
        $wanted  = @(Get-ConfigValue $Definition 'choices' @())
        $current = Get-FieldChoiceValue -Field $field
        # Compared as an ordered string: the order of the choices is what the user
        # sees in the dropdown, so a reshuffle counts as drift too. Compare-Object
        # would refuse the empty-array case a brand new column can produce.
        if (($current -join "`u{241F}") -ne ($wanted -join "`u{241F}")) {
            if ($PSCmdlet.ShouldProcess($name, "Set choices to: $($wanted -join ', ')")) {
                Set-FieldChoice -Field $field -Choices $wanted -Connection $Connection
                Write-Change "column '$name' choices updated: $($wanted -join ', ')"
                $script:changeCount++
            }
        } else {
            Write-Ok "column '$($Definition.displayName)' ($name)"
        }
    } elseif ($updates.Count -eq 0) {
        Write-Ok "column '$($Definition.displayName)' ($name)"
    }
}

function Set-FieldChoice {
    <#
        Replace the CHOICES block in a field's schema XML and push the change to every
        list that already uses the column. Set-PnPField cannot express a collection
        property, so this is the reliable route.
    #>
    param($Field, [string[]] $Choices, $Connection)

    $schema = [xml] $Field.SchemaXml
    $node   = $schema.Field.SelectSingleNode('CHOICES')
    if (-not $node) {
        $node = $schema.CreateElement('CHOICES')
        $schema.Field.AppendChild($node) | Out-Null
    }
    $node.RemoveAll()
    foreach ($choice in $Choices) {
        $element = $schema.CreateElement('CHOICE')
        $element.InnerText = $choice
        $node.AppendChild($element) | Out-Null
    }

    $Field.SchemaXml = $schema.OuterXml
    # $true: push down to the lists, otherwise only new lists get the new choices.
    $Field.UpdateAndPushChanges($true)
    Invoke-PnPQuery -Connection $Connection
}

# -- 3. Content types ----------------------------------------------------------
function Set-StructureContentType {
    <#
        One content type: created under Document when missing, then its field links
        reconciled - missing columns added, Required brought in line with the config.
        Columns are never unlinked; that would silently drop metadata from documents
        already filed under the type.
    #>
    param($Definition, $Connection)

    $contentType = Get-PnPContentType -Identity $Definition.name -Connection $Connection -ErrorAction SilentlyContinue

    if (-not $contentType) {
        if (-not $PSCmdlet.ShouldProcess($Definition.name, 'Create content type')) { return }

        $parent = Get-PnPContentType -Identity '0x0101' -Connection $Connection
        $addSplat = @{
            Name              = $Definition.name
            Description       = (Get-ConfigValue $Definition 'description' '')
            Group             = (Get-ConfigValue $config 'contentTypeGroup' 'Custom')
            ParentContentType = $parent
            Connection        = $Connection
        }
        # A fixed ID keeps the structure reproducible across tenants, which is what
        # makes the drift check able to say "this is the same content type".
        $id = Get-ConfigValue $Definition 'id'
        if ($id) { $addSplat['ContentTypeId'] = $id }

        $contentType = Add-PnPContentType @addSplat
        Write-Change "content type '$($Definition.name)' created"
        $script:changeCount++
    }

    if (-not $contentType) { return }

    $context = Get-PnPContext -Connection $Connection
    $context.Load($contentType.FieldLinks)
    Invoke-PnPQuery -Connection $Connection
    $linked = @($contentType.FieldLinks | ForEach-Object { $_.Name })

    $touched = @()
    foreach ($fieldDef in $Definition.fields) {
        $required = [bool] (Get-ConfigValue $fieldDef 'required' $false)

        if ($fieldDef.internalName -notin $linked) {
            if ($PSCmdlet.ShouldProcess("$($Definition.name) / $($fieldDef.internalName)", "Add column (required: $required)")) {
                Add-PnPFieldToContentType -Field $fieldDef.internalName -ContentType $contentType `
                    -Required:$required -Connection $Connection
                $touched += "+$($fieldDef.internalName)$(if ($required) { ' (required)' })"
                $script:changeCount++
            }
            continue
        }

        $wouldChange = Set-FieldLinkRequired -ContentType $contentType -InternalName $fieldDef.internalName `
            -Required $required -Connection $Connection -WhatIfMode:$simulate
        if ($wouldChange) {
            if ($simulate -or $PSCmdlet.ShouldProcess("$($Definition.name) / $($fieldDef.internalName)", "Set required to $required")) {
                $touched += "$($fieldDef.internalName) required=$required"
                $script:changeCount++
            }
        }
    }

    if ($touched.Count -gt 0) {
        Write-Change "content type '$($Definition.name)': $($touched -join ', ')"
    } else {
        Write-Ok "content type '$($Definition.name)'"
    }
}

# -- Run -----------------------------------------------------------------------
try {
    $first = $true
    foreach ($key in $siteKeys) {
        $url = Get-StructureSiteUrl -Config $config -SiteKey $key
        Write-Head "Site '$key' - $url"
        $connection = Connect-Structure -Url $url @connectSplat

        # The term store is tenant-wide, so it is done once - on the first site we
        # happen to connect to.
        if ($first -and $Only -in @('All', 'TermSet')) {
            Write-Step '1. Managed metadata'
            Set-StructureTermSet -Connection $connection
        }
        $first = $false

        if ($Only -in @('All', 'Columns')) {
            Write-Step '2. Site columns'
            foreach ($column in $config.columns) {
                Set-StructureColumn -Definition $column -Connection $connection
            }
        }

        if ($Only -in @('All', 'ContentTypes')) {
            Write-Step '3. Content types'
            foreach ($contentType in $config.contentTypes) {
                Set-StructureContentType -Definition $contentType -Connection $connection
            }
        }
    }

    Write-Head 'Summary'
    if ($simulate) {
        Write-Host "    $changeCount change(s) would be made - rerun without -WhatIf to apply." -ForegroundColor Yellow
    } elseif ($changeCount -eq 0) {
        Write-Ok 'Metadata model already matches the configuration - nothing changed.'
    } else {
        Write-Change "$changeCount change(s) applied."
        Write-Host '    Next: Set-SharePointLibraries.ps1 to bind these to the libraries and set the permissions.' -ForegroundColor DarkGray
    }
    Write-Host ''
} finally {
    if ($Disconnect) { Disconnect-Structure; Write-Host '  Disconnected.' -ForegroundColor DarkGray }
}
