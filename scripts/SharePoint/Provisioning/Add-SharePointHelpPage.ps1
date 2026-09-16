#Requires -Version 7.0
<#
.SYNOPSIS
    Put the end-user explanation on the team site itself, as a SharePoint page built
    from the same configuration the structure was built from. Supports -WhatIf.

.DESCRIPTION
    A handleiding in a repo is read by nobody. This writes it where the people who
    upload files already are: a page on the team site, linked from the left-hand
    navigation.

    The page is generated from the configuration, not typed out, so it cannot drift
    from what the libraries actually do. The pillars it lists are the pillars that
    exist; the labels it explains are the labels on the columns, with the same help
    text that appears under each field in the upload form; the required fields per
    document type are read off the content types.

    Rename a channel or add a brand, rerun this, and the page says the new thing.

    What goes on it
    ---------------
        1. The idea      labels instead of folders, and what "Beide" buys you
        2. Where to put  a line per channel and library
        3. What to fill  every label, its choices and its help text, and which
                         document type insists on which
        4. Uploading     the two ways of adding a file and why they differ
        5. Finding back  the views that exist, named as they are named
        6. Getting help  whoever you tell it to be

    Written for the person uploading a catalogue, not for whoever maintains this. No
    group names, no column internal names, no talk of content types or site columns.
    Two things are deliberately left out of the generated text as well: the note on
    each container, which names security groups, and the description on each view,
    which talks about pillars and synced folders. Their titles are plain enough.

    Nothing about permissions is on it either. Who may see what is not a thing users
    can act on, and a page that explains it invites the question of why they cannot.

.PARAMETER ConfigPath
    Path to the structure configuration JSON. Default: the single *.config.json next
    to this script that no longer contains the CHANGEME placeholders.

.PARAMETER Title
    Title of the page. Default: "Zo werken we hier met bestanden".

.PARAMETER Name
    File name of the page, without .aspx. Default: derived from the title.

.PARAMETER Contact
    Who to contact, printed at the bottom. Default: "de IT-servicedesk".

.PARAMETER SkipNavigation
    Do not add a link to the site's left-hand navigation.

.PARAMETER Force
    Rewrite the page if it already exists. Without this an existing page is left
    alone - somebody may have edited it.

.PARAMETER Interactive
    Sign in interactively (the default).

.PARAMETER ClientId
    Client ID of the Entra app registration used to sign in.

.PARAMETER Tenant
    Tenant name or ID. Defaults to the tenant in the configuration file.

.EXAMPLE
    # What would the page say?
    .\Add-SharePointHelpPage.ps1 -WhatIf

.EXAMPLE
    # Put it on the team site and link it in the navigation
    .\Add-SharePointHelpPage.ps1 -Interactive -ClientId <app-id>

.EXAMPLE
    # Rewrite it after the model changed
    .\Add-SharePointHelpPage.ps1 -Interactive -ClientId <app-id> -Force

.NOTES
    Author  : Sjoerd Kanon
    Requires: PnP.PowerShell 2.x
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $ConfigPath,
    [string] $Title = 'Zo werken we hier met bestanden',
    [string] $Name,
    [string] $Contact = 'de IT-servicedesk',
    [switch] $SkipNavigation,
    [switch] $Force,

    [switch] $Interactive,
    [string] $ClientId,
    [string] $Tenant
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

. (Join-Path $PSScriptRoot 'SharePointStructure.Common.ps1')

Assert-PnPModule

if (-not $ConfigPath) {
    $candidates = @(Get-ChildItem -Path $PSScriptRoot -Filter '*.config.json' -ErrorAction SilentlyContinue |
                    Where-Object { (Get-Content $_.FullName -Raw) -notmatch 'CHANGEME' })
    if ($candidates.Count -eq 1) { $ConfigPath = $candidates[0].FullName }
    elseif ($candidates.Count -eq 0) { throw 'No configuration found - run New-StructureConfig.ps1 first.' }
    else { throw ("More than one configuration here - pass -ConfigPath. Found: {0}" -f (($candidates | ForEach-Object { $_.Name }) -join ', ')) }
}

$config = Import-StructureConfig -Path $ConfigPath
if (-not $Tenant) { $Tenant = $config.tenant }
if (-not $Name)   { $Name   = ($Title -replace '[^a-zA-Z0-9]+', '-').Trim('-') }

$simulate = [bool] $WhatIfPreference

function Get-Html {
    <# HTML-safe text: a brand called "Smith & Co" must not break the page. #>
    param([string] $Text)
    return ($Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;')
}

function Get-FieldRole {
    param([string] $Role, [string] $Default)
    $value = Get-ConfigValue (Get-ConfigValue $config 'fieldRoles') $Role
    if ($value) { return $value }
    return $Default
}

$brandField = Get-FieldRole 'brand' 'PsMerk'
$brandColumn = @($config.columns | Where-Object { $_.internalName -eq $brandField }) | Select-Object -First 1
$brands      = @(Get-ConfigValue $brandColumn 'choices' @())
# The last choice of a brand column is the "belongs to all of them" one by convention
# - it is what the wizard appends. With one brand there is no such thing.
$bothName    = if ($brands.Count -gt 2) { $brands[-1] } else { $null }

# -- The sections --------------------------------------------------------------
function New-IntroSection {
    $html = "<h2>Geen mappen meer &ndash; labels doen het werk</h2>"
    $html += "<p>Vroeger bepaalde de map waar iets stond wat het was. Nu doen <strong>labels</strong> dat: "
    $html += "extra kenmerken naast de bestandsnaam, die je als kolommen naast elkaar ziet staan.</p>"
    if ($bothName) {
        $html += "<p>Het grote voordeel: <strong>&eacute;&eacute;n bestand kan in meerdere overzichten opduiken.</strong> "
        $html += "Hoort iets bij $((Get-Html ($brands[0..($brands.Count-2)] -join ' en '))), kies dan <strong>$(Get-Html $bothName)</strong>. "
        $html += "Het staat &eacute;&eacute;n keer opgeslagen en verschijnt in beide merkoverzichten. "
        $html += "Geen twee kopie&euml;n die uit elkaar gaan lopen.</p>"
    }
    $html += "<p><strong>Maak geen submappen aan.</strong> Dat hoeft niet meer, en het maakt terugvinden juist moeilijker.</p>"
    return $html
}

function New-WhereSection {
    $html = "<h2>Waar zet ik iets neer?</h2><p>Kies de plek die erbij hoort, verder niets:</p><ul>"
    foreach ($entry in $config.containers) {
        # Never the config's note field: that is written for whoever maintains this
        # and mentions group names. Nothing on this page should name a group.
        $where = switch (Get-ConfigValue $entry 'kind' 'Library') {
            'Library' { 'eigen plek met bestanden' }
            default   { 'kanaal in Teams' }
        }
        $html += "<li><strong>$(Get-Html $entry.title)</strong> &ndash; $where</li>"
    }
    $html += '</ul>'
    return $html
}

function New-LabelSection {
    $html = "<h2>Wat vul je in?</h2>"
    $html += "<p>Onder elk veld in het formulier staat dezelfde uitleg als hieronder.</p><ul>"
    foreach ($column in $config.columns) {
        $line = "<li><strong>$(Get-Html $column.displayName)</strong>"
        $choices = @(Get-ConfigValue $column 'choices' @())
        if ($choices.Count -gt 0) { $line += " &ndash; $(Get-Html ($choices -join ' / '))" }
        elseif ($column.type -eq 'Taxonomy') { $line += ' &ndash; kies uit de lijst, typ de eerste letters' }
        $description = Get-ConfigValue $column 'description'
        if ($description) { $line += "<br />$(Get-Html $description)" }
        $html += "$line</li>"
    }
    $html += '</ul>'

    $html += "<h3>Wat moet je invullen?</h3><p>Dat hangt af van het soort document:</p><ul>"
    foreach ($contentType in $config.contentTypes) {
        # Only the labels people see - the internal column names never appear here.
        $labels = @()
        foreach ($field in $contentType.fields) {
            if (-not (Get-ConfigValue $field 'required' $false)) { continue }
            $column = @($config.columns | Where-Object { $_.internalName -eq $field.internalName }) | Select-Object -First 1
            if ($column) { $labels += $column.displayName }
        }
        $html += "<li><strong>$(Get-Html $contentType.name)</strong> &ndash; $(Get-Html ($labels -join ', '))</li>"
    }
    $html += '</ul>'
    return $html
}

function New-UploadSection {
    $html = "<h2>Een bestand toevoegen</h2>"
    $html += "<p>Er zijn twee manieren, en ze gedragen zich <strong>niet</strong> hetzelfde.</p>"
    $html += "<h3>1. Nieuw aanmaken of uploaden met de knop</h3>"
    $html += "<p>In het kanaal, tabblad Bestanden, via <strong>+ Nieuw</strong> of <strong>Uploaden</strong>. "
    $html += "Je krijgt meteen de vragen te zien en je bent klaar. Dit is de goede manier.</p>"
    $html += "<h3>2. Bestanden erin slepen, of ze in de map op je eigen computer zetten</h3>"
    $html += "<p>Dan wordt er <strong>niets gevraagd</strong> en komen de bestanden binnen zonder labels. "
    $html += "Ze zijn niet weg, maar niemand vindt ze nog terug. Je herkent ze aan de melding "
    $html += "<em>Vereiste info</em> achter de bestandsnaam.</p>"
    $html += "<p><strong>Zo los je dat op:</strong> kies bovenaan de weergave <em>Nog te taggen</em>. "
    $html += "Daar staan ze bij elkaar. Vink er meerdere aan, klik rechtsboven op het rondje met de <em>i</em>, "
    $html += "en vul de labels in &eacute;&eacute;n keer voor allemaal in.</p>"
    $html += "<p>Het makkelijkst blijft: toevoegen doe je in de browser of in Teams. "
    $html += "De map op je eigen computer is handig om te lezen en om offline te werken.</p>"
    $html += "<h3>Wat gebeurt er als je een label geeft?</h3><ul>"
    $html += "<li>Het bestand <strong>verhuist niet</strong>. Gedeelde links blijven werken.</li>"
    $html += "<li>De inhoud verandert niet &ndash; een label staat n&aacute;&aacute;st het bestand.</li>"
    $html += "<li>Overzichten passen zich meteen aan. De zoekbalk loopt een paar minuten achter.</li>"
    $html += "<li>Verkeerd gelabeld? Verander het gewoon. Er gaat niets verloren.</li>"
    $html += "</ul>"
    $html += "<p><strong>Let op: een label is geen slot.</strong> Een bestand als vertrouwelijk markeren "
    $html += "sluit niemand buiten &ndash; het is een afspraak, geen beveiliging. Wie op deze plek mag komen, "
    $html += "ziet het bestand nog steeds. Wie waar bij mag, regelt $(Get-Html $Contact).</p>"
    return $html
}

function New-FindSection {
    $html = "<h2>Iets terugvinden</h2>"
    $html += "<p>Bovenaan de bestandenlijst staat een menu met weergaven. Deze staan klaar:</p><ul>"

    foreach ($entry in $config.containers) {
        $view = Get-ConfigValue $entry 'view'
        if (-not $view) { continue }
        $title = if ((Get-ConfigValue $entry 'kind' 'Library') -eq 'ChannelFolder') { "$($entry.title) - $($view.title)" } else { $view.title }
        $html += "<li><strong>$(Get-Html $title)</strong> &ndash; in $(Get-Html $entry.title)</li>"
    }
    foreach ($group in (Get-ConfigValue $config 'libraryViews' @())) {
        foreach ($view in $group.views) {
            # Only the title. The description in the configuration is written for
            # whoever maintains this and talks about pillars and synced folders -
            # words this page has no business using.
            $html += "<li><strong>$(Get-Html $view.title)</strong> &ndash; over alles heen</li>"
        }
    }
    $html += '</ul>'
    $html += "<p>Of klik rechtsboven op het filterpictogram en vink meerdere labels tegelijk aan. "
    $html += "De zoekbalk zoekt ook <em>in</em> documenten, niet alleen in bestandsnamen.</p>"
    return $html
}

function New-HelpSection {
    return "<h2>Hulp nodig?</h2><p>Neem contact op met $(Get-Html $Contact). " +
           'Vermeld het kanaal en de bestandsnaam &ndash; dan is meteen duidelijk waar het over gaat.</p>'
}

# -- Run -----------------------------------------------------------------------
try {
    Write-Host ''
    Write-Host "  Client : $(Get-ConfigValue $config 'client' '(unnamed)')" -ForegroundColor Cyan
    Write-Host "  Page   : $Title ($Name.aspx)" -ForegroundColor Cyan
    Write-Host "  Mode   : $(if ($simulate) { '-WhatIf - nothing will be written' } else { 'APPLY' })" -ForegroundColor Cyan

    $sections = [ordered]@{
        'De bedoeling'   = New-IntroSection
        'Waar wat hoort' = New-WhereSection
        'De labels'      = New-LabelSection
        'Toevoegen'      = New-UploadSection
        'Terugvinden'    = New-FindSection
        'Hulp'           = New-HelpSection
    }

    Write-Head 'Sections'
    foreach ($key in $sections.Keys) {
        Write-Ok "$key ($($sections[$key].Length) tekens)"
    }

    if ($simulate) {
        Write-Head 'Summary'
        Write-Skip "Would create the page '$Title' on $(Get-StructureSiteUrl -Config $config -SiteKey 'team')"
        Write-Host '    Rerun without -WhatIf to write it.' -ForegroundColor Yellow
        Write-Host ''
        exit 0
    }

    $url = Get-StructureSiteUrl -Config $config -SiteKey 'team'
    Write-Head "Site - $url"
    $connection = Connect-Structure -Url $url -Tenant $Tenant -ClientId $ClientId -Interactive:$Interactive

    $existing = Get-PnPPage -Identity $Name -Connection $connection -ErrorAction SilentlyContinue
    if ($existing -and -not $Force) {
        Write-Warn "The page '$Name' already exists - left alone. Pass -Force to rewrite it."
        Write-Host ''
        exit 0
    }

    if (-not $PSCmdlet.ShouldProcess($Title, 'Create or rewrite the help page')) { exit 0 }

    # Overwrites when it exists, which with -Force is what was asked for.
    $page = Add-PnPPage -Name $Name -Title $Title -LayoutType Article -Connection $connection
    Write-Change "page '$Title' created"

    foreach ($key in $sections.Keys) {
        Add-PnPPageTextPart -Page $Name -Text $sections[$key] -Connection $connection | Out-Null
        Write-Ok "section '$key' added"
    }

    Set-PnPPage -Identity $Name -Publish -Connection $connection | Out-Null
    Write-Change 'page published'

    $pageUrl = "$url/SitePages/$Name.aspx"

    if (-not $SkipNavigation) {
        $nodes = @(Get-PnPNavigationNode -Location QuickLaunch -Connection $connection -ErrorAction SilentlyContinue)
        if ($nodes | Where-Object { (Get-ConfigValue $_ 'Title') -eq $Title }) {
            Write-Ok 'already in the navigation'
        } else {
            Add-PnPNavigationNode -Location QuickLaunch -Title $Title -Url $pageUrl -Connection $connection | Out-Null
            Write-Change 'added to the left-hand navigation'
        }
    }

    Write-Head 'Summary'
    Write-Ok "The explanation is on the site: $pageUrl"
    Write-Host '    Point people at it, or pin it as a tab in the General channel.' -ForegroundColor DarkGray
    Write-Host ''
} catch {
    Write-Host ''
    Write-Bad "Aborted at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    if ($_.InvocationInfo.Line) { Write-Host "    $($_.InvocationInfo.Line.Trim())" -ForegroundColor DarkRed }
    Write-Host ''
    exit 1
}

exit 0
