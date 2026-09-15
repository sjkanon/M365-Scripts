#Requires -Version 7.0
<#
.SYNOPSIS
    Ask what everything should be called and write the structure configuration file.
    No JSON editing - answer the questions, or press Enter to take the suggestion.

.DESCRIPTION
    The configuration file is what the other scripts read, what the drift check
    compares against, and what makes a second client a second file instead of a second
    fork. But nobody should have to write it by hand, so this asks instead.

    You are asked about the things that differ per client:

        the tenant, the team and its owner
        the brands
        the pillars, and which of them is private
        the prefix for the security groups
        the languages, regions, document kinds and starting suppliers
        whether you want the share-status column maintained

    Everything else is derived. Per pillar you get a channel, a content type, two
    security groups and a grouped view; per brand you get a view that spans every
    pillar.

    -All asks for those derived names too - the channel name, the folder inside the
    library, the content type, both group names, the view title, the library behind
    the channels, the column and content type groups, the term set, and the label
    each column carries for the user. Nothing about the finished structure is then
    decided behind your back.

    The column internal names stay fixed either way. They are never shown to anyone,
    and changing one after documents carry it loses the data on those documents.

    Press Enter at any question to accept the suggestion in brackets. Lists are
    comma-separated. Where a question is optional the hint says (of "geen") - type
    that to turn the suggestion down, because Enter means "take it".

    What you cannot change later
    ----------------------------
    Display names, channel names and group names can all be changed afterwards.
    Two things cannot, because SharePoint keys data to them: the column internal
    names, and the content type IDs. Those are generated once and then left alone,
    which is why this script refuses to overwrite a configuration that already exists
    unless you pass -Force.

.PARAMETER Path
    Where to write the configuration.
    Default: <client>.config.json next to this script, named after the client.

.PARAMETER All
    Ask for every name, including the ones that are otherwise derived: channel,
    folder, content type, group and view names per pillar, plus the library, column
    group, content type group, term set and the label of every column.

.PARAMETER Force
    Overwrite an existing configuration file. Generates new content type IDs, so only
    do this for a tenant where nothing has been provisioned yet.

.PARAMETER ShowJson
    Print the configuration instead of writing it.

.EXAMPLE
    # Answer the questions, get a configuration, then build
    .\New-StructureConfig.ps1
    .\Install-SharePointStructure.ps1 -ConfigPath .\petsolutions-nv.config.json

.EXAMPLE
    # Decide every name yourself, nothing derived
    .\New-StructureConfig.ps1 -All

.EXAMPLE
    # See what it would write without writing it
    .\New-StructureConfig.ps1 -ShowJson

.NOTES
    Author  : Sjoerd Kanon
#>
[CmdletBinding()]
param(
    [string] $Path,
    [switch] $All,
    [switch] $Force,
    [switch] $ShowJson
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'SharePointStructure.Common.ps1')

# -- Asking --------------------------------------------------------------------
function Read-Value {
    param([string] $Question, [string] $Default, [switch] $AllowEmpty)

    while ($true) {
        $hint = if ($Default) { " [$Default]" } else { '' }
        if ($AllowEmpty) { $hint += ' (of "geen")' }
        Write-Host "  $Question$hint" -NoNewline -ForegroundColor Cyan
        Write-Host ': ' -NoNewline
        $answer = Read-Host

        # Enter means "take the suggestion", so an optional question needs a word for
        # "actually, none" - otherwise the default can never be turned down.
        if ($AllowEmpty -and $answer -match '^\s*(geen|none|nee|-)\s*$') { return '' }

        if (-not $answer) { $answer = $Default }
        if ($answer -or $AllowEmpty) { return $answer }
        Write-Host '    Dit is verplicht.' -ForegroundColor Yellow
    }
}

function Read-List {
    param([string] $Question, [string[]] $Default, [switch] $AllowEmpty)

    $answer = Read-Value -Question "$Question (komma's ertussen)" -Default ($Default -join ', ') -AllowEmpty:$AllowEmpty
    $list   = @($answer -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    # The comma keeps a one-item list a list: PowerShell unrolls a single-element
    # array on return, and one brand would come back as a bare string.
    return ,$list
}

function Read-YesNo {
    param([string] $Question, [bool] $Default = $false)

    $hint   = if ($Default) { 'J/n' } else { 'j/N' }
    Write-Host "  $Question [$hint]" -NoNewline -ForegroundColor Cyan
    Write-Host ': ' -NoNewline
    $answer = Read-Host
    if (-not $answer) { return $Default }
    return $answer -match '^[jJyY]'
}

function Read-Detail {
    <#
        A name that has a sensible derivation. Without -All it is taken silently;
        with -All it becomes a question like any other, so nothing about the finished
        structure is decided behind the operator's back.
    #>
    param([string] $Question, [string] $Default, [switch] $AllowEmpty)

    if (-not $All) { return $Default }
    return Read-Value -Question $Question -Default $Default -AllowEmpty:$AllowEmpty
}

function Get-Slug {
    <# A name SharePoint and Entra both accept: letters and digits, nothing else. #>
    param([string] $Text)
    return (($Text -replace '[^a-zA-Z0-9]', '')).ToLowerInvariant()
}

function New-ContentTypeId {
    <#
        A child of Document (0x0101) is 0x0101 + 00 + a GUID without dashes. Generated
        once and then fixed: it is how the drift check knows two content types on two
        tenants are the same one.
    #>
    return '0x010100' + ([guid]::NewGuid().ToString('N').ToUpperInvariant())
}

Write-Host ''
Write-Host '  ┌────────────────────────────────────────────────────────────┐' -ForegroundColor Cyan
Write-Host '  │  Hoe moet alles heten?                                     │' -ForegroundColor Cyan
Write-Host '  └────────────────────────────────────────────────────────────┘' -ForegroundColor Cyan
Write-Host '  Enter = de waarde tussen haakjes overnemen.' -ForegroundColor DarkGray
if ($All) {
    Write-Host '  -All: elke naam wordt gevraagd, ook de afleidbare.' -ForegroundColor DarkGray
} else {
    Write-Host '  Tip: -All vraagt ook naar kanaal-, map-, documenttype- en groepsnamen.' -ForegroundColor DarkGray
}
Write-Host ''

# -- 1. Client and tenant ------------------------------------------------------
Write-Host '  ── De klant ──────────────────────────────────────────────' -ForegroundColor DarkGray
$client = Read-Value -Question 'Naam van de klant' -Default 'Petsolutions NV'
$tenant = Read-Value -Question 'Tenant (bv. petsolutions.onmicrosoft.com)'
if ($tenant -notmatch '\.') { $tenant = "$tenant.onmicrosoft.com" }

$shortName = Get-Slug ($client -split '\s')[0]

# -- 2. The team ---------------------------------------------------------------
Write-Host ''
Write-Host '  ── Het Team ──────────────────────────────────────────────' -ForegroundColor DarkGray
$teamName  = Read-Value -Question 'Naam van het Team, zoals het in Teams komt te staan' -Default (($client -split '\s')[0])
$teamAlias = Read-Value -Question 'Alias - bepaalt ook de site-URL (.../sites/<alias>)' -Default (Get-Slug $teamName)
$teamOwner = Read-Value -Question 'Eigenaar van het Team (UPN)'

# -- 3. Brands -----------------------------------------------------------------
Write-Host ''
Write-Host '  ── De merken ─────────────────────────────────────────────' -ForegroundColor DarkGray
Write-Host '  Merken worden een label, geen aparte site. Er komt automatisch' -ForegroundColor DarkGray
Write-Host '  een keuze "Beide" bij, plus een overzicht per merk.' -ForegroundColor DarkGray
$brands   = Read-List -Question 'Merken' -Default @('Butterstone', 'Laseto')
$bothName = if ($brands.Count -gt 1) { Read-Value -Question 'Hoe heet "hoort bij allemaal"' -Default 'Beide' } else { $null }

# -- 4. Pillars ----------------------------------------------------------------
Write-Host ''
Write-Host '  ── De pijlers ────────────────────────────────────────────' -ForegroundColor DarkGray
Write-Host '  Elke pijler wordt een kanaal, een documenttype, twee groepen' -ForegroundColor DarkGray
Write-Host '  en een gegroepeerde weergave.' -ForegroundColor DarkGray
$pillars = Read-List -Question 'Pijlers' -Default @('MGMT', 'Leveranciers', 'Verkopers', 'Klanten', 'Marketing', 'TD')

$privatePillars = Read-List -Question 'Welke daarvan zijn een privekanaal (eigen site, alleen leden)' -Default @($pillars[0]) -AllowEmpty
foreach ($p in $privatePillars) {
    if ($p -notin $pillars) { throw "'$p' staat niet in de pijlerlijst." }
}

$supplierPillar = Read-Value -Question 'Welke pijler gaat over leveranciers' -Default 'Leveranciers' -AllowEmpty
$salesPillar    = Read-Value -Question 'Welke pijler gaat over verkoop, met regios' -Default 'Verkopers' -AllowEmpty

# -- 5. External library -------------------------------------------------------
Write-Host ''
Write-Host '  ── Bibliotheek voor klanten ──────────────────────────────' -ForegroundColor DarkGray
$extLibrary = Read-Value -Question 'Naam van de aparte bibliotheek voor klanten' -Default 'FUTECH Images and videos' -AllowEmpty

# -- 6. Groups -----------------------------------------------------------------
Write-Host ''
Write-Host '  ── De groepen ────────────────────────────────────────────' -ForegroundColor DarkGray
$groupPrefix = Read-Value -Question 'Voorvoegsel voor de securitygroepen' -Default ("SG-" + $shortName.ToUpperInvariant())
$rwSuffix    = Read-Value -Question 'Achtervoegsel voor bewerken'    -Default 'RW'
$roSuffix    = Read-Value -Question 'Achtervoegsel voor alleen lezen' -Default 'RO'

# -- 7. The labels -------------------------------------------------------------
Write-Host ''
Write-Host '  ── De labels ─────────────────────────────────────────────' -ForegroundColor DarkGray
$languages  = Read-List -Question 'Talen' -Default @('NL', 'FR', 'DE', 'EN', 'Geen taal')
# Assigned before the if: an if-block whose only output is @() yields $null, not an
# empty array, and everything downstream asks these for their Count.
$regions = @()
if ($salesPillar) { $regions = Read-List -Question 'Verkoopregios' -Default @('Benelux', 'Duitsland', 'Frankrijk', 'Export') -AllowEmpty }
$docKinds   = Read-List -Question 'Soorten document' -Default @('Catalogus', 'Prijslijst', 'Schrijfrichtlijn', 'Afbeelding+certificaat', 'Marketingslag')
$confLevels = Read-List -Question 'Vertrouwelijkheidsniveaus' -Default @('Intern', 'Deelbaar met klant', 'Vertrouwelijk')
$lifecycle  = Read-List -Question 'Statuswaarden' -Default @('Actief', 'Te archiveren', 'Verouderd')
$suppliers = @()
if ($supplierPillar) { $suppliers = Read-List -Question 'Leveranciers om mee te beginnen (later uitbreidbaar)' -Default @('Lev. 1', 'Lev. 2', 'Lev. 3') -AllowEmpty }

# -- 8. The one real choice ----------------------------------------------------
Write-Host ''
Write-Host '  ── Onderhoud ─────────────────────────────────────────────' -ForegroundColor DarkGray
Write-Host '  De deelstatus-kolom houdt per bestand bij of het buiten de' -ForegroundColor DarkGray
Write-Host '  organisatie open staat. Dat is het enige onderdeel dat een' -ForegroundColor DarkGray
Write-Host '  script nodig heeft dat elke nacht draait.' -ForegroundColor DarkGray
$wantShareStatus = Read-YesNo -Question 'Deelstatus bijhouden' -Default $false

$tightenChannels = Read-YesNo -Question 'Rechten per pijler afdwingen op de kanaalmappen (niet ondersteund door Microsoft)' -Default $false

# -- 9. Everything that is otherwise derived -----------------------------------
$tenantName = ($tenant -split '\.')[0]
$teamSite   = "https://$tenantName.sharepoint.com/sites/$teamAlias"

if ($All) {
    Write-Host ''
    Write-Host '  ── Namen die anders worden afgeleid ──────────────────────' -ForegroundColor DarkGray
}
$teamSite     = (Read-Detail -Question 'URL van de teamsite' -Default $teamSite).TrimEnd('/')
$channelList  = Read-Detail -Question 'Naam van de bibliotheek achter de kanalen' -Default 'Documents'
$columnGroup  = Read-Detail -Question 'Groepsnaam voor de sitekolommen (in de kolommenlijst)' -Default $teamName
$ctGroup      = Read-Detail -Question 'Groepsnaam voor de documenttypes' -Default $teamName
$termSetName  = if ($suppliers.Count -gt 0) { Read-Detail -Question 'Naam van de leveranciers-termenset' -Default 'Leveranciers' } else { 'Leveranciers' }

# Column display names. The internal names stay fixed - they are never shown and
# changing one after documents carry it loses the data on those documents.
if ($All) {
    Write-Host ''
    Write-Host '  ── Hoe de labels heten voor de gebruiker ─────────────────' -ForegroundColor DarkGray
}
$labelBrand = Read-Detail -Question 'Label voor het merk'            -Default 'Merk'
$labelPil   = Read-Detail -Question 'Label voor de pijler'           -Default 'Pijler'
$labelReg   = Read-Detail -Question 'Label voor de regio'            -Default 'Regio'
$labelSup   = Read-Detail -Question 'Label voor de leverancier'      -Default 'Leverancier'
$labelLang  = Read-Detail -Question 'Label voor de taal'             -Default 'Taal'
$labelKind  = Read-Detail -Question 'Label voor het soort document'  -Default 'Contenttype'
$labelConf  = Read-Detail -Question 'Label voor de vertrouwelijkheid' -Default 'Vertrouwelijkheid'
$labelShare = Read-Detail -Question 'Label voor de deelstatus'       -Default 'Deelstatus'
$labelLife  = Read-Detail -Question 'Label voor de status'           -Default 'Status'

# Per pillar: the channel, its folder, its content type, its two groups and its view.
$detail = [ordered]@{}
foreach ($pillar in $pillars) {
    if ($All) {
        Write-Host ''
        Write-Host "  ── Pijler $pillar ─────────────────────────────────" -ForegroundColor DarkGray
    }
    $defaultGroupBy = if ($pillar -eq $supplierPillar -and $suppliers.Count -gt 0) { $labelSup }
                      elseif ($pillar -eq $salesPillar -and $regions.Count -gt 0)  { $labelReg }
                      else                                                          { $labelKind }
    $groupByField   = switch ($defaultGroupBy) {
        $labelSup  { 'PsLeverancier' }
        $labelReg  { 'PsRegio' }
        default    { 'PsContenttype' }
    }

    $channel = Read-Detail -Question "Kanaalnaam"                  -Default $pillar
    $detail[$pillar] = [ordered]@{
        Channel     = $channel
        Folder      = (Read-Detail -Question 'Mapnaam in de bibliotheek'  -Default $channel)
        ContentType = (Read-Detail -Question 'Naam van het documenttype'  -Default "$pillar-document")
        GroupRw     = (Read-Detail -Question 'Groep die mag bewerken'     -Default "$groupPrefix-$pillar-$rwSuffix")
        GroupRo     = (Read-Detail -Question 'Groep die mag lezen'        -Default "$groupPrefix-$pillar-$roSuffix")
        ViewTitle   = (Read-Detail -Question 'Naam van de weergave'       -Default "Op $($defaultGroupBy.ToLowerInvariant())")
        GroupBy     = $groupByField
    }
}

$extGroupName = ''
$extCtName    = ''
if ($extLibrary) {
    if ($All) {
        Write-Host ''
        Write-Host "  ── Klantbibliotheek ──────────────────────────────" -ForegroundColor DarkGray
    }
    $extGroupName = Read-Detail -Question 'Groep voor de externe klanten' -Default "$groupPrefix-Klanten-Extern"
    $extCtName    = Read-Detail -Question 'Naam van het documenttype'     -Default 'Klantmedia'
}

# -- Build ---------------------------------------------------------------------

$brandChoices = @($brands)
if ($bothName) { $brandChoices += $bothName }

$columns = [System.Collections.Generic.List[object]]::new()
$columns.Add([ordered]@{ internalName = 'PsMerk'; displayName = $labelBrand; type = 'Choice'
    description = "Voor welk merk is dit?$(if ($bothName) { " Hoort het bij allebei, kies $bothName - dan verschijnt het in beide merkoverzichten. Geen kopie nodig." })"
    choices = $brandChoices })
$columns.Add([ordered]@{ internalName = 'PsPijler'; displayName = $labelPil; type = 'Choice'
    description = 'Vult zichzelf in op basis van het kanaal. Je hoeft hier niets te doen.'
    choices = @($pillars) })
if ($regions.Count -gt 0) {
    $columns.Add([ordered]@{ internalName = 'PsRegio'; displayName = $labelReg; type = 'Choice'
        description = 'Voor welke verkoopregio is dit bedoeld?'; choices = @($regions) })
}
if ($suppliers.Count -gt 0) {
    $columns.Add([ordered]@{ internalName = 'PsLeverancier'; displayName = $labelSup; type = 'Taxonomy'
        description = 'Typ de eerste letters van de leverancier, dan vult hij aan. Staat hij er niet bij? Vraag IT om hem toe te voegen.'
        termSet = $termSetName; multiValue = $false })
}
$columns.Add([ordered]@{ internalName = 'PsTaal'; displayName = $labelLang; type = 'MultiChoice'
    description = 'In welke taal of talen is dit document? Meerdere aanvinken mag.'
    choices = @($languages) })
$columns.Add([ordered]@{ internalName = 'PsContenttype'; displayName = $labelKind; type = 'Choice'
    description = 'Wat voor soort document is dit?'; choices = @($docKinds) })
$columns.Add([ordered]@{ internalName = 'PsVertrouwelijkheid'; displayName = $labelConf; type = 'Choice'
    description = "Mag dit naar buiten? $($confLevels[0]) blijft binnen $teamName."
    choices = @($confLevels); defaultValue = $confLevels[0] })
if ($wantShareStatus) {
    $columns.Add([ordered]@{ internalName = 'PsDeelstatus'; displayName = $labelShare; type = 'Choice'
        description = 'Wordt automatisch bijgehouden en toont of dit bestand buiten de organisatie open staat. Niet zelf invullen.'
        choices = @('Niet gedeeld', 'Intern gedeeld', 'Extern - alleen bekijken', 'Extern - bewerken')
        defaultValue = 'Niet gedeeld'; readOnlyInForms = $true })
}
$columns.Add([ordered]@{ internalName = 'PsStatus'; displayName = $labelLife; type = 'Choice'
    description = "Laat $($lifecycle[0]) staan. Zet op $($lifecycle[1]) wat weg mag maar nog niet verwijderd hoeft."
    choices = @($lifecycle); defaultValue = $lifecycle[0] })

$columnNames = @($columns | ForEach-Object { $_.internalName })

function New-PillarFields {
    <# Which labels a pillar's content type asks for, and which it insists on. #>
    param([string] $Pillar)

    $fields = [System.Collections.Generic.List[object]]::new()
    $fields.Add([ordered]@{ internalName = 'PsMerk'; required = $true })

    if ($Pillar -eq $supplierPillar -and 'PsLeverancier' -in $columnNames) {
        $fields.Add([ordered]@{ internalName = 'PsLeverancier'; required = $true })
    }
    if ($Pillar -eq $salesPillar -and 'PsRegio' -in $columnNames) {
        $fields.Add([ordered]@{ internalName = 'PsRegio'; required = $true })
    }
    $fields.Add([ordered]@{ internalName = 'PsContenttype'; required = $true })
    $fields.Add([ordered]@{ internalName = 'PsTaal'; required = $true })
    $fields.Add([ordered]@{ internalName = 'PsVertrouwelijkheid'; required = ($Pillar -in $privatePillars) })
    $fields.Add([ordered]@{ internalName = 'PsPijler'; required = $false })
    $fields.Add([ordered]@{ internalName = 'PsStatus'; required = $false })
    if ($wantShareStatus) { $fields.Add([ordered]@{ internalName = 'PsDeelstatus'; required = $false }) }
    return $fields
}

$contentTypes = [System.Collections.Generic.List[object]]::new()
$groups       = [System.Collections.Generic.List[object]]::new()
$containers   = [System.Collections.Generic.List[object]]::new()
$sites        = [ordered]@{ team = $teamSite }

$viewFieldsBase = @('DocIcon', 'LinkFilename')
$viewFieldsTail = @('PsMerk', 'PsTaal', 'PsVertrouwelijkheid') +
                  $(if ($wantShareStatus) { @('PsDeelstatus') } else { @() }) +
                  @('PsStatus', 'Modified')

foreach ($pillar in $pillars) {
    $d         = $detail[$pillar]
    $ctName    = $d.ContentType
    $isPrivate = $pillar -in $privatePillars
    $siteKey   = if ($isPrivate) { Get-Slug $pillar } else { 'team' }
    # SharePoint builds a private channel's site as <teamsite>-<channel>; the team step
    # reads the real URL back from Graph afterwards, this is only the starting guess.
    if ($isPrivate) { $sites[$siteKey] = "$teamSite-$(Get-Slug $d.Channel)" }

    $contentTypes.Add([ordered]@{
        name        = $ctName
        id          = New-ContentTypeId
        description = "Documenten van de pijler $pillar."
        fields      = (New-PillarFields -Pillar $pillar)
    })

    foreach ($pair in @(@($d.GroupRw, 'bewerken'), @($d.GroupRo, 'alleen lezen'))) {
        $groups.Add([ordered]@{
            displayName  = $pair[0]
            mailNickname = (Get-Slug $pair[0])
            description  = "$client - $pillar - $($pair[1])"
        })
    }

    # The view groups on the label that actually distinguishes documents in this
    # pillar, but only where that column exists at all.
    $groupBy = $d.GroupBy
    if ($groupBy -notin $columnNames) { $groupBy = 'PsContenttype' }

    $container = [ordered]@{
        key                 = $pillar
        title               = $d.Channel
        kind                = 'ChannelFolder'
        channelType         = $(if ($isPrivate) { 'Private' } else { 'Standard' })
        site                = $siteKey
        list                = $channelList
        folder              = $d.Folder
        pillar              = $pillar
        defaultColumnValues = [ordered]@{ PsPijler = $pillar; PsStatus = $lifecycle[0] }
        contentTypes        = @($ctName)
        defaultContentType  = $ctName
        # A private channel has its own membership - breaking inheritance there would
        # fight Teams for control of the same thing.
        uniquePermissions   = (-not $isPrivate -and $tightenChannels)
        view                = [ordered]@{
            title   = $d.ViewTitle
            groupBy = $groupBy
            fields  = $viewFieldsBase + @($groupBy) + $viewFieldsTail
        }
        permissions         = @(
            [ordered]@{ group = $d.GroupRw; role = 'Contribute' }
            [ordered]@{ group = $d.GroupRo; role = 'Read' }
        )
    }
    if ($isPrivate) {
        $container['note'] = 'Privekanaal - Teams beheert de rechten via het lidmaatschap van de kanaalsite.'
        $container['permissions'] = @()
    }
    $containers.Add($container)
}

# The external customer library, if there is one.
if ($extLibrary) {
    $extGroup = $extGroupName
    $groups.Add([ordered]@{
        displayName = $extGroup; mailNickname = (Get-Slug $extGroup)
        description = "$client - externe klanten met leestoegang tot $extLibrary"
    })
    $extCt = $extCtName
    $extFields = [System.Collections.Generic.List[object]]::new()
    $extFields.Add([ordered]@{ internalName = 'PsMerk'; required = $true })
    $extFields.Add([ordered]@{ internalName = 'PsContenttype'; required = $true })
    $extFields.Add([ordered]@{ internalName = 'PsTaal'; required = $true })
    $extFields.Add([ordered]@{ internalName = 'PsVertrouwelijkheid'; required = $false })
    $extFields.Add([ordered]@{ internalName = 'PsStatus'; required = $false })
    if ($wantShareStatus) { $extFields.Add([ordered]@{ internalName = 'PsDeelstatus'; required = $false }) }

    $contentTypes.Add([ordered]@{
        name = $extCt; id = (New-ContentTypeId)
        description = "Materiaal dat met klanten gedeeld wordt."
        fields = $extFields
    })

    $defaults = [ordered]@{ PsStatus = $lifecycle[0] }
    if ($confLevels.Count -gt 1) { $defaults['PsVertrouwelijkheid'] = $confLevels[1] }

    # Whoever curates marketing keeps write access to the customer library.
    $marketingPillar = @($pillars | Where-Object { $_ -match 'Marketing' } | Select-Object -First 1)
    $ownerGroups = @()
    if ($marketingPillar) { $ownerGroups += [ordered]@{ group = $detail[$marketingPillar].GroupRw; role = 'Contribute' } }
    $ownerGroups += [ordered]@{ group = $extGroup; role = 'Read' }

    $containers.Add([ordered]@{
        key = 'KlantBibliotheek'; title = $extLibrary; kind = 'Library'; site = 'team'
        list = $extLibrary
        defaultColumnValues = $defaults
        contentTypes = @($extCt); defaultContentType = $extCt
        uniquePermissions = $true
        keepExistingPermissions = $false
        view = [ordered]@{ title = 'Op merk'; groupBy = 'PsMerk'; fields = $viewFieldsBase + @('PsMerk', 'PsContenttype', 'PsTaal', 'PsStatus', 'Modified') }
        permissions = $ownerGroups
    })
}

# Cross-cutting views: one per brand, spanning every pillar folder.
$crossViews = [System.Collections.Generic.List[object]]::new()
foreach ($brand in $brands) {
    $clauses = "<Eq><FieldRef Name='PsMerk' /><Value Type='Text'>$brand</Value></Eq>"
    if ($bothName) { $clauses = "<Or>$clauses<Eq><FieldRef Name='PsMerk' /><Value Type='Text'>$bothName</Value></Eq></Or>" }
    $crossViews.Add([ordered]@{
        title = "Alles - $brand"
        description = "Elk bestand van $brand$(if ($bothName) { " of $bothName" }), dwars door alle pijlers heen."
        recursive = $true; groupBy = 'PsPijler'; where = $clauses
        fields = $viewFieldsBase + @('PsPijler', 'PsContenttype', 'PsTaal', 'PsStatus', 'Modified')
    })
}
$crossViews.Add([ordered]@{
    title = 'Nog te taggen'
    description = 'Bestanden zonder merk - meestal binnengesleept of via de gesynchroniseerde map gekopieerd.'
    recursive = $true; where = "<IsNull><FieldRef Name='PsMerk' /></IsNull>"
    fields = $viewFieldsBase + @('PsPijler', 'Editor', 'Modified')
})
if ($wantShareStatus) {
    $crossViews.Add([ordered]@{
        title = 'Extern gedeeld'
        description = 'Alles wat buiten de organisatie open staat.'
        recursive = $true; groupBy = 'PsVertrouwelijkheid'
        where = "<Contains><FieldRef Name='PsDeelstatus' /><Value Type='Text'>Extern</Value></Contains>"
        fields = $viewFieldsBase + @('PsDeelstatus', 'PsVertrouwelijkheid', 'PsPijler', 'PsMerk', 'Modified')
    })
}
$crossViews.Add([ordered]@{
    title = 'Te archiveren'
    description = "Alles met status $($lifecycle[1])$(if ($lifecycle.Count -gt 2) { " of $($lifecycle[2])" })."
    recursive = $true; groupBy = 'PsStatus'
    where = $(
        $parts = @($lifecycle | Select-Object -Skip 1 | ForEach-Object {
            "<Eq><FieldRef Name='PsStatus' /><Value Type='Text'>$_</Value></Eq>" })
        if ($parts.Count -gt 1) { "<Or>$($parts -join '')</Or>" } else { $parts[0] }
    )
    fields = $viewFieldsBase + @('PsStatus', 'PsPijler', 'PsMerk', 'Editor', 'Modified')
})

$fieldRoles = [ordered]@{
    brand = 'PsMerk'; pillar = 'PsPijler'; confidentiality = 'PsVertrouwelijkheid'; lifecycle = 'PsStatus'
}
if ($wantShareStatus) { $fieldRoles['shareStatus'] = 'PsDeelstatus' }

$config = [ordered]@{
    client      = $client
    description = "SharePoint-structuur, metadata en groepsrechten voor $client."
    tenant      = $tenant
    team        = [ordered]@{
        displayName  = $teamName
        mailNickname = $teamAlias
        description  = "Documentbeheer $client."
        visibility   = 'Private'
        owners       = @($teamOwner)
    }
    sites            = $sites
    columnGroup      = $columnGroup
    contentTypeGroup = $ctGroup
    fieldRoles       = $fieldRoles
}
if ($suppliers.Count -gt 0) {
    $config['termStore'] = [ordered]@{
        group = $columnGroup; termSet = $termSetName
        description = "Leveranciers van $client - uitbreidbaar vanuit de term store."
        terms = @($suppliers)
    }
}
$config['columns']      = @($columns)
$config['contentTypes'] = @($contentTypes)
$config['groups']       = @($groups)
$config['libraryViews'] = @(@{ site = 'team'; list = $channelList; views = @($crossViews) })
$config['containers']   = @($containers)

$json = $config | ConvertTo-Json -Depth 12

# -- Write ---------------------------------------------------------------------
if (-not $Path) { $Path = Join-Path $PSScriptRoot ((Get-Slug $client) + '.config.json') }

Write-Host ''
Write-Host '  ── Dit wordt het ─────────────────────────────────────────' -ForegroundColor DarkGray
Write-Host "    Team              $teamName  ($teamSite)" -ForegroundColor Gray
Write-Host "    Kanalen           $($pillars -join ', ')" -ForegroundColor Gray
Write-Host "    Privekanaal       $(if ($privatePillars) { $privatePillars -join ', ' } else { 'geen' })" -ForegroundColor Gray
if ($extLibrary) { Write-Host "    Klantbibliotheek  $extLibrary" -ForegroundColor Gray }
Write-Host "    Merken            $($brandChoices -join ', ')" -ForegroundColor Gray
Write-Host "    Groepen           $($groups.Count) stuks, $groupPrefix-<pijler>-$rwSuffix / -$roSuffix" -ForegroundColor Gray
Write-Host "    Labels            $($columns.Count)" -ForegroundColor Gray
Write-Host "    Documenttypes     $($contentTypes.Count)" -ForegroundColor Gray
Write-Host "    Weergaven         $($crossViews.Count) over alle pijlers + 1 per kanaal" -ForegroundColor Gray
Write-Host "    Onderhoud         $(if ($wantShareStatus) { 'een nachtelijk script voor de deelstatus' } else { 'geen - nul scripts na het bouwen' })" -ForegroundColor Gray
Write-Host ''

if ($ShowJson) {
    Write-Host $json
    exit 0
}

if ((Test-Path $Path) -and -not $Force) {
    Write-Host "  Er staat al een configuratie op $Path." -ForegroundColor Yellow
    Write-Host '  Overschrijven maakt nieuwe documenttype-IDs aan - doe dat alleen als er nog' -ForegroundColor Yellow
    Write-Host '  niets is uitgerold. Gebruik -Force als je het zeker weet.' -ForegroundColor Yellow
    exit 1
}

$json | Set-Content -Path $Path -Encoding UTF8
Write-Host "  Geschreven: $Path" -ForegroundColor Green
Write-Host ''
Write-Host '  Nu bouwen:' -ForegroundColor Cyan
Write-Host "    .\Install-SharePointStructure.ps1 -ConfigPath `"$Path`" -WhatIf" -ForegroundColor Gray
Write-Host "    .\Install-SharePointStructure.ps1 -ConfigPath `"$Path`"" -ForegroundColor Gray
Write-Host ''

# The path is the only thing on the pipeline - everything above is Write-Host - so a
# caller can pick up what was written instead of guessing at the folder afterwards.
Write-Output $Path
