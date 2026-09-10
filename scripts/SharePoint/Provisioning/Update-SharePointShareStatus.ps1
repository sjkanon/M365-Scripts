#Requires -Version 7.0
<#
.SYNOPSIS
    Work out how every document is actually shared and write that back to the
    Deelstatus column. Reports files shared wider than their Vertrouwelijkheid tag
    allows. Supports -WhatIf.

.DESCRIPTION
    Step three of the structure set, and the only one meant to run on a schedule.
    Deelstatus is the one column nobody fills in by hand: this script derives it from
    the permissions that are really on the file and writes it back.

    Per document, in order of severity - the widest wins:

        Extern - bewerken        an Anyone-edit link, or a guest with Contribute or
                                 more, or a specific-people edit link that contains
                                 a guest
        Extern - alleen bekijken the same, but view-only
        Intern gedeeld           a link or a direct grant that stays inside the
                                 tenant, on top of what the library already gives
        Niet gedeeld             the file just inherits from its folder or library

    How it decides
    --------------
    A file that inherits its permissions is not shared, full stop - one round trip
    per page of a thousand items settles that for most of the library. Only files
    that broke inheritance get their role assignments read.

    Behind every "Copy link" SharePoint keeps a group called SharingLinks.<guid>.*.
    The name says what kind of link it is: Anonymous (Anyone with the link - external
    by definition), Organization (people in the tenant), or Flexible (specific
    people, which may be a guest). Only Flexible links need their members expanded,
    so the expensive lookup happens for the handful of links where the answer is not
    already in the name.

    Guests are recognised by #ext# in their login name, which is how SharePoint
    stores every B2B guest account.

    Vertrouwelijkheid violations
    ----------------------------
    Deriving the deelstatus makes one question cheap to answer: is anything tagged
    Intern or Vertrouwelijk sitting behind an external link? Every one of those is
    written to the report and the run exits with code 2, so a scheduled job in an RMM
    turns up in the activity feed exactly when there is something to look at.

    Nothing is changed about the sharing itself. This script reads permissions and
    writes one metadata column - it never revokes a link. Revoking is a decision, and
    a decision belongs with a person.

    The write uses SystemUpdate, so Modified and Modified By stay as they were and no
    new version is created. A nightly run does not push every document to the top of
    "recently changed".

    Exit codes
    ----------
        0  done, nothing shared wider than its tag allows
        1  failure
        2  at least one Vertrouwelijkheid violation found

.PARAMETER ConfigPath
    Path to the structure configuration JSON.
    Default: petsolutions.config.json next to this script.

.PARAMETER Container
    Only audit these containers (keys from the containers section). Default: all.

.PARAMETER ReportOnly
    Work out the deelstatus and report it, but do not write the column.

.PARAMETER MaxItems
    Stop after this many documents per container (0 = no limit, the default).

.PARAMETER PageSize
    Items fetched per server call (default 1000).

.PARAMETER Interactive
    Sign in interactively.

.PARAMETER ClientId
    Client ID of the Entra app registration used to sign in.

.PARAMETER Thumbprint
    Certificate thumbprint for app-only sign-in. Implies app-only - this is what a
    scheduled run uses.

.PARAMETER CertificatePath
    PFX file for app-only sign-in, as an alternative to -Thumbprint.

.PARAMETER CertificatePassword
    Password for -CertificatePath.

.PARAMETER Tenant
    Tenant name or ID. Defaults to the tenant in the configuration file.

.PARAMETER ReportPath
    CSV with one row per document whose deelstatus changed or that violates its
    Vertrouwelijkheid tag.
    Default: C:\Temp\SharePointStructure_ShareStatus_<timestamp>.csv

.PARAMETER Quiet
    Only print the summary and anything that needs attention. For scheduled runs.

.PARAMETER Disconnect
    Sign out of PnP when finished.

.EXAMPLE
    # What would the audit change, and is anything shared too widely?
    .\Update-SharePointShareStatus.ps1 -Interactive -ClientId <app-id> -WhatIf

.EXAMPLE
    # Nightly, unattended: update Deelstatus and flag violations
    .\Update-SharePointShareStatus.ps1 -ClientId <app-id> -Thumbprint <thumbprint> -Quiet

.EXAMPLE
    # Just the external customer library, reporting only
    .\Update-SharePointShareStatus.ps1 -Interactive -ClientId <app-id> `
        -Container FUTECH -ReportOnly

.NOTES
    Author  : Sjoerd Kanon
    Requires: PnP.PowerShell 2.x
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $ConfigPath,
    [string[]] $Container,

    [switch] $ReportOnly,

    [ValidateRange(0, [int]::MaxValue)]
    [int] $MaxItems = 0,

    [ValidateRange(100, 5000)]
    [int] $PageSize = 1000,

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
    $ReportPath = Join-Path $outputDir "SharePointStructure_ShareStatus_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
}

$connectSplat = @{
    Tenant              = $Tenant
    ClientId            = $ClientId
    Thumbprint          = $Thumbprint
    CertificatePath     = $CertificatePath
    CertificatePassword = $CertificatePassword
    Interactive         = $Interactive
}

# Column names come from the config so a client that renamed them still works.
$statusField = 'PsDeelstatus'
$secretField = 'PsVertrouwelijkheid'
$statusChoices = @($config.columns | Where-Object { $_.internalName -eq $statusField } |
                   ForEach-Object { Get-ConfigValue $_ 'choices' @() })
if ($statusChoices.Count -lt 4) {
    throw "The configuration's $statusField column needs the four deelstatus choices."
}
$NotShared  = $statusChoices[0]
$Internal   = $statusChoices[1]
$ExternView = $statusChoices[2]
$ExternEdit = $statusChoices[3]

# Widest wins. The index into this list is the severity.
$severity = @($NotShared, $Internal, $ExternView, $ExternEdit)

# The tags that must never end up behind an external link.
$sensitiveTags = @('Intern', 'Vertrouwelijk')

$report     = [System.Collections.Generic.List[object]]::new()
$scanned    = 0
$updated    = 0
$violations = 0
$failures   = 0

function Write-Detail { param([string] $Message) if (-not $Quiet) { Write-Ok $Message } }

function Get-ItemValue {
    <#
        FieldValues is a Dictionary, so asking it for a column the library does not
        have throws instead of returning nothing. That happens on any library where
        the metadata script has not run yet, and it should read as "no value", not as
        a crash halfway through an audit.
    #>
    param($Item, [string] $Name)

    if (-not $Item.FieldValues.ContainsKey($Name)) { return $null }
    return $Item.FieldValues[$Name]
}

function Get-ShareVerdict {
    <#
        Read the role assignments of one item that broke inheritance and return the
        widest way it is shared, plus the evidence for it.
    #>
    param($Item, $Connection)

    $context = Get-PnPContext -Connection $Connection
    $context.Load($Item.RoleAssignments)
    $context.ExecuteQuery()
    foreach ($assignment in $Item.RoleAssignments) {
        $context.Load($assignment.Member)
        $context.Load($assignment.RoleDefinitionBindings)
    }
    $context.ExecuteQuery()

    $verdict  = $NotShared
    $evidence = @()

    foreach ($assignment in $Item.RoleAssignments) {
        $login = $assignment.Member.LoginName
        $title = $assignment.Member.Title
        $roles = @($assignment.RoleDefinitionBindings | ForEach-Object { $_.Name })

        # Limited Access is the bookkeeping entry SharePoint adds so someone can
        # reach a deeper item. It grants nothing on its own.
        $roles = @($roles | Where-Object { $_ -notin @('Limited Access', 'Beperkte toegang', 'Web-Only Limited Access') })
        if ($roles.Count -eq 0) { continue }

        $canEdit = [bool] @($roles | Where-Object { $_ -notin @('Read', 'Lezen', 'View Only', 'Restricted View') }).Count

        if ($title -like 'SharingLinks*') {
            switch -Regex ($title) {
                '\.AnonymousEdit\.'    { $found = $ExternEdit; $why = 'Anyone-link met bewerkrechten' }
                '\.AnonymousView\.'    { $found = $ExternView; $why = 'Anyone-link, alleen bekijken' }
                '\.OrganizationEdit\.' { $found = $Internal;   $why = 'organisatielink met bewerkrechten' }
                '\.OrganizationView\.' { $found = $Internal;   $why = 'organisatielink, alleen bekijken' }
                default {
                    # A specific-people link: the only kind whose name does not say
                    # whether a guest is behind it, so this one gets expanded.
                    $guests = @()
                    try {
                        $guests = @(Get-PnPGroupMember -Group $title -Connection $Connection -ErrorAction Stop |
                                    Where-Object { $_.LoginName -match '#ext#|urn:spo:guest' })
                    } catch {
                        $guests = @()
                    }
                    if ($guests.Count -gt 0) {
                        $found = if ($canEdit) { $ExternEdit } else { $ExternView }
                        $why   = "link voor specifieke personen met gast(en): $(($guests | ForEach-Object { $_.Email ?? $_.Title }) -join ', ')"
                    } else {
                        $found = $Internal
                        $why   = 'link voor specifieke personen, intern'
                    }
                }
            }
        } elseif ($login -match '#ext#|urn:spo:guest') {
            $found = if ($canEdit) { $ExternEdit } else { $ExternView }
            $why   = "gast $title rechtstreeks toegevoegd ($($roles -join ', '))"
        } else {
            $found = $Internal
            $why   = "$title rechtstreeks toegevoegd ($($roles -join ', '))"
        }

        $evidence += $why
        if ($severity.IndexOf($found) -gt $severity.IndexOf($verdict)) { $verdict = $found }
    }

    return [PSCustomObject]@{
        Status   = $verdict
        Evidence = ($evidence -join '; ')
    }
}

function Update-ContainerStatus {
    <# Walk one container and bring its Deelstatus column up to date. #>
    param($Definition, $Connection)

    $listTitle = Get-ConfigValue $Definition 'list' $Definition.title
    $kind      = Get-ConfigValue $Definition 'kind' 'Library'

    $list = Get-PnPList -Identity $listTitle -Connection $Connection -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-Bad "library '$listTitle' not found - skipped"
        $script:failures++
        return
    }

    $itemSplat = @{
        List       = $listTitle
        PageSize   = $PageSize
        Connection = $Connection
        Fields     = @('ID', 'FileRef', 'FileLeafRef', 'FSObjType', $statusField, $secretField)
    }
    if ($kind -eq 'ChannelFolder') {
        $context = Get-PnPContext -Connection $Connection
        $context.Load($list.RootFolder)
        $context.ExecuteQuery()
        $folderName = Get-ConfigValue $Definition 'folder' $Definition.title
        $itemSplat['FolderServerRelativeUrl'] = "$($list.RootFolder.ServerRelativeUrl.TrimEnd('/'))/$folderName"
    }

    $items = @(Get-PnPListItem @itemSplat)
    if ($MaxItems -gt 0 -and $items.Count -gt $MaxItems) { $items = $items[0..($MaxItems - 1)] }

    # Loading a list item without an expression brings its default scalar properties
    # along, HasUniqueRoleAssignments included - so one round trip answers "is this
    # shared at all" for a whole batch. That is what keeps a library of 20.000 files
    # down to a handful of calls: only items that broke inheritance cost more.
    # Batched at 100 because a single request carrying thousands of loads is refused.
    $context = Get-PnPContext -Connection $Connection
    for ($offset = 0; $offset -lt $items.Count; $offset += 100) {
        $batch = $items[$offset..([Math]::Min($offset + 99, $items.Count - 1))]
        foreach ($item in $batch) { $context.Load($item) }
        $context.ExecuteQuery()
    }

    $containerUpdated = 0
    foreach ($item in $items) {
        # Folders carry permissions too, but the deelstatus column is about documents.
        if ((Get-ItemValue $item 'FSObjType') -ne 0) { continue }
        $script:scanned++

        $current = Get-ItemValue $item $statusField
        $secret  = Get-ItemValue $item $secretField
        $path    = Get-ItemValue $item 'FileRef'

        if ($item.HasUniqueRoleAssignments) {
            $verdict = Get-ShareVerdict -Item $item -Connection $Connection
        } else {
            $verdict = [PSCustomObject]@{ Status = $NotShared; Evidence = 'erft van de map/bibliotheek' }
        }

        $isViolation = ($verdict.Status -in @($ExternView, $ExternEdit)) -and ($secret -in $sensitiveTags)
        if ($isViolation) {
            $script:violations++
            Write-Warn "$path is '$secret' maar staat extern open: $($verdict.Evidence)"
        }

        if ($current -eq $verdict.Status) {
            if ($isViolation) {
                $report.Add([PSCustomObject]@{
                    Container = $Definition.key; Path = $path
                    Was = $current; Becomes = $verdict.Status
                    Vertrouwelijkheid = $secret; Violation = $true; Evidence = $verdict.Evidence
                })
            }
            continue
        }

        $report.Add([PSCustomObject]@{
            Container = $Definition.key; Path = $path
            Was = $current; Becomes = $verdict.Status
            Vertrouwelijkheid = $secret; Violation = $isViolation; Evidence = $verdict.Evidence
        })

        if ($ReportOnly) { continue }
        if (-not $PSCmdlet.ShouldProcess($path, "Set Deelstatus to '$($verdict.Status)'")) { continue }

        try {
            # SystemUpdate: no new version, and Modified/Modified By stay put, so a
            # nightly run does not push the whole library to the top of "recent".
            Set-PnPListItem -List $listTitle -Identity $item.Id -Values @{ $statusField = $verdict.Status } `
                -UpdateType SystemUpdate -Connection $Connection | Out-Null
            $script:updated++
            $containerUpdated++
        } catch {
            Write-Bad "could not update $path : $($_.Exception.Message)"
            $script:failures++
        }
    }

    Write-Detail "$($Definition.key): $($items.Count) item(s) read, $containerUpdated updated"
}

# -- Run -----------------------------------------------------------------------
$exitCode = 0
try {
    if (-not $Quiet) {
        Write-Host ''
        Write-Host "  Client : $(Get-ConfigValue $config 'client' '(unnamed)')" -ForegroundColor Cyan
        Write-Host "  Scope  : $(($containers | ForEach-Object { $_.key }) -join ', ')" -ForegroundColor Cyan
        Write-Host "  Mode   : $(if ($simulate) { '-WhatIf' } elseif ($ReportOnly) { 'report only' } else { 'APPLY' })" -ForegroundColor Cyan
    }

    foreach ($siteKey in @($containers | ForEach-Object { $_.site } | Select-Object -Unique)) {
        $url = Get-StructureSiteUrl -Config $config -SiteKey $siteKey
        if (-not $Quiet) { Write-Head "Site '$siteKey' - $url" }
        $connection = Connect-Structure -Url $url @connectSplat

        foreach ($entry in ($containers | Where-Object { $_.site -eq $siteKey })) {
            Update-ContainerStatus -Definition $entry -Connection $connection
        }
    }

    if ($report.Count -gt 0) {
        if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }
        $report | Sort-Object -Property @{ Expression = 'Violation'; Descending = $true }, Container, Path |
            Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    }

    if ($violations -gt 0) { $exitCode = 2 }
    if ($failures -gt 0)   { $exitCode = 1 }

    # In -Quiet mode this is the only thing a scheduled run prints when all is well.
    Write-Head 'Summary'
    Write-Host "    Scanned          : $scanned document(s)" -ForegroundColor Cyan
    Write-Host "    Deelstatus $(if ($simulate -or $ReportOnly) { 'to change ' } else { 'updated   ' }): $(if ($simulate -or $ReportOnly) { $report.Count } else { $updated })" -ForegroundColor Cyan
    if ($violations -gt 0) {
        Write-Host "    Too widely shared: $violations file(s) tagged Intern/Vertrouwelijk behind an external link" -ForegroundColor Red
    } else {
        Write-Ok 'Nothing shared wider than its Vertrouwelijkheid tag allows'
    }
    if ($failures -gt 0) { Write-Bad "$failures item(s) could not be updated" }
    if ($report.Count -gt 0) { Write-Host "    Report           : $ReportPath" -ForegroundColor Cyan }
    Write-Host ''
} catch {
    Write-Host ''
    Write-Bad "Aborted: $($_.Exception.Message)"
    Write-Host ''
    $exitCode = 1
} finally {
    if ($Disconnect) { Disconnect-Structure }
}

exit $exitCode
