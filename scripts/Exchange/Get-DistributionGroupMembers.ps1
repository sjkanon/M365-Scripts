#Requires -Version 5.1
<#
.SYNOPSIS
    Export the members of every distribution list to a customer-readable Excel workbook.

.DESCRIPTION
    Connects to Exchange Online and reports who is in which distribution list:

      - Distribution groups and mail-enabled security groups (always)
      - Dynamic distribution groups          (-IncludeDynamic)
      - Microsoft 365 groups                 (-IncludeM365Groups)

    The workbook has two sheets, both filterable tables with a frozen header row:

      Overzicht  one row per list: name, address, type, member count, owners
      Leden      one row per member: which list, who, which address, what kind of member

    The sheet headers are Dutch because the workbook is what goes to the customer;
    the script itself stays English like the rest of the repo.

    Without ImportExcel the report is written as two CSV files instead, so the
    script never fails just because a module is missing.

.PARAMETER Group
    One list (name, alias or e-mail). Without it every list is reported.

.PARAMETER Member
    Only report the lists this address is a member of - the answer to "which lists
    is this person on?". Resolved server-side, so it stays fast in a large tenant.
    The matched lists are still exported in full, so the customer sees who else is on
    them. Direct membership only: a person inside a nested group is not a match, but
    the nested group itself shows up as a member row.

.PARAMETER IncludeDynamic
    Also report dynamic distribution groups. Their membership is evaluated live, which
    costs one query per group.

.PARAMETER IncludeM365Groups
    Also report Microsoft 365 groups (Teams-backed groups included).

.PARAMETER OutputPath
    Path of the .xlsx to write. Defaults to C:\Temp\Distributielijsten_<timestamp>.xlsx.

.PARAMETER Csv
    Write CSV instead of Excel, even when ImportExcel is available.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Get-DistributionGroupMembers.ps1
    Every distribution list with all of its members, as one Excel workbook.

.EXAMPLE
    .\Get-DistributionGroupMembers.ps1 -Member "jan@contoso.com"
    Only the lists Jan is on - including the other members of those lists.

.EXAMPLE
    .\Get-DistributionGroupMembers.ps1 -Group "helpdesk@contoso.com" -OutputPath "C:\Reports\helpdesk.xlsx"

.EXAMPLE
    .\Get-DistributionGroupMembers.ps1 -IncludeDynamic -IncludeM365Groups
    Everything that can receive mail as a group.

.NOTES
    Author: Sjoerd Kanon
#>
[CmdletBinding()]
param(
    [string] $Group,
    [string] $Member,
    [switch] $IncludeDynamic,
    [switch] $IncludeM365Groups,
    [string] $OutputPath,
    [switch] $Csv,
    [string] $TenantId
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
} catch {
    $connectParams = @{ ShowBanner = $false }
    if ($TenantId) { $connectParams['Organization'] = $TenantId }
    Connect-ExchangeOnline @connectParams
    $script:ConnectedHere = $true
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Distribution List Members" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# Exchange recipient types say nothing to a customer, so every one of them gets a
# plain Dutch name. Anything unmapped falls through with its raw value, which is
# still better than an empty cell when Microsoft adds a type.
$typeNames = @{
    'MailUniversalDistributionGroup' = 'Distributielijst'
    'MailUniversalSecurityGroup'     = 'Beveiligingsgroep (mail-enabled)'
    'MailNonUniversalGroup'          = 'Distributielijst (legacy)'
    'DynamicDistributionGroup'       = 'Dynamische distributielijst'
    'GroupMailbox'                   = 'Microsoft 365-groep'
    'RoomList'                       = 'Ruimtelijst'
    'UserMailbox'                    = 'Gebruiker'
    'SharedMailbox'                  = 'Gedeelde postbus'
    'RoomMailbox'                    = 'Vergaderruimte'
    'EquipmentMailbox'               = 'Apparatuur'
    'MailUser'                       = 'Externe gebruiker'
    'MailContact'                    = 'Externe contactpersoon'
    'GuestMailUser'                  = 'Gast'
    'PublicFolder'                   = 'Openbare map'
}

function Get-FriendlyType {
    param([string] $RecipientType)
    if (-not $RecipientType)                    { return 'Onbekend' }
    if ($typeNames.ContainsKey($RecipientType)) { return $typeNames[$RecipientType] }
    return $RecipientType
}

function Get-DynamicMember {
    param($DynamicGroup)
    Get-Recipient -ResultSize Unlimited `
        -RecipientPreviewFilter $DynamicGroup.RecipientFilter `
        -OrganizationalUnit $DynamicGroup.RecipientContainer -ErrorAction Stop
}

# ── Collect the lists ─────────────────────────────────────────────────────────
# Every entry is normalised to { Object, Kind } so the member lookup below only has
# to branch on Kind, not on which cmdlet happened to produce the object.
$lists = [System.Collections.Generic.List[PSObject]]::new()

if ($Group) {
    $found = $null
    foreach ($getter in @(
        { Get-DistributionGroup        -Identity $Group -ErrorAction Stop },
        { Get-DynamicDistributionGroup -Identity $Group -ErrorAction Stop },
        { Get-UnifiedGroup             -Identity $Group -ErrorAction Stop }
    )) {
        try { $found = & $getter; break } catch { }
    }
    if (-not $found) { throw "No distribution list found for '$Group'." }

    $kind = switch ($found.RecipientTypeDetails) {
        'DynamicDistributionGroup' { 'Dynamic' }
        'GroupMailbox'             { 'Unified' }
        default                    { 'Static'  }
    }
    $lists.Add([PSCustomObject]@{ Object = $found; Kind = $kind })

} elseif ($Member) {
    # Resolve the address first: the membership filter matches on DN, not on SMTP.
    $recipient = Get-Recipient -Identity $Member -ErrorAction Stop
    Write-Host "  Looking up lists for $($recipient.PrimarySmtpAddress)..." -ForegroundColor DarkGray

    $dn   = $recipient.DistinguishedName -replace "'", "''"
    $hits = @(Get-Recipient -ResultSize Unlimited -Filter "Members -eq '$dn'" -ErrorAction Stop)

    foreach ($hit in $hits) {
        if ($hit.RecipientTypeDetails -eq 'GroupMailbox') {
            if (-not $IncludeM365Groups) { continue }
            $lists.Add([PSCustomObject]@{ Object = Get-UnifiedGroup -Identity $hit.Identity; Kind = 'Unified' })
        } else {
            $lists.Add([PSCustomObject]@{ Object = Get-DistributionGroup -Identity $hit.Identity; Kind = 'Static' })
        }
    }

    # Dynamic groups store no membership, so the filter above cannot see them - they
    # have to be evaluated one by one.
    if ($IncludeDynamic) {
        foreach ($ddg in @(Get-DynamicDistributionGroup -ResultSize Unlimited)) {
            try {
                $onList = Get-DynamicMember -DynamicGroup $ddg |
                          Where-Object { $_.DistinguishedName -eq $recipient.DistinguishedName }
                if ($onList) { $lists.Add([PSCustomObject]@{ Object = $ddg; Kind = 'Dynamic' }) }
            } catch {
                Write-Host "  [WARN] $($ddg.DisplayName): $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }

} else {
    Write-Host "  Retrieving distribution lists..." -ForegroundColor DarkGray
    foreach ($dg in @(Get-DistributionGroup -ResultSize Unlimited)) {
        $lists.Add([PSCustomObject]@{ Object = $dg; Kind = 'Static' })
    }
    if ($IncludeDynamic) {
        foreach ($ddg in @(Get-DynamicDistributionGroup -ResultSize Unlimited)) {
            $lists.Add([PSCustomObject]@{ Object = $ddg; Kind = 'Dynamic' })
        }
    }
    if ($IncludeM365Groups) {
        foreach ($ug in @(Get-UnifiedGroup -ResultSize Unlimited)) {
            $lists.Add([PSCustomObject]@{ Object = $ug; Kind = 'Unified' })
        }
    }
}

if ($lists.Count -eq 0) {
    if ($Member) { Write-Host "  $Member is not a direct member of any distribution list." -ForegroundColor Yellow }
    else         { Write-Host "  No distribution lists found." -ForegroundColor Yellow }
    if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
    return
}

Write-Host "  Reading members of $($lists.Count) list(s)..." -ForegroundColor DarkGray
Write-Host ""

# ── Read the members ──────────────────────────────────────────────────────────
$overview = [System.Collections.Generic.List[PSObject]]::new()
$members  = [System.Collections.Generic.List[PSObject]]::new()
$index    = 0

foreach ($entry in $lists) {
    $list = $entry.Object
    $index++
    Write-Progress -Activity "Distribution lists" -Status $list.DisplayName `
        -PercentComplete ([int](100 * $index / $lists.Count))

    $groupMembers = @()
    $owners       = ''

    try {
        switch ($entry.Kind) {
            'Static' {
                $groupMembers = @(Get-DistributionGroupMember -Identity $list.Identity -ResultSize Unlimited -ErrorAction Stop)
                $owners = ($list.ManagedBy | ForEach-Object { ($_ -split '/')[-1] }) -join '; '
            }
            'Dynamic' {
                $groupMembers = @(Get-DynamicMember -DynamicGroup $list)
                $owners = ($list.ManagedBy | ForEach-Object { ($_ -split '/')[-1] }) -join '; '
            }
            'Unified' {
                $groupMembers = @(Get-UnifiedGroupLinks -Identity $list.Identity -LinkType Members -ResultSize Unlimited -ErrorAction Stop)
                $owners = (@(Get-UnifiedGroupLinks -Identity $list.Identity -LinkType Owners -ResultSize Unlimited -ErrorAction SilentlyContinue) |
                           ForEach-Object { $_.DisplayName }) -join '; '
            }
        }
    } catch {
        Write-Host "  [WARN] $($list.DisplayName): $($_.Exception.Message)" -ForegroundColor Yellow
    }

    $internalOnly = if ($null -ne $list.RequireSenderAuthenticationEnabled) {
        if ($list.RequireSenderAuthenticationEnabled) { 'Ja' } else { 'Nee' }
    } else { '' }

    $overview.Add([PSCustomObject]@{
        'Lijst'                    = $list.DisplayName
        'E-mailadres'              = [string]$list.PrimarySmtpAddress
        'Type'                     = Get-FriendlyType $list.RecipientTypeDetails
        'Aantal leden'             = $groupMembers.Count
        'Eigenaar(s)'              = $owners
        'Alias'                    = $list.Alias
        'Verborgen in adresboek'   = if ($list.HiddenFromAddressListsEnabled) { 'Ja' } else { 'Nee' }
        'Alleen interne afzenders' = $internalOnly
        'Aangemaakt op'            = if ($list.WhenCreated) { (Get-Date $list.WhenCreated -Format 'dd-MM-yyyy') } else { '' }
    })

    if ($groupMembers.Count -eq 0) {
        # An empty list is exactly what a customer wants to spot, so it gets a row of
        # its own instead of quietly missing from the member sheet.
        $members.Add([PSCustomObject]@{
            'Lijst'             = $list.DisplayName
            'E-mailadres lijst' = [string]$list.PrimarySmtpAddress
            'Type lijst'        = Get-FriendlyType $list.RecipientTypeDetails
            'Lid'               = '(geen leden)'
            'E-mailadres lid'   = ''
            'Type lid'          = ''
        })
        continue
    }

    foreach ($m in ($groupMembers | Sort-Object DisplayName)) {
        $members.Add([PSCustomObject]@{
            'Lijst'             = $list.DisplayName
            'E-mailadres lijst' = [string]$list.PrimarySmtpAddress
            'Type lijst'        = Get-FriendlyType $list.RecipientTypeDetails
            'Lid'               = $m.DisplayName
            'E-mailadres lid'   = [string]$m.PrimarySmtpAddress
            'Type lid'          = Get-FriendlyType $m.RecipientTypeDetails
        })
    }
}
Write-Progress -Activity "Distribution lists" -Completed

$overviewSorted = @($overview | Sort-Object 'Lijst')
$membersSorted  = @($members  | Sort-Object 'Lijst', 'Lid')

# ── Console summary ───────────────────────────────────────────────────────────
$overviewSorted | Format-Table 'Lijst', 'E-mailadres', 'Type', 'Aantal leden', 'Eigenaar(s)' -AutoSize

# ── Report ────────────────────────────────────────────────────────────────────
$ts   = Get-Date -Format 'yyyyMMdd_HHmmss'
$stem = if ($Member) { "Distributielijsten_$($Member -replace '[^\w.-]', '_')" } else { 'Distributielijsten' }

$useExcel = -not $Csv
if ($useExcel -and -not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "  ImportExcel is not installed - it is what writes the .xlsx." -ForegroundColor Yellow
    if ((Read-Host "  Install it now (CurrentUser)? [Y/n]") -notmatch '^[Nn]') {
        try {
            Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        } catch {
            Write-Host "  [WARN] Install failed: $($_.Exception.Message)" -ForegroundColor Yellow
            $useExcel = $false
        }
    } else {
        $useExcel = $false
    }
}

if ($useExcel) {
    Import-Module ImportExcel -ErrorAction Stop

    if (-not $OutputPath) { $OutputPath = Join-Path $outputDir ("{0}_{1}.xlsx" -f $stem, $ts) }
    if ($OutputPath -notmatch '\.xlsx$') { $OutputPath = [System.IO.Path]::ChangeExtension($OutputPath, 'xlsx') }
    # Export-Excel appends to an existing workbook, which would stack a second run on
    # top of the first - so a file at this exact path is replaced, not extended.
    if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

    # -TableName gives each sheet a real Excel table, which is what carries the filter
    # dropdowns and the banded rows - so no separate -AutoFilter is needed.
    $common = @{
        Path         = $OutputPath
        AutoSize     = $true
        BoldTopRow   = $true
        FreezeTopRow = $true
        TableStyle   = 'Medium2'
    }
    $overviewSorted | Export-Excel @common -WorksheetName 'Overzicht' -TableName 'Overzicht'
    $membersSorted  | Export-Excel @common -WorksheetName 'Leden'     -TableName 'Leden'

    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
} else {
    if (-not $OutputPath) { $OutputPath = Join-Path $outputDir ("{0}_{1}.csv" -f $stem, $ts) }
    $base        = Join-Path ([System.IO.Path]::GetDirectoryName($OutputPath)) ([System.IO.Path]::GetFileNameWithoutExtension($OutputPath))
    $overviewCsv = "$base-overzicht.csv"
    $membersCsv  = "$base-leden.csv"

    # Delimiter follows the local culture, otherwise a Dutch Excel opens the file as
    # one column of comma soup.
    $overviewSorted | Export-Csv -Path $overviewCsv -NoTypeInformation -Encoding UTF8 -UseCulture
    $membersSorted  | Export-Csv -Path $membersCsv  -NoTypeInformation -Encoding UTF8 -UseCulture

    Write-Host "  Report saved: $overviewCsv" -ForegroundColor Green
    Write-Host "  Report saved: $membersCsv"  -ForegroundColor Green
}

$memberRows = @($membersSorted | Where-Object { $_.'E-mailadres lid' }).Count
Write-Host ""
if ($Member) {
    Write-Host ("  {0} is a member of {1} list(s) - {2} member row(s) exported." -f $Member, $overviewSorted.Count, $memberRows) -ForegroundColor Cyan
} else {
    Write-Host ("  {0} list(s) - {1} member row(s) exported." -f $overviewSorted.Count, $memberRows) -ForegroundColor Cyan
}
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
