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
      Leden      one row per member: which list, who, which address, what kind of
                 member, and the external address when the member is a contact

    With -Member both sheets gain a column for the filter: 'Treffers' (how many
    members of this list matched) and 'Treffer op' (the address this member matched
    on, empty when it did not). With -Recurse they gain 'Aantal personen' and
    'Via groep'.

    The sheet headers are Dutch because the workbook is what goes to the customer;
    the script itself stays English like the rest of the repo.

    Without ImportExcel the report is written as two CSV files instead, so the
    script never fails just because a module is missing.

.PARAMETER Group
    One list (name, alias or e-mail). Without it every list is reported.

.PARAMETER Member
    Narrow the report to the lists that contain a given member. Takes either:

      one address     jan@contoso.com    which lists is this person on?
      one domain      @be.verizon.com    members on exactly that domain
                      (also "be.verizon.com", "*@be.verizon.com")
      a domain tree   *.verizon.com      verizon.com AND every subdomain of it -
                      (also ".verizon.com",   be.verizon.com, us.verizon.com, ...
                       "*@*.verizon.com")

    Without the leading "*." the match is on that one domain, so @be.verizon.com
    deliberately does not match @notbe.verizon.com or @us.verizon.com. With it, the
    apex and every subdomain are in scope. The run says which of the two it is doing.

    An address is resolved server-side and stays fast in a large tenant. A domain
    cannot be: Exchange has no filter for "member whose address ends in @x", so every
    list is read and then filtered, and only the lists with a hit are kept.

    Matching covers the primary address, every alias, and - for mail contacts and
    mail users - ExternalEmailAddress, which is where an external party's real
    address lives. The matched lists are exported in full, so the customer sees who
    else is on them; a 'Treffer op' column names the address each hit matched on -
    which is the only way to see why someone matched on an alias.

    Direct membership only, unless -Recurse is used: without it a person inside a
    nested group is not a match, and the nested group itself shows up as one member
    row. See -Recurse - for a domain filter this is usually the difference between a
    complete answer and a confident wrong one.

.PARAMETER Recurse
    Expand nested groups, so the report lists the people who actually receive the mail
    instead of the group standing in for them.

    Exchange only ever returns DIRECT members. A list containing another list therefore
    reports that list as one member and never the people inside it - so by default
    someone who only receives mail through a nested group is invisible, and -Member
    reports "no hits" on a list that does deliver to them. For a domain filter that is
    usually the difference between a complete answer and a confident wrong one.

    Costs one extra query per nested group. A group already expanded is not expanded
    again, which is also what keeps a membership cycle (A contains B, B contains A)
    from recursing forever; nesting deeper than 20 levels is reported and left alone.

    The nested group itself stays in the report as its own row, so the structure is
    still visible. 'Via groep' names the group a person came in through, empty for a
    direct member; someone reachable by several routes gets one row with the routes
    joined. 'Aantal leden' keeps counting direct members - that is the number Exchange
    and the EAC show - and 'Aantal personen' counts the real recipients reached.

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
    .\Get-DistributionGroupMembers.ps1 -Member "@be.verizon.com"
    Every list with a member on that domain, with the matching address of each hit in
    the 'Treffer op' column.

.EXAMPLE
    .\Get-DistributionGroupMembers.ps1 -Member "*.verizon.com" -Recurse
    Everything Verizon: verizon.com and every subdomain, nested lists expanded.

.EXAMPLE
    .\Get-DistributionGroupMembers.ps1 -Member "@be.verizon.com" -Recurse
    The same, but also finding the people who sit inside a nested group. This is the
    form to use when the question is "does anything still reach that domain?".

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
    [switch] $Recurse,
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

# Which recipient types are themselves a group, and which cmdlet reads their members.
$groupKinds = @{
    'MailUniversalDistributionGroup' = 'Static'
    'MailUniversalSecurityGroup'     = 'Static'
    'MailNonUniversalGroup'          = 'Static'
    'RoomList'                       = 'Static'
    'DynamicDistributionGroup'       = 'Dynamic'
    'GroupMailbox'                   = 'Unified'
}

function Get-DirectMember {
    param($List, [string] $Kind)
    switch ($Kind) {
        'Dynamic' { @(Get-DynamicMember -DynamicGroup $List) }
        'Unified' { @(Get-UnifiedGroupLinks -Identity $List.Identity -LinkType Members -ResultSize Unlimited -ErrorAction Stop) }
        default   { @(Get-DistributionGroupMember -Identity $List.Identity -ResultSize Unlimited -ErrorAction Stop) }
    }
}

# Walk nested groups. Exchange only ever hands back DIRECT members, so without this a
# list containing another list reports that list as one member and never the people
# inside it - and a filter then reports "no hits" on a list that does deliver mail to
# the person you asked about.
$script:MaxNestDepth = 20

function Expand-Member {
    param($Direct, [hashtable] $Seen, [string] $Via = '', [int] $Depth = 0)
    foreach ($m in $Direct) {
        [PSCustomObject]@{ Recipient = $m; Via = $Via }

        $kind = $groupKinds[[string]$m.RecipientTypeDetails]
        if (-not $kind) { continue }

        # A group already expanded is skipped, which is also what stops a membership
        # cycle (A contains B, B contains A) from recursing forever.
        $dn = [string]$m.DistinguishedName
        if (-not $dn -or $Seen.ContainsKey($dn)) { continue }
        $Seen[$dn] = $true

        if ($Depth -ge $script:MaxNestDepth) {
            Write-Host "  [WARN] Nesting deeper than $script:MaxNestDepth levels at '$($m.DisplayName)' - not expanded." -ForegroundColor Yellow
            continue
        }

        try {
            Expand-Member -Direct (Get-DirectMember -List $m -Kind $kind) -Seen $Seen `
                          -Via $m.DisplayName -Depth ($Depth + 1)
        } catch {
            Write-Host "  [WARN] Nested group '$($m.DisplayName)': $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

# Every address a member can be reached on. An external party is usually a mail
# contact, and then the address that matters is ExternalEmailAddress - the primary
# SMTP is an internal placeholder. Missing that is missing exactly the members a
# domain filter is used to find.
function Get-AllAddress {
    param($Recipient)
    $raw = @($Recipient.PrimarySmtpAddress, $Recipient.ExternalEmailAddress) + @($Recipient.EmailAddresses)
    foreach ($a in $raw) {
        if (-not $a) { continue }
        $s = [string]$a
        # Proxy addresses are prefixed: SMTP: (primary), smtp: (alias), X500:, SIP: ...
        # Only the SMTP ones are e-mail addresses; the rest would never match anyway.
        if ($s -match '^(?i)smtp:') { $s = $s -replace '^(?i)smtp:', '' }
        elseif ($s -match '^[A-Za-z0-9]+:')  { continue }
        if ($s) { $s }
    }
}

# ── Filter mode ───────────────────────────────────────────────────────────────
# -Member takes one address, one domain, or a domain and everything under it:
#
#   jan@contoso.com                            one address
#   @be.verizon.com  be.verizon.com  *@be...   that domain exactly
#   *.verizon.com    .verizon.com   *@*.ver…   verizon.com AND every subdomain
#
# Anything carrying a wildcard is a domain - an address never does.
$filterMode   = 'None'
$filterDomain = $null      # the bare domain, no @ and no leading dot
$filterSubs   = $false     # also match subdomains of it
$filterLabel  = $null      # how the filter is written back to the user

if ($Member) {
    $raw = $Member.Trim()

    if ($raw -match '^[^@*]+@[^@*]+$') {
        $filterMode = 'Address'
    } else {
        $filterMode = 'Domain'

        # Take what sits after the last @, so "*@*.verizon.com" reduces to "*.verizon.com".
        $dom = if ($raw -match '@') { ($raw -split '@')[-1] } else { $raw }
        $dom = $dom.TrimStart('*')
        $filterSubs = $dom.StartsWith('.')
        $dom = $dom.TrimStart('.')

        # A wildcard anywhere else is not a thing this supports, and passing it into
        # -like would quietly match something nobody asked for.
        $filterDomain = [System.Management.Automation.WildcardPattern]::Escape($dom)
        $filterLabel  = if ($filterSubs) { "$dom and its subdomains" } else { "@$dom" }
    }
}

# A member row, built the same way for a real member and for the placeholder an empty
# list gets - so every row in the sheet keeps the same columns.
function New-MemberRow {
    param($List, $Name, $Address = '', $External = '', $Type = '', $Via = '', $Hit = '')
    $row = [ordered]@{
        'Lijst'             = $List.DisplayName
        'E-mailadres lijst' = [string]$List.PrimarySmtpAddress
        'Type lijst'        = Get-FriendlyType $List.RecipientTypeDetails
        'Lid'               = $Name
        'E-mailadres lid'   = $Address
        'Extern adres'      = $External
        'Type lid'          = $Type
    }
    if ($Recurse)               { $row['Via groep']  = $Via }
    if ($filterMode -ne 'None') { $row['Treffer op'] = $Hit }
    [PSCustomObject]$row
}

# Returns the address that matched, not just $true. A member can match on an alias
# that appears nowhere else in the report, and "Treffer: Ja" on a row whose visible
# address is sara@contoso.com only raises the question why.
function Get-MemberMatch {
    param($Recipient)
    switch ($filterMode) {
        'Domain' {
            foreach ($a in Get-AllAddress $Recipient) {
                # "*@x" is the domain itself; "*.x" is anything under it. Matching on
                # the dot is what keeps @notbe.verizon.com out of a @be.verizon.com run.
                if ($a -like "*@$filterDomain")                 { return $a }
                if ($filterSubs -and $a -like "*.$filterDomain") { return $a }
            }
        }
        'Address' {
            foreach ($a in Get-AllAddress $Recipient) { if ($a -eq $Member) { return $a } }
            if ($script:FilterDn -and $Recipient.DistinguishedName -eq $script:FilterDn) {
                return [string]$Recipient.PrimarySmtpAddress
            }
        }
    }
    return $null
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

} elseif ($filterMode -eq 'Address') {
    # Resolve the address first: the membership filter matches on DN, not on SMTP.
    $recipient = Get-Recipient -Identity $Member -ErrorAction Stop
    $script:FilterDn = $recipient.DistinguishedName
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
    # Also the path for a domain filter: no server-side query can answer "has a member
    # whose address ends in @be.verizon.com", so every list is read and then filtered.
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
    if ($filterMode -eq 'Address') { Write-Host "  $Member is not a direct member of any distribution list." -ForegroundColor Yellow }
    else                           { Write-Host "  No distribution lists found." -ForegroundColor Yellow }
    if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
    return
}

if ($filterMode -eq 'Domain') {
    Write-Host "  Scanning $($lists.Count) list(s) for members on $filterLabel..." -ForegroundColor DarkGray
} else {
    Write-Host "  Reading members of $($lists.Count) list(s)..." -ForegroundColor DarkGray
}
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

    $directMembers = @()
    $owners        = ''

    try {
        $directMembers = Get-DirectMember -List $list -Kind $entry.Kind
        if ($entry.Kind -eq 'Unified') {
            $owners = (@(Get-UnifiedGroupLinks -Identity $list.Identity -LinkType Owners -ResultSize Unlimited -ErrorAction SilentlyContinue) |
                       ForEach-Object { $_.DisplayName }) -join '; '
        } else {
            $owners = ($list.ManagedBy | ForEach-Object { ($_ -split '/')[-1] }) -join '; '
        }
    } catch {
        Write-Host "  [WARN] $($list.DisplayName): $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Direct members become { Recipient, Via } either way, so one code path serves both
    # modes. Without -Recurse a nested group stays a single row, as Exchange reports it.
    if ($Recurse) {
        $seen = @{ ([string]$list.DistinguishedName) = $true }
        $walked = @(Expand-Member -Direct $directMembers -Seen $seen)

        # The same person can hang under several nested groups. One row each, with the
        # routes joined - a duplicate row per route reads as a data error to a customer.
        $byDn = [ordered]@{}
        foreach ($w in $walked) {
            $dn = [string]$w.Recipient.DistinguishedName
            if (-not $dn) { $dn = "$($w.Recipient.DisplayName)|$($w.Recipient.PrimarySmtpAddress)" }
            if (-not $byDn.Contains($dn)) { $byDn[$dn] = $w; continue }

            # Direct membership wins over any route: it is the plainer fact.
            $kept = $byDn[$dn]
            if (-not $kept.Via -or -not $w.Via) { continue }
            if (($kept.Via -split '; ') -notcontains $w.Via) { $kept.Via = "$($kept.Via); $($w.Via)" }
        }
        $groupMembers = @($byDn.Values)
    } else {
        $groupMembers = @($directMembers | ForEach-Object { [PSCustomObject]@{ Recipient = $_; Via = '' } })
    }

    # Which members the filter hits, keyed by DN, holding the address that matched. A
    # domain filter drops a list without a single hit; a filter on one address arrives
    # here already narrowed down by the server, so there this only fills the column.
    $matchHits = @{}
    if ($filterMode -ne 'None') {
        foreach ($gm in $groupMembers) {
            $hitAddress = Get-MemberMatch $gm.Recipient
            if ($hitAddress) { $matchHits[[string]$gm.Recipient.DistinguishedName] = $hitAddress }
        }
    }
    if ($filterMode -eq 'Domain' -and $matchHits.Count -eq 0) { continue }

    $internalOnly = if ($null -ne $list.RequireSenderAuthenticationEnabled) {
        if ($list.RequireSenderAuthenticationEnabled) { 'Ja' } else { 'Nee' }
    } else { '' }

    $listRow = [ordered]@{
        'Lijst'        = $list.DisplayName
        'E-mailadres'  = [string]$list.PrimarySmtpAddress
        'Type'         = Get-FriendlyType $list.RecipientTypeDetails
        'Aantal leden' = $directMembers.Count
    }
    if ($Recurse) {
        # Direct members are what Exchange and the EAC show, so that count stays put.
        # What people actually want to know when nesting is in play is how many real
        # recipients the list reaches, which is a different number.
        $listRow['Aantal personen'] = @($groupMembers | Where-Object { -not $groupKinds[[string]$_.Recipient.RecipientTypeDetails] }).Count
    }
    if ($filterMode -ne 'None') { $listRow['Treffers'] = $matchHits.Count }
    $listRow['Eigenaar(s)']              = $owners
    $listRow['Alias']                    = $list.Alias
    $listRow['Verborgen in adresboek']   = if ($list.HiddenFromAddressListsEnabled) { 'Ja' } else { 'Nee' }
    $listRow['Alleen interne afzenders'] = $internalOnly
    $listRow['Aangemaakt op']            = if ($list.WhenCreated) { (Get-Date $list.WhenCreated -Format 'dd-MM-yyyy') } else { '' }
    $overview.Add([PSCustomObject]$listRow)

    if ($groupMembers.Count -eq 0) {
        # An empty list is exactly what a customer wants to spot, so it gets a row of
        # its own instead of quietly missing from the member sheet.
        $members.Add((New-MemberRow -List $list -Name '(geen leden)'))
        continue
    }

    foreach ($w in ($groupMembers | Sort-Object { $_.Recipient.DisplayName })) {
        $m = $w.Recipient

        # A mail contact's real address lives in ExternalEmailAddress; its primary SMTP
        # is an internal placeholder, so both belong in the report.
        $external = [string]$m.ExternalEmailAddress -replace '^(?i)smtp:', ''
        if ($external -eq [string]$m.PrimarySmtpAddress) { $external = '' }

        $members.Add((New-MemberRow -List $list `
                                    -Name     $m.DisplayName `
                                    -Address  ([string]$m.PrimarySmtpAddress) `
                                    -External $external `
                                    -Type     (Get-FriendlyType $m.RecipientTypeDetails) `
                                    -Via      $w.Via `
                                    -Hit      ([string]$matchHits[[string]$m.DistinguishedName])))
    }
}
Write-Progress -Activity "Distribution lists" -Completed

$overviewSorted = @($overview | Sort-Object 'Lijst')
$membersSorted  = @($members  | Sort-Object 'Lijst', 'Lid')

# ── Console summary ───────────────────────────────────────────────────────────
if ($overviewSorted.Count -eq 0) {
    # Only a domain filter can get this far and end up empty: the lists existed, none
    # of them had a member on that domain. Writing an empty workbook would read as
    # "the report failed" rather than as the answer it is.
    Write-Host "  No distribution list has a member on $filterLabel." -ForegroundColor Yellow
    Write-Host ""
    if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
    return
}

$columns = @('Lijst', 'E-mailadres', 'Type', 'Aantal leden')
if ($Recurse)               { $columns += 'Aantal personen' }
if ($filterMode -ne 'None') { $columns += 'Treffers' }
$columns += 'Eigenaar(s)'
$overviewSorted | Format-Table $columns -AutoSize

# ── Report ────────────────────────────────────────────────────────────────────
$ts   = Get-Date -Format 'yyyyMMdd_HHmmss'
$stem = switch ($filterMode) {
    'Domain'  { "Distributielijsten_$(if ($filterSubs) { 'sub_' })$($filterDomain -replace '[^\w.-]', '_')" }
    'Address' { "Distributielijsten_$($Member -replace '[^\w.-]', '_')" }
    default   { 'Distributielijsten' }
}

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
$hitRows    = @($membersSorted | Where-Object { $_.'Treffer op' }).Count
Write-Host ""
switch ($filterMode) {
    'Domain' {
        Write-Host ("  {0} list(s) have a member on {1} - {2} matching member(s), {3} member row(s) exported." -f
                    $overviewSorted.Count, $filterLabel, $hitRows, $memberRows) -ForegroundColor Cyan
    }
    'Address' {
        Write-Host ("  {0} is a member of {1} list(s) - {2} member row(s) exported." -f
                    $Member, $overviewSorted.Count, $memberRows) -ForegroundColor Cyan
    }
    default {
        Write-Host ("  {0} list(s) - {1} member row(s) exported." -f $overviewSorted.Count, $memberRows) -ForegroundColor Cyan
    }
}
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
