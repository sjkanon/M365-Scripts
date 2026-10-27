#Requires -Version 5.1
<#
.SYNOPSIS
    Create one security group per printer from a printer x user matrix - in
    Active Directory or in Entra ID - fill it with the users ticked for that
    printer, and optionally write those groups into printers.json as the queue's
    permissions, so printing runs through the groups.

.DESCRIPTION
    The matrix is a tab- (or semicolon-, comma-) separated file the way it comes out
    of Excel: the first column holds the printer names, the header row the users,
    and any non-empty cell (x) means "this user prints to this printer".

                            Jan Peeters   An Claes   ...
        Office - Duplex     x             x
        Office - Labels                   x

    A header without a cell above the printer column, as Excel copies it, is fine.
    A user is a UPN (contains @), DOMAIN\account (AD), or a display name;
    "Name (DEPARTMENT)" also matches the display name "Name". A name that matches
    no user, or more than one, is reported and left out - the group is still
    created for the others.

    Per printer:

      1. The group <Prefix><printer name> - spaces removed, so "Office - Duplex"
         becomes SG-U-PRINTM-Office-Duplex. Created as a security group when it
         does not exist yet, found by its name when it does.
      2. The ticked users who are not a member yet are added. Members the matrix
         does not tick are reported and left alone, unless -RemoveExtraMembers.
      3. With -ConfigPath: the printer's entry in printers.json gets
         "permissions": [ { "principal": <the group>, "access": "print" } ], so
         Install-Printer gives that group print access and takes it from Everyone.
         Entries with manage or manageDocuments that are not Everyone are kept.
         The file is backed up to printers.json.bak first.

    Default is a dry run: it reports what it would create and add. -Apply changes
    the directory and the JSON.

    AD or Entra ID?
    ---------------
    A printer ACL is checked against the user's Windows sign-in token, so the
    group has to be one that token carries:

      -Directory AD     Users sign in with an Active Directory account - a
                        domain-joined or hybrid-joined server or session host,
                        a classic print server. That token holds AD groups only.
                        The groups are created in -OU (pick one Entra Connect
                        synchronises, and they appear in Entra ID too) and land
                        in printers.json as DOMAIN\SG-U-PRINTM-....
                        Needs the ActiveDirectory module (RSAT), on a
                        domain-joined machine; Windows PowerShell 5.1 is enough.
      -Directory Entra  Users sign in with an Entra ID account on an Entra-joined
                        device (Entra-joined AVD session hosts included). The
                        groups are cloud groups and land in printers.json as
                        their SID (S-1-12-1-..., computed from the object id) -
                        the only way Windows names one. Needs PowerShell 7 and
                        Microsoft Graph. Warns when the matrix users are synced
                        from AD, because then -Directory AD is the one that works.

    Either way a new member only gets access at their next sign-in: a token
    holds the groups it was issued with.

.PARAMETER MatrixPath
    The matrix file (tab, semicolon or comma separated).

.PARAMETER Directory
    AD or Entra (default). See "AD or Entra ID?" above.

.PARAMETER OU
    -Directory AD: the distinguished name of the OU new groups are created in,
    e.g. "OU=Printers,OU=Groups,DC=bvl,DC=local". Required as soon as a group has
    to be created. Pick one in Entra Connect's sync scope.

.PARAMETER Prefix
    Prefix of the group names (default: SG-U-PRINTM-).

.PARAMETER ConfigPath
    printers.json (from Install-Printer or Export-Printer) to write the groups into
    as permissions. Printers in the matrix that the JSON does not have are reported.

.PARAMETER RemoveExtraMembers
    Also remove members the matrix does not tick, so the group matches it exactly.

.PARAMETER TenantId
    -Directory Entra: tenant ID or domain. Defaults to the GDAP customer tenant
    ($global:cid / $env:M365_CUSTOMER_TENANTID), else the tenant you sign in to -
    or the one an existing Graph session in this window is already in. The run
    names the tenant (or the AD domain) first, and -Apply asks before creating
    anything there (-Confirm:$false skips that question).

.PARAMETER ClientId
    -Directory Entra: app registration for app-only sign-in, with -CertificateThumbprint.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for -ClientId.

.PARAMETER AppOnly
    -Directory Entra: app-only sign-in with ClientId and CertificateThumbprint from
    graph.appid.json.

.PARAMETER Apply
    Create the groups, change their members and write the JSON. Without it, only
    report.

.EXAMPLE
    .\New-PrinterGroup.ps1 -MatrixPath .\printer-matrix.tsv -Directory AD

    Hybrid / on-premises: what would be created and added in AD - nothing changes.

.EXAMPLE
    .\New-PrinterGroup.ps1 -MatrixPath .\printer-matrix.tsv -Directory AD -OU 'OU=Printers,OU=Groups,DC=contoso,DC=local' -ConfigPath .\printers.json -Apply

    Create the AD groups, add the members and write DOMAIN\SG-U-PRINTM-... into
    printers.json as each queue's print permission.

.EXAMPLE
    .\New-PrinterGroup.ps1 -MatrixPath .\printer-matrix.tsv -TenantId contoso.com -ConfigPath .\printers.json -Apply

    Entra-joined devices: cloud groups, and their SIDs into printers.json.

.NOTES
    Author          : Sjoerd Kanon
    -Directory AD   : ActiveDirectory module, rights to create groups in -OU
    -Directory Entra: PowerShell 7, Microsoft.Graph.Groups/Users,
                      scopes Group.ReadWrite.All, User.Read.All
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)] [string] $MatrixPath,
    [ValidateSet('Entra', 'AD')]
    [string] $Directory = 'Entra',
    [string] $OU,
    [string] $Prefix = 'SG-U-PRINTM-',
    [string] $ConfigPath,
    [switch] $RemoveExtraMembers,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly,
    [switch] $Apply
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -Confirm:$false skips the tenant/domain question, for a run that already knows.
$confirmSuppressed = $PSBoundParameters.ContainsKey('Confirm') -and -not $PSBoundParameters['Confirm']

function Write-Ok   { param([string] $Message) Write-Host "  [ OK ] $Message" -ForegroundColor Green }
function Write-Skip { param([string] $Message) Write-Host "  [SKIP] $Message" -ForegroundColor DarkGray }
function Write-Warn { param([string] $Message) Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
function Write-Bad  { param([string] $Message) Write-Host "  [FAIL] $Message" -ForegroundColor Red }
function Write-Plan { param([string] $Message) Write-Host "  [PLAN] $Message" -ForegroundColor Magenta }

function Get-PropertyValue {
    <# Property access that returns $null instead of tripping Set-StrictMode. #>
    param($Object, [string] $Name)
    if ($null -eq $Object) { return $null }
    if ($Object.PSObject.Properties.Name -notcontains $Name) { return $null }
    return $Object.$Name
}

function Read-PrinterMatrix {
    <#
        The matrix as { Users; Printers = @({ Name; Users }) }. The delimiter is the
        one the header line holds most of. Excel copies a block without the empty
        top-left cell, so a header one cell shorter than the rows is shifted.
    #>
    param([Parameter(Mandatory)] [string] $Path)

    $lines = @(Get-Content -LiteralPath $Path -Encoding UTF8 | Where-Object { $_.Trim() })
    if ($lines.Count -lt 2) { throw "$Path needs a header row and at least one printer row" }

    $delimiter = @("`t", ';', ',') | Sort-Object { -($lines[0].Split($_).Count) } | Select-Object -First 1
    $header    = @($lines[0].Split($delimiter) | ForEach-Object { $_.Trim() })
    $width     = ($lines | Select-Object -Skip 1 | ForEach-Object { $_.Split($delimiter).Count } | Measure-Object -Maximum).Maximum
    $users     = if ($header.Count -lt $width) { $header } else { @($header | Select-Object -Skip 1) }
    $users     = @($users | Where-Object { $_ })

    $printers = foreach ($line in ($lines | Select-Object -Skip 1)) {
        $cells = $line.Split($delimiter)
        $name  = $cells[0].Trim()
        if (-not $name) { continue }
        $ticked = for ($i = 0; $i -lt $users.Count; $i++) {
            if ($i + 1 -lt $cells.Count -and $cells[$i + 1].Trim()) { $users[$i] }
        }
        [PSCustomObject]@{ Name = $name; Users = @($ticked) }
    }
    return [PSCustomObject]@{ Users = $users; Printers = @($printers) }
}

function Get-GroupName {
    param([Parameter(Mandatory)] [string] $Printer)
    return $Prefix + ($Printer -replace '\s+', '')
}

function ConvertTo-EntraSid {
    <# An Entra ID object id as the SID Windows gives it: S-1-12-1 and the GUID as four 32-bit numbers. #>
    param([Parameter(Mandatory)] [string] $ObjectId)
    $b = [guid]::new($ObjectId).ToByteArray()
    return 'S-1-12-1-{0}-{1}-{2}-{3}' -f [BitConverter]::ToUInt32($b, 0), [BitConverter]::ToUInt32($b, 4),
                                        [BitConverter]::ToUInt32($b, 8), [BitConverter]::ToUInt32($b, 12)
}

# -- Directory operations -------------------------------------------------------
# Everything below the matrix works on { Id; DisplayName; Label } for users and
# members, so the two directories differ only in these functions. A user's Id is
# what the group's member list returns for them: the object id in Entra ID, the
# SID in AD.

function Find-DirectoryUser {
    <# All users matching one way of naming them: 'upn', 'account' or 'display'. #>
    param([string] $By, [string] $Value)
    $q = $Value -replace "'", "''"
    if ($Directory -eq 'AD') {
        $filter = switch ($By) { 'upn' { "UserPrincipalName -eq '$q'" } 'account' { "SamAccountName -eq '$q'" } default { "DisplayName -eq '$q'" } }
        foreach ($u in @(Get-ADUser -Filter $filter -Properties DisplayName, UserPrincipalName -Server $script:dc)) {
            [PSCustomObject]@{ Id = $u.SID.Value; DisplayName = $(if ($u.DisplayName) { $u.DisplayName } else { $u.Name })
                               Label = "$($script:netbios)\$($u.SamAccountName)"; Synced = $false }
        }
    } else {
        $props = 'Id', 'DisplayName', 'UserPrincipalName', 'OnPremisesSyncEnabled'
        $found = switch ($By) {
            'upn'   { try { @(Get-MgUser -UserId $Value -Property $props -ErrorAction Stop) } catch { @() } }
            default { @(Get-MgUser -Filter "displayName eq '$q'" -Property $props -All -ErrorAction Stop) }
        }
        foreach ($u in $found) {
            [PSCustomObject]@{ Id = $u.Id; DisplayName = $u.DisplayName; Label = $u.UserPrincipalName
                               Synced = ((Get-PropertyValue $u 'OnPremisesSyncEnabled') -eq $true) }
        }
    }
}

function Resolve-MatrixUser {
    <#
        A matrix header to one user: a UPN or DOMAIN\account as given, a display
        name as is, then without a trailing "(...)". $null with a reason when it is
        none or more than one.
    #>
    param([Parameter(Mandatory)] [string] $Name)

    # The @() around the whole if: one try on its own would otherwise be unrolled
    # into its two strings.
    $tries = @(if ($Name -match '@') { ,@('upn', $Name) }
               elseif ($Name -match '^[^\\]+\\(.+)$' -and $Directory -eq 'AD') { ,@('account', $Matches[1]) }
               else { @(@('display', $Name), @('display', ($Name -replace '\s*\([^)]*\)\s*$', '').Trim())) })
    $seen = @{}
    foreach ($try in $tries) {
        if ($seen.ContainsKey($try[1])) { continue }
        $seen[$try[1]] = $true
        $hits = @(Find-DirectoryUser -By $try[0] -Value $try[1])
        if ($hits.Count -eq 1) { return [PSCustomObject]@{ User = $hits[0]; Reason = $null } }
        if ($hits.Count -gt 1) {
            return [PSCustomObject]@{ User = $null; Reason = "$($hits.Count) users are called '$($try[1])' ($(($hits | ForEach-Object Label) -join ', ')) - put the UPN in the header instead" }
        }
    }
    return [PSCustomObject]@{ User = $null; Reason = 'no user by that name - put the UPN in the header instead' }
}

function Find-DirectoryGroup {
    param([Parameter(Mandatory)] [string] $Name)
    $q = $Name -replace "'", "''"
    if ($Directory -eq 'AD') { return @(Get-ADGroup -Filter "SamAccountName -eq '$q' -or Name -eq '$q'" -Server $script:dc) | Select-Object -First 1 }
    return @(Get-MgGroup -Filter "displayName eq '$q'" -ErrorAction Stop) | Select-Object -First 1
}

function New-DirectoryGroup {
    param([Parameter(Mandatory)] [string] $Name, [Parameter(Mandatory)] [string] $Printer)
    $description = "Print access to the printer '$Printer'"
    if ($Directory -eq 'AD') {
        return New-ADGroup -Name $Name -SamAccountName $Name -GroupScope Global -GroupCategory Security -Path $OU `
                           -Description $description -Server $script:dc -PassThru
    }
    $nickname = ($Name -replace '[^A-Za-z0-9\-_.]', '')
    if ($nickname.Length -gt 64) { $nickname = $nickname.Substring(0, 64) }
    return New-MgGroup -DisplayName $Name -MailEnabled:$false -MailNickname $nickname -SecurityEnabled:$true `
                       -Description $description -ErrorAction Stop
}

function Get-DirectoryGroupMember {
    param([Parameter(Mandatory)] $Group)
    if ($Directory -eq 'AD') {
        foreach ($m in @(Get-ADGroupMember -Identity $Group -Server $script:dc)) {
            [PSCustomObject]@{ Id = $m.SID.Value; DisplayName = $m.Name }
        }
    } else {
        foreach ($m in @(Get-MgGroupMember -GroupId $Group.Id -All -ErrorAction Stop)) {
            [PSCustomObject]@{ Id = $m.Id; DisplayName = $m.AdditionalProperties['displayName'] }
        }
    }
}

function Add-DirectoryGroupMember {
    param([Parameter(Mandatory)] $Group, [Parameter(Mandatory)] $User)
    if ($Directory -eq 'AD') { Add-ADGroupMember -Identity $Group -Members $User.Id -Server $script:dc }
    else { New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $User.Id -ErrorAction Stop }
}

function Remove-DirectoryGroupMember {
    param([Parameter(Mandatory)] $Group, [Parameter(Mandatory)] $Member)
    if ($Directory -eq 'AD') { Remove-ADGroupMember -Identity $Group -Members $Member.Id -Server $script:dc -Confirm:$false }
    else { Remove-MgGroupMemberDirectoryObjectByRef -GroupId $Group.Id -DirectoryObjectId $Member.Id -ErrorAction Stop }
}

function Get-GroupPrincipal {
    <# How Install-Printer names the group: DOMAIN\name in AD, the SID for Entra ID. #>
    param([Parameter(Mandatory)] [string] $Name, $Group)
    if ($Directory -eq 'AD') { return "$($script:netbios)\$Name" }
    if ($Group) { return ConvertTo-EntraSid $Group.Id }
    return $null
}

# -- Matrix ---------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $MatrixPath)) { throw "Matrix not found: $MatrixPath" }
$matrix = Read-PrinterMatrix -Path $MatrixPath
if ($ConfigPath -and -not (Test-Path -LiteralPath $ConfigPath)) { throw "printers.json not found: $ConfigPath" }

# An AD group's name (CN) stops at 64 characters; a printer name rarely does.
$tooLong = @($matrix.Printers | ForEach-Object { Get-GroupName $_.Name } | Where-Object { $_.Length -gt 64 })
if ($Directory -eq 'AD' -and $tooLong.Count -gt 0) { throw "Group names longer than 64 characters, which AD refuses: $($tooLong -join ', ') - shorten the printer name or -Prefix" }

Write-Host ''
Write-Host "  Mode: $(if ($Apply) { 'APPLY - groups and members are changed' } else { 'PREVIEW - nothing is changed' }) - $Directory" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'Cyan' })
Write-Host "  Matrix: $($matrix.Printers.Count) printer(s), $($matrix.Users.Count) user(s)" -ForegroundColor Cyan

# -- Connect --------------------------------------------------------------------
$graph = $null
if ($Directory -eq 'AD') {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $adDomain       = Get-ADDomain -ErrorAction Stop
        $script:netbios = $adDomain.NetBIOSName
        $script:dc      = $adDomain.PDCEmulator   # one DC for the whole run: no replication lag between create and add
    } catch {
        Write-Bad "Active Directory is not reachable: $($_.Exception.Message). Run this on a domain-joined machine with RSAT (ActiveDirectory module)."
        exit 1
    }
    $where = "domain $($adDomain.DNSRoot) ($($script:netbios)) via $($script:dc)"
} else {
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        Write-Bad '-Directory Entra needs PowerShell 7 (pwsh). For AD groups use -Directory AD, which runs in Windows PowerShell too.'
        exit 1
    }
    . (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')
    try {
        $graph = Connect-M365Graph -Scopes 'Group.ReadWrite.All', 'User.Read.All' -TenantId $TenantId `
            -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly
    } catch {
        Write-Bad "Could not connect to Microsoft Graph: $($_.Exception.Message)"
        exit 1
    }
}

$exitCode = 0
try {
    # -- Where ------------------------------------------------------------------
    # Without -TenantId an existing Graph session is reused as it is - often the
    # partner's own tenant. So the target is named before anything is read, and
    # confirmed before anything is created.
    if ($Directory -eq 'Entra') {
        $org    = @(Get-MgOrganization -Property Id, DisplayName, VerifiedDomains -ErrorAction Stop) | Select-Object -First 1
        $domain = @($org.VerifiedDomains | Where-Object IsDefault | ForEach-Object Name) | Select-Object -First 1
        $where  = "tenant $($org.DisplayName) ($domain, $($org.Id)) - signed in as $($graph.Account)"
    }
    Write-Host "  Target: $where" -ForegroundColor Cyan
    if ($Directory -eq 'AD') { Write-Host "  OU for new groups: $(if ($OU) { $OU } else { '(not given - needed as soon as a group has to be created)' })" -ForegroundColor Cyan }
    if ($Apply -and -not $confirmSuppressed -and [Environment]::UserInteractive) {
        $answer = Read-Host "  Create the groups there? [y/N]"
        if ($answer -notmatch '^[Yy]') { throw 'Cancelled - wrong target? -TenantId (Entra) or running on the right domain (AD) chooses it.' }
    }
    if ($Directory -eq 'AD' -and $OU) {
        try { $null = Get-ADOrganizationalUnit -Identity $OU -Server $script:dc } catch { throw "OU not found: $OU" }
    }
    Write-Host ''

    # -- Users ------------------------------------------------------------------
    Write-Host '  Users' -ForegroundColor Cyan
    $resolved = @{}
    foreach ($name in $matrix.Users) {
        $r = Resolve-MatrixUser -Name $name
        if ($r.User) {
            $resolved[$name] = $r.User
            Write-Ok "$name -> $($r.User.Label)"
        } else {
            Write-Bad "${name}: $($r.Reason)"
            $exitCode = 1
        }
    }

    # Synced users sign in to domain-joined or hybrid machines with their AD
    # account, and that token holds AD groups only - a cloud group in a printer
    # ACL then shuts them out. Worth saying before twenty groups are created.
    $synced = @($resolved.Values | Where-Object { $_.Synced })
    if ($synced.Count -gt 0) {
        Write-Host ''
        Write-Warn ("{0} of {1} user(s) are synced from Active Directory (hybrid). A cloud group does not reach their Windows token on a domain-joined or hybrid-joined machine, so as a printer permission it would lock them out - use -Directory AD for those machines" -f $synced.Count, $resolved.Count)
        if ($ConfigPath) { Write-Warn '-ConfigPath is given: the queues would only accept these cloud groups. Only continue if the target machine is Entra-joined and users sign in with Entra ID' }
    }

    # -- Groups -----------------------------------------------------------------
    $principals = [ordered]@{}
    foreach ($printer in $matrix.Printers) {
        $groupName = Get-GroupName $printer.Name
        Write-Host ''
        Write-Host "  $groupName" -ForegroundColor Cyan

        $group = Find-DirectoryGroup -Name $groupName
        if ($group) {
            Write-Skip "exists ($(if ($Directory -eq 'AD') { $group.DistinguishedName } else { $group.Id }))"
        } elseif (-not $Apply) {
            Write-Plan "create the security group$(if ($Directory -eq 'AD' -and $OU) { " in $OU" })"
        } elseif ($Directory -eq 'AD' -and -not $OU) {
            throw "-OU is needed to create $groupName - give the OU (in Entra Connect's sync scope) for the new groups"
        } elseif ($PSCmdlet.ShouldProcess($groupName, "Create $Directory security group")) {
            $group = New-DirectoryGroup -Name $groupName -Printer $printer.Name
            Write-Ok "created ($(if ($Directory -eq 'AD') { $group.DistinguishedName } else { $group.Id }))"
        }

        $wanted    = @($printer.Users | Where-Object { $resolved.ContainsKey($_) } | ForEach-Object { $resolved[$_] })
        $members   = @(if ($group) { Get-DirectoryGroupMember -Group $group })
        $memberIds = @($members | ForEach-Object Id)
        $wantedIds = @($wanted | ForEach-Object Id)

        foreach ($user in $wanted) {
            if ($user.Id -in $memberIds) { Write-Skip "$($user.DisplayName) is a member"; continue }
            if (-not $Apply -or -not $group) { Write-Plan "add $($user.DisplayName)"; continue }
            if ($PSCmdlet.ShouldProcess($groupName, "Add $($user.Label)")) {
                Add-DirectoryGroupMember -Group $group -User $user
                Write-Ok "added $($user.DisplayName)"
            }
        }
        foreach ($member in $members) {
            if ($member.Id -in $wantedIds) { continue }
            if (-not $RemoveExtraMembers) { Write-Warn "$($member.DisplayName) is a member but not ticked in the matrix - left alone (-RemoveExtraMembers removes it)"; continue }
            if (-not $Apply) { Write-Plan "remove $($member.DisplayName)"; continue }
            if ($PSCmdlet.ShouldProcess($groupName, "Remove $($member.DisplayName)")) {
                Remove-DirectoryGroupMember -Group $group -Member $member
                Write-Ok "removed $($member.DisplayName)"
            }
        }
        if ($wanted.Count -eq 0) { Write-Warn 'nobody ticked (or resolved) for this printer - the group stays empty' }

        $principal = Get-GroupPrincipal -Name $groupName -Group $group
        if ($principal) {
            $principals[$printer.Name] = [PSCustomObject]@{ Principal = $principal; Exists = [bool] $group }
            if ($Directory -eq 'Entra') { Write-Skip "SID $principal" }
        }
    }

    # -- printers.json ----------------------------------------------------------
    if ($ConfigPath) {
        Write-Host ''
        Write-Host "  $ConfigPath" -ForegroundColor Cyan
        $config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $byName = @{}
        foreach ($p in @($config.printers)) { $byName[$p.name] = $p }

        $changed = 0
        foreach ($printer in $matrix.Printers) {
            if (-not $byName.ContainsKey($printer.Name)) { Write-Warn "'$($printer.Name)' is not in the JSON"; continue }
            if (-not $principals.Contains($printer.Name)) { Write-Plan "'$($printer.Name)': permissions once its group exists (-Apply)"; continue }

            # Install-Printer fails closed on a group it cannot resolve - the queue
            # is then not created at all - so a group that does not exist yet is
            # only written by a run that creates it.
            $target = $principals[$printer.Name]
            if (-not $target.Exists -and -not $Apply) { Write-Plan "'$($printer.Name)': print -> $($target.Principal) once the group exists (-Apply)"; continue }

            $entry = $byName[$printer.Name]
            # Keep who manages the queue; replace who may print on it.
            $keep = @(if ($entry.PSObject.Properties.Name -contains 'permissions') {
                $entry.permissions | Where-Object { $_.access -ne 'print' -and $_.principal -ne 'Everyone' -and $_.principal -ne $target.Principal }
            })
            $new = @([PSCustomObject]@{ principal = $target.Principal; access = 'print' }) + $keep
            $entry | Add-Member -NotePropertyName permissions -NotePropertyValue $new -Force
            $changed++
            Write-Ok "'$($printer.Name)': print -> $($target.Principal)$(if ($keep.Count) { " (kept: $(($keep | ForEach-Object { "$($_.principal) $($_.access)" }) -join ', '))" })"
        }

        if ($changed -gt 0 -and $Apply -and $PSCmdlet.ShouldProcess($ConfigPath, "Write permissions for $changed printer(s)")) {
            Copy-Item -LiteralPath $ConfigPath -Destination "$ConfigPath.bak" -Force
            [IO.File]::WriteAllText((Resolve-Path -LiteralPath $ConfigPath).ProviderPath, ($config | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding $false))
            Write-Ok "written ($changed printer(s)), previous version in $ConfigPath.bak"
        } elseif ($changed -gt 0) {
            Write-Plan "write permissions for $changed printer(s) (-Apply)"
        }
    }

    Write-Host ''
    if (-not $Apply) { Write-Host '  Preview only - rerun with -Apply to make these changes.' -ForegroundColor Yellow }
} catch {
    Write-Bad "Aborted: $($_.Exception.Message)"
    $exitCode = 1
} finally {
    if ($graph) { Disconnect-M365Graph $graph }
}
Write-Host ''
exit $exitCode
