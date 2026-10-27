#Requires -Version 7.0
<#
.SYNOPSIS
    Create one Entra ID security group per printer from a printer x user matrix,
    fill it with the users ticked for that printer, and optionally write those
    groups into printers.json as the queue's permissions.

.DESCRIPTION
    The matrix is a tab- (or semicolon-, comma-) separated file the way it comes out
    of Excel: the first column holds the printer names, the header row the users,
    and any non-empty cell (x) means "this user prints to this printer".

                            Jan Peeters   An Claes   ...
        Office - Duplex     x             x
        Office - Labels                   x

    A header without a cell above the printer column, as Excel copies it, is fine.
    A user is a UPN (contains @) or a display name; "Name (DEPARTMENT)" also matches
    the display name "Name". A name that matches no user, or more than one, is
    reported and left out - the group is still created for the others.

    Per printer:

      1. The group <Prefix><printer name> - spaces removed, so "Office - Duplex"
         becomes SG-U-PRINTM-Office-Duplex. Created as an assigned security group
         when it does not exist yet, found by its display name when it does.
      2. The ticked users who are not a member yet are added. Members the matrix
         does not tick are reported and left alone, unless -RemoveExtraMembers.
      3. With -ConfigPath: the printer's entry in printers.json gets
         "permissions": [ { "principal": "<group SID>", "access": "print" } ]. The
         SID is the Entra ID form (S-1-12-1-...) computed from the group's object
         id - the only way Windows names an Entra ID group in a printer ACL.
         Entries with manage or manageDocuments that are not Everyone are kept.
         The file is backed up to printers.json.bak first.

    Default is a dry run: it reports what it would create and add. -Apply changes
    the tenant and the JSON.

    Does an Entra ID group work in a printer ACL?
    ---------------------------------------------
    Only where the user's Windows sign-in token carries it: a user who signs in
    with an Entra ID account on an Entra-joined device (Entra-joined AVD session
    hosts included) - and only after the next sign-in, as a token holds the groups
    it was issued with. A user who signs in with an Active Directory account
    (domain-joined or hybrid server, classic print server) gets AD groups in the
    token, never cloud groups, and the queue then refuses everybody but the
    administrators. Use AD groups (DOMAIN\Group) there.

.PARAMETER MatrixPath
    The matrix file (tab, semicolon or comma separated).

.PARAMETER Prefix
    Prefix of the group names (default: SG-U-PRINTM-).

.PARAMETER ConfigPath
    printers.json (from Install-Printer or Export-Printer) to write the groups into
    as permissions. Printers in the matrix that the JSON does not have are reported.

.PARAMETER RemoveExtraMembers
    Also remove members the matrix does not tick, so the group matches it exactly.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Defaults to the GDAP customer tenant ($global:cid /
    $env:M365_CUSTOMER_TENANTID), else the tenant you sign in to.

.PARAMETER ClientId
    App registration for app-only sign-in, with -CertificateThumbprint.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.PARAMETER Apply
    Create the groups, change their members and write the JSON. Without it, only
    report.

.EXAMPLE
    .\New-PrinterGroup.ps1 -MatrixPath C:\Temp\printer-matrix.tsv

    What would be created and added - nothing changes.

.EXAMPLE
    .\New-PrinterGroup.ps1 -MatrixPath C:\Temp\printer-matrix.tsv -ConfigPath C:\Temp\export\printers.json -Apply

    Create the groups, add the members and write the group SIDs into printers.json.

.NOTES
    Author          : Sjoerd Kanon
    Required scopes : Group.ReadWrite.All, User.Read.All
    Required module : Microsoft.Graph.Groups, Microsoft.Graph.Users
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)] [string] $MatrixPath,
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

function Write-Ok   { param([string] $Message) Write-Host "  [ OK ] $Message" -ForegroundColor Green }
function Write-Skip { param([string] $Message) Write-Host "  [SKIP] $Message" -ForegroundColor DarkGray }
function Write-Warn { param([string] $Message) Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
function Write-Bad  { param([string] $Message) Write-Host "  [FAIL] $Message" -ForegroundColor Red }
function Write-Plan { param([string] $Message) Write-Host "  [PLAN] $Message" -ForegroundColor Magenta }

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

function Resolve-MatrixUser {
    <#
        A matrix header to one Entra ID user: a UPN as given, a display name as is,
        then without a trailing "(...)". $null with a reason when it is none or
        more than one.
    #>
    param([Parameter(Mandatory)] [string] $Name)

    if ($Name -match '@') {
        try { return [PSCustomObject]@{ User = (Get-MgUser -UserId $Name -Property Id, DisplayName, UserPrincipalName -ErrorAction Stop); Reason = $null } }
        catch { return [PSCustomObject]@{ User = $null; Reason = "no user with UPN $Name" } }
    }
    foreach ($candidate in @($Name, ($Name -replace '\s*\([^)]*\)\s*$', '').Trim()) | Select-Object -Unique) {
        $hits = @(Get-MgUser -Filter "displayName eq '$($candidate -replace "'", "''")'" -Property Id, DisplayName, UserPrincipalName -All -ErrorAction Stop)
        if ($hits.Count -eq 1) { return [PSCustomObject]@{ User = $hits[0]; Reason = $null } }
        if ($hits.Count -gt 1) {
            return [PSCustomObject]@{ User = $null; Reason = "$($hits.Count) users are called '$candidate' ($(($hits | ForEach-Object UserPrincipalName) -join ', ')) - put the UPN in the header instead" }
        }
    }
    return [PSCustomObject]@{ User = $null; Reason = 'no user with that display name - put the UPN in the header instead' }
}

# -- Matrix ---------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $MatrixPath)) { throw "Matrix not found: $MatrixPath" }
$matrix = Read-PrinterMatrix -Path $MatrixPath
if ($ConfigPath -and -not (Test-Path -LiteralPath $ConfigPath)) { throw "printers.json not found: $ConfigPath" }

Write-Host ''
Write-Host "  Mode: $(if ($Apply) { 'APPLY - groups and members are changed' } else { 'PREVIEW - nothing is changed' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'Cyan' })
Write-Host "  Matrix: $($matrix.Printers.Count) printer(s), $($matrix.Users.Count) user(s)" -ForegroundColor Cyan
Write-Host ''

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')
try {
    $graph = Connect-M365Graph -Scopes 'Group.ReadWrite.All', 'User.Read.All' -TenantId $TenantId `
        -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly
} catch {
    Write-Bad "Could not connect to Microsoft Graph: $($_.Exception.Message)"
    exit 1
}

$exitCode = 0
try {
    # -- Users ------------------------------------------------------------------
    Write-Host '  Users' -ForegroundColor Cyan
    $resolved = @{}
    foreach ($name in $matrix.Users) {
        $r = Resolve-MatrixUser -Name $name
        if ($r.User) {
            $resolved[$name] = $r.User
            Write-Ok "$name -> $($r.User.UserPrincipalName)"
        } else {
            Write-Bad "${name}: $($r.Reason)"
            $exitCode = 1
        }
    }

    # -- Groups -----------------------------------------------------------------
    $sids = [ordered]@{}
    foreach ($printer in $matrix.Printers) {
        $groupName = Get-GroupName $printer.Name
        Write-Host ''
        Write-Host "  $groupName" -ForegroundColor Cyan

        $group = Get-MgGroup -Filter "displayName eq '$($groupName -replace "'", "''")'" -ErrorAction Stop | Select-Object -First 1
        if ($group) {
            Write-Skip "exists ($($group.Id))"
        } elseif (-not $Apply) {
            Write-Plan 'create the security group'
        } elseif ($PSCmdlet.ShouldProcess($groupName, 'Create Entra ID security group')) {
            $nickname = ($groupName -replace '[^A-Za-z0-9\-_.]', '')
            if ($nickname.Length -gt 64) { $nickname = $nickname.Substring(0, 64) }
            $group = New-MgGroup -DisplayName $groupName -MailEnabled:$false -MailNickname $nickname -SecurityEnabled:$true `
                                 -Description "Print access to the printer '$($printer.Name)'" -ErrorAction Stop
            Write-Ok "created ($($group.Id))"
        }

        $wanted  = @($printer.Users | Where-Object { $resolved.ContainsKey($_) } | ForEach-Object { $resolved[$_] })
        $members = @(if ($group) { Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop })
        $memberIds = @($members | ForEach-Object Id)

        foreach ($user in $wanted) {
            if ($user.Id -in $memberIds) { Write-Skip "$($user.DisplayName) is a member"; continue }
            if (-not $Apply -or -not $group) { Write-Plan "add $($user.DisplayName)"; continue }
            if ($PSCmdlet.ShouldProcess($groupName, "Add $($user.UserPrincipalName)")) {
                New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction Stop
                Write-Ok "added $($user.DisplayName)"
            }
        }
        foreach ($member in $members) {
            if ($member.Id -in @($wanted | ForEach-Object Id)) { continue }
            $label = $member.AdditionalProperties['displayName']
            if (-not $RemoveExtraMembers) { Write-Warn "$label is a member but not ticked in the matrix - left alone (-RemoveExtraMembers removes it)"; continue }
            if (-not $Apply) { Write-Plan "remove $label"; continue }
            if ($PSCmdlet.ShouldProcess($groupName, "Remove $label")) {
                Remove-MgGroupMemberDirectoryObjectByRef -GroupId $group.Id -DirectoryObjectId $member.Id -ErrorAction Stop
                Write-Ok "removed $label"
            }
        }
        if ($wanted.Count -eq 0) { Write-Warn 'nobody ticked (or resolved) for this printer - the group stays empty' }

        if ($group) {
            $sids[$printer.Name] = ConvertTo-EntraSid $group.Id
            Write-Skip "SID $($sids[$printer.Name])"
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
            if (-not $sids.Contains($printer.Name)) { Write-Plan "'$($printer.Name)': permissions once its group exists (-Apply)"; continue }

            $entry = $byName[$printer.Name]
            # Keep who manages the queue; replace who may print on it.
            $keep = @(if ($entry.PSObject.Properties.Name -contains 'permissions') {
                $entry.permissions | Where-Object { $_.access -ne 'print' -and $_.principal -ne 'Everyone' -and $_.principal -ne $sids[$printer.Name] }
            })
            $new = @([PSCustomObject]@{ principal = $sids[$printer.Name]; access = 'print' }) + $keep
            $entry | Add-Member -NotePropertyName permissions -NotePropertyValue $new -Force
            $changed++
            Write-Ok "'$($printer.Name)': print -> $(Get-GroupName $printer.Name)$(if ($keep.Count) { " (kept $($keep.Count) manage entr$(if ($keep.Count -eq 1) { 'y' } else { 'ies' }))" })"
        }

        if ($changed -gt 0 -and $Apply -and $PSCmdlet.ShouldProcess($ConfigPath, "Write permissions for $changed printer(s)")) {
            Copy-Item -LiteralPath $ConfigPath -Destination "$ConfigPath.bak" -Force
            [IO.File]::WriteAllText((Resolve-Path -LiteralPath $ConfigPath).ProviderPath, ($config | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
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
    Disconnect-M365Graph $graph
}
Write-Host ''
exit $exitCode
