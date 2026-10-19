#Requires -Version 7.0
<#
.SYNOPSIS
    Give each subfolder of one SharePoint folder its own set of people, so everyone sees only
    the subfolders they are listed for - created, broken from inheritance and reconciled in one run.

.DESCRIPTION
    The usual request: "under General we made a folder Teamleads; make it so that people only
    reach their own subfolder, read and write, and nothing else". This script does exactly that,
    for any folder, from one config:

      Parent folder (e.g. General/Teamleads)
        * Inheritance broken; the site Owners group keeps Full Control
        * Everyone named in any subfolder gets Read here (-ParentAccess Read, the default), so
          the folder shows up for them and they can click through. -ParentAccess None leaves
          only the owners - the subfolders then work by direct link only
        * Everyone else (Members, Visitors, earlier direct grants, sharing links) is removed

      Each configured subfolder (e.g. Team Clara)
        * Created if it does not exist yet
        * Inheritance broken; the site Owners group keeps Full Control
        * The people listed for it get -Role (Contribute by default: read, add, change and
          delete files - "lezen/schrijven")
        * Everyone else is removed, including the parent's Read

      Everything else directly in the parent (files, subfolders not in the config)
        * Locked to the owners: inheritance broken, everyone else removed. Without this the
          parent's Read would let everyone listed anywhere open them, and people would reach
          more than their own subfolder. -KeepOtherItems leaves them alone (reported only)

    A subfolder you have no rights on is not listed at all, so people do not even see the other
    teams' folder names.

    Kept on purpose, everywhere: the site's associated Owners group, the owners claim of the
    Microsoft 365 group behind a Teams site, and system "Limited Access" entries (SharePoint
    manages those itself so people can reach a deeper item; removing them breaks that access).
    Site collection administrators are not a role assignment and always keep access.

    Reporting is the default: the script reads everything, resolves every person, and prints
    what it would change. Nothing is written until -Apply. With -Apply it reads the permissions
    back afterwards and prints what is actually on each folder. Every run writes a CSV.

    Roles are picked by type, not by name, so it works the same on a Dutch or French site
    ("Bijdragen", "Collaboration") as on an English one.

    Sign-in: delegated PnP through Connect-M365PnP (device code / GDAP customer per
    load.config.ps1), with the client ID for the tenant from pnp.appid.json. If the tenant has
    no entry there yet, run Find-SiteContent.ps1 once against it (it registers the app) or pass
    -ClientId. -AppOnly signs in with the certificate app from graph.appid.json instead.

.PARAMETER ConfigPath
    JSON file describing the folder and who goes where. See Examples. Values given as parameters
    win over the ones in the file.

.PARAMETER SiteUrl
    The site the folder lives in, e.g. https://contoso.sharepoint.com/sites/Sales. For a Teams
    channel folder: the team's site.

.PARAMETER FolderPath
    The parent folder, site-relative, e.g. "Shared Documents/General/Teamleads". A full URL or a
    server-relative path (/sites/Sales/Shared Documents/...) is accepted too. Note that the
    library's URL is usually "Shared Documents" even when Teams and the browser show "Documents"
    or "Documenten".

.PARAMETER Access
    Subfolder name -> the people for it, as a hashtable, instead of (or on top of) the config:
    @{ 'Team A' = @('jan@contoso.com', 'piet@contoso.com'); 'Team B' = @('kim@contoso.com') }
    UPNs or e-mail addresses of the people.

.PARAMETER Role
    What the listed people get on their subfolder: Contribute (default, read/write), Edit
    (also manage the library's views and columns - rarely what you want on a folder) or Read.

.PARAMETER ParentAccess
    Read (default): everyone in any subfolder gets Read on the parent so they can browse to it.
    None: only the owners keep access to the parent.

.PARAMETER KeepOtherItems
    Leave files and unconfigured subfolders directly in the parent alone instead of locking them
    to the owners. They are still reported, with who can reach them.

.PARAMETER Apply
    Actually make the changes. Without it nothing is written.

.PARAMETER TenantId
    Tenant ID or domain for the sign-in. Defaults to what Connect-M365 resolves.

.PARAMETER ClientId
    Client ID of the PnP app registration to sign in with, instead of pnp.appid.json.

.PARAMETER AppOnly
    App-only sign-in with the certificate app from graph.appid.json (needs SharePoint
    Sites.FullControl.All as an application permission).

.PARAMETER OutputPath
    CSV report path. Default: C:\Temp\SharePointFolderAccess_<timestamp>.csv (~/Downloads off Windows).

.EXAMPLE
    # What would change? Reads only.
    .\Set-SharePointFolderAccess.ps1 -ConfigPath .\teamleads.json

.EXAMPLE
    # Do it.
    .\Set-SharePointFolderAccess.ps1 -ConfigPath .\teamleads.json -Apply

.EXAMPLE
    # Without a config file.
    .\Set-SharePointFolderAccess.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Sales `
        -FolderPath "Shared Documents/General/Teamleads" `
        -Access @{ 'Team A' = @('jan@contoso.com', 'kim@contoso.com'); 'Leads' = @('jan@contoso.com') } `
        -Apply

.EXAMPLE
    # The config file:
    # {
    #   "siteUrl":      "https://contoso.sharepoint.com/sites/Sales",
    #   "folder":       "Shared Documents/General/Teamleads",
    #   "role":         "Contribute",
    #   "parentAccess": "Read",
    #   "subfolders": {
    #     "Team A": [ "jan@contoso.com", "kim@contoso.com" ],
    #     "Leads":  [ "jan@contoso.com" ]
    #   }
    # }

.NOTES
    Idempotent: a second run reports "already correct" for everything the first one set.
    Only the parent, what sits directly in it and the configured subfolders are touched. Files or
    folders deeper inside a subfolder that have their own unique permissions (an earlier "Share"
    on one file) keep them - the subfolders' contents are not crawled. Something added directly
    to the parent later inherits its Read again; a re-run locks it.
#>
[CmdletBinding()]
param(
    [string] $ConfigPath,
    [string] $SiteUrl,
    [string] $FolderPath,
    [hashtable] $Access,

    [ValidateSet('Contribute', 'Edit', 'Read')]
    [string] $Role,

    [ValidateSet('Read', 'None')]
    [string] $ParentAccess,

    [switch] $KeepOtherItems,

    [switch] $Apply,
    [string] $TenantId,
    [string] $ClientId,
    [switch] $AppOnly,
    [string] $OutputPath
)

$ErrorActionPreference = 'Stop'

# -- Modules -------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name 'PnP.PowerShell')) {
    throw "Module 'PnP.PowerShell' is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser"
}
Import-Module PnP.PowerShell -ErrorAction Stop
. (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')

# -- Settings: config file, then parameters on top ------------------------------
$subfolders = [ordered]@{}
if ($ConfigPath) {
    if (-not (Test-Path -LiteralPath $ConfigPath)) { throw "Config file '$ConfigPath' not found." }
    $config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
    if (-not $SiteUrl      -and $config.siteUrl)      { $SiteUrl      = [string]$config.siteUrl }
    if (-not $FolderPath   -and $config.folder)       { $FolderPath   = [string]$config.folder }
    if (-not $Role         -and $config.role)         { $Role         = [string]$config.role }
    if (-not $ParentAccess -and $config.parentAccess) { $ParentAccess = [string]$config.parentAccess }
    if ($config.subfolders) {
        foreach ($p in $config.subfolders.PSObject.Properties) { $subfolders[$p.Name] = @($p.Value) }
    }
}
if ($Access) {
    foreach ($key in $Access.Keys) { $subfolders[[string]$key] = @($Access[$key]) }
}
if (-not $Role)         { $Role = 'Contribute' }
if (-not $ParentAccess) { $ParentAccess = 'Read' }

if ($Role -notin 'Contribute', 'Edit', 'Read') { throw "Role '$Role' is not one of Contribute, Edit, Read." }
if ($ParentAccess -notin 'Read', 'None')       { throw "parentAccess '$ParentAccess' is not one of Read, None." }
if (-not $SiteUrl)    { throw 'No site: pass -SiteUrl or set siteUrl in the config.' }
if (-not $FolderPath) { throw 'No folder: pass -FolderPath or set folder in the config.' }
if ($subfolders.Count -eq 0) { throw 'No subfolders: pass -Access or set subfolders in the config.' }

# Clean the people lists: trimmed, no blanks, no duplicates, case-insensitive.
foreach ($name in @($subfolders.Keys)) {
    if ($name -match '[\\/:*?"<>|#%]' -or $name.Trim() -ne $name -or $name.EndsWith('.')) {
        throw "Subfolder name '$name' contains characters SharePoint does not allow in a folder name."
    }
    $people = @($subfolders[$name] | ForEach-Object { "$_".Trim() } | Where-Object { $_ } |
                Sort-Object -Unique)
    if ($people.Count -eq 0) { throw "Subfolder '$name' has nobody listed. Leave it out, or list at least one person." }
    $subfolders[$name] = $people
}

$SiteUrl = $SiteUrl.Trim().TrimEnd('/')
if ($SiteUrl -notmatch '^https://[^/]+\.sharepoint\.(com|de|us|cn)(/|$)') {
    throw "'$SiteUrl' does not look like a SharePoint site URL."
}

# -- Output --------------------------------------------------------------------
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not $OutputPath) {
    if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
    $OutputPath = Join-Path $outputDir "SharePointFolderAccess_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
}

$results = New-Object System.Collections.Generic.List[object]
function Add-Result {
    param([string] $Folder, [string] $Action, [string] $Principal = '', [string] $Detail = '', [string] $Result)
    $results.Add([pscustomobject]@{
        Folder    = $Folder
        Action    = $Action
        Principal = $Principal
        Detail    = $Detail
        Result    = $Result
    })
}

function Write-Ok     { param([string] $Message) Write-Host "    [ OK ] $Message" -ForegroundColor Green }
function Write-Change { param([string] $Message) Write-Host "    [ >> ] $Message" -ForegroundColor Magenta }
function Write-Warn   { param([string] $Message) Write-Host "    [WARN] $Message" -ForegroundColor Yellow }
function Write-Bad    { param([string] $Message) Write-Host "    [FAIL] $Message" -ForegroundColor Red }

# -- Header --------------------------------------------------------------------
Write-Host ''
Write-Host "  Site    : $SiteUrl" -ForegroundColor Cyan
Write-Host "  Folder  : $FolderPath" -ForegroundColor Cyan
Write-Host "  Role    : $Role on each subfolder, parent: $ParentAccess" -ForegroundColor Cyan
foreach ($name in $subfolders.Keys) {
    Write-Host ("    {0,-24} {1}" -f $name, ($subfolders[$name] -join ', ')) -ForegroundColor DarkGray
}
if ($Apply) {
    Write-Host '  Mode    : APPLY - permissions will be changed' -ForegroundColor Yellow
} else {
    Write-Host '  Mode    : DRY RUN - nothing is changed without -Apply' -ForegroundColor DarkGray
}
Write-Host ''

# -- Connect -------------------------------------------------------------------
$connectParams = @{ Url = $SiteUrl }
if ($TenantId) { $connectParams['TenantId'] = $TenantId }
if ($ClientId) { $connectParams['ClientId'] = $ClientId }
if ($AppOnly)  { $connectParams['AppOnly']  = $true }
$connection = Connect-M365PnP @connectParams

$context = Get-PnPContext -Connection $connection
$web     = Get-PnPWeb -Connection $connection -Includes ServerRelativeUrl, AssociatedOwnerGroup
$ownerGroupId = if ($web.AssociatedOwnerGroup -and -not $web.AssociatedOwnerGroup.ServerObjectIsNull) { $web.AssociatedOwnerGroup.Id } else { $null }
if (-not $ownerGroupId) {
    Write-Warn 'This site has no associated Owners group. Only the people in the config and site collection admins will have access.'
}

# -- Folder path: site-relative, whatever form it was given in ------------------
$relative = [System.Uri]::UnescapeDataString($FolderPath.Trim())
if ($relative -match '^https://[^/]+(/.*)$') { $relative = $Matches[1] }
$webPath = $web.ServerRelativeUrl.TrimEnd('/')
if ($webPath -and $relative.StartsWith("$webPath/", [System.StringComparison]::OrdinalIgnoreCase)) {
    $relative = $relative.Substring($webPath.Length)
}
$relative = $relative.Trim('/')

try {
    $parentFolder = Get-PnPFolder -Url $relative -Connection $connection -Includes ListItemAllFields
} catch {
    throw "Folder '$relative' not found in $SiteUrl. The library URL is usually 'Shared Documents', even when it is shown as 'Documents' or 'Documenten'. ($($_.Exception.Message))"
}
if (-not $parentFolder.ListItemAllFields -or $parentFolder.ListItemAllFields.ServerObjectIsNull) {
    throw "'$relative' is the root of a library, not a folder inside one. Point -FolderPath at a folder."
}

# -- Roles, by type so the site language does not matter -----------------------
$roleType = @{
    Read       = [Microsoft.SharePoint.Client.RoleType]::Reader
    Contribute = [Microsoft.SharePoint.Client.RoleType]::Contributor
    Edit       = [Microsoft.SharePoint.Client.RoleType]::Editor
}
$roleDefs = @{}
foreach ($kind in @('Read', 'Contribute', 'Edit', 'FullControl')) {
    $type = if ($kind -eq 'FullControl') { [Microsoft.SharePoint.Client.RoleType]::Administrator } else { $roleType[$kind] }
    $def  = $web.RoleDefinitions.GetByType($type)
    $context.Load($def)
    $roleDefs[$kind] = $def
}
$context.ExecuteQuery()

# -- Resolve everyone before touching anything ---------------------------------
# One unknown address must stop the run here, not halfway through, with the parent already
# closed off and half the subfolders done.
Write-Host '  Resolving people...' -ForegroundColor Cyan
$principals = @{}
$unresolved = @()
foreach ($login in @($subfolders.Values | ForEach-Object { $_ } | Sort-Object -Unique)) {
    try {
        $user = $web.EnsureUser($login)
        $context.Load($user)
        $context.ExecuteQuery()
        $principals[$login] = $user
        Write-Host ("    {0,-40} {1}" -f $login, $user.Title) -ForegroundColor DarkGray
    } catch {
        $unresolved += $login
        Write-Bad "$login - $($_.Exception.Message.Trim())"
    }
}
if ($unresolved.Count -gt 0) {
    foreach ($login in $unresolved) { Add-Result -Folder $relative -Action 'Resolve' -Principal $login -Result 'Not found' }
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    throw "$($unresolved.Count) address(es) could not be resolved - nothing was changed. Fix them in the config and run again. Report: $OutputPath"
}

# -- Helpers -------------------------------------------------------------------
function Get-Assignments {
    <# The role assignments on a list item, with member and bindings loaded. Inherited ones too:
       on an item that inherits, SharePoint returns those of the ancestor it inherits from. #>
    param($Item)
    $context.Load($Item)
    $context.Load($Item.RoleAssignments)
    $context.ExecuteQuery()
    foreach ($a in $Item.RoleAssignments) {
        $context.Load($a.Member)
        $context.Load($a.RoleDefinitionBindings)
    }
    $context.ExecuteQuery()
    return @($Item.RoleAssignments)
}

function Test-Protected {
    <# Assignments that are never removed: the site Owners group, the owners of the M365 group
       behind a Teams site, and system Limited Access entries. #>
    param($Assignment)
    if ($ownerGroupId -and $Assignment.Member.Id -eq $ownerGroupId) { return 'site owners' }
    if ($Assignment.Member.LoginName -match 'federateddirectoryclaimprovider\|.+_o$') { return 'group owners' }
    $bindings = @($Assignment.RoleDefinitionBindings)
    if ($bindings.Count -gt 0 -and -not ($bindings | Where-Object { -not $_.Hidden -and "$($_.RoleTypeKind)" -ne 'Guest' })) {
        return 'limited access'
    }
    return $null
}

function Set-FolderAccess {
    <#
        Bring one folder in line: unique permissions, owners Full Control, exactly the desired
        people with exactly the desired role, nobody else. Returns the number of changes
        (planned in a dry run, made with -Apply).
    #>
    param(
        [string] $Label,
        $Item,
        [hashtable] $Desired,   # login -> role kind (Read / Contribute / Edit)
        [switch] $IsNew
    )

    Write-Host ''
    Write-Host "  $Label" -ForegroundColor Cyan
    $changes = 0

    if ($IsNew -and -not $Apply) {
        Write-Change 'create folder'
        Add-Result -Folder $Label -Action 'Create folder' -Result 'Planned'
        Write-Change 'break inheritance'
        foreach ($login in $Desired.Keys) {
            Write-Change "grant $($Desired[$login]) to $login"
            Add-Result -Folder $Label -Action 'Grant' -Principal $login -Detail $Desired[$login] -Result 'Planned'
        }
        Write-Change 'remove everything inherited from the parent except the owners'
        return 2 + $Desired.Count
    }

    $assignments = Get-Assignments -Item $Item

    if (-not $Item.HasUniqueRoleAssignments) {
        $changes++
        Write-Change 'break inheritance (current permissions copied, then trimmed below)'
        if ($Apply) {
            # Copy, not clear: clearing would hand the signed-in admin Full Control as a side
            # effect, and the reconcile below removes what should not stay anyway.
            $Item.BreakRoleInheritance($true, $false)
            $context.ExecuteQuery()
            $assignments = Get-Assignments -Item $Item
            Add-Result -Folder $Label -Action 'Break inheritance' -Result 'Done'
        } else {
            Add-Result -Folder $Label -Action 'Break inheritance' -Result 'Planned'
        }
    }

    $desiredById = @{}
    foreach ($login in $Desired.Keys) { $desiredById[$principals[$login].Id] = $login }

    # Owners first, so the folder can never end up without them.
    if ($ownerGroupId -and -not ($assignments | Where-Object { $_.Member.Id -eq $ownerGroupId })) {
        $changes++
        Write-Change "grant Full Control to the site Owners group"
        if ($Apply) {
            $bindings = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($context)
            $bindings.Add($roleDefs['FullControl'])
            $Item.RoleAssignments.Add($web.AssociatedOwnerGroup, $bindings) | Out-Null
            $context.ExecuteQuery()
        }
        Add-Result -Folder $Label -Action 'Grant' -Principal 'Site Owners' -Detail 'FullControl' -Result $(if ($Apply) { 'Done' } else { 'Planned' })
    }

    $seen = @{}
    foreach ($a in $assignments) {
        $member = $a.Member
        $roles  = @($a.RoleDefinitionBindings | Where-Object { -not $_.Hidden } | ForEach-Object { $_.Name })

        if ($desiredById.ContainsKey($member.Id)) {
            $login = $desiredById[$member.Id]
            $seen[$login] = $true
            $want  = $roleDefs[$Desired[$login]]
            $have  = @($a.RoleDefinitionBindings | Where-Object { -not $_.Hidden })
            if ($have.Count -eq 1 -and $have[0].Id -eq $want.Id) {
                Write-Ok "$login has $($want.Name)"
                Add-Result -Folder $Label -Action 'Check' -Principal $login -Detail $want.Name -Result 'Already correct'
                continue
            }
            $changes++
            Write-Change "set $login to $($want.Name) (was: $($roles -join ', '))"
            if ($Apply) {
                $a.RoleDefinitionBindings.RemoveAll()
                $a.RoleDefinitionBindings.Add($want)
                $a.Update()
                $context.ExecuteQuery()
            }
            Add-Result -Folder $Label -Action 'Change role' -Principal $login -Detail "$($roles -join ', ') -> $($want.Name)" -Result $(if ($Apply) { 'Done' } else { 'Planned' })
            continue
        }

        $keep = Test-Protected -Assignment $a
        if ($keep) {
            Write-Host "    [KEEP] $($member.Title) ($keep)" -ForegroundColor DarkGray
            Add-Result -Folder $Label -Action 'Keep' -Principal $member.LoginName -Detail ($roles -join ', ') -Result $keep
            continue
        }

        $changes++
        $kind = if ($member.LoginName -like 'SharingLinks.*') { 'sharing link' } else { "$($roles -join ', ')" }
        Write-Change "remove $($member.Title) ($kind)"
        Add-Result -Folder $Label -Action 'Remove' -Principal $member.LoginName -Detail $kind -Result $(if ($Apply) { 'Pending' } else { 'Planned' })
        if ($Apply) {
            $a.DeleteObject()
            $context.ExecuteQuery()
            $results[$results.Count - 1].Result = 'Done'
        }
    }

    foreach ($login in $Desired.Keys) {
        if ($seen[$login]) { continue }
        $changes++
        $want = $roleDefs[$Desired[$login]]
        Write-Change "grant $($want.Name) to $login"
        if ($Apply) {
            $bindings = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($context)
            $bindings.Add($want)
            $Item.RoleAssignments.Add($principals[$login], $bindings) | Out-Null
            $context.ExecuteQuery()
        }
        Add-Result -Folder $Label -Action 'Grant' -Principal $login -Detail $want.Name -Result $(if ($Apply) { 'Done' } else { 'Planned' })
    }

    if ($changes -eq 0) { Write-Ok 'already correct' }
    return $changes
}

function Show-FolderAccess {
    <# Read back what is actually on a folder now - the proof after -Apply. #>
    param([string] $Label, $Item)
    $assignments = Get-Assignments -Item $Item
    Write-Host "    $Label" -ForegroundColor Cyan
    foreach ($a in $assignments) {
        $roles = @($a.RoleDefinitionBindings | Where-Object { -not $_.Hidden -and "$($_.RoleTypeKind)" -ne 'Guest' } | ForEach-Object { $_.Name })
        if ($roles.Count -eq 0) { continue }
        Write-Host ("      {0,-40} {1}" -f $a.Member.Title, ($roles -join ', ')) -ForegroundColor DarkGray
    }
}

# Wrapped so the CSV is written even when a change fails halfway: the report then shows what
# was done before it stopped, which is what you need to judge the state the folder is in.
$total = 0
try {
    # -- Parent folder -------------------------------------------------------------
    $parentDesired = @{}
    if ($ParentAccess -eq 'Read') {
        foreach ($login in $principals.Keys) { $parentDesired[$login] = 'Read' }
    }
    $total += Set-FolderAccess -Label $relative -Item $parentFolder.ListItemAllFields -Desired $parentDesired

    # -- Subfolders ----------------------------------------------------------------
    $existing = @{}
    $others   = @()
    foreach ($f in @(Get-PnPFolderItem -FolderSiteRelativeUrl $relative -Connection $connection)) {
        if ($f -is [Microsoft.SharePoint.Client.Folder]) { $existing[$f.Name] = $f }
        if (-not ($f -is [Microsoft.SharePoint.Client.Folder] -and $subfolders.Contains($f.Name))) { $others += $f }
    }

    $managed = @{}
    foreach ($name in $subfolders.Keys) {
        $path    = "$relative/$name"
        $desired = @{}
        foreach ($login in $subfolders[$name]) { $desired[$login] = $Role }

        $isNew = -not ($existing.Keys | Where-Object { $_ -eq $name })
        if ($isNew -and $Apply) {
            Add-PnPFolder -Name $name -Folder $relative -Connection $connection | Out-Null
            Add-Result -Folder $path -Action 'Create folder' -Result 'Done'
            Write-Host ''
            Write-Host "  Created $path" -ForegroundColor Magenta
            $total++
        }

        $item = if ($isNew -and -not $Apply) { $null }
                else { (Get-PnPFolder -Url $path -Connection $connection -Includes ListItemAllFields).ListItemAllFields }
        $managed[$name] = $item
        $total += Set-FolderAccess -Label $path -Item $item -Desired $desired -IsNew:($isNew -and -not $Apply)
    }

    # Anything else directly in the parent - files, subfolders the config does not name - inherits
    # the parent's Read, so everyone listed anywhere could open it. That is more than "only their
    # own subfolder", so it is locked to the owners: the same reconcile with nobody desired.
    foreach ($f in $others) {
        $kind = if ($f -is [Microsoft.SharePoint.Client.Folder]) { 'folder' } else { 'file' }
        $item = $f.ListItemAllFields
        $context.Load($item); $context.ExecuteQuery()
        $open = $item.HasUniqueRoleAssignments -or $ParentAccess -eq 'Read'

        if ($KeepOtherItems -or -not $open) {
            $how = if ($item.HasUniqueRoleAssignments) { 'has its own permissions' }
                   elseif ($ParentAccess -eq 'Read') { 'inherits from the parent: everyone in the config can READ it' }
                   else { 'inherits from the parent: owners only' }
            if ($open) {
                Write-Host ''
                Write-Warn "$kind '$($f.Name)' is not in the config and was left alone (-KeepOtherItems) - it $how."
            }
            Add-Result -Folder "$relative/$($f.Name)" -Action 'Not managed' -Detail $how -Result 'Untouched'
            continue
        }
        $total += Set-FolderAccess -Label "$relative/$($f.Name) ($kind, not in the config: owners only)" -Item $item -Desired @{}
    }

    # -- Read back -----------------------------------------------------------------
    if ($Apply) {
        Write-Host ''
        Write-Host '  Permissions now in place' -ForegroundColor Cyan
        Show-FolderAccess -Label $relative -Item $parentFolder.ListItemAllFields
        foreach ($name in $managed.Keys) { Show-FolderAccess -Label "$relative/$name" -Item $managed[$name] }
    }
} finally {
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
}

# -- Summary -------------------------------------------------------------------
Write-Host ''
if ($total -eq 0) {
    Write-Host '  Everything already matches the config.' -ForegroundColor Green
} elseif ($Apply) {
    Write-Host "  $total change(s) made." -ForegroundColor Green
} else {
    Write-Host "  $total change(s) planned. Run again with -Apply to make them." -ForegroundColor Yellow
}
Write-Host "  Report: $OutputPath" -ForegroundColor DarkGray
Write-Host ''
