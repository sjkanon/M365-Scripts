#Requires -Version 7.0
<#
.SYNOPSIS
    Give each subfolder of one SharePoint folder its own SharePoint group, a group of owners
    who manage it all, and optionally a personal folder per person - so everyone reaches only
    their own folders. Created, broken from inheritance and reconciled in one run.

.DESCRIPTION
    The usual request: "under General we made a folder Teamleads; make it so that people only
    reach their own subfolder, read and write, and nothing else - and let the team leads manage
    it". This script does that for any folder, from one config.

    Groups (default; -DirectAccess grants the people directly instead)
      * "<prefix> - Owners"       the people in "owners". Full Control on the parent and on
                                  everything the script manages below it
      * "<prefix> - <subfolder>"  the people listed for that subfolder
      Entra ID security groups by default (-GroupType Entra): visible and manageable in Entra,
      no mailbox, no team. The people in "owners" become owner of every group, so they manage
      membership themselves in My Groups (myaccount.microsoft.com/groups) - no IT needed. A
      membership change can take a while to reach SharePoint. <prefix> is "groupPrefix" from
      the config, else "<site> - <folder>" (Entra names are tenant-wide, so the site is in it).
      -GroupType SharePoint makes site-local SharePoint groups instead, owned by the Owners
      group, prefix default the folder name: effective at once, nothing in the directory.

    Parent folder (e.g. General/Teamleads)
      * Inheritance broken. Site Owners and the Owners group: Full Control
      * Every subfolder group: Read (-ParentAccess Read, the default), so the folder shows up
        and people can click through. -ParentAccess None: owners only
      * Everything else directly in the parent (files, subfolders not in the config): owners
        only, otherwise the parent's Read would let everyone open them. -KeepOtherItems leaves
        those alone (reported only)
      * Everyone else (Members, Visitors, earlier direct grants, sharing links) is removed

    Each configured subfolder (e.g. Team Clara)
      * Created if missing; inheritance broken; owners Full Control
      * Its group gets -Role (Contribute by default: read, add, change, delete - "lezen/schrijven")
      * Everyone else is removed, including the parent's Read

    Personal folders (per subfolder "personalFolders": true or "all", or -PersonalFolders for all)
      * Inside the subfolder, one folder per listed person who is not an owner ("all": the
        owners too), named after their display name. Only that person (Contribute) and the
        owners (Full Control) - so owners can open each other's personal folders
      * Everything else in the subfolder stays the team's, inheriting the subfolder

    A folder you have no rights on is not listed at all, so people do not see the other teams'
    folder names. Kept on purpose, everywhere: the site's associated Owners group, the owners
    claim of the Microsoft 365 group behind a Teams site, and system "Limited Access" entries.
    Site collection administrators are not a role assignment and always keep access.

    Group membership: missing members are added; people in a group who are not in the config
    are reported and left in, because the owners manage membership themselves and a re-run must
    not undo that. -SyncMembers makes the groups match the config exactly.

    Reporting is the default: the script reads everything, resolves every person, and prints
    what it would change. Nothing is written until -Apply. With -Apply it reads the permissions
    back afterwards and prints the group pages where the owners manage membership. Every run
    writes a CSV. Roles are picked by type, not by name, so a Dutch or French site works the same.

    Sign-in, the same as the other SharePoint scripts: PnP.PowerShell no longer ships a shared
    multi-tenant app, so the first run against a tenant creates one:

      1. Signs in to Microsoft Graph as an admin (Application.ReadWrite.All), through
         Connect-M365Graph: device code / GDAP customer per load.config.ps1
      2. Creates (or reuses) a public-client app named after -AppName
      3. Grants and admin-consents the delegated SharePoint scope AllSites.FullControl
      4. Caches the client ID per tenant in pnp.appid.json (gitignored)

    Later runs - and the other SharePoint scripts - read the cached client ID and go straight to
    the delegated PnP sign-in (Connect-M365PnP: a device code when $global:useDeviceCodeAuth is
    set in load.config.ps1, otherwise the browser). A client ID another script already cached
    for the tenant is reused. Pass -ClientId to skip app registration entirely; -AppOnly signs in
    with the certificate app from graph.appid.json instead.

.PARAMETER ConfigPath
    JSON file describing the folder, the owners and who goes where. See Examples. Values given as
    parameters win over the ones in the file.

.PARAMETER SiteUrl
    The site the folder lives in, e.g. https://contoso.sharepoint.com/sites/Sales. For a Teams
    channel folder: the team's site.

.PARAMETER FolderPath
    The parent folder, site-relative, e.g. "Shared Documents/General/Teamleads". A full URL or a
    server-relative path (/sites/Sales/Shared Documents/...) is accepted too. The library's URL is
    usually "Shared Documents" even when Teams and the browser show "Documents" or "Documenten".

.PARAMETER Access
    Subfolder name -> the people for it, as a hashtable, instead of (or on top of) the config:
    @{ 'Team A' = @('jan@contoso.com', 'piet@contoso.com'); 'Team B' = @('kim@contoso.com') }

.PARAMETER Owners
    The people who get Full Control everywhere and manage the groups. Overrides "owners" in the
    config.

.PARAMETER GroupPrefix
    Start of the group names. Default: "groupPrefix" from the config, else the parent folder name.

.PARAMETER PersonalFolders
    Give every listed non-owner a personal folder in every subfolder. Without it, per subfolder
    "personalFolders" in the config decides.

.PARAMETER DirectAccess
    Grant the people directly on the folders instead of through SharePoint groups. Membership
    then changes only by editing the config and running again.

.PARAMETER GroupType
    Entra (default): Entra ID security groups, created through Microsoft Graph (a delegated
    admin sign-in with Group.ReadWrite.All). SharePoint: SharePoint groups in this site only.
    Overrides "groupType" in the config.

.PARAMETER SyncMembers
    Remove group members who are not in the config. Without it they are only reported.

.PARAMETER Role
    What a subfolder's people get on it: Contribute (default, read/write), Edit (also manage the
    library's views and columns - rarely what you want on a folder) or Read. Personal folders
    always get Contribute.

.PARAMETER ParentAccess
    Read (default): every subfolder's people get Read on the parent so they can browse to it.
    None: only the owners keep access to the parent.

.PARAMETER KeepOtherItems
    Leave files and unconfigured subfolders directly in the parent alone instead of locking them
    to the owners. They are still reported, with who can reach them.

.PARAMETER Apply
    Actually make the changes. Without it nothing is written.

.PARAMETER TenantId
    Tenant ID or domain for the sign-in and the one-time Graph admin sign-in. Defaults to the
    tenant derived from -SiteUrl.

.PARAMETER ClientId
    Client ID of an existing Entra app to sign in with. Skips app registration.

.PARAMETER AppName
    Display name of the app registration to create or reuse.
    Default: "M365-Scripts SharePoint Folder Access".

.PARAMETER GrantSiteAdmin
    Make the signed-in admin site collection administrator for the duration of the run
    (requires SharePoint Administrator) and remove the rights again afterwards, also when the
    run fails. A Global Administrator has no access to a team's files by role alone, so this is
    what you normally need on a customer's Teams site you are not a member of. Without it such
    a run stops at "Attempted to perform an unauthorized operation".

.PARAMETER KeepSiteAdmin
    Keep the site collection admin rights granted by -GrantSiteAdmin.

.PARAMETER AdminUpn
    UPN to grant site collection admin rights to. Defaults to the signed-in account.

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
    # Without a config file, people granted directly instead of through groups.
    .\Set-SharePointFolderAccess.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Sales `
        -FolderPath "Shared Documents/General/Teamleads" -Owners jan@contoso.com `
        -Access @{ 'Team A' = @('kim@contoso.com', 'els@contoso.com') } -DirectAccess -Apply

.EXAMPLE
    # The config file. A subfolder is either a list of people, or an object with "members"
    # and "personalFolders".
    # {
    #   "siteUrl":      "https://contoso.sharepoint.com/sites/Sales",
    #   "folder":       "Shared Documents/General/Teamleads",
    #   "groupPrefix":  "Teamleads",
    #   "owners":       [ "jan@contoso.com", "kim@contoso.com" ],
    #   "role":         "Contribute",
    #   "parentAccess": "Read",
    #   "subfolders": {
    #     "Team A": { "members": [ "jan@contoso.com", "els@contoso.com", "piet@contoso.com" ], "personalFolders": true },
    #     "Leads":  [ "jan@contoso.com", "kim@contoso.com" ]
    #   }
    # }

.NOTES
    Idempotent: a second run reports "already correct" for everything the first one set.
    Personal folders are matched by name: rename one and the next run creates a new one under the
    display name (and reports the renamed one as not managed).
    Only the parent, what sits directly in it, the configured subfolders and their personal
    folders are touched. Anything deeper that was shared separately keeps its own permissions.
    Something added directly to the parent later inherits its Read again; a re-run locks it.
#>
[CmdletBinding()]
param(
    [string] $ConfigPath,
    [string] $SiteUrl,
    [string] $FolderPath,
    [hashtable] $Access,
    [string[]] $Owners,
    [string] $GroupPrefix,
    [switch] $PersonalFolders,
    [switch] $DirectAccess,
    [switch] $SyncMembers,

    [ValidateSet('Entra', 'SharePoint')]
    [string] $GroupType,

    [ValidateSet('Contribute', 'Edit', 'Read')]
    [string] $Role,

    [ValidateSet('Read', 'None')]
    [string] $ParentAccess,

    [switch] $KeepOtherItems,
    [switch] $Apply,
    [string] $TenantId,
    [string] $ClientId,
    [string] $AppName = 'M365-Scripts SharePoint Folder Access',
    [switch] $AppOnly,
    [switch] $GrantSiteAdmin,
    [switch] $KeepSiteAdmin,
    [string] $AdminUpn,
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
function ConvertTo-LoginList {
    # Trimmed, no blanks, no duplicates (case-insensitive).
    param($Value)
    return @(@($Value) | ForEach-Object { "$_".Trim() } | Where-Object { $_ } | Sort-Object -Unique)
}

$plan = [ordered]@{}   # subfolder name -> @{ Members; Personal }
function Add-PlanEntry {
    param([string] $Name, $Value)
    $members  = $Value
    $personal = $false
    if ($Value -is [System.Management.Automation.PSCustomObject] -or $Value -is [hashtable]) {
        $members  = $Value.members
        # true: every listed non-owner; "all": the owners as well (a team leads' folder, where
        # the leads are the owners and still want a folder of their own).
        $personal = switch ("$($Value.personalFolders)".ToLowerInvariant()) {
            'all'   { 'all' }
            'true'  { $true }
            ''      { $false }
            'false' { $false }
            default { throw "personalFolders for '$Name' is '$($Value.personalFolders)' - use true, false or ""all""." }
        }
    }
    $plan[$Name] = @{ Members = (ConvertTo-LoginList $members); Personal = $personal }
}

$ownerLogins = @()
if ($ConfigPath) {
    if (-not (Test-Path -LiteralPath $ConfigPath)) { throw "Config file '$ConfigPath' not found." }
    $config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
    if (-not $SiteUrl      -and $config.siteUrl)      { $SiteUrl      = [string]$config.siteUrl }
    if (-not $FolderPath   -and $config.folder)       { $FolderPath   = [string]$config.folder }
    if (-not $Role         -and $config.role)         { $Role         = [string]$config.role }
    if (-not $ParentAccess -and $config.parentAccess) { $ParentAccess = [string]$config.parentAccess }
    if (-not $GroupPrefix  -and $config.groupPrefix)  { $GroupPrefix  = [string]$config.groupPrefix }
    if (-not $GroupType    -and $config.groupType)    { $GroupType    = [string]$config.groupType }
    if ($config.owners) { $ownerLogins = ConvertTo-LoginList $config.owners }
    if ($config.subfolders) {
        foreach ($p in $config.subfolders.PSObject.Properties) { Add-PlanEntry -Name $p.Name -Value $p.Value }
    }
}
if ($Access) {
    foreach ($key in $Access.Keys) { Add-PlanEntry -Name ([string]$key) -Value $Access[$key] }
}
if ($Owners)          { $ownerLogins = ConvertTo-LoginList $Owners }
if ($PersonalFolders) { foreach ($name in $plan.Keys) { $plan[$name].Personal = $true } }
if (-not $Role)         { $Role = 'Contribute' }
if (-not $GroupType)    { $GroupType = 'Entra' }
if ($GroupType -notin 'Entra', 'SharePoint') { throw "groupType '$GroupType' is not one of Entra, SharePoint." }
$useEntra = $GroupType -eq 'Entra' -and -not $DirectAccess
if (-not $ParentAccess) { $ParentAccess = 'Read' }

if ($Role -notin 'Contribute', 'Edit', 'Read') { throw "Role '$Role' is not one of Contribute, Edit, Read." }
if ($ParentAccess -notin 'Read', 'None')       { throw "parentAccess '$ParentAccess' is not one of Read, None." }
if (-not $SiteUrl)    { throw 'No site: pass -SiteUrl or set siteUrl in the config.' }
if (-not $FolderPath) { throw 'No folder: pass -FolderPath or set folder in the config.' }
if ($plan.Count -eq 0) { throw 'No subfolders: pass -Access or set subfolders in the config.' }

$invalidName = '[\\/:*?"<>|#%]'
foreach ($name in $plan.Keys) {
    if ($name -match $invalidName -or $name.Trim() -ne $name -or $name.EndsWith('.')) {
        throw "Subfolder name '$name' contains characters SharePoint does not allow in a folder name."
    }
    if ($plan[$name].Members.Count -eq 0) { throw "Subfolder '$name' has nobody listed. Leave it out, or list at least one person." }
}
$ownerKeys = @{}
foreach ($login in $ownerLogins) { $ownerKeys[$login.ToLowerInvariant()] = $true }
if (($plan.Values | Where-Object Personal) -and $ownerLogins.Count -eq 0) {
    Write-Warning 'Personal folders without owners: only the person and the site Owners group will reach each one.'
}

$SiteUrl = $SiteUrl.Trim().TrimEnd('/')
if ($SiteUrl -notmatch '^https://([a-zA-Z0-9-]+?)(-my)?\.sharepoint\.(com|de|us|cn)(/|$)') {
    throw "'$SiteUrl' does not look like a SharePoint site URL."
}
$tenantName = $Matches[1]
$adminUrl   = "https://$tenantName-admin.sharepoint.$($Matches[3])"
if (-not $TenantId) { $TenantId = "$tenantName.onmicrosoft.com" }

$repoRoot   = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$appIdStore = Join-Path $repoRoot 'pnp.appid.json'

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
$status = if ($Apply) { 'Done' } else { 'Planned' }

function Write-Ok     { param([string] $Message) Write-Host "    [ OK ] $Message" -ForegroundColor Green }
function Write-Change { param([string] $Message) Write-Host "    [ >> ] $Message" -ForegroundColor Magenta }
function Write-Warn   { param([string] $Message) Write-Host "    [WARN] $Message" -ForegroundColor Yellow }
function Write-Bad    { param([string] $Message) Write-Host "    [FAIL] $Message" -ForegroundColor Red }

# -- Header --------------------------------------------------------------------
Write-Host ''
Write-Host "  Site    : $SiteUrl" -ForegroundColor Cyan
Write-Host "  Folder  : $FolderPath" -ForegroundColor Cyan
Write-Host "  Access  : $(if ($DirectAccess) { 'people granted directly' } elseif ($useEntra) { 'through Entra security groups' } else { 'through SharePoint groups' }); $Role on each subfolder, parent: $ParentAccess" -ForegroundColor Cyan
Write-Host ("  Owners  : {0}" -f $(if ($ownerLogins) { $ownerLogins -join ', ' } else { '(none - site Owners only)' })) -ForegroundColor Cyan
foreach ($name in $plan.Keys) {
    $tag = if ($plan[$name].Personal) { ' + personal folders' } else { '' }
    Write-Host ("    {0,-24} {1}{2}" -f $name, ($plan[$name].Members -join ', '), $tag) -ForegroundColor DarkGray
}
if ($Apply) {
    Write-Host '  Mode    : APPLY - groups and permissions will be changed' -ForegroundColor Yellow
} else {
    Write-Host '  Mode    : DRY RUN - nothing is changed without -Apply' -ForegroundColor DarkGray
}
Write-Host ''

# -- App registration ----------------------------------------------------------
function Get-CachedClientId {
    param([string] $Tenant)

    if (-not (Test-Path $appIdStore)) { return $null }
    try {
        $store = Get-Content $appIdStore -Raw | ConvertFrom-Json
        return $store.$Tenant
    } catch {
        Write-Warning "Could not read ${appIdStore}: $($_.Exception.Message)"
        return $null
    }
}

function Set-CachedClientId {
    param([string] $Tenant, [string] $Id)

    $store = @{}
    if (Test-Path $appIdStore) {
        try {
            (Get-Content $appIdStore -Raw | ConvertFrom-Json).PSObject.Properties |
                ForEach-Object { $store[$_.Name] = $_.Value }
        } catch { }
    }
    $store[$Tenant] = $Id
    $store | ConvertTo-Json | Set-Content -Path $appIdStore -Encoding UTF8
    Write-Host "  Client ID cached in $appIdStore" -ForegroundColor DarkGray
}

function New-FolderAccessApp {
    <#
        Creates (or reuses) a public-client app in the target tenant and
        admin-consents the delegated SharePoint scope PnP needs. Returns the app id.
        Requires a Graph sign-in as Global / Application Administrator.
    #>
    param([string] $Tenant, [string] $DisplayName)

    foreach ($module in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications')) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            throw "Module '$module' is required to create the app registration. Run: Install-Module $module -Scope CurrentUser"
        }
    }
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop

    Write-Host '  No app registration known for this tenant - creating one.' -ForegroundColor Yellow
    Write-Host "  Sign in as a Global Administrator of $Tenant." -ForegroundColor Yellow

    $bootstrap = Connect-M365Graph -TenantId $Tenant -Scopes @(
        'Application.ReadWrite.All'
        'DelegatedPermissionGrant.ReadWrite.All'
        'Directory.Read.All'
    )

    Write-Host "  Signed in as $((Get-MgContext).Account)" -ForegroundColor Green

    $escaped = $DisplayName -replace "'", "''"
    $app = @(Get-MgApplication -Filter "displayName eq '$escaped'" -All) | Select-Object -First 1

    if ($app) {
        Write-Host "  Reusing existing app '$DisplayName' ($($app.AppId))" -ForegroundColor Green
    } else {
        $app = New-MgApplication -DisplayName $DisplayName `
            -SignInAudience 'AzureADMyOrg' `
            -IsFallbackPublicClient `
            -PublicClient @{ RedirectUris = @('http://localhost') }
        Write-Host "  App created: $($app.AppId)" -ForegroundColor Green
    }

    $sp = @(Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -All) | Select-Object -First 1
    if (-not $sp) {
        $sp = New-MgServicePrincipal -AppId $app.AppId
        Write-Host '  Service principal created.' -ForegroundColor Green
    }

    # Delegated scopes: SharePoint for the CSOM calls PnP makes, Graph for sign-in.
    $resources = @(
        @{ AppId = '00000003-0000-0ff1-ce00-000000000000'; Name = 'SharePoint'; Scopes = @('AllSites.FullControl', 'User.Read.All') }
        @{ AppId = '00000003-0000-0000-c000-000000000000'; Name = 'Graph';      Scopes = @('User.Read') }
    )

    foreach ($resource in $resources) {
        $resourceSp = @(Get-MgServicePrincipal -Filter "appId eq '$($resource.AppId)'" -All) | Select-Object -First 1
        if (-not $resourceSp) {
            Write-Warning "  Service principal for $($resource.Name) not found - skipped."
            continue
        }

        $valid = @($resource.Scopes | Where-Object { $resourceSp.Oauth2PermissionScopes.Value -contains $_ })
        if ($valid.Count -eq 0) { continue }

        $grant = @(Get-MgOauth2PermissionGrant -All -Filter "clientId eq '$($sp.Id)' and consentType eq 'AllPrincipals'") |
                 Where-Object { $_.ResourceId -eq $resourceSp.Id } | Select-Object -First 1

        $existing = if ($grant -and $grant.Scope) { $grant.Scope -split ' ' } else { @() }
        $missing  = @($valid | Where-Object { $_ -notin $existing })
        if ($missing.Count -eq 0) {
            Write-Host "    $($resource.Name): scopes already consented." -ForegroundColor DarkGray
            continue
        }

        $merged = ((@($existing) + $valid) | Where-Object { $_ } | Select-Object -Unique) -join ' '
        if ($grant) {
            Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $grant.Id -Scope $merged | Out-Null
        } else {
            New-MgOauth2PermissionGrant -ClientId $sp.Id -ResourceId $resourceSp.Id `
                -ConsentType 'AllPrincipals' -Scope $merged | Out-Null
        }
        Write-Host "    $($resource.Name): consented $($missing -join ', ')" -ForegroundColor Green
    }

    Disconnect-M365Graph $bootstrap
    return $app.AppId
}

# -AppOnly brings its own certificate app from graph.appid.json; nothing to register.
$appIsNew = $false
if (-not $AppOnly) {
    if (-not $ClientId) {
        $ClientId = Get-CachedClientId -Tenant $TenantId
        if ($ClientId) { Write-Host "  Using cached app registration $ClientId" -ForegroundColor DarkGray }
    }
    if (-not $ClientId) {
        $ClientId = New-FolderAccessApp -Tenant $TenantId -DisplayName $AppName
        Set-CachedClientId -Tenant $TenantId -Id $ClientId
        $appIsNew = $true
    }
}

# -- Connect -------------------------------------------------------------------
if ($appIsNew) {
    Write-Host '  Waiting 20s for the new app registration to propagate...' -ForegroundColor DarkGray
    Start-Sleep -Seconds 20
}

$connectParams = @{ TenantId = $TenantId }
if ($ClientId) { $connectParams['ClientId'] = $ClientId }
if ($AppOnly)  { $connectParams['AppOnly']  = $true }

function Connect-Site {
    <#
        A freshly created app registration is not replicated everywhere yet, so the first
        sign-in can fail with "application not found". Retry briefly.
    #>
    param([string] $Url, [int] $Retries = 6)

    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        try {
            # Device code or browser per load.config.ps1 - see Connect-M365PnP.
            return Connect-M365PnP -Url $Url @connectParams
        } catch {
            if ($attempt -eq $Retries) { throw }
            Write-Host "  Sign-in attempt $attempt failed ($($_.Exception.Message.Trim())) - retrying in 10s..." -ForegroundColor DarkYellow
            Start-Sleep -Seconds 10
        }
    }
}

# -GrantSiteAdmin: a Global Administrator cannot read a team's files by role alone. Become site
# collection admin through the tenant admin site, and give it back in the finally at the end.
$grantedSiteAdmin = $false
$script:graph     = $null
if ($GrantSiteAdmin) {
    Write-Host "  Connecting to $adminUrl..." -ForegroundColor Cyan
    $adminConnection = Connect-Site -Url $adminUrl
    if (-not $AdminUpn) {
        try {
            # Loaded with the web, not through Get-PnPProperty: that cmdlet wants a "current" PnP
            # connection before it looks at the object, and this script only holds returned ones -
            # the first live -GrantSiteAdmin run failed on exactly that ("You are not signed in").
            $currentUser = (Get-PnPWeb -Connection $adminConnection -Includes CurrentUser).CurrentUser
            $AdminUpn = if ($currentUser.Email) { $currentUser.Email } else { ($currentUser.LoginName -split '\|')[-1] }
            Write-Host "  Signed in as $AdminUpn" -ForegroundColor Green
        } catch {
            throw "Could not determine the signed-in account - pass -AdminUpn explicitly. $($_.Exception.Message)"
        }
    }
    try {
        Set-PnPTenantSite -Identity $SiteUrl -Owners @($AdminUpn) -Connection $adminConnection -ErrorAction Stop
        $grantedSiteAdmin = $true
        Write-Host "  Granted site collection admin to $AdminUpn on $SiteUrl" -ForegroundColor DarkGray
    } catch {
        throw "Could not grant site collection admin on ${SiteUrl} (SharePoint Administrator needed): $($_.Exception.Message)"
    }
}

$connection = Connect-Site -Url $SiteUrl

# Everything after the grant runs inside this try, so the site collection admin rights are
# given back however the run ends - a typo in an address or a failed change included.
try {
    $context = Get-PnPContext -Connection $connection
    try {
        $web = Get-PnPWeb -Connection $connection -Includes ServerRelativeUrl, AssociatedOwnerGroup, CurrentUser
    } catch {
        if (-not $GrantSiteAdmin -and $_.Exception.Message -match 'unauthorized|access denied|403') {
            throw "No access to $SiteUrl with this account. A Global Administrator cannot read a team's files by role alone - run again with -GrantSiteAdmin to be site collection admin for the duration of the run. ($($_.Exception.Message))"
        }
        throw
    }
    $ownerGroupId = if ($web.AssociatedOwnerGroup -and -not $web.AssociatedOwnerGroup.ServerObjectIsNull) { $web.AssociatedOwnerGroup.Id } else { $null }
    if (-not $ownerGroupId) {
        Write-Warn 'This site has no associated Owners group. Only the configured people and site collection admins will have access.'
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
    if (-not $GroupPrefix) {
        # Entra group names are tenant-wide, so the site goes in; SharePoint groups live in the site.
        $GroupPrefix = if ($useEntra) { "$($SiteUrl.Split('/')[-1]) - $($relative.Split('/')[-1])" } else { $relative.Split('/')[-1] }
    }

    # -- Roles, by type so the site language does not matter -----------------------
    $roleDefs = @{}
    foreach ($pair in @(
            @('Read',        [Microsoft.SharePoint.Client.RoleType]::Reader),
            @('Contribute',  [Microsoft.SharePoint.Client.RoleType]::Contributor),
            @('Edit',        [Microsoft.SharePoint.Client.RoleType]::Editor),
            @('FullControl', [Microsoft.SharePoint.Client.RoleType]::Administrator))) {
        $def = $web.RoleDefinitions.GetByType($pair[1])
        $context.Load($def)
        $roleDefs[$pair[0]] = $def
    }
    $context.ExecuteQuery()
    $roleRank = @{ Read = 1; Contribute = 2; Edit = 3; FullControl = 4 }

    # -- Microsoft Graph, for the Entra groups ------------------------------------
    if ($useEntra) {
        # Read-only for a dry run: a report has no business holding write access to every group.
        $graphScopes = if ($Apply) { @('Group.ReadWrite.All', 'User.Read.All') } else { @('Group.Read.All', 'User.Read.All') }
        $script:graph = Connect-M365Graph -TenantId $TenantId -Scopes $graphScopes
        Write-Host "  Graph signed in as $($script:graph.Account)" -ForegroundColor DarkGray
    }

    # -- Resolve everyone before touching anything ---------------------------------
    # One unknown address must stop the run here, not halfway through, with the parent already
    # closed off and half the subfolders done.
    Write-Host '  Resolving people...' -ForegroundColor Cyan
    $principals = @{}
    $entraUsers = @{}
    $unresolved = @()
    $everyone = @($ownerLogins) + @($plan.Values | ForEach-Object { $_.Members }) | Sort-Object -Unique
    foreach ($login in $everyone) {
        try {
            $user = $web.EnsureUser($login)
            $context.Load($user)
            $context.ExecuteQuery()
            $principals[$login] = $user
            if ($useEntra) {
                # By UPN or primary address - the config may hold either.
                $escaped = $login -replace "'", "''"
                $filter  = [uri]::EscapeDataString("userPrincipalName eq '$escaped' or mail eq '$escaped'")
                # An "or" across two properties is an advanced query: it needs the eventual-consistency header.
                $hits = @(Invoke-M365GraphPaged "/v1.0/users?`$filter=$filter&`$count=true&`$select=id,displayName,userPrincipalName" -Headers @{ ConsistencyLevel = 'eventual' })
                if ($hits.Count -ne 1) { throw "$($hits.Count) Entra users match this address" }
                $entraUsers[$login] = $hits[0]
            }
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
    function Test-Owner { param([string] $Login) return $ownerKeys.ContainsKey($Login.ToLowerInvariant()) }

    # -- Helpers -------------------------------------------------------------------
    function New-Grant {
        <# One desired assignment. Principal is $null for a group that a dry run has not created. #>
        param([string] $Name, $Principal, [string] $Role)
        [pscustomobject]@{ Name = $Name; Principal = $Principal; Role = $Role }
    }

    function Merge-Grants {
        <# One grant per principal, the highest role winning: an owner who is also listed as a
           member must end up with Full Control, not with whichever entry came last. #>
        param([array] $Grants)
        $byKey = [ordered]@{}
        foreach ($g in $Grants) {
            $key = if ($g.Principal) { "id:$($g.Principal.Id)" } else { "name:$($g.Name)" }
            if (-not $byKey.Contains($key) -or $roleRank[$g.Role] -gt $roleRank[$byKey[$key].Role]) { $byKey[$key] = $g }
        }
        return @($byKey.Values)
    }

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
            Bring one item in line: unique permissions, site owners Full Control, exactly the desired
            grants with exactly their role, nobody else. Returns the number of changes (planned in a
            dry run, made with -Apply).
        #>
        param(
            [string] $Label,
            $Item,
            [array] $Desired,
            [switch] $IsNew
        )

        Write-Host ''
        Write-Host "  $Label" -ForegroundColor Cyan
        $Desired = Merge-Grants $Desired

        if ($IsNew -and -not $Apply) {
            Write-Change 'create folder'
            Add-Result -Folder $Label -Action 'Create folder' -Result 'Planned'
            Write-Change 'break inheritance, keep only the site owners'
            foreach ($g in $Desired) {
                Write-Change "grant $($g.Role) to $($g.Name)"
                Add-Result -Folder $Label -Action 'Grant' -Principal $g.Name -Detail $g.Role -Result 'Planned'
            }
            return 2 + $Desired.Count
        }

        $changes     = 0
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
            }
            Add-Result -Folder $Label -Action 'Break inheritance' -Result $status
        }

        # Site owners first, so the item can never end up without them.
        if ($ownerGroupId -and -not ($assignments | Where-Object { $_.Member.Id -eq $ownerGroupId })) {
            $changes++
            Write-Change 'grant Full Control to the site Owners group'
            if ($Apply) {
                $bindings = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($context)
                $bindings.Add($roleDefs['FullControl'])
                $Item.RoleAssignments.Add($web.AssociatedOwnerGroup, $bindings) | Out-Null
                $context.ExecuteQuery()
            }
            Add-Result -Folder $Label -Action 'Grant' -Principal 'Site Owners' -Detail 'FullControl' -Result $status
        }

        $desiredById = @{}
        foreach ($g in $Desired | Where-Object Principal) { $desiredById[$g.Principal.Id] = $g }

        $seen = @{}
        foreach ($a in $assignments) {
            $member = $a.Member
            $roles  = @($a.RoleDefinitionBindings | Where-Object { -not $_.Hidden } | ForEach-Object { $_.Name })

            if ($desiredById.ContainsKey($member.Id)) {
                $g = $desiredById[$member.Id]
                $seen[$member.Id] = $true
                $want = $roleDefs[$g.Role]
                $have = @($a.RoleDefinitionBindings | Where-Object { -not $_.Hidden })
                if ($have.Count -eq 1 -and $have[0].Id -eq $want.Id) {
                    Write-Ok "$($g.Name) has $($want.Name)"
                    Add-Result -Folder $Label -Action 'Check' -Principal $g.Name -Detail $want.Name -Result 'Already correct'
                    continue
                }
                $changes++
                Write-Change "set $($g.Name) to $($want.Name) (was: $($roles -join ', '))"
                if ($Apply) {
                    $a.RoleDefinitionBindings.RemoveAll()
                    $a.RoleDefinitionBindings.Add($want)
                    $a.Update()
                    $context.ExecuteQuery()
                }
                Add-Result -Folder $Label -Action 'Change role' -Principal $g.Name -Detail "$($roles -join ', ') -> $($want.Name)" -Result $status
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

        foreach ($g in $Desired) {
            if ($g.Principal -and $seen[$g.Principal.Id]) { continue }
            $changes++
            $want = $roleDefs[$g.Role]
            Write-Change "grant $($want.Name) to $($g.Name)"
            if ($Apply) {
                $bindings = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($context)
                $bindings.Add($want)
                $Item.RoleAssignments.Add($g.Principal, $bindings) | Out-Null
                $context.ExecuteQuery()
            }
            Add-Result -Folder $Label -Action 'Grant' -Principal $g.Name -Detail $want.Name -Result $status
        }

        if ($changes -eq 0) { Write-Ok 'already correct' }
        return $changes
    }

    function Get-FolderItem {
        <# The list item behind a folder, or $null when the folder does not exist. #>
        param([string] $Path)
        try { return (Get-PnPFolder -Url $Path -Connection $connection -Includes ListItemAllFields).ListItemAllFields }
        catch { return $null }
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

    function Get-SiteGroup {
        param([string] $Title)
        try { return Get-PnPGroup -Identity $Title -Connection $connection } catch { return $null }
    }

    function Set-SiteGroup {
        <#
            Make sure a SharePoint group exists, holds at least the configured people (exactly them
            with -SyncMembers) and is owned by $OwnerTitle. Returns the group, or $null in a dry run
            when it does not exist yet.
        #>
        param([string] $Title, [string[]] $Logins, [string] $OwnerTitle)

        Write-Host ''
        Write-Host "  Group '$Title'" -ForegroundColor Cyan
        $changes = 0
        $group   = Get-SiteGroup -Title $Title
        $created = $false

        if (-not $group) {
            $changes++
            Write-Change 'create group'
            Add-Result -Folder '' -Action 'Create group' -Principal $Title -Result $status
            if (-not $Apply) {
                foreach ($login in $Logins) {
                    Write-Change "add $login"
                    Add-Result -Folder '' -Action 'Add member' -Principal $login -Detail $Title -Result 'Planned'
                }
                Write-Change "owner: $OwnerTitle"
                $script:groupChanges += $changes + $Logins.Count
                return $null
            }
            $group   = New-PnPGroup -Title $Title -Description "Managed by Set-SharePointFolderAccess.ps1 for $relative" -Connection $connection
            $created = $true
        }

        # Members, by login name, compared case-insensitively.
        $current = @{}
        foreach ($m in @(Get-PnPGroupMember -Group $group -Connection $connection)) { $current[$m.LoginName.ToLowerInvariant()] = $m }
        $wanted = @{}
        foreach ($login in $Logins) { $wanted[$principals[$login].LoginName.ToLowerInvariant()] = $login }

        foreach ($key in $wanted.Keys) {
            if ($current.ContainsKey($key)) { continue }
            $changes++
            $login = $wanted[$key]
            Write-Change "add $login"
            if ($Apply) { Add-PnPGroupMember -Group $group -LoginName $principals[$login].LoginName -Connection $connection | Out-Null }
            Add-Result -Folder '' -Action 'Add member' -Principal $login -Detail $Title -Result $status
        }
        foreach ($key in $current.Keys) {
            if ($wanted.ContainsKey($key)) { continue }
            $member = $current[$key]
            # Creating a group can put the signed-in admin in it; that one always goes.
            $isCreator = $created -and $web.CurrentUser -and $member.Id -eq $web.CurrentUser.Id
            if ($SyncMembers -or $isCreator) {
                $changes++
                Write-Change "remove $($member.Title)"
                if ($Apply) { Remove-PnPGroupMember -Group $group -LoginName $member.LoginName -Connection $connection | Out-Null }
                Add-Result -Folder '' -Action 'Remove member' -Principal $member.LoginName -Detail $Title -Result $status
            } else {
                Write-Warn "$($member.Title) is in the group but not in the config - left in (-SyncMembers removes)"
                Add-Result -Folder '' -Action 'Extra member' -Principal $member.LoginName -Detail $Title -Result 'Left in'
            }
        }

        $owner = Get-PnPProperty -ClientObject $group -Property Owner -Connection $connection
        if ("$($owner.Title)" -ne $OwnerTitle) {
            $changes++
            Write-Change "owner: $OwnerTitle (was: $($owner.Title))"
            if ($Apply) { Set-PnPGroup -Identity $group -Owner $OwnerTitle -Connection $connection | Out-Null }
            Add-Result -Folder '' -Action 'Set owner' -Principal $Title -Detail $OwnerTitle -Result $status
        }

        if ($changes -eq 0) { Write-Ok 'already correct' }
        $script:groupChanges += $changes
        return $group
    }

    function Invoke-GraphWrite {
        <# A Graph write, retried while a group created seconds ago is not replicated yet (404). #>
        param([string] $Method, [string] $Uri, $Body)
        for ($i = 1; $i -le 6; $i++) {
            try {
                $p = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
                if ($null -ne $Body) { $p['Body'] = ($Body | ConvertTo-Json -Depth 5); $p['ContentType'] = 'application/json' }
                return Invoke-MgGraphRequest @p
            } catch {
                if ($i -eq 6 -or "$($_.Exception.Message)" -notmatch '404|NotFound|does not exist') { throw }
                Start-Sleep -Seconds 5
            }
        }
    }

    function Get-SharePointGroupPrincipal {
        <# The SharePoint principal for an Entra security group, by its claim. EnsureUser adds the
           site's entry for it; a group created seconds ago may not be known to SharePoint yet. #>
        param([string] $ObjectId, [string] $Title)
        for ($i = 1; $i -le 6; $i++) {
            try {
                $principal = $web.EnsureUser("c:0t.c|tenant|$ObjectId")
                $context.Load($principal)
                $context.ExecuteQuery()
                return $principal
            } catch {
                if ($i -eq 6) { throw "SharePoint does not know the Entra group '$Title' ($ObjectId) yet: $($_.Exception.Message)" }
                Write-Host "    SharePoint does not see '$Title' yet - retrying in 10s..." -ForegroundColor DarkYellow
                Start-Sleep -Seconds 10
            }
        }
    }

    function Set-EntraGroup {
        <#
            Make sure an Entra security group exists, holds at least the configured members and
            owners (exactly them with -SyncMembers). Returns its SharePoint principal, or $null in
            a dry run when the group does not exist yet.
        #>
        param([string] $Title, [string[]] $Logins, [string[]] $OwnerList)

        Write-Host ''
        Write-Host "  Entra group '$Title'" -ForegroundColor Cyan
        $changes = 0

        $escaped = $Title -replace "'", "''"
        $filter  = [uri]::EscapeDataString("displayName eq '$escaped'")
        $found   = @(Invoke-M365GraphPaged "/v1.0/groups?`$filter=$filter&`$select=id,displayName,securityEnabled,mailEnabled,groupTypes,onPremisesSyncEnabled")
        if ($found.Count -gt 1) { throw "There are $($found.Count) Entra groups called '$Title'. Remove the extra ones, or pick another groupPrefix." }
        $group = $found | Select-Object -First 1
        if ($group) {
            # Refused rather than forced: each of these would fail halfway or do the wrong thing.
            if ($group.onPremisesSyncEnabled) { throw "Entra group '$Title' is synced from on-premises AD and cannot be changed in the cloud. Pick another groupPrefix." }
            if (-not $group.securityEnabled)  { throw "Entra group '$Title' exists but is not a security group, so SharePoint cannot grant it access. Pick another groupPrefix." }
            if (@($group.groupTypes) -contains 'DynamicMembership') { throw "Entra group '$Title' has dynamic membership; its members follow a rule. Pick another groupPrefix." }
        }

        $created = $false
        if (-not $group) {
            $changes++
            Write-Change 'create security group'
            Add-Result -Folder '' -Action 'Create Entra group' -Principal $Title -Result $status
            if (-not $Apply) {
                foreach ($login in $Logins)    { Write-Change "add member $login"; Add-Result -Folder '' -Action 'Add member' -Principal $login -Detail $Title -Result 'Planned' }
                foreach ($login in $OwnerList) { Write-Change "add owner $login";  Add-Result -Folder '' -Action 'Add owner'  -Principal $login -Detail $Title -Result 'Planned' }
                $script:groupChanges += $changes + $Logins.Count + $OwnerList.Count
                return $null
            }
            $nick = ($Title -replace '[^A-Za-z0-9]', '')
            if ($nick.Length -gt 50) { $nick = $nick.Substring(0, 50) }
            $body = @{
                displayName     = $Title
                description     = "Managed by Set-SharePointFolderAccess.ps1 - $SiteUrl/$relative"
                mailEnabled     = $false
                mailNickname    = "$nick-$([guid]::NewGuid().ToString('N').Substring(0, 6))"
                securityEnabled = $true
            }
            # Owners set at creation, so the signed-in admin is not made owner as a side effect.
            if ($OwnerList) { $body['owners@odata.bind'] = @($OwnerList | ForEach-Object { "https://graph.microsoft.com/v1.0/users/$($entraUsers[$_].id)" }) }
            $group   = Invoke-GraphWrite -Method POST -Uri '/v1.0/groups' -Body $body
            $created = $true
            foreach ($login in $OwnerList) { Add-Result -Folder '' -Action 'Add owner' -Principal $login -Detail $Title -Result 'Done' }
        }
        $script:entraGroupIds[$Title] = $group.id

        # Members, and then owners, by object id.
        foreach ($kind in @('members', 'owners')) {
            $wantLogins = if ($kind -eq 'members') { $Logins } else { $OwnerList }
            if ($kind -eq 'owners' -and $created) { continue }   # set at creation above

            $current = @{}
            if (-not $created) {
                foreach ($m in @(Invoke-M365GraphPaged "/v1.0/groups/$($group.id)/$kind`?`$select=id,displayName,userPrincipalName")) { $current[$m.id] = $m }
            }
            $wanted = @{}
            foreach ($login in $wantLogins) { $wanted[$entraUsers[$login].id] = $login }
            $label = $kind.TrimEnd('s')

            foreach ($id in $wanted.Keys) {
                if ($current.ContainsKey($id)) { continue }
                $changes++
                Write-Change "add $label $($wanted[$id])"
                if ($Apply) { Invoke-GraphWrite -Method POST -Uri "/v1.0/groups/$($group.id)/$kind/`$ref" -Body @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$id" } | Out-Null }
                Add-Result -Folder '' -Action "Add $label" -Principal $wanted[$id] -Detail $Title -Result $status
            }
            foreach ($id in $current.Keys) {
                if ($wanted.ContainsKey($id)) { continue }
                $who = if ($current[$id].userPrincipalName) { $current[$id].userPrincipalName } else { $current[$id].displayName }
                # An owner is only removed when the configured owners are there, so the group is never left ownerless.
                if ($SyncMembers -and ($kind -eq 'members' -or $wanted.Count -gt 0)) {
                    $changes++
                    Write-Change "remove $label $who"
                    if ($Apply) { Invoke-GraphWrite -Method DELETE -Uri "/v1.0/groups/$($group.id)/$kind/$id/`$ref" | Out-Null }
                    Add-Result -Folder '' -Action "Remove $label" -Principal $who -Detail $Title -Result $status
                } else {
                    Write-Warn "$who is $label of the group but not in the config - left in (-SyncMembers removes)"
                    Add-Result -Folder '' -Action "Extra $label" -Principal $who -Detail $Title -Result 'Left in'
                }
            }
        }

        if ($changes -eq 0) { Write-Ok 'already correct' }
        $script:groupChanges += $changes
        return Get-SharePointGroupPrincipal -ObjectId $group.id -Title $Title
    }

    # Wrapped so the CSV is written even when a change fails halfway: the report then shows what
    # was done before it stopped, which is what you need to judge the state the folder is in.
    $total = 0
    $script:entraGroupIds = [ordered]@{}
    $script:groupChanges = 0
    $groupsMade = [ordered]@{}
    try {
        # -- Who gets what -------------------------------------------------------------
        $ownerGrants = @()
        $teamGrants  = @{}   # subfolder -> grants
        if ($DirectAccess) {
            foreach ($login in $ownerLogins) { $ownerGrants += New-Grant -Name $login -Principal $principals[$login] -Role 'FullControl' }
            foreach ($name in $plan.Keys) {
                $teamGrants[$name] = @($plan[$name].Members | ForEach-Object { New-Grant -Name $_ -Principal $principals[$_] -Role $Role })
            }
        } elseif ($useEntra) {
            if (-not $ownerLogins) { Write-Warn 'No owners in the config: the Entra groups get no owner, so only admins can manage their members.' }
            if ($ownerLogins) {
                $title = "$GroupPrefix - Owners"
                $g = Set-EntraGroup -Title $title -Logins $ownerLogins -OwnerList $ownerLogins
                $ownerGrants += New-Grant -Name $title -Principal $g -Role 'FullControl'
            }
            foreach ($name in $plan.Keys) {
                $title = "$GroupPrefix - $name"
                $g = Set-EntraGroup -Title $title -Logins $plan[$name].Members -OwnerList $ownerLogins
                $teamGrants[$name] = @(New-Grant -Name $title -Principal $g -Role $Role)
            }
        } else {
            $ownerTitle = if ($ownerLogins) { "$GroupPrefix - Owners" }
                          elseif ($ownerGroupId) { $web.AssociatedOwnerGroup.Title }
                          else { $null }
            if (-not $ownerTitle) { throw 'Groups need an owner: list "owners" in the config, or use a site that has an Owners group.' }

            if ($ownerLogins) {
                $g = Set-SiteGroup -Title $ownerTitle -Logins $ownerLogins -OwnerTitle $ownerTitle
                $groupsMade[$ownerTitle] = $g
                $ownerGrants += New-Grant -Name $ownerTitle -Principal $g -Role 'FullControl'
            }
            foreach ($name in $plan.Keys) {
                $title = "$GroupPrefix - $name"
                $g = Set-SiteGroup -Title $title -Logins $plan[$name].Members -OwnerTitle $ownerTitle
                $groupsMade[$title] = $g
                $teamGrants[$name] = @(New-Grant -Name $title -Principal $g -Role $Role)
            }
        }

        # -- Parent folder -------------------------------------------------------------
        $parentGrants = @($ownerGrants)
        if ($ParentAccess -eq 'Read') {
            foreach ($name in $plan.Keys) {
                $parentGrants += @($teamGrants[$name] | ForEach-Object { New-Grant -Name $_.Name -Principal $_.Principal -Role 'Read' })
            }
        }
        $total += Set-FolderAccess -Label $relative -Item $parentFolder.ListItemAllFields -Desired $parentGrants

        # -- Subfolders ----------------------------------------------------------------
        $others = @()
        foreach ($f in @(Get-PnPFolderItem -FolderSiteRelativeUrl $relative -Connection $connection)) {
            if (-not ($f -is [Microsoft.SharePoint.Client.Folder] -and $plan.Contains($f.Name))) { $others += $f }
        }

        $managed = [ordered]@{}
        foreach ($name in $plan.Keys) {
            $path  = "$relative/$name"
            $item  = Get-FolderItem -Path $path
            $isNew = -not $item
            if ($isNew -and $Apply) {
                Add-PnPFolder -Name $name -Folder $relative -Connection $connection | Out-Null
                Add-Result -Folder $path -Action 'Create folder' -Result 'Done'
                Write-Host ''
                Write-Host "  Created $path" -ForegroundColor Magenta
                $total++
                $item = Get-FolderItem -Path $path
            }
            $total += Set-FolderAccess -Label $path -Item $item -Desired (@($ownerGrants) + @($teamGrants[$name])) -IsNew:($isNew -and -not $Apply)
            if ($item) { $managed[$path] = $item }

            # Personal folders: one per listed non-owner (owners too with "all"), only that person
            # and the owners.
            if (-not $plan[$name].Personal) { continue }
            # Compared as a string: $true -eq 'all' is $true in PowerShell, which would make every
            # "personalFolders": true behave like "all".
            $withOwners = "$($plan[$name].Personal)" -eq 'all'
            foreach ($login in $plan[$name].Members | Where-Object { $withOwners -or -not (Test-Owner $_) }) {
                $folderName = ($principals[$login].Title -replace $invalidName, '-').Trim().TrimEnd('.')
                if (-not $folderName) { $folderName = ($login -split '@')[0] }
                $personalPath = "$path/$folderName"
                $pItem  = if ($isNew) { $null } else { Get-FolderItem -Path $personalPath }
                $pIsNew = -not $pItem
                if ($pIsNew -and $Apply) {
                    Add-PnPFolder -Name $folderName -Folder $path -Connection $connection | Out-Null
                    Add-Result -Folder $personalPath -Action 'Create folder' -Result 'Done'
                    Write-Host ''
                    Write-Host "  Created $personalPath" -ForegroundColor Magenta
                    $total++
                    $pItem = Get-FolderItem -Path $personalPath
                }
                $personalGrants = @($ownerGrants) + @(New-Grant -Name $login -Principal $principals[$login] -Role 'Contribute')
                $total += Set-FolderAccess -Label $personalPath -Item $pItem -Desired $personalGrants -IsNew:($pIsNew -and -not $Apply)
                if ($pItem) { $managed[$personalPath] = $pItem }
            }
        }

        # Anything else directly in the parent - files, subfolders the config does not name - inherits
        # the parent's Read, so everyone listed anywhere could open it. That is more than "only their
        # own subfolder", so it is locked to the owners: the same reconcile with only the owners desired.
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
            $total += Set-FolderAccess -Label "$relative/$($f.Name) ($kind, not in the config: owners only)" -Item $item -Desired $ownerGrants
        }

        # -- Read back -----------------------------------------------------------------
        if ($Apply) {
            Write-Host ''
            Write-Host '  Permissions now in place' -ForegroundColor Cyan
            Show-FolderAccess -Label $relative -Item $parentFolder.ListItemAllFields
            foreach ($path in $managed.Keys) { Show-FolderAccess -Label $path -Item $managed[$path] }

            if ($script:entraGroupIds.Count -gt 0) {
                Write-Host ''
                Write-Host '  Entra groups - the owners manage members at https://myaccount.microsoft.com/groups' -ForegroundColor Cyan
                foreach ($title in $script:entraGroupIds.Keys) {
                    Write-Host ("    {0,-40} {1}" -f $title, $script:entraGroupIds[$title]) -ForegroundColor DarkGray
                }
                Write-Host '    A membership change can take a while to reach SharePoint.' -ForegroundColor DarkGray
            }
            if ($groupsMade.Count -gt 0) {
                Write-Host ''
                Write-Host '  Group pages - the owners add and remove members here' -ForegroundColor Cyan
                foreach ($title in $groupsMade.Keys) {
                    $g = Get-SiteGroup -Title $title
                    if ($g) { Write-Host ("    {0,-36} {1}/_layouts/15/people.aspx?MembershipGroupId={2}" -f $title, $SiteUrl, $g.Id) -ForegroundColor DarkGray }
                }
            }
        }
    } finally {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    }
} finally {
    if ($script:graph) { Disconnect-M365Graph $script:graph }
    if ($grantedSiteAdmin -and -not $KeepSiteAdmin) {
        try {
            Remove-PnPSiteCollectionAdmin -Owners @($AdminUpn) -Connection $connection -ErrorAction Stop
            Write-Host "  Removed site collection admin rights for $AdminUpn again." -ForegroundColor DarkGray
        } catch {
            Write-Warning "$AdminUpn is still site collection admin on ${SiteUrl}: $($_.Exception.Message)"
        }
    }
}

# -- Summary -------------------------------------------------------------------
$total += $script:groupChanges
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
