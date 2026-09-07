#Requires -Version 7.0
<#
.SYNOPSIS
    Search an entire SharePoint site for content - by name, path, type, size, date
    or full text - and report exactly which permissions apply to every hit.

.DESCRIPTION
    Answers the two questions you usually have at the same time: *where does this
    live* and *who can get at it*.

    Two search engines:

      Crawl  (default)  Walks every list and library in the site item by item and
                        filters client-side on -Name, -Path, -Extension, -ItemType,
                        -ModifiedBy and the date/size filters. Sees everything,
                        including items the search index has not picked up yet.

      Search (-Content) Runs a KQL query against the SharePoint search index scoped
                        to the site, so it also matches text *inside* documents.
                        Fast, but only finds what has been indexed and what the
                        signed-in account may see. The other filters still apply on
                        top of the search results.

    For every hit the script then resolves the permissions:

      Item   the item has broken inheritance and carries its own role assignments
      List   the item inherits from its list/library, which has unique permissions
      Site   the item inherits all the way up to the (sub)site

    Role assignments are flattened to one CSV row per principal, with the principal
    type, login name, e-mail and the role names ("Full Control", "Edit", ...).
    Noise is trimmed: "Limited Access" assignments are hidden unless you pass
    -IncludeLimitedAccess.

    Three kinds of access are called out separately because they are the ones that
    cause surprises:

      Sharing links    the SharingLinks.* groups SharePoint creates behind every
                       "Copy link". Always expanded to the people in them, and
                       labelled Anyone / Organization / Specific people.
      External users   guest accounts (#ext#) in any role assignment
      Everyone         "Everyone" and "Everyone except external users"

    Nothing is changed - this script only reads.

    Sign-in: PnP.PowerShell no longer ships a shared multi-tenant app, so an Entra
    app registration is required. The first run against a tenant creates one and
    caches the client ID in pnp.appid.json (gitignored) - a cached ID from another
    script in this repo is reused. Pass -ClientId to skip app creation entirely.

.PARAMETER SiteUrl
    Full URL of the site collection to search, e.g.
    https://contoso.sharepoint.com/sites/Finance
    A user's OneDrive works too:
    https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com

.PARAMETER IncludeSubsites
    Also search every subsite below -SiteUrl. Without it only the top web is
    searched; the script tells you when subsites exist that it skipped.

.PARAMETER Content
    Free-text/KQL query run against the search index - this is what finds words
    *inside* documents. Switches the script from crawling to searching.
    Examples: "salarisschaal", "vertrouwelijk AND 2026", "author:jane".

.PARAMETER Name
    Filter on the name. Wildcards allowed, e.g. "*offerte*" or "Budget*.xlsx".
    Matched against the file name, the item title and the last segment of the URL,
    so an item that carries a different title than its file name is still found.

.PARAMETER Path
    Filter on the folder the item lives in - substring match on the server relative
    URL, e.g. "Gedeelde documenten/Directie".

.PARAMETER Extension
    One or more file extensions to keep, with or without the dot: "xlsx","pdf".

.PARAMETER ItemType
    Restrict to File, Folder or ListItem. Default All.

.PARAMETER ListName
    Only search these lists/libraries (by title). Wildcards allowed. Repeatable.

.PARAMETER ModifiedBy
    Filter on who last changed the item - display name or e-mail, wildcards allowed.

.PARAMETER ModifiedAfter
    Only items changed on or after this date/time.

.PARAMETER ModifiedBefore
    Only items changed before this date/time.

.PARAMETER MinSizeMB
    Only files of at least this size.

.PARAMETER IncludeHidden
    Also search hidden lists, catalogs and system libraries. Off by default - they
    are almost never what you are looking for and they are big.

.PARAMETER Everything
    Leave nothing out: every subsite, every hidden and system list, and no cap on
    the number of hits or on how many of them get their permissions resolved.
    Equivalent to -IncludeSubsites -IncludeHidden -MaxItems 0 -MaxPermissionLookups 0.
    Note that a crawl still matches names and metadata only - add -Content to look
    inside the documents themselves.

.PARAMETER Permissions
    How much permission detail to resolve per hit:
      Effective  (default) the permissions that actually apply, whether they come
                 from the item, its list or the site
      Unique     only report permissions where inheritance is broken - the fastest
                 way to find "what is shared differently from the rest"
      None       just find the content, do not touch permissions

.PARAMETER ExpandGroups
    Also list the members of SharePoint groups that hold permissions. Sharing-link
    groups are always expanded; this switch adds the regular ones ("Site Members",
    ...). Costs one extra call per group, cached per run.

.PARAMETER IncludeLimitedAccess
    Keep "Limited Access" role assignments in the report. These are the bookkeeping
    entries SharePoint adds so someone can reach a deeper item; they grant nothing
    on their own and are hidden by default.

.PARAMETER MaxItems
    Stop after this many matching items. Default 5000; 0 means no limit.

.PARAMETER MaxPermissionLookups
    Cap on how many hits get their permissions resolved (default 1000, 0 means no
    limit). Resolving permissions costs a few server calls per item, so a very broad
    search would otherwise run for hours. Hits above the cap are still reported,
    without permission rows.

.PARAMETER PageSize
    Items fetched per server call while crawling (default 500, max 5000).

.PARAMETER GrantSiteAdmin
    Add the signed-in admin as site collection administrator before searching
    (requires SharePoint Administrator). Needed to read another user's OneDrive or
    a site you are not a member of. The rights are removed again afterwards unless
    -KeepSiteAdmin.

.PARAMETER KeepSiteAdmin
    Keep the site collection admin rights granted by -GrantSiteAdmin.

.PARAMETER AdminUpn
    UPN to grant site collection admin rights to. Defaults to the signed-in account.

.PARAMETER TenantId
    Tenant ID or domain for the one-time Graph admin sign-in. Defaults to the tenant
    derived from -SiteUrl.

.PARAMETER ClientId
    Client ID of an existing Entra app to sign in with. Skips app registration.

.PARAMETER AppName
    Display name of the app registration to create or reuse.
    Default: "M365-Scripts SharePoint Search".

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\SharePointFind_<timestamp>.csv
    (~/Downloads on non-Windows).

.PARAMETER Disconnect
    Sign out of PnP when finished.

.EXAMPLE
    # Where does anything with "offerte" in the name live, and who can see it?
    .\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Sales -Name "*offerte*"

.EXAMPLE
    # Full text: which documents mention "salarisschaal", anywhere in the site tree?
    .\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/HR `
        -Content "salarisschaal"

.EXAMPLE
    # Everything in the site that is shared differently from the rest
    .\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
        -Permissions Unique -IncludeSubsites

.EXAMPLE
    # Large PDFs in one library, with the groups behind the permissions expanded
    .\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
        -ListName "Gedeelde documenten" -Extension pdf -MinSizeMB 10 -ExpandGroups

.EXAMPLE
    # Leave nothing out: all subsites, all hidden and system lists, no caps
    .\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
        -Name "*veiligheid*" -Everything

.EXAMPLE
    # Search someone's OneDrive you have no rights on
    .\Find-SiteContent.ps1 `
        -SiteUrl https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com `
        -Name "*.xlsx" -GrantSiteAdmin
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $SiteUrl,

    [switch] $IncludeSubsites,

    [string] $Content,
    [string] $Name,
    [string] $Path,
    [string[]] $Extension,

    [ValidateSet('All', 'File', 'Folder', 'ListItem')]
    [string] $ItemType = 'All',

    [string[]] $ListName,
    [string] $ModifiedBy,
    [datetime] $ModifiedAfter,
    [datetime] $ModifiedBefore,
    [double] $MinSizeMB,

    [switch] $IncludeHidden,
    [switch] $Everything,

    [ValidateSet('Effective', 'Unique', 'None')]
    [string] $Permissions = 'Effective',

    [switch] $ExpandGroups,
    [switch] $IncludeLimitedAccess,

    [ValidateRange(0, [int]::MaxValue)]
    [int] $MaxItems = 5000,

    [ValidateRange(0, [int]::MaxValue)]
    [int] $MaxPermissionLookups = 1000,

    [ValidateRange(100, 5000)]
    [int] $PageSize = 500,

    [switch] $GrantSiteAdmin,
    [switch] $KeepSiteAdmin,
    [string] $AdminUpn,

    [string] $TenantId,
    [string] $ClientId,
    [string] $AppName = 'M365-Scripts SharePoint Search',

    [string] $OutputPath,
    [switch] $Disconnect
)

$ErrorActionPreference = 'Stop'

# A terminating error anywhere below should name the line it came from - without it
# PowerShell reports only the script name, which is no help in a script this size.
trap {
    Write-Host ''
    Write-Host "  FAILED at line $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    Write-Host "    $($_.InvocationInfo.Line.Trim())" -ForegroundColor DarkRed
    Write-Host "    $($_.Exception.GetType().Name): $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ''
    break
}

# -- Leave nothing out ---------------------------------------------------------
# -Everything is the "do not make me think about it" switch: every subsite, every
# hidden and system list, and no caps on what is reported.
if ($Everything) {
    $IncludeSubsites = $true
    $IncludeHidden   = $true
    if (-not $PSBoundParameters.ContainsKey('MaxItems'))             { $MaxItems = 0 }
    if (-not $PSBoundParameters.ContainsKey('MaxPermissionLookups')) { $MaxPermissionLookups = 0 }
}

# -- Modules -------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name 'PnP.PowerShell')) {
    throw "Module 'PnP.PowerShell' is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser"
}
Import-Module PnP.PowerShell -ErrorAction Stop

# -- Output folder -------------------------------------------------------------
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
if (-not $OutputPath) {
    $OutputPath = Join-Path $outputDir "SharePointFind_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
}

# -- Resolve tenant + admin URL ------------------------------------------------
$SiteUrl = $SiteUrl.Trim().TrimEnd('/')
if ($SiteUrl -notmatch '^https://([a-zA-Z0-9-]+?)(-admin|-my)?\.sharepoint\.(com|de|us|cn)(/|$)') {
    throw "'$SiteUrl' does not look like a SharePoint or OneDrive URL."
}
$tenantName = $Matches[1]
$tld        = $Matches[3]
$adminUrl   = "https://$tenantName-admin.sharepoint.$tld"
if (-not $TenantId) { $TenantId = "$tenantName.onmicrosoft.com" }

$repoRoot   = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$appIdStore = Join-Path $repoRoot 'pnp.appid.json'

$searchMode = [bool]$Content

# Normalise the extension filter once: "xlsx", ".xlsx" and "XLSX" all mean the same.
$extensionFilter = @()
if ($Extension) {
    $extensionFilter = @($Extension | ForEach-Object { $_.TrimStart('.').ToLowerInvariant() } | Where-Object { $_ })
}

$useModifiedAfter  = $PSBoundParameters.ContainsKey('ModifiedAfter')
$useModifiedBefore = $PSBoundParameters.ContainsKey('ModifiedBefore')
$useMinSize        = $PSBoundParameters.ContainsKey('MinSizeMB')

Write-Host ''
Write-Host "  Site   : $SiteUrl" -ForegroundColor Cyan
Write-Host "  Engine : $(if ($searchMode) { "search index - $Content" } else { 'crawl - every list and library' })" -ForegroundColor Cyan
$filterParts = @()
if ($Name)               { $filterParts += "name $Name" }
if ($Path)               { $filterParts += "path *$Path*" }
if ($extensionFilter)    { $filterParts += "ext $($extensionFilter -join ',')" }
if ($ItemType -ne 'All') { $filterParts += "type $ItemType" }
if ($ListName)           { $filterParts += "list $($ListName -join ', ')" }
if ($ModifiedBy)         { $filterParts += "modified by $ModifiedBy" }
if ($useModifiedAfter)   { $filterParts += "after $($ModifiedAfter.ToString('yyyy-MM-dd HH:mm'))" }
if ($useModifiedBefore)  { $filterParts += "before $($ModifiedBefore.ToString('yyyy-MM-dd HH:mm'))" }
if ($useMinSize)         { $filterParts += "min $MinSizeMB MB" }
if ($filterParts.Count)  { Write-Host "  Filter : $($filterParts -join ' | ')" -ForegroundColor Cyan }
Write-Host "  Rights : $Permissions" -ForegroundColor Cyan
if ($Everything) {
    Write-Host '  Scope  : EVERYTHING - all subsites, hidden and system lists included, no caps' -ForegroundColor Yellow
}
if (-not $searchMode) {
    Write-Host '  Note   : a crawl matches names and metadata - add -Content to search inside documents.' -ForegroundColor DarkGray
}
Write-Host ''

# -- Helpers -------------------------------------------------------------------
function Get-ClientProperty {
    <#
        CSOM only ships the properties it was asked for, and PnP's -Includes cannot
        express value types - asking it for Hidden, ItemCount or BaseType fails with
        "Argument types do not match". So read the property and only go back to the
        server when it turns out not to be loaded.
    #>
    param($ClientObject, [string] $Property, $Default = $null)

    try { return $ClientObject.$Property } catch { }
    try {
        Get-PnPProperty -ClientObject $ClientObject -Property $Property -ErrorAction Stop | Out-Null
        return $ClientObject.$Property
    } catch {
        return $Default
    }
}

function Format-Duration {
    param([TimeSpan] $Span)

    if ($Span.TotalSeconds -lt 60) { return "{0:n0}s" -f $Span.TotalSeconds }
    if ($Span.TotalMinutes -lt 60) { return "{0}m {1}s" -f [int]$Span.TotalMinutes, $Span.Seconds }
    return "{0}h {1}m" -f [int]$Span.TotalHours, $Span.Minutes
}

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

function New-SearchApp {
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

    Connect-MgGraph -TenantId $Tenant -NoWelcome -ContextScope Process -Scopes @(
        'Application.ReadWrite.All'
        'DelegatedPermissionGrant.ReadWrite.All'
        'Directory.Read.All'
    ) | Out-Null

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

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return $app.AppId
}

if (-not $ClientId) {
    $ClientId = Get-CachedClientId -Tenant $TenantId
    if ($ClientId) { Write-Host "  Using cached app registration $ClientId" -ForegroundColor DarkGray }
}

$appIsNew = $false
if (-not $ClientId) {
    $ClientId = New-SearchApp -Tenant $TenantId -DisplayName $AppName
    Set-CachedClientId -Tenant $TenantId -Id $ClientId
    $appIsNew = $true
}

# -- Connections ---------------------------------------------------------------
# One connection per web URL, cached, so the admin connection and the per-web
# connections live side by side without reconnecting all the time.
$script:Connections = @{}

function Connect-Site {
    <#
        A freshly created app registration is not replicated everywhere yet, so the
        first sign-in can fail with "application not found". Retry briefly.
    #>
    param([string] $Url, [int] $Retries = 6)

    $key = $Url.TrimEnd('/').ToLowerInvariant()
    if ($script:Connections.ContainsKey($key)) { return $script:Connections[$key] }

    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        try {
            $connection = Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId -ReturnConnection -ErrorAction Stop
            $script:Connections[$key] = $connection
            return $connection
        } catch {
            if ($attempt -eq $Retries) { throw }
            Write-Host "  Sign-in attempt $attempt failed ($($_.Exception.Message.Trim())) - retrying in 10s..." -ForegroundColor DarkYellow
            Start-Sleep -Seconds 10
        }
    }
}

if ($appIsNew) {
    Write-Host '  Waiting 20s for the new app registration to propagate...' -ForegroundColor DarkGray
    Start-Sleep -Seconds 20
}

# Permission caches - webs and lists are resolved once, items only when they broke
# inheritance. Script scope so the helper functions share them.
$script:Unreadable = [System.Collections.Generic.List[object]]::new()
$script:PermCache  = @{}
$script:GroupCache = @{}
$script:Webs       = @()
$script:SkipRoles  = if ($IncludeLimitedAccess) { @() } else { @('Limited Access', 'Beperkte toegang', 'Web-Only Limited Access') }

function Get-SharingLinkKind {
    param([string] $Title)

    switch -Regex ($Title) {
        '\.AnonymousEdit\.'    { return 'Anyone - edit' }
        '\.AnonymousView\.'    { return 'Anyone - view' }
        '\.OrganizationEdit\.' { return 'Organization - edit' }
        '\.OrganizationView\.' { return 'Organization - view' }
        '\.Flexible\.'         { return 'Specific people' }
        default                { return 'Sharing link' }
    }
}

function Get-GroupMemberSummary {
    param([string] $GroupTitle, $Connection)

    if ($script:GroupCache.ContainsKey($GroupTitle)) { return $script:GroupCache[$GroupTitle] }

    $summary = ''
    try {
        $members = @(Get-PnPGroupMember -Group $GroupTitle -Connection $Connection -ErrorAction Stop)
        $summary = ($members | ForEach-Object { if ($_.Email) { $_.Email } else { $_.Title } }) -join '; '
    } catch {
        $summary = "<could not read members: $($_.Exception.Message.Trim())>"
    }
    $script:GroupCache[$GroupTitle] = $summary
    return $summary
}

function Get-PrincipalRows {
    <#
        Flattens the role assignments of a securable object (web, list or item) into
        one row per principal.
    #>
    param($SecurableObject, $Connection)

    $rows = [System.Collections.Generic.List[object]]::new()

    try {
        Get-PnPProperty -ClientObject $SecurableObject -Property RoleAssignments -ErrorAction Stop | Out-Null
    } catch {
        $rows.Add([pscustomobject]@{
            Principal = '<could not read>'; PrincipalType = ''; Login = ''; Email = ''
            Permission = ''; SharingLink = ''; IsExternal = $false; IsEveryone = $false
            Members = $_.Exception.Message.Trim()
        })
        return $rows
    }

    foreach ($assignment in $SecurableObject.RoleAssignments) {
        try {
            Get-PnPProperty -ClientObject $assignment -Property Member, RoleDefinitionBindings -ErrorAction Stop | Out-Null
        } catch { continue }

        $roles = @($assignment.RoleDefinitionBindings | ForEach-Object { $_.Name } | Where-Object { $_ -notin $script:SkipRoles })
        if ($roles.Count -eq 0) { continue }

        $member = $assignment.Member
        $title  = "$($member.Title)"
        $login  = "$($member.LoginName)"
        $type   = "$($member.PrincipalType)"

        $sharingLink = ''
        $members     = ''

        if ($title -like 'SharingLinks.*') {
            # The people a link was actually shared with only show up inside the group.
            $sharingLink = Get-SharingLinkKind -Title $title
            $members     = Get-GroupMemberSummary -GroupTitle $title -Connection $Connection
            $title       = $sharingLink
        } elseif ($type -eq 'SharePointGroup' -and $ExpandGroups) {
            $members = Get-GroupMemberSummary -GroupTitle $title -Connection $Connection
        }

        $email = ''
        try { if ($member.Email) { $email = "$($member.Email)" } } catch { }

        $isEveryone = $login -match 'spo-grid-all-users' -or $login -eq 'c:0(.s|true' -or
                      $title -in @('Everyone', 'Everyone except external users',
                                   'Iedereen', 'Iedereen behalve externe gebruikers')
        $isExternal = $login -match '#ext#|urn:spo:guest' -or $members -match '#ext#'

        $rows.Add([pscustomobject]@{
            Principal     = $title
            PrincipalType = $type
            Login         = ($login -split '\|')[-1]
            Email         = $email
            Permission    = ($roles -join ', ')
            SharingLink   = $sharingLink
            IsExternal    = [bool]$isExternal
            IsEveryone    = [bool]$isEveryone
            Members       = $members
        })
    }

    return $rows
}

function Resolve-WebPermissions {
    param([string] $WebUrl)

    $key = "web:$WebUrl"
    if ($script:PermCache.ContainsKey($key)) { return $script:PermCache[$key] }

    $entry = [pscustomobject]@{ Source = 'Site'; SourceName = $WebUrl; Rows = @() }
    try {
        $connection = Connect-Site -Url $WebUrl
        $web = Get-PnPWeb -Connection $connection
        $unique = [bool](Get-ClientProperty -ClientObject $web -Property 'HasUniqueRoleAssignments' -Default $true)

        $parentUrl = ($script:Webs | Where-Object { $_.Url -eq $WebUrl } | Select-Object -First 1).ParentUrl
        if ($unique -or -not $parentUrl) {
            $entry.Rows = @(Get-PrincipalRows -SecurableObject $web -Connection $connection)
        } else {
            $parentEntry = Resolve-WebPermissions -WebUrl $parentUrl
            $entry.SourceName = $parentEntry.SourceName
            $entry.Rows       = $parentEntry.Rows
        }
    } catch {
        Write-Warning "Could not read site permissions on ${WebUrl}: $($_.Exception.Message)"
    }

    $script:PermCache[$key] = $entry
    return $entry
}

function Resolve-ListPermissions {
    param($List, [string] $WebUrl, $Connection)

    $key = "list:$($List.Id)"
    if ($script:PermCache.ContainsKey($key)) { return $script:PermCache[$key] }

    $entry = $null
    try {
        Get-PnPProperty -ClientObject $List -Property HasUniqueRoleAssignments -ErrorAction Stop | Out-Null
        if ($List.HasUniqueRoleAssignments) {
            $entry = [pscustomobject]@{
                Source     = 'List'
                SourceName = $List.Title
                Rows       = @(Get-PrincipalRows -SecurableObject $List -Connection $Connection)
            }
        }
    } catch {
        Write-Warning "Could not read permissions on list '$($List.Title)': $($_.Exception.Message)"
    }

    if (-not $entry) { $entry = Resolve-WebPermissions -WebUrl $WebUrl }

    $script:PermCache[$key] = $entry
    return $entry
}

function Resolve-ItemPermissions {
    <#
        Returns the permission entry that actually applies to one item: its own if
        inheritance is broken, otherwise the list's, otherwise the web's. With
        -Permissions Unique an inheriting item returns nothing at all.
    #>
    param($Item, $List, [string] $WebUrl, $Connection, [string] $Mode)

    try {
        Get-PnPProperty -ClientObject $Item -Property HasUniqueRoleAssignments -ErrorAction Stop | Out-Null
    } catch {
        Write-Warning "Could not read permissions on item $($Item.Id) in '$($List.Title)': $($_.Exception.Message)"
        return $null
    }

    if ($Item.HasUniqueRoleAssignments) {
        return [pscustomobject]@{
            Source     = 'Item'
            SourceName = 'this item'
            Rows       = @(Get-PrincipalRows -SecurableObject $Item -Connection $Connection)
        }
    }

    if ($Mode -eq 'Unique') { return $null }
    return Resolve-ListPermissions -List $List -WebUrl $WebUrl -Connection $Connection
}

# -- Matching ------------------------------------------------------------------
function Test-Match {
    param($Hit)

    # A file has a name, a list item has a title, and plenty of items have both with
    # different values - so -Name matches either, plus the leaf of the URL.
    if ($Name) {
        $candidates = @($Hit.Name, $Hit.AltName, [System.IO.Path]::GetFileName("$($Hit.Url)")) |
            Where-Object { $_ }
        if (-not @($candidates | Where-Object { $_ -like $Name })) { return $false }
    }
    if ($Path -and ($Hit.Url -notlike "*$Path*")) { return $false }
    if ($ItemType -ne 'All' -and $Hit.ItemType -ne $ItemType) { return $false }

    if ($extensionFilter.Count -gt 0) {
        $ext = [System.IO.Path]::GetExtension("$($Hit.Name)").TrimStart('.').ToLowerInvariant()
        if ($ext -notin $extensionFilter) { return $false }
    }

    if ($ModifiedBy) {
        if ("$($Hit.ModifiedBy)" -notlike $ModifiedBy -and "$($Hit.ModifiedByEmail)" -notlike $ModifiedBy) { return $false }
    }

    if ($useModifiedAfter  -and $Hit.Modified -and $Hit.Modified -lt $ModifiedAfter)  { return $false }
    if ($useModifiedBefore -and $Hit.Modified -and $Hit.Modified -ge $ModifiedBefore) { return $false }
    if ($useMinSize -and (($Hit.SizeBytes / 1MB) -lt $MinSizeMB)) { return $false }

    return $true
}

# -- Run -----------------------------------------------------------------------
$adminConnection  = $null
$grantedSiteAdmin = $false

try {
    if ($GrantSiteAdmin) {
        Write-Host "  Connecting to $adminUrl..." -ForegroundColor Cyan
        $adminConnection = Connect-Site -Url $adminUrl

        if (-not $AdminUpn) {
            try {
                $currentUser = Get-PnPProperty -ClientObject (Get-PnPWeb -Connection $adminConnection) -Property CurrentUser
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
            Write-Warning "Could not grant site collection admin on ${SiteUrl}: $($_.Exception.Message)"
        }
    }

    Write-Host "  Connecting to $SiteUrl..." -ForegroundColor Cyan
    $siteConnection = Connect-Site -Url $SiteUrl

    # -- Webs to search --------------------------------------------------------
    $rootWeb = Get-PnPWeb -Connection $siteConnection
    $webList = [System.Collections.Generic.List[object]]::new()
    $webList.Add([pscustomobject]@{ Url = "$($rootWeb.Url)".TrimEnd('/'); Title = $rootWeb.Title; ParentUrl = $null })

    $subWebs = @()
    try {
        $subWebs = @(Get-PnPSubWeb -Recurse -Connection $siteConnection -ErrorAction Stop)
    } catch {
        Write-Warning "Could not enumerate subsites: $($_.Exception.Message)"
    }

    if ($subWebs.Count -gt 0 -and -not $IncludeSubsites -and -not $searchMode) {
        Write-Host "  Note   : $($subWebs.Count) subsite(s) found and skipped - add -IncludeSubsites to search them too." -ForegroundColor Yellow
    }

    if ($IncludeSubsites -or $searchMode) {
        # A subweb's URL is always its parent's URL plus one segment, so the parent
        # map is rebuilt from the URLs alone - no extra round trips.
        $knownUrls = [System.Collections.Generic.List[string]]@("$($rootWeb.Url)".TrimEnd('/'))
        foreach ($sub in ($subWebs | Sort-Object { "$($_.Url)".Length })) {
            $url = "$($sub.Url)".TrimEnd('/')
            $parent = $knownUrls |
                Where-Object { $url.StartsWith("$_/", [StringComparison]::OrdinalIgnoreCase) } |
                Sort-Object { $_.Length } -Descending | Select-Object -First 1
            $webList.Add([pscustomobject]@{ Url = $url; Title = $sub.Title; ParentUrl = $parent })
            $knownUrls.Add($url)
        }
        if ($IncludeSubsites) {
            Write-Host "  Webs   : $($webList.Count) (site + $($webList.Count - 1) subsite(s))" -ForegroundColor Cyan
        }
    }
    # ToArray, not @(): on PowerShell 7.6.5 / .NET 10 an array subexpression over a
    # List[object] throws "Argument types do not match".
    # $script:Webs and $webs would also be one and the same variable at script scope -
    # PowerShell ignores case - hence the separate name for the webs to crawl.
    $script:Webs = $webList.ToArray()
    $targetWebs  = if ($IncludeSubsites) { $webList.ToArray() } else { @($webList[0]) }
    Write-Host ''

    # -- Collect hits ----------------------------------------------------------
    $hits    = [System.Collections.Generic.List[object]]::new()
    $scanned = 0
    $capped  = $false
    $timer   = [System.Diagnostics.Stopwatch]::StartNew()

    if ($searchMode) {
        # KQL: scoped to the site path so the query cannot wander off into the tenant.
        $kql = "$Content path:`"$SiteUrl`""
        Write-Host "  Query  : $kql" -ForegroundColor DarkGray
        Write-Host '  Note   : the search index covers the whole site tree, subsites included.' -ForegroundColor DarkGray

        $select = 'Title,Path,FileExtension,Size,LastModifiedTime,Created,Author,EditorOWSUSER,ListID,ListItemID,SPWebUrl,IsDocument,IsContainer'
        $rows = @()
        try {
            $result = Submit-PnPSearchQuery -Query $kql -All -SelectProperties $select -Connection $siteConnection -ErrorAction Stop
            $rows = @($result.ResultRows)
        } catch {
            throw "Search failed: $($_.Exception.Message)"
        }

        foreach ($row in $rows) {
            $scanned++
            $itemUrl     = "$($row['Path'])"
            $isContainer = "$($row['IsContainer'])" -eq 'True'
            $isDocument  = "$($row['IsDocument'])"  -eq 'True'

            $modified = $null
            if ($row['LastModifiedTime']) { try { $modified = [datetime]$row['LastModifiedTime'] } catch { } }
            $created = $null
            if ($row['Created']) { try { $created = [datetime]$row['Created'] } catch { } }

            $hit = [pscustomobject]@{
                WebUrl          = if ($row['SPWebUrl']) { "$($row['SPWebUrl'])".TrimEnd('/') } else { $SiteUrl }
                List            = ''
                ListId          = "$($row['ListID'])"
                ItemId          = if ($row['ListItemID']) { [int]$row['ListItemID'] } else { 0 }
                ItemType        = if ($isContainer) { 'Folder' } elseif ($isDocument) { 'File' } else { 'ListItem' }
                Name            = if ($row['Title']) { "$($row['Title'])" } else { [System.IO.Path]::GetFileName($itemUrl) }
                AltName         = [System.IO.Path]::GetFileName($itemUrl)
                Url             = $itemUrl
                SizeBytes       = if ($row['Size']) { [long]$row['Size'] } else { 0 }
                Modified        = $modified
                ModifiedBy      = "$($row['EditorOWSUSER'])" -replace '^.*\|', ''
                ModifiedByEmail = ''
                Created         = $created
                CreatedBy       = "$($row['Author'])"
                ItemObject      = $null
                ListObject      = $null
                Connection      = $null
            }

            if (-not (Test-Match -Hit $hit)) { continue }
            if ($MaxItems -gt 0 -and $hits.Count -ge $MaxItems) { $capped = $true; break }
            $hits.Add($hit)
        }

        Write-Host "  Search returned $scanned result(s) in $(Format-Duration $timer.Elapsed); $($hits.Count) matched the filters." -ForegroundColor Cyan
    } else {
        # System libraries that are never what someone is looking for.
        $systemLists = @('Form Templates', 'Site Assets', 'Style Library', 'Composed Looks',
                         'Master Page Gallery', 'Preservation Hold Library', 'Sharing Links',
                         'TaxonomyHiddenList', 'User Information List')

        $webIndex = 0
        foreach ($web in $targetWebs) {
            $webIndex++
            if ($capped) { break }

            # No access to a subsite must not kill the run - record it and carry on,
            # so the summary can say the search was not complete.
            $webConnection = $null
            $lists = @()
            try {
                $webConnection = Connect-Site -Url $web.Url
                $lists = @(Get-PnPList -Connection $webConnection)
            } catch {
                Write-Warning "No access to $($web.Url): $($_.Exception.Message)"
                $script:Unreadable.Add([pscustomobject]@{ Scope = $web.Url; Reason = $_.Exception.Message.Trim() })
                continue
            }

            $skippedHidden = 0
            $skippedByName = 0
            $targetLists = @($lists | Where-Object {
                if (-not $IncludeHidden) {
                    $hidden = [bool](Get-ClientProperty -ClientObject $_ -Property 'Hidden' -Default $false)
                    if ($hidden -or $_.Title -in $systemLists) { $skippedHidden++; return $false }
                }
                if ($ListName) {
                    $title = $_.Title
                    if (-not @($ListName | Where-Object { $title -like $_ })) { $skippedByName++; return $false }
                }
                return $true
            })

            $skipNote = @()
            if ($skippedHidden) { $skipNote += "$skippedHidden hidden/system" }
            if ($skippedByName) { $skipNote += "$skippedByName filtered out" }
            $skipText = if ($skipNote) { " (skipped: $($skipNote -join ', '))" } else { '' }
            Write-Host "  [$webIndex/$($targetWebs.Count)] $($web.Url) - $($targetLists.Count) of $($lists.Count) list(s)/librar(ies)$skipText" -ForegroundColor DarkGray

            $listIndex = 0
            foreach ($list in $targetLists) {
                $listIndex++
                if ($capped) { break }

                $isDocLib  = "$(Get-ClientProperty -ClientObject $list -Property 'BaseType')" -eq 'DocumentLibrary'
                $itemCount = Get-ClientProperty -ClientObject $list -Property 'ItemCount' -Default 0
                Write-Progress -Id 1 -Activity "Crawling $($web.Url)" `
                    -Status "$($list.Title) ($listIndex/$($targetLists.Count)) - $itemCount item(s), $($hits.Count) hit(s), $(Format-Duration $timer.Elapsed)" `
                    -PercentComplete ([math]::Min(100, ($listIndex / [math]::Max(1, $targetLists.Count)) * 100))

                $fields = @('ID', 'Title', 'FileLeafRef', 'FileRef', 'FileDirRef', 'FSObjType', 'Modified', 'Created', 'Author', 'Editor')
                if ($isDocLib) { $fields += 'File_x0020_Size' }

                try {
                    $items = @(Get-PnPListItem -List $list -PageSize $PageSize -Fields $fields -Connection $webConnection -ErrorAction Stop)
                } catch {
                    Write-Warning "Skipping list '$($list.Title)' on $($web.Url): $($_.Exception.Message)"
                    $script:Unreadable.Add([pscustomobject]@{ Scope = "$($web.Url) > $($list.Title)"; Reason = $_.Exception.Message.Trim() })
                    continue
                }

                foreach ($item in $items) {
                    $scanned++
                    $fv = $item.FieldValues

                    $type     = if ("$($fv.FSObjType)" -eq '1') { 'Folder' } elseif ($isDocLib) { 'File' } else { 'ListItem' }
                    $itemName = if ($isDocLib -or -not $fv.Title) { "$($fv.FileLeafRef)" } else { "$($fv.Title)" }
                    $editor   = $fv.Editor
                    $author   = $fv.Author

                    $hit = [pscustomobject]@{
                        WebUrl          = $web.Url
                        List            = $list.Title
                        ListId          = "$($list.Id)"
                        ItemId          = [int]$item.Id
                        ItemType        = $type
                        Name            = $itemName
                        AltName         = if ($isDocLib) { "$($fv.Title)" } else { "$($fv.FileLeafRef)" }
                        Url             = "$($fv.FileRef)"
                        SizeBytes       = if ($fv.File_x0020_Size) { [long]$fv.File_x0020_Size } else { 0 }
                        Modified        = $fv.Modified
                        ModifiedBy      = if ($editor) { "$($editor.LookupValue)" } else { '' }
                        ModifiedByEmail = if ($editor -and $editor.Email) { "$($editor.Email)" } else { '' }
                        Created         = $fv.Created
                        CreatedBy       = if ($author) { "$($author.LookupValue)" } else { '' }
                        ItemObject      = $item
                        ListObject      = $list
                        Connection      = $webConnection
                    }

                    if (-not (Test-Match -Hit $hit)) { continue }
                    if ($MaxItems -gt 0 -and $hits.Count -ge $MaxItems) { $capped = $true; break }
                    $hits.Add($hit)
                }
            }
            Write-Progress -Id 1 -Activity "Crawling $($web.Url)" -Completed
        }

        Write-Host "  Scanned $scanned item(s) in $(Format-Duration $timer.Elapsed); $($hits.Count) matched." -ForegroundColor Cyan
    }

    if ($capped) {
        Write-Host "  Stopped at -MaxItems $MaxItems - narrow the filters or raise the cap for the full picture." -ForegroundColor Yellow
    }

    if ($hits.Count -eq 0) {
        Write-Host ''
        Write-Host '  Nothing matched. Nothing to report.' -ForegroundColor Yellow
        Write-Host ''
        return
    }

    # -- Resolve permissions ---------------------------------------------------
    $results   = [System.Collections.Generic.List[object]]::new()
    $permTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $resolved  = 0

    function Add-Row {
        param($Hit, $Entry, $Principal)

        $results.Add([pscustomobject]@{
            Site             = $Hit.WebUrl
            List             = $Hit.List
            ItemType         = $Hit.ItemType
            Name             = $Hit.Name
            Url              = $Hit.Url
            SizeMB           = if ($Hit.SizeBytes) { [math]::Round($Hit.SizeBytes / 1MB, 2) } else { 0 }
            Modified         = $Hit.Modified
            ModifiedBy       = $Hit.ModifiedBy
            Created          = $Hit.Created
            CreatedBy        = $Hit.CreatedBy
            ItemId           = $Hit.ItemId
            PermissionSource = if ($Entry) { $Entry.Source } else { '' }
            InheritedFrom    = if ($Entry -and $Entry.Source -ne 'Item') { $Entry.SourceName } else { '' }
            UniqueRights     = [bool]($Entry -and $Entry.Source -eq 'Item')
            Principal        = if ($Principal) { $Principal.Principal } else { '' }
            PrincipalType    = if ($Principal) { $Principal.PrincipalType } else { '' }
            PrincipalLogin   = if ($Principal) { $Principal.Login } else { '' }
            PrincipalEmail   = if ($Principal) { $Principal.Email } else { '' }
            Permission       = if ($Principal) { $Principal.Permission } else { '' }
            SharingLink      = if ($Principal) { $Principal.SharingLink } else { '' }
            External         = if ($Principal) { $Principal.IsExternal } else { $false }
            Everyone         = if ($Principal) { $Principal.IsEveryone } else { $false }
            GroupMembers     = if ($Principal) { $Principal.Members } else { '' }
        })
    }

    if ($Permissions -eq 'None') {
        foreach ($hit in $hits) { Add-Row -Hit $hit -Entry $null -Principal $null }
    } else {
        $index = 0
        foreach ($hit in $hits) {
            $index++
            Write-Progress -Id 2 -Activity 'Resolving permissions' `
                -Status "$index/$($hits.Count) - $($hit.Name) - $(Format-Duration $permTimer.Elapsed)" `
                -PercentComplete ([math]::Min(100, ($index / $hits.Count) * 100))

            if ($MaxPermissionLookups -gt 0 -and $resolved -ge $MaxPermissionLookups) {
                Add-Row -Hit $hit -Entry $null -Principal $null
                continue
            }

            # Crawled hits carry their web connection; search hits only a web URL,
            # which is empty for the odd result that has no web of its own.
            $connection = if ($hit.Connection) {
                $hit.Connection
            } elseif ($hit.WebUrl) {
                Connect-Site -Url $hit.WebUrl
            } else {
                $siteConnection
            }
            $item = $hit.ItemObject
            $list = $hit.ListObject

            # Search results only carry ids, so the item and its list are fetched here.
            if (-not $item) {
                if (-not $hit.ListId -or -not $hit.ItemId) {
                    Add-Row -Hit $hit -Entry $null -Principal $null
                    continue
                }
                try {
                    $list = Get-PnPList -Identity $hit.ListId -Connection $connection -ErrorAction Stop
                    $item = Get-PnPListItem -List $hit.ListId -Id $hit.ItemId -Connection $connection -ErrorAction Stop
                    $hit.List = $list.Title
                } catch {
                    Write-Warning "Could not open $($hit.Url): $($_.Exception.Message)"
                    Add-Row -Hit $hit -Entry $null -Principal $null
                    continue
                }
            }

            $entry = Resolve-ItemPermissions -Item $item -List $list -WebUrl $hit.WebUrl -Connection $connection -Mode $Permissions
            $resolved++

            if (-not $entry) { continue }   # -Permissions Unique and this item inherits

            if ($entry.Rows.Count -eq 0) {
                Add-Row -Hit $hit -Entry $entry -Principal $null
            } else {
                foreach ($principal in $entry.Rows) { Add-Row -Hit $hit -Entry $entry -Principal $principal }
            }
        }
        Write-Progress -Id 2 -Activity 'Resolving permissions' -Completed
    }
    $permTimer.Stop()
    $timer.Stop()

    # -- Report ----------------------------------------------------------------
    if ($results.Count -gt 0) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    }

    $permissionRows = @($results | Where-Object { $_.Principal })
    $uniqueItems    = @($results | Where-Object { $_.UniqueRights } | Group-Object Url).Count
    $linkRows       = @($results | Where-Object { $_.SharingLink })
    $externalRows   = @($results | Where-Object { $_.External })
    $everyoneRows   = @($results | Where-Object { $_.Everyone })

    Write-Host ''
    if ($Permissions -eq 'Unique' -and $permissionRows.Count -eq 0) {
        Write-Host "  $($hits.Count) item(s) matched and every one of them inherits - nothing is shared differently." -ForegroundColor Green
    } else {
        $hits | Group-Object List | Sort-Object Count -Descending | Select-Object -First 15 |
            ForEach-Object {
                [pscustomobject]@{
                    List = if ($_.Name) { $_.Name } else { '(from search index)' }
                    Hits = $_.Count
                }
            } | Format-Table -AutoSize | Out-Host

        $preview = @($permissionRows | Select-Object -First 25)
        if ($preview.Count -gt 0) {
            $preview | Select-Object `
                @{ N = 'Item';   E = { if ("$($_.Name)".Length -gt 38) { "$($_.Name)".Substring(0, 37) + '...' } else { $_.Name } } },
                @{ N = 'Src';    E = { $_.PermissionSource } },
                @{ N = 'Who';    E = { if ("$($_.Principal)".Length -gt 34) { "$($_.Principal)".Substring(0, 33) + '...' } else { $_.Principal } } },
                @{ N = 'Rights'; E = { $_.Permission } },
                @{ N = 'Link';   E = { $_.SharingLink } } |
                Format-Table -AutoSize | Out-Host

            if ($permissionRows.Count -gt $preview.Count) {
                Write-Host "  ... $($permissionRows.Count - $preview.Count) more permission row(s) in the CSV." -ForegroundColor DarkGray
            }
        }
    }

    Write-Host "  Items found      : $($hits.Count)" -ForegroundColor Cyan
    if ($Permissions -ne 'None') {
        Write-Host "  Unique rights on : $uniqueItems item(s)" -ForegroundColor $(if ($uniqueItems) { 'Yellow' } else { 'Green' })
        if ($linkRows.Count)     { Write-Host "  Sharing links    : $($linkRows.Count) - $((($linkRows.SharingLink | Select-Object -Unique) -join ', '))" -ForegroundColor Yellow }
        if ($externalRows.Count) { Write-Host "  External access  : $($externalRows.Count) assignment(s)" -ForegroundColor Red }
        if ($everyoneRows.Count) { Write-Host "  Everyone (-ish)  : $($everyoneRows.Count) assignment(s)" -ForegroundColor Yellow }
        if ($MaxPermissionLookups -gt 0 -and $resolved -ge $MaxPermissionLookups) {
            Write-Host "  Note             : stopped resolving rights after $MaxPermissionLookups item(s) (-MaxPermissionLookups)" -ForegroundColor Yellow
        }
    }

    # "Did we actually get everywhere?" - answer it, do not leave it to the warnings
    # that scrolled past.
    if ($script:Unreadable.Count -gt 0) {
        Write-Host "  Not readable     : $($script:Unreadable.Count) - the result is INCOMPLETE" -ForegroundColor Red
        $script:Unreadable | Select-Object -First 10 |
            Select-Object @{ N = 'Skipped'; E = { $_.Scope } }, @{ N = 'Why'; E = { $_.Reason } } |
            Format-Table -AutoSize | Out-Host
        if ($script:Unreadable.Count -gt 10) { Write-Host "  ... and $($script:Unreadable.Count - 10) more." -ForegroundColor DarkGray }
        Write-Host '  Tip: -GrantSiteAdmin makes you site collection admin for the duration of the run.' -ForegroundColor DarkGray
    } else {
        Write-Host '  Access           : everything in scope was readable' -ForegroundColor Green
    }
    Write-Host "  Duration         : $(Format-Duration $timer.Elapsed) total, of which $(Format-Duration $permTimer.Elapsed) on permissions" -ForegroundColor Cyan
    if ($results.Count -gt 0) {
        Write-Host "  Report           : $OutputPath" -ForegroundColor Cyan
    } else {
        Write-Host '  Report           : not written - nothing to report' -ForegroundColor DarkGray
    }
    Write-Host ''

} finally {
    if ($grantedSiteAdmin -and -not $KeepSiteAdmin) {
        try {
            $siteConn = Connect-Site -Url $SiteUrl
            Remove-PnPSiteCollectionAdmin -Owners @($AdminUpn) -Connection $siteConn -ErrorAction Stop
            Write-Host "  Removed site collection admin rights for $AdminUpn again." -ForegroundColor DarkGray
        } catch {
            Write-Warning "$AdminUpn is still site collection admin on ${SiteUrl}: $($_.Exception.Message)"
        }
    }
    if ($Disconnect) {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Host '  Disconnected.' -ForegroundColor DarkGray
    }
}
