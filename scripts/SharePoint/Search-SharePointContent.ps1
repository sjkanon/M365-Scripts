#Requires -Version 7.0
<#
.SYNOPSIS
    Search SharePoint and OneDrive content tenant-wide through Microsoft Graph -
    app-only, no interactive sign-in - and report the permissions on every hit.

.DESCRIPTION
    The Graph counterpart of Find-SiteContent.ps1. Where that script drives CSOM
    through PnP and signs in as you, this one runs on an app registration with
    application permissions: no interactive login, no site collection admin rights
    to arrange, and it reaches every site and every OneDrive in the tenant.

    Two engines, same filters:

      Delta  (default)  Walks every document library with /drives/{id}/root/delta,
                        which returns a whole library tree in pages of a thousand
                        items. Filters client-side, so "*contains*" wildcards work.

      Search (-Content) Runs the query through /search/query, the same index the
                        SharePoint search box uses, so it matches text *inside*
                        documents. KQL only does prefix wildcards (veiligheid*),
                        not leading ones.

    Permissions come from /drives/{id}/items/{id}/permissions, fetched 20 at a time
    through /$batch. That one call already carries everything that matters:

      roles              read / write / owner, reported as Read / Edit / Full Control
      grantedToV2        the user, Entra group, SharePoint group or site user
      link               a sharing link, with its scope (anyone / organization /
                         specific people), whether it can edit, and its expiry
      inheritedFrom      absent means the permission is set on the item itself -
                         that is what "unique rights" means here

    Called out separately, because these are the ones that surprise people:
    sharing links (in particular "anyone with the link"), external guests (#EXT#),
    and "Everyone except external users".

    Nothing is changed - this script only reads.

    LIMITS OF GRAPH. Graph has no API for SharePoint role assignments, so this
    script cannot tell you that "Site Members has Edit on this site" - it reports
    per item whether the permission sits on the item or is inherited, and from
    where. Graph also exposes permissions for driveItems only, so items in ordinary
    lists (not libraries) are out of scope. For the site- and list-level picture,
    and for non-library lists, use Find-SiteContent.ps1 in this folder.

    SIGN-IN. Application permissions need an app registration with admin consent.
    The first run against a tenant creates one for you:

      1. Sign in to Microsoft Graph as a Global Administrator (once)
      2. Create (or reuse) an app named after -AppName
      3. Grant and admin-consent the application role Sites.Read.All
      4. Create a self-signed certificate in CurrentUser\My and upload its public
         key to the app - no secret is ever written to disk
      5. Cache the client ID and thumbprint per tenant in graph.appid.json
         (gitignored) in the repo root

    Later runs read that cache and connect app-only without any prompt, which also
    makes the script usable from a scheduled task. Pass -ClientId together with
    -CertificateThumbprint or -ClientSecret to use an app you already have.

.PARAMETER SiteUrl
    Search one site collection, e.g. https://contoso.sharepoint.com/sites/Finance
    or a OneDrive: https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com

.PARAMETER AllSites
    Search every site in the tenant instead. OneDrive personal sites are excluded
    unless you add -IncludePersonalSites.

.PARAMETER SiteFilter
    With -AllSites: only sites whose URL matches this wildcard, e.g. "*/sites/HR*".

.PARAMETER MaxSites
    With -AllSites: stop after this many sites. Use it to try a tenant sweep out on
    a handful of sites first.

.PARAMETER IncludePersonalSites
    With -AllSites: include everybody's OneDrive in the sweep. This multiplies the
    work by the number of users - start with -MaxSites.

.PARAMETER IncludeSubsites
    Also search the subsites of -SiteUrl.

.PARAMETER Content
    Free-text/KQL query run through /search/query - the only way to match text
    inside documents. KQL supports trailing wildcards (veiligheid*) but not leading
    ones; for "*contains*" use the default delta engine with -Name.

.PARAMETER Name
    Filter on the file or folder name. Wildcards allowed, e.g. "*veiligheid*".

.PARAMETER Path
    Filter on the folder path, substring match, e.g. "Gedeelde documenten/Directie".

.PARAMETER Extension
    One or more extensions to keep, with or without the dot: "xlsx","pdf".

.PARAMETER ItemType
    Restrict to File or Folder. Default All.

.PARAMETER LibraryName
    Only these document libraries, by name. Wildcards allowed.

.PARAMETER ModifiedBy
    Filter on who last changed the item - display name or e-mail, wildcards allowed.

.PARAMETER ModifiedAfter
    Only items changed on or after this date/time.

.PARAMETER ModifiedBefore
    Only items changed before this date/time.

.PARAMETER MinSizeMB
    Only files of at least this size.

.PARAMETER Permissions
    How much permission detail to resolve per hit:
      Effective  (default) every permission that applies, inherited ones included
      Unique     only items that carry permissions of their own - the fast way to
                 find what is shared differently from the rest
      None       just find the content

.PARAMETER ExpandGroups
    Also list the members of Entra groups that hold permissions. Needs the extra
    application role Group.Read.All; without it the script says so and carries on.

.PARAMETER Everything
    Leave nothing out: subsites, personal sites, and no cap on hits or permission
    lookups. Equivalent to -IncludeSubsites -IncludePersonalSites -MaxItems 0
    -MaxPermissionLookups 0.

.PARAMETER MaxItems
    Stop after this many matching items. Default 5000; 0 means no limit.

.PARAMETER MaxPermissionLookups
    Cap on how many hits get their permissions resolved. Default 2000; 0 means no
    limit. Lookups are batched 20 per request, so this is cheaper than it sounds.

.PARAMETER TenantId
    Tenant ID or domain. Defaults to the tenant derived from -SiteUrl.

.PARAMETER ClientId
    Client ID of an existing app registration with Sites.Read.All (application).

.PARAMETER CertificateThumbprint
    Thumbprint of a certificate in CurrentUser\My to authenticate that app with.

.PARAMETER ClientSecret
    Client secret to authenticate that app with, if you would rather not use a
    certificate. Never written to disk by this script.

.PARAMETER AppName
    Display name of the app registration to create or reuse.
    Default: "M365-Scripts Graph SharePoint Search".

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\GraphSharePointFind_<timestamp>.csv
    (~/Downloads on non-Windows).

.PARAMETER MaxRetries
    How often a throttled (429) or failed Graph call is retried. Default 5.

.EXAMPLE
    # One site, everything with "veiligheid" in the name, plus who can reach it
    .\Search-SharePointContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
        -Name "*veiligheid*" -IncludeSubsites

.EXAMPLE
    # Tenant-wide: which documents mention "salarisschaal" anywhere?
    .\Search-SharePointContent.ps1 -AllSites -TenantId contoso.onmicrosoft.com `
        -Content "salarisschaal"

.EXAMPLE
    # Everything in the tenant that hangs on an "anyone with the link" link
    .\Search-SharePointContent.ps1 -AllSites -TenantId contoso.onmicrosoft.com `
        -Permissions Unique -MaxSites 25

.EXAMPLE
    # Leave nothing out on one site, including its subsites, no caps
    .\Search-SharePointContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
        -Name "*offerte*" -Everything

.EXAMPLE
    # Use an app registration you already have, from a scheduled task
    .\Search-SharePointContent.ps1 -AllSites -TenantId contoso.onmicrosoft.com `
        -ClientId 00000000-1111-2222-3333-444444444444 -CertificateThumbprint A1B2C3... `
        -Name "*.pfx" -OutputPath C:\Reports\keys.csv
#>
[CmdletBinding(DefaultParameterSetName = 'Site')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Site', Position = 0)]
    [string] $SiteUrl,

    [Parameter(Mandatory = $true, ParameterSetName = 'Tenant')]
    [switch] $AllSites,

    [Parameter(ParameterSetName = 'Tenant')]
    [string] $SiteFilter,

    [Parameter(ParameterSetName = 'Tenant')]
    [int] $MaxSites,

    [Parameter(ParameterSetName = 'Tenant')]
    [switch] $IncludePersonalSites,

    [Parameter(ParameterSetName = 'Site')]
    [switch] $IncludeSubsites,

    [string] $Content,
    [string] $Name,
    [string] $Path,
    [string[]] $Extension,

    [ValidateSet('All', 'File', 'Folder')]
    [string] $ItemType = 'All',

    [string[]] $LibraryName,
    [string] $ModifiedBy,
    [datetime] $ModifiedAfter,
    [datetime] $ModifiedBefore,
    [double] $MinSizeMB,

    [ValidateSet('Effective', 'Unique', 'None')]
    [string] $Permissions = 'Effective',

    [switch] $ExpandGroups,
    [switch] $Everything,

    [ValidateRange(0, [int]::MaxValue)]
    [int] $MaxItems = 5000,

    [ValidateRange(0, [int]::MaxValue)]
    [int] $MaxPermissionLookups = 2000,

    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [string] $ClientSecret,
    [string] $AppName = 'M365-Scripts Graph SharePoint Search',

    [string] $OutputPath,

    [ValidateRange(1, 20)]
    [int] $MaxRetries = 5
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

if ($Everything) {
    $IncludeSubsites      = $true
    $IncludePersonalSites = $true
    if (-not $PSBoundParameters.ContainsKey('MaxItems'))             { $MaxItems = 0 }
    if (-not $PSBoundParameters.ContainsKey('MaxPermissionLookups')) { $MaxPermissionLookups = 0 }
}

$tenantMode = $PSCmdlet.ParameterSetName -eq 'Tenant'

# -- Modules -------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication')) {
    throw "Module 'Microsoft.Graph.Authentication' is not installed. Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# -- Output folder -------------------------------------------------------------
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
if (-not $OutputPath) {
    $OutputPath = Join-Path $outputDir "GraphSharePointFind_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
}

# -- Tenant --------------------------------------------------------------------
if (-not $tenantMode) {
    $SiteUrl = $SiteUrl.Trim().TrimEnd('/')
    if ($SiteUrl -notmatch '^https://([a-zA-Z0-9-]+?)(-admin|-my)?\.sharepoint\.(com|de|us|cn)(/|$)') {
        throw "'$SiteUrl' does not look like a SharePoint or OneDrive URL."
    }
    if (-not $TenantId) { $TenantId = "$($Matches[1]).onmicrosoft.com" }
}
if (-not $TenantId) {
    throw 'Pass -TenantId when using -AllSites, e.g. -TenantId contoso.onmicrosoft.com'
}

$repoRoot   = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$appIdStore = Join-Path $repoRoot 'graph.appid.json'

$searchMode      = [bool]$Content
$extensionFilter = @()
if ($Extension) {
    $extensionFilter = @($Extension | ForEach-Object { $_.TrimStart('.').ToLowerInvariant() } | Where-Object { $_ })
}

$useModifiedAfter  = $PSBoundParameters.ContainsKey('ModifiedAfter')
$useModifiedBefore = $PSBoundParameters.ContainsKey('ModifiedBefore')
$useMinSize        = $PSBoundParameters.ContainsKey('MinSizeMB')

Write-Host ''
Write-Host "  Scope  : $(if ($tenantMode) { 'every site in the tenant' } else { $SiteUrl })" -ForegroundColor Cyan
Write-Host "  Tenant : $TenantId" -ForegroundColor Cyan
Write-Host "  Engine : $(if ($searchMode) { 'Graph search index' } else { 'Graph delta - every document library' })" -ForegroundColor Cyan
$filterParts = @()
if ($Name)               { $filterParts += "name $Name" }
if ($Path)               { $filterParts += "path *$Path*" }
if ($extensionFilter)    { $filterParts += "ext $($extensionFilter -join ',')" }
if ($ItemType -ne 'All') { $filterParts += "type $ItemType" }
if ($LibraryName)        { $filterParts += "library $($LibraryName -join ', ')" }
if ($ModifiedBy)         { $filterParts += "modified by $ModifiedBy" }
if ($useModifiedAfter)   { $filterParts += "after $($ModifiedAfter.ToString('yyyy-MM-dd HH:mm'))" }
if ($useModifiedBefore)  { $filterParts += "before $($ModifiedBefore.ToString('yyyy-MM-dd HH:mm'))" }
if ($useMinSize)         { $filterParts += "min $MinSizeMB MB" }
if ($filterParts.Count)  { Write-Host "  Filter : $($filterParts -join ' | ')" -ForegroundColor Cyan }
Write-Host "  Rights : $Permissions" -ForegroundColor Cyan
if ($Everything) { Write-Host '  Scope  : EVERYTHING - subsites and personal sites included, no caps' -ForegroundColor Yellow }
Write-Host '  Note   : Graph reports permissions on files and folders. Items in ordinary lists,' -ForegroundColor DarkGray
Write-Host '           and site/list level rights, need Find-SiteContent.ps1 instead.' -ForegroundColor DarkGray
Write-Host ''

function Format-Duration {
    param([TimeSpan] $Span)

    if ($Span.TotalSeconds -lt 60) { return "{0:n0}s" -f $Span.TotalSeconds }
    if ($Span.TotalMinutes -lt 60) { return "{0}m {1}s" -f [int]$Span.TotalMinutes, $Span.Seconds }
    return "{0}h {1}m" -f [int]$Span.TotalHours, $Span.Minutes
}

# -- App registration ----------------------------------------------------------
function Get-CachedApp {
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

function Set-CachedApp {
    param([string] $Tenant, [string] $Id, [string] $Thumbprint)

    $store = @{}
    if (Test-Path $appIdStore) {
        try {
            (Get-Content $appIdStore -Raw | ConvertFrom-Json).PSObject.Properties |
                ForEach-Object { $store[$_.Name] = $_.Value }
        } catch { }
    }
    $store[$Tenant] = [pscustomobject]@{ ClientId = $Id; CertificateThumbprint = $Thumbprint }
    $store | ConvertTo-Json -Depth 5 | Set-Content -Path $appIdStore -Encoding UTF8
    Write-Host "  App details cached in $appIdStore" -ForegroundColor DarkGray
}

function New-GraphSearchApp {
    <#
        Creates (or reuses) an app registration with the application role
        Sites.Read.All, consents it, and gives it a self-signed certificate so no
        secret has to be stored anywhere. Returns client id + thumbprint.
    #>
    param([string] $Tenant, [string] $DisplayName, [switch] $WithGroups)

    foreach ($module in @('Microsoft.Graph.Applications')) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            throw "Module '$module' is required to create the app registration. Run: Install-Module $module -Scope CurrentUser"
        }
    }
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop

    if (-not ($IsWindows -or $env:OS -eq 'Windows_NT')) {
        throw 'Automatic app registration uses a self-signed certificate and only runs on Windows. Create the app yourself and pass -ClientId with -CertificateThumbprint or -ClientSecret.'
    }

    Write-Host '  No app registration known for this tenant - creating one.' -ForegroundColor Yellow
    Write-Host "  Sign in as a Global Administrator of $Tenant." -ForegroundColor Yellow

    Connect-MgGraph -TenantId $Tenant -NoWelcome -ContextScope Process -Scopes @(
        'Application.ReadWrite.All'
        'AppRoleAssignment.ReadWrite.All'
        'Directory.Read.All'
    ) | Out-Null
    Write-Host "  Signed in as $((Get-MgContext).Account)" -ForegroundColor Green

    $escaped = $DisplayName -replace "'", "''"
    $app = @(Get-MgApplication -Filter "displayName eq '$escaped'" -All) | Select-Object -First 1
    if ($app) {
        Write-Host "  Reusing existing app '$DisplayName' ($($app.AppId))" -ForegroundColor Green
    } else {
        $app = New-MgApplication -DisplayName $DisplayName -SignInAudience 'AzureADMyOrg'
        Write-Host "  App created: $($app.AppId)" -ForegroundColor Green
    }

    $sp = @(Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -All) | Select-Object -First 1
    if (-not $sp) {
        $sp = New-MgServicePrincipal -AppId $app.AppId
        Write-Host '  Service principal created.' -ForegroundColor Green
    }

    # Resolve the app role ids from the Graph service principal itself rather than
    # hardcoding GUIDs that may differ per cloud.
    $graphSp = @(Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -All) | Select-Object -First 1
    if (-not $graphSp) { throw 'Could not find the Microsoft Graph service principal in this tenant.' }

    $wanted = @('Sites.Read.All')
    if ($WithGroups) { $wanted += 'Group.Read.All' }

    $existingGrants = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All)
    foreach ($roleName in $wanted) {
        $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $roleName -and $_.AllowedMemberTypes -contains 'Application' } | Select-Object -First 1
        if (-not $role) {
            Write-Warning "  Application role $roleName not found on Microsoft Graph - skipped."
            continue
        }
        if ($existingGrants | Where-Object { $_.AppRoleId -eq $role.Id -and $_.ResourceId -eq $graphSp.Id }) {
            Write-Host "    $roleName : already consented." -ForegroundColor DarkGray
            continue
        }
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id `
            -PrincipalId $sp.Id -ResourceId $graphSp.Id -AppRoleId $role.Id | Out-Null
        Write-Host "    $roleName : consented." -ForegroundColor Green
    }

    # Certificate instead of a secret - nothing sensitive ends up in the cache file.
    $subject = "CN=$DisplayName"
    $cert = Get-ChildItem Cert:\CurrentUser\My |
        Where-Object { $_.Subject -eq $subject -and $_.NotAfter -gt (Get-Date).AddDays(30) } |
        Sort-Object NotAfter -Descending | Select-Object -First 1

    if ($cert) {
        Write-Host "  Reusing certificate $($cert.Thumbprint)" -ForegroundColor DarkGray
    } else {
        $cert = New-SelfSignedCertificate -Subject $subject -CertStoreLocation 'Cert:\CurrentUser\My' `
            -KeyExportPolicy NonExportable -KeySpec Signature -KeyAlgorithm RSA -KeyLength 2048 `
            -NotAfter (Get-Date).AddYears(2)
        Write-Host "  Certificate created: $($cert.Thumbprint) (valid until $($cert.NotAfter.ToString('yyyy-MM-dd')))" -ForegroundColor Green
    }

    $app = Get-MgApplication -ApplicationId $app.Id
    $known = @($app.KeyCredentials | Where-Object { $_.CustomKeyIdentifier -and ([System.Convert]::ToBase64String($_.CustomKeyIdentifier)) -eq ([System.Convert]::ToBase64String($cert.GetCertHash())) })
    if ($known.Count -eq 0) {
        # Update-MgApplication replaces the whole collection, so keep what is there.
        $keys = @($app.KeyCredentials | Where-Object { $_.Key } | ForEach-Object {
            @{ Type = $_.Type; Usage = $_.Usage; Key = $_.Key; DisplayName = $_.DisplayName }
        })
        $keys += @{ Type = 'AsymmetricX509Cert'; Usage = 'Verify'; Key = $cert.RawData; DisplayName = $subject }
        Update-MgApplication -ApplicationId $app.Id -KeyCredentials $keys | Out-Null
        Write-Host '  Certificate uploaded to the app registration.' -ForegroundColor Green
    }

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return [pscustomobject]@{ ClientId = $app.AppId; CertificateThumbprint = $cert.Thumbprint; IsNew = $true }
}

$appIsNew = $false
if (-not $ClientId) {
    $cached = Get-CachedApp -Tenant $TenantId
    if ($cached -and $cached.ClientId) {
        $ClientId = $cached.ClientId
        if (-not $CertificateThumbprint -and -not $ClientSecret) { $CertificateThumbprint = $cached.CertificateThumbprint }
        Write-Host "  Using cached app registration $ClientId" -ForegroundColor DarkGray
    }
}

if (-not $ClientId) {
    $created = New-GraphSearchApp -Tenant $TenantId -DisplayName $AppName -WithGroups:$ExpandGroups
    $ClientId              = $created.ClientId
    $CertificateThumbprint = $created.CertificateThumbprint
    Set-CachedApp -Tenant $TenantId -Id $ClientId -Thumbprint $CertificateThumbprint
    $appIsNew = $true
}

if (-not $CertificateThumbprint -and -not $ClientSecret) {
    throw "No credential for app $ClientId - pass -CertificateThumbprint or -ClientSecret."
}

# -- Connect app-only ----------------------------------------------------------
if ($appIsNew) {
    Write-Host '  Waiting 30s for the new app registration and its consent to propagate...' -ForegroundColor DarkGray
    Start-Sleep -Seconds 30
}

$connected = $false
for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
    try {
        if ($ClientSecret) {
            $secure = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
            $credential = [pscredential]::new($ClientId, $secure)
            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome -ContextScope Process | Out-Null
        } else {
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ContextScope Process | Out-Null
        }
        $connected = $true
        break
    } catch {
        if ($attempt -eq $MaxRetries) { throw }
        Write-Host "  Sign-in attempt $attempt failed ($($_.Exception.Message.Trim())) - retrying in 10s..." -ForegroundColor DarkYellow
        Start-Sleep -Seconds 10
    }
}
if (-not $connected) { throw 'Could not connect to Microsoft Graph.' }
Write-Host "  Connected app-only as $ClientId" -ForegroundColor Green

# App-only means the token carries roles, not user rights: with Sites.Read.All it
# reads every site in the tenant regardless of who is a member. Say so, and warn
# when the role is missing instead of letting sites come back mysteriously empty.
$grantedRoles = @()
try { $grantedRoles = @((Get-MgContext).Scopes) } catch { }
if ($grantedRoles.Count -gt 0) {
    Write-Host "  Roles  : $($grantedRoles -join ', ')" -ForegroundColor DarkGray
    $readAll = @('Sites.Read.All', 'Sites.FullControl.All', 'Sites.Manage.All', 'Sites.ReadWrite.All', 'Files.Read.All') |
        Where-Object { $grantedRoles -contains $_ }
    if ($readAll) {
        Write-Host "  Access : tenant-wide through $($readAll[0]) - site membership does not apply to an app role" -ForegroundColor Green
    } else {
        Write-Warning 'The token carries no tenant-wide read role (Sites.Read.All). Sites that were not granted to this app individually (Sites.Selected) will come back empty rather than refused.'
    }
    if ($ExpandGroups -and $grantedRoles -notcontains 'Group.Read.All') {
        Write-Warning '-ExpandGroups needs the application role Group.Read.All; group members will be reported as unavailable.'
    }
}
Write-Host ''

# Anything in scope the run could not read - the summary reports this, so "nothing
# found" can be told apart from "could not look".
$script:Unreadable = [System.Collections.Generic.List[object]]::new()

# -- Graph plumbing ------------------------------------------------------------
$script:GraphRoot = 'https://graph.microsoft.com/v1.0'

function Invoke-Graph {
    <#
        One Graph call with throttling handled: 429 and 503 are retried, honouring
        Retry-After when Graph sends one.
    #>
    param(
        [string] $Uri,
        [string] $Method = 'GET',
        $Body,
        [switch] $Quiet
    )

    if ($Uri -notmatch '^https://') { $Uri = "$script:GraphRoot$Uri" }

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            if ($Body) {
                return Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body ($Body | ConvertTo-Json -Depth 10) `
                    -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop
            }
            return Invoke-MgGraphRequest -Method $Method -Uri $Uri -OutputType PSObject -ErrorAction Stop
        } catch {
            $status = 0
            try { $status = [int]$_.Exception.Response.StatusCode } catch { }
            $message = "$($_.Exception.Message)"
            if ($status -eq 0 -and $message -match '429|TooManyRequests|throttl|ServiceUnavailable') { $status = 429 }

            if ($status -in @(429, 503, 504) -and $attempt -lt $MaxRetries) {
                $wait = 10 * $attempt
                try {
                    $retryAfter = $_.Exception.Response.Headers.RetryAfter.Delta.TotalSeconds
                    if ($retryAfter) { $wait = [int]$retryAfter }
                } catch { }
                Write-Host "    Throttled by Graph - waiting $wait s..." -ForegroundColor DarkYellow
                Start-Sleep -Seconds $wait
                continue
            }
            if (-not $Quiet) { throw }
            return $null
        }
    }
    return $null
}

function Get-GraphAll {
    <#
        Follows @odata.nextLink and returns every value. -Activity turns on a
        progress line, because a big library takes a while.
    #>
    param([string] $Uri, [string] $Activity, [int] $ProgressId = 3)

    $items = [System.Collections.Generic.List[object]]::new()
    $next  = $Uri
    $page  = 0

    while ($next) {
        $page++
        $response = Invoke-Graph -Uri $next
        if (-not $response) { break }
        if ($response.value) { $items.AddRange(@($response.value)) }

        if ($Activity) {
            Write-Progress -Id $ProgressId -Activity $Activity -Status "page $page - $($items.Count) item(s)"
        }
        $next = $response.'@odata.nextLink'
    }
    if ($Activity) { Write-Progress -Id $ProgressId -Activity $Activity -Completed }
    return $items
}

function Invoke-GraphBatch {
    <#
        Sends up to 20 GETs per round trip. Returns a hashtable id -> response body.
        Sub-requests that come back throttled are retried in a smaller batch.
    #>
    param([hashtable] $Requests)   # id -> relative url

    $results = @{}
    $ids = @($Requests.Keys)

    for ($offset = 0; $offset -lt $ids.Count; $offset += 20) {
        $chunk = $ids[$offset..([math]::Min($offset + 19, $ids.Count - 1))]
        $payload = @{
            requests = @($chunk | ForEach-Object { @{ id = "$_"; method = 'GET'; url = $Requests[$_] } })
        }

        $response = Invoke-Graph -Uri '/$batch' -Method POST -Body $payload
        if (-not $response) { continue }

        $retry = @{}
        foreach ($sub in @($response.responses)) {
            if ($sub.status -eq 200) {
                $results[$sub.id] = $sub.body
            } elseif ($sub.status -in @(429, 503, 504)) {
                $retry[$sub.id] = $Requests[$sub.id]
            } else {
                $results[$sub.id] = $null
            }
        }

        if ($retry.Count -gt 0) {
            Write-Host "    $($retry.Count) permission call(s) throttled - waiting 15s and retrying..." -ForegroundColor DarkYellow
            Start-Sleep -Seconds 15
            foreach ($entry in (Invoke-GraphBatch -Requests $retry).GetEnumerator()) { $results[$entry.Key] = $entry.Value }
        }
    }

    return $results
}

# -- Sites ---------------------------------------------------------------------
function Get-SiteByUrl {
    param([string] $Url)

    $uri = [uri]$Url
    $relative = $uri.AbsolutePath.TrimEnd('/')
    $lookup = if ($relative -and $relative -ne '/') { "/sites/$($uri.Host):$($relative)" } else { "/sites/$($uri.Host)" }
    return Invoke-Graph -Uri $lookup
}

$sites = [System.Collections.Generic.List[object]]::new()
$timer = [System.Diagnostics.Stopwatch]::StartNew()

if ($tenantMode) {
    Write-Host '  Enumerating sites...' -ForegroundColor DarkGray
    $allSites = Get-GraphAll -Uri '/sites/getAllSites' -Activity 'Enumerating sites'

    $skippedPersonal = 0
    $skippedFiltered = 0
    foreach ($site in $allSites) {
        $url = "$($site.webUrl)".TrimEnd('/')
        if (-not $url) { continue }
        $personal = $site.isPersonalSite -eq $true -or $url -match '-my\.sharepoint\.'
        if ($personal -and -not $IncludePersonalSites) { $skippedPersonal++; continue }
        if ($SiteFilter -and ($url -notlike $SiteFilter)) { $skippedFiltered++; continue }
        $sites.Add([pscustomobject]@{ Id = $site.id; Url = $url; Title = $site.displayName })
    }

    Write-Host "  Sites found      : $($allSites.Count) in $(Format-Duration $timer.Elapsed)" -ForegroundColor Cyan
    Write-Host "  Skipped          : $skippedPersonal personal, $skippedFiltered filtered out" -ForegroundColor DarkGray

    if ($MaxSites -and $sites.Count -gt $MaxSites) {
        Write-Host "  Limited to the first $MaxSites site(s) by -MaxSites (of $($sites.Count) matching)." -ForegroundColor Yellow
        $sites = [System.Collections.Generic.List[object]]@($sites | Select-Object -First $MaxSites)
    }
    Write-Host "  Sites to search  : $($sites.Count)" -ForegroundColor Cyan
} else {
    $site = Get-SiteByUrl -Url $SiteUrl
    if (-not $site -or -not $site.id) { throw "Site not found through Graph: $SiteUrl" }
    $sites.Add([pscustomobject]@{ Id = $site.id; Url = "$($site.webUrl)".TrimEnd('/'); Title = $site.displayName })

    if ($IncludeSubsites) {
        $subs = Get-GraphAll -Uri "/sites/$($site.id)/sites?`$select=id,displayName,webUrl"
        foreach ($sub in $subs) {
            $sites.Add([pscustomobject]@{ Id = $sub.id; Url = "$($sub.webUrl)".TrimEnd('/'); Title = $sub.displayName })
        }
        Write-Host "  Sites to search  : $($sites.Count) (site + $($sites.Count - 1) subsite(s))" -ForegroundColor Cyan
    }
}
Write-Host ''

if ($sites.Count -eq 0) {
    Write-Host '  No sites to search. Nothing to do.' -ForegroundColor Yellow
    return
}

# -- Matching ------------------------------------------------------------------
function Test-Match {
    param($Hit)

    if ($Name) {
        $candidates = @($Hit.Name, [System.IO.Path]::GetFileName("$($Hit.Url)")) | Where-Object { $_ }
        if (-not @($candidates | Where-Object { $_ -like $Name })) { return $false }
    }
    if ($Path -and ("$($Hit.FolderPath)" -notlike "*$Path*")) { return $false }
    if ($ItemType -ne 'All' -and $Hit.ItemType -ne $ItemType) { return $false }

    if ($extensionFilter.Count -gt 0) {
        $ext = [System.IO.Path]::GetExtension("$($Hit.Name)").TrimStart('.').ToLowerInvariant()
        if ($ext -notin $extensionFilter) { return $false }
    }
    if ($ModifiedBy -and "$($Hit.ModifiedBy)" -notlike $ModifiedBy -and "$($Hit.ModifiedByEmail)" -notlike $ModifiedBy) { return $false }

    if ($useModifiedAfter  -and $Hit.Modified -and $Hit.Modified -lt $ModifiedAfter)  { return $false }
    if ($useModifiedBefore -and $Hit.Modified -and $Hit.Modified -ge $ModifiedBefore) { return $false }
    if ($useMinSize -and (($Hit.SizeBytes / 1MB) -lt $MinSizeMB)) { return $false }

    return $true
}

function ConvertTo-Hit {
    <#
        driveItem (from delta or from search) -> the flat shape the rest works with.
    #>
    param($Item, [string] $SiteUrlValue, [string] $Library, [string] $DriveId)

    $folderPath = ''
    if ($Item.parentReference -and $Item.parentReference.path) {
        # "/drive/root:/Gedeelde documenten/Map" -> "Gedeelde documenten/Map"
        $folderPath = "$($Item.parentReference.path)" -replace '^/drive[^:]*:?/?', ''
        $folderPath = [uri]::UnescapeDataString($folderPath)
    }

    $modified = $null
    if ($Item.lastModifiedDateTime) { try { $modified = [datetime]$Item.lastModifiedDateTime } catch { } }
    $created = $null
    if ($Item.createdDateTime) { try { $created = [datetime]$Item.createdDateTime } catch { } }

    [pscustomobject]@{
        Site            = $SiteUrlValue
        Library         = $Library
        DriveId         = if ($DriveId) { $DriveId } else { "$($Item.parentReference.driveId)" }
        ItemId          = "$($Item.id)"
        ItemType        = if ($Item.folder) { 'Folder' } else { 'File' }
        Name            = "$($Item.name)"
        Url             = "$($Item.webUrl)"
        FolderPath      = $folderPath
        SizeBytes       = if ($Item.size) { [long]$Item.size } else { 0 }
        Modified        = $modified
        ModifiedBy      = "$($Item.lastModifiedBy.user.displayName)"
        ModifiedByEmail = "$($Item.lastModifiedBy.user.email)"
        Created         = $created
        CreatedBy       = "$($Item.createdBy.user.displayName)"
        Shared          = [bool]$Item.shared
    }
}

# -- Collect hits --------------------------------------------------------------
$hits    = [System.Collections.Generic.List[object]]::new()
$scanned = 0
$capped  = $false

$selectFields = 'id,name,size,webUrl,lastModifiedDateTime,createdDateTime,lastModifiedBy,createdBy,file,folder,parentReference,shared,deleted'

if ($searchMode) {
    Write-Host "  Query  : $Content" -ForegroundColor DarkGray

    $scopes = @($sites | ForEach-Object { "path:`"$($_.Url)`"" })
    $scopeClause = if ($scopes.Count -eq 1) { $scopes[0] } elseif ($scopes.Count -le 20) { "($($scopes -join ' OR '))" } else { '' }
    if (-not $scopeClause -and -not $tenantMode) { $scopeClause = "path:`"$SiteUrl`"" }
    $kql = if ($scopeClause) { "$Content $scopeClause" } else { $Content }
    if (-not $scopeClause) { Write-Host '  Note   : too many sites to scope the query - searching the whole tenant and filtering afterwards.' -ForegroundColor DarkGray }

    $siteUrls = @($sites.Url)
    $from = 0
    do {
        $payload = @{
            requests = @(@{
                entityTypes = @('driveItem')
                query       = @{ queryString = $kql }
                from        = $from
                size        = 200
            })
        }
        $response = Invoke-Graph -Uri '/search/query' -Method POST -Body $payload
        $container = @($response.value)[0]
        $hitsContainer = @($container.hitsContainers)[0]
        $rows = @($hitsContainer.hits)

        foreach ($row in $rows) {
            $scanned++
            $resource = $row.resource
            if (-not $resource) { continue }

            $hit = ConvertTo-Hit -Item $resource -SiteUrlValue '' -Library '' -DriveId ''
            # The search index is tenant-wide; keep only what falls inside the scope.
            $inScope = $siteUrls | Where-Object { $hit.Url -like "$_*" } | Select-Object -First 1
            if (-not $inScope) { continue }
            $hit.Site = $inScope

            if (-not (Test-Match -Hit $hit)) { continue }
            if ($MaxItems -gt 0 -and $hits.Count -ge $MaxItems) { $capped = $true; break }
            $hits.Add($hit)
        }

        $from += 200
        $more = $hitsContainer.moreResultsAvailable -eq $true
    } while ($more -and -not $capped -and $from -lt 1000)

    Write-Host "  Search returned $scanned result(s) in $(Format-Duration $timer.Elapsed); $($hits.Count) in scope and matching." -ForegroundColor Cyan
} else {
    $siteIndex = 0
    foreach ($site in $sites) {
        $siteIndex++
        if ($capped) { break }

        Write-Progress -Id 1 -Activity 'Walking sites' `
            -Status "$siteIndex/$($sites.Count) - $($site.Url) - $($hits.Count) hit(s), $(Format-Duration $timer.Elapsed)" `
            -PercentComplete ([math]::Min(100, ($siteIndex / $sites.Count) * 100))

        $drives = @()
        try {
            $drives = @(Get-GraphAll -Uri "/sites/$($site.Id)/drives?`$select=id,name,webUrl,driveType")
        } catch {
            Write-Warning "Could not list libraries on $($site.Url): $($_.Exception.Message)"
            $script:Unreadable.Add([pscustomobject]@{ Scope = $site.Url; Reason = $_.Exception.Message.Trim() })
            continue
        }

        $targetDrives = @($drives | Where-Object {
            if (-not $LibraryName) { return $true }
            $driveName = $_.name
            return [bool]@($LibraryName | Where-Object { $driveName -like $_ })
        })

        Write-Host "  [$siteIndex/$($sites.Count)] $($site.Url) - $($targetDrives.Count) of $($drives.Count) librar(ies)" -ForegroundColor DarkGray

        foreach ($drive in $targetDrives) {
            if ($capped) { break }

            $items = @()
            try {
                $items = @(Get-GraphAll -Uri "/drives/$($drive.id)/root/delta?`$select=$selectFields&`$top=999" `
                    -Activity "$($drive.name) - $($site.Url)")
            } catch {
                Write-Warning "Could not read library '$($drive.name)' on $($site.Url): $($_.Exception.Message)"
                $script:Unreadable.Add([pscustomobject]@{ Scope = "$($site.Url) > $($drive.name)"; Reason = $_.Exception.Message.Trim() })
                continue
            }

            foreach ($item in $items) {
                if ($item.deleted) { continue }
                if (-not $item.parentReference -or -not $item.parentReference.path) { continue }  # the root itself
                $scanned++

                $hit = ConvertTo-Hit -Item $item -SiteUrlValue $site.Url -Library $drive.name -DriveId $drive.id
                if (-not (Test-Match -Hit $hit)) { continue }
                if ($MaxItems -gt 0 -and $hits.Count -ge $MaxItems) { $capped = $true; break }
                $hits.Add($hit)
            }
        }
    }
    Write-Progress -Id 1 -Activity 'Walking sites' -Completed

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

# -- Permissions ---------------------------------------------------------------
$script:GroupCache = @{}

function Get-GroupMemberSummary {
    param([string] $GroupId)

    if (-not $GroupId) { return '' }
    if ($script:GroupCache.ContainsKey($GroupId)) { return $script:GroupCache[$GroupId] }

    $summary = ''
    $response = Invoke-Graph -Uri "/groups/$GroupId/members?`$select=displayName,userPrincipalName&`$top=100" -Quiet
    if ($response -and $response.value) {
        $summary = (@($response.value) | ForEach-Object {
            if ($_.userPrincipalName) { $_.userPrincipalName } else { $_.displayName }
        }) -join '; '
    } elseif (-not $response) {
        $summary = '<no Group.Read.All - members not read>'
    }
    $script:GroupCache[$GroupId] = $summary
    return $summary
}

function ConvertFrom-GraphPermission {
    <#
        One Graph permission object -> one row per principal it grants to.
        A sharing link grants to everyone in grantedToIdentitiesV2, or to the whole
        world / organisation when its scope says so.
    #>
    param($Permission)

    $rows = [System.Collections.Generic.List[object]]::new()

    $roles = @($Permission.roles | ForEach-Object {
        switch ("$_") {
            'read'  { 'Read' }
            'write' { 'Edit' }
            'owner' { 'Full Control' }
            default { "$_" }
        }
    })
    $roleText = ($roles -join ', ')

    $inheritedFrom = ''
    if ($Permission.inheritedFrom) {
        $inheritedFrom = "$($Permission.inheritedFrom.path)" -replace '^/drive[^:]*:?/?', ''
        if (-not $inheritedFrom) { $inheritedFrom = 'parent folder' }
        $inheritedFrom = [uri]::UnescapeDataString($inheritedFrom)
    }

    $linkLabel = ''
    $linkUrl   = ''
    $expires   = if ($Permission.expirationDateTime) { "$($Permission.expirationDateTime)" } else { '' }

    if ($Permission.link) {
        $scope = "$($Permission.link.scope)"
        $type  = "$($Permission.link.type)"
        $linkLabel = switch ($scope) {
            'anonymous'    { "Anyone - $type" }
            'organization' { "Organization - $type" }
            'users'        { "Specific people - $type" }
            default        { "Sharing link - $type" }
        }
        $linkUrl = "$($Permission.link.webUrl)"
    }

    # Who it is granted to. grantedToIdentitiesV2 is the list behind a link;
    # grantedToV2 is a direct assignment.
    $identities = @()
    if ($Permission.grantedToIdentitiesV2) { $identities += @($Permission.grantedToIdentitiesV2) }
    if ($Permission.grantedToV2)           { $identities += @($Permission.grantedToV2) }

    if ($identities.Count -eq 0) {
        # An anonymous or organisation-wide link grants to nobody in particular.
        $rows.Add([pscustomobject]@{
            Principal   = if ($linkLabel) { $linkLabel } else { '(unspecified)' }
            Type        = if ($Permission.link.scope -eq 'anonymous') { 'AnyoneWithLink' } else { 'Link' }
            Login       = ''
            Email       = ''
            Permission  = $roleText
            SharingLink = $linkLabel
            LinkUrl     = $linkUrl
            Expires     = $expires
            Inherited   = $inheritedFrom
            IsExternal  = $Permission.link.scope -eq 'anonymous'
            IsEveryone  = $Permission.link.scope -in @('anonymous', 'organization')
            Members     = ''
        })
        return $rows
    }

    foreach ($identity in $identities) {
        $display = ''
        $login   = ''
        $email   = ''
        $type    = ''
        $members = ''
        $everyone = $false

        if ($identity.user) {
            $type    = 'User'
            $display = "$($identity.user.displayName)"
            $email   = "$($identity.user.email)"
            $login   = if ($identity.user.userPrincipalName) { "$($identity.user.userPrincipalName)" } else { $email }
        } elseif ($identity.siteUser) {
            $type    = 'SiteUser'
            $display = "$($identity.siteUser.displayName)"
            $login   = "$($identity.siteUser.loginName)"
            $email   = "$($identity.siteUser.email)"
        } elseif ($identity.group) {
            $type    = 'Group'
            $display = "$($identity.group.displayName)"
            $email   = "$($identity.group.email)"
            if ($ExpandGroups -and $identity.group.id) { $members = Get-GroupMemberSummary -GroupId "$($identity.group.id)" }
        } elseif ($identity.siteGroup) {
            $type    = 'SharePointGroup'
            $display = "$($identity.siteGroup.displayName)"
            $login   = "$($identity.siteGroup.loginName)"
        } elseif ($identity.application) {
            $type    = 'Application'
            $display = "$($identity.application.displayName)"
        } else {
            $type    = 'Unknown'
            $display = '(unknown principal)'
        }

        $everyone = $display -in @('Everyone', 'Everyone except external users',
                                   'Iedereen', 'Iedereen behalve externe gebruikers') -or
                    $login -match 'spo-grid-all-users'
        $external = "$login$email" -match '#EXT#|#ext#' -or $display -match '\(Guest\)'

        $rows.Add([pscustomobject]@{
            Principal   = if ($display) { $display } else { $login }
            Type        = $type
            Login       = $login
            Email       = $email
            Permission  = $roleText
            SharingLink = $linkLabel
            LinkUrl     = $linkUrl
            Expires     = $expires
            Inherited   = $inheritedFrom
            IsExternal  = [bool]$external
            IsEveryone  = [bool]$everyone
            Members     = $members
        })
    }

    return $rows
}

$results   = [System.Collections.Generic.List[object]]::new()
$permTimer = [System.Diagnostics.Stopwatch]::StartNew()
$resolved  = 0
$permFailed = 0

function Add-Row {
    param($Hit, $Principal, [string] $Source)

    $results.Add([pscustomobject]@{
        Site             = $Hit.Site
        Library          = $Hit.Library
        ItemType         = $Hit.ItemType
        Name             = $Hit.Name
        Url              = $Hit.Url
        FolderPath       = $Hit.FolderPath
        SizeMB           = if ($Hit.SizeBytes) { [math]::Round($Hit.SizeBytes / 1MB, 2) } else { 0 }
        Modified         = $Hit.Modified
        ModifiedBy       = $Hit.ModifiedBy
        Created          = $Hit.Created
        CreatedBy        = $Hit.CreatedBy
        ItemId           = $Hit.ItemId
        PermissionSource = $Source
        InheritedFrom    = if ($Principal) { $Principal.Inherited } else { '' }
        UniqueRights     = [bool]($Source -eq 'Item')
        Principal        = if ($Principal) { $Principal.Principal } else { '' }
        PrincipalType    = if ($Principal) { $Principal.Type } else { '' }
        PrincipalLogin   = if ($Principal) { $Principal.Login } else { '' }
        PrincipalEmail   = if ($Principal) { $Principal.Email } else { '' }
        Permission       = if ($Principal) { $Principal.Permission } else { '' }
        SharingLink      = if ($Principal) { $Principal.SharingLink } else { '' }
        LinkUrl          = if ($Principal) { $Principal.LinkUrl } else { '' }
        LinkExpires      = if ($Principal) { $Principal.Expires } else { '' }
        External         = if ($Principal) { $Principal.IsExternal } else { $false }
        Everyone         = if ($Principal) { $Principal.IsEveryone } else { $false }
        GroupMembers     = if ($Principal) { $Principal.Members } else { '' }
    })
}

if ($Permissions -eq 'None') {
    foreach ($hit in $hits) { Add-Row -Hit $hit -Principal $null -Source '' }
} else {
    # Batched 20 at a time; the hit list is indexed so each response finds its item.
    # ToArray rather than @($hits): an array subexpression over a List[object] throws
    # "Argument types do not match" on PowerShell 7.6.5 / .NET 10.
    $lookupHits = if ($MaxPermissionLookups -gt 0) { @($hits | Select-Object -First $MaxPermissionLookups) } else { $hits.ToArray() }
    $overflow   = @($hits | Select-Object -Skip $lookupHits.Count)

    for ($offset = 0; $offset -lt $lookupHits.Count; $offset += 20) {
        $slice = $lookupHits[$offset..([math]::Min($offset + 19, $lookupHits.Count - 1))]

        $requests = @{}
        for ($i = 0; $i -lt $slice.Count; $i++) {
            $hit = $slice[$i]
            if (-not $hit.DriveId -or -not $hit.ItemId) { continue }
            $requests["$i"] = "/drives/$($hit.DriveId)/items/$($hit.ItemId)/permissions"
        }
        if ($requests.Count -eq 0) { continue }

        $responses = Invoke-GraphBatch -Requests $requests

        for ($i = 0; $i -lt $slice.Count; $i++) {
            $hit = $slice[$i]
            $resolved++
            Write-Progress -Id 2 -Activity 'Resolving permissions' `
                -Status "$resolved/$($lookupHits.Count) - $($hit.Name) - $(Format-Duration $permTimer.Elapsed)" `
                -PercentComplete ([math]::Min(100, ($resolved / $lookupHits.Count) * 100))

            $body = $responses["$i"]
            if (-not $body) {
                $permFailed++
                Add-Row -Hit $hit -Principal $null -Source 'unreadable'
                continue
            }

            $permissions = @($body.value)
            $rowsForItem = [System.Collections.Generic.List[object]]::new()
            foreach ($permission in $permissions) {
                $direct = -not $permission.inheritedFrom
                if ($Permissions -eq 'Unique' -and -not $direct) { continue }
                foreach ($principal in (ConvertFrom-GraphPermission -Permission $permission)) {
                    $rowsForItem.Add([pscustomobject]@{ Principal = $principal; Source = if ($direct) { 'Item' } else { 'Inherited' } })
                }
            }

            if ($rowsForItem.Count -eq 0) {
                if ($Permissions -eq 'Unique') { continue }   # inherits everything - not interesting
                Add-Row -Hit $hit -Principal $null -Source 'Inherited'
                continue
            }
            foreach ($row in $rowsForItem) { Add-Row -Hit $hit -Principal $row.Principal -Source $row.Source }
        }
    }
    Write-Progress -Id 2 -Activity 'Resolving permissions' -Completed

    foreach ($hit in $overflow) { Add-Row -Hit $hit -Principal $null -Source '' }
}
$permTimer.Stop()
$timer.Stop()

# -- Report --------------------------------------------------------------------
if ($results.Count -gt 0) {
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
}

$permissionRows = @($results | Where-Object { $_.Principal })
$uniqueItems    = @($results | Where-Object { $_.UniqueRights } | Group-Object Url).Count
$linkRows       = @($results | Where-Object { $_.SharingLink })
$anyoneRows     = @($results | Where-Object { $_.SharingLink -like 'Anyone*' })
$externalRows   = @($results | Where-Object { $_.External })
$everyoneRows   = @($results | Where-Object { $_.Everyone })

Write-Host ''
if ($Permissions -eq 'Unique' -and $permissionRows.Count -eq 0) {
    Write-Host "  $($hits.Count) item(s) matched and every one of them inherits - nothing is shared differently." -ForegroundColor Green
} else {
    $hits | Group-Object Site | Sort-Object Count -Descending | Select-Object -First 15 |
        ForEach-Object {
            [pscustomobject]@{
                Site = $_.Name -replace '^https://[^/]+', ''
                Hits = $_.Count
            }
        } | Format-Table -AutoSize | Out-Host

    $preview = @($permissionRows | Select-Object -First 25)
    if ($preview.Count -gt 0) {
        $preview | Select-Object `
            @{ N = 'Item';   E = { if ("$($_.Name)".Length -gt 38) { "$($_.Name)".Substring(0, 37) + '...' } else { $_.Name } } },
            @{ N = 'Src';    E = { $_.PermissionSource } },
            @{ N = 'Who';    E = { if ("$($_.Principal)".Length -gt 32) { "$($_.Principal)".Substring(0, 31) + '...' } else { $_.Principal } } },
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
    if ($anyoneRows.Count)   { Write-Host "  Anyone links     : $($anyoneRows.Count) - reachable without signing in" -ForegroundColor Red }
    if ($externalRows.Count) { Write-Host "  External access  : $($externalRows.Count) assignment(s)" -ForegroundColor Red }
    if ($everyoneRows.Count) { Write-Host "  Everyone (-ish)  : $($everyoneRows.Count) assignment(s)" -ForegroundColor Yellow }
    if ($permFailed)         { Write-Host "  Unreadable       : $permFailed item(s) - permissions could not be read" -ForegroundColor DarkYellow }
    if ($MaxPermissionLookups -gt 0 -and $hits.Count -gt $MaxPermissionLookups) {
        Write-Host "  Note             : rights resolved for the first $MaxPermissionLookups item(s) (-MaxPermissionLookups)" -ForegroundColor Yellow
    }
}
if ($script:Unreadable.Count -gt 0) {
    Write-Host "  Not readable     : $($script:Unreadable.Count) - the result is INCOMPLETE" -ForegroundColor Red
    $script:Unreadable | Select-Object -First 10 |
        Select-Object @{ N = 'Skipped'; E = { $_.Scope } }, @{ N = 'Why'; E = { $_.Reason } } |
        Format-Table -AutoSize | Out-Host
    if ($script:Unreadable.Count -gt 10) { Write-Host "  ... and $($script:Unreadable.Count - 10) more." -ForegroundColor DarkGray }
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

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
