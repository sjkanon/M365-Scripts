#Requires -Version 7.0
<#
.SYNOPSIS
    Restore deleted files and folders from SharePoint or OneDrive recycle bins -
    one site, or every SharePoint site in a tenant - with interactive admin sign-in.

.DESCRIPTION
    Reads the first- and/or second-stage recycle bin, filters the items, and
    restores the matches to their original location.

    Two modes:
      -SiteUrl <url>  one site collection: team site, communication site, or a
                      user's OneDrive
      -AllSites       every SharePoint site in the tenant. OneDrive personal sites
                      are deliberately excluded from this sweep - restore those
                      one at a time with -SiteUrl.

    Defaults to dry-run mode - pass -Apply to actually restore.

    Items are restored in batches (see -BatchSize): SharePoint restores a whole
    set in a single server call, so a few hundred files take about as long as one.
    Folders and files never share a batch, and a batch that fails as a whole is
    retried item by item so the good items still come back and the CSV names the
    ones that did not.

    Timing is reported throughout: how long reading each recycle bin took, an
    up-front estimate, a progress bar with elapsed time and a live ETA based on
    the measured rate, and the real duration in the summary. The CSV records the
    site, batch number and duration per item.

    Sign-in: PnP.PowerShell no longer ships a shared multi-tenant app, so an Entra
    app registration is required. This script creates one automatically the first
    time it runs against a tenant:

      1. Signs in to Microsoft Graph as an admin (Application.ReadWrite.All)
      2. Creates (or reuses) a public-client app named after -AppName
      3. Grants and admin-consents the delegated SharePoint scope AllSites.FullControl
      4. Caches the resulting client ID per tenant in pnp.appid.json (gitignored)

    Later runs read the cached client ID and go straight to the interactive login,
    so the Graph admin sign-in only happens once per tenant. Pass an existing
    -ClientId to skip app creation entirely.

    Reading a site's recycle bin requires being site collection administrator
    there. -GrantSiteAdmin arranges that per site via the tenant admin site and
    removes the rights again afterwards (keep them with -KeepSiteAdmin). For
    -AllSites that is effectively required unless you already administer every
    site.

.PARAMETER SiteUrl
    Full URL of a single site collection to restore in. Works for team and
    communication sites and for OneDrive, e.g.
    https://contoso.sharepoint.com/sites/Finance
    https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com

.PARAMETER AllSites
    Walk every SharePoint site in the tenant instead of a single site. OneDrive
    personal sites, the My Site host and redirect sites are skipped, as are sites
    that are locked (read-only or no access).

.PARAMETER TenantUrl
    Tenant SharePoint URL for -AllSites, e.g. https://contoso.sharepoint.com
    (the -admin URL is accepted too).

.PARAMETER SiteFilter
    With -AllSites: only process sites whose URL matches this wildcard, e.g.
    "*/sites/Finance*" or "*/teams/*".

.PARAMETER MaxSites
    With -AllSites: stop after this many sites. Useful to try a tenant-wide run
    on a handful of sites first.

.PARAMETER TenantId
    Tenant ID or domain used for the one-time Graph admin sign-in. Defaults to the
    tenant derived from the site or tenant URL.

.PARAMETER ClientId
    Client ID of an existing Entra app to sign in with. Skips app registration.

.PARAMETER AppName
    Display name of the app registration to create or reuse.
    Default: "M365-Scripts SharePoint Restore".

.PARAMETER Name
    Filter on item name. Wildcards allowed, e.g. "*.xlsx" or "Budget*".

.PARAMETER Path
    Filter on the original location (the folder the item was deleted from).
    Substring match, e.g. "Shared Documents/Finance".

.PARAMETER DeletedBy
    Filter on who deleted the item - matches display name or e-mail, wildcards allowed.

.PARAMETER DeletedAfter
    Only items deleted on or after this date/time.

.PARAMETER DeletedBefore
    Only items deleted before this date/time.

.PARAMETER ItemType
    Restrict to File, Folder or ListItem. Default All.

.PARAMETER Stage
    Which recycle bin to read: FirstStage (user), SecondStage (site collection
    admin) or All (default).

.PARAMETER RowLimit
    Maximum number of recycle bin entries to fetch per site. Use on large recycle
    bins to avoid pulling hundreds of thousands of rows.

.PARAMETER BatchSize
    How many items to restore per server call (1-200, default 200). SharePoint
    restores a whole batch in one round trip, which is what makes a bulk restore
    fast. A batch is all-or-nothing, so a failed batch is automatically retried
    one item at a time. Use -BatchSize 1 to restore strictly item by item.

.PARAMETER GrantSiteAdmin
    Add the signed-in admin as site collection administrator on each site before
    reading its recycle bin (requires SharePoint Administrator). Rights are
    removed again afterwards unless -KeepSiteAdmin.

.PARAMETER KeepSiteAdmin
    Keep the site collection admin rights granted by -GrantSiteAdmin.

.PARAMETER AdminUpn
    UPN to grant site collection admin rights to. Defaults to the signed-in account.

.PARAMETER Apply
    Actually restore the matching items. Without it the script only reports.

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\RecycleBinRestore_<timestamp>.csv
    (~/Downloads on non-Windows).

.PARAMETER Disconnect
    Sign out of PnP when finished.

.EXAMPLE
    # Dry run - show everything in both recycle bins of one site
    .\Restore-RecycleBinItems.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance

.EXAMPLE
    # Restore everything one user deleted last night, in one site
    .\Restore-RecycleBinItems.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
        -DeletedBy "jane.doe@contoso.com" -DeletedAfter "2026-08-27 18:00" -Apply

.EXAMPLE
    # Tenant-wide dry run: which sites hold files this account deleted today?
    .\Restore-RecycleBinItems.ps1 -AllSites -TenantUrl https://contoso.sharepoint.com `
        -DeletedBy "jane.doe@contoso.com" -DeletedAfter "2026-08-28" -GrantSiteAdmin

.EXAMPLE
    # Tenant-wide restore after a bulk delete, trying 5 sites first
    .\Restore-RecycleBinItems.ps1 -AllSites -TenantUrl https://contoso.sharepoint.com `
        -DeletedBy "jane.doe@contoso.com" -DeletedAfter "2026-08-28" `
        -GrantSiteAdmin -MaxSites 5 -Apply

.EXAMPLE
    # Restore a user's OneDrive, temporarily granting yourself site admin
    .\Restore-RecycleBinItems.ps1 `
        -SiteUrl https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com `
        -GrantSiteAdmin -Apply
#>
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Site')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Site', Position = 0)]
    [string] $SiteUrl,

    [Parameter(Mandatory = $true, ParameterSetName = 'Tenant')]
    [switch] $AllSites,

    [Parameter(Mandatory = $true, ParameterSetName = 'Tenant')]
    [string] $TenantUrl,

    [Parameter(ParameterSetName = 'Tenant')]
    [string] $SiteFilter,

    [Parameter(ParameterSetName = 'Tenant')]
    [int] $MaxSites,

    [string] $TenantId,
    [string] $ClientId,
    [string] $AppName = 'M365-Scripts SharePoint Restore',

    [string] $Name,
    [string] $Path,
    [string] $DeletedBy,
    [datetime] $DeletedAfter,
    [datetime] $DeletedBefore,

    [ValidateSet('All', 'File', 'Folder', 'ListItem')]
    [string] $ItemType = 'All',

    [ValidateSet('All', 'FirstStage', 'SecondStage')]
    [string] $Stage = 'All',

    [int] $RowLimit,

    [ValidateRange(1, 200)]
    [int] $BatchSize = 200,

    [switch] $GrantSiteAdmin,
    [switch] $KeepSiteAdmin,
    [string] $AdminUpn,

    [switch] $Apply,
    [string] $OutputPath,
    [switch] $Disconnect
)

$ErrorActionPreference = 'Stop'

# -- Modules -------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name 'PnP.PowerShell')) {
    throw "Module 'PnP.PowerShell' is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser"
}
Import-Module PnP.PowerShell -ErrorAction Stop

# -- Output folder -------------------------------------------------------------
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
if (-not $OutputPath) {
    $OutputPath = Join-Path $outputDir "RecycleBinRestore_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
}

# -- Resolve tenant + admin URL ------------------------------------------------
$tenantMode = $PSCmdlet.ParameterSetName -eq 'Tenant'
$inputUrl   = if ($tenantMode) { $TenantUrl } else { $SiteUrl }
$inputUrl   = $inputUrl.Trim().TrimEnd('/')

if ($inputUrl -notmatch '^https://([a-zA-Z0-9-]+?)(-admin|-my)?\.sharepoint\.(com|de|us|cn)(/|$)') {
    throw "'$inputUrl' does not look like a SharePoint or OneDrive URL."
}
$tenantName = $Matches[1]
$tld        = $Matches[3]
$adminUrl   = "https://$tenantName-admin.sharepoint.$tld"
if (-not $tenantMode) { $SiteUrl = $inputUrl }
if (-not $TenantId)   { $TenantId = "$tenantName.onmicrosoft.com" }

$repoRoot   = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$appIdStore = Join-Path $repoRoot 'pnp.appid.json'

Write-Host ''
if ($tenantMode) {
    Write-Host "  Scope  : all SharePoint sites in $tenantName (OneDrive excluded)" -ForegroundColor Cyan
    if ($SiteFilter) { Write-Host "  Filter : $SiteFilter" -ForegroundColor Cyan }
} else {
    Write-Host "  Scope  : $SiteUrl" -ForegroundColor Cyan
}
Write-Host "  Tenant : $TenantId" -ForegroundColor Cyan
if ($Apply) {
    Write-Host '  Mode   : APPLY - matching items will be restored' -ForegroundColor Yellow
} else {
    Write-Host '  Mode   : DRY RUN - nothing is restored' -ForegroundColor DarkGray
}
Write-Host ''

if ($tenantMode -and -not $GrantSiteAdmin) {
    Write-Warning 'Without -GrantSiteAdmin every site where you are not already site collection admin will be skipped.'
}

# -- Helpers -------------------------------------------------------------------
function Format-Duration {
    param([TimeSpan] $Span)

    if ($Span.TotalSeconds -lt 60)  { return "{0:n0}s" -f $Span.TotalSeconds }
    if ($Span.TotalMinutes -lt 60)  { return "{0}m {1}s" -f [int]$Span.TotalMinutes, $Span.Seconds }
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

function New-RestoreApp {
    <#
        Creates (or reuses) a public-client app in the target tenant and
        admin-consents the delegated SharePoint scope PnP needs. Returns the app
        id. Requires a Graph sign-in as Global / Application Administrator.
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
    $ClientId = New-RestoreApp -Tenant $TenantId -DisplayName $AppName
    Set-CachedClientId -Tenant $TenantId -Id $ClientId
    $appIsNew = $true
}

# -- Connections ---------------------------------------------------------------
function Connect-Site {
    <#
        Returns a connection object instead of relying on the implicit "current"
        connection, so the admin connection and the per-site connection can live
        side by side without reconnecting all the time.

        A freshly created app registration is not replicated everywhere yet, so
        the first sign-in can fail with "application not found". Retry briefly.
    #>
    param([string] $Url, [int] $Retries = 6)

    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        try {
            return Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId -ReturnConnection -ErrorAction Stop
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

# The admin connection is needed to enumerate sites and to grant site admin rights.
$adminConnection = $null
if ($tenantMode -or $GrantSiteAdmin) {
    Write-Host "  Connecting to $adminUrl..." -ForegroundColor Cyan
    $adminConnection = Connect-Site -Url $adminUrl

    if ($GrantSiteAdmin -and -not $AdminUpn) {
        try {
            $currentUser = Get-PnPProperty -ClientObject (Get-PnPWeb -Connection $adminConnection) -Property CurrentUser
            $AdminUpn = if ($currentUser.Email) { $currentUser.Email } else { ($currentUser.LoginName -split '\|')[-1] }
            Write-Host "  Signed in as $AdminUpn" -ForegroundColor Green
        } catch {
            throw "Could not determine the signed-in account - pass -AdminUpn explicitly. $($_.Exception.Message)"
        }
    }
}

# -- Build the site list -------------------------------------------------------
$targetSites = New-Object System.Collections.Generic.List[object]

if ($tenantMode) {
    Write-Host '  Enumerating site collections...' -ForegroundColor DarkGray
    $enumTimer = [System.Diagnostics.Stopwatch]::StartNew()

    # No -IncludeOneDriveSites: personal sites stay out of the tenant sweep.
    $allTenantSites = @(Get-PnPTenantSite -Connection $adminConnection)
    $enumTimer.Stop()

    # Templates without a usable recycle bin, plus anything that is really a
    # OneDrive, plus sites that are locked (a restore there fails anyway).
    $skipTemplates = @('SPSMSITEHOST#0', 'REDIRECTSITE#0')
    $skipped = [ordered]@{ OneDrive = 0; Template = 0; Locked = 0; Filtered = 0 }

    foreach ($site in $allTenantSites) {
        $url = "$($site.Url)".TrimEnd('/')

        if ($url -match '-my\.sharepoint\.' -or "$($site.Template)" -like 'SPSPERS*') { $skipped.OneDrive++; continue }
        if ("$($site.Template)" -in $skipTemplates) { $skipped.Template++; continue }
        if ($site.LockState -and "$($site.LockState)" -ne 'Unlock') { $skipped.Locked++; continue }
        if ($SiteFilter -and ($url -notlike $SiteFilter)) { $skipped.Filtered++; continue }

        $targetSites.Add([pscustomobject]@{
            Url      = $url
            Title    = $site.Title
            Template = $site.Template
        })
    }

    Write-Host "  Sites found          : $($allTenantSites.Count) in $(Format-Duration $enumTimer.Elapsed)" -ForegroundColor Cyan
    Write-Host "  Skipped              : $($skipped.OneDrive) OneDrive, $($skipped.Template) system, $($skipped.Locked) locked, $($skipped.Filtered) filtered out" -ForegroundColor DarkGray

    if ($MaxSites -and $targetSites.Count -gt $MaxSites) {
        Write-Host "  Limited to the first $MaxSites site(s) by -MaxSites (of $($targetSites.Count) matching)." -ForegroundColor Yellow
        $targetSites = [System.Collections.Generic.List[object]]@($targetSites | Select-Object -First $MaxSites)
    }

    Write-Host "  Sites to process     : $($targetSites.Count)" -ForegroundColor Cyan
    Write-Host ''

    if ($targetSites.Count -eq 0) {
        Write-Host '  No sites match. Nothing to do.' -ForegroundColor Yellow
        return
    }
} else {
    $targetSites.Add([pscustomobject]@{ Url = $SiteUrl; Title = ''; Template = '' })
}

# -- Per-site restore ----------------------------------------------------------
# $PSBoundParameters is per-scope, so settle the "was a date filter given?"
# question here rather than inside the function.
$useDeletedAfter  = $PSBoundParameters.ContainsKey('DeletedAfter')
$useDeletedBefore = $PSBoundParameters.ContainsKey('DeletedBefore')

$allResults = New-Object System.Collections.Generic.List[object]

function Add-Result {
    param($Item, [string] $Site, [string] $Status, [double] $Seconds, [string] $Failure, [int] $BatchNumber)

    $allResults.Add([pscustomobject]@{
        Site             = $Site
        Status           = $Status
        ItemType         = $Item.ItemType
        Title            = $Item.Title
        OriginalLocation = $Item.DirName
        DeletedBy        = if ($Item.DeletedByEmail) { $Item.DeletedByEmail } else { $Item.DeletedByName }
        DeletedDate      = $Item.DeletedDate
        SizeBytes        = $Item.Size
        RecycleBin       = $Item.ItemState
        Id               = $Item.Id
        Batch            = $BatchNumber
        DurationSeconds  = [math]::Round($Seconds, 2)
        Error            = $Failure
    })
}

function Invoke-SiteRestore {
    <#
        Reads one site's recycle bin, filters it, and restores the matches in
        batches. Returns a summary object; individual items land in $allResults.
    #>
    param([string] $Url, $Connection)

    $summary = [pscustomobject]@{
        Site           = $Url
        InBin          = 0
        Matched        = 0
        Restored       = 0
        Failed         = 0
        Batches        = 0
        FellBack       = 0
        FetchSeconds   = 0.0
        RestoreSeconds = 0.0
        Error          = ''
    }

    $binParams = @{ Connection = $Connection }
    if ($Stage -eq 'FirstStage')  { $binParams['FirstStage']  = $true }
    if ($Stage -eq 'SecondStage') { $binParams['SecondStage'] = $true }
    if ($RowLimit)                { $binParams['RowLimit']    = $RowLimit }

    $fetchTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $items = @(Get-PnPRecycleBinItem @binParams)
    $fetchTimer.Stop()
    $summary.FetchSeconds = [math]::Round($fetchTimer.Elapsed.TotalSeconds, 1)
    $summary.InBin = $items.Count

    # -- Filter ----------------------------------------------------------------
    $matched = $items | Where-Object {
        $item = $_
        $keep = $true

        if ($Name) { $keep = $keep -and ($item.Title -like $Name) }
        if ($Path) { $keep = $keep -and ("$($item.DirName)" -like "*$Path*") }
        if ($ItemType -ne 'All') { $keep = $keep -and ("$($item.ItemType)" -eq $ItemType) }
        if ($DeletedBy) {
            $keep = $keep -and (("$($item.DeletedByEmail)" -like $DeletedBy) -or ("$($item.DeletedByName)" -like $DeletedBy))
        }
        if ($useDeletedAfter)  { $keep = $keep -and ($item.DeletedDate -ge $DeletedAfter) }
        if ($useDeletedBefore) { $keep = $keep -and ($item.DeletedDate -lt $DeletedBefore) }

        $keep
    }

    # Folders before files, shallow paths before deep ones: an item cannot be
    # restored into a folder that is itself still in the recycle bin.
    $matched = @($matched | Sort-Object `
        @{ Expression = { if ("$($_.ItemType)" -eq 'Folder') { 0 } else { 1 } } }, `
        @{ Expression = { "$($_.DirName)".Length } }, `
        DeletedDate)

    $summary.Matched = $matched.Count
    if ($matched.Count -eq 0) { return $summary }

    # -- Build the batches -----------------------------------------------------
    # Restore-PnPRecycleBinItem -IdList restores a whole set in a single server
    # call, which is where the speed comes from: 200 items in one round trip
    # instead of 200 round trips. Folders and files never share a batch, so the
    # folders-first ordering survives batching.
    $batches = New-Object System.Collections.Generic.List[object]
    foreach ($group in @('Folder', 'Other')) {
        $subset = @($matched | Where-Object {
            if ($group -eq 'Folder') { "$($_.ItemType)" -eq 'Folder' } else { "$($_.ItemType)" -ne 'Folder' }
        })
        for ($i = 0; $i -lt $subset.Count; $i += $BatchSize) {
            $batches.Add(@($subset[$i..([math]::Min($i + $BatchSize - 1, $subset.Count - 1))]))
        }
    }
    $summary.Batches = $batches.Count

    # One site: say up front how long this is going to take. Tenant-wide that
    # would be a wall of text, so there the per-site line and the ETA bar do it.
    if (-not $tenantMode) {
        $estimate = [TimeSpan]::FromSeconds($batches.Count * $script:SecondsPerBatch)
        Write-Host "      Estimated restore time: $(Format-Duration $estimate) - $($matched.Count) item(s) in $($batches.Count) batch(es) of up to $BatchSize" -ForegroundColor DarkGray
    }

    if (-not $Apply) {
        foreach ($item in $matched) { Add-Result -Item $item -Site $Url -Status 'DryRun' -Seconds 0 -Failure '' -BatchNumber 0 }
        return $summary
    }

    # -- Restore ---------------------------------------------------------------
    $done     = 0
    $batchNr  = 0
    $runTimer = [System.Diagnostics.Stopwatch]::StartNew()

    foreach ($batch in $batches) {
        $batchNr++

        $label = "$Url - $($batch.Count) item(s) starting with $($batch[0].Title)"
        if (-not $PSCmdlet.ShouldProcess($label, 'Restore from recycle bin')) {
            foreach ($item in $batch) { Add-Result -Item $item -Site $Url -Status 'Skipped' -Seconds 0 -Failure '' -BatchNumber $batchNr }
            $done += $batch.Count
            continue
        }

        # Average per item over the work done so far, so the ETA reflects the
        # real throughput of this tenant rather than an up-front guess.
        $average   = if ($done -gt 0) { $runTimer.Elapsed.TotalSeconds / $done } else { $script:SecondsPerBatch / [math]::Max($batch.Count, 1) }
        $remaining = [TimeSpan]::FromSeconds($average * ($matched.Count - $done))

        Write-Progress -Id 2 -ParentId 1 -Activity "Restoring in $Url" `
            -Status "Batch $batchNr / $($batches.Count) - $done / $($matched.Count) item(s) done" `
            -CurrentOperation "Elapsed $(Format-Duration $runTimer.Elapsed) - about $(Format-Duration $remaining) to go" `
            -PercentComplete ([int](100 * $done / $matched.Count)) `
            -SecondsRemaining ([int]$remaining.TotalSeconds)

        $batchTimer = [System.Diagnostics.Stopwatch]::StartNew()
        $batchFailure = ''
        try {
            if ($batch.Count -eq 1) {
                Restore-PnPRecycleBinItem -Identity $batch[0].Id -Force -Connection $Connection -ErrorAction Stop
            } else {
                Restore-PnPRecycleBinItem -IdList @($batch.Id | ForEach-Object { "$_" }) -Connection $Connection -ErrorAction Stop
            }
        } catch {
            $batchFailure = $_.Exception.Message.Trim()
        }
        $batchTimer.Stop()

        if (-not $batchFailure) {
            $perItem = $batchTimer.Elapsed.TotalSeconds / $batch.Count
            foreach ($item in $batch) { Add-Result -Item $item -Site $Url -Status 'Restored' -Seconds $perItem -Failure '' -BatchNumber $batchNr }
            $summary.Restored += $batch.Count
            $done += $batch.Count
            continue
        }

        # A batch is all-or-nothing: one bad item takes the whole call down and
        # the error does not say which. Retry the batch one item at a time so the
        # good items still come back and the CSV names the ones that did not.
        if ($batch.Count -eq 1) {
            Add-Result -Item $batch[0] -Site $Url -Status 'Failed' -Seconds $batchTimer.Elapsed.TotalSeconds -Failure $batchFailure -BatchNumber $batchNr
            $summary.Failed++
            $done++
            Write-Warning "  $($batch[0].Title): $batchFailure"
            continue
        }

        $summary.FellBack++
        Write-Warning "  Batch $batchNr failed ($batchFailure) - retrying its $($batch.Count) items one by one."

        foreach ($item in $batch) {
            $itemTimer = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                Restore-PnPRecycleBinItem -Identity $item.Id -Force -Connection $Connection -ErrorAction Stop
                $itemTimer.Stop()
                Add-Result -Item $item -Site $Url -Status 'Restored' -Seconds $itemTimer.Elapsed.TotalSeconds -Failure '' -BatchNumber $batchNr
                $summary.Restored++
            } catch {
                $itemTimer.Stop()
                $message = $_.Exception.Message.Trim()
                Add-Result -Item $item -Site $Url -Status 'Failed' -Seconds $itemTimer.Elapsed.TotalSeconds -Failure $message -BatchNumber $batchNr
                $summary.Failed++
                Write-Warning "  $($item.Title): $message"
            }
            $done++
        }
    }
    $runTimer.Stop()
    Write-Progress -Id 2 -Activity 'Restoring' -Completed
    $summary.RestoreSeconds = [math]::Round($runTimer.Elapsed.TotalSeconds, 1)

    return $summary
}

# A batch call costs roughly a fixed round trip; used only until the first
# measurement comes in.
$script:SecondsPerBatch = 4.0

$siteSummaries = New-Object System.Collections.Generic.List[object]
$totalTimer    = [System.Diagnostics.Stopwatch]::StartNew()
$siteIndex     = 0

try {
    foreach ($site in $targetSites) {
        $siteIndex++

        if ($tenantMode) {
            $average   = if ($siteIndex -gt 1) { $totalTimer.Elapsed.TotalSeconds / ($siteIndex - 1) } else { 0 }
            $remaining = [TimeSpan]::FromSeconds($average * ($targetSites.Count - $siteIndex + 1))
            $eta = if ($average -gt 0) { "elapsed $(Format-Duration $totalTimer.Elapsed) - about $(Format-Duration $remaining) to go" } else { 'estimating...' }

            Write-Progress -Id 1 -Activity 'Walking site collections' `
                -Status "Site $siteIndex / $($targetSites.Count) - $($site.Url)" `
                -CurrentOperation $eta `
                -PercentComplete ([int](100 * ($siteIndex - 1) / $targetSites.Count)) `
                -SecondsRemaining $(if ($average -gt 0) { [int]$remaining.TotalSeconds } else { -1 })
        }

        Write-Host "  [$siteIndex/$($targetSites.Count)] $($site.Url)" -ForegroundColor White

        $grantedHere = $false
        $siteConnection = $null

        try {
            if ($GrantSiteAdmin) {
                if ($PSCmdlet.ShouldProcess($site.Url, "Add $AdminUpn as site collection administrator")) {
                    Set-PnPTenantSite -Identity $site.Url -Owners @($AdminUpn) -Connection $adminConnection -ErrorAction Stop
                    $grantedHere = $true
                }
            }

            $siteConnection = Connect-Site -Url $site.Url
            $summary = Invoke-SiteRestore -Url $site.Url -Connection $siteConnection

            if ($summary.Matched -eq 0) {
                Write-Host "      $($summary.InBin) in bin, nothing matched ($(Format-Duration ([TimeSpan]::FromSeconds($summary.FetchSeconds))) to read)" -ForegroundColor DarkGray
            } elseif ($Apply) {
                Write-Host "      $($summary.Matched) matched - restored $($summary.Restored), failed $($summary.Failed) in $(Format-Duration ([TimeSpan]::FromSeconds($summary.RestoreSeconds)))" `
                    -ForegroundColor $(if ($summary.Failed) { 'Yellow' } else { 'Green' })
            } else {
                Write-Host "      $($summary.Matched) item(s) would be restored in $($summary.Batches) batch(es)" -ForegroundColor Yellow
            }

            $siteSummaries.Add($summary)

        } catch {
            $message = $_.Exception.Message.Trim()
            Write-Warning "  $($site.Url): $message"
            $siteSummaries.Add([pscustomobject]@{
                Site = $site.Url; InBin = 0; Matched = 0; Restored = 0; Failed = 0
                Batches = 0; FellBack = 0; FetchSeconds = 0.0; RestoreSeconds = 0.0; Error = $message
            })
        } finally {
            # Removing site admin rights runs against the site itself, so it needs
            # the site connection. If the connection never came up, say so loudly -
            # the rights are still standing.
            if ($grantedHere -and -not $KeepSiteAdmin) {
                if ($siteConnection) {
                    try {
                        Remove-PnPSiteCollectionAdmin -Owners @($AdminUpn) -Connection $siteConnection -ErrorAction Stop
                    } catch {
                        Write-Warning "Could not remove site collection admin rights for $AdminUpn on $($site.Url): $($_.Exception.Message)"
                    }
                } else {
                    Write-Warning "$AdminUpn is still site collection admin on $($site.Url) - could not connect to remove the rights again."
                }
            }
        }
    }
    $totalTimer.Stop()
    Write-Progress -Id 1 -Activity 'Walking site collections' -Completed

    # -- Report ----------------------------------------------------------------
    if ($allResults.Count -gt 0) {
        $allResults | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    }

    $totals = [pscustomobject]@{
        Sites    = $siteSummaries.Count
        Matched  = ($siteSummaries | Measure-Object -Property Matched  -Sum).Sum
        Restored = ($siteSummaries | Measure-Object -Property Restored -Sum).Sum
        Failed   = ($siteSummaries | Measure-Object -Property Failed   -Sum).Sum
        FellBack = ($siteSummaries | Measure-Object -Property FellBack -Sum).Sum
        Errors   = @($siteSummaries | Where-Object { $_.Error }).Count
    }

    Write-Host ''
    if ($tenantMode) {
        $withHits = @($siteSummaries | Where-Object { $_.Matched -gt 0 -or $_.Error })
        if ($withHits.Count -gt 0) {
            $withHits | Select-Object `
                @{ N = 'Site'; E = { $_.Site -replace '^https://[^/]+', '' } },
                @{ N = 'In bin'; E = { $_.InBin } },
                @{ N = 'Matched'; E = { $_.Matched } },
                @{ N = 'Restored'; E = { $_.Restored } },
                @{ N = 'Failed'; E = { $_.Failed } },
                @{ N = 'Read'; E = { "$($_.FetchSeconds)s" } },
                @{ N = 'Restore'; E = { "$($_.RestoreSeconds)s" } },
                @{ N = 'Error'; E = { $_.Error } } |
                Format-Table -AutoSize | Out-Host
        }
        Write-Host "  Sites processed : $($totals.Sites)" -ForegroundColor Cyan
        if ($totals.Errors) { Write-Host "  Sites in error  : $($totals.Errors)" -ForegroundColor Red }
    }

    if ($Apply) {
        Write-Host "  Restored : $($totals.Restored)" -ForegroundColor Green
        if ($totals.Failed) {
            Write-Host "  Failed   : $($totals.Failed)" -ForegroundColor Red
            Write-Host '  Tip: an item cannot be restored while its original folder is still deleted,' -ForegroundColor DarkGray
            Write-Host '       or when an item with the same name already exists in that location.' -ForegroundColor DarkGray
        }
        if ($totals.FellBack) {
            Write-Host "  Note     : $($totals.FellBack) batch(es) failed as a whole and were retried item by item." -ForegroundColor DarkGray
        }
    } else {
        Write-Host "  DRY RUN - $($totals.Matched) item(s) would be restored. Re-run with -Apply." -ForegroundColor Yellow
        if ($totals.Matched -gt 0) {
            $totalBatches = ($siteSummaries | Measure-Object -Property Batches -Sum).Sum
            $estimateAll  = [TimeSpan]::FromSeconds($totalBatches * $script:SecondsPerBatch)
            Write-Host "  Estimate : about $(Format-Duration $estimateAll) of restoring in $totalBatches batch(es), on top of the $(Format-Duration $totalTimer.Elapsed) spent reading" -ForegroundColor Yellow
        }
    }

    Write-Host "  Duration : $(Format-Duration $totalTimer.Elapsed) total" -ForegroundColor Cyan
    if ($allResults.Count -gt 0) {
        Write-Host "  Report   : $OutputPath" -ForegroundColor Cyan
    } else {
        Write-Host '  Report   : not written - no matching items' -ForegroundColor DarkGray
    }
    Write-Host ''

} finally {
    if ($Disconnect) {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Host '  Disconnected.' -ForegroundColor DarkGray
    }
}
