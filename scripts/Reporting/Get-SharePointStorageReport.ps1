#Requires -Version 5.1
<#
.SYNOPSIS
    Report SharePoint storage usage across all sites in a tenant, including version history.

.DESCRIPTION
    Connects to Microsoft Graph and scans all SharePoint sites in the tenant (or a single
    site if -SiteUrl is provided). For each document library, all files are enumerated
    recursively. Version history is included by default.

    Output:
      - Summary CSV  : one row per site with totals
      - Detail CSV   : one row per file with size + version info
      - Both saved to C:\Temp\ (Windows) or ~/Downloads/ (macOS)

    Run without -Apply for a fast summary (site quota data only, no file enumeration).
    Run with -Apply to perform the full recursive scan including version history.

    Authentication:
      By default the script connects interactively (delegated), creates a temporary App
      Registration with Sites.Read.All application permission, fetches a short-lived
      app-only token for site enumeration, and deletes the app when done. File/drive
      operations use the delegated session throughout.

      Enumerating all sites requires app-only auth — delegated is not supported by Microsoft.

      To skip auto-create and use your own app, pass -ClientId + -TenantId + -ClientSecret
      (or -CertificateThumbprint). The script will then connect fully app-only.

.PARAMETER SiteUrl
    Scan a single site. If omitted, all sites in the tenant are scanned.

.PARAMETER SkipVersions
    Skip version history analysis. Faster but only reports current file sizes.

.PARAMETER OutputPath
    Override the default output folder.

.PARAMETER TenantId
    Entra ID tenant ID. Detected automatically from the connected account when omitted.
    Required when using -ClientId.

.PARAMETER ClientId
    Existing App Registration client ID. Skips auto-create. Use with -TenantId and
    -ClientSecret or -CertificateThumbprint.

.PARAMETER ClientSecret
    Client secret for an existing app registration.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for an existing app registration.

.PARAMETER Apply
    Perform the full recursive file scan. Without this switch, only quota data
    from the Graph sites API is retrieved (fast, no file enumeration).

.PARAMETER UseHighPrivilege
    Optional. In auto mode, grants Sites.FullControl.All application permission
    to the temporary app instead of Sites.Read.All. Use this only when stricter
    tenant settings block read-only enumeration.

.EXAMPLE
    # Auto mode — creates and deletes a temporary App Registration automatically
    .\Get-SharePointStorageReport.ps1 -Apply

.EXAMPLE
    # Quick summary — site quotas only, no file scan
    .\Get-SharePointStorageReport.ps1

.EXAMPLE
    # Full scan — single site
    .\Get-SharePointStorageReport.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Finance" -Apply

.EXAMPLE
    # Full scan using an existing app registration
    .\Get-SharePointStorageReport.ps1 -Apply -ClientId "..." -TenantId "..." -ClientSecret "..."

.EXAMPLE
    # Full scan — skip version history (faster)
    .\Get-SharePointStorageReport.ps1 -Apply -SkipVersions
#>
[CmdletBinding()]
param (
    [string] $SiteUrl,
    [switch] $SkipVersions,
    [string] $OutputPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $ClientSecret,
    [string] $CertificateThumbprint,
    [switch] $Apply,
    [switch] $UseHighPrivilege
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($OutputPath) { $OutputPath }
             elseif ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' }
             else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

$ts          = Get-Date -Format 'yyyyMMdd_HHmmss'
$summaryCsv  = Join-Path $outputDir "SharePoint_Summary_$ts.csv"
$reportCsv   = Join-Path $outputDir "SharePoint_StorageRanked_$ts.csv"
$reportMd    = Join-Path $outputDir "SharePoint_VersionReport_$ts.md"

# ── Cleanup tracking ───────────────────────────────────────────────────────────
$script:TempAppObjectId = $null
$script:ConnectedHere   = $false
$script:AppOnlyHeaders  = $null   # set in auto mode for site enumeration REST calls

function Remove-TempApp {
    # Delegated session is still open here — Remove-MgApplication works
    if ($script:TempAppObjectId) {
        Write-Host "  Removing temporary App Registration..." -ForegroundColor DarkGray
        try {
            Remove-MgApplication -ApplicationId $script:TempAppObjectId -ErrorAction Stop
            Write-Host "  [OK]   Temporary App Registration removed." -ForegroundColor DarkGray
        } catch {
            Write-Host ("  [WARN] Could not remove temp App Registration (ID: {0})" -f $script:TempAppObjectId) -ForegroundColor Yellow
            Write-Host "         Remove it manually in Entra ID > App registrations." -ForegroundColor Yellow
        }
        $script:TempAppObjectId = $null
    }
    if ($script:ConnectedHere) {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
        $script:ConnectedHere = $false
    }
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Get-SharePointStorageReport" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

if (-not $Apply) {
    Write-Host "  ================================================" -ForegroundColor Yellow
    Write-Host "   QUICK MODE — quota data only (no file scan)" -ForegroundColor Yellow
    Write-Host "   Add -Apply for a full recursive scan." -ForegroundColor Yellow
    Write-Host "  ================================================" -ForegroundColor Yellow
    Write-Host ""
} elseif ($SkipVersions) {
    Write-Host "  Mode      : Full scan (version history skipped)" -ForegroundColor Cyan
} else {
    Write-Host "  Mode      : Full scan including version history" -ForegroundColor Cyan
}

if ($UseHighPrivilege) {
    Write-Host "  Privilege : High (Sites.FullControl.All for temporary app)" -ForegroundColor Yellow
} else {
    Write-Host "  Privilege : Standard (Sites.Read.All for temporary app)" -ForegroundColor DarkGray
}

# ── Connection ────────────────────────────────────────────────────────────────
try {
    if ($ClientId -and $TenantId) {
        # ── Provided app credentials → full app-only SDK connection ──────────
        if ($CertificateThumbprint) {
            Connect-MgGraph -ClientId $ClientId -TenantId $TenantId `
                -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
        } elseif ($ClientSecret) {
            $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
            $cred = [System.Management.Automation.PSCredential]::new($ClientId, $secureSecret)
            Connect-MgGraph -ClientId $ClientId -TenantId $TenantId `
                -ClientSecretCredential $cred -NoWelcome -ErrorAction Stop
        } else {
            Write-Host "  [ERROR] -ClientId requires -ClientSecret or -CertificateThumbprint." -ForegroundColor Red
            exit 1
        }
        $script:ConnectedHere = $true
        Write-Host "  [OK]   Connected with provided app credentials." -ForegroundColor DarkGray

    } else {
        # ── Auto mode: delegated session stays open throughout ────────────────
        # The delegated session is used for:  app create/delete, drive ops, file ops
        # A separate short-lived app-only REST token is used only for getAllSites
        Write-Host "  Connecting interactively..." -ForegroundColor Cyan
        Write-Host "  Required role: Global Administrator or Application Administrator" -ForegroundColor DarkGray
        Connect-MgGraph -Scopes @(
            'Application.ReadWrite.All'
            'AppRoleAssignment.ReadWrite.All'
            'Sites.Read.All'
            'Files.Read.All'
        ) -NoWelcome -ErrorAction Stop
        $script:ConnectedHere = $true

        $ctx          = Get-MgContext
        $usedTenantId = if ($TenantId) { $TenantId } else { $ctx.TenantId }
        $requiredSiteRole = if ($UseHighPrivilege) { 'Sites.FullControl.All' } else { 'Sites.Read.All' }
        if (-not $usedTenantId) {
            Write-Host "  [ERROR] Could not determine tenant ID. Provide -TenantId." -ForegroundColor Red
            Remove-TempApp; exit 1
        }

        # Create temporary App Registration
        $appName = "SP-StorageReport-Temp-$ts"
        Write-Host "  Creating temporary App Registration '$appName'..." -ForegroundColor Cyan
        $app = New-MgApplication -DisplayName $appName -ErrorAction Stop
        $script:TempAppObjectId = $app.Id

        # Service Principal
        $sp = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop

        # Assign site application permission + grant admin consent
        $graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ErrorAction Stop
        $appRole = $graphSp.AppRoles | Where-Object { $_.Value -eq $requiredSiteRole }
        if (-not $appRole) {
            Write-Host "  [ERROR] Could not resolve app role '$requiredSiteRole'." -ForegroundColor Red
            Remove-TempApp; exit 1
        }
        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $sp.Id `
            -PrincipalId        $sp.Id `
            -ResourceId         $graphSp.Id `
            -AppRoleId          $appRole.Id `
            -ErrorAction Stop | Out-Null
        Write-Host ("  [OK]   {0} granted." -f $requiredSiteRole) -ForegroundColor DarkGray

        # Create short-lived client secret (expires in 1 day)
        $secret = Add-MgApplicationPassword `
            -ApplicationId      $app.Id `
            -PasswordCredential @{
                displayName = 'temp'
                endDateTime = (Get-Date).AddDays(1)
            } -ErrorAction Stop

        # Get app-only OAuth token via REST — no SDK reconnect needed
        # The delegated session stays open so Remove-MgApplication works at the end
        Write-Host "  Obtaining app-only token for site enumeration..." -ForegroundColor Cyan
        $tokenBody = @{
            grant_type    = 'client_credentials'
            scope         = 'https://graph.microsoft.com/.default'
            client_id     = $app.AppId
            client_secret = $secret.SecretText
        }

        $appOnlyToken = $null
        for ($i = 1; $i -le 6; $i++) {
            try {
                $tokenResp    = Invoke-RestMethod -Method POST -ErrorAction Stop `
                    -Uri  "https://login.microsoftonline.com/$usedTenantId/oauth2/v2.0/token" `
                    -Body $tokenBody
                $appOnlyToken = $tokenResp.access_token
                break
            } catch {
                if ($i -lt 6) {
                    Write-Host ("  [INFO] Waiting for app registration propagation (attempt {0}/6)..." -f $i) -ForegroundColor DarkGray
                    Start-Sleep -Seconds 5
                }
            }
        }

        if (-not $appOnlyToken) {
            Write-Host "  [ERROR] Could not obtain app-only token. Try again in a moment." -ForegroundColor Red
            Remove-TempApp; exit 1
        }

        $script:AppOnlyHeaders  = @{ Authorization = "Bearer $appOnlyToken" }
        $script:TokenExpiry     = (Get-Date).AddSeconds($tokenResp.expires_in - 300)  # refresh 5 min early
        $script:TokenBody       = $tokenBody
        $script:TokenTenantId   = $usedTenantId
        Write-Host "  [OK]   Token obtained (valid until ~$($script:TokenExpiry.ToString('HH:mm')))." -ForegroundColor DarkGray
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Remove-TempApp; exit 1
}

# ── Get sites ─────────────────────────────────────────────────────────────────
Write-Host "  Retrieving sites..." -ForegroundColor Cyan

if ($SiteUrl) {
    if ($SiteUrl -notmatch 'https://([^/]+)/(sites|teams)/([^/?#]+)') {
        Write-Host "  [ERROR] Invalid URL format. Expected: https://tenant.sharepoint.com/sites/<name> or /teams/<name>" -ForegroundColor Red
        Remove-TempApp; exit 1
    }
    try {
        $sites = @(Get-MgSite -Search $Matches[3] -ErrorAction Stop |
                   Where-Object { $_.WebUrl -eq $SiteUrl })
        if ($sites.Count -eq 0) {
            Write-Host "  [ERROR] Site not found: $SiteUrl" -ForegroundColor Red
            Remove-TempApp; exit 1
        }
    } catch {
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        Remove-TempApp; exit 1
    }
} elseif ($script:AppOnlyHeaders) {
    # Auto mode: enumerate all sites using app-only REST token
    # Retry first call — consent may take a few seconds to propagate
    $sites = [System.Collections.Generic.List[object]]::new()
    $firstUri = 'https://graph.microsoft.com/v1.0/sites/getAllSites?$select=id,displayName,webUrl&$top=200'
    $firstDone = $false

    for ($i = 1; $i -le 6; $i++) {
        try {
            $response = Invoke-RestMethod -Uri $firstUri -Headers $script:AppOnlyHeaders -ErrorAction Stop
            $response.value | Where-Object { $_.id } | ForEach-Object { $sites.Add($_) }
            $nextUri = $response.'@odata.nextLink'
            $firstDone = $true
            break
        } catch {
            if ($i -lt 6) {
                Write-Host ("  [INFO] Waiting for consent propagation (attempt {0}/6)..." -f $i) -ForegroundColor DarkGray
                Start-Sleep -Seconds 5
            } else {
                Write-Host "  [ERROR] Failed to retrieve sites: $($_.Exception.Message)" -ForegroundColor Red
                Remove-TempApp; exit 1
            }
        }
    }

    # Continue pagination
    while ($firstDone -and $nextUri) {
        $response = Invoke-RestMethod -Uri $nextUri -Headers $script:AppOnlyHeaders -ErrorAction Stop
        $response.value | Where-Object { $_.id } | ForEach-Object { $sites.Add($_) }
        $nextUri = $response.'@odata.nextLink'
    }
} else {
    # Provided credentials — app-only SDK connection, use Get-MgAllSite
    try {
        $sites = @(Get-MgAllSite -All -Property 'id,displayName,webUrl' -ErrorAction Stop)
    } catch {
        Write-Host "  [ERROR] Failed to retrieve sites: $($_.Exception.Message)" -ForegroundColor Red
        Remove-TempApp; exit 1
    }
}

# Exclude personal OneDrive sites (URLs contain -my.sharepoint.com/personal/)
$sites = [System.Collections.Generic.List[object]]::new(
    @($sites | Where-Object { $_.webUrl -notmatch '-my\.sharepoint\.com/personal/' })
)

# Add sub-sites at all depths — getAllSites/Get-MgAllSite primarily return site collections.
# Standard Teams channels appear as document libraries in the parent site (handled by /lists).
# Private/shared Teams channels appear as separate site collections.
# Classic SharePoint sub-webs require explicit enumeration via /sites/{id}/sites.
$subSiteQueue = [System.Collections.Generic.Queue[object]]::new()
$knownSiteIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

$sites | ForEach-Object {
    if ($_.id -and $knownSiteIds.Add($_.id)) {
        $subSiteQueue.Enqueue($_)
    }
}

while ($subSiteQueue.Count -gt 0) {
    $parent = $subSiteQueue.Dequeue()
    try {
        $subSites = @()

        if ($script:AppOnlyHeaders) {
            Update-AppOnlyToken
            $subResp = Invoke-RestMethod `
                -Uri     "https://graph.microsoft.com/v1.0/sites/$($parent.id)/sites" `
                -Headers $script:AppOnlyHeaders -ErrorAction Stop
            $subSites = @($subResp.value)
        } else {
            $subSites = @(Get-MgSiteSubSite -SiteId $parent.id -All -ErrorAction Stop)
        }

        $subSites | Where-Object { $_.id } | ForEach-Object {
            if ($knownSiteIds.Add($_.id)) {
                $sites.Add($_)               # add to scan list
                $subSiteQueue.Enqueue($_)    # also check its children
            }
        }
    } catch {
        # Most sites have no sub-sites or may be inaccessible with current permissions.
    }
}

Write-Host ("  Found {0} site(s) (site collections + sub-sites included, OneDrive excluded)" -f $sites.Count) -ForegroundColor Green
Write-Host ""

# ── Helpers ───────────────────────────────────────────────────────────────────
$knownExtensions = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
@(
    '.doc','.docx','.docm','.odt','.rtf','.txt','.xls','.xlsx','.xlsm','.xlsb','.ods','.csv',
    '.ppt','.pptx','.pptm','.odp','.pdf','.jpg','.jpeg','.png','.gif','.bmp','.tif','.tiff',
    '.svg','.webp','.psd','.ai','.mp4','.avi','.mkv','.mov','.wmv','.mp3','.wav','.flac',
    '.aac','.zip','.rar','.7z','.tar','.gz','.html','.htm','.css','.js','.ts','.json','.xml',
    '.yaml','.yml','.py','.cs','.java','.cpp','.go','.rs','.sh','.exe','.dll','.msi','.sql',
    '.db','.dwg','.stl','.ttf','.otf','.epub','.msg','.eml','.md','.log','.bak','.ndpi','.prism'
) | ForEach-Object { $knownExtensions.Add($_) | Out-Null }

function Test-IsFile {
    param([object]$Item)
    if ($null -ne $Item.file) { return $true }
    if ($null -ne $Item.folder) { return $false }
    $ext = if ($Item.name -match '\.([^.]+)$') { ".$($Matches[1].ToLower())" } else { '' }
    return $knownExtensions.Contains($ext)
}

function Update-AppOnlyToken {
    # Silently refreshes the app-only token if it expires within 5 minutes
    if (-not $script:TokenBody) { return }
    if ((Get-Date) -lt $script:TokenExpiry) { return }

    try {
        $resp = Invoke-RestMethod -Method POST -ErrorAction Stop `
            -Uri  "https://login.microsoftonline.com/$($script:TokenTenantId)/oauth2/v2.0/token" `
            -Body $script:TokenBody
        $script:AppOnlyHeaders = @{ Authorization = "Bearer $($resp.access_token)" }
        $script:TokenExpiry    = (Get-Date).AddSeconds($resp.expires_in - 300)
        Write-Host "  [INFO] App-only token refreshed (valid until ~$($script:TokenExpiry.ToString('HH:mm')))." -ForegroundColor DarkGray
    } catch {
        Write-Host "  [WARN] Token refresh failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Get-SiteDrives {
    # Uses /lists?$expand=drive to return ALL document libraries per site,
    # including Site Pages, Site Assets, Teams channels, and custom libraries
    # that may not surface in the /drives endpoint.
    param([string]$SiteId)
    if ($script:AppOnlyHeaders) {
        Update-AppOnlyToken
        $drives  = [System.Collections.Generic.List[object]]::new()
        $listUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists" +
                   '?$select=id,displayName,list&$expand=drive($select=id,name,webUrl)&$top=200'
        do {
            $resp = Invoke-RestMethod -Uri $listUri -Headers $script:AppOnlyHeaders -ErrorAction Stop
            # Keep any list that has an associated drive — covers document libraries,
            # Teams channel libraries, picture libraries, form libraries, etc.
            $resp.value |
                Where-Object { $_.drive } |
                ForEach-Object {
                    $driveObj = $_.drive
                    $driveObj | Add-Member -NotePropertyName 'VersioningEnabled' -NotePropertyValue $_.list.enableVersioning  -Force -ErrorAction SilentlyContinue
                    $driveObj | Add-Member -NotePropertyName 'MajorVersionLimit'  -NotePropertyValue $_.list.majorVersionLimit -Force -ErrorAction SilentlyContinue
                    $drives.Add($driveObj)
                }
            $listUri = $resp.'@odata.nextLink'
        } while ($listUri)
        return $drives
    } else {
        return Get-MgSiteDrive -SiteId $SiteId -ErrorAction Stop
    }
}

function Get-VersionSize {
    param([string]$DriveId, [string]$ItemId)
    try {
        if ($script:AppOnlyHeaders) {
            $resp     = Invoke-RestMethod `
                -Uri     "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ItemId/versions" `
                -Headers $script:AppOnlyHeaders -ErrorAction Stop
            $versions = $resp.value
        } else {
            $versions = Get-MgDriveItemVersion -DriveId $DriveId -DriveItemId $ItemId -ErrorAction Stop
        }
        $size = ($versions | Where-Object { $_.size } | Measure-Object -Property size -Sum).Sum
        return @{ Count = $versions.Count; Size = [int64]($size ?? 0) }
    } catch {
        return @{ Count = 0; Size = [int64]0 }
    }
}

function Get-AllDriveItems {
    param([string]$DriveId)

    # Iterative breadth-first traversal — no call stack limit, handles any folder depth
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $queue   = [System.Collections.Generic.Queue[PSCustomObject]]::new()

    $queue.Enqueue([PSCustomObject]@{ Id = 'root'; Path = '' })

    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()

        # Collect all children (paginated)
        $children = [System.Collections.Generic.List[object]]::new()
        try {
            if ($script:AppOnlyHeaders) {
                Update-AppOnlyToken
                $childUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$($current.Id)/children" +
                            '?$select=id,name,size,file,folder,lastModifiedDateTime&$top=200'
                do {
                    $resp = Invoke-RestMethod -Uri $childUri -Headers $script:AppOnlyHeaders -ErrorAction Stop
                    $resp.value | ForEach-Object { $children.Add($_) }
                    $childUri = $resp.'@odata.nextLink'
                } while ($childUri)
            } else {
                Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $current.Id -All -ErrorAction Stop |
                    ForEach-Object { $children.Add($_) }
            }
        } catch {
            Write-Host ("          [ERROR] Cannot read folder '{0}': {1}" -f $current.Path, $_.Exception.Message) -ForegroundColor Red
            continue
        }

        foreach ($child in $children) {
            $path = if ($current.Path) { "$($current.Path)/$($child.name)" } else { $child.name }

            if (Test-IsFile -Item $child) {
                $fileSize = [int64]($child.size ?? 0)
                $verCount = 0
                $verSize  = [int64]0

                if (-not $SkipVersions) {
                    $ver      = Get-VersionSize -DriveId $DriveId -ItemId $child.id
                    $verCount = $ver.Count
                    $verSize  = $ver.Size
                }

                $results.Add([PSCustomObject]@{
                    ItemType         = 'File'
                    Path             = $path
                    Level            = (($path -split '/').Count)
                    ParentPath       = $(if ($path -match '/') { ($path -replace '/[^/]+$','') } else { '/' })
                    SizeBytes        = $fileSize
                    SizeMB           = [math]::Round($fileSize / 1MB, 3)
                    VersionCount     = $verCount
                    VersionSizeBytes = $verSize
                    VersionSizeMB    = [math]::Round($verSize / 1MB, 3)
                    TotalSizeBytes   = $fileSize + $verSize
                    TotalSizeMB      = [math]::Round(($fileSize + $verSize) / 1MB, 3)
                    Modified         = $child.lastModifiedDateTime
                }) | Out-Null
            } else {
                $results.Add([PSCustomObject]@{
                    ItemType         = 'Folder'
                    Path             = $path
                    Level            = (($path -split '/').Count)
                    ParentPath       = $(if ($path -match '/') { ($path -replace '/[^/]+$','') } else { '/' })
                    SizeBytes        = $null
                    SizeMB           = $null
                    VersionCount     = $null
                    VersionSizeBytes = $null
                    VersionSizeMB    = $null
                    TotalSizeBytes   = $null
                    TotalSizeMB      = $null
                    Modified         = $child.lastModifiedDateTime
                }) | Out-Null

                $queue.Enqueue([PSCustomObject]@{ Id = $child.id; Path = $path })
            }
        }
    }

    return $results
}

# ── Phase 1: Enumerate all document libraries ─────────────────────────────────
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Phase 1: Enumerating document libraries" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

$siteLibraries = [System.Collections.Generic.List[PSCustomObject]]::new()
$siteIndex     = 0

foreach ($site in $sites) {
    $siteIndex++
    $siteName = $site.displayName ?? $site.name
    $siteId   = $site.id

    Write-Host ("  [{0}/{1}] {2}" -f $siteIndex, $sites.Count, $siteName) -ForegroundColor White

    try {
        $drives = Get-SiteDrives -SiteId $siteId
        foreach ($drive in $drives) {
            $siteLibraries.Add([PSCustomObject]@{
                Site  = $site
                Drive = $drive
            }) | Out-Null
            Write-Host ("        {0}" -f $drive.name) -ForegroundColor DarkGray
        }
    } catch {
        Write-Host ("        [ERROR] Cannot enumerate libraries: {0}" -f $_.Exception.Message) -ForegroundColor Red
    }
}

Write-Host ""
Write-Host ("  Found {0} document libraries across {1} site(s)" -f
    $siteLibraries.Count, $sites.Count) -ForegroundColor Green
Write-Host ""

# ── Phase 2: Retrieve storage data ────────────────────────────────────────────
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Phase 2: Retrieving storage data" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

$summaryRows = [System.Collections.Generic.List[PSCustomObject]]::new()
$detailRows  = [System.Collections.Generic.List[PSCustomObject]]::new()
$libIndex    = 0

foreach ($entry in $siteLibraries) {
    $libIndex++
    $site     = $entry.Site
    $drive    = $entry.Drive
    $siteName = $site.displayName ?? $site.name

    Write-Host ("  [{0}/{1}] {2} › {3}" -f $libIndex, $siteLibraries.Count, $siteName, $drive.name) -ForegroundColor White

    # Quick mode: use quota data from drives (no file enumeration)
    if (-not $Apply) {
        $quota = $drive.quota
        $summaryRows.Add([PSCustomObject]@{
            SiteName           = $siteName
            SiteUrl            = $site.webUrl
            Library            = $drive.name
            VersioningEnabled  = $drive.VersioningEnabled
            MajorVersionLimit  = if ($drive.MajorVersionLimit -eq 0) { 'Unlimited' } else { $drive.MajorVersionLimit }
            UsedGB             = if ($quota.used)      { [math]::Round($quota.used      / 1GB, 3) } else { $null }
            TotalGB            = if ($quota.total)     { [math]::Round($quota.total     / 1GB, 3) } else { $null }
            RemainingGB        = if ($quota.remaining) { [math]::Round($quota.remaining / 1GB, 3) } else { $null }
            State              = $quota.state
            FileCount          = $null
            FolderCount        = $null
            VersionSizeMB      = $null
            TotalSizeMB        = $null
        }) | Out-Null
        Write-Host ("        used: {0} GB" -f ([math]::Round(($quota.used ?? 0) / 1GB, 2))) -ForegroundColor DarkGray
        continue
    }

    # Full scan mode
    $items = Get-AllDriveItems -DriveId $drive.id

    $fileItems   = @($items | Where-Object { $_.ItemType -eq 'File' })
    $folderItems = @($items | Where-Object { $_.ItemType -eq 'Folder' })

    # Build per-folder aggregated sizes from descendant files
    $folderStats = @{}
    foreach ($folder in $folderItems) {
        $folderStats[$folder.Path] = [PSCustomObject]@{
            Path             = $folder.Path
            Level            = $folder.Level
            ParentPath       = $folder.ParentPath
            Modified         = $folder.Modified
            SizeBytes        = [int64]0
            VersionSizeBytes = [int64]0
            TotalSizeBytes   = [int64]0
            VersionCount     = 0
        }
    }

    # Explicit root level per library
    if (-not $folderStats.ContainsKey('/')) {
        $folderStats['/'] = [PSCustomObject]@{
            Path             = '/'
            Level            = 0
            ParentPath       = ''
            Modified         = $null
            SizeBytes        = [int64]0
            VersionSizeBytes = [int64]0
            TotalSizeBytes   = [int64]0
            VersionCount     = 0
        }
    }

    foreach ($file in $fileItems) {
        $folderStats['/'].SizeBytes        += [int64]($file.SizeBytes ?? 0)
        $folderStats['/'].VersionSizeBytes += [int64]($file.VersionSizeBytes ?? 0)
        $folderStats['/'].TotalSizeBytes   += [int64]($file.TotalSizeBytes ?? 0)
        $folderStats['/'].VersionCount     += [int]($file.VersionCount ?? 0)

        if ($file.Path -notmatch '/') {
            continue
        }

        $parts = $file.Path -split '/'
        for ($i = 0; $i -lt ($parts.Count - 1); $i++) {
            $ancestorPath = ($parts[0..$i] -join '/')
            if (-not $folderStats.ContainsKey($ancestorPath)) {
                $folderStats[$ancestorPath] = [PSCustomObject]@{
                    Path             = $ancestorPath
                    Level            = $i + 1
                    ParentPath       = $(if ($ancestorPath -match '/') { ($ancestorPath -replace '/[^/]+$','') } else { '/' })
                    Modified         = $null
                    SizeBytes        = [int64]0
                    VersionSizeBytes = [int64]0
                    TotalSizeBytes   = [int64]0
                    VersionCount     = 0
                }
            }

            $folderStats[$ancestorPath].SizeBytes        += [int64]($file.SizeBytes ?? 0)
            $folderStats[$ancestorPath].VersionSizeBytes += [int64]($file.VersionSizeBytes ?? 0)
            $folderStats[$ancestorPath].TotalSizeBytes   += [int64]($file.TotalSizeBytes ?? 0)
            $folderStats[$ancestorPath].VersionCount     += [int]($file.VersionCount ?? 0)
        }
    }

    $folderReportRows = @(
        $folderStats.Values | ForEach-Object {
            [PSCustomObject]@{
                ItemType         = 'Folder'
                Path             = $_.Path
                Level            = $_.Level
                ParentPath       = $_.ParentPath
                SizeMB           = [math]::Round($_.SizeBytes / 1MB, 3)
                VersionCount     = $_.VersionCount
                VersionSizeMB    = [math]::Round($_.VersionSizeBytes / 1MB, 3)
                TotalSizeMB      = [math]::Round($_.TotalSizeBytes / 1MB, 3)
                Modified         = $_.Modified
            }
        }
    )
    $totalFiles  = $fileItems.Count
    $totalFolders = $folderReportRows.Count
    $currentSize = ($fileItems | Measure-Object -Property SizeBytes -Sum).Sum ?? 0
    $versionSize = ($fileItems | Measure-Object -Property VersionSizeBytes -Sum).Sum ?? 0
    $totalSize   = $currentSize + $versionSize

    Write-Host ("        {0} folders | {1} files | current: {2} MB | versions: {3} MB | total: {4} MB" -f
        $totalFolders,
        $totalFiles,
        [math]::Round($currentSize / 1MB, 1),
        [math]::Round($versionSize / 1MB, 1),
        [math]::Round($totalSize   / 1MB, 1)) -ForegroundColor DarkGray

    $summaryRows.Add([PSCustomObject]@{
        SiteName          = $siteName
        SiteUrl           = $site.webUrl
        Library           = $drive.name
        VersioningEnabled = $drive.VersioningEnabled
        MajorVersionLimit = if ($drive.MajorVersionLimit -eq 0) { 'Unlimited' } else { $drive.MajorVersionLimit }
        UsedGB            = $null
        TotalGB           = $null
        RemainingGB       = $null
        State             = $null
        FileCount         = $totalFiles
        FolderCount       = $totalFolders
        VersionSizeMB     = [math]::Round($versionSize / 1MB, 2)
        TotalSizeMB       = [math]::Round($totalSize   / 1MB, 2)
    }) | Out-Null

    $verEnabled = $drive.VersioningEnabled
    $verLimit   = if ($drive.MajorVersionLimit -eq 0) { 'Unlimited' } else { $drive.MajorVersionLimit }

    foreach ($item in $folderReportRows) {
        $detailRows.Add([PSCustomObject]@{
            SiteName          = $siteName
            SiteUrl           = $site.webUrl
            Library           = $drive.name
            VersioningEnabled = $verEnabled
            MajorVersionLimit = $verLimit
            ItemType          = $item.ItemType
            Path              = $item.Path
            Level             = $item.Level
            ParentPath        = $item.ParentPath
            SizeMB            = $item.SizeMB
            VersionCount      = $item.VersionCount
            VersionSizeMB     = $item.VersionSizeMB
            TotalSizeMB       = $item.TotalSizeMB
            Modified          = $item.Modified
        }) | Out-Null
    }

    foreach ($item in $fileItems) {
        $detailRows.Add([PSCustomObject]@{
            SiteName          = $siteName
            SiteUrl           = $site.webUrl
            Library           = $drive.name
            VersioningEnabled = $verEnabled
            MajorVersionLimit = $verLimit
            ItemType          = $item.ItemType
            Path              = $item.Path
            Level             = $item.Level
            ParentPath        = $item.ParentPath
            SizeMB            = $item.SizeMB
            VersionCount      = $item.VersionCount
            VersionSizeMB     = $item.VersionSizeMB
            TotalSizeMB       = $item.TotalSizeMB
            Modified          = $item.Modified
        }) | Out-Null
    }
}

# ── Phase 2b: Recycle bins ───────────────────────────────────────────────────
if ($Apply) {
    Write-Host ""
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "   Phase 2b: Recycle bins" -ForegroundColor Cyan
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  NOTE: Recycle bin items count towards SharePoint storage quota." -ForegroundColor DarkGray
    Write-Host ""

    # SharePoint has two stages: first-stage (user) and second-stage (site collection admin).
    # The Graph recycleBin/items endpoint returns items from both stages.
    $processedRbSiteIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    foreach ($site in $sites) {
        $siteId   = $site.id
        $siteName = $site.displayName ?? $site.name

        # One recycle bin per site collection — skip sub-sites that share the same collection root
        if (-not $processedRbSiteIds.Add($siteId)) { continue }

        Write-Host ("  Recycle bin: {0}" -f $siteName) -ForegroundColor White

        try {
            $rbItems = [System.Collections.Generic.List[object]]::new()

            if ($script:AppOnlyHeaders) {
                Update-AppOnlyToken
                $rbUri = "https://graph.microsoft.com/v1.0/sites/$siteId/recycleBin/items" +
                         '?$select=id,name,size,deletedDateTime&$top=200'
                do {
                    $resp = Invoke-RestMethod -Uri $rbUri -Headers $script:AppOnlyHeaders -ErrorAction Stop
                    $resp.value | ForEach-Object { $rbItems.Add($_) }
                    $rbUri = $resp.'@odata.nextLink'
                } while ($rbUri)
            } else {
                # SDK fallback — cmdlet available in Microsoft.Graph.Sites >= 2.x
                Get-MgSiteRecycleBinItem -SiteId $siteId -All `
                    -Property 'id,name,size,deletedDateTime' -ErrorAction Stop |
                    ForEach-Object { $rbItems.Add($_) }
            }

            $rbSizeBytes = [int64](($rbItems | Where-Object { $_.size } |
                               Measure-Object -Property size -Sum).Sum ?? 0)
            $rbCount     = $rbItems.Count

            Write-Host ("        {0} item(s) | {1} MB" -f
                $rbCount, [math]::Round($rbSizeBytes / 1MB, 1)) -ForegroundColor DarkGray

            $summaryRows.Add([PSCustomObject]@{
                SiteName          = $siteName
                SiteUrl           = $site.webUrl
                Library           = 'Recycle Bin (stage 1 + 2)'
                VersioningEnabled = $null
                MajorVersionLimit = $null
                UsedGB            = $null
                TotalGB           = $null
                RemainingGB       = $null
                State             = $null
                FileCount         = $rbCount
                FolderCount       = $null
                VersionSizeMB     = $null
                TotalSizeMB       = [math]::Round($rbSizeBytes / 1MB, 2)
            }) | Out-Null

            foreach ($rbItem in $rbItems) {
                $detailRows.Add([PSCustomObject]@{
                    SiteName          = $siteName
                    SiteUrl           = $site.webUrl
                    Library           = 'Recycle Bin (stage 1 + 2)'
                    VersioningEnabled = $null
                    MajorVersionLimit = $null
                    ItemType          = 'Deleted'
                    Path              = $rbItem.name
                    Level             = $null
                    ParentPath        = $null
                    SizeMB            = [math]::Round([int64]($rbItem.size ?? 0) / 1MB, 3)
                    VersionCount      = $null
                    VersionSizeMB     = $null
                    TotalSizeMB       = [math]::Round([int64]($rbItem.size ?? 0) / 1MB, 3)
                    Modified          = $rbItem.deletedDateTime
                }) | Out-Null
            }

        } catch {
            Write-Host ("        [WARN] Cannot read recycle bin: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
        }
    }
}

# ── Export ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Exporting results" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

$summaryRows = @(
    if ($Apply) {
        $summaryRows | Sort-Object `
            @{ Expression = { if ($null -ne $_.TotalSizeMB) { [double]$_.TotalSizeMB } else { -1 } }; Descending = $true },
            SiteName,
            Library
    } else {
        $summaryRows | Sort-Object `
            @{ Expression = { if ($null -ne $_.UsedGB) { [double]$_.UsedGB } else { -1 } }; Descending = $true },
            SiteName,
            Library
    }
)

if (-not $Apply) {
    $summaryRows | Export-Csv -Path $summaryCsv -NoTypeInformation -Encoding UTF8
    Write-Host ("  Summary  : {0}" -f $summaryCsv) -ForegroundColor Green
}

if ($Apply -and $detailRows.Count -gt 0) {
    $detailRows = @(
        $detailRows | Sort-Object `
            @{ Expression = { if ($null -ne $_.TotalSizeMB) { [double]$_.TotalSizeMB } else { -1 } }; Descending = $true },
            SiteName,
            Library,
            ItemType,
            Level,
            Path
    )
    $detailRows | Export-Csv -Path $reportCsv -NoTypeInformation -Encoding UTF8
    Write-Host ("  Ranked   : {0}" -f $reportCsv) -ForegroundColor Green

    # ── Markdown version report ───────────────────────────────────────────────
    if (-not $SkipVersions) {
        $top10Files = @(
            $detailRows |
                Where-Object { $_.ItemType -eq 'File' -and $null -ne $_.VersionSizeMB -and $_.VersionSizeMB -gt 0 } |
                Sort-Object { [double]$_.VersionSizeMB } -Descending |
                Select-Object -First 10
        )
        $top5Libs = @(
            $summaryRows |
                Where-Object { $null -ne $_.VersionSizeMB -and $_.VersionSizeMB -gt 0 } |
                Sort-Object { [double]$_.VersionSizeMB } -Descending |
                Select-Object -First 5
        )

        $mdLines = [System.Collections.Generic.List[string]]::new()
        $mdLines.Add('# SharePoint Version History Report')
        $mdLines.Add('')
        $mdLines.Add("> Gegenereerd op: $(Get-Date -Format 'dd MMMM yyyy HH:mm')")
        if ($SiteUrl) { $mdLines.Add("> Site: ``$SiteUrl``") }
        $mdLines.Add('')
        $mdLines.Add('---')
        $mdLines.Add('')
        $mdLines.Add('## Samenvatting')
        $mdLines.Add('')
        $mdLines.Add("| | |")
        $mdLines.Add("|---|---|")
        $mdLines.Add(("| Sites gescand | {0} |" -f $sites.Count))
        $mdLines.Add(("| Totaal bestanden | {0} |" -f $grandFiles))
        $mdLines.Add(("| Versiedata | {0} MB ({1} GB) |" -f [math]::Round($grandVer, 0), [math]::Round($grandVer / 1024, 2)))
        $mdLines.Add(("| Totaal (huidig + versies) | {0} MB ({1} GB) |" -f [math]::Round($grandTotal, 0), [math]::Round($grandTotal / 1024, 2)))
        $mdLines.Add('')
        $mdLines.Add('---')
        $mdLines.Add('')
        $mdLines.Add('## Top 5 libraries op versiegrootte')
        $mdLines.Add('')
        $mdLines.Add('| # | Versiegrootte (MB) | Versiebeheer | Max. versies | Site | Library |')
        $mdLines.Add('|---|-------------------:|:------------:|:------------:|------|---------|')
        $i = 0
        foreach ($lib in $top5Libs) {
            $i++
            $mdLines.Add(("| {0} | {1} | {2} | {3} | {4} | {5} |" -f
                $i,
                [math]::Round($lib.VersionSizeMB, 1),
                $(if ($lib.VersioningEnabled) { 'Aan' } else { 'Uit' }),
                $(if ($lib.MajorVersionLimit) { $lib.MajorVersionLimit } else { '—' }),
                $lib.SiteName,
                $lib.Library))
        }
        $mdLines.Add('')
        $mdLines.Add('---')
        $mdLines.Add('')
        $mdLines.Add('## Top 10 bestanden op versiegrootte')
        $mdLines.Add('')
        $mdLines.Add('| # | Versiegrootte (MB) | Versies | Library | Pad | Site |')
        $mdLines.Add('|---|-------------------:|--------:|---------|-----|------|')
        $i = 0
        foreach ($f in $top10Files) {
            $i++
            $mdLines.Add(("| {0} | {1} | {2} | {3} | {4} | {5} |" -f
                $i,
                [math]::Round($f.VersionSizeMB, 1),
                $f.VersionCount,
                $f.Library,
                $f.Path,
                $f.SiteName))
        }
        $mdLines.Add('')

        $mdLines | Set-Content -Path $reportMd -Encoding UTF8
        Write-Host ("  Rapport  : {0}" -f $reportMd) -ForegroundColor Green
    }
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Summary" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ("  Sites scanned : {0}" -f $sites.Count)

if ($Apply) {
    $grandFiles = ($summaryRows | Where-Object { $_.Library -ne 'Recycle Bin (stage 1 + 2)' } | Measure-Object -Property FileCount -Sum).Sum ?? 0
    $grandVer   = ($summaryRows | Where-Object { $_.Library -ne 'Recycle Bin (stage 1 + 2)' } | Measure-Object -Property VersionSizeMB -Sum).Sum ?? 0
    $grandRB    = ($summaryRows | Where-Object { $_.Library -eq 'Recycle Bin (stage 1 + 2)' } | Measure-Object -Property TotalSizeMB  -Sum).Sum ?? 0
    $grandTotal = ($summaryRows | Measure-Object -Property TotalSizeMB -Sum).Sum ?? 0
    Write-Host ("  Total files   : {0}"    -f $grandFiles)
    Write-Host ("  Version data  : {0} MB ({1} GB)" -f [math]::Round($grandVer, 0), [math]::Round($grandVer / 1024, 2)) -ForegroundColor Yellow
    Write-Host ("  Recycle bins  : {0} MB ({1} GB)" -f [math]::Round($grandRB, 0),  [math]::Round($grandRB  / 1024, 2)) -ForegroundColor Magenta
    Write-Host ("  Grand total   : {0} MB ({1} GB)" -f [math]::Round($grandTotal, 0), [math]::Round($grandTotal / 1024, 2)) -ForegroundColor Green

    # ── Top libraries by version history size ─────────────────────────────────
    Write-Host ""
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "   Top 5 libraries by version history size" -ForegroundColor Cyan
    Write-Host "  ================================================" -ForegroundColor Cyan
    $summaryRows |
        Where-Object { $null -ne $_.VersionSizeMB -and $_.VersionSizeMB -gt 0 } |
        Sort-Object { [double]$_.VersionSizeMB } -Descending |
        Select-Object -First 5 |
        ForEach-Object {
            Write-Host ("  {0} MB  [{1}] › {2}" -f
                [math]::Round($_.VersionSizeMB, 1),
                $_.SiteName,
                $_.Library) -ForegroundColor Yellow
        }

    # ── Top files by version history size ─────────────────────────────────────
    if (-not $SkipVersions -and $detailRows.Count -gt 0) {
        Write-Host ""
        Write-Host "  ================================================" -ForegroundColor Cyan
        Write-Host "   Top 10 files by version history size" -ForegroundColor Cyan
        Write-Host "  ================================================" -ForegroundColor Cyan
        $detailRows |
            Where-Object { $_.ItemType -eq 'File' -and $null -ne $_.VersionSizeMB -and $_.VersionSizeMB -gt 0 } |
            Sort-Object { [double]$_.VersionSizeMB } -Descending |
            Select-Object -First 10 |
            ForEach-Object {
                Write-Host ("  {0} MB  ({1} versies)  {2} › {3}" -f
                    [math]::Round($_.VersionSizeMB, 1),
                    $_.VersionCount,
                    $_.Library,
                    $_.Path) -ForegroundColor Yellow
                Write-Host ("          Site: {0}" -f $_.SiteName) -ForegroundColor DarkGray
            }
    }
}
Write-Host ""

# ── Cleanup ───────────────────────────────────────────────────────────────────
Remove-TempApp
