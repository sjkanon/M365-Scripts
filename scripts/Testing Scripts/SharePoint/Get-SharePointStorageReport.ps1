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

.PARAMETER SiteUrl
    Scan a single site. If omitted, all sites in the tenant are scanned.

.PARAMETER SkipVersions
    Skip version history analysis. Faster but only reports current file sizes.

.PARAMETER OutputPath
    Override the default output folder.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Required when using app-only auth (-ClientId / -ClientSecret).

.PARAMETER ClientId
    App Registration client ID. Use together with -TenantId and -ClientSecret for app-only auth.
    Required to enumerate all sites — delegated auth cannot list all SharePoint sites by design.

.PARAMETER ClientSecret
    Client secret for app-only auth. Use together with -TenantId and -ClientId.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only auth (alternative to -ClientSecret).

.PARAMETER Apply
    Perform the full recursive file scan. Without this switch, only quota data
    from the Graph sites API is retrieved (fast, no file enumeration).

.EXAMPLE
    # Quick summary — site quotas only (no file scan)
    .\Get-SharePointStorageReport.ps1

.EXAMPLE
    # Full scan — all sites, all files, including version history
    .\Get-SharePointStorageReport.ps1 -Apply

.EXAMPLE
    # Full scan — single site
    .\Get-SharePointStorageReport.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Finance" -Apply

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
    [switch] $Apply
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($OutputPath) { $OutputPath }
             elseif ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' }
             else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

$ts          = Get-Date -Format 'yyyyMMdd_HHmmss'
$summaryCsv  = Join-Path $outputDir "SharePoint_Summary_$ts.csv"
$detailCsv   = Join-Path $outputDir "SharePoint_Detail_$ts.csv"

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false

if ($ClientId -and $TenantId) {
    # App-only auth — required for enumerating all sites
    $connectParams = @{ ClientId = $ClientId; TenantId = $TenantId; NoWelcome = $true }

    if ($CertificateThumbprint) {
        $connectParams['CertificateThumbprint'] = $CertificateThumbprint
    } elseif ($ClientSecret) {
        $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
        $connectParams['ClientSecretCredential'] = [System.Management.Automation.PSCredential]::new($ClientId, $secureSecret)
    } else {
        Write-Host "  [ERROR] App-only auth requires -ClientSecret or -CertificateThumbprint." -ForegroundColor Red
        exit 1
    }

    try {
        Connect-MgGraph @connectParams -ErrorAction Stop
        $script:ConnectedHere = $true
    } catch {
        Write-Host "  [ERROR] App-only authentication failed: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
} else {
    # Fall back to existing session or interactive delegated auth
    # Note: enumerating all sites requires app-only auth — use -ClientId / -TenantId / -ClientSecret
    try {
        $null = Get-MgSite -SiteId 'root' -ErrorAction Stop
    } catch {
        $connectParams = @{ Scopes = @('Sites.Read.All', 'Files.Read.All'); NoWelcome = $true }
        if ($TenantId) { $connectParams['TenantId'] = $TenantId }
        Connect-MgGraph @connectParams
        $script:ConnectedHere = $true
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

# ── Get sites ─────────────────────────────────────────────────────────────────
Write-Host "  Retrieving sites..." -ForegroundColor Cyan

if ($SiteUrl) {
    if ($SiteUrl -notmatch 'https://([^/]+)/sites/([^/]+)') {
        Write-Host "  [ERROR] Invalid URL format. Expected: https://tenant.sharepoint.com/sites/sitename" -ForegroundColor Red
        if ($script:ConnectedHere) { Disconnect-MgGraph }
        exit 1
    }
    try {
        $sites = @(Get-MgSite -Search $Matches[2] -ErrorAction Stop |
                   Where-Object { $_.WebUrl -eq $SiteUrl })
        if ($sites.Count -eq 0) {
            Write-Host "  [ERROR] Site not found: $SiteUrl" -ForegroundColor Red
            if ($script:ConnectedHere) { Disconnect-MgGraph }
            exit 1
        }
    } catch {
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        if ($script:ConnectedHere) { Disconnect-MgGraph }
        exit 1
    }
} else {
    # App-only auth required — Get-MgAllSite enumerates all sites including multi-geo
    try {
        $sites = @(Get-MgAllSite -All -Property 'id,displayName,webUrl' -ErrorAction Stop)
    } catch {
        Write-Host "  [ERROR] Failed to retrieve sites: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  [INFO]  Enumerating all sites requires app-only auth." -ForegroundColor Yellow
        Write-Host "          Use: -ClientId <id> -TenantId <id> -ClientSecret <secret>" -ForegroundColor Yellow
        if ($script:ConnectedHere) { Disconnect-MgGraph }
        exit 1
    }
}

Write-Host ("  Found {0} site(s)" -f $sites.Count) -ForegroundColor Green
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

function Get-VersionSize {
    param([string]$DriveId, [string]$ItemId)
    try {
        $versions = Get-MgDriveItemVersion -DriveId $DriveId -DriveItemId $ItemId -ErrorAction Stop
        $size = ($versions | Where-Object { $_.Size } | Measure-Object -Property Size -Sum).Sum
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

    # Seed with root
    $queue.Enqueue([PSCustomObject]@{ Id = 'root'; Path = '' })

    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()

        try {
            $children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $current.Id -All -ErrorAction Stop
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
                    Path             = $path
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
                # Queue folder for processing — no recursion depth limit
                $queue.Enqueue([PSCustomObject]@{ Id = $child.id; Path = $path })
            }
        }
    }

    return $results
}

# ── Process sites ─────────────────────────────────────────────────────────────
$summaryRows = [System.Collections.Generic.List[PSCustomObject]]::new()
$detailRows  = [System.Collections.Generic.List[PSCustomObject]]::new()
$siteIndex   = 0

foreach ($site in $sites) {
    $siteIndex++
    $siteName = $site.displayName ?? $site.name
    $siteId   = $site.id

    Write-Host ("  [{0}/{1}] {2}" -f $siteIndex, $sites.Count, $siteName) -ForegroundColor White

    # Quick mode: use quota data from drives (no file enumeration)
    if (-not $Apply) {
        try {
            $drives = Get-MgSiteDrive -SiteId $siteId -ErrorAction Stop
            foreach ($drive in $drives) {
                $quota = $drive.quota
                $summaryRows.Add([PSCustomObject]@{
                    SiteName        = $siteName
                    SiteUrl         = $site.webUrl
                    Library         = $drive.name
                    UsedGB          = if ($quota.used)  { [math]::Round($quota.used  / 1GB, 3) } else { $null }
                    TotalGB         = if ($quota.total) { [math]::Round($quota.total / 1GB, 3) } else { $null }
                    RemainingGB     = if ($quota.remaining) { [math]::Round($quota.remaining / 1GB, 3) } else { $null }
                    State           = $quota.state
                    FileCount       = $null
                    FolderCount     = $null
                    VersionSizeMB   = $null
                    TotalSizeMB     = $null
                }) | Out-Null
                Write-Host ("        {0,-30} used: {1} GB" -f $drive.name,
                    ([math]::Round(($quota.used ?? 0) / 1GB, 2))) -ForegroundColor DarkGray
            }
        } catch {
            Write-Host "        [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        }
        continue
    }

    # Full scan mode
    try {
        $drives = Get-MgSiteDrive -SiteId $siteId -ErrorAction Stop
    } catch {
        Write-Host "        [ERROR] Cannot access drives: $($_.Exception.Message)" -ForegroundColor Red
        continue
    }

    foreach ($drive in $drives) {
        Write-Host ("        Scanning '{0}'..." -f $drive.name) -ForegroundColor DarkGray

        $items = Get-AllDriveItems -DriveId $drive.id

        $fileItems   = $items
        $totalFiles  = $fileItems.Count
        $currentSize = ($fileItems | Measure-Object -Property SizeBytes -Sum).Sum ?? 0
        $versionSize = ($fileItems | Measure-Object -Property VersionSizeBytes -Sum).Sum ?? 0
        $totalSize   = $currentSize + $versionSize

        Write-Host ("        {0} files | current: {1} MB | versions: {2} MB | total: {3} MB" -f
            $totalFiles,
            [math]::Round($currentSize / 1MB, 1),
            [math]::Round($versionSize / 1MB, 1),
            [math]::Round($totalSize   / 1MB, 1)) -ForegroundColor DarkGray

        $summaryRows.Add([PSCustomObject]@{
            SiteName        = $siteName
            SiteUrl         = $site.webUrl
            Library         = $drive.name
            UsedGB          = $null
            TotalGB         = $null
            RemainingGB     = $null
            State           = $null
            FileCount       = $totalFiles
            FolderCount     = $null
            VersionSizeMB   = [math]::Round($versionSize / 1MB, 2)
            TotalSizeMB     = [math]::Round($totalSize   / 1MB, 2)
        }) | Out-Null

        foreach ($item in $fileItems) {
            $detailRows.Add([PSCustomObject]@{
                SiteName         = $siteName
                SiteUrl          = $site.webUrl
                Library          = $drive.name
                Path             = $item.Path
                SizeMB           = $item.SizeMB
                VersionCount     = $item.VersionCount
                VersionSizeMB    = $item.VersionSizeMB
                TotalSizeMB      = $item.TotalSizeMB
                Modified         = $item.Modified
            }) | Out-Null
        }
    }
}

# ── Export ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Exporting results" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

$summaryRows | Export-Csv -Path $summaryCsv -NoTypeInformation -Encoding UTF8
Write-Host ("  Summary  : {0}" -f $summaryCsv) -ForegroundColor Green

if ($Apply -and $detailRows.Count -gt 0) {
    $detailRows | Export-Csv -Path $detailCsv -NoTypeInformation -Encoding UTF8
    Write-Host ("  Detail   : {0}" -f $detailCsv) -ForegroundColor Green
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Summary" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ("  Sites scanned : {0}" -f $sites.Count)

if ($Apply) {
    $grandTotal = ($summaryRows | Measure-Object -Property TotalSizeMB -Sum).Sum ?? 0
    $grandVer   = ($summaryRows | Measure-Object -Property VersionSizeMB -Sum).Sum ?? 0
    $grandFiles = ($summaryRows | Measure-Object -Property FileCount -Sum).Sum ?? 0
    Write-Host ("  Total files   : {0}"    -f $grandFiles)
    Write-Host ("  Version data  : {0} MB ({1} GB)" -f [math]::Round($grandVer, 0), [math]::Round($grandVer / 1024, 2)) -ForegroundColor Yellow
    Write-Host ("  Grand total   : {0} MB ({1} GB)" -f [math]::Round($grandTotal, 0), [math]::Round($grandTotal / 1024, 2)) -ForegroundColor Green
}
Write-Host ""

# ── Disconnect ────────────────────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-MgGraph }
