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

        # Assign Sites.Read.All application permission + grant admin consent
        $graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ErrorAction Stop
        $appRole = $graphSp.AppRoles | Where-Object { $_.Value -eq 'Sites.Read.All' }
        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $sp.Id `
            -PrincipalId        $sp.Id `
            -ResourceId         $graphSp.Id `
            -AppRoleId          $appRole.Id `
            -ErrorAction Stop | Out-Null
        Write-Host "  [OK]   Sites.Read.All granted." -ForegroundColor DarkGray

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

        $script:AppOnlyHeaders = @{ Authorization = "Bearer $appOnlyToken" }
        Write-Host "  [OK]   Token obtained." -ForegroundColor DarkGray
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Remove-TempApp; exit 1
}

# ── Get sites ─────────────────────────────────────────────────────────────────
Write-Host "  Retrieving sites..." -ForegroundColor Cyan

if ($SiteUrl) {
    if ($SiteUrl -notmatch 'https://([^/]+)/sites/([^/]+)') {
        Write-Host "  [ERROR] Invalid URL format. Expected: https://tenant.sharepoint.com/sites/sitename" -ForegroundColor Red
        Remove-TempApp; exit 1
    }
    try {
        $sites = @(Get-MgSite -Search $Matches[2] -ErrorAction Stop |
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
$sites = @($sites | Where-Object { $_.webUrl -notmatch '-my\.sharepoint\.com/personal/' })

Write-Host ("  Found {0} site(s) (personal sites excluded)" -f $sites.Count) -ForegroundColor Green
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

function Invoke-GraphGet {
    # Unified GET helper: uses app-only REST headers when available, otherwise SDK
    param([string]$Uri)
    if ($script:AppOnlyHeaders) {
        return Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0$Uri" `
            -Headers $script:AppOnlyHeaders -ErrorAction Stop
    } else {
        return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
    }
}

function Get-SiteDrives {
    param([string]$SiteId)
    if ($script:AppOnlyHeaders) {
        $resp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/drives" `
            -Headers $script:AppOnlyHeaders -ErrorAction Stop
        return $resp.value
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
            $drives = Get-SiteDrives -SiteId $siteId
            foreach ($drive in $drives) {
                $quota = $drive.quota
                $summaryRows.Add([PSCustomObject]@{
                    SiteName        = $siteName
                    SiteUrl         = $site.webUrl
                    Library         = $drive.name
                    UsedGB          = if ($quota.used)      { [math]::Round($quota.used      / 1GB, 3) } else { $null }
                    TotalGB         = if ($quota.total)     { [math]::Round($quota.total     / 1GB, 3) } else { $null }
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
        $drives = Get-SiteDrives -SiteId $siteId
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

# ── Cleanup ───────────────────────────────────────────────────────────────────
Remove-TempApp
