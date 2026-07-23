#Requires -Version 5.1
<#
.SYNOPSIS
    Remove SharePoint Online file versions older than a cutoff date while preserving the current version.

.DESCRIPTION
    Connects to Microsoft Graph and scans either a single site or all SharePoint sites in the
    tenant. For each document library, files are enumerated recursively and their version history
    is inspected via Graph. Versions older than -BeforeDate are reported by default. The current
    (most recent) version of a file is always preserved — Graph does not allow deleting it.

    Only when -Apply is specified are matching versions actually removed
    (DELETE /drives/{drive-id}/items/{item-id}/versions/{version-id}). Note that this delete
    operation is not in Microsoft's official Graph API reference, but is widely used and
    confirmed working for both OneDrive and SharePoint document libraries.

    Default behavior is safe preview mode.

    Authentication:
      By default the script connects interactively (delegated) with Sites.ReadWrite.All and
      Files.ReadWrite.All — no Entra ID app registration is required. Scanning all sites in the
      tenant additionally needs app-only auth for the site-listing call only (Microsoft does not
      support delegated tenant-wide site enumeration); the script creates a short-lived, read-only
      temporary App Registration for that single lookup and removes it when done. All file reads
      and version deletions always go through your own delegated permissions, never the temp app.

      To skip the temporary app and use your own existing app registration instead, pass
      -ClientId + -TenantId + -ClientSecret (or -CertificateThumbprint). That app must already
      have Sites.ReadWrite.All application permission granted.

.PARAMETER BeforeDate
    Delete versions older than this date/time.

.PARAMETER SiteUrl
    Optional. Scan a single SharePoint site.

.PARAMETER TenantUrl
    Optional. Tenant root URL, for example https://contoso.sharepoint.com.
    Required when scanning all sites.

.PARAMETER TenantId
    Entra ID tenant ID. Detected automatically from the connected account when omitted.
    Required when using -ClientId.

.PARAMETER ClientId
    Existing App Registration client ID. Skips the automatic temporary app. Use with -TenantId
    and -ClientSecret or -CertificateThumbprint.

.PARAMETER ClientSecret
    Client secret for an existing app registration.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for an existing app registration.

.PARAMETER OutputPath
    Override the default output folder.

.PARAMETER Apply
    Actually remove matching versions. Without this switch, the script only reports.

.PARAMETER IncludeOneDriveSites
    Include personal OneDrive sites in an all-sites scan.

.PARAMETER IncludeHiddenLibraries
    Also inspect hidden document libraries. Disabled by default for safety.

.PARAMETER LibraryTitle
    Optional filter. Limit the scan to one or more library titles.

.PARAMETER GraphTimeoutSec
    Timeout in seconds per Graph call (default: 120).

.PARAMETER MaxGraphRetry
    Max retries on Graph throttling/timeouts (default: 6).

.EXAMPLE
    .\Remove-SharePointFileVersionsByDate.ps1 -TenantUrl "https://contoso.sharepoint.com" -BeforeDate "2025-01-01"

.EXAMPLE
    .\Remove-SharePointFileVersionsByDate.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Finance" -BeforeDate "2025-01-01" -Apply
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [datetime] $BeforeDate,
    [string] $SiteUrl,
    [string] $TenantUrl,
    [string] $TenantId,
    [string] $ClientId,
    [string] $ClientSecret,
    [string] $CertificateThumbprint,
    [string] $OutputPath,
    [switch] $Apply,
    [switch] $IncludeOneDriveSites,
    [switch] $IncludeHiddenLibraries,
    [string[]] $LibraryTitle = @(),
    [int] $GraphTimeoutSec = 120,
    [int] $MaxGraphRetry = 6
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($OutputPath) { $OutputPath }
             elseif ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' }
             else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$detailCsv = Join-Path $outputDir "SharePoint_VersionCleanup_Detail_$ts.csv"
$summaryCsv = Join-Path $outputDir "SharePoint_VersionCleanup_Summary_$ts.csv"

# ── Cleanup tracking ──────────────────────────────────────────────────────────
$script:TempAppObjectId = $null
$script:ConnectedHere   = $false
$script:AppOnlyHeaders  = $null   # set only for the tenant-wide site-listing call in auto mode
$script:TokenBody       = $null
$script:TokenTenantId   = $null
$script:TokenExpiry     = $null

function Write-ProgressHost {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ConsoleColor]$ForegroundColor = [ConsoleColor]::DarkGray
    )
    Write-Host ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message) -ForegroundColor $ForegroundColor
}

function Format-SizeAuto {
    # Picks MB/GB/TB automatically based on magnitude instead of a fixed unit.
    param([Parameter(Mandatory = $true)][double]$MB)
    $absMB = [math]::Abs($MB)
    if ($absMB -ge 1024 * 1024) {
        return "{0:N2} TB" -f ($MB / 1024 / 1024)
    } elseif ($absMB -ge 1024) {
        return "{0:N2} GB" -f ($MB / 1024)
    } else {
        return "{0:N1} MB" -f $MB
    }
}

function Remove-TempApp {
    if ($script:TempAppObjectId) {
        Write-ProgressHost -Message "Removing temporary App Registration..." -ForegroundColor DarkGray
        try {
            Remove-MgApplication -ApplicationId $script:TempAppObjectId -ErrorAction Stop
            Write-ProgressHost -Message "[OK] Temporary App Registration removed." -ForegroundColor DarkGray
        } catch {
            Write-ProgressHost -Message ("[WARN] Could not remove temp App Registration (ID: {0})" -f $script:TempAppObjectId) -ForegroundColor Yellow
            Write-ProgressHost -Message "[WARN] Remove it manually in Entra ID > App registrations." -ForegroundColor Yellow
        }
        $script:TempAppObjectId = $null
    }
    if ($script:ConnectedHere) {
        $prevWarningPreference = $WarningPreference
        try {
            $WarningPreference = 'SilentlyContinue'
            Disconnect-MgGraph -ErrorAction SilentlyContinue
        } catch {} finally {
            $WarningPreference = $prevWarningPreference
        }
        $script:ConnectedHere = $false
    }
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host '   Remove-SharePointFileVersionsByDate' -ForegroundColor Cyan
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host ("  Cutoff    : older than {0}" -f $BeforeDate.ToString('yyyy-MM-dd HH:mm:ss')) -ForegroundColor Cyan
Write-Host ("  Mode      : {0}" -f $(if ($Apply) { 'Apply (delete matching versions)' } else { 'Preview only (no deletion)' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ''

# ── Module preflight ──────────────────────────────────────────────────────────
$missingModules = @('Microsoft.Graph.Authentication') | Where-Object { -not (Get-Module -ListAvailable -Name $_) }
if ($missingModules.Count -gt 0) {
    Write-Host "  [ERROR] Missing required module(s): $($missingModules -join ', ')" -ForegroundColor Red
    Write-Host "  Install with: .\scripts\Startup\Install-Modules.ps1" -ForegroundColor Yellow
    exit 1
}

$allSitesMode = [string]::IsNullOrWhiteSpace($SiteUrl)
if ($allSitesMode -and [string]::IsNullOrWhiteSpace($TenantUrl)) {
    Write-Host '  [ERROR] -TenantUrl is required when scanning all sites.' -ForegroundColor Red
    exit 1
}

$useTempApp = $allSitesMode -and -not $ClientId
if ($useTempApp -and -not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Applications')) {
    Write-Host "  [ERROR] Missing required module: Microsoft.Graph.Applications (needed for the temporary App Registration in all-sites mode)." -ForegroundColor Red
    Write-Host "  Install with: .\scripts\Startup\Install-Modules.ps1" -ForegroundColor Yellow
    exit 1
}

# ── Tenant resolution (GDAP-aware, consistent with Get-SharePointStorageReport.ps1) ──
$effectiveTenantId = $TenantId
if (-not $effectiveTenantId) {
    try {
        if ($global:authMode -and ([string]$global:authMode).ToUpperInvariant() -eq 'GDAP' -and $global:cid) {
            $effectiveTenantId = [string]$global:cid
        } elseif ($env:M365_CUSTOMER_TENANTID) {
            $effectiveTenantId = [string]$env:M365_CUSTOMER_TENANTID
        }
    } catch {}
}

if ($ClientId -and -not $effectiveTenantId) {
    Write-Host '  [ERROR] -ClientId requires -TenantId (or a resolvable GDAP customer tenant context).' -ForegroundColor Red
    exit 1
}

# ── Connection ────────────────────────────────────────────────────────────────
try {
    if ($ClientId -and $effectiveTenantId) {
        # ── Provided app credentials → full app-only SDK connection ──────────
        if ($CertificateThumbprint) {
            Connect-MgGraph -ClientId $ClientId -TenantId $effectiveTenantId `
                -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
        } elseif ($ClientSecret) {
            $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
            $cred = [System.Management.Automation.PSCredential]::new($ClientId, $secureSecret)
            Connect-MgGraph -ClientId $ClientId -TenantId $effectiveTenantId `
                -ClientSecretCredential $cred -NoWelcome -ErrorAction Stop
        } else {
            Write-Host "  [ERROR] -ClientId requires -ClientSecret or -CertificateThumbprint." -ForegroundColor Red
            exit 1
        }
        $script:ConnectedHere = $true
        Write-Host "  [OK]   Connected with provided app credentials." -ForegroundColor DarkGray
    } else {
        Write-Host "  Connecting interactively..." -ForegroundColor Cyan
        $delegatedScopes = @('Sites.ReadWrite.All', 'Files.ReadWrite.All')
        if ($useTempApp) {
            $delegatedScopes += @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
            Write-Host "  Required role: Global Administrator or Application Administrator (for the temporary all-sites lookup app)" -ForegroundColor DarkGray
        }
        $connectParams = @{ Scopes = $delegatedScopes; NoWelcome = $true }
        if ($effectiveTenantId) { $connectParams['TenantId'] = $effectiveTenantId }
        Connect-MgGraph @connectParams -ErrorAction Stop
        $script:ConnectedHere = $true
        Write-Host "  [OK]   Connected (delegated)." -ForegroundColor DarkGray

        if ($useTempApp) {
            $ctx = Get-MgContext
            $usedTenantId = if ($effectiveTenantId) { $effectiveTenantId } else { $ctx.TenantId }
            if (-not $usedTenantId) {
                Write-Host "  [ERROR] Could not determine tenant ID. Provide -TenantId." -ForegroundColor Red
                Remove-TempApp; exit 1
            }

            $appName = "SP-VersionCleanup-Temp-$ts"
            Write-Host "  Creating temporary read-only App Registration '$appName' for site enumeration..." -ForegroundColor Cyan
            $app = New-MgApplication -DisplayName $appName -ErrorAction Stop
            $script:TempAppObjectId = $app.Id

            $sp = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop
            $graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ErrorAction Stop
            $appRole = $graphSp.AppRoles | Where-Object { $_.Value -eq 'Sites.Read.All' }
            if (-not $appRole) {
                Write-Host "  [ERROR] Could not resolve app role 'Sites.Read.All'." -ForegroundColor Red
                Remove-TempApp; exit 1
            }
            New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $sp.Id `
                -PrincipalId        $sp.Id `
                -ResourceId         $graphSp.Id `
                -AppRoleId          $appRole.Id `
                -ErrorAction Stop | Out-Null
            Write-Host "  [OK]   Sites.Read.All granted (enumeration only)." -ForegroundColor DarkGray

            $secret = Add-MgApplicationPassword `
                -ApplicationId      $app.Id `
                -PasswordCredential @{
                    displayName = 'temp'
                    endDateTime = (Get-Date).AddDays(1)
                } -ErrorAction Stop

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
                    $tokenResp = Invoke-RestMethod -Method POST -ErrorAction Stop `
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
            $script:TokenExpiry     = (Get-Date).AddSeconds($tokenResp.expires_in - 300)
            $script:TokenBody       = $tokenBody
            $script:TokenTenantId   = $usedTenantId
            Write-Host "  [OK]   Token obtained (valid until ~$($script:TokenExpiry.ToString('HH:mm')))." -ForegroundColor DarkGray
        }
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Remove-TempApp; exit 1
}

function Update-AppOnlyToken {
    if (-not $script:TokenBody) { return }
    if ((Get-Date) -lt $script:TokenExpiry) { return }
    try {
        $resp = Invoke-RestMethod -Method POST -ErrorAction Stop `
            -Uri  "https://login.microsoftonline.com/$($script:TokenTenantId)/oauth2/v2.0/token" `
            -Body $script:TokenBody
        $script:AppOnlyHeaders = @{ Authorization = "Bearer $($resp.access_token)" }
        $script:TokenExpiry    = (Get-Date).AddSeconds($resp.expires_in - 300)
    } catch {
        Write-Host "  [WARN] Token refresh failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Get-GraphRetryDelaySeconds {
    param([int]$Attempt, [object]$ErrorRecord)
    $retryAfter = $null
    try {
        $resp = $ErrorRecord.Exception.Response
        if ($resp -and $resp.Headers) {
            $retryHeader = $resp.Headers['Retry-After']
            if ($retryHeader) { [void][int]::TryParse([string]$retryHeader, [ref]$retryAfter) }
        }
    } catch {}
    if ($retryAfter -and $retryAfter -gt 0) { return [Math]::Min($retryAfter, 120) }
    return [Math]::Min([int][Math]::Pow(2, [Math]::Max(1, $Attempt)), 60)
}

function Invoke-GraphGet {
    # Uses the app-only bridge token only when set (tenant-wide site enumeration in auto mode).
    # Every other call goes through the delegated (or provided app-only) SDK session, so file
    # reads/deletes always run under the caller's own permissions.
    param([string]$Uri)
    for ($attempt = 1; $attempt -le $MaxGraphRetry; $attempt++) {
        try {
            if ($script:AppOnlyHeaders) {
                Update-AppOnlyToken
                return Invoke-RestMethod -Uri $Uri -Headers $script:AppOnlyHeaders -TimeoutSec $GraphTimeoutSec -ErrorAction Stop
            }
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
        } catch {
            $statusCode = $null
            try {
                if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                }
            } catch {}

            $isRetryable = $statusCode -in @(408, 429, 500, 502, 503, 504)
            if (-not $isRetryable -and -not $statusCode) {
                $isRetryable = $_.Exception.Message -match 'timed out|timeout|temporar|connection|EOF|name resolution'
            }
            if (-not $isRetryable -or $attempt -eq $MaxGraphRetry) { throw }

            $delay = Get-GraphRetryDelaySeconds -Attempt $attempt -ErrorRecord $_
            Write-ProgressHost -Message ("[INFO] Graph request retry ({0}/{1}) in {2}s: {3}" -f $attempt, $MaxGraphRetry, $delay, $Uri) -ForegroundColor DarkGray
            Start-Sleep -Seconds $delay
        }
    }
}

function Invoke-GraphDelete {
    # Deletes a specific file version. Not in Microsoft's official Graph API reference, but a
    # confirmed-working operation (same URL pattern as the documented "get individual version").
    # The current/latest version cannot be deleted this way — Graph rejects it, which is exactly
    # the safety net that keeps the current version intact.
    param([string]$Uri)
    for ($attempt = 1; $attempt -le $MaxGraphRetry; $attempt++) {
        try {
            Invoke-MgGraphRequest -Method DELETE -Uri $Uri -ErrorAction Stop
            return
        } catch {
            $statusCode = $null
            try {
                if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                }
            } catch {}
            $isRetryable = $statusCode -in @(408, 429, 500, 502, 503, 504)
            if (-not $isRetryable -or $attempt -eq $MaxGraphRetry) { throw }
            Start-Sleep -Seconds (Get-GraphRetryDelaySeconds -Attempt $attempt -ErrorRecord $_)
        }
    }
}

function Get-FileVersionsBatch {
    # Resolves version lists for up to 20 files per Graph $batch call. A failed lookup for a
    # file just falls back to "no versions retrievable" instead of stalling the whole scan.
    param([System.Collections.Generic.List[object]] $Requests)   # each: @{ Id; Url }

    $results = @{}
    if ($Requests.Count -eq 0) { return $results }

    for ($i = 0; $i -lt $Requests.Count; $i += 20) {
        $end   = [Math]::Min($i + 19, $Requests.Count - 1)
        $chunk = $Requests.GetRange($i, $end - $i + 1)
        $batchBody = @{
            requests = @($chunk | ForEach-Object { @{ id = $_.Id; method = 'GET'; url = $_.Url } })
        } | ConvertTo-Json -Depth 6

        $done = $false
        for ($attempt = 1; $attempt -le 3 -and -not $done; $attempt++) {
            try {
                $resp = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' `
                    -Body $batchBody -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop
                foreach ($r in $resp.responses) {
                    $results[[string]$r.id] = if ($r.status -eq 200) { $r.body } else { $null }
                }
                $done = $true
            } catch {
                if ($attempt -eq 3) {
                    foreach ($req in $chunk) { $results[$req.Id] = $null }
                } else {
                    Start-Sleep -Seconds ($attempt * 3)
                }
            }
        }
    }
    return $results
}

# ── Get sites ─────────────────────────────────────────────────────────────────
Write-ProgressHost -Message "Retrieving sites..." -ForegroundColor Cyan
$targetSites = [System.Collections.Generic.List[object]]::new()

if (-not $allSitesMode) {
    try {
        $siteUri  = [System.Uri]$SiteUrl.TrimEnd('/')
        $siteHost = $siteUri.Host
        $sitePath = $siteUri.AbsolutePath.TrimEnd('/')
        $siteGraphUri = "https://graph.microsoft.com/v1.0/sites/${siteHost}:${sitePath}?`$select=id,displayName,webUrl"
        $siteObj = Invoke-GraphGet -Uri $siteGraphUri

        if (-not $siteObj -or -not $siteObj.id) {
            Write-Host "  [ERROR] Site not found: $SiteUrl" -ForegroundColor Red
            Remove-TempApp; exit 1
        }
        $targetSites.Add($siteObj) | Out-Null
    } catch {
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        Remove-TempApp; exit 1
    }
} else {
    $firstUri = 'https://graph.microsoft.com/v1.0/sites/getAllSites?$select=id,displayName,webUrl&$top=200'
    $firstDone = $false
    for ($i = 1; $i -le 6; $i++) {
        try {
            $response = Invoke-GraphGet -Uri $firstUri
            $response.value | Where-Object { $_.id } | ForEach-Object { $targetSites.Add($_) }
            $nextUri = $response.'@odata.nextLink'
            $firstDone = $true
            break
        } catch {
            if ($i -lt 6) {
                Write-ProgressHost -Message ("[INFO] Waiting for consent propagation (attempt {0}/6)..." -f $i) -ForegroundColor DarkGray
                Start-Sleep -Seconds 5
            } else {
                Write-Host "  [ERROR] Failed to retrieve sites: $($_.Exception.Message)" -ForegroundColor Red
                Remove-TempApp; exit 1
            }
        }
    }
    while ($firstDone -and $nextUri) {
        $response = Invoke-GraphGet -Uri $nextUri
        $response.value | Where-Object { $_.id } | ForEach-Object { $targetSites.Add($_) }
        $nextUri = $response.'@odata.nextLink'
    }

    if (-not $IncludeOneDriveSites) {
        $targetSites = [System.Collections.Generic.List[object]]::new(
            @($targetSites | Where-Object { $_.webUrl -notmatch '-my\.sharepoint\.com/personal/' })
        )
    }

    # Sub-sites at all depths — getAllSites primarily returns root site collections.
    $subSiteQueue = [System.Collections.Generic.Queue[object]]::new()
    $knownSiteIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $targetSites | ForEach-Object { if ($_.id -and $knownSiteIds.Add($_.id)) { $subSiteQueue.Enqueue($_) } }

    while ($subSiteQueue.Count -gt 0) {
        $parent = $subSiteQueue.Dequeue()
        try {
            $subSitesUri = "https://graph.microsoft.com/v1.0/sites/$($parent.id)/sites`?`$select=id,displayName,webUrl&`$top=200"
            do {
                $subResp = Invoke-GraphGet -Uri $subSitesUri
                $subResp.value | Where-Object { $_.id } | ForEach-Object {
                    if ($knownSiteIds.Add($_.id)) {
                        $targetSites.Add($_)
                        $subSiteQueue.Enqueue($_)
                    }
                }
                $subSitesUri = $subResp.'@odata.nextLink'
            } while ($subSitesUri)
        } catch {
            # Most sites have no sub-sites or are inaccessible with current permissions.
        }
    }
}

Write-ProgressHost -Message ("Target sites: {0}" -f $targetSites.Count) -ForegroundColor Green

# ── Helpers ───────────────────────────────────────────────────────────────────
function Get-SiteDocumentLibraries {
    # /lists?$expand=drive returns every document library (incl. Teams channel libraries),
    # filtered down here to real document libraries (list.template) with drive access.
    param([string]$SiteId)
    $libraries = [System.Collections.Generic.List[object]]::new()
    $listUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists" +
               '?$select=id,displayName,list&$expand=drive($select=id,name,webUrl)&$top=200'
    do {
        $resp = Invoke-GraphGet -Uri $listUri
        $resp.value | Where-Object { $_.drive -and $_.list -and $_.list.template -eq 'documentLibrary' } | ForEach-Object {
            $driveObj = $_.drive
            $driveObj | Add-Member -NotePropertyName 'LibraryTitle' -NotePropertyValue $_.displayName -Force
            $driveObj | Add-Member -NotePropertyName 'Hidden' -NotePropertyValue ([bool]$_.list.hidden) -Force
            $libraries.Add($driveObj)
        }
        $listUri = $resp.'@odata.nextLink'
    } while ($listUri)
    return $libraries
}

function Get-DriveFiles {
    # Iterative breadth-first traversal of a document library's folder tree.
    param([string]$DriveId)
    $files = [System.Collections.Generic.List[object]]::new()
    $queue = [System.Collections.Generic.Queue[string]]::new()
    $queue.Enqueue('root')
    $processedFolders = 0
    while ($queue.Count -gt 0) {
        $folderId = $queue.Dequeue()
        $processedFolders++
        if ($processedFolders % 25 -eq 0) {
            Write-ProgressHost -Message ("progress: {0} folders, {1} files scanned..." -f $processedFolders, $files.Count) -ForegroundColor DarkGray
        }
        $childUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$folderId/children" +
                    '?$select=id,name,file,folder,webUrl&$top=200'
        do {
            $resp = Invoke-GraphGet -Uri $childUri
            foreach ($item in $resp.value) {
                if ($item.folder) { $queue.Enqueue($item.id) }
                elseif ($item.file) { $files.Add($item) }
            }
            $childUri = $resp.'@odata.nextLink'
        } while ($childUri)
    }
    return $files
}

# ── Scan ──────────────────────────────────────────────────────────────────────
$detailRows = [System.Collections.Generic.List[object]]::new()
$summaryRows = [System.Collections.Generic.List[object]]::new()

try {
    $siteIndex = 0
    foreach ($targetSite in $targetSites) {
        $siteIndex++
        Write-ProgressHost -Message ("[{0}/{1}] {2}" -f $siteIndex, $targetSites.Count, $targetSite.webUrl) -ForegroundColor White

        try {
            $libraries = @(Get-SiteDocumentLibraries -SiteId $targetSite.id | Where-Object {
                $IncludeHiddenLibraries -or -not $_.Hidden
            })

            if ($LibraryTitle.Count -gt 0) {
                $libraries = @($libraries | Where-Object { $_.LibraryTitle -in $LibraryTitle })
            }

            $libIndex = 0
            foreach ($library in $libraries) {
                $libIndex++
                Write-ProgressHost -Message ("[{0}/{1}] Library: {2}" -f $libIndex, $libraries.Count, $library.LibraryTitle) -ForegroundColor White

                try {
                    $files = @(Get-DriveFiles -DriveId $library.id)
                } catch {
                    Write-ProgressHost -Message ("[WARN] Cannot enumerate files in {0}: {1}" -f $library.LibraryTitle, $_.Exception.Message) -ForegroundColor Yellow
                    continue
                }

                $candidateCount = 0
                $deletedCount = 0
                $candidateBytes = [int64]0
                $deletedBytes = [int64]0

                $versionRequests = [System.Collections.Generic.List[object]]::new()
                foreach ($file in $files) {
                    $versionRequests.Add([PSCustomObject]@{
                        Id  = $file.id
                        Url = "/drives/$($library.id)/items/$($file.id)/versions?`$select=id,size,lastModifiedDateTime"
                    }) | Out-Null
                }
                if ($versionRequests.Count -gt 0) {
                    Write-ProgressHost -Message ("resolving version history for {0} file(s)..." -f $versionRequests.Count) -ForegroundColor DarkGray
                }
                $versionResults = Get-FileVersionsBatch -Requests $versionRequests

                foreach ($file in $files) {
                    $body = $versionResults[$file.id]
                    if (-not $body) {
                        $detailRows.Add([PSCustomObject]@{
                            SiteUrl         = $targetSite.webUrl
                            Library         = $library.LibraryTitle
                            FileUrl         = $file.webUrl
                            FileName        = $file.name
                            VersionLabel    = $null
                            VersionIdentity = $null
                            VersionCreated  = $null
                            VersionSizeMB   = $null
                            Action          = 'Error'
                            Message         = 'Could not retrieve version history.'
                        }) | Out-Null
                        continue
                    }

                    $versions = @($body.value) | Sort-Object { [datetime]$_.lastModifiedDateTime } -Descending
                    if ($versions.Count -le 1) { continue }   # only the current version exists

                    # Skip index 0 — the most recent version is always the current one and can't be deleted.
                    foreach ($version in ($versions | Select-Object -Skip 1)) {
                        $versionCreated = [datetime]$version.lastModifiedDateTime
                        if ($versionCreated -ge $BeforeDate) { continue }

                        $versionSizeBytes = [int64]($version.size ?? 0)
                        $candidateCount++
                        $candidateBytes += $versionSizeBytes

                        $action = 'WouldDelete'
                        $message = $null
                        if ($Apply) {
                            try {
                                Invoke-GraphDelete -Uri "https://graph.microsoft.com/v1.0/drives/$($library.id)/items/$($file.id)/versions/$($version.id)"
                                $action = 'Deleted'
                                $deletedCount++
                                $deletedBytes += $versionSizeBytes
                            } catch {
                                $action = 'Error'
                                $message = $_.Exception.Message
                            }
                        }

                        $detailRows.Add([PSCustomObject]@{
                            SiteUrl         = $targetSite.webUrl
                            Library         = $library.LibraryTitle
                            FileUrl         = $file.webUrl
                            FileName        = $file.name
                            VersionLabel    = $version.id
                            VersionIdentity = $version.id
                            VersionCreated  = $versionCreated.ToString('s')
                            VersionSizeMB   = [math]::Round($versionSizeBytes / 1MB, 3)
                            Action          = $action
                            Message         = $message
                        }) | Out-Null
                    }
                }

                Write-Host ("        {0} files scanned | candidates: {1} version(s) ({2}){3}" -f
                    $files.Count,
                    $candidateCount,
                    (Format-SizeAuto -MB ($candidateBytes / 1MB)),
                    $(if ($Apply) { " | deleted: {0} version(s) ({1})" -f $deletedCount, (Format-SizeAuto -MB ($deletedBytes / 1MB)) } else { '' })
                ) -ForegroundColor DarkGray

                $summaryRows.Add([PSCustomObject]@{
                    SiteUrl           = $targetSite.webUrl
                    Library           = $library.LibraryTitle
                    FilesScanned      = $files.Count
                    CandidateVersions = $candidateCount
                    CandidateSizeMB   = [math]::Round($candidateBytes / 1MB, 2)
                    DeletedVersions   = $deletedCount
                    DeletedSizeMB     = [math]::Round($deletedBytes / 1MB, 2)
                    Mode              = if ($Apply) { 'Apply' } else { 'Preview' }
                }) | Out-Null
            }
        } catch {
            $summaryRows.Add([PSCustomObject]@{
                SiteUrl           = $targetSite.webUrl
                Library           = $null
                FilesScanned      = 0
                CandidateVersions = 0
                CandidateSizeMB   = 0
                DeletedVersions   = 0
                DeletedSizeMB     = 0
                Mode              = if ($Apply) { 'Apply' } else { 'Preview' }
                Error             = $_.Exception.Message
            }) | Out-Null
            Write-ProgressHost -Message ("[ERROR] Site failed: {0}" -f $_.Exception.Message) -ForegroundColor Red
        }
    }
} finally {
    Remove-TempApp
}

$detailRows | Export-Csv -Path $detailCsv -NoTypeInformation -Encoding UTF8
$summaryRows | Export-Csv -Path $summaryCsv -NoTypeInformation -Encoding UTF8

$totalCandidates = ($summaryRows | Measure-Object -Property CandidateVersions -Sum).Sum ?? 0
$totalDeleted = ($summaryRows | Measure-Object -Property DeletedVersions -Sum).Sum ?? 0
$totalCandidateMB = ($summaryRows | Measure-Object -Property CandidateSizeMB -Sum).Sum ?? 0
$totalDeletedMB = ($summaryRows | Measure-Object -Property DeletedSizeMB -Sum).Sum ?? 0

Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host '   Summary' -ForegroundColor Cyan
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host ("  Detail   : {0}" -f $detailCsv) -ForegroundColor Green
Write-Host ("  Summary  : {0}" -f $summaryCsv) -ForegroundColor Green
Write-Host ("  Candidates: {0} version(s) | {1}" -f $totalCandidates, (Format-SizeAuto -MB $totalCandidateMB)) -ForegroundColor Yellow
Write-Host ("  Deleted   : {0} version(s) | {1}" -f $totalDeleted, (Format-SizeAuto -MB $totalDeletedMB)) -ForegroundColor $(if ($Apply) { 'Magenta' } else { 'DarkGray' })
Write-Host ''
