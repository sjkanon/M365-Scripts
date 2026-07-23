#Requires -Version 5.1
<#
.SYNOPSIS
    Report SharePoint storage usage across all sites in a tenant, including version history.

.DESCRIPTION
    Connects to Microsoft Graph and scans all SharePoint sites in the tenant (or a single
    site if -SiteUrl is provided). For each document library, all files are enumerated
    recursively. Version history is included by default.

    Output:
      - Summary CSV           : one row per site with totals (quick mode only)
      - Detail CSV            : one row per file/folder with size + version info (-Apply)
      - Site collection totals: one row per root site collection — sub-sites/channels and
                                 the recycle bin rolled up together, comparable 1:1 with the
                                 SharePoint admin center's per-site storage figure (-Apply)
      - All saved to C:\Temp\ (Windows) or ~/Downloads/ (macOS)

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

.PARAMETER FastMode
    Perform the recursive scan with a smaller output footprint: skip version history
    and per-item detail rows, while still producing summary totals and checkpoints.

.PARAMETER UseHighPrivilege
    Optional. In auto mode, grants Sites.FullControl.All application permission
    to the temporary app instead of Sites.Read.All. Use this only when stricter
    tenant settings block read-only enumeration.

.PARAMETER RecycleBinOnly
    Skip storage/library scanning and only read SharePoint recycle bin items
    (stage 1 + stage 2) for each site collection.

.PARAMETER IncludeOneDriveUsers
    One or more user principal names whose OneDrive personal site should be included
    in the scan alongside SharePoint sites (e.g. dilara.beerten@d-build.be).
    OneDrive personal sites are excluded by default.

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
    [switch] $FastMode,
    [switch] $UseHighPrivilege,
    [switch] $RecycleBinOnly,
    [switch] $ForceAppOnlySingleSite,
    [string[]] $IncludeOneDriveUsers = @(),
    [int] $GraphTimeoutSec = 120,
    [int] $MaxGraphRetry = 6,
    [ValidateRange(1, 8)]
    [int] $VersionBatchConcurrency = 4
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
$collectionCsv = Join-Path $outputDir "SharePoint_SiteCollectionTotals_$ts.csv"

# ── Cleanup tracking ───────────────────────────────────────────────────────────
$script:TempAppObjectId = $null
$script:ConnectedHere   = $false
$script:AppOnlyHeaders  = $null   # set in auto mode for site enumeration REST calls
$script:SpoHostTokenCache = @{}
$script:VersionLookupFailureCount = 0
$script:VersionLookupFailureSamples = [System.Collections.Generic.List[string]]::new()

function Write-ProgressHost {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ConsoleColor]$ForegroundColor = [ConsoleColor]::DarkGray
    )
    Write-Host ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message) -ForegroundColor $ForegroundColor
}

function Format-SizeAuto {
    # Picks MB/GB/TB automatically based on magnitude instead of a fixed unit,
    # so small libraries don't print "0.00 GB" and tenant totals don't print in millions of MB.
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
    # Delegated session is still open here — Remove-MgApplication works
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
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Get-SharePointStorageReport" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

if ($RecycleBinOnly) {
    Write-Host "  Mode      : Recycle bin only" -ForegroundColor Cyan
    Write-Host "  Scope     : Site collection recycle bins (stage 1 + 2)" -ForegroundColor DarkGray
} elseif (-not $Apply) {
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

if ($FastMode) {
    Write-Host "  Mode      : Fast scan (no version history, no detail rows)" -ForegroundColor Cyan
    $SkipVersions = $true
}

# ── Module preflight ─────────────────────────────────────────────────────────
$requiredGraphModules = @(
    'Microsoft.Graph.Authentication'
    'Microsoft.Graph.Sites'
)

$missingGraphModules = $requiredGraphModules | Where-Object {
    -not (Get-Module -ListAvailable -Name $_)
}

if ($missingGraphModules.Count -gt 0) {
    Write-Host "  [ERROR] Missing required module(s): $($missingGraphModules -join ', ')" -ForegroundColor Red
    Write-Host "  Install with: .\scripts\Startup\Install-Modules.ps1" -ForegroundColor Yellow
    exit 1
}

# ── Determine scan mode early (impacts auth flow and performance) ───────────
$scanAllSites = $false
$isSingleSiteScan = $false

if ($SiteUrl) {
    $normalizedSiteUrl = $SiteUrl.TrimEnd('/')

    # A tenant root URL should behave like omitting -SiteUrl (scan all sites).
    if ($normalizedSiteUrl -match '^https://[^/]+$') {
        Write-ProgressHost -Message "[INFO] Tenant root URL detected; running tenant-wide scan." -ForegroundColor DarkGray
        $scanAllSites = $true
        $SiteUrl = $null
    } elseif ($normalizedSiteUrl -match '^https://[^/]+/(sites|teams)/[^?#]+$') {
        $SiteUrl = $normalizedSiteUrl
        $isSingleSiteScan = $true
    } else {
        Write-ProgressHost -Message "[ERROR] Invalid URL format. Expected: https://tenant.sharepoint.com/sites/<name>, /teams/<name>, or tenant root URL." -ForegroundColor Red
        exit 1
    }
}

$isGdapMode = $false
try {
    if ($global:authMode -and ([string]$global:authMode).ToUpperInvariant() -eq 'GDAP') {
        $isGdapMode = $true
    } elseif ($env:M365_AUTH_MODE -and ([string]$env:M365_AUTH_MODE).ToUpperInvariant() -eq 'GDAP') {
        $isGdapMode = $true
    }
} catch {}

$effectiveTenantId = $TenantId
if (-not $effectiveTenantId) {
    try {
        if ($isGdapMode -and $global:cid) {
            $effectiveTenantId = [string]$global:cid
        } elseif ($env:M365_CUSTOMER_TENANTID) {
            $effectiveTenantId = [string]$env:M365_CUSTOMER_TENANTID
        }
    } catch {}
}

$useAppOnlyForSingleSite = ($isSingleSiteScan -and ($ForceAppOnlySingleSite -or $isGdapMode))
if ($useAppOnlyForSingleSite) {
    if ($ForceAppOnlySingleSite) {
        Write-ProgressHost -Message '[INFO] Single-site app-only mode enabled by -ForceAppOnlySingleSite.' -ForegroundColor DarkGray
    } elseif ($isGdapMode) {
        Write-ProgressHost -Message '[INFO] GDAP mode detected; using app-only path for single-site reliability.' -ForegroundColor DarkGray
    }
}

$needsAppOnlyEnumeration = ((-not $SiteUrl) -or $scanAllSites -or $useAppOnlyForSingleSite)

if ($isGdapMode -and $needsAppOnlyEnumeration -and -not $effectiveTenantId) {
    Write-ProgressHost -Message '[ERROR] GDAP mode detected but no customer TenantId found. Run Connect-Tenant first or pass -TenantId.' -ForegroundColor Red
    exit 1
}

if ($ClientId -and -not $effectiveTenantId) {
    Write-ProgressHost -Message '[ERROR] -ClientId requires -TenantId (or a resolvable GDAP customer tenant context).' -ForegroundColor Red
    exit 1
}

# ── Connection ────────────────────────────────────────────────────────────────
try {
    if ($ClientId -and ($TenantId -or $effectiveTenantId)) {
        $resolvedTenantId = if ($TenantId) { $TenantId } else { $effectiveTenantId }

        # ── Provided app credentials → full app-only SDK connection ──────────
        if ($CertificateThumbprint) {
            Connect-MgGraph -ClientId $ClientId -TenantId $resolvedTenantId `
                -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
        } elseif ($ClientSecret) {
            $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
            $cred = [System.Management.Automation.PSCredential]::new($ClientId, $secureSecret)
            Connect-MgGraph -ClientId $ClientId -TenantId $resolvedTenantId `
                -ClientSecretCredential $cred -NoWelcome -ErrorAction Stop

            # Keep raw app credentials for non-Graph fallback token requests (for example SPO REST).
            $script:TokenBody = @{
                grant_type    = 'client_credentials'
                scope         = 'https://graph.microsoft.com/.default'
                client_id     = $ClientId
                client_secret = $ClientSecret
            }
            $script:TokenTenantId = $resolvedTenantId
        } else {
            Write-Host "  [ERROR] -ClientId requires -ClientSecret or -CertificateThumbprint." -ForegroundColor Red
            exit 1
        }
        $script:ConnectedHere = $true
        Write-Host "  [OK]   Connected with provided app credentials." -ForegroundColor DarkGray

    } else {
        # ── Auto mode: delegated session stays open throughout ────────────────
        # The delegated session is always used for drive/file operations.
        # For tenant-wide enumeration we temporarily add app-only getAllSites.
        if ($needsAppOnlyEnumeration) {
            Write-Host "  Connecting interactively..." -ForegroundColor Cyan
            Write-Host "  Required role: Global Administrator or Application Administrator" -ForegroundColor DarkGray
            $connectParams = @{
                Scopes    = @(
                'Application.ReadWrite.All'
                'AppRoleAssignment.ReadWrite.All'
                'Sites.Read.All'
                'Files.Read.All'
                )
                NoWelcome = $true
            }
            if ($effectiveTenantId) {
                $connectParams['TenantId'] = $effectiveTenantId
            }
            Connect-MgGraph @connectParams -ErrorAction Stop
            $script:ConnectedHere = $true

            $ctx          = Get-MgContext
            $usedTenantId = if ($effectiveTenantId) { $effectiveTenantId } else { $ctx.TenantId }
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
            Write-Host ("  [OK]   {0} granted (Graph)." -f $requiredSiteRole) -ForegroundColor DarkGray

            # Also grant a role on the SharePoint service principal so the app can obtain a
            # SharePoint-scoped token for SPO REST calls (recycle bin, etc.). Reading a site's
            # recycle bin needs elevated (manage/full-control level) rights — Sites.Read.All is
            # not enough and the call fails silently — so this must follow -UseHighPrivilege the
            # same way the Graph-scoped role above does, instead of always requesting read-only.
            try {
                $spoSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0ff1-ce00-000000000000'" -ErrorAction Stop
                $spoRoleCandidates = if ($UseHighPrivilege) {
                    @('Sites.FullControl.All', 'AllSites.FullControl', 'AllSites.Manage', 'Sites.Manage.All', 'Sites.Read.All', 'AllSites.Read')
                } else {
                    @('Sites.Read.All', 'AllSites.Read')
                }
                $spoRole = $null
                foreach ($candidate in $spoRoleCandidates) {
                    $spoRole = $spoSp.AppRoles | Where-Object { $_.Value -eq $candidate } | Select-Object -First 1
                    if ($spoRole) { break }
                }
                if ($spoRole) {
                    New-MgServicePrincipalAppRoleAssignment `
                        -ServicePrincipalId $sp.Id `
                        -PrincipalId        $sp.Id `
                        -ResourceId         $spoSp.Id `
                        -AppRoleId          $spoRole.Id `
                        -ErrorAction Stop | Out-Null
                    Write-Host ("  [OK]   {0} granted (SharePoint REST)." -f $spoRole.Value) -ForegroundColor DarkGray
                }
            } catch {
                Write-Host "  [WARN] Could not grant SharePoint REST permission to temp app. Recycle bin data may be unavailable." -ForegroundColor Yellow
            }

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
        } else {
            Write-Host "  Connecting interactively (single-site optimized mode)..." -ForegroundColor Cyan
            $connectParams = @{
                Scopes    = @(
                'Sites.Read.All'
                'Files.Read.All'
                )
                NoWelcome = $true
            }
            if ($effectiveTenantId) {
                $connectParams['TenantId'] = $effectiveTenantId
            }
            Connect-MgGraph @connectParams -ErrorAction Stop
            $script:ConnectedHere = $true
            Write-Host "  [OK]   Connected (delegated single-site mode, no temporary app)." -ForegroundColor DarkGray
        }
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Remove-TempApp; exit 1
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

function Get-GraphRetryDelaySeconds {
    param(
        [int]$Attempt,
        [object]$ErrorRecord
    )

    $retryAfter = $null
    try {
        $resp = $ErrorRecord.Exception.Response
        if ($resp -and $resp.Headers) {
            $retryHeader = $resp.Headers['Retry-After']
            if ($retryHeader) {
                [void][int]::TryParse([string]$retryHeader, [ref]$retryAfter)
            }
        }
    } catch {}

    if ($retryAfter -and $retryAfter -gt 0) {
        return [Math]::Min($retryAfter, 120)
    }

    return [Math]::Min([int][Math]::Pow(2, [Math]::Max(1, $Attempt)), 60)
}

function Invoke-GraphGet {
    param(
        [string]$Uri,
        [hashtable]$Headers
    )

    for ($attempt = 1; $attempt -le $MaxGraphRetry; $attempt++) {
        try {
            Update-AppOnlyToken
            return Invoke-RestMethod -Uri $Uri -Headers $Headers -TimeoutSec $GraphTimeoutSec -ErrorAction Stop
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

            if (-not $isRetryable -or $attempt -eq $MaxGraphRetry) {
                throw
            }

            $delay = Get-GraphRetryDelaySeconds -Attempt $attempt -ErrorRecord $_
            Write-ProgressHost -Message (
                "[INFO] Graph request retry ({0}/{1}) in {2}s: {3}" -f
                $attempt,
                $MaxGraphRetry,
                $delay,
                $Uri
            ) -ForegroundColor DarkGray
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-TextHashHex {
    param([string]$Text)

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $hash  = $sha256.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hash) -replace '-', '').ToLowerInvariant()
    } finally {
        $sha256.Dispose()
    }
}

function Get-LibraryCheckpointKey {
    param(
        [string]$SiteId,
        [string]$DriveId
    )

    return ('{0}|{1}' -f $SiteId, $DriveId)
}

function Append-CheckpointRows {
    param(
        [string]$Path,
        [object[]]$Rows
    )

    if ($Rows.Count -eq 0) { return }

    if (Test-Path $Path) {
        $Rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Append
    } else {
        $Rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
    }
}

# ── Get sites ─────────────────────────────────────────────────────────────────
Write-ProgressHost -Message "Retrieving sites..." -ForegroundColor Cyan

if ($isSingleSiteScan) {
    try {
        $siteUri = [System.Uri]$SiteUrl
        $siteHost = $siteUri.Host
        $sitePath = $siteUri.AbsolutePath.TrimEnd('/')
        $siteGraphUri = "https://graph.microsoft.com/v1.0/sites/${siteHost}:${sitePath}?`$select=id,displayName,webUrl,name"

        if ($script:AppOnlyHeaders) {
            $siteObj = Invoke-GraphGet -Uri $siteGraphUri -Headers $script:AppOnlyHeaders
        } else {
            $siteObj = Invoke-MgGraphRequest -Method GET -Uri $siteGraphUri -OutputType PSObject -ErrorAction Stop
        }

        if (-not $siteObj -or -not $siteObj.id) {
            Write-ProgressHost -Message "[ERROR] Site not found: $SiteUrl" -ForegroundColor Red
            Remove-TempApp; exit 1
        }

        $sites = @($siteObj)
    } catch {
        Write-ProgressHost -Message "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
        Remove-TempApp; exit 1
    }
}

if ((-not $SiteUrl) -or $scanAllSites) {
    if ($script:AppOnlyHeaders) {
    # Auto mode: enumerate all sites using app-only REST token
    # Retry first call — consent may take a few seconds to propagate
    $sites = [System.Collections.Generic.List[object]]::new()
    $firstUri = 'https://graph.microsoft.com/v1.0/sites/getAllSites?$select=id,displayName,webUrl&$top=200'
    $firstDone = $false

    for ($i = 1; $i -le 6; $i++) {
        try {
            $response = Invoke-GraphGet -Uri $firstUri -Headers $script:AppOnlyHeaders
            $response.value | Where-Object { $_.id } | ForEach-Object { $sites.Add($_) }
            $nextUri = $response.'@odata.nextLink'
            $firstDone = $true
            break
        } catch {
            if ($i -lt 6) {
                Write-ProgressHost -Message ("[INFO] Waiting for consent propagation (attempt {0}/6)..." -f $i) -ForegroundColor DarkGray
                Start-Sleep -Seconds 5
            } else {
                Write-ProgressHost -Message "[ERROR] Failed to retrieve sites: $($_.Exception.Message)" -ForegroundColor Red
                Remove-TempApp; exit 1
            }
        }
    }

    # Continue pagination
    while ($firstDone -and $nextUri) {
        $response = Invoke-GraphGet -Uri $nextUri -Headers $script:AppOnlyHeaders
        $response.value | Where-Object { $_.id } | ForEach-Object { $sites.Add($_) }
        $nextUri = $response.'@odata.nextLink'
    }
    } else {
    # Provided credentials — app-only SDK connection, use Get-MgAllSite
    try {
        $sites = @(Get-MgAllSite -All -Property 'id,displayName,webUrl' -ErrorAction Stop)
    } catch {
        Write-ProgressHost -Message "[ERROR] Failed to retrieve sites: $($_.Exception.Message)" -ForegroundColor Red
        Remove-TempApp; exit 1
    }
    }
}

# Exclude personal OneDrive sites (URLs contain -my.sharepoint.com/personal/) — from both the
# storage/library scan and the recycle bin phases. Only real SharePoint site collections count.
$sites = [System.Collections.Generic.List[object]]::new(
    @($sites | Where-Object {
        -not [string]::IsNullOrWhiteSpace([string]$_.id) -and
        ($_.webUrl -notmatch '-my\.sharepoint\.com/personal/')
    })
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
            $subSitesList = [System.Collections.Generic.List[object]]::new()
            $subSitesUri = "https://graph.microsoft.com/v1.0/sites/$($parent.id)/sites`?$select=id,displayName,webUrl&`$top=200"
            do {
                $subResp = Invoke-GraphGet -Uri $subSitesUri -Headers $script:AppOnlyHeaders
                $subResp.value | Where-Object { $_.id } | ForEach-Object { $subSitesList.Add($_) }
                $subSitesUri = $subResp.'@odata.nextLink'
            } while ($subSitesUri)
            $subSites = @($subSitesList)
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

# ── Add explicitly requested OneDrive personal sites ─────────────────────────
if ($IncludeOneDriveUsers.Count -gt 0) {
    Write-ProgressHost -Message "Resolving OneDrive sites for specified users..." -ForegroundColor Cyan
    foreach ($upn in $IncludeOneDriveUsers) {
        try {
            $idsUri = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($upn))/drive/root/sharepointIds"

            if ($script:AppOnlyHeaders) {
                $ids = Invoke-GraphGet -Uri $idsUri -Headers $script:AppOnlyHeaders
            } else {
                $ids = Invoke-MgGraphRequest -Method GET -Uri $idsUri -OutputType PSObject -ErrorAction Stop
            }

            $odUri    = [System.Uri]$ids.siteUrl
            $odHost   = $odUri.Host
            $odPath   = $odUri.AbsolutePath.TrimEnd('/')
            $siteGraphUri = "https://graph.microsoft.com/v1.0/sites/${odHost}:${odPath}?`$select=id,displayName,webUrl,name"

            if ($script:AppOnlyHeaders) {
                $odSite = Invoke-GraphGet -Uri $siteGraphUri -Headers $script:AppOnlyHeaders
            } else {
                $odSite = Invoke-MgGraphRequest -Method GET -Uri $siteGraphUri -OutputType PSObject -ErrorAction Stop
            }

            if ($odSite -and $odSite.id -and $knownSiteIds.Add($odSite.id)) {
                $sites.Add($odSite) | Out-Null
                Write-ProgressHost -Message ("[OK] Added OneDrive: {0} ({1})" -f $upn, $odSite.webUrl) -ForegroundColor DarkGray
            } else {
                Write-ProgressHost -Message ("[INFO] OneDrive for {0} is already in the scan list." -f $upn) -ForegroundColor DarkGray
            }
        } catch {
            Write-ProgressHost -Message ("[WARN] Could not resolve OneDrive for {0}: {1}" -f $upn, $_.Exception.Message) -ForegroundColor Yellow
        }
    }
}

$oneDriveNote = if ($IncludeOneDriveUsers.Count -gt 0) { ", $($IncludeOneDriveUsers.Count) OneDrive user(s) included" } else { ', OneDrive excluded' }
Write-ProgressHost -Message ("Found {0} site(s) (site collections + sub-sites included{1})" -f $sites.Count, $oneDriveNote) -ForegroundColor Green
Write-Host ""

# ── Helpers ───────────────────────────────────────────────────────────────────
function Test-IsFile {
    param([object]$Item)
    return $null -ne $Item.file
}

function Get-SiteCollectionKey {
    # Resolves a site's webUrl to its root site collection URL, so sub-sites and
    # classic sub-webs can be grouped back with the root they share a storage quota
    # with. Private/shared Teams channels are genuine separate site collections
    # (own /sites/<name> or /teams/<name> segment) and correctly form their own key —
    # only extra path segments beyond that (classic sub-webs) collapse into the root.
    param([string]$WebUrl)
    if ($WebUrl -match '^(https://[^/]+(?:/sites/[^/]+|/teams/[^/]+)?)') {
        return $Matches[1].TrimEnd('/')
    }
    return $WebUrl.TrimEnd('/')
}

function Get-SiteDrives {
    # Uses /lists?$expand=drive to return ALL document libraries per site,
    # including Site Pages, Site Assets, Teams channels, and custom libraries
    # that may not surface in the /drives endpoint.
    param([string]$SiteId)
    if ($script:AppOnlyHeaders) {
        try {
            Update-AppOnlyToken
            $drives  = [System.Collections.Generic.List[object]]::new()
            $listUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists" +
                       '?$select=id,displayName,list&$expand=drive($select=id,name,webUrl,quota)&$top=200'
            do {
                $resp = Invoke-GraphGet -Uri $listUri -Headers $script:AppOnlyHeaders
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
        } catch {
            # Some tenants/sites return 400 on /lists?$expand=drive in app-only mode.
            # Fallback 1: enumerate document-library lists, then resolve /lists/{id}/drive.
            # This preserves broader coverage versus using /drives directly.
            try {
                Update-AppOnlyToken
                $drives = [System.Collections.Generic.List[object]]::new()
                $seenDriveIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                $listUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists" +
                           '?$select=id,displayName,list&$top=200'

                do {
                    $resp = Invoke-GraphGet -Uri $listUri -Headers $script:AppOnlyHeaders
                    $docLibLists = @(
                        $resp.value | Where-Object {
                            $_.id -and $_.list -and
                            # Include all list templates that can have an associated drive.
                            # documentLibrary is the most common, but picture libraries, form
                            # libraries and wiki libraries also expose a drive endpoint.
                            $_.list.template -in @(
                                'documentLibrary',
                                'pictureLibrary',
                                'webPageLibrary',
                                'htmlFormLibrary',
                                'homePageLibrary',
                                'assetLibrary'
                            )
                        }
                    )

                    foreach ($list in $docLibLists) {
                        try {
                            $driveObj = Invoke-GraphGet `
                                -Uri ("https://graph.microsoft.com/v1.0/sites/$SiteId/lists/{0}/drive" -f $list.id) `
                                -Headers $script:AppOnlyHeaders

                            if ($driveObj.id -and $seenDriveIds.Add($driveObj.id)) {
                                $driveObj | Add-Member -NotePropertyName 'VersioningEnabled' -NotePropertyValue $list.list.enableVersioning  -Force -ErrorAction SilentlyContinue
                                $driveObj | Add-Member -NotePropertyName 'MajorVersionLimit'  -NotePropertyValue $list.list.majorVersionLimit -Force -ErrorAction SilentlyContinue
                                $drives.Add($driveObj)
                            }
                        } catch {
                            # Not every list exposes a drive endpoint in every tenant configuration.
                        }
                    }

                    $listUri = $resp.'@odata.nextLink'
                } while ($listUri)

                if ($drives.Count -gt 0) {
                    return $drives
                }
            } catch {
                # Continue to final fallback below.
            }

            # Fallback 2: delegated Graph REST call (module-agnostic).
            if ($script:ConnectedHere) {
                $drives = [System.Collections.Generic.List[object]]::new()
                $drivesUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives?`$select=id,name,webUrl,quota&`$top=200"
                do {
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $drivesUri -OutputType PSObject -ErrorAction Stop
                    @($resp.value) | ForEach-Object { $drives.Add($_) }
                    $drivesUri = $resp.'@odata.nextLink'
                } while ($drivesUri)
                return $drives
            }
            throw
        }
    } else {
        $drives = [System.Collections.Generic.List[object]]::new()
        $drivesUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives?`$select=id,name,webUrl,quota&`$top=200"
        do {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $drivesUri -OutputType PSObject -ErrorAction Stop
            @($resp.value) | ForEach-Object { $drives.Add($_) }
            $drivesUri = $resp.'@odata.nextLink'
        } while ($drivesUri)
        return $drives
    }
}

function Get-VersionsNextPage {
    # Follows @odata.nextLink for a single file's version list. The versions endpoint normally
    # returns everything in one page, but if a library ever has enough version history to trigger
    # paging, we must not silently drop the older versions on later pages.
    param([string]$NextLink, [System.Collections.Generic.List[object]] $Values)
    $link = $NextLink
    while ($link) {
        $page = if ($script:AppOnlyHeaders) { Invoke-GraphGet -Uri $link -Headers $script:AppOnlyHeaders }
                else { Invoke-MgGraphRequest -Method GET -Uri $link -OutputType PSObject -ErrorAction Stop }
        if ($page.value) { $Values.AddRange(@($page.value)) }
        $link = $page.'@odata.nextLink'
    }
}

function Get-BatchItemRetryDelaySeconds {
    # "activityLimitReached" (SharePoint/OneDrive resource-level quota, surfaced as HTTP 429) needs
    # a much longer cooldown than ordinary throttling — it's a rolling quota against the same
    # site/list, so retrying quickly just keeps re-tripping it. Honor a Retry-After header when
    # Graph gives us one; otherwise back off hard specifically for activityLimitReached, and more
    # gently for plain 429/5xx.
    param([int]$Pass, [object]$SubResponse)
    $retryAfter = $null
    try {
        if ($SubResponse -and $SubResponse.headers) {
            $h = $SubResponse.headers.'Retry-After'
            if ($h) { [void][int]::TryParse([string]$h, [ref]$retryAfter) }
        }
    } catch {}
    if ($retryAfter -and $retryAfter -gt 0) { return [Math]::Min($retryAfter + 2, 180) }

    $errorCode = $null
    try { $errorCode = $SubResponse.body.error.code } catch {}
    if ($errorCode -eq 'activityLimitReached') { return [Math]::Min(30 * $Pass, 180) }

    return [Math]::Min(5 * $Pass, 30)
}

function Invoke-GraphBatchGet {
    # Resolves a set of GET requests via Microsoft Graph's $batch endpoint (max 20 per call).
    # Individual sub-requests inside an otherwise-successful batch response can come back
    # throttled (429) or with a transient server error under sustained load (thousands of files).
    # Those are retried in subsequent passes with backoff instead of being treated as a permanent
    # "0 versions" — the previous single-attempt-per-chunk approach was silently discarding the
    # vast majority of version history on large libraries, which is why tenant-wide version totals
    # could come out lower than a single site's real Storage Metrics usage.
    param(
        [System.Collections.Generic.List[object]] $Requests   # each: @{ Id = 'string'; Url = '/relative/path' }
    )

    $results = @{}
    if ($Requests.Count -eq 0) { return $results }

    $pending = $Requests
    $maxPasses = 8
    for ($pass = 1; $pass -le $maxPasses -and $pending.Count -gt 0; $pass++) {
        $chunkSpecs = [System.Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $pending.Count; $i += 20) {
            $end   = [Math]::Min($i + 19, $pending.Count - 1)
            $chunk = $pending.GetRange($i, $end - $i + 1)

            $chunkSpecs.Add([PSCustomObject]@{
                Requests = $chunk
                Body     = (@{
                    requests = @($chunk | ForEach-Object { @{ id = $_.Id; method = 'GET'; url = $_.Url } })
                } | ConvertTo-Json -Depth 6)
            }) | Out-Null
        }

        $retryList = [System.Collections.Generic.List[object]]::new()
        $nextDelay = 0

        if ($script:AppOnlyHeaders -and $VersionBatchConcurrency -gt 1 -and $chunkSpecs.Count -gt 1) {
            # Runspace workers get a plain copy of the bearer token and cannot see later updates to
            # $script:AppOnlyHeaders, so refresh it here before dispatch. Without this, a token that
            # expires mid-scan silently breaks every subsequent parallel version lookup for the rest
            # of the run.
            Update-AppOnlyToken
            $poolSize = [Math]::Min($VersionBatchConcurrency, $chunkSpecs.Count)
            $runspacePool = [RunspaceFactory]::CreateRunspacePool(1, $poolSize)
            $runspacePool.Open()

            $workers = [System.Collections.Generic.List[object]]::new()
            $workerScript = {
                param(
                    [string]$BatchBody,
                    [hashtable]$Headers,
                    [int]$TimeoutSec
                )

                for ($attempt = 1; $attempt -le 3; $attempt++) {
                    try {
                        $resp = Invoke-RestMethod -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' `
                            -Headers $Headers -ContentType 'application/json' `
                            -Body $BatchBody -TimeoutSec $TimeoutSec -ErrorAction Stop

                        return [PSCustomObject]@{
                            Success   = $true
                            Responses = @($resp.responses)
                        }
                    } catch {
                        if ($attempt -eq 3) {
                            return [PSCustomObject]@{
                                Success      = $false
                                Responses    = @()
                                ErrorMessage = $_.Exception.Message
                            }
                        }

                        Start-Sleep -Seconds ($attempt * 3)
                    }
                }
            }

            try {
                foreach ($spec in $chunkSpecs) {
                    # Stagger dispatch slightly — the pool size already caps true concurrency, but
                    # spacing out when each request starts further reduces burst rate against the
                    # same site/list, which is what activityLimitReached actually tracks.
                    Start-Sleep -Milliseconds 75

                    $ps = [PowerShell]::Create()
                    $ps.RunspacePool = $runspacePool
                    [void]$ps.AddScript($workerScript)
                    [void]$ps.AddParameter('BatchBody', $spec.Body)
                    [void]$ps.AddParameter('Headers', $script:AppOnlyHeaders)
                    [void]$ps.AddParameter('TimeoutSec', $GraphTimeoutSec)

                    $workers.Add([PSCustomObject]@{
                        PowerShell = $ps
                        Handle     = $ps.BeginInvoke()
                        Requests   = $spec.Requests
                    }) | Out-Null
                }

                foreach ($worker in $workers) {
                    $payload = $null

                    try {
                        $payload = $worker.PowerShell.EndInvoke($worker.Handle)
                    } catch {
                        $payload = $null
                    } finally {
                        $worker.PowerShell.Dispose()
                    }

                    if ($payload -and $payload.Success) {
                        $byId = @{}
                        foreach ($r in $payload.Responses) { $byId[[string]$r.id] = $r }

                        foreach ($req in $worker.Requests) {
                            $r = $byId[[string]$req.Id]
                            if (-not $r) {
                                if ($pass -lt $maxPasses) { $retryList.Add($req) } else { $results[$req.Id] = "No response for request id in batch." }
                                continue
                            }
                            if ($r.status -eq 200) {
                                try {
                                    $values = [System.Collections.Generic.List[object]]::new(@($r.body.value))
                                    if ($r.body.'@odata.nextLink') {
                                        Get-VersionsNextPage -NextLink $r.body.'@odata.nextLink' -Values $values
                                    }
                                    $results[[string]$r.id] = @{ value = $values }
                                } catch {
                                    # Pagination follow-up hit a hard failure (e.g. sustained throttling) —
                                    # fail just this one file's lookup instead of crashing the whole scan.
                                    if ($pass -lt $maxPasses) { $retryList.Add($req) } else { $results[$req.Id] = "Version page fetch failed: $($_.Exception.Message)" }
                                }
                            } elseif ($r.status -in @(429, 500, 502, 503, 504) -and $pass -lt $maxPasses) {
                                $retryList.Add($req)
                                $nextDelay = [Math]::Max($nextDelay, (Get-BatchItemRetryDelaySeconds -Pass $pass -SubResponse $r))
                            } else {
                                $errBody = try { $r.body | ConvertTo-Json -Compress -Depth 4 } catch { [string]$r.body }
                                $results[[string]$req.Id] = "HTTP $($r.status): $errBody"
                            }
                        }
                        continue
                    }

                    if ($pass -lt $maxPasses) {
                        foreach ($req in $worker.Requests) { $retryList.Add($req) }
                    } else {
                        $reason = if ($payload -and $payload.ErrorMessage) { $payload.ErrorMessage } else { 'Batch call failed after retries.' }
                        foreach ($req in $worker.Requests) { $results[$req.Id] = $reason }
                    }
                }
            } finally {
                $runspacePool.Close()
                $runspacePool.Dispose()
            }
        } else {
            foreach ($spec in $chunkSpecs) {
                $chunk = $spec.Requests
                $batchBody = $spec.Body

                $batchDone = $false
                for ($attempt = 1; $attempt -le 3 -and -not $batchDone; $attempt++) {
                    try {
                        if ($script:AppOnlyHeaders) {
                            Update-AppOnlyToken
                            $resp = Invoke-RestMethod -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' `
                                -Headers $script:AppOnlyHeaders -ContentType 'application/json' `
                                -Body $batchBody -TimeoutSec $GraphTimeoutSec -ErrorAction Stop
                        } else {
                            $resp = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' `
                                -Body $batchBody -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop
                        }
                        $byId = @{}
                        foreach ($r in $resp.responses) { $byId[[string]$r.id] = $r }

                        foreach ($req in $chunk) {
                            $r = $byId[[string]$req.Id]
                            if (-not $r) {
                                if ($pass -lt $maxPasses) { $retryList.Add($req) } else { $results[$req.Id] = "No response for request id in batch." }
                                continue
                            }
                            if ($r.status -eq 200) {
                                try {
                                    $values = [System.Collections.Generic.List[object]]::new(@($r.body.value))
                                    if ($r.body.'@odata.nextLink') {
                                        Get-VersionsNextPage -NextLink $r.body.'@odata.nextLink' -Values $values
                                    }
                                    $results[[string]$r.id] = @{ value = $values }
                                } catch {
                                    if ($pass -lt $maxPasses) { $retryList.Add($req) } else { $results[$req.Id] = "Version page fetch failed: $($_.Exception.Message)" }
                                }
                            } elseif ($r.status -in @(429, 500, 502, 503, 504) -and $pass -lt $maxPasses) {
                                $retryList.Add($req)
                            } else {
                                $errBody = try { $r.body | ConvertTo-Json -Compress -Depth 4 } catch { [string]$r.body }
                                $results[[string]$req.Id] = "HTTP $($r.status): $errBody"
                            }
                        }
                        $batchDone = $true
                    } catch {
                        if ($attempt -eq 3) {
                            if ($pass -lt $maxPasses) { foreach ($req in $chunk) { $retryList.Add($req) } }
                            else { foreach ($req in $chunk) { $results[$req.Id] = "Batch call failed: $($_.Exception.Message)" } }
                        } else {
                            Start-Sleep -Seconds ($attempt * 3)
                        }
                    }
                }
            }
        }

        if ($retryList.Count -gt 0 -and $pass -lt $maxPasses) {
            Start-Sleep -Seconds ([Math]::Min(5 * $pass, 30))
        }
        $pending = $retryList
    }
    foreach ($req in $pending) { $results[[string]$req.Id] = "Gave up after $maxPasses retry pass(es)." }

    return $results
}

function Get-AllDriveItems {
    param(
        [string]$DriveId,
        [bool]$FetchVersions = $true
    )

    # Iterative breadth-first traversal — no call stack limit, handles any folder depth.
    # File records are created with zeroed version fields; if $FetchVersions is set, their
    # version history is resolved afterwards in batches of 20 via Invoke-GraphBatchGet
    # instead of one Graph call per file during the traversal.
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $queue   = [System.Collections.Generic.Queue[PSCustomObject]]::new()
    $pendingVersions  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $processedFolders = 0
    $processedFiles   = 0

    $queue.Enqueue([PSCustomObject]@{ Id = 'root'; Path = '' })

    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        $processedFolders++
        if ($processedFolders % 25 -eq 0) {
            Write-ProgressHost -Message (
                "progress: {0} folders, {1} files scanned..." -f
                $processedFolders,
                $processedFiles
            ) -ForegroundColor DarkGray
        }

        # Collect all children (paginated)
        $children = [System.Collections.Generic.List[object]]::new()
        try {
            if ($script:AppOnlyHeaders) {
                $childUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$($current.Id)/children" +
                            '?$select=id,name,size,file,folder,lastModifiedDateTime&$top=200'
                do {
                    $resp = Invoke-GraphGet -Uri $childUri -Headers $script:AppOnlyHeaders
                    $resp.value | ForEach-Object { $children.Add($_) }
                    $childUri = $resp.'@odata.nextLink'
                } while ($childUri)
            } else {
                $childUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$($current.Id)/children" +
                            '?$select=id,name,size,file,folder,lastModifiedDateTime&$top=200'
                do {
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $childUri -OutputType PSObject -ErrorAction Stop
                    @($resp.value) | ForEach-Object { $children.Add($_) }
                    $childUri = $resp.'@odata.nextLink'
                } while ($childUri)
            }
        } catch {
            Write-Host ("          [ERROR] Cannot read folder '{0}': {1}" -f $current.Path, $_.Exception.Message) -ForegroundColor Red
            continue
        }

        foreach ($child in $children) {
            $path = if ($current.Path) { "$($current.Path)/$($child.name)" } else { $child.name }

            if (Test-IsFile -Item $child) {
                $processedFiles++
                $fileSize = [int64]($child.size ?? 0)

                $fileRecord = [PSCustomObject]@{
                    ItemType         = 'File'
                    Path             = $path
                    Level            = (($path -split '/').Count)
                    ParentPath       = $(if ($path -match '/') { ($path -replace '/[^/]+$','') } else { '/' })
                    SizeBytes        = $fileSize
                    SizeMB           = [math]::Round($fileSize / 1MB, 3)
                    VersionCount     = 0
                    VersionSizeBytes = [int64]0
                    VersionSizeMB    = 0.0
                    TotalSizeBytes   = $fileSize
                    TotalSizeMB      = [math]::Round($fileSize / 1MB, 3)
                    Modified         = $child.lastModifiedDateTime
                }
                $results.Add($fileRecord) | Out-Null

                if ($FetchVersions) {
                    $pendingVersions.Add([PSCustomObject]@{ ItemId = $child.id; Record = $fileRecord }) | Out-Null
                }
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

    # ── Resolve version history in batches of 20 (Graph $batch limit) ──────────
    if ($pendingVersions.Count -gt 0) {
        Write-ProgressHost -Message ("resolving version history for {0} file(s)..." -f $pendingVersions.Count) -ForegroundColor DarkGray

        $batchRequests = [System.Collections.Generic.List[object]]::new()
        $byId  = @{}
        $reqId = 0
        foreach ($pending in $pendingVersions) {
            $reqId++
            $rid = [string]$reqId
            $byId[$rid] = $pending.Record
            $batchRequests.Add([PSCustomObject]@{
                Id  = $rid
                Url = "/drives/$DriveId/items/$($pending.ItemId)/versions?`$select=id,size"
            }) | Out-Null
        }

        $batchResults = Invoke-GraphBatchGet -Requests $batchRequests

        foreach ($rid in $byId.Keys) {
            $record = $byId[$rid]
            $body   = $batchResults[$rid]
            if ($body -and $body.value) {
                $verSize = [int64](($body.value | Where-Object { $_.size } | Measure-Object -Property size -Sum).Sum ?? 0)
                $record.VersionCount     = $body.value.Count
                $record.VersionSizeBytes = $verSize
                $record.VersionSizeMB    = [math]::Round($verSize / 1MB, 3)
                $record.TotalSizeBytes   = $record.SizeBytes + $verSize
                $record.TotalSizeMB      = [math]::Round($record.TotalSizeBytes / 1MB, 3)
            } else {
                # Lookup failed after retries — record keeps its zeroed version defaults. Track the
                # reason so the run summary can surface *why*, instead of silently under-reporting.
                $script:VersionLookupFailureCount++
                if ($body -is [string] -and $script:VersionLookupFailureSamples.Count -lt 20) {
                    $script:VersionLookupFailureSamples.Add($body)
                }
            }
        }
    }

    return $results
}

function Get-SpoAppOnlyTokenForHost {
    param([string]$HostName)

    if ([string]::IsNullOrWhiteSpace($HostName)) { return $null }

    if ($script:SpoHostTokenCache.ContainsKey($HostName)) {
        return $script:SpoHostTokenCache[$HostName]
    }

    if (-not $script:TokenBody -or -not $script:TokenTenantId) {
        return $null
    }

    $body = @{
        grant_type    = 'client_credentials'
        scope         = "https://$HostName/.default"
        client_id     = $script:TokenBody.client_id
        client_secret = $script:TokenBody.client_secret
    }

    try {
        $resp = Invoke-RestMethod -Method POST -ErrorAction Stop `
            -Uri  "https://login.microsoftonline.com/$($script:TokenTenantId)/oauth2/v2.0/token" `
            -Body $body

        if ($resp.access_token) {
            $script:SpoHostTokenCache[$HostName] = $resp.access_token
            return $resp.access_token
        }
    } catch {}

    return $null
}

function Get-SpoResponseRows {
    param([object]$Response)

    if ($null -eq $Response) { return @() }
    if ($Response.value) { return @($Response.value) }
    if ($Response.d -and $Response.d.results) { return @($Response.d.results) }
    return @()
}

function Get-AllSpoLibraryItems {
    param(
        [string]$SiteWebUrl,
        [string]$RootFolderServerRelativeUrl,
        [bool]$FetchVersions = $true
    )

    if ([string]::IsNullOrWhiteSpace($SiteWebUrl)) {
        throw 'SPO REST scan requires SiteWebUrl.'
    }
    if ([string]::IsNullOrWhiteSpace($RootFolderServerRelativeUrl)) {
        throw 'SPO REST scan requires RootFolderServerRelativeUrl.'
    }

    $siteUri  = [Uri]$SiteWebUrl
    $hostName = $siteUri.Host
    $spoToken = Get-SpoAppOnlyTokenForHost -HostName $hostName
    if (-not $spoToken) {
        throw 'Could not obtain SharePoint-scoped token for hidden library scan.'
    }

    $spoHeaders = @{
        Authorization = "Bearer $spoToken"
        Accept        = 'application/json;odata=nometadata'
    }

    $libraryRoot      = $RootFolderServerRelativeUrl.TrimEnd('/')
    $results          = [System.Collections.Generic.List[PSCustomObject]]::new()
    $queue            = [System.Collections.Generic.Queue[PSCustomObject]]::new()
    $pendingVersions  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $seenFolders      = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $processedFolders = 0
    $processedFiles   = 0

    $queue.Enqueue([PSCustomObject]@{ ServerRelativeUrl = $libraryRoot })
    $seenFolders.Add($libraryRoot) | Out-Null

    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        $processedFolders++
        if ($processedFolders % 25 -eq 0) {
            Write-ProgressHost -Message (
                "progress (SPO REST): {0} folders, {1} files scanned..." -f
                $processedFolders,
                $processedFiles
            ) -ForegroundColor DarkGray
        }

        $encodedFolderUrl = [Uri]::EscapeDataString($current.ServerRelativeUrl)

        $filesUri = "https://$hostName/_api/web/GetFolderByServerRelativeUrl('$encodedFolderUrl')/Files?`$select=Name,ServerRelativeUrl,TimeLastModified,Length&`$top=5000"
        do {
            $filesResp = Invoke-RestMethod -Method GET -Uri $filesUri -Headers $spoHeaders -TimeoutSec $GraphTimeoutSec -ErrorAction Stop
            foreach ($file in (Get-SpoResponseRows -Response $filesResp)) {
                $processedFiles++
                $fileSize = [int64]($file.Length ?? 0)
                $path = ([string]$file.ServerRelativeUrl).Substring($libraryRoot.Length).TrimStart('/')
                if ([string]::IsNullOrWhiteSpace($path)) { $path = $file.Name }

                $fileRecord = [PSCustomObject]@{
                    ItemType         = 'File'
                    Path             = $path
                    Level            = (($path -split '/').Count)
                    ParentPath       = $(if ($path -match '/') { ($path -replace '/[^/]+$','') } else { '/' })
                    SizeBytes        = $fileSize
                    SizeMB           = [math]::Round($fileSize / 1MB, 3)
                    VersionCount     = 0
                    VersionSizeBytes = [int64]0
                    VersionSizeMB    = 0.0
                    TotalSizeBytes   = $fileSize
                    TotalSizeMB      = [math]::Round($fileSize / 1MB, 3)
                    Modified         = $file.TimeLastModified
                }
                $results.Add($fileRecord) | Out-Null

                if ($FetchVersions) {
                    $pendingVersions.Add([PSCustomObject]@{
                        ServerRelativeUrl = [string]$file.ServerRelativeUrl
                        Record            = $fileRecord
                    }) | Out-Null
                }
            }

            if ($filesResp.'@odata.nextLink') {
                $filesUri = $filesResp.'@odata.nextLink'
            } elseif ($filesResp.d -and $filesResp.d.__next) {
                $filesUri = $filesResp.d.__next
            } else {
                $filesUri = $null
            }
        } while ($filesUri)

        $foldersUri = "https://$hostName/_api/web/GetFolderByServerRelativeUrl('$encodedFolderUrl')/Folders?`$select=Name,ServerRelativeUrl,TimeLastModified,ItemCount&`$top=5000"
        do {
            $foldersResp = Invoke-RestMethod -Method GET -Uri $foldersUri -Headers $spoHeaders -TimeoutSec $GraphTimeoutSec -ErrorAction Stop
            foreach ($folder in (Get-SpoResponseRows -Response $foldersResp)) {
                $folderUrl = [string]$folder.ServerRelativeUrl
                if ([string]::IsNullOrWhiteSpace($folderUrl) -or $folderUrl -eq $libraryRoot) { continue }

                $path = $folderUrl.Substring($libraryRoot.Length).TrimStart('/')
                if ([string]::IsNullOrWhiteSpace($path)) { continue }

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
                    Modified         = $folder.TimeLastModified
                }) | Out-Null

                if ($seenFolders.Add($folderUrl)) {
                    $queue.Enqueue([PSCustomObject]@{ ServerRelativeUrl = $folderUrl })
                }
            }

            if ($foldersResp.'@odata.nextLink') {
                $foldersUri = $foldersResp.'@odata.nextLink'
            } elseif ($foldersResp.d -and $foldersResp.d.__next) {
                $foldersUri = $foldersResp.d.__next
            } else {
                $foldersUri = $null
            }
        } while ($foldersUri)
    }

    if ($FetchVersions -and $pendingVersions.Count -gt 0) {
        Write-ProgressHost -Message ("resolving version history via SPO REST for {0} file(s)..." -f $pendingVersions.Count) -ForegroundColor DarkGray
        foreach ($pending in $pendingVersions) {
            try {
                $encodedFileUrl = [Uri]::EscapeDataString($pending.ServerRelativeUrl)
                $versionUri = "https://$hostName/_api/web/GetFileByServerRelativeUrl('$encodedFileUrl')/Versions?`$select=Size&`$top=5000"
                $versionRows = [System.Collections.Generic.List[object]]::new()

                do {
                    $versionResp = Invoke-RestMethod -Method GET -Uri $versionUri -Headers $spoHeaders -TimeoutSec $GraphTimeoutSec -ErrorAction Stop
                    foreach ($row in (Get-SpoResponseRows -Response $versionResp)) {
                        $versionRows.Add($row) | Out-Null
                    }

                    if ($versionResp.'@odata.nextLink') {
                        $versionUri = $versionResp.'@odata.nextLink'
                    } elseif ($versionResp.d -and $versionResp.d.__next) {
                        $versionUri = $versionResp.d.__next
                    } else {
                        $versionUri = $null
                    }
                } while ($versionUri)

                if ($versionRows.Count -gt 0) {
                    $verSize = [int64](($versionRows | Where-Object { $_.Size } | Measure-Object -Property Size -Sum).Sum ?? 0)
                    $pending.Record.VersionCount     = $versionRows.Count
                    $pending.Record.VersionSizeBytes = $verSize
                    $pending.Record.VersionSizeMB    = [math]::Round($verSize / 1MB, 3)
                    $pending.Record.TotalSizeBytes   = $pending.Record.SizeBytes + $verSize
                    $pending.Record.TotalSizeMB      = [math]::Round($pending.Record.TotalSizeBytes / 1MB, 3)
                }
            } catch {
                # Version lookup failed for this file — keep zeroed version defaults.
            }
        }
    }

    return $results
}

function Get-SiteRecycleBinItems {
    param(
        [string]$SiteId,
        [string]$SiteWebUrl
    )

    $items = [System.Collections.Generic.List[object]]::new()

    if ($script:AppOnlyHeaders) {
        # NOTE: Graph /sites/{id}/recycleBin/items is ONLY for SharePoint Embedded
        # fileStorageContainers and always returns 400 for regular SharePoint sites.
        # The correct API is SharePoint REST /_api/site/RecycleBin, which requires
        # a SharePoint-scoped token (not a Graph token).
        # SPO REST is therefore tried first; Graph is only kept as a last-resort fallback.

        # Primary: SharePoint REST with an app-only token for the specific site host.
        if ($SiteWebUrl) {
            try {
                $siteUri  = [Uri]$SiteWebUrl
                $hostName = $siteUri.Host
                $basePath = $siteUri.AbsolutePath.TrimEnd('/')
                if ($basePath -eq '/') { $basePath = '' }

                $spoToken = Get-SpoAppOnlyTokenForHost -HostName $hostName
                if ($spoToken) {
                    $spoHeaders = @{
                        Authorization = "Bearer $spoToken"
                        Accept        = 'application/json;odata=nometadata'
                    }

                    $spoUri = "https://$hostName$basePath/_api/site/RecycleBin?`$select=Id,Title,Size,DeletedDate,ItemState&`$top=5000"
                    do {
                        $resp = Invoke-RestMethod -Method GET -Uri $spoUri -Headers $spoHeaders -TimeoutSec $GraphTimeoutSec -ErrorAction Stop

                        $rows = @()
                        if ($resp.value) {
                            $rows = @($resp.value)
                        } elseif ($resp.d -and $resp.d.results) {
                            $rows = @($resp.d.results)
                        }

                        foreach ($row in $rows) {
                            $items.Add([PSCustomObject]@{
                                id              = $row.Id
                                name            = $row.Title
                                size            = [int64]($row.Size ?? 0)
                                deletedDateTime = $row.DeletedDate
                            }) | Out-Null
                        }

                        if ($resp.'@odata.nextLink') {
                            $spoUri = $resp.'@odata.nextLink'
                        } elseif ($resp.d -and $resp.d.__next) {
                            $spoUri = $resp.d.__next
                        } else {
                            $spoUri = $null
                        }
                    } while ($spoUri)

                    return $items
                }
            } catch {}
        }

        # Final fallback: Graph REST recycle bin endpoint (delegated/app-only behavior
        # depends on tenant/site type; if unsupported this call can return 400).
        try {
            $rbUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/recycleBin/items?`$select=id,name,size,deletedDateTime&`$top=200"
            do {
                $resp = Invoke-MgGraphRequest -Method GET -Uri $rbUri -OutputType PSObject -ErrorAction Stop
                @($resp.value) | ForEach-Object { $items.Add($_) }
                $rbUri = $resp.'@odata.nextLink'
            } while ($rbUri)

            if ($items.Count -gt 0) {
                return $items
            }
        } catch {}

        throw "Recycle bin: all retrieval methods failed for this site."
    }

    $rbUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/recycleBin/items?`$select=id,name,size,deletedDateTime&`$top=200"
    do {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $rbUri -OutputType PSObject -ErrorAction Stop
        @($resp.value) | ForEach-Object { $items.Add($_) }
        $rbUri = $resp.'@odata.nextLink'
    } while ($rbUri)

    return $items
}

# ── Resume / checkpoint state ────────────────────────────────────────────────
$checkpointSignature = Get-TextHashHex -Text (@"
$PSCommandPath
$outputDir
$SiteUrl
$SkipVersions
$Apply
$UseHighPrivilege
$RecycleBinOnly
$ForceAppOnlySingleSite
$($IncludeOneDriveUsers -join ',')
$TenantId
$effectiveTenantId
$ClientId
$CertificateThumbprint
$GraphTimeoutSec
$MaxGraphRetry
$VersionBatchConcurrency
"@)
$script:CheckpointStatePath   = Join-Path $outputDir "SharePoint_StorageReport_$checkpointSignature.state.json"
$script:CheckpointSummaryPath = Join-Path $outputDir "SharePoint_StorageReport_$checkpointSignature.summary.partial.csv"
$script:CheckpointDetailPath  = Join-Path $outputDir "SharePoint_StorageReport_$checkpointSignature.detail.partial.csv"
$script:CompletedLibraryKeys  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$script:LoadedCheckpoint      = $false

if (Test-Path $script:CheckpointStatePath) {
    try {
        $checkpointState = Get-Content -Path $script:CheckpointStatePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($checkpointState.RunSignature -eq $checkpointSignature) {
            $script:LoadedCheckpoint = $true
            foreach ($key in @($checkpointState.CompletedLibraryKeys)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$key)) {
                    $script:CompletedLibraryKeys.Add([string]$key) | Out-Null
                }
            }
        }
    } catch {
        $script:LoadedCheckpoint = $false
    }
}

function Save-CheckpointState {
    param(
        [System.Collections.Generic.HashSet[string]]$CompletedLibraryKeys
    )

    $state = [PSCustomObject]@{
        Version              = 1
        RunSignature         = $checkpointSignature
        UpdatedUtc           = (Get-Date).ToUniversalTime().ToString('o')
        CompletedLibraryKeys = @($CompletedLibraryKeys | Sort-Object)
    }

    $state | ConvertTo-Json -Depth 4 | Set-Content -Path $script:CheckpointStatePath -Encoding UTF8
}

function Finalize-CheckpointFiles {
    foreach ($path in @($script:CheckpointStatePath, $script:CheckpointSummaryPath, $script:CheckpointDetailPath)) {
        if (Test-Path $path) {
            try { Remove-Item -Path $path -Force -ErrorAction Stop } catch {}
        }
    }
}

function Convert-CheckpointCsvRows {
    param([string]$Path)

    if (-not (Test-Path $Path)) { return @() }
    return @(Import-Csv -Path $Path -ErrorAction Stop)
}

function Complete-CheckpointUnit {
    param(
        [string]$CheckpointKey,
        [object[]]$SummaryRows,
        [object[]]$DetailRows
    )

    if ($SummaryRows.Count -gt 0) {
        Append-CheckpointRows -Path $script:CheckpointSummaryPath -Rows $SummaryRows
    }
    if ($DetailRows.Count -gt 0) {
        Append-CheckpointRows -Path $script:CheckpointDetailPath -Rows $DetailRows
    }

    if (-not [string]::IsNullOrWhiteSpace($CheckpointKey)) {
        $script:CompletedLibraryKeys.Add($CheckpointKey) | Out-Null
    }
    Save-CheckpointState -CompletedLibraryKeys $script:CompletedLibraryKeys
}

$summaryRows = [System.Collections.Generic.List[PSCustomObject]]::new()
$detailRows  = [System.Collections.Generic.List[PSCustomObject]]::new()

if ($script:LoadedCheckpoint) {
    foreach ($row in (Convert-CheckpointCsvRows -Path $script:CheckpointSummaryPath)) {
        $summaryRows.Add($row) | Out-Null
    }
    foreach ($row in (Convert-CheckpointCsvRows -Path $script:CheckpointDetailPath)) {
        $detailRows.Add($row) | Out-Null
    }
    Write-ProgressHost -Message ("Resuming with {0} completed library checkpoint(s)." -f $script:CompletedLibraryKeys.Count) -ForegroundColor DarkGray
}

if ($RecycleBinOnly) {
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "   Phase 1: Recycle bins" -ForegroundColor Cyan
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

        if ([string]::IsNullOrWhiteSpace([string]$siteId)) { continue }

        # Only root site collections have their own recycle bin.
        $isRootSiteCollection = $site.webUrl -match '^https://[^/]+(/sites/[^/]+|/teams/[^/]+|/personal/[^/]+)?/?$'
        if (-not $isRootSiteCollection) { continue }

        if (-not $processedRbSiteIds.Add($siteId)) { continue }

        Write-Host ("  Recycle bin: {0}" -f $siteName) -ForegroundColor White

        try {
            $rbItems = Get-SiteRecycleBinItems -SiteId $siteId -SiteWebUrl $site.webUrl

            $rbSizeBytes = [int64](($rbItems | Where-Object { $_.size } |
                               Measure-Object -Property size -Sum).Sum ?? 0)
            $rbCount     = $rbItems.Count

            Write-Host ("        {0} item(s) | {1}" -f
                $rbCount, (Format-SizeAuto -MB ($rbSizeBytes / 1MB))) -ForegroundColor DarkGray

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

    Write-Host ""
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "   Exporting results" -ForegroundColor Cyan
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host ""

    $summaryRows = @(
        $summaryRows | Sort-Object `
            @{ Expression = { if ($null -ne $_.TotalSizeMB) { [double]$_.TotalSizeMB } else { -1 } }; Descending = $true },
            SiteName
    )
    $summaryRows | Export-Csv -Path $summaryCsv -NoTypeInformation -Encoding UTF8
    Write-Host ("  Summary  : {0}" -f $summaryCsv) -ForegroundColor Green

    if ($detailRows.Count -gt 0) {
        $detailRows = @(
            $detailRows | Sort-Object `
                @{ Expression = { if ($null -ne $_.Modified) { $_.Modified } else { [datetime]'1900-01-01' } }; Descending = $true },
                SiteName,
                Path
        )
        $detailRows | Export-Csv -Path $reportCsv -NoTypeInformation -Encoding UTF8
        Write-Host ("  Detail   : {0}" -f $reportCsv) -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "   Summary" -ForegroundColor Cyan
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host ("  Sites scanned : {0}" -f $sites.Count)
    $grandRB = ($summaryRows | Measure-Object -Property TotalSizeMB -Sum).Sum ?? 0
    $grandItems = ($summaryRows | Measure-Object -Property FileCount -Sum).Sum ?? 0
    Write-Host ("  Deleted items : {0}" -f $grandItems)
    Write-Host ("  Recycle bins  : {0}" -f (Format-SizeAuto -MB $grandRB)) -ForegroundColor Magenta
    Write-Host ""

    Finalize-CheckpointFiles
    Remove-TempApp
    return
}

# ── Phase 1: Enumerate all document libraries ─────────────────────────────────
Write-Host "  ================================================" -ForegroundColor Cyan
Write-ProgressHost -Message "Phase 1: Enumerating document libraries" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

$siteLibraries = [System.Collections.Generic.List[PSCustomObject]]::new()
$siteIndex     = 0

foreach ($site in $sites) {
    $siteIndex++
    $siteName = $site.displayName ?? $site.name
    $siteId   = $site.id

    Write-ProgressHost -Message ("[{0}/{1}] {2}" -f $siteIndex, $sites.Count, $siteName) -ForegroundColor White

    if ([string]::IsNullOrWhiteSpace([string]$siteId)) {
        Write-ProgressHost -Message "[WARN] Skipping site without valid id." -ForegroundColor Yellow
        continue
    }

    try {
        $drives = Get-SiteDrives -SiteId $siteId
        foreach ($drive in $drives) {
            $siteLibraries.Add([PSCustomObject]@{
                Site  = $site
                Drive = $drive
            }) | Out-Null
            Write-ProgressHost -Message ("{0}" -f $drive.name) -ForegroundColor DarkGray
        }

        # ── Hidden library scan via SPO REST ──────────────────────────────────
        # The Preservation Hold Library (and other hidden document libraries) are
        # created automatically by Microsoft Purview/Compliance retention policies.
        # They count toward the SharePoint storage quota shown in the admin portal,
        # but are NOT returned by the Graph /lists or /drives endpoints without
        # Sites.FullControl.All. The SPO REST /_api/web/lists endpoint with a
        # SharePoint-scoped token is the most reliable way to discover them.
        # Requires -UseHighPrivilege (Sites.FullControl.All) on the temp app, or a
        # provided app with FullControl, for the subsequent Graph drive lookup to succeed.
        if ($site.webUrl -and $script:TokenBody) {
            try {
                $hSiteUri  = [Uri]$site.webUrl
                $hHost     = $hSiteUri.Host
                $hBasePath = $hSiteUri.AbsolutePath.TrimEnd('/')
                if ($hBasePath -eq '/') { $hBasePath = '' }
                $hSpoToken = Get-SpoAppOnlyTokenForHost -HostName $hHost
                if ($hSpoToken) {
                    $hSpoHdrs  = @{
                        Authorization = "Bearer $hSpoToken"
                        Accept        = 'application/json;odata=nometadata'
                    }
                    # Enumerate all hidden document libraries and capture the root folder URL,
                    # so they can still be scanned via SPO REST when Graph won't expose a drive.
                    $hListUri  = "https://$hHost$hBasePath/_api/web/lists?`$filter=Hidden eq true and BaseTemplate eq 101&`$select=Id,Title,BaseTemplate,RootFolder/ServerRelativeUrl&`$expand=RootFolder&`$top=500"
                    $hListResp = Invoke-RestMethod -Method GET -Uri $hListUri -Headers $hSpoHdrs `
                                    -TimeoutSec $GraphTimeoutSec -ErrorAction Stop
                    # Build a set of drive IDs already found to avoid duplicates
                    $hKnownIds = [System.Collections.Generic.HashSet[string]]::new(
                        [string[]]@($drives | Where-Object { $_.id } | ForEach-Object { $_.id }),
                        [StringComparer]::OrdinalIgnoreCase
                    )
                    foreach ($hList in (Get-SpoResponseRows -Response $hListResp)) {
                        try {
                            $hDriveUri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$($hList.Id)/drive"
                            $hDrive    = if ($script:AppOnlyHeaders) {
                                Invoke-GraphGet -Uri $hDriveUri -Headers $script:AppOnlyHeaders
                            } else {
                                Invoke-MgGraphRequest -Method GET -Uri $hDriveUri -OutputType PSObject -ErrorAction Stop
                            }
                            if ($hDrive -and $hDrive.id -and $hKnownIds.Add($hDrive.id)) {
                                $hDrive | Add-Member -NotePropertyName 'VersioningEnabled' -NotePropertyValue $null -Force -ErrorAction SilentlyContinue
                                $hDrive | Add-Member -NotePropertyName 'MajorVersionLimit'  -NotePropertyValue $null -Force -ErrorAction SilentlyContinue
                                $siteLibraries.Add([PSCustomObject]@{
                                    Site  = $site
                                    Drive = $hDrive
                                }) | Out-Null
                                Write-ProgressHost -Message ("[hidden] {0}" -f $hList.Title) -ForegroundColor DarkYellow
                            }
                        } catch {
                            if ($hList.RootFolder -and $hList.RootFolder.ServerRelativeUrl) {
                                $pseudoDrive = [PSCustomObject]@{
                                    id                          = "spo-list|$($hList.Id)"
                                    name                        = $hList.Title
                                    webUrl                      = "https://$hHost$($hList.RootFolder.ServerRelativeUrl)"
                                    quota                       = $null
                                    VersioningEnabled           = $null
                                    MajorVersionLimit           = $null
                                    ScanMode                    = 'SpoRest'
                                    RootFolderServerRelativeUrl = $hList.RootFolder.ServerRelativeUrl
                                }
                                $siteLibraries.Add([PSCustomObject]@{
                                    Site  = $site
                                    Drive = $pseudoDrive
                                }) | Out-Null
                                Write-ProgressHost -Message ("[hidden-rest] {0}" -f $hList.Title) -ForegroundColor DarkYellow
                            }
                        }
                    }
                }
            } catch {
                # SPO REST hidden library scan failed — non-critical, standard libraries already collected.
            }
        }
    } catch {
        Write-ProgressHost -Message ("[ERROR] Cannot enumerate libraries: {0}" -f $_.Exception.Message) -ForegroundColor Red
    }
}

Write-Host ""
Write-ProgressHost -Message ("Found {0} document libraries across {1} site(s)" -f
    $siteLibraries.Count, $sites.Count) -ForegroundColor Green
Write-Host ""

# ── Phase 2: Retrieve storage data ────────────────────────────────────────────
Write-Host "  ================================================" -ForegroundColor Cyan
Write-ProgressHost -Message "Phase 2: Retrieving storage data" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

$libIndex    = 0

foreach ($entry in $siteLibraries) {
    $libIndex++
    $site     = $entry.Site
    $drive    = $entry.Drive
    $siteName = $site.displayName ?? $site.name
    $libraryKey = Get-LibraryCheckpointKey -SiteId $site.id -DriveId $drive.id

    Write-ProgressHost -Message ("[{0}/{1}] {2} > {3}" -f $libIndex, $siteLibraries.Count, $siteName, $drive.name) -ForegroundColor White

    if ($script:LoadedCheckpoint -and $script:CompletedLibraryKeys.Contains($libraryKey)) {
        Write-ProgressHost -Message "    [SKIP] Already completed in a previous run." -ForegroundColor DarkGray
        continue
    }

    $librarySummaryRows = [System.Collections.Generic.List[object]]::new()
    $libraryDetailRows  = [System.Collections.Generic.List[object]]::new()

    # Quick mode: use quota data from drives (no file enumeration)
    if (-not $Apply) {
        $quota = $drive.quota
        $librarySummaryRows.Add([PSCustomObject]@{
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
        Write-Host ("        used: {0}" -f (Format-SizeAuto -MB (($quota.used ?? 0) / 1MB))) -ForegroundColor DarkGray

        foreach ($row in $librarySummaryRows) { $summaryRows.Add($row) | Out-Null }
        Complete-CheckpointUnit -CheckpointKey $libraryKey -SummaryRows @($librarySummaryRows) -DetailRows @()
        continue
    }

    # Full scan mode
    # Skip per-file version lookups only when we positively know versioning is off for this
    # library — $drive.VersioningEnabled is $null (unknown) for the Get-MgSiteDrive fallback,
    # in which case we still fetch to avoid silently under-reporting.
    $isSpoRestLibrary   = ($drive.PSObject.Properties.Name -contains 'ScanMode') -and ($drive.ScanMode -eq 'SpoRest')
    $versioningKnownOff = ($null -ne $drive.VersioningEnabled) -and (-not [bool]$drive.VersioningEnabled)
    $fetchVersions      = (-not $SkipVersions) -and (-not $FastMode) -and (-not $versioningKnownOff)
    $includeDetailRows  = -not $FastMode
    if (-not $SkipVersions -and $versioningKnownOff) {
        Write-Host "        versioning disabled on this library — skipping version lookups" -ForegroundColor DarkGray
    }
    $items = if ($isSpoRestLibrary) {
        Get-AllSpoLibraryItems -SiteWebUrl $site.webUrl -RootFolderServerRelativeUrl $drive.RootFolderServerRelativeUrl -FetchVersions $fetchVersions
    } else {
        Get-AllDriveItems -DriveId $drive.id -FetchVersions $fetchVersions
    }

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

    $folderReportRows = if ($includeDetailRows) {
        @(
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
    } else {
        @()
    }
    $totalFiles  = $fileItems.Count
    $totalFolders = $folderReportRows.Count
    $currentSize = ($fileItems | Measure-Object -Property SizeBytes -Sum).Sum ?? 0
    $versionSize = ($fileItems | Measure-Object -Property VersionSizeBytes -Sum).Sum ?? 0
    $totalSize   = $currentSize + $versionSize
    $folderCountDisplay = if ($FastMode) { 'n/a' } else { $totalFolders }

    Write-Host ("        {0} folders | {1} files | current: {2} | versions: {3} | total: {4}" -f
        $folderCountDisplay,
        $totalFiles,
        (Format-SizeAuto -MB ($currentSize / 1MB)),
        (Format-SizeAuto -MB ($versionSize / 1MB)),
        (Format-SizeAuto -MB ($totalSize   / 1MB))) -ForegroundColor DarkGray

    $librarySummaryRows.Add([PSCustomObject]@{
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
        FolderCount       = if ($FastMode) { $null } else { $totalFolders }
        VersionSizeMB     = [math]::Round($versionSize / 1MB, 2)
        TotalSizeMB       = [math]::Round($totalSize   / 1MB, 2)
    }) | Out-Null

    $verEnabled = $drive.VersioningEnabled
    $verLimit   = if ($drive.MajorVersionLimit -eq 0) { 'Unlimited' } else { $drive.MajorVersionLimit }

    if ($includeDetailRows) {
        foreach ($item in $folderReportRows) {
            $libraryDetailRows.Add([PSCustomObject]@{
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
            $libraryDetailRows.Add([PSCustomObject]@{
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

    foreach ($row in $librarySummaryRows) { $summaryRows.Add($row) | Out-Null }
    foreach ($row in $libraryDetailRows) { $detailRows.Add($row) | Out-Null }
    Complete-CheckpointUnit -CheckpointKey $libraryKey -SummaryRows @($librarySummaryRows) -DetailRows @($libraryDetailRows)
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
        $rbKey    = "rb|$siteId"

        # Only root site collections have their own recycle bin.
        # Sub-webs (URLs with extra path segments beyond /sites/<name>) share the root's bin.
        # Root patterns: https://tenant.sharepoint.com  or  .../sites/name  or  .../teams/name  or  .../personal/name
        $isRootSiteCollection = $site.webUrl -match '^https://[^/]+(/sites/[^/]+|/teams/[^/]+|/personal/[^/]+)?/?$'
        if (-not $isRootSiteCollection) { continue }

        if ($script:LoadedCheckpoint -and $script:CompletedLibraryKeys.Contains($rbKey)) {
            Write-Host ("  Recycle bin: {0} [SKIP]" -f $siteName) -ForegroundColor DarkGray
            continue
        }

        if (-not $processedRbSiteIds.Add($siteId)) { continue }

        Write-Host ("  Recycle bin: {0}" -f $siteName) -ForegroundColor White

        $rbSummaryRows = [System.Collections.Generic.List[object]]::new()
        $rbDetailRows  = [System.Collections.Generic.List[object]]::new()

        try {
            $rbItems = Get-SiteRecycleBinItems -SiteId $siteId -SiteWebUrl $site.webUrl

            $rbSizeBytes = [int64](($rbItems | Where-Object { $_.size } |
                               Measure-Object -Property size -Sum).Sum ?? 0)
            $rbCount     = $rbItems.Count

            Write-Host ("        {0} item(s) | {1}" -f
                $rbCount, (Format-SizeAuto -MB ($rbSizeBytes / 1MB))) -ForegroundColor DarkGray

            $rbSummaryRows.Add([PSCustomObject]@{
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
                $rbDetailRows.Add([PSCustomObject]@{
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

            foreach ($row in $rbSummaryRows) { $summaryRows.Add($row) | Out-Null }
            foreach ($row in $rbDetailRows) { $detailRows.Add($row) | Out-Null }
            Complete-CheckpointUnit -CheckpointKey $rbKey -SummaryRows @($rbSummaryRows) -DetailRows @($rbDetailRows)

        } catch {
            Write-Host ("        [WARN] Cannot read recycle bin: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
        }
    }
}

# ── Phase 2c: Site collection totals ─────────────────────────────────────────
# Sub-sites and classic sub-webs share their root's storage quota in the SharePoint
# admin center, but are scanned here as separate site entries — and the recycle bin
# is tracked separately from the library scan. Neither is directly comparable to the
# single "storage used" figure the admin center shows per site collection. This phase
# regroups everything (all sub-sites' libraries + that collection's recycle bin) back
# to its root site collection so the grand total lines up 1:1 with the admin portal.
$siteCollectionRows = @()
if ($Apply) {
    $collectionMap = [System.Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)

    foreach ($site in $sites) {
        $key = Get-SiteCollectionKey -WebUrl $site.webUrl
        if (-not $collectionMap.ContainsKey($key)) {
            $collectionMap[$key] = [PSCustomObject]@{
                CollectionKey = $key
                DisplayName   = $null
                LibrariesMB   = 0.0
                RecycleBinMB  = 0.0
                FileCount     = 0
                LibraryCount  = 0
                SiteCount     = 0
            }
        }
        $entry = $collectionMap[$key]
        $entry.SiteCount++
        if ($site.webUrl.TrimEnd('/') -eq $key) {
            $entry.DisplayName = $site.displayName ?? $site.name
        }
    }

    foreach ($row in $summaryRows) {
        if ([string]::IsNullOrWhiteSpace([string]$row.SiteUrl)) { continue }
        $key = Get-SiteCollectionKey -WebUrl $row.SiteUrl
        if (-not $collectionMap.ContainsKey($key)) { continue }
        $entry = $collectionMap[$key]
        if ($row.Library -eq 'Recycle Bin (stage 1 + 2)') {
            $entry.RecycleBinMB += [double]($row.TotalSizeMB ?? 0)
        } else {
            $entry.LibrariesMB += [double]($row.TotalSizeMB ?? 0)
            $entry.FileCount   += [int]($row.FileCount ?? 0)
            $entry.LibraryCount++
        }
    }

    $siteCollectionRows = @(
        $collectionMap.Values | ForEach-Object {
            [PSCustomObject]@{
                SiteCollection    = $(if ($_.DisplayName) { $_.DisplayName } else { $_.CollectionKey })
                SiteCollectionUrl = $_.CollectionKey
                SubSiteCount      = $_.SiteCount
                LibraryCount      = $_.LibraryCount
                FileCount         = $_.FileCount
                LibrariesMB       = [math]::Round($_.LibrariesMB, 2)
                RecycleBinMB      = [math]::Round($_.RecycleBinMB, 2)
                GrandTotalMB      = [math]::Round($_.LibrariesMB + $_.RecycleBinMB, 2)
                GrandTotalGB      = [math]::Round(($_.LibrariesMB + $_.RecycleBinMB) / 1024, 3)
            }
        } | Sort-Object GrandTotalMB -Descending
    )

    $siteCollectionRows | Export-Csv -Path $collectionCsv -NoTypeInformation -Encoding UTF8
    Write-ProgressHost -Message ("Site collections : {0}" -f $collectionCsv) -ForegroundColor Green

    Write-Host ""
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-ProgressHost -Message "Top 10 site collections (libraries + recycle bin, sub-sites combined)" -ForegroundColor Cyan
    Write-Host "  ================================================" -ForegroundColor Cyan
    $siteCollectionRows | Select-Object -First 10 | ForEach-Object {
        Write-ProgressHost -Message ("{0}  [{1}]  (libraries: {2}, prullenbak: {3}, {4} sub-site(n)/kanalen)" -f
            (Format-SizeAuto -MB ($_.LibrariesMB + $_.RecycleBinMB)),
            $_.SiteCollection,
            (Format-SizeAuto -MB $_.LibrariesMB),
            (Format-SizeAuto -MB $_.RecycleBinMB),
            $_.SubSiteCount) -ForegroundColor Green
    }
}

# ── Export ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-ProgressHost -Message "Exporting results" -ForegroundColor Cyan
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

if ($Apply) {
    $grandFiles = ($summaryRows | Where-Object { $_.Library -ne 'Recycle Bin (stage 1 + 2)' } | Measure-Object -Property FileCount -Sum).Sum ?? 0
    $grandVer   = ($summaryRows | Where-Object { $_.Library -ne 'Recycle Bin (stage 1 + 2)' } | Measure-Object -Property VersionSizeMB -Sum).Sum ?? 0
    $grandRB    = ($summaryRows | Where-Object { $_.Library -eq 'Recycle Bin (stage 1 + 2)' } | Measure-Object -Property TotalSizeMB  -Sum).Sum ?? 0
    $grandTotal = ($summaryRows | Measure-Object -Property TotalSizeMB -Sum).Sum ?? 0
}

if (-not $Apply) {
    $summaryRows | Export-Csv -Path $summaryCsv -NoTypeInformation -Encoding UTF8
    Write-ProgressHost -Message ("Summary  : {0}" -f $summaryCsv) -ForegroundColor Green
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
    Write-ProgressHost -Message ("Ranked   : {0}" -f $reportCsv) -ForegroundColor Green

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
        $mdLines.Add(("| Versiedata | {0} |" -f (Format-SizeAuto -MB $grandVer)))
        $mdLines.Add(("| Totaal (huidig + versies) | {0} |" -f (Format-SizeAuto -MB $grandTotal)))
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
        Write-ProgressHost -Message ("Rapport  : {0}" -f $reportMd) -ForegroundColor Green
    }
}

Finalize-CheckpointFiles

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-ProgressHost -Message "Summary" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-ProgressHost -Message ("Sites scanned : {0}" -f $sites.Count)

if ($Apply) {
    Write-ProgressHost -Message ("Site collections : {0}" -f $siteCollectionRows.Count)
    Write-ProgressHost -Message ("Total files   : {0}"    -f $grandFiles)
    Write-ProgressHost -Message ("Version data  : {0}" -f (Format-SizeAuto -MB $grandVer)) -ForegroundColor Yellow
    Write-ProgressHost -Message ("Recycle bins  : {0}" -f (Format-SizeAuto -MB $grandRB)) -ForegroundColor Magenta
    Write-ProgressHost -Message ("Grand total   : {0}" -f (Format-SizeAuto -MB $grandTotal)) -ForegroundColor Green

    if ($script:VersionLookupFailureCount -gt 0) {
        Write-ProgressHost -Message ("[WARN] Version history could not be resolved for {0} file(s) — Version data above is understated by that much." -f $script:VersionLookupFailureCount) -ForegroundColor Yellow
        foreach ($sample in ($script:VersionLookupFailureSamples | Select-Object -Unique | Select-Object -First 5)) {
            Write-ProgressHost -Message ("        e.g. {0}" -f $sample) -ForegroundColor DarkYellow
        }
    }

    # ── Top libraries by version history size ─────────────────────────────────
    Write-Host ""
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-ProgressHost -Message "Top 5 libraries by version history size" -ForegroundColor Cyan
    Write-Host "  ================================================" -ForegroundColor Cyan
    $summaryRows |
        Where-Object { $null -ne $_.VersionSizeMB -and $_.VersionSizeMB -gt 0 } |
        Sort-Object { [double]$_.VersionSizeMB } -Descending |
        Select-Object -First 5 |
        ForEach-Object {
            Write-ProgressHost -Message ("{0}  [{1}] > {2}" -f
                (Format-SizeAuto -MB $_.VersionSizeMB),
                $_.SiteName,
                $_.Library) -ForegroundColor Yellow
        }

    # ── Top files by version history size ─────────────────────────────────────
    if (-not $SkipVersions -and $detailRows.Count -gt 0) {
        Write-Host ""
        Write-Host "  ================================================" -ForegroundColor Cyan
        Write-ProgressHost -Message "Top 10 files by version history size" -ForegroundColor Cyan
        Write-Host "  ================================================" -ForegroundColor Cyan
        $detailRows |
            Where-Object { $_.ItemType -eq 'File' -and $null -ne $_.VersionSizeMB -and $_.VersionSizeMB -gt 0 } |
            Sort-Object { [double]$_.VersionSizeMB } -Descending |
            Select-Object -First 10 |
            ForEach-Object {
                Write-ProgressHost -Message ("{0}  ({1} versies)  {2} > {3}" -f
                    (Format-SizeAuto -MB $_.VersionSizeMB),
                    $_.VersionCount,
                    $_.Library,
                    $_.Path) -ForegroundColor Yellow
                Write-ProgressHost -Message ("Site: {0}" -f $_.SiteName) -ForegroundColor DarkGray
            }
    }
}
Write-Host ""

# ── Cleanup ───────────────────────────────────────────────────────────────────
Remove-TempApp
