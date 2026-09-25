#Requires -Version 5.1
<#
.SYNOPSIS
    Exhaustive SharePoint Online permissions report — every site, sub-site, list/library, folder
    and file that carries its own permissions, exported to CSV.

.DESCRIPTION
    Walks the tenant (or a single site) and reports who has access to what, at every level where
    SharePoint actually stores an access decision:

      * Site collection administrators
      * Web (site and sub-site) role assignments, including inheritance breaks
      * SharePoint groups (Owners/Members/Visitors and custom) and their full membership
      * List and document library role assignments
      * Folder and file/list-item role assignments — every item with broken inheritance
      * Sharing links (anonymous / organization / specific people) and who they were shared with
      * External and guest users, "Everyone" and "Everyone except external users" grants
      * Entra ID group grants, resolved to their transitive user membership

    Inheritance is followed the way SharePoint models it: an item is only reported as its own scope
    when HasUniqueRoleAssignments is true. Everything else inherits from the nearest parent scope,
    which is reported once instead of being duplicated per item — so the CSV stays a faithful map
    of the permission structure rather than a row per file.

    Discovery is deliberately redundant: sites come from Graph's tenant-wide getAllSites, are then
    re-crawled for sub-sites through both Graph (/sites/{id}/sites) and SharePoint REST
    (/_api/web/webs) and de-duplicated on URL. Classic sub-webs that Graph omits are therefore
    still scanned. Within a web, every list is enumerated — including hidden and system lists with
    -IncludeHiddenLists — and every item in it is checked for a unique scope, folders included.

    Authentication:
      Reading role assignments is not something Microsoft Graph can do, and it is not covered by
      SharePoint's Read/Write/Manage application roles either — enumerating permissions requires
      the SharePoint "Sites.FullControl.All" application role. The script therefore connects you
      interactively once, then creates a short-lived App Registration granted:

        * SharePoint    Sites.FullControl.All  — role assignments, site groups, item scopes
        * Graph         Sites.Read.All         — tenant-wide site enumeration
        * Graph         GroupMember.Read.All   — resolving Entra group membership

      That app is authenticated with a certificate, not a secret. This is not a preference:
      SharePoint Online refuses any app-only token obtained with a client secret, answering 401
      with x-ms-diagnostics "Unsupported app only token". The certificate is generated in memory
      for the run, registered on the temporary app, and never written to the certificate store or
      to disk — there is nothing to clean up afterwards.

      The app is deleted again when the run finishes. Despite the Full Control role, this script
      only ever issues HTTP GET requests — it never writes, and it never changes a permission.

      To avoid the temporary app, pass -ClientId + -TenantId + -CertificateThumbprint for an
      existing app registration that already holds those roles. -ClientSecret is accepted for the
      Graph half but will fail against SharePoint for the reason above.

.PARAMETER SiteUrl
    Optional. Report on a single site collection (including its sub-sites) instead of the tenant.

.PARAMETER TenantUrl
    Tenant root URL, for example https://contoso.sharepoint.com. Required for a tenant-wide run.

.PARAMETER TenantId
    Entra ID tenant ID. Detected from the connected account when omitted. Required with -ClientId.

.PARAMETER ClientId
    Existing App Registration client ID. Skips the temporary app. Use with -TenantId and
    -ClientSecret or -CertificateThumbprint.

.PARAMETER ClientSecret
    Client secret for an existing app registration. Works for the Graph calls but NOT for the
    SharePoint ones — SharePoint Online rejects secret-based app-only tokens outright. Prefer
    -CertificateThumbprint; the script warns if you use this.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for an existing app registration. The certificate must be in
    Cert:\CurrentUser\My or Cert:\LocalMachine\My and have a private key. This is the supported
    way to authenticate an existing app against SharePoint Online.

.PARAMETER OutputPath
    Override the default output folder (C:\Temp on Windows).

.PARAMETER Scope
    How deep to descend. Site = webs only; List = webs + lists/libraries; Item = webs + lists +
    every folder/file/list item with unique permissions (default, and what "everything" means).

.PARAMETER IncludeOneDriveSites
    Also scan personal OneDrive sites. Off by default — these add one site per user.

.PARAMETER IncludeHiddenLists
    Also scan hidden and system lists (Form Templates, Style Library, workflow history, ...).
    Off by default because they rarely carry meaningful grants and are numerous.

.PARAMETER ListTitle
    Optional filter. Limit the scan to one or more list/library titles.

.PARAMETER ExcludeLimitedAccess
    Drop "Limited Access" role assignments. SharePoint adds these automatically so a principal can
    traverse to something it was granted deeper down; they are noise in most reviews but they are
    reported by default because leaving them out hides why a principal can see a folder path.

.PARAMETER SkipGroupExpansion
    Do not resolve SharePoint group and Entra ID group membership. Faster, but the report then only
    tells you which group has access, not who is in it.

.PARAMETER IncludeEffectiveAccess
    Also write an "effective access" CSV: one row per resolved user per scope, with the group they
    inherited the access through. Off by default — on a large tenant this file can be orders of
    magnitude larger than the detail CSV.

.PARAMETER GraphTimeoutSec
    Timeout in seconds per Graph/SharePoint call (default: 120).

.PARAMETER MaxGraphRetry
    Max retries on throttling/timeouts (default: 6).

.PARAMETER Concurrency
    Parallel workers (1-8, default: 4) for per-item role assignment lookups, which dominate the
    runtime of a -Scope Item run. Set to 1 to disable parallel dispatch.

.PARAMETER Restart
    Discard any existing checkpoint for this run (same parameters + output folder) and start over
    instead of resuming from the last completed list.

.EXAMPLE
    .\Get-SharePointPermissionsReport.ps1 -TenantUrl "https://contoso.sharepoint.com"

    Everything, tenant-wide: all sites, sub-sites, libraries, folders and files with unique rights.

.EXAMPLE
    .\Get-SharePointPermissionsReport.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Finance" -IncludeEffectiveAccess

    One site collection, plus a per-user effective access CSV.

.EXAMPLE
    .\Get-SharePointPermissionsReport.ps1 -TenantUrl "https://contoso.sharepoint.com" -Scope List -ExcludeLimitedAccess

    Faster overview: stops at list/library level and hides automatic traversal grants.
#>
[CmdletBinding()]
param(
    [string] $SiteUrl,
    [string] $TenantUrl,
    [string] $TenantId,
    [string] $ClientId,
    [string] $ClientSecret,
    [string] $CertificateThumbprint,
    [string] $OutputPath,
    [ValidateSet('Site', 'List', 'Item')]
    [string] $Scope = 'Item',
    [switch] $IncludeOneDriveSites,
    [switch] $IncludeHiddenLists,
    [string[]] $ListTitle = @(),
    [switch] $ExcludeLimitedAccess,
    [switch] $SkipGroupExpansion,
    [switch] $IncludeEffectiveAccess,
    [int] $GraphTimeoutSec = 120,
    [int] $MaxGraphRetry = 6,
    [ValidateRange(1, 8)]
    [int] $Concurrency = 4,
    [switch] $Restart
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($OutputPath) { $OutputPath }
             elseif ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' }
             else { "$HOME/Downloads" }
# Prove the output folder is usable before anything else happens. A tenant-wide scan can run for
# hours; discovering at the end that nothing could be written — or worse, failing on the first
# checkpoint after authenticating and creating an app registration — is entirely avoidable.
try {
    if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -ErrorAction Stop | Out-Null }
    $writeProbe = Join-Path $outputDir ".sp-permissions-write-test-$PID.tmp"
    [System.IO.File]::WriteAllText($writeProbe, 'probe')
    Remove-Item -Path $writeProbe -Force -ErrorAction SilentlyContinue
} catch {
    Write-Host "  [ERROR] Output folder '$outputDir' is not writable: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Pass -OutputPath to write somewhere else." -ForegroundColor Yellow
    exit 1
}

$ts           = Get-Date -Format 'yyyyMMdd_HHmmss'
$detailCsv    = Join-Path $outputDir "SharePoint_Permissions_Detail_$ts.csv"
$summaryCsv   = Join-Path $outputDir "SharePoint_Permissions_Summary_$ts.csv"
$groupsCsv    = Join-Path $outputDir "SharePoint_Permissions_Groups_$ts.csv"
$effectiveCsv = Join-Path $outputDir "SharePoint_Permissions_EffectiveAccess_$ts.csv"

# ── Well-known application IDs ────────────────────────────────────────────────
$GraphAppId      = '00000003-0000-0000-c000-000000000000'
$SharePointAppId = '00000003-0000-0ff1-ce00-000000000000'
$GraphResource   = 'https://graph.microsoft.com'

# ── Cleanup / shared state ────────────────────────────────────────────────────
$script:TempAppObjectId = $null
$script:ConnectedHere   = $false
$script:AppClientId     = $null
$script:AppClientSecret = $null
$script:AppCertificate  = $null
$script:AppSigningKey   = $null
$script:AppTenantId     = $null
$script:TokenCache      = @{}   # resource root URI -> @{ Headers; Expiry }
$script:CurrentScanLabel = ''

function Write-ProgressHost {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ConsoleColor]$ForegroundColor = [ConsoleColor]::DarkGray
    )
    Write-Host ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message) -ForegroundColor $ForegroundColor
}

function Set-ScanProgress {
    # Thin wrapper around Write-Progress so long phases show a real progress UI (percent + ETA in
    # interactive hosts) instead of relying purely on scrolling Write-Host log lines.
    param(
        [Parameter(Mandatory = $true)][int]$Id,
        [int]$ParentId = -1,
        [Parameter(Mandatory = $true)][string]$Activity,
        [Parameter(Mandatory = $true)][string]$Status,
        [int]$PercentComplete = -1
    )
    $params = @{ Id = $Id; Activity = $Activity; Status = $Status }
    if ($ParentId -ge 0) { $params['ParentId'] = $ParentId }
    if ($PercentComplete -ge 0) { $params['PercentComplete'] = [Math]::Min($PercentComplete, 100) }
    Write-Progress @params
}

function Complete-ScanProgress {
    param([Parameter(Mandatory = $true)][int]$Id, [int]$ParentId = -1)
    $params = @{ Id = $Id; Activity = 'Done'; Completed = $true }
    if ($ParentId -ge 0) { $params['ParentId'] = $ParentId }
    Write-Progress @params
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

function Append-CheckpointRows {
    # Retries on a locked or briefly unavailable file. The realistic cause is someone opening the
    # partial CSV in Excel mid-scan, which takes an exclusive lock: without this the write throws,
    # the unit is still marked complete, and those rows are gone from the report for good.
    param([string]$Path, [object[]]$Rows)
    if ($Rows.Count -eq 0) { return }

    $maxAttempts = 5
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            if (Test-Path $Path) {
                $Rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Append -ErrorAction Stop
            } else {
                $Rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            }
            return
        } catch {
            if ($attempt -eq $maxAttempts) {
                # Deliberately terminating: silently dropping rows would leave a report that looks
                # complete and is not. Losing the run is recoverable — the checkpoint resumes it.
                throw ("Could not write {0} row(s) to {1} after {2} attempts: {3}" -f $Rows.Count, $Path, $maxAttempts, $_.Exception.Message)
            }
            Write-ProgressHost -Message ("[WARN] Could not write to {0} (attempt {1}/{2}) — is it open in another program? Retrying..." -f (Split-Path $Path -Leaf), $attempt, $maxAttempts) -ForegroundColor Yellow
            Start-Sleep -Seconds (3 * $attempt)
        }
    }
}

function Set-GraphRequestTimeoutOptions {
    # The Microsoft.Graph SDK cmdlets have no per-call timeout and silently retry 429/503 with
    # their own backoff before surfacing anything to the caller, which under SharePoint throttling
    # is indistinguishable from a hang. Route every retry decision through this script instead.
    param([int]$TimeoutSec)
    if (Get-Command Set-MgRequestContext -ErrorAction SilentlyContinue) {
        try {
            Set-MgRequestContext -ClientTimeout $TimeoutSec -MaxRetry 0 -ErrorAction Stop
        } catch {
            Write-Host "  [WARN] Could not configure Graph SDK request timeout/retry options: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

function Remove-TempApp {
    # The delegated session is still open here, so Remove-MgApplication works.
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
    # The in-memory signing key holds unmanaged crypto state; release it once it can no longer be
    # needed rather than waiting for the finalizer.
    foreach ($disposable in @($script:AppSigningKey, $script:AppCertificate)) {
        if ($disposable -is [System.IDisposable]) { try { $disposable.Dispose() } catch {} }
    }
    $script:AppSigningKey  = $null
    $script:AppCertificate = $null
    $script:TokenCache     = @{}
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

# Last line of defence for the temporary App Registration. The scan itself cleans up in a finally,
# and the known failure paths call Remove-TempApp explicitly, but an unexpected terminating error
# anywhere between creating the app and reaching the scan would otherwise leave a Full Control app
# registration behind in the customer's tenant. That must not depend on having predicted the error.
trap {
    Write-Host ''
    Write-Host "  [ERROR] Unhandled error: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.ScriptStackTrace) { Write-Host "  $($_.ScriptStackTrace)" -ForegroundColor DarkGray }
    try { Remove-TempApp } catch {}
    exit 1
}

# ── Token handling ────────────────────────────────────────────────────────────
# One app registration, several resources: Graph for site enumeration and group expansion,
# and one SharePoint resource per host (contoso.sharepoint.com and, for OneDrive scans,
# contoso-my.sharepoint.com are separate audiences and need separate tokens).

function New-SelfSignedAppCertificate {
    # SharePoint Online refuses an app-only token that was obtained with a client secret — the
    # request comes back 401 with x-ms-diagnostics "Unsupported app only token." Only a
    # certificate-backed client credential is accepted against the SharePoint audience, so the
    # temporary app is given a certificate instead of a password.
    #
    # The key is generated in memory and never written to the certificate store or to disk: it
    # lives as long as the process does, which is already longer than the app registration it
    # authenticates. Nothing to clean up, and nothing left behind if the run is interrupted.
    param([string]$Subject = 'CN=SP-PermissionsReport-Temp')

    $rsa = [System.Security.Cryptography.RSA]::Create(2048)
    $req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
        $Subject, $rsa,
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $cert = $req.CreateSelfSigned([DateTimeOffset]::UtcNow.AddMinutes(-5), [DateTimeOffset]::UtcNow.AddDays(1))

    return [PSCustomObject]@{
        Certificate = $cert
        # Keep the generating RSA: on Windows the private key of an in-memory self-signed
        # certificate is not always retrievable through GetRSAPrivateKey(), and signing with the
        # object we already hold sidesteps that entirely.
        SigningKey  = $rsa
    }
}

function New-ClientAssertion {
    # Certificate credentials cannot be exchanged for a raw token the way a secret can, so build
    # the RFC 7523 client assertion ourselves. The Graph SDK does this internally, but SharePoint
    # REST is not reachable through the SDK and needs a token for the SharePoint audience.
    param(
        [Parameter(Mandatory = $true)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [Parameter(Mandatory = $true)][string]$ClientIdValue,
        [Parameter(Mandatory = $true)][string]$Authority,
        [System.Security.Cryptography.RSA]$SigningKey
    )
    $rsa = $SigningKey
    if (-not $rsa) { $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate) }
    if (-not $rsa) { throw "Certificate $($Certificate.Thumbprint) has no usable RSA private key." }

    function ConvertTo-Base64Url {
        param([byte[]]$Bytes)
        return [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    }

    $x5t    = ConvertTo-Base64Url -Bytes $Certificate.GetCertHash()
    $now    = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $header = @{ alg = 'RS256'; typ = 'JWT'; x5t = $x5t } | ConvertTo-Json -Compress
    $claims = @{
        aud = $Authority
        iss = $ClientIdValue
        sub = $ClientIdValue
        jti = [guid]::NewGuid().ToString()
        nbf = $now - 60
        exp = $now + 600
    } | ConvertTo-Json -Compress

    $encodedHeader = ConvertTo-Base64Url -Bytes ([Text.Encoding]::UTF8.GetBytes($header))
    $encodedClaims = ConvertTo-Base64Url -Bytes ([Text.Encoding]::UTF8.GetBytes($claims))
    $toSign        = "$encodedHeader.$encodedClaims"
    $signature     = $rsa.SignData(
        [Text.Encoding]::UTF8.GetBytes($toSign),
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
    )
    return "$toSign.$(ConvertTo-Base64Url -Bytes $signature)"
}

function Get-ResourceToken {
    # Returns an Authorization header hashtable for the given resource root, minting and caching
    # a client-credentials token per resource. Refreshes 5 minutes before expiry.
    param([Parameter(Mandatory = $true)][string]$Resource)

    $cached = $script:TokenCache[$Resource]
    if ($cached -and (Get-Date) -lt $cached.Expiry) { return $cached.Headers }

    $authority = "https://login.microsoftonline.com/$($script:AppTenantId)/oauth2/v2.0/token"
    $body = @{
        grant_type = 'client_credentials'
        scope      = "$Resource/.default"
        client_id  = $script:AppClientId
    }
    if ($script:AppCertificate) {
        $body['client_assertion_type'] = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        $body['client_assertion']      = New-ClientAssertion -Certificate $script:AppCertificate `
                                            -ClientIdValue $script:AppClientId -Authority $authority `
                                            -SigningKey $script:AppSigningKey
    } else {
        $body['client_secret'] = $script:AppClientSecret
    }

    # A freshly registered certificate credential and a freshly granted app role both need time to
    # replicate before Entra will mint a token against them — noticeably longer than the few
    # seconds a secret needs. Give it up to ~two minutes rather than failing the whole run on a
    # race that resolves itself.
    $resp = $null
    $lastError = $null
    $maxAttempts = 12
    for ($i = 1; $i -le $maxAttempts; $i++) {
        try {
            $resp = Invoke-RestMethod -Method POST -Uri $authority -Body $body -ErrorAction Stop
            break
        } catch {
            $lastError = $_
            if ($i -lt $maxAttempts) {
                Write-ProgressHost -Message ("[INFO] Waiting for credential/consent propagation for {0} (attempt {1}/{2})..." -f $Resource, $i, $maxAttempts)
                Start-Sleep -Seconds ([Math]::Min(5 * $i, 20))
            }
        }
    }
    if (-not $resp -or -not $resp.access_token) {
        $detail = $null
        try { $detail = ($lastError.ErrorDetails.Message | ConvertFrom-Json).error_description } catch {}
        if (-not $detail) { $detail = $lastError.Exception.Message }
        throw "Could not obtain app-only token for $Resource. $detail"
    }

    # SharePoint and Graph disagree on how to ask for a lean payload: SharePoint wants
    # odata=nometadata — which is also what keeps Get-SPCollection's paging shape predictable —
    # while Graph does not understand that parameter. Pick per resource instead of sending one
    # Accept header that is wrong for half the calls.
    $accept = if ($Resource -eq $GraphResource) { 'application/json' } else { 'application/json;odata=nometadata' }
    $headers = @{
        Authorization = "Bearer $($resp.access_token)"
        Accept        = $accept
        # Microsoft asks callers to identify themselves; traffic without a User-Agent is
        # throttled more aggressively than traffic that declares itself.
        'User-Agent'  = 'NONISV|M365-Scripts|SharePointPermissionsReport/1.0'
    }
    $script:TokenCache[$Resource] = @{
        Headers = $headers
        Expiry  = (Get-Date).AddSeconds([int]$resp.expires_in - 300)
    }
    return $headers
}

function Get-ResourceRootFromUrl {
    param([Parameter(Mandatory = $true)][string]$Url)
    $uri = [System.Uri]$Url
    return ("{0}://{1}" -f $uri.Scheme, $uri.Host)
}

# ── HTTP helpers ──────────────────────────────────────────────────────────────
function Get-RetryDelaySeconds {
    param([int]$Attempt, [object]$ErrorRecord)
    $retryAfter = $null
    try {
        $resp = $ErrorRecord.Exception.Response
        if ($resp -and $resp.Headers) {
            $retryHeader = $resp.Headers['Retry-After']
            if ($retryHeader) { [void][int]::TryParse([string]$retryHeader, [ref]$retryAfter) }
        }
    } catch {}
    if ($retryAfter -and $retryAfter -gt 0) { return [Math]::Min($retryAfter, 180) }
    return [Math]::Min([int][Math]::Pow(2, [Math]::Max(1, $Attempt)), 60)
}

function Get-ResponseStatusCode {
    param([object]$ErrorRecord)
    try {
        if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.StatusCode) {
            return [int]$ErrorRecord.Exception.Response.StatusCode
        }
    } catch {}
    return $null
}

function Invoke-GraphGet {
    param([Parameter(Mandatory = $true)][string]$Uri)
    for ($attempt = 1; $attempt -le $MaxGraphRetry; $attempt++) {
        try {
            $headers = Get-ResourceToken -Resource $GraphResource
            return Invoke-RestMethod -Uri $Uri -Headers $headers -TimeoutSec $GraphTimeoutSec -ErrorAction Stop
        } catch {
            $statusCode = Get-ResponseStatusCode -ErrorRecord $_
            $isRetryable = $statusCode -in @(408, 429, 500, 502, 503, 504)
            if (-not $isRetryable -and -not $statusCode) {
                $isRetryable = $_.Exception.Message -match 'timed out|timeout|temporar|connection|EOF|name resolution'
            }
            if (-not $isRetryable -or $attempt -eq $MaxGraphRetry) { throw }
            $delay = Get-RetryDelaySeconds -Attempt $attempt -ErrorRecord $_
            Write-ProgressHost -Message ("[INFO] Graph retry ({0}/{1}) in {2}s: {3}" -f $attempt, $MaxGraphRetry, $delay, $Uri)
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-ResponseHeaderValue {
    # SharePoint puts the real reason for a refusal in x-ms-diagnostics, not in the status line.
    # "Unsupported app only token" in particular is the difference between "this app may not read
    # this site" and "this credential type can never read any site".
    param([object]$ErrorRecord, [string]$Name)
    try {
        $headers = $ErrorRecord.Exception.Response.Headers
        if (-not $headers) { return $null }
        # PS 7 exposes HttpResponseHeaders (TryGetValues); PS 5.1 a WebHeaderCollection (indexer).
        if ($headers -is [System.Net.Http.Headers.HttpResponseHeaders]) {
            $values = $null
            if ($headers.TryGetValues($Name, [ref]$values)) { return ($values -join '; ') }
            return $null
        }
        return [string]$headers[$Name]
    } catch { return $null }
}

function Invoke-SPGet {
    # Single SharePoint REST GET. Returns $null for 403/404 — a site the app cannot open, or a
    # list/endpoint that does not exist on this template — because a tenant-wide sweep always hits
    # some of both and neither should abort the run.
    #
    # A 401 is deliberately NOT treated that way. It means the token itself is not accepted, which
    # is never per-site: it is the same answer for every site in the tenant. Swallowing it turned a
    # single credential fault into 130 lines of "not accessible with the current permissions",
    # which reads like a permissions finding instead of the bug it is.
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [switch] $ThrowOnDenied
    )
    $resourceRoot = Get-ResourceRootFromUrl -Url $Uri
    $reauthTried  = $false
    for ($attempt = 1; $attempt -le $MaxGraphRetry; $attempt++) {
        try {
            $headers = Get-ResourceToken -Resource $resourceRoot
            return Invoke-RestMethod -Uri $Uri -Headers $headers -TimeoutSec $GraphTimeoutSec -ErrorAction Stop
        } catch {
            $statusCode = Get-ResponseStatusCode -ErrorRecord $_
            if ($statusCode -eq 401) {
                # A 401 is usually fatal, but not always: a token can be rejected right on the
                # expiry boundary, or after the service principal is re-replicated. Throw away the
                # cached token and try once more before concluding the credential itself is wrong.
                if (-not $reauthTried) {
                    $reauthTried = $true
                    $script:TokenCache.Remove($resourceRoot)
                    Write-ProgressHost -Message "[INFO] Token refused (401) — re-authenticating once before giving up..."
                    Start-Sleep -Seconds 2
                    continue
                }
                $diag = Get-ResponseHeaderValue -ErrorRecord $_ -Name 'x-ms-diagnostics'
                $hint = ''
                if ($diag -match 'Unsupported app only token') {
                    $hint = "`n         SharePoint Online does not accept an app-only token obtained with a client secret." +
                            "`n         Use -CertificateThumbprint, or omit -ClientId so the script creates its own certificate-backed app."
                }
                throw ("SharePoint refused the token (401) for {0}.{1}{2}" -f $resourceRoot, $(if ($diag) { " $diag" } else { '' }), $hint)
            }
            if ($statusCode -in @(403, 404) -and -not $ThrowOnDenied) { return $null }
            $isRetryable = $statusCode -in @(408, 429, 500, 502, 503, 504)
            if (-not $isRetryable -and -not $statusCode) {
                $isRetryable = $_.Exception.Message -match 'timed out|timeout|temporar|connection|EOF|name resolution'
            }
            if (-not $isRetryable -or $attempt -eq $MaxGraphRetry) { throw }
            $delay = Get-RetryDelaySeconds -Attempt $attempt -ErrorRecord $_
            Write-ProgressHost -Message ("[WAIT] SharePoint throttled — retry ({0}/{1}) in {2}s" -f $attempt, $MaxGraphRetry, $delay) -ForegroundColor Yellow
            Start-Sleep -Seconds $delay
        }
    }
}

function Invoke-SPCollectionPaged {
    # Walks SharePoint's paging (odata.nextLink under nometadata, __next under verbose) and hands
    # each page to -OnPage. Nothing is accumulated here, so a library with a million items costs
    # one page of memory instead of a million live objects.
    #
    # Item paging uses $skiptoken internally, so this does not trip the 5000-item list view
    # threshold the way a $filter or $orderby query would.
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][scriptblock]$OnPage,
        [int] $MaxPages = 200000
    )
    $next = $Uri
    $page = 0
    # SharePoint has been seen to echo back a nextLink identical to the request under some error
    # conditions. Without this guard that is an infinite loop against a live tenant.
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    while ($next -and $page -lt $MaxPages) {
        if (-not $seen.Add($next)) {
            Write-ProgressHost -Message "[WARN] SharePoint repeated a paging link — stopping this collection to avoid looping." -ForegroundColor Yellow
            break
        }
        $resp = Invoke-SPGet -Uri $next
        if (-not $resp) { break }
        $page++

        $values = $null
        if ($null -ne $resp.value) { $values = $resp.value }
        elseif ($resp.d -and $null -ne $resp.d.results) { $values = $resp.d.results }
        if ($values) { & $OnPage @($values) }

        $next = $null
        if ($resp.'odata.nextLink') { $next = [string]$resp.'odata.nextLink' }
        elseif ($resp.d -and $resp.d.__next) { $next = [string]$resp.d.__next }
    }
    if ($page -ge $MaxPages) {
        Write-ProgressHost -Message ("[WARN] Stopped paging after {0} pages — the collection may be incomplete." -f $MaxPages) -ForegroundColor Yellow
    }
}

function Get-SPCollection {
    # Accumulating wrapper, for the collections that are inherently small: site groups, lists,
    # role assignments, sub-webs. List *items* deliberately do not go through this — see
    # Invoke-SPCollectionPaged.
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [int] $MaxPages = 200000
    )
    $rows = [System.Collections.Generic.List[object]]::new()
    Invoke-SPCollectionPaged -Uri $Uri -MaxPages $MaxPages -OnPage {
        param($Values)
        foreach ($v in $Values) { $rows.Add($v) | Out-Null }
    }
    return $rows
}

# ── Principal classification ──────────────────────────────────────────────────
# SharePoint stores every grantee as a claims login name. The shape of that string is the only
# reliable way to tell a person from a security group from a sharing link, so parse it explicitly
# rather than trusting PrincipalType alone (which reports both Entra groups and SharePoint groups
# in ways that collapse distinctions we care about here).

function Get-PrincipalInfo {
    param([object]$Member)

    $login = [string]$Member.LoginName
    $title = [string]$Member.Title
    $email = [string]$Member.Email
    $spType = [int]0
    if ($null -ne $Member.PrincipalType) { $spType = [int]$Member.PrincipalType }

    $kind        = 'Unknown'
    $directoryId = $null
    $isExternal  = $false
    $linkKind    = $null

    switch -Regex ($login) {
        '^c:0\(\.s\|true$' {
            $kind = 'Everyone'
            if (-not $title) { $title = 'Everyone' }
            break
        }
        '^c:0-\.f\|rolemanager\|spo-grid-all-users' {
            $kind = 'EveryoneExceptExternalUsers'
            if (-not $title) { $title = 'Everyone except external users' }
            break
        }
        '^c:0!\.s\|windows$' {
            $kind = 'AllAuthenticatedUsers'
            break
        }
        '^SharingLinks\.' {
            # SharingLinks.<itemGuid>.<LinkKind>.<linkGuid>
            $kind = 'SharingLink'
            $parts = $login.Split('.')
            if ($parts.Count -ge 3) { $linkKind = $parts[2] }
            break
        }
        '^c:0o\.c\|federateddirectoryclaimprovider\|' {
            # Microsoft 365 group. A trailing _o addresses the owners, not the members.
            $raw = $login.Split('|')[-1]
            if ($raw -match '_o$') {
                $kind = 'M365GroupOwners'
                $directoryId = $raw -replace '_o$', ''
            } else {
                $kind = 'M365Group'
                $directoryId = $raw
            }
            break
        }
        '^c:0t\.c\|tenant\|' {
            $kind = 'SecurityGroup'
            $directoryId = $login.Split('|')[-1]
            break
        }
        '^i:0#\.f\|membership\|' {
            $kind = 'User'
            $upn  = $login.Split('|')[-1]
            if (-not $email) { $email = $upn }
            if ($upn -match '#ext#') { $isExternal = $true }
            break
        }
        '^i:0#\.w\|' {
            $kind = 'User'
            break
        }
        default {
            if ($spType -eq 8) { $kind = 'SharePointGroup' }
            elseif ($spType -eq 4) { $kind = 'SecurityGroup' }
            elseif ($spType -eq 2) { $kind = 'DistributionList' }
            elseif ($spType -eq 1) { $kind = 'User' }
        }
    }

    # PrincipalType 8 always wins for SharePoint's own groups — their login name is just the group
    # title, which matches none of the claim patterns above.
    if ($spType -eq 8 -and $kind -notin @('SharingLink')) { $kind = 'SharePointGroup' }

    if (-not $isExternal -and ($email -match '#ext#' -or $login -match 'urn:spo:(guest|anon)')) {
        $isExternal = $true
    }

    return [PSCustomObject]@{
        Kind        = $kind
        Title       = $title
        LoginName   = $login
        Email       = $email
        DirectoryId = $directoryId
        IsExternal  = $isExternal
        LinkKind    = $linkKind
        SpTypeId    = $spType
        SpGroupId   = $(if ($spType -eq 8) { [int]$Member.Id } else { $null })
    }
}

function Get-SharingLinkDescription {
    # The link kind is embedded in the SharingLinks.* principal name. Translating it beats an extra
    # round trip to GetSharingInformation for every shared item, which would multiply the call
    # count on exactly the items a tenant tends to have thousands of.
    param([string]$LinkKind)
    switch -Regex ($LinkKind) {
        '^Anonymous'    { return 'Anyone with the link (anonymous)' }
        '^Flexible'     { return 'Specific people / custom link' }
        '^Organization' { return 'Anyone in the organization' }
        '^AnonymousEdit'{ return 'Anyone with the link (anonymous, edit)' }
        default         { if ($LinkKind) { return $LinkKind } else { return 'Sharing link' } }
    }
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host '   Get-SharePointPermissionsReport' -ForegroundColor Cyan
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host ("  Scope     : {0}" -f $(switch ($Scope) {
    'Site' { 'Sites and sub-sites only' }
    'List' { 'Sites, sub-sites, lists and libraries' }
    'Item' { 'Everything — sites, lists, folders and files with unique permissions' }
})) -ForegroundColor Cyan
Write-Host ("  Target    : {0}" -f $(if ($SiteUrl) { $SiteUrl } else { "$TenantUrl (tenant-wide)" })) -ForegroundColor Cyan
Write-Host  '  Mode      : Read-only — this script never changes a permission' -ForegroundColor DarkGray
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

# Unlike the storage/version reports, a single-site run needs the app-only path too: SharePoint
# role assignments are only readable with a Sites.FullControl.All token, and a delegated Graph
# token is the wrong audience for the /_api endpoints entirely.
$useTempApp = -not $ClientId
if ($useTempApp -and -not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Applications')) {
    Write-Host "  [ERROR] Missing required module: Microsoft.Graph.Applications (needed for the temporary App Registration)." -ForegroundColor Red
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

function Resolve-ClientCertificate {
    param([Parameter(Mandatory = $true)][string]$Thumbprint)
    $normalized = $Thumbprint.Replace(' ', '').ToUpperInvariant()
    foreach ($store in @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My')) {
        $cert = Get-ChildItem -Path $store -ErrorAction SilentlyContinue |
                Where-Object { $_.Thumbprint -eq $normalized } |
                Select-Object -First 1
        if ($cert) { return $cert }
    }
    throw "Certificate with thumbprint $Thumbprint was not found in CurrentUser\My or LocalMachine\My."
}

# ── Connection ────────────────────────────────────────────────────────────────
try {
    if ($ClientId) {
        # Existing app registration — no interactive sign-in and no temp app at all. Tokens for
        # both Graph and SharePoint are minted directly from these credentials.
        $script:AppClientId = $ClientId
        $script:AppTenantId = $effectiveTenantId
        if ($CertificateThumbprint) {
            $script:AppCertificate = Resolve-ClientCertificate -Thumbprint $CertificateThumbprint
        } elseif ($ClientSecret) {
            $script:AppClientSecret = $ClientSecret
            # Not an error — Graph accepts it — but every SharePoint call will come back 401, and
            # that failure is invisible unless it is called out here: the scan would simply report
            # every site as inaccessible.
            Write-Host "  [WARN] SharePoint Online rejects app-only tokens obtained with a client secret" -ForegroundColor Yellow
            Write-Host "         ('Unsupported app only token'). Use -CertificateThumbprint instead, or omit" -ForegroundColor Yellow
            Write-Host "         -ClientId entirely and let the script create its own certificate-backed app." -ForegroundColor Yellow
        } else {
            Write-Host "  [ERROR] -ClientId requires -ClientSecret or -CertificateThumbprint." -ForegroundColor Red
            exit 1
        }
        [void](Get-ResourceToken -Resource $GraphResource)
        Write-Host "  [OK]   Connected with provided app credentials." -ForegroundColor DarkGray
    } else {
        Write-Host "  Connecting interactively..." -ForegroundColor Cyan
        Write-Host "  Required role: Global Administrator or Application Administrator (to create the temporary lookup app)" -ForegroundColor DarkGray
        $connectParams = @{
            Scopes    = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
            NoWelcome = $true
        }
        if ($effectiveTenantId) { $connectParams['TenantId'] = $effectiveTenantId }
        Connect-MgGraph @connectParams -ErrorAction Stop
        $script:ConnectedHere = $true
        Write-Host "  [OK]   Connected (delegated)." -ForegroundColor DarkGray

        $ctx = Get-MgContext
        $usedTenantId = if ($effectiveTenantId) { $effectiveTenantId } else { $ctx.TenantId }
        if (-not $usedTenantId) {
            Write-Host "  [ERROR] Could not determine tenant ID. Provide -TenantId." -ForegroundColor Red
            Remove-TempApp; exit 1
        }

        $tempAppName = "SP-PermissionsReport-Temp-$ts"
        Write-Host "  Creating temporary App Registration '$tempAppName'..." -ForegroundColor Cyan

        # Certificate, not a password: SharePoint Online returns 401 "Unsupported app only token"
        # for any app-only token that was obtained with a client secret, so a secret-backed temp
        # app would authenticate fine against Graph and then fail on every single /_api call.
        $tempCert = New-SelfSignedAppCertificate -Subject "CN=$tempAppName"
        $keyCredential = @{
            Type        = 'AsymmetricX509Cert'
            Usage       = 'Verify'
            Key         = $tempCert.Certificate.GetRawCertData()
            DisplayName = "CN=$tempAppName"
        }
        $app = New-MgApplication -DisplayName $tempAppName -KeyCredentials @($keyCredential) -ErrorAction Stop
        $script:TempAppObjectId = $app.Id
        $sp = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop
        Write-Host ("  [OK]   Certificate credential registered (thumbprint {0}, valid 1 day, never written to disk)." -f $tempCert.Certificate.Thumbprint) -ForegroundColor DarkGray

        # Sites.FullControl.All is not an oversight here: SharePoint gates reading role
        # assignments behind the EnumeratePermissions right, which only Full Control carries.
        # Read, Write and Manage all return 403 on /roleassignments. The app stays read-only in
        # practice — every call this script makes is a GET — and it is deleted when the run ends.
        $requiredRoles = @(
            @{ ResourceAppId = $SharePointAppId; Role = 'Sites.FullControl.All'; Why = 'role assignments, site groups, unique scopes' }
            @{ ResourceAppId = $GraphAppId;      Role = 'Sites.Read.All';        Why = 'tenant-wide site enumeration' }
            @{ ResourceAppId = $GraphAppId;      Role = 'GroupMember.Read.All';  Why = 'Entra group membership expansion' }
        )
        foreach ($required in $requiredRoles) {
            $resourceSp = Get-MgServicePrincipal -Filter "appId eq '$($required.ResourceAppId)'" -ErrorAction Stop
            if (-not $resourceSp) { throw "Could not resolve service principal for resource $($required.ResourceAppId)." }
            $appRole = $resourceSp.AppRoles | Where-Object { $_.Value -eq $required.Role -and $_.AllowedMemberTypes -contains 'Application' }
            if (-not $appRole) { throw "Could not resolve app role '$($required.Role)' on $($resourceSp.DisplayName)." }
            New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $sp.Id `
                -PrincipalId        $sp.Id `
                -ResourceId         $resourceSp.Id `
                -AppRoleId          $appRole.Id `
                -ErrorAction Stop | Out-Null
            Write-Host ("  [OK]   {0} / {1} granted ({2})." -f $resourceSp.DisplayName, $required.Role, $required.Why) -ForegroundColor DarkGray
        }

        $script:AppClientId    = $app.AppId
        $script:AppCertificate = $tempCert.Certificate
        $script:AppSigningKey  = $tempCert.SigningKey
        $script:AppTenantId    = $usedTenantId

        Write-Host "  Obtaining app-only tokens..." -ForegroundColor Cyan
        [void](Get-ResourceToken -Resource $GraphResource)
        Write-Host "  [OK]   Token obtained." -ForegroundColor DarkGray
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Remove-TempApp; exit 1
}

# Set-MgRequestContext returns the context object; without this it lands on stdout and prints a
# stray "ClientTimeout RetryDelay MaxRetry" table at the end of an otherwise clean run.
Set-GraphRequestTimeoutOptions -TimeoutSec $GraphTimeoutSec | Out-Null

# ── SharePoint access preflight ───────────────────────────────────────────────
# One call against the tenant root, before enumerating anything. Whether SharePoint accepts this
# credential is a single yes/no for the whole tenant, so finding out here costs one request and
# turns an otherwise silent 130-site sweep of "not accessible" into one actionable error.
$preflightRoot = if ($SiteUrl) { Get-ResourceRootFromUrl -Url $SiteUrl } else { Get-ResourceRootFromUrl -Url $TenantUrl }
Write-ProgressHost -Message ("Verifying SharePoint access against {0}..." -f $preflightRoot) -ForegroundColor Cyan
try {
    $preflightWeb = Invoke-SPGet -Uri ("{0}/_api/web?`$select=Title" -f $preflightRoot) -ThrowOnDenied
    if (-not $preflightWeb) { throw "SharePoint returned no data for $preflightRoot/_api/web." }
    Write-Host ("  [OK]   SharePoint accepted the token (root web: {0})." -f $preflightWeb.Title) -ForegroundColor DarkGray
} catch {
    Write-Host ''
    Write-Host "  [ERROR] Cannot read SharePoint with this credential — stopping before the scan." -ForegroundColor Red
    Write-Host ("  {0}" -f $_.Exception.Message) -ForegroundColor Red
    Write-Host ''
    Write-Host "  Reading role assignments needs the SharePoint application role Sites.FullControl.All" -ForegroundColor Yellow
    Write-Host "  on a certificate-backed app registration. Graph permissions alone are not enough, and" -ForegroundColor Yellow
    Write-Host "  a client secret is not accepted by SharePoint Online for app-only access." -ForegroundColor Yellow
    Remove-TempApp; exit 1
}

# ── Site discovery ────────────────────────────────────────────────────────────
# Three sources, de-duplicated on URL, because no single one is complete:
#   1. Graph getAllSites          — tenant-wide, but misses some classic sub-webs
#   2. Graph /sites/{id}/sites    — sub-sites Graph does know about
#   3. SharePoint /_api/web/webs  — the authoritative sub-web list, including classic ones
Write-ProgressHost -Message "Retrieving sites..." -ForegroundColor Cyan

$targetSites  = [System.Collections.Generic.List[object]]::new()
$knownSiteUrl = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

function Add-TargetSite {
    param([string]$WebUrl, [string]$Title, [string]$GraphId)
    if ([string]::IsNullOrWhiteSpace($WebUrl)) { return $false }
    $normalized = $WebUrl.TrimEnd('/')
    if (-not $knownSiteUrl.Add($normalized)) { return $false }
    $targetSites.Add([PSCustomObject]@{
        WebUrl  = $normalized
        Title   = $Title
        GraphId = $GraphId
    }) | Out-Null
    return $true
}

if (-not $allSitesMode) {
    try {
        $siteUri      = [System.Uri]$SiteUrl.TrimEnd('/')
        $siteHost     = $siteUri.Host
        $sitePath     = $siteUri.AbsolutePath.TrimEnd('/')
        $siteGraphUri = "https://graph.microsoft.com/v1.0/sites/${siteHost}:${sitePath}?`$select=id,displayName,webUrl"
        $siteObj      = Invoke-GraphGet -Uri $siteGraphUri
        if (-not $siteObj -or -not $siteObj.id) {
            Write-Host "  [ERROR] Site not found: $SiteUrl" -ForegroundColor Red
            Remove-TempApp; exit 1
        }
        [void](Add-TargetSite -WebUrl $siteObj.webUrl -Title $siteObj.displayName -GraphId $siteObj.id)
    } catch {
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        Remove-TempApp; exit 1
    }
} else {
    $firstUri  = 'https://graph.microsoft.com/v1.0/sites/getAllSites?$select=id,displayName,webUrl&$top=200'
    $nextUri   = $null
    $firstDone = $false
    for ($i = 1; $i -le 6; $i++) {
        try {
            $response = Invoke-GraphGet -Uri $firstUri
            foreach ($s in @($response.value)) { [void](Add-TargetSite -WebUrl $s.webUrl -Title $s.displayName -GraphId $s.id) }
            $nextUri   = $response.'@odata.nextLink'
            $firstDone = $true
            break
        } catch {
            if ($i -lt 6) {
                Write-ProgressHost -Message ("[INFO] Waiting for consent propagation (attempt {0}/6)..." -f $i)
                Start-Sleep -Seconds 5
            } else {
                Write-Host "  [ERROR] Failed to retrieve sites: $($_.Exception.Message)" -ForegroundColor Red
                Remove-TempApp; exit 1
            }
        }
    }
    while ($firstDone -and $nextUri) {
        $response = Invoke-GraphGet -Uri $nextUri
        foreach ($s in @($response.value)) { [void](Add-TargetSite -WebUrl $s.webUrl -Title $s.displayName -GraphId $s.id) }
        $nextUri = $response.'@odata.nextLink'
    }

    if (-not $IncludeOneDriveSites) {
        $filtered = @($targetSites | Where-Object { $_.WebUrl -notmatch '-my\.sharepoint\.com/personal/' })
        $targetSites = [System.Collections.Generic.List[object]]::new($filtered)
        $knownSiteUrl.Clear()
        foreach ($s in $targetSites) { [void]$knownSiteUrl.Add($s.WebUrl) }
    }
}

# Sub-webs at all depths. Graph's getAllSites (and a single-site lookup) returns site collections,
# not the classic sub-webs underneath them, so a library living on a sub-web would otherwise never
# be scanned. Crawl both sources and let the URL hash set sort out the overlap.
$discoveryQueue = [System.Collections.Generic.Queue[object]]::new()
foreach ($s in $targetSites) { $discoveryQueue.Enqueue($s) }

while ($discoveryQueue.Count -gt 0) {
    $parent = $discoveryQueue.Dequeue()

    if ($parent.GraphId) {
        try {
            $subSitesUri = "https://graph.microsoft.com/v1.0/sites/$($parent.GraphId)/sites`?`$select=id,displayName,webUrl&`$top=200"
            do {
                $subResp = Invoke-GraphGet -Uri $subSitesUri
                foreach ($s in @($subResp.value)) {
                    if (Add-TargetSite -WebUrl $s.webUrl -Title $s.displayName -GraphId $s.id) {
                        $discoveryQueue.Enqueue($targetSites[$targetSites.Count - 1])
                    }
                }
                $subSitesUri = $subResp.'@odata.nextLink'
            } while ($subSitesUri)
        } catch {
            # Most sites have no sub-sites, or are inaccessible with the current permissions.
        }
    }

    try {
        $webs = Get-SPCollection -Uri ("{0}/_api/web/webs?`$select=Title,Url" -f $parent.WebUrl)
        foreach ($w in $webs) {
            if (Add-TargetSite -WebUrl $w.Url -Title $w.Title -GraphId $null) {
                $discoveryQueue.Enqueue($targetSites[$targetSites.Count - 1])
            }
        }
    } catch {
        # A web the app cannot open, or a template without sub-webs.
    }
}

Write-ProgressHost -Message ("Target webs: {0}" -f $targetSites.Count) -ForegroundColor Green

if ($targetSites.Count -eq 0) {
    Write-Host ''
    Write-Host "  [ERROR] No sites to scan — nothing was discovered." -ForegroundColor Red
    if ($allSitesMode -and -not $IncludeOneDriveSites) {
        Write-Host "  Every site found was a personal OneDrive site, which is excluded by default." -ForegroundColor Yellow
        Write-Host "  Add -IncludeOneDriveSites to include them." -ForegroundColor Yellow
    } else {
        Write-Host "  Check that the tenant URL is right and that the app's Sites.Read.All grant has replicated." -ForegroundColor Yellow
    }
    Remove-TempApp; exit 1
}

# ── Principal membership resolution ───────────────────────────────────────────
# Two caches, because the same handful of groups grant access all over a tenant and resolving
# them once per grant would dwarf the cost of the scan itself.
$script:SiteGroupCache   = @{}   # site collection URL -> @{ groupId -> group object with Users }
$script:EntraGroupCache  = @{}   # Entra group object id -> resolved member list
$script:GroupRowsWritten = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

function Get-SiteCollectionUrl {
    # SharePoint groups live on the site collection, not the web, so sub-webs must share the cache
    # entry of their parent — otherwise every sub-web re-fetches the same group membership.
    param([Parameter(Mandatory = $true)][string]$WebUrl)
    $uri  = [System.Uri]$WebUrl
    $path = $uri.AbsolutePath.TrimEnd('/')
    if ($path -match '^(/(?:sites|teams|personal)/[^/]+)') {
        return ("{0}://{1}{2}" -f $uri.Scheme, $uri.Host, $Matches[1])
    }
    return ("{0}://{1}" -f $uri.Scheme, $uri.Host)
}

function Get-SiteGroups {
    # One call returns every SharePoint group on the site collection together with its members,
    # which is what turns "Contoso Owners has Full Control" into a list of actual people.
    param([Parameter(Mandatory = $true)][string]$WebUrl)

    $scUrl = Get-SiteCollectionUrl -WebUrl $WebUrl
    if ($script:SiteGroupCache.ContainsKey($scUrl)) { return $script:SiteGroupCache[$scUrl] }

    $map = @{}
    try {
        $groups = Get-SPCollection -Uri ("{0}/_api/web/sitegroups?`$select=Id,Title,LoginName,Description,OwnerTitle,Users/Id,Users/Title,Users/LoginName,Users/Email,Users/PrincipalType&`$expand=Users" -f $scUrl)
        foreach ($g in $groups) { $map[[string]$g.Id] = $g }
    } catch {
        Write-ProgressHost -Message ("[WARN] Could not read site groups for {0}: {1}" -f $scUrl, $_.Exception.Message) -ForegroundColor Yellow
    }
    $script:SiteGroupCache[$scUrl] = $map
    return $map
}

function Get-EntraGroupMembers {
    # Transitive, so a grant to a parent group reports the people in its nested groups too — which
    # is the whole point of an access review and the thing nested groups hide from you.
    param([Parameter(Mandatory = $true)][string]$GroupId, [switch]$OwnersOnly)

    $cacheKey = if ($OwnersOnly) { "$GroupId|owners" } else { "$GroupId|members" }
    if ($script:EntraGroupCache.ContainsKey($cacheKey)) { return $script:EntraGroupCache[$cacheKey] }

    $members = [System.Collections.Generic.List[object]]::new()
    $segment = if ($OwnersOnly) { 'owners' } else { 'transitiveMembers' }
    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/$segment/microsoft.graph.user" +
           '?$select=id,displayName,userPrincipalName,mail,userType,accountEnabled&$top=999'
    try {
        while ($uri) {
            $resp = Invoke-GraphGet -Uri $uri
            foreach ($u in @($resp.value)) {
                $members.Add([PSCustomObject]@{
                    DisplayName = $u.displayName
                    Upn         = $u.userPrincipalName
                    Email       = if ($u.mail) { $u.mail } else { $u.userPrincipalName }
                    IsExternal  = ($u.userType -eq 'Guest' -or ([string]$u.userPrincipalName) -match '#ext#')
                    Enabled     = $u.accountEnabled
                }) | Out-Null
            }
            $uri = $resp.'@odata.nextLink'
        }
    } catch {
        Write-ProgressHost -Message ("[WARN] Could not expand Entra group {0}: {1}" -f $GroupId, $_.Exception.Message) -ForegroundColor Yellow
    }

    $script:EntraGroupCache[$cacheKey] = $members
    return $members
}

function Resolve-PrincipalMembers {
    # Returns the individual people behind a principal. A SharePoint group can itself contain Entra
    # groups, so its members are re-resolved rather than reported as-is.
    param(
        [Parameter(Mandatory = $true)][object]$Principal,
        [Parameter(Mandatory = $true)][string]$WebUrl,
        [int]$Depth = 0
    )

    $resolved = [System.Collections.Generic.List[object]]::new()
    if ($SkipGroupExpansion -or $Depth -gt 3) { return $resolved }

    switch ($Principal.Kind) {
        'User' {
            $resolved.Add([PSCustomObject]@{
                DisplayName = $Principal.Title
                Upn         = $Principal.LoginName.Split('|')[-1]
                Email       = $Principal.Email
                IsExternal  = $Principal.IsExternal
                Enabled     = $null
            }) | Out-Null
        }
        { $_ -in @('SharePointGroup', 'SharingLink') } {
            $groups = Get-SiteGroups -WebUrl $WebUrl
            $group  = $null
            if ($Principal.SpGroupId -and $groups.ContainsKey([string]$Principal.SpGroupId)) {
                $group = $groups[[string]$Principal.SpGroupId]
            }
            if (-not $group) {
                # Sharing-link groups are not always returned by /sitegroups; ask for this one.
                try {
                    $group = Invoke-SPGet -Uri ("{0}/_api/web/sitegroups({1})?`$select=Id,Title,Users/Id,Users/Title,Users/LoginName,Users/Email,Users/PrincipalType&`$expand=Users" -f $WebUrl, $Principal.SpGroupId)
                } catch { $group = $null }
            }
            foreach ($u in @($group.Users)) {
                $inner = Get-PrincipalInfo -Member $u
                if ($inner.Kind -eq 'User') {
                    $resolved.Add([PSCustomObject]@{
                        DisplayName = $inner.Title
                        Upn         = $inner.LoginName.Split('|')[-1]
                        Email       = $inner.Email
                        IsExternal  = $inner.IsExternal
                        Enabled     = $null
                    }) | Out-Null
                } else {
                    foreach ($m in (Resolve-PrincipalMembers -Principal $inner -WebUrl $WebUrl -Depth ($Depth + 1))) {
                        $resolved.Add($m) | Out-Null
                    }
                }
            }
        }
        { $_ -in @('SecurityGroup', 'M365Group') } {
            if ($Principal.DirectoryId) {
                foreach ($m in (Get-EntraGroupMembers -GroupId $Principal.DirectoryId)) { $resolved.Add($m) | Out-Null }
            }
        }
        'M365GroupOwners' {
            if ($Principal.DirectoryId) {
                foreach ($m in (Get-EntraGroupMembers -GroupId $Principal.DirectoryId -OwnersOnly)) { $resolved.Add($m) | Out-Null }
            }
        }
        default { }
    }

    return $resolved
}

# ── Role assignment reading ───────────────────────────────────────────────────
function ConvertTo-PermissionRows {
    # Turns one SharePoint roleassignments response into flat CSV rows — one per principal per
    # scope, with the permission levels joined, plus (optionally) the resolved people behind it.
    param(
        [Parameter(Mandatory = $true)][object[]]$RoleAssignments,
        [Parameter(Mandatory = $true)][hashtable]$ScopeInfo,
        [System.Collections.Generic.List[object]]$EffectiveRows
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($ra in $RoleAssignments) {
        if (-not $ra.Member) { continue }

        $levels = @()
        foreach ($binding in @($ra.RoleDefinitionBindings)) {
            if ($binding.Name) { $levels += [string]$binding.Name }
        }
        if ($ExcludeLimitedAccess) {
            $levels = @($levels | Where-Object { $_ -ne 'Limited Access' })
            if ($levels.Count -eq 0) { continue }
        }
        if ($levels.Count -eq 0) { continue }

        $principal = Get-PrincipalInfo -Member $ra.Member
        $members   = Resolve-PrincipalMembers -Principal $principal -WebUrl $ScopeInfo.WebUrl

        $externalCount = @($members | Where-Object { $_.IsExternal }).Count
        if ($principal.IsExternal) { $externalCount = [Math]::Max($externalCount, 1) }

        $rows.Add([PSCustomObject]@{
            SiteUrl           = $ScopeInfo.SiteUrl
            WebUrl            = $ScopeInfo.WebUrl
            WebTitle          = $ScopeInfo.WebTitle
            ScopeType         = $ScopeInfo.ScopeType
            ScopeTitle        = $ScopeInfo.ScopeTitle
            ScopeUrl          = $ScopeInfo.ScopeUrl
            ItemType          = $ScopeInfo.ItemType
            ListTitle         = $ScopeInfo.ListTitle
            ListTemplate      = $ScopeInfo.ListTemplate
            HasUniquePerms    = $ScopeInfo.HasUniquePerms
            InheritsFrom      = $ScopeInfo.InheritsFrom
            PrincipalType     = $principal.Kind
            PrincipalName     = $principal.Title
            PrincipalLogin    = $principal.LoginName
            PrincipalEmail    = $principal.Email
            DirectoryObjectId = $principal.DirectoryId
            PermissionLevels  = ($levels -join '; ')
            IsSharingLink     = ($principal.Kind -eq 'SharingLink')
            SharingLinkType   = $(if ($principal.Kind -eq 'SharingLink') { Get-SharingLinkDescription -LinkKind $principal.LinkKind } else { $null })
            IsExternal        = $principal.IsExternal
            MemberCount       = $members.Count
            ExternalMembers   = $externalCount
            MemberPreview     = (($members | Select-Object -First 10 | ForEach-Object { $_.Upn }) -join '; ')
            LastModified      = $ScopeInfo.LastModified
            ScannedUtc        = (Get-Date).ToUniversalTime().ToString('s')
            # Always present, even when empty: every detail row is appended to the same CSV, and
            # Export-Csv -Append rejects a row whose column set differs from the existing header.
            Error             = $null
        }) | Out-Null

        if ($IncludeEffectiveAccess -and $EffectiveRows) {
            foreach ($m in $members) {
                $EffectiveRows.Add([PSCustomObject]@{
                    SiteUrl          = $ScopeInfo.SiteUrl
                    WebUrl           = $ScopeInfo.WebUrl
                    ScopeType        = $ScopeInfo.ScopeType
                    ScopeTitle       = $ScopeInfo.ScopeTitle
                    ScopeUrl         = $ScopeInfo.ScopeUrl
                    UserDisplayName  = $m.DisplayName
                    UserPrincipalName= $m.Upn
                    UserEmail        = $m.Email
                    IsExternal       = $m.IsExternal
                    AccountEnabled   = $m.Enabled
                    PermissionLevels = ($levels -join '; ')
                    GrantedVia       = $(if ($principal.Kind -eq 'User') { 'Direct' } else { "$($principal.Kind): $($principal.Title)" })
                }) | Out-Null
            }
        }
    }
    return $rows
}

function Get-ScopeRoleAssignments {
    param([Parameter(Mandatory = $true)][string]$Uri)
    $expand = "?`$expand=Member,RoleDefinitionBindings&`$select=PrincipalId,Member/Id,Member/Title,Member/LoginName,Member/Email,Member/PrincipalType,RoleDefinitionBindings/Name,RoleDefinitionBindings/Id"
    return Get-SPCollection -Uri ($Uri + $expand)
}

function Get-ItemRoleAssignmentsParallel {
    # Per-item role assignment lookups dominate a -Scope Item run: a library with 400 uniquely
    # permissioned files needs 400 round trips no matter how the query is shaped (SharePoint has
    # no bulk "give me every unique scope with its assignments" endpoint). Dispatching them across
    # a small runspace pool is what makes a tenant-wide run finish in hours instead of days, while
    # staying well under the request rate that trips SharePoint's per-site throttle.
    param(
        [Parameter(Mandatory = $true)][object[]]$Items,
        [Parameter(Mandatory = $true)][string]$WebUrl,
        [Parameter(Mandatory = $true)][string]$ListId
    )

    $results = @{}
    if ($Items.Count -eq 0) { return $results }

    $resourceRoot = Get-ResourceRootFromUrl -Url $WebUrl
    $headers = Get-ResourceToken -Resource $resourceRoot
    $select  = "?`$expand=Member,RoleDefinitionBindings&`$select=PrincipalId,Member/Id,Member/Title,Member/LoginName,Member/Email,Member/PrincipalType,RoleDefinitionBindings/Name,RoleDefinitionBindings/Id"

    if ($Concurrency -le 1 -or $Items.Count -eq 1) {
        foreach ($item in $Items) {
            $uri = "{0}/_api/web/lists(guid'{1}')/items({2})/roleassignments{3}" -f $WebUrl, $ListId, $item.Id, $select
            try {
                $results[[string]$item.Id] = @(Get-SPCollection -Uri $uri)
            } catch {
                $results[[string]$item.Id] = $_.Exception.Message
            }
        }
        return $results
    }

    $workerScript = {
        param([string]$Uri, [hashtable]$Headers, [int]$TimeoutSec, [int]$MaxAttempts)
        $rows = @()
        $next = $Uri
        while ($next) {
            $ok = $false
            for ($attempt = 1; $attempt -le $MaxAttempts -and -not $ok; $attempt++) {
                try {
                    $resp = Invoke-RestMethod -Uri $next -Headers $Headers -TimeoutSec $TimeoutSec -ErrorAction Stop
                    if ($resp.value) { $rows += @($resp.value) }
                    $next = $null
                    if ($resp.'odata.nextLink') { $next = [string]$resp.'odata.nextLink' }
                    $ok = $true
                } catch {
                    $status = $null
                    try { if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode } } catch {}
                    # 404 only: the item was deleted between the sweep and this lookup, so "no
                    # permissions" is the honest answer. 401 and 403 are not — reporting those as
                    # an empty permission set would quietly claim an item nobody can reach is an
                    # item nobody has rights on.
                    if ($status -eq 404) {
                        return [PSCustomObject]@{ Success = $true; Rows = @() }
                    }
                    if ($status -in @(401, 403)) {
                        return [PSCustomObject]@{
                            Success      = $false
                            Rows         = @()
                            ErrorMessage = "HTTP $status reading role assignments - permissions for this item are unknown, not empty."
                        }
                    }
                    if ($attempt -eq $MaxAttempts) {
                        return [PSCustomObject]@{ Success = $false; Rows = @(); ErrorMessage = $_.Exception.Message }
                    }
                    $delay = $null
                    try {
                        if ($_.Exception.Response -and $_.Exception.Response.Headers) {
                            $h = $_.Exception.Response.Headers['Retry-After']
                            if ($h) { [void][int]::TryParse([string]$h, [ref]$delay) }
                        }
                    } catch {}
                    if (-not $delay -or $delay -le 0) { $delay = [Math]::Min([int][Math]::Pow(2, $attempt), 60) }
                    Start-Sleep -Seconds $delay
                }
            }
        }
        return [PSCustomObject]@{ Success = $true; Rows = $rows }
    }

    $poolSize     = [Math]::Min($Concurrency, $Items.Count)
    $runspacePool = [RunspaceFactory]::CreateRunspacePool(1, $poolSize)
    $runspacePool.Open()

    # Dispatch in waves rather than queueing one PowerShell instance per item up front. A library
    # can easily have tens of thousands of uniquely permissioned items, and holding a live
    # PowerShell object and pending async handle for every one of them costs far more memory than
    # the scan itself — the pool would still only run $Concurrency of them at a time.
    $waveSize = [Math]::Max($poolSize * 25, 100)
    $done     = 0

    try {
        for ($offset = 0; $offset -lt $Items.Count; $offset += $waveSize) {
            $wave    = $Items[$offset..([Math]::Min($offset + $waveSize - 1, $Items.Count - 1))]
            $workers = [System.Collections.Generic.List[object]]::new()

            # Re-read the token every wave, not once per library. Workers get a plain copy of the
            # bearer header and cannot see a later refresh, and a library with enough unique scopes
            # runs longer than a token lives — the tail of it would 401 with nothing to show why.
            $headers = Get-ResourceToken -Resource $resourceRoot

            foreach ($item in $wave) {
                # Spacing out dispatch reduces the burst rate against one site, which is what
                # SharePoint's throttle actually measures — the pool size alone does not do that.
                Start-Sleep -Milliseconds 40
                $uri = "{0}/_api/web/lists(guid'{1}')/items({2})/roleassignments{3}" -f $WebUrl, $ListId, $item.Id, $select

                $ps = [PowerShell]::Create()
                $ps.RunspacePool = $runspacePool
                [void]$ps.AddScript($workerScript)
                [void]$ps.AddParameter('Uri', $uri)
                [void]$ps.AddParameter('Headers', $headers)
                [void]$ps.AddParameter('TimeoutSec', $GraphTimeoutSec)
                [void]$ps.AddParameter('MaxAttempts', $MaxGraphRetry)

                $workers.Add([PSCustomObject]@{
                    PowerShell = $ps
                    Handle     = $ps.BeginInvoke()
                    ItemId     = [string]$item.Id
                }) | Out-Null
            }

            foreach ($worker in $workers) {
                $payload = $null
                try { $payload = $worker.PowerShell.EndInvoke($worker.Handle) } catch { $payload = $null } finally { $worker.PowerShell.Dispose() }

                if ($payload -and $payload.Success) {
                    $results[$worker.ItemId] = @($payload.Rows)
                } elseif ($payload -and $payload.ErrorMessage) {
                    $results[$worker.ItemId] = [string]$payload.ErrorMessage
                } else {
                    $results[$worker.ItemId] = 'Role assignment lookup failed after retries.'
                }

                $done++
                if (($done % 100) -eq 0 -or $done -eq $Items.Count) {
                    Set-ScanProgress -Id 3 -ParentId 2 -Activity 'Unieke rechten ophalen' -Status (
                        "{0} — {1}/{2} objecten" -f $script:CurrentScanLabel, $done, $Items.Count
                    ) -PercentComplete ([int](($done / [Math]::Max($Items.Count, 1)) * 100))
                }
            }
        }
    } finally {
        $runspacePool.Close()
        $runspacePool.Dispose()
    }

    return $results
}

# ── Resume / checkpoint state ────────────────────────────────────────────────
$checkpointSignature = Get-TextHashHex -Text (@"
$PSCommandPath
$outputDir
$SiteUrl
$TenantUrl
$Scope
$IncludeOneDriveSites
$IncludeHiddenLists
$($ListTitle -join ',')
$ExcludeLimitedAccess
$SkipGroupExpansion
$IncludeEffectiveAccess
$TenantId
$effectiveTenantId
$ClientId
$CertificateThumbprint
"@)

$script:CheckpointStatePath     = Join-Path $outputDir "SharePoint_Permissions_$checkpointSignature.state.json"
$script:CheckpointDetailPath    = Join-Path $outputDir "SharePoint_Permissions_$checkpointSignature.detail.partial.csv"
$script:CheckpointGroupsPath    = Join-Path $outputDir "SharePoint_Permissions_$checkpointSignature.groups.partial.csv"
$script:CheckpointEffectivePath = Join-Path $outputDir "SharePoint_Permissions_$checkpointSignature.effective.partial.csv"
$script:CompletedUnitKeys       = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$script:LoadedCheckpoint        = $false

# Completed keys live in an append-only log, not in the JSON. Rewriting a sorted list of every
# key after every finished list is quadratic: on a tenant with thousands of lists the last writes
# serialise thousands of keys each, and the checkpoint ends up costing more than the scanning.
$script:CheckpointKeysPath = Join-Path $outputDir "SharePoint_Permissions_$checkpointSignature.keys.partial.log"

$allCheckpointPaths = @(
    $script:CheckpointStatePath, $script:CheckpointKeysPath, $script:CheckpointDetailPath,
    $script:CheckpointGroupsPath, $script:CheckpointEffectivePath
)

if ($Restart) {
    $discarded = $false
    foreach ($path in $allCheckpointPaths) {
        if (Test-Path $path) {
            try { Remove-Item -Path $path -Force -ErrorAction Stop; $discarded = $true } catch {}
        }
    }
    if ($discarded) {
        Write-ProgressHost -Message "[INFO] -Restart specified: discarded existing checkpoint, starting from scratch." -ForegroundColor Yellow
    }
} elseif (Test-Path $script:CheckpointStatePath) {
    try {
        $checkpointState = Get-Content -Path $script:CheckpointStatePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($checkpointState.RunSignature -eq $checkpointSignature) {
            $script:LoadedCheckpoint = $true
            # Read the key log line by line rather than slurping it: it can hold hundreds of
            # thousands of lines, and a truncated last line (killed mid-write) must not abort the
            # resume — it just means that one unit gets scanned again, which is harmless.
            if (Test-Path $script:CheckpointKeysPath) {
                foreach ($line in [System.IO.File]::ReadLines($script:CheckpointKeysPath)) {
                    if ([string]::IsNullOrWhiteSpace($line)) { continue }
                    $kind, $key = $line.Split('|', 2)
                    if ([string]::IsNullOrWhiteSpace($key)) { continue }
                    switch ($kind) {
                        'U' { [void]$script:CompletedUnitKeys.Add($key) }
                        'G' { [void]$script:GroupRowsWritten.Add($key) }
                    }
                }
            }
        }
    } catch {
        $script:LoadedCheckpoint = $false
    }
}

function Save-CheckpointState {
    # Only the signature and a timestamp — the keys themselves are appended to the key log as they
    # happen, so this stays a constant-size write no matter how large the tenant is.
    $state = [PSCustomObject]@{
        Version      = 2
        RunSignature = $checkpointSignature
        UpdatedUtc   = (Get-Date).ToUniversalTime().ToString('o')
    }
    try {
        $state | ConvertTo-Json -Depth 4 | Set-Content -Path $script:CheckpointStatePath -Encoding UTF8 -ErrorAction Stop
    } catch {
        Write-ProgressHost -Message ("[WARN] Could not update the checkpoint state file: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}

function Add-CheckpointKey {
    param([ValidateSet('U', 'G')][string]$Kind, [string]$Key)
    if ([string]::IsNullOrWhiteSpace($Key)) { return }
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            Add-Content -Path $script:CheckpointKeysPath -Value ("{0}|{1}" -f $Kind, $Key) -Encoding UTF8 -ErrorAction Stop
            return
        } catch {
            if ($attempt -eq 5) {
                # Non-fatal by design, unlike a failed row write: a lost key means the unit is
                # re-scanned on resume, which costs time but never costs data.
                Write-ProgressHost -Message ("[WARN] Could not record checkpoint key — this unit will be re-scanned if the run is resumed.") -ForegroundColor Yellow
                return
            }
            Start-Sleep -Milliseconds (200 * $attempt)
        }
    }
}

function Complete-CheckpointUnit {
    # Rows are streamed to the partial CSVs as each unit finishes rather than held in memory: a
    # tenant-wide item-level run can produce millions of rows, and the version/storage reports'
    # habit of keeping everything until the end would not survive that.
    param(
        [string]$UnitKey,
        [object[]]$DetailRows = @(),
        [object[]]$GroupRows = @(),
        [object[]]$EffectiveRows = @(),
        [string[]]$GroupKeys = @()
    )
    # Order matters: rows first, keys second. A crash between the two re-scans the unit on resume,
    # which is wasteful but correct. The reverse order would mark a unit done whose rows never
    # landed, and that gap would never be noticed.
    if ($DetailRows.Count -gt 0)    { Append-CheckpointRows -Path $script:CheckpointDetailPath    -Rows $DetailRows }
    if ($GroupRows.Count -gt 0)     { Append-CheckpointRows -Path $script:CheckpointGroupsPath    -Rows $GroupRows }
    if ($EffectiveRows.Count -gt 0) { Append-CheckpointRows -Path $script:CheckpointEffectivePath -Rows $EffectiveRows }

    foreach ($groupKey in $GroupKeys) { Add-CheckpointKey -Kind 'G' -Key $groupKey }
    if (-not [string]::IsNullOrWhiteSpace($UnitKey)) {
        $script:CompletedUnitKeys.Add($UnitKey) | Out-Null
        Add-CheckpointKey -Kind 'U' -Key $UnitKey
    }

    # The state file only carries the signature, so refreshing it every few units is enough to
    # keep its timestamp meaningful without writing a file per list.
    $script:UnitsSinceStateSave++
    if ($script:UnitsSinceStateSave -ge 25) {
        Save-CheckpointState
        $script:UnitsSinceStateSave = 0
    }
}

$script:UnitsSinceStateSave = 0

# Write the state file up front, not on the first periodic save. It is what a later run matches
# the signature against, so a scan interrupted in its first few units must still be resumable.
Save-CheckpointState

if ($script:LoadedCheckpoint) {
    Write-ProgressHost -Message ("Resuming with {0} completed unit(s) from a previous run." -f $script:CompletedUnitKeys.Count) -ForegroundColor DarkGray
}

# ── Scan ──────────────────────────────────────────────────────────────────────
$runTotals = [PSCustomObject]@{
    Webs            = 0
    WebsFailed      = 0
    Lists           = 0
}

try {
    $webIndex = 0
    foreach ($web in $targetSites) {
        $webIndex++
        $webUrl = $web.WebUrl
        $siteCollectionUrl = Get-SiteCollectionUrl -WebUrl $webUrl

        Write-ProgressHost -Message ("[{0}/{1}] {2}" -f $webIndex, $targetSites.Count, $webUrl) -ForegroundColor White
        Set-ScanProgress -Id 1 -Activity 'Sites scannen' -Status ("[{0}/{1}] {2}" -f $webIndex, $targetSites.Count, $webUrl) `
            -PercentComplete ([int](($webIndex / [Math]::Max($targetSites.Count, 1)) * 100))

        try {
            $webUnitKey = "WEB|$webUrl"
            $webDetailRows    = [System.Collections.Generic.List[object]]::new()
            $webEffectiveRows = [System.Collections.Generic.List[object]]::new()
            $webGroupRows     = [System.Collections.Generic.List[object]]::new()
            # Group keys are only persisted once their rows are safely written, so an interrupted
            # run does not remember having emitted membership it never got round to writing.
            $webGroupKeys     = [System.Collections.Generic.List[string]]::new()

            $webInfo = Invoke-SPGet -Uri ("{0}/_api/web?`$select=Title,Url,WebTemplate,Created,HasUniqueRoleAssignments,LastItemModifiedDate" -f $webUrl)
            if (-not $webInfo) {
                Write-ProgressHost -Message "    [SKIP] Web not accessible with the current permissions." -ForegroundColor Yellow
                $runTotals.WebsFailed++
                continue
            }
            $runTotals.Webs++
            $webTitle = [string]$webInfo.Title

            if (-not ($script:LoadedCheckpoint -and $script:CompletedUnitKeys.Contains($webUnitKey))) {
                # ── Site collection administrators ────────────────────────────
                # Site collection admins bypass every role assignment below, so a permissions
                # report that omits them is actively misleading about who can reach the content.
                if ($webUrl.TrimEnd('/') -eq $siteCollectionUrl.TrimEnd('/')) {
                    $admins = Get-SPCollection -Uri ("{0}/_api/web/siteusers?`$select=Id,Title,LoginName,Email,PrincipalType,IsSiteAdmin&`$filter=IsSiteAdmin%20eq%20true" -f $webUrl)
                    foreach ($admin in $admins) {
                        $principal = Get-PrincipalInfo -Member $admin
                        $webDetailRows.Add([PSCustomObject]@{
                            SiteUrl           = $siteCollectionUrl
                            WebUrl            = $webUrl
                            WebTitle          = $webTitle
                            ScopeType         = 'SiteCollection'
                            ScopeTitle        = $webTitle
                            ScopeUrl          = $siteCollectionUrl
                            ItemType          = 'SiteCollectionAdmin'
                            ListTitle         = $null
                            ListTemplate      = $null
                            HasUniquePerms    = $true
                            InheritsFrom      = $null
                            PrincipalType     = $principal.Kind
                            PrincipalName     = $principal.Title
                            PrincipalLogin    = $principal.LoginName
                            PrincipalEmail    = $principal.Email
                            DirectoryObjectId = $principal.DirectoryId
                            PermissionLevels  = 'Site Collection Administrator'
                            IsSharingLink     = $false
                            SharingLinkType   = $null
                            IsExternal        = $principal.IsExternal
                            MemberCount       = 1
                            ExternalMembers   = $(if ($principal.IsExternal) { 1 } else { 0 })
                            MemberPreview     = $principal.Email
                            LastModified      = $null
                            ScannedUtc        = (Get-Date).ToUniversalTime().ToString('s')
                            Error             = $null
                        }) | Out-Null
                    }
                }

                # ── SharePoint groups and their membership ────────────────────
                # Written once per site collection, keyed in the checkpoint, so sub-webs sharing
                # the same groups do not repeat thousands of membership rows.
                $siteGroups = Get-SiteGroups -WebUrl $webUrl
                foreach ($groupId in $siteGroups.Keys) {
                    $group    = $siteGroups[$groupId]
                    $groupKey = "$siteCollectionUrl|$groupId"
                    if (-not $script:GroupRowsWritten.Add($groupKey)) { continue }
                    $webGroupKeys.Add($groupKey) | Out-Null

                    $groupUsers = @($group.Users)
                    if ($groupUsers.Count -eq 0) {
                        $webGroupRows.Add([PSCustomObject]@{
                            SiteUrl        = $siteCollectionUrl
                            GroupId        = $groupId
                            GroupTitle     = $group.Title
                            GroupOwner     = $group.OwnerTitle
                            GroupType      = 'SharePointGroup'
                            MemberType     = $null
                            MemberName     = $null
                            MemberLogin    = $null
                            MemberEmail    = $null
                            MemberIsExternal = $null
                            IsEmpty        = $true
                        }) | Out-Null
                        continue
                    }
                    foreach ($u in $groupUsers) {
                        $memberPrincipal = Get-PrincipalInfo -Member $u
                        $webGroupRows.Add([PSCustomObject]@{
                            SiteUrl        = $siteCollectionUrl
                            GroupId        = $groupId
                            GroupTitle     = $group.Title
                            GroupOwner     = $group.OwnerTitle
                            GroupType      = 'SharePointGroup'
                            MemberType     = $memberPrincipal.Kind
                            MemberName     = $memberPrincipal.Title
                            MemberLogin    = $memberPrincipal.LoginName
                            MemberEmail    = $memberPrincipal.Email
                            MemberIsExternal = $memberPrincipal.IsExternal
                            IsEmpty        = $false
                        }) | Out-Null

                        # A SharePoint group containing an Entra group is the most common way real
                        # membership hides from a site owner — flatten it here too.
                        if ($memberPrincipal.Kind -in @('SecurityGroup', 'M365Group', 'M365GroupOwners')) {
                            foreach ($nested in (Resolve-PrincipalMembers -Principal $memberPrincipal -WebUrl $webUrl)) {
                                $webGroupRows.Add([PSCustomObject]@{
                                    SiteUrl        = $siteCollectionUrl
                                    GroupId        = $groupId
                                    GroupTitle     = $group.Title
                                    GroupOwner     = $group.OwnerTitle
                                    GroupType      = "via $($memberPrincipal.Kind): $($memberPrincipal.Title)"
                                    MemberType     = 'User'
                                    MemberName     = $nested.DisplayName
                                    MemberLogin    = $nested.Upn
                                    MemberEmail    = $nested.Email
                                    MemberIsExternal = $nested.IsExternal
                                    IsEmpty        = $false
                                }) | Out-Null
                            }
                        }
                    }
                }

                # ── Web level role assignments ────────────────────────────────
                $webScope = @{
                    SiteUrl        = $siteCollectionUrl
                    WebUrl         = $webUrl
                    WebTitle       = $webTitle
                    ScopeType      = 'Web'
                    ScopeTitle     = $webTitle
                    ScopeUrl       = $webUrl
                    ItemType       = [string]$webInfo.WebTemplate
                    ListTitle      = $null
                    ListTemplate   = $null
                    HasUniquePerms = [bool]$webInfo.HasUniqueRoleAssignments
                    InheritsFrom   = $(if ([bool]$webInfo.HasUniqueRoleAssignments) { $null } else { 'Parent web' })
                    LastModified   = $webInfo.LastItemModifiedDate
                }
                $webRoleAssignments = Get-ScopeRoleAssignments -Uri ("{0}/_api/web/roleassignments" -f $webUrl)
                foreach ($row in (ConvertTo-PermissionRows -RoleAssignments @($webRoleAssignments) -ScopeInfo $webScope -EffectiveRows $webEffectiveRows)) {
                    $webDetailRows.Add($row) | Out-Null
                }

                Complete-CheckpointUnit -UnitKey $webUnitKey -DetailRows @($webDetailRows) `
                    -GroupRows @($webGroupRows) -EffectiveRows @($webEffectiveRows) `
                    -GroupKeys @($webGroupKeys)
            } else {
                Write-ProgressHost -Message "    [SKIP] Web level already completed in a previous run." -ForegroundColor DarkGray
            }

            if ($Scope -eq 'Site') { continue }

            # ── Lists and libraries ───────────────────────────────────────────
            $lists = @(Get-SPCollection -Uri ("{0}/_api/web/lists?`$select=Id,Title,Hidden,BaseTemplate,BaseType,ItemCount,HasUniqueRoleAssignments,EntityTypeName,RootFolder/ServerRelativeUrl&`$expand=RootFolder" -f $webUrl))
            if (-not $IncludeHiddenLists) { $lists = @($lists | Where-Object { -not $_.Hidden }) }
            if ($ListTitle.Count -gt 0)   { $lists = @($lists | Where-Object { $_.Title -in $ListTitle }) }

            $listIndex = 0
            foreach ($list in $lists) {
                $listIndex++
                $listId  = [string]$list.Id
                $unitKey = "LIST|$webUrl|$listId"
                $script:CurrentScanLabel = "$webUrl > $($list.Title)"

                Set-ScanProgress -Id 2 -ParentId 1 -Activity 'Lijsten en bibliotheken scannen' `
                    -Status ("[{0}/{1}] {2}" -f $listIndex, $lists.Count, $script:CurrentScanLabel) `
                    -PercentComplete ([int](($listIndex / [Math]::Max($lists.Count, 1)) * 100))

                if ($script:LoadedCheckpoint -and $script:CompletedUnitKeys.Contains($unitKey)) {
                    Write-ProgressHost -Message ("    [SKIP] {0} — already completed in a previous run." -f $list.Title) -ForegroundColor DarkGray
                    continue
                }

                $runTotals.Lists++
                $listDetailRows    = [System.Collections.Generic.List[object]]::new()
                $listEffectiveRows = [System.Collections.Generic.List[object]]::new()

                # Everything from here to the checkpoint runs inside its own try. One list that
                # throws — a template that does not answer /roleassignments, a view threshold, a
                # transient server error — must cost that list, not the rest of the web behind it.
                try {

                $listUrl = if ($list.RootFolder -and $list.RootFolder.ServerRelativeUrl) {
                    ("{0}://{1}{2}" -f ([Uri]$webUrl).Scheme, ([Uri]$webUrl).Host, $list.RootFolder.ServerRelativeUrl)
                } else { $webUrl }

                # A list that inherits is reported through its parent web, not duplicated here —
                # that is what makes the CSV a map of the permission structure rather than a dump.
                if ([bool]$list.HasUniqueRoleAssignments) {
                    $listScope = @{
                        SiteUrl        = $siteCollectionUrl
                        WebUrl         = $webUrl
                        WebTitle       = $webTitle
                        ScopeType      = 'List'
                        ScopeTitle     = [string]$list.Title
                        ScopeUrl       = $listUrl
                        ItemType       = $(if ([int]$list.BaseType -eq 1) { 'DocumentLibrary' } else { 'List' })
                        ListTitle      = [string]$list.Title
                        ListTemplate   = [string]$list.BaseTemplate
                        HasUniquePerms = $true
                        InheritsFrom   = $null
                        LastModified   = $null
                    }
                    $listRoleAssignments = Get-ScopeRoleAssignments -Uri ("{0}/_api/web/lists(guid'{1}')/roleassignments" -f $webUrl, $listId)
                    foreach ($row in (ConvertTo-PermissionRows -RoleAssignments @($listRoleAssignments) -ScopeInfo $listScope -EffectiveRows $listEffectiveRows)) {
                        $listDetailRows.Add($row) | Out-Null
                    }
                }

                # ── Folders, files and list items with unique permissions ─────
                if ($Scope -eq 'Item' -and [int]$list.ItemCount -gt 0) {
                    # Ask for HasUniqueRoleAssignments on every item in one paged sweep, then only
                    # fetch role assignments for the items that actually have their own scope.
                    # Paging by $top uses $skiptoken under the hood, so this does not trip the
                    # 5000-item list view threshold the way a $filter or $orderby query would.
                    $itemsUri = "{0}/_api/web/lists(guid'{1}')/items?`$select=Id,FileSystemObjectType,HasUniqueRoleAssignments,FileRef,FileLeafRef,Title,Modified&`$top=2000" -f $webUrl, $listId

                    # Only the items that actually have their own scope are kept. Materialising a
                    # whole library first would mean holding every one of a million rows in memory
                    # to discard nearly all of them — the unique ones are typically a rounding error.
                    $uniqueItems = [System.Collections.Generic.List[object]]::new()
                    # A plain counter would not survive: & $scriptblock runs in a child scope, so
                    # "$n += 1" inside -OnPage writes to a local copy and the outer stays zero.
                    # Property writes on an object do reach back, so the tally lives on one.
                    $sweep = [PSCustomObject]@{ Scanned = 0 }
                    $itemSweepFailed = $null
                    try {
                        Invoke-SPCollectionPaged -Uri $itemsUri -OnPage {
                            param($PageItems)
                            $sweep.Scanned += $PageItems.Count
                            foreach ($it in $PageItems) {
                                if ([bool]$it.HasUniqueRoleAssignments) { $uniqueItems.Add($it) | Out-Null }
                            }
                            Set-ScanProgress -Id 3 -ParentId 2 -Activity 'Items doorzoeken op eigen rechten' -Status (
                                "{0} — {1} items, {2} met eigen rechten" -f $script:CurrentScanLabel, $sweep.Scanned, $uniqueItems.Count)
                        }
                    } catch {
                        $itemSweepFailed = $_.Exception.Message
                        Write-ProgressHost -Message ("    [WARN] Cannot enumerate items in {0}: {1}" -f $list.Title, $itemSweepFailed) -ForegroundColor Yellow
                    }

                    # A failed sweep is a hole in the report, so it goes in the CSV. Otherwise the
                    # list simply looks like it had nothing with unique permissions.
                    if ($itemSweepFailed) {
                        $listDetailRows.Add([PSCustomObject]@{
                            SiteUrl = $siteCollectionUrl; WebUrl = $webUrl; WebTitle = $webTitle
                            ScopeType = 'List'; ScopeTitle = [string]$list.Title; ScopeUrl = $listUrl
                            ItemType = 'Error'; ListTitle = [string]$list.Title; ListTemplate = [string]$list.BaseTemplate
                            HasUniquePerms = $null; InheritsFrom = $null
                            PrincipalType = $null; PrincipalName = $null; PrincipalLogin = $null; PrincipalEmail = $null
                            DirectoryObjectId = $null; PermissionLevels = $null
                            IsSharingLink = $false; SharingLinkType = $null; IsExternal = $null
                            MemberCount = 0; ExternalMembers = 0; MemberPreview = $null
                            LastModified = $null
                            ScannedUtc = (Get-Date).ToUniversalTime().ToString('s')
                            Error = "Item sweep failed, items with unique permissions were not checked: $itemSweepFailed"
                        }) | Out-Null
                    }

                    if ($uniqueItems.Count -gt 0) {
                        Write-ProgressHost -Message ("    {0}: {1} item(s) scanned, {2} with unique permissions" -f $list.Title, $sweep.Scanned, $uniqueItems.Count)
                        $itemAssignments = Get-ItemRoleAssignmentsParallel -Items $uniqueItems -WebUrl $webUrl -ListId $listId

                        foreach ($item in $uniqueItems) {
                            $assignments = $itemAssignments[[string]$item.Id]
                            if ($assignments -is [string]) {
                                $listDetailRows.Add([PSCustomObject]@{
                                    SiteUrl = $siteCollectionUrl; WebUrl = $webUrl; WebTitle = $webTitle
                                    ScopeType = 'Item'; ScopeTitle = [string]$item.FileLeafRef; ScopeUrl = [string]$item.FileRef
                                    ItemType = 'Error'; ListTitle = [string]$list.Title; ListTemplate = [string]$list.BaseTemplate
                                    HasUniquePerms = $true; InheritsFrom = $null
                                    PrincipalType = $null; PrincipalName = $null; PrincipalLogin = $null; PrincipalEmail = $null
                                    DirectoryObjectId = $null; PermissionLevels = $null
                                    IsSharingLink = $false; SharingLinkType = $null; IsExternal = $null
                                    MemberCount = 0; ExternalMembers = 0; MemberPreview = $null
                                    LastModified = $item.Modified
                                    ScannedUtc = (Get-Date).ToUniversalTime().ToString('s')
                                    Error = $assignments
                                }) | Out-Null
                                continue
                            }

                            $itemScope = @{
                                SiteUrl        = $siteCollectionUrl
                                WebUrl         = $webUrl
                                WebTitle       = $webTitle
                                ScopeType      = 'Item'
                                ScopeTitle     = [string]$item.FileLeafRef
                                ScopeUrl       = [string]$item.FileRef
                                ItemType       = $(switch ([int]$item.FileSystemObjectType) { 1 { 'Folder' } 0 { 'File' } default { 'ListItem' } })
                                ListTitle      = [string]$list.Title
                                ListTemplate   = [string]$list.BaseTemplate
                                HasUniquePerms = $true
                                InheritsFrom   = $null
                                LastModified   = $item.Modified
                            }
                            foreach ($row in (ConvertTo-PermissionRows -RoleAssignments @($assignments) -ScopeInfo $itemScope -EffectiveRows $listEffectiveRows)) {
                                $listDetailRows.Add($row) | Out-Null
                            }
                        }
                    }
                }

                # Sharing-link, external and Everyone tallies are deliberately not accumulated
                # here: the closing summary derives them from the detail CSV, which also covers
                # the rows a resumed run wrote in an earlier session. Counting them twice would
                # mean three more passes over every row of every list for a number nobody reads.
                Complete-CheckpointUnit -UnitKey $unitKey -DetailRows @($listDetailRows) -EffectiveRows @($listEffectiveRows)

                } catch {
                    # A 401 means the credential itself stopped working, which is not survivable
                    # and not specific to this list — let it out so the run stops loudly rather
                    # than filling the CSV with one error row per remaining list.
                    if ($_.Exception.Message -match 'SharePoint refused the token \(401\)') { throw }

                    Write-ProgressHost -Message ("    [ERROR] {0} failed: {1}" -f $list.Title, $_.Exception.Message) -ForegroundColor Red
                    # Whatever this list did produce before failing is still written, and the
                    # unit is deliberately left unmarked so a resumed run retries it.
                    Complete-CheckpointUnit -UnitKey $null -EffectiveRows @($listEffectiveRows) -DetailRows (
                        @($listDetailRows) + @([PSCustomObject]@{
                            SiteUrl = $siteCollectionUrl; WebUrl = $webUrl; WebTitle = $webTitle
                            ScopeType = 'List'; ScopeTitle = [string]$list.Title; ScopeUrl = $webUrl
                            ItemType = 'Error'; ListTitle = [string]$list.Title; ListTemplate = [string]$list.BaseTemplate
                            HasUniquePerms = $null; InheritsFrom = $null
                            PrincipalType = $null; PrincipalName = $null; PrincipalLogin = $null; PrincipalEmail = $null
                            DirectoryObjectId = $null; PermissionLevels = $null
                            IsSharingLink = $false; SharingLinkType = $null; IsExternal = $null
                            MemberCount = 0; ExternalMembers = 0; MemberPreview = $null
                            LastModified = $null
                            ScannedUtc = (Get-Date).ToUniversalTime().ToString('s')
                            Error = $_.Exception.Message
                        })
                    )
                }
            }

            Complete-ScanProgress -Id 3 -ParentId 2
            Complete-ScanProgress -Id 2 -ParentId 1
        } catch {
            # Same reasoning as the per-list handler, one level up: a refused token is not a
            # property of this web. Letting it be recorded here would walk the whole tenant
            # writing an error row per site — which is exactly the failure mode this scan already
            # had once, and the reason it looked like a permissions finding instead of a bug.
            if ($_.Exception.Message -match 'SharePoint refused the token \(401\)') { throw }

            $runTotals.WebsFailed++
            Write-ProgressHost -Message ("[ERROR] Web failed: {0} — {1}" -f $webUrl, $_.Exception.Message) -ForegroundColor Red
            # Recorded in the detail file rather than a separate error file: a web that could not
            # be read is a gap in the permissions picture, and it belongs where someone reviewing
            # that site will actually see it.
            Complete-CheckpointUnit -UnitKey $null -DetailRows @([PSCustomObject]@{
                SiteUrl = $siteCollectionUrl; WebUrl = $webUrl; WebTitle = $web.Title
                ScopeType = 'Web'; ScopeTitle = $web.Title; ScopeUrl = $webUrl
                ItemType = 'Error'; ListTitle = $null; ListTemplate = $null
                HasUniquePerms = $null; InheritsFrom = $null
                PrincipalType = $null; PrincipalName = $null; PrincipalLogin = $null; PrincipalEmail = $null
                DirectoryObjectId = $null; PermissionLevels = $null
                IsSharingLink = $false; SharingLinkType = $null; IsExternal = $null
                MemberCount = 0; ExternalMembers = 0; MemberPreview = $null
                LastModified = $null
                ScannedUtc = (Get-Date).ToUniversalTime().ToString('s')
                Error = $_.Exception.Message
            })
        }
    }
} finally {
    1, 2, 3 | ForEach-Object { Complete-ScanProgress -Id $_ }
    Remove-TempApp
}

# ── Summary ───────────────────────────────────────────────────────────────────
# Built by streaming the detail file back rather than from in-memory rows, so a resumed run
# summarises everything it has ever written for this signature, not just this session's share.
function New-PermissionSummary {
    param([string]$DetailPath)

    $bySite = [ordered]@{}
    if (-not (Test-Path $DetailPath)) { return @() }

    Import-Csv -Path $DetailPath | ForEach-Object {
        $key = [string]$_.SiteUrl
        if (-not $bySite.Contains($key)) {
            $bySite[$key] = [PSCustomObject]@{
                SiteUrl                  = $key
                Grants                   = 0
                UniqueScopes             = 0
                Webs                     = 0
                Lists                    = 0
                FoldersWithUniquePerms   = 0
                FilesWithUniquePerms     = 0
                SiteCollectionAdmins     = 0
                SharingLinks             = 0
                AnonymousLinks           = 0
                ExternalPrincipals       = 0
                EveryoneGrants           = 0
                DistinctPrincipals       = 0
                Errors                   = 0
            }
            $bySite[$key] | Add-Member -NotePropertyName '_webs'       -NotePropertyValue ([System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)) -Force
            $bySite[$key] | Add-Member -NotePropertyName '_lists'      -NotePropertyValue ([System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)) -Force
            $bySite[$key] | Add-Member -NotePropertyName '_scopes'     -NotePropertyValue ([System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)) -Force
            $bySite[$key] | Add-Member -NotePropertyName '_principals' -NotePropertyValue ([System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)) -Force
        }
        $entry = $bySite[$key]

        # Error rows record a gap in coverage, not an access grant — counting them as grants
        # would quietly inflate exactly the number someone uses to judge the site.
        if ($_.ItemType -eq 'Error') {
            $entry.Errors++
            if ($_.WebUrl) { [void]$entry._webs.Add([string]$_.WebUrl) }
            return
        }

        $entry.Grants++
        if ($_.WebUrl)         { [void]$entry._webs.Add([string]$_.WebUrl) }
        if ($_.ListTitle)      { [void]$entry._lists.Add("$($_.WebUrl)|$($_.ListTitle)") }
        if ($_.ScopeUrl)       { [void]$entry._scopes.Add("$($_.ScopeType)|$($_.ScopeUrl)") }
        if ($_.PrincipalLogin) { [void]$entry._principals.Add([string]$_.PrincipalLogin) }

        if ($_.ItemType -eq 'SiteCollectionAdmin') { $entry.SiteCollectionAdmins++ }
        if ($_.ItemType -eq 'Folder')              { $entry.FoldersWithUniquePerms++ }
        if ($_.ItemType -eq 'File')                { $entry.FilesWithUniquePerms++ }
        if ($_.IsSharingLink -eq 'True') {
            $entry.SharingLinks++
            if ([string]$_.SharingLinkType -match 'anonymous') { $entry.AnonymousLinks++ }
        }
        if ($_.IsExternal -eq 'True' -or ([int]([string]$_.ExternalMembers -as [int]) -gt 0)) { $entry.ExternalPrincipals++ }
        if ($_.PrincipalType -in @('Everyone', 'EveryoneExceptExternalUsers', 'AllAuthenticatedUsers')) { $entry.EveryoneGrants++ }
    }

    $summary = foreach ($key in $bySite.Keys) {
        $entry = $bySite[$key]
        $entry.Webs               = $entry._webs.Count
        $entry.Lists              = $entry._lists.Count
        $entry.UniqueScopes       = $entry._scopes.Count
        $entry.DistinctPrincipals = $entry._principals.Count
        $entry | Select-Object SiteUrl, Webs, Lists, UniqueScopes, Grants, DistinctPrincipals,
                               FoldersWithUniquePerms, FilesWithUniquePerms, SiteCollectionAdmins,
                               SharingLinks, AnonymousLinks, ExternalPrincipals, EveryoneGrants, Errors
    }
    return @($summary)
}

$summaryRows = @(New-PermissionSummary -DetailPath $script:CheckpointDetailPath)

# ── Write final CSVs ──────────────────────────────────────────────────────────
function Publish-ReportFile {
    param([string]$PartialPath, [string]$FinalPath, [string]$Label)
    if (-not (Test-Path $PartialPath)) {
        Write-Host ("  {0,-18}: (no rows)" -f $Label) -ForegroundColor DarkGray
        return $false
    }
    Copy-Item -Path $PartialPath -Destination $FinalPath -Force
    Write-Host ("  {0,-18}: {1}" -f $Label, $FinalPath) -ForegroundColor Green
    return $true
}

Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host '   Summary' -ForegroundColor Cyan
Write-Host '  ================================================' -ForegroundColor Cyan

if ($summaryRows.Count -gt 0) {
    $summaryRows | Export-Csv -Path $summaryCsv -NoTypeInformation -Encoding UTF8
    Write-Host ("  {0,-18}: {1}" -f 'Summary', $summaryCsv) -ForegroundColor Green
} else {
    Write-Host ("  {0,-18}: (no rows)" -f 'Summary') -ForegroundColor DarkGray
}
[void](Publish-ReportFile -PartialPath $script:CheckpointDetailPath    -FinalPath $detailCsv    -Label 'Detail')
[void](Publish-ReportFile -PartialPath $script:CheckpointGroupsPath    -FinalPath $groupsCsv    -Label 'Groups')
if ($IncludeEffectiveAccess) {
    [void](Publish-ReportFile -PartialPath $script:CheckpointEffectivePath -FinalPath $effectiveCsv -Label 'Effective access')
}

# Checkpoint files are only removed once the finals are on disk — an interrupted run keeps them
# so the next invocation resumes instead of re-scanning the tenant from the start.
foreach ($path in $allCheckpointPaths) {
    if (Test-Path $path) {
        try { Remove-Item -Path $path -Force -ErrorAction Stop } catch {}
    }
}

$totalGrants       = ($summaryRows | Measure-Object -Property Grants -Sum).Sum
$totalUniqueScopes = ($summaryRows | Measure-Object -Property UniqueScopes -Sum).Sum
$totalSharingLinks = ($summaryRows | Measure-Object -Property SharingLinks -Sum).Sum
$totalAnonymous    = ($summaryRows | Measure-Object -Property AnonymousLinks -Sum).Sum
$totalExternal     = ($summaryRows | Measure-Object -Property ExternalPrincipals -Sum).Sum
$totalEveryone     = ($summaryRows | Measure-Object -Property EveryoneGrants -Sum).Sum
$totalErrors       = ($summaryRows | Measure-Object -Property Errors -Sum).Sum
if (-not $totalErrors)       { $totalErrors = 0 }
if (-not $totalGrants)       { $totalGrants = 0 }
if (-not $totalUniqueScopes) { $totalUniqueScopes = 0 }
if (-not $totalSharingLinks) { $totalSharingLinks = 0 }
if (-not $totalAnonymous)    { $totalAnonymous = 0 }
if (-not $totalExternal)     { $totalExternal = 0 }
if (-not $totalEveryone)     { $totalEveryone = 0 }

Write-Host ''
Write-Host ("  Webs scanned      : {0}{1}" -f $runTotals.Webs, $(if ($runTotals.WebsFailed) { " ({0} failed/inaccessible)" -f $runTotals.WebsFailed } else { '' })) -ForegroundColor Cyan
Write-Host ("  Lists scanned     : {0}" -f $runTotals.Lists) -ForegroundColor Cyan
Write-Host ("  Unique scopes     : {0}" -f $totalUniqueScopes) -ForegroundColor Cyan
Write-Host ("  Permission grants : {0}" -f $totalGrants) -ForegroundColor Cyan
Write-Host ("  Sharing links     : {0}{1}" -f $totalSharingLinks, $(if ($totalAnonymous) { " ({0} anonymous)" -f $totalAnonymous } else { '' })) -ForegroundColor $(if ($totalAnonymous -gt 0) { 'Yellow' } else { 'Cyan' })
Write-Host ("  External access   : {0} grant(s)" -f $totalExternal) -ForegroundColor $(if ($totalExternal -gt 0) { 'Yellow' } else { 'Cyan' })
Write-Host ("  Everyone grants   : {0}" -f $totalEveryone) -ForegroundColor $(if ($totalEveryone -gt 0) { 'Yellow' } else { 'Cyan' })

# Coverage before conclusions: a report with gaps in it must say so, or the absence of a finding
# reads as the absence of a risk.
if ($totalErrors -gt 0 -or $runTotals.WebsFailed -gt 0) {
    Write-Host ''
    Write-Host ("  [WARN] {0} scope(s) could not be read — the report is incomplete." -f $totalErrors) -ForegroundColor Yellow
    Write-Host "         Filter the detail CSV on ItemType = 'Error' to see exactly what was missed." -ForegroundColor Yellow
    Write-Host "         Re-running with the same parameters resumes and retries them." -ForegroundColor Yellow
} else {
    Write-Host ''
    Write-Host "  [OK]   Every targeted scope was read without errors." -ForegroundColor Green
}
if ($SkipGroupExpansion) {
    Write-Host '  [NOTE] -SkipGroupExpansion was used: group membership was not resolved.' -ForegroundColor DarkYellow
}
if (-not $IncludeEffectiveAccess) {
    Write-Host '  [NOTE] Add -IncludeEffectiveAccess for a per-user "who can reach what" CSV.' -ForegroundColor DarkGray
}
Write-Host ''
