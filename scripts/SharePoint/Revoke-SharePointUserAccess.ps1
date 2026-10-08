#Requires -Version 5.1
<#
.SYNOPSIS
    Revoke one user's access to SharePoint Online everywhere it is granted — site collection
    admin, role assignments at every level, SharePoint group membership and sharing links.

.DESCRIPTION
    The counterpart to Get-SharePointPermissionsReport.ps1: that one answers who can reach what,
    this one takes it away. It finds every place a named user holds access and removes it:

      * Site collection administrator
      * Role assignments granted directly to them, on the site, a sub-site, a list or library,
        a folder or a single file
      * Membership of SharePoint groups (Owners, Members, Visitors and custom ones)
      * Sharing links — the SharingLinks.* groups a shared link puts its recipients in, which is
        how "Anyone with the link" and "Specific people" actually carry a person

    Reporting is the default. Nothing is changed until -Apply is given, and every run writes a
    CSV of exactly what was found and what happened to it.

    What it deliberately does NOT do:

      * It does not change Entra ID group membership unless -RemoveFromEntraGroups is given. A
        user who reaches a site through a security group or a Microsoft 365 group keeps that
        access otherwise, and removing them from SharePoint will not take it away — the group is
        the grant. Those routes are reported, loudly, with the group named, so the access is not
        silently believed to be gone. Use -IncludeGroupAccess to have them listed even where the
        user has no SharePoint-level grant at all. With -RemoveFromEntraGroups, only the groups
        this run actually saw granting access are touched, never every group the user is in.
      * It does not touch grants to Everyone, Everyone except external users, or authenticated
        users. Removing one of those revokes access for the whole tenant, not for this person.
        They are reported for the same reason.
      * It leaves ownership and authorship metadata alone. A revoked user stays the author of the
        documents they created.

    Authentication is identical to the permissions report: a short-lived certificate-backed app
    registration with SharePoint Sites.FullControl.All, deleted again when the run ends. A client
    secret cannot work — SharePoint Online refuses secret-based app-only tokens, and a delegated
    Graph token is the wrong audience for SharePoint REST, where role assignments, SharePoint
    groups and site collection admins live (Graph has no API for them). The one interactive
    step, creating that app, is delegated through scripts\Startup\Connect-M365.ps1 (PowerShell 7):
    a device code when $global:useDeviceCodeAuth is set in load.config.ps1, the GDAP customer
    from Connect-Tenant. -ClientId + -CertificateThumbprint, or -AppOnly (graph.appid.json), use an
    app of your own instead.

.PARAMETER UserPrincipalName
    The user to revoke, for example jan@contoso.com. For a guest, either their UPN in this tenant
    (jan_partner.com#ext#@contoso.onmicrosoft.com) or their real address (jan@partner.com).

.PARAMETER TenantUrl
    Tenant root URL, for example https://contoso.sharepoint.com. Required for a tenant-wide run.

.PARAMETER SiteUrl
    Optional. Revoke within a single site collection (including its sub-sites) instead of the
    whole tenant.

.PARAMETER Apply
    Actually revoke. Without it the script only reports what it would remove.

.PARAMETER TenantId
    Entra ID tenant ID. Detected from the connected account when omitted. Required with -ClientId.

.PARAMETER ClientId
    Existing App Registration client ID. Skips the temporary app. Use with -TenantId and
    -CertificateThumbprint.

.PARAMETER ClientSecret
    Client secret for an existing app registration. Works for the Graph calls but NOT for the
    SharePoint ones — SharePoint Online rejects secret-based app-only tokens.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for an existing app registration, from Cert:\CurrentUser\My or
    Cert:\LocalMachine\My. This is the supported way to authenticate an existing app.

.PARAMETER AppOnly
    Use the app registration for the tenant in graph.appid.json (ClientId + CertificateThumbprint)
    instead of creating a temporary one. It needs the same roles as the temporary app: SharePoint
    Sites.FullControl.All and Graph Sites.Read.All, GroupMember.Read.All and User.Read.All.
    PowerShell 7 only.

.PARAMETER OutputPath
    Override the default output folder (C:\Temp on Windows).

.PARAMETER Scope
    How deep to look for direct grants. Site = site and sub-sites; List = also lists and
    libraries; Item = also folders, files and list items with unique permissions (default).

.PARAMETER IncludeOneDriveSites
    Also search personal OneDrive sites. Off by default — these add one site per user, and a
    leaver's own OneDrive is usually handled separately.

.PARAMETER IncludeHiddenLists
    Also search hidden and system lists.

.PARAMETER FromReport
    Take the webs to visit from a Get-SharePointPermissionsReport.ps1 run instead of walking the
    tenant again. Point it at the detail CSV, any other file from the same run, or the folder
    they are in.

    This is the pairing between the two scripts: the report answers who can reach what, you read
    it and decide, and the revoke acts on exactly the webs you were looking at. On a tenant where
    a full sweep takes a quarter of an hour, a user with access to a handful of sites is revoked
    in seconds.

    The report decides where to look, never what to remove. Every web it names is still read
    live, so a grant that disappeared between the two runs is reported as already gone rather
    than failing, and one that was removed by hand is not resurrected. The reverse does not hold:
    anything granted *after* the report was written is invisible here, and so is anything the
    report itself could not read — both are named in the summary.

    Without -TenantUrl the tenant is taken from the report.

.PARAMETER IncludeGroupAccess
    Also report the sites the user reaches through Entra ID groups, including sites where they
    have no SharePoint-level grant at all. Reported only; never revoked.

.PARAMETER KeepSharingLinks
    Leave sharing-link groups alone. The user keeps access through any link already shared with
    them; every other route is still revoked.

.PARAMETER RemoveFromEntraGroups
    Also remove the user from the Entra ID groups that were found granting access, completing
    the second half of an offboarding instead of only reporting it.

    Only the groups this run actually caught holding a role assignment on a scope in range are
    touched — never every group the user belongs to. Even so, an Entra group is not a SharePoint
    object: the same membership commonly carries Teams, a mailbox, licences and app assignments,
    so removing someone from one reaches well beyond anything this report can see. Read the
    report first, then re-run with this.

    Four cases are reported rather than forced, because forcing them would either fail or do
    the wrong thing: a dynamic group (membership follows a rule, so there is nothing to remove),
    a group synced from on-premises AD (read-only in the cloud), a membership inherited through
    a nested group (the access has to be cut at the group that actually holds the user), and a
    user who could not be resolved in Entra at all.

    Needs Graph GroupMember.ReadWrite.All, which the temporary app only asks for when this
    switch is given.

.PARAMETER RemoveFromSite
    After revoking, also remove the user from each site collection's user list. This clears any
    grant the per-scope pass could not see, but it also drops their entry from the site — people
    pickers stop suggesting them, and their name in older metadata renders as a deleted account.
    Off by default because the targeted removals above are enough in almost every case.

.PARAMETER GraphTimeoutSec
    Timeout in seconds per Graph/SharePoint call (default: 120).

.PARAMETER MaxGraphRetry
    Max retries on throttling/timeouts (default: 6).

.EXAMPLE
    .\Revoke-SharePointUserAccess.ps1 -UserPrincipalName jan@contoso.com -TenantUrl "https://contoso.sharepoint.com"

    Report everything Jan can reach, tenant-wide. Changes nothing.

.EXAMPLE
    .\Revoke-SharePointUserAccess.ps1 -UserPrincipalName jan@contoso.com -TenantUrl "https://contoso.sharepoint.com" -Apply

    The same, and actually revoke it.

.EXAMPLE
    .\Revoke-SharePointUserAccess.ps1 -UserPrincipalName gast@partner.com -SiteUrl "https://contoso.sharepoint.com/sites/Finance" -Apply

    Remove a guest from one site collection, sharing links included.

.EXAMPLE
    .\Revoke-SharePointUserAccess.ps1 -UserPrincipalName jan@contoso.com -TenantUrl "https://contoso.sharepoint.com" -IncludeGroupAccess

    Report SharePoint-level access and the Entra groups that also let Jan in — the offboarding
    checklist, since those groups have to be handled in Entra.

.EXAMPLE
    .\Revoke-SharePointUserAccess.ps1 -UserPrincipalName jan@contoso.com -TenantUrl "https://contoso.sharepoint.com" -RemoveFromEntraGroups -Apply -Confirm:$false

    The whole offboarding in one run: SharePoint access, and the Entra groups that were seen
    granting it. Read the report from a run without -Apply first — those memberships usually
    carry more than SharePoint.

.NOTES
    There is no checkpoint and no resume, unlike the permissions report. Revoking is idempotent —
    a second run finds only what the first did not remove — so re-running after an interruption
    is both the recovery and the verification, and it is safer than resuming a partly applied
    destructive operation from a saved position.

#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [string] $UserPrincipalName,
    [string] $TenantUrl,
    [string] $SiteUrl,
    [switch] $Apply,
    [string] $TenantId,
    [string] $ClientId,
    [string] $ClientSecret,
    [string] $CertificateThumbprint,
    [switch] $AppOnly,
    [string] $OutputPath,
    [ValidateSet('Site', 'List', 'Item')]
    [string] $Scope = 'Item',
    [switch] $IncludeOneDriveSites,
    [switch] $IncludeHiddenLists,
    [string] $FromReport,
    [switch] $IncludeGroupAccess,
    [switch] $KeepSharingLinks,
    [switch] $RemoveFromEntraGroups,
    [switch] $RemoveFromSite,
    [int] $GraphTimeoutSec = 120,
    [int] $MaxGraphRetry = 6
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($OutputPath) { $OutputPath }
             elseif ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' }
             else { "$HOME/Downloads" }

# Prove the output folder is usable before anything else happens: this run creates an app
# registration and deletes permissions, and finding out at the end that none of it could be
# recorded would leave a change with no audit trail.
try {
    if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -ErrorAction Stop | Out-Null }
    $writeProbe = Join-Path $outputDir ".sp-revoke-write-test-$PID.tmp"
    [System.IO.File]::WriteAllText($writeProbe, 'probe')
    Remove-Item -Path $writeProbe -Force -ErrorAction SilentlyContinue
} catch {
    Write-Host "  [ERROR] Output folder '$outputDir' is not writable: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Pass -OutputPath to write somewhere else." -ForegroundColor Yellow
    exit 1
}

$ts        = Get-Date -Format 'yyyyMMdd_HHmmss'
$safeUser  = ($UserPrincipalName -replace '[^\w.@-]', '_')
$actionCsv = Join-Path $outputDir "SharePoint_Revoke_${safeUser}_$ts.csv"

$TempAppNamePrefix = 'SP-RevokeAccess'

# ── Tenant from the report ────────────────────────────────────────────────────
# Authenticating needs a SharePoint host before anything is read, and -FromReport on its own
# does not give one. Rather than making the caller repeat a tenant URL the report already
# contains, take the first site out of it. Deliberately a cheap peek, not the full read: the
# real parse happens later, once the connection exists.
if ($FromReport -and -not $TenantUrl -and -not $SiteUrl) {
    try {
        $peekPath = $FromReport
        if ((Get-Item -LiteralPath $peekPath -ErrorAction Stop).PSIsContainer) {
            $peekPath = (Get-ChildItem -LiteralPath $FromReport -Filter 'SharePoint_Permissions_Detail_*.csv' |
                         Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
        }
        foreach ($row in (Import-Csv -LiteralPath $peekPath | Select-Object -First 50)) {
            $candidate = if ($row.SiteUrl) { [string]$row.SiteUrl } else { [string]$row.WebUrl }
            if ($candidate -match '^https?://[^/]+') {
                $TenantUrl = $Matches[0]
                break
            }
        }
        if ($TenantUrl) {
            Write-Host "  Tenant taken from the report: $TenantUrl" -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "  [ERROR] -FromReport could not be read to find the tenant: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Pass -TenantUrl as well, or point -FromReport at the detail CSV." -ForegroundColor Yellow
        exit 1
    }
    if (-not $TenantUrl) {
        Write-Host "  [ERROR] No site URL found in the report, so the tenant is unknown. Pass -TenantUrl." -ForegroundColor Red
        exit 1
    }
}

# ── Well-known application IDs ────────────────────────────────────────────────
# Defined before the shared block, not inside it: each script builds $RequiredAppRoles from
# these, and that happens before the block runs.
$GraphAppId      = '00000003-0000-0000-c000-000000000000'
$SharePointAppId = '00000003-0000-0ff1-ce00-000000000000'
$GraphResource   = 'https://graph.microsoft.com'

# What the temporary app is granted. User.Read.All is what the permissions report does not need
# and this script does: it has to resolve the named user before it can revoke anything, and
# GroupMember.Read.All does not allow reading an arbitrary user object. Without it every lookup
# comes back 403 and the run reports a real account as "not found in Entra ID".
$RequiredAppRoles = @(
    @{ ResourceAppId = $SharePointAppId; Role = 'Sites.FullControl.All'; Why = 'role assignments, site groups, sharing links' }
    @{ ResourceAppId = $GraphAppId;      Role = 'Sites.Read.All';        Why = 'tenant-wide site enumeration' }
    @{ ResourceAppId = $GraphAppId;      Role = 'User.Read.All';         Why = 'resolving the user to revoke' }
    @{ ResourceAppId = $GraphAppId;      Role = 'GroupMember.Read.All';  Why = 'the Entra groups that also grant access' }
)
# Write access to directory groups is asked for only when it will be used. A report-only run has
# no business holding a permission that can change group membership tenant-wide.
if ($RemoveFromEntraGroups) {
    $RequiredAppRoles += @{ ResourceAppId = $GraphAppId; Role = 'GroupMember.ReadWrite.All'; Why = 'removing the user from the groups that grant access' }
}

# Concurrency is fixed at 1: the shared block's parallel helper is for reading, and revocations
# are writes against the same site. Serial is slower and is the right trade here.
$Concurrency = 1

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host '   Revoke-SharePointUserAccess' -ForegroundColor Cyan
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host ("  User      : {0}" -f $UserPrincipalName) -ForegroundColor Cyan
Write-Host ("  Target    : {0}" -f $(if ($FromReport) { "webs named in $(Split-Path $FromReport -Leaf)" } elseif ($SiteUrl) { $SiteUrl } else { "$TenantUrl (tenant-wide)" })) -ForegroundColor Cyan
Write-Host ("  Scope     : {0}" -f $(switch ($Scope) {
    'Site' { 'Sites and sub-sites' }
    'List' { 'Sites, sub-sites, lists and libraries' }
    'Item' { 'Everything — sites, lists, folders and files with unique permissions' }
})) -ForegroundColor Cyan
Write-Host ("  Mode      : {0}" -f $(if ($Apply) { 'APPLY — access will be removed' } else { 'Report only — nothing is changed without -Apply' })) `
    -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ("  Links     : {0}" -f $(if ($KeepSharingLinks) { 'sharing links are left alone' } else { 'sharing links are revoked too' })) -ForegroundColor DarkGray
Write-Host ("  Entra     : {0}" -f $(if ($RemoveFromEntraGroups) { 'ALSO removing from the Entra groups seen granting access' } else { 'Entra group membership is reported, never changed' })) `
    -ForegroundColor $(if ($RemoveFromEntraGroups) { 'Yellow' } else { 'DarkGray' })
Write-Host ''

# ── SHARED BLOCK START ────────────────────────────────────────────────────────
# Everything from here to SHARED BLOCK END is kept byte-identical with the copy in
# scripts/SharePoint/Revoke-SharePointUserAccess.ps1. It is the app-only authentication and
# SharePoint REST layer, and it took four live runs against a tenant to get right: certificate
# credentials because SharePoint refuses secret-based app-only tokens, tokens that must prove
# they carry their app roles before being cached, 401 treated as fatal rather than per-site, and
# paging that cannot loop. A second, drifting copy of that is a correctness risk in a script that
# deletes permissions, so a test asserts the two are identical. Set $TempAppNamePrefix before it.
# ── Cleanup / shared state ────────────────────────────────────────────────────
$script:TempAppObjectId = $null
$script:ConnectedHere   = $false
$script:AppClientId     = $null
$script:AppClientSecret = $null
$script:AppCertificate  = $null
$script:AppSigningKey   = $null
$script:AppTenantId     = $null
$script:TokenCache      = @{}   # resource root URI -> @{ Headers; Expiry }
$script:ResourceRequiredRoles = @{}   # resource root URI -> app roles its token must carry
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
            # Returns the context it just disconnected; without Out-Null that object lands on
            # stdout and prints a stray ClientId/TenantId/Scopes table after the summary.
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
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

function Get-JwtClaim {
    # Reads one claim out of a JWT payload. No validation and no library: the token was just
    # handed to us by Entra over TLS, and all we want to know is what it says about itself.
    param(
        # Explicitly allows empty: a caller asking about a token it does not have should get back
        # "no claim", not a parameter binding error it then has to guard against separately.
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Token,
        [Parameter(Mandatory = $true)][string]$Claim
    )
    if ([string]::IsNullOrWhiteSpace($Token)) { return $null }
    try {
        $parts = $Token.Split('.')
        if ($parts.Count -lt 2) { return $null }
        $payload = $parts[1].Replace('-', '+').Replace('_', '/')
        switch ($payload.Length % 4) { 2 { $payload += '==' } 3 { $payload += '=' } 1 { return $null } }
        $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
        return $json.$Claim
    } catch { return $null }
}

function Get-ResourceToken {
    # Returns an Authorization header hashtable for the given resource root, minting and caching
    # a client-credentials token per resource. Refreshes 5 minutes before expiry.
    #
    # -RequiredRoles is what makes this safe to call straight after granting app roles. Entra will
    # happily issue a token before a freshly granted role has replicated, and that token carries no
    # roles claim at all — which Graph answers with 401, not 403. Cached for the hour it is valid,
    # one such token poisons the entire run and no amount of retrying the request can recover it,
    # because every retry is handed the same dead token back. So the token has to prove it carries
    # the roles before it is allowed into the cache.
    param(
        [Parameter(Mandatory = $true)][string]$Resource,
        [string[]]$RequiredRoles = @()
    )

    $cached = $script:TokenCache[$Resource]
    if ($cached -and (Get-Date) -lt $cached.Expiry) { return $cached.Headers }

    # Remember what this resource needs, so a refresh later in the run is held to the same bar.
    if ($RequiredRoles.Count -gt 0) { $script:ResourceRequiredRoles[$Resource] = $RequiredRoles }
    $expectedRoles = @($script:ResourceRequiredRoles[$Resource])

    $authority = "https://login.microsoftonline.com/$($script:AppTenantId)/oauth2/v2.0/token"

    # A freshly registered certificate credential and a freshly granted app role both need time to
    # replicate before Entra will mint a token against them — noticeably longer than the few
    # seconds a secret needs. Give it up to ~two minutes rather than failing the whole run on a
    # race that resolves itself.
    $resp        = $null
    $lastError   = $null
    $missingRoles = @()
    $maxAttempts = 15
    for ($i = 1; $i -le $maxAttempts; $i++) {
        # Rebuild the credential every attempt: the assertion carries its own expiry and a
        # single-use jti, and over a propagation wait this loop can outlive the first one.
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

        $resp = $null
        try {
            $resp = Invoke-RestMethod -Method POST -Uri $authority -Body $body -ErrorAction Stop
        } catch {
            $lastError = $_
            $resp = $null
        }

        if ($resp -and $resp.access_token) {
            if ($expectedRoles.Count -eq 0) { break }
            $granted     = @(Get-JwtClaim -Token $resp.access_token -Claim 'roles')
            $missingRoles = @($expectedRoles | Where-Object { $_ -notin $granted })
            if ($missingRoles.Count -eq 0) { break }
            # Token is technically valid but useless — discard it rather than cache it.
            $resp = $null
        }

        if ($i -lt $maxAttempts) {
            $reason = if ($missingRoles.Count -gt 0) {
                "app role(s) {0} not in the token yet" -f ($missingRoles -join ', ')
            } else { 'credential not accepted yet' }
            Write-ProgressHost -Message ("[INFO] Waiting for {0} to replicate for {1} (attempt {2}/{3})..." -f $reason, $Resource, $i, $maxAttempts)
            Start-Sleep -Seconds ([Math]::Min(5 * $i, 20))
        }
    }

    if (-not $resp -or -not $resp.access_token) {
        if ($missingRoles.Count -gt 0) {
            throw ("Entra issued a token for {0} without the required app role(s): {1}. The grant has not replicated, or it was not applied to this application." -f $Resource, ($missingRoles -join ', '))
        }
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
    $reauthTried = $false
    for ($attempt = 1; $attempt -le $MaxGraphRetry; $attempt++) {
        try {
            $headers = Get-ResourceToken -Resource $GraphResource
            return Invoke-RestMethod -Uri $Uri -Headers $headers -TimeoutSec $GraphTimeoutSec -ErrorAction Stop
        } catch {
            $statusCode = Get-ResponseStatusCode -ErrorRecord $_

            # Graph answers an app-only token that carries no usable role with 401, not 403, and a
            # cached token never improves on its own. Retrying the request alone is therefore
            # futile — the cache has to be dropped so the next attempt mints a fresh one.
            if ($statusCode -eq 401 -and -not $reauthTried) {
                $reauthTried = $true
                $script:TokenCache.Remove($GraphResource)
                Write-ProgressHost -Message "[INFO] Graph refused the token (401) — minting a fresh one and retrying..."
                Start-Sleep -Seconds 3
                continue
            }

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
            if ($statusCode -in @(403, 404)) {
                if (-not $ThrowOnDenied) { return $null }
                $what = if ($statusCode -eq 403) { 'Access denied' } else { 'Not found' }
                throw ("{0} (HTTP {1}) reading {2} — this scope could not be read, so its permissions are unknown rather than empty." -f $what, $statusCode, $Uri)
            }
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
        [int] $MaxPages = 200000,
        [switch] $ThrowOnDenied
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
        $resp = Invoke-SPGet -Uri $next -ThrowOnDenied:$ThrowOnDenied
        if (-not $resp) { break }
        $page++

        $values = $null
        if ($null -ne $resp.value) { $values = $resp.value }
        elseif ($resp.d -and $null -ne $resp.d.results) { $values = $resp.d.results }
        elseif ($page -eq 1) {
            # A collection endpoint always answers with value (or d.results), even when empty.
            # An object carrying neither is not an empty collection, it is the wrong URL - and
            # reading it as empty is a silent zero instead of a visible mistake.
            throw ("{0} did not return a collection - a scope object was asked for where its collection was meant." -f $next)
        }
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
        [int] $MaxPages = 200000,
        [switch] $ThrowOnDenied
    )
    $rows = [System.Collections.Generic.List[object]]::new()
    Invoke-SPCollectionPaged -Uri $Uri -MaxPages $MaxPages -ThrowOnDenied:$ThrowOnDenied -OnPage {
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

# ── Module preflight ──────────────────────────────────────────────────────────
$missingModules = @('Microsoft.Graph.Authentication') | Where-Object { -not (Get-Module -ListAvailable -Name $_) }
if ($missingModules.Count -gt 0) {
    Write-Host "  [ERROR] Missing required module(s): $($missingModules -join ', ')" -ForegroundColor Red
    Write-Host "  Install with: .\scripts\Startup\Install-Modules.ps1" -ForegroundColor Yellow
    exit 1
}

$allSitesMode = [string]::IsNullOrWhiteSpace($SiteUrl)
if ($allSitesMode -and -not $FromReport -and [string]::IsNullOrWhiteSpace($TenantUrl)) {
    Write-Host '  [ERROR] -TenantUrl is required when scanning all sites.' -ForegroundColor Red
    exit 1
}

# Unlike the storage/version reports, a single-site run needs the app-only path too: SharePoint
# role assignments are only readable with a Sites.FullControl.All token, and a delegated Graph
# token is the wrong audience for the /_api endpoints entirely.
# Sign-in goes through scripts\Startup\Connect-M365.ps1 where it can (PowerShell 7): the bootstrap
# sign-in below is delegated - device code and the GDAP customer per load.config.ps1 - and
# -AppOnly takes ClientId and CertificateThumbprint for the tenant from graph.appid.json. That app
# then needs the same roles as the temporary one, SharePoint Sites.FullControl.All included.
if ($PSVersionTable.PSVersion.Major -ge 7) { . (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1') }
if ($AppOnly -and -not $ClientId) {
    if (-not (Get-Command Get-M365AppRegistration -ErrorAction SilentlyContinue)) {
        Write-Host '  [ERROR] -AppOnly needs PowerShell 7. Pass -ClientId, -TenantId and -CertificateThumbprint instead.' -ForegroundColor Red
        exit 1
    }
    try {
        $appRegistration = Get-M365AppRegistration -TenantId $(if ($TenantId) { $TenantId } else { Resolve-M365TenantId })
    } catch {
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
    $ClientId              = $appRegistration.ClientId
    $CertificateThumbprint = $appRegistration.CertificateThumbprint
    if (-not $TenantId) { $TenantId = $appRegistration.Tenant }
}

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
        Write-Host "  Connecting (delegated)..." -ForegroundColor Cyan
        Write-Host "  Required role: Global Administrator or Application Administrator (to create the temporary lookup app)" -ForegroundColor DarkGray
        $bootstrapScopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
        if (Get-Command Connect-M365Graph -ErrorAction SilentlyContinue) {
            # Device code per load.config.ps1, the GDAP customer, and an existing session reused
            # when it already holds these scopes - which is then also left open at the end.
            $bootstrap = Connect-M365Graph -Scopes $bootstrapScopes -TenantId $effectiveTenantId
            $script:ConnectedHere = [bool]$bootstrap.ConnectedHere
        } else {
            # Windows PowerShell 5.1 cannot load Connect-M365.ps1: the same sign-in, inline.
            $connectParams = @{ Scopes = $bootstrapScopes; NoWelcome = $true }
            if ($effectiveTenantId) { $connectParams['TenantId'] = $effectiveTenantId }
            if ($global:useDeviceCodeAuth) { $connectParams['UseDeviceCode'] = $true }
            Connect-MgGraph @connectParams -ErrorAction Stop
            $script:ConnectedHere = $true
        }
        Write-Host "  [OK]   Connected (delegated)." -ForegroundColor DarkGray

        $ctx = Get-MgContext
        $usedTenantId = if ($effectiveTenantId) { $effectiveTenantId } else { $ctx.TenantId }
        if (-not $usedTenantId) {
            Write-Host "  [ERROR] Could not determine tenant ID. Provide -TenantId." -ForegroundColor Red
            Remove-TempApp; exit 1
        }

        # Named from a variable so this whole block stays byte-identical to the copy in
        # Revoke-SharePointUserAccess.ps1 — see the shared-block note above Remove-TempApp.
        $tempAppName = "$TempAppNamePrefix-Temp-$ts"
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
        # Set by each script before the shared block: the report and the revoke script need
        # different directory permissions, and granting a temporary Full Control app more
        # than it uses is not a detail worth being sloppy about.
        $requiredRoles = $RequiredAppRoles
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

        # Both tokens are minted here, and both must already carry their app roles. Waiting for
        # that now — while the run has produced nothing yet — is the difference between a clear
        # "the grant has not replicated" and a scan that walks the whole tenant on a dead token.
        Write-Host "  Obtaining app-only tokens (waiting for the grants to replicate)..." -ForegroundColor Cyan
        [void](Get-ResourceToken -Resource $GraphResource -RequiredRoles @($RequiredAppRoles | Where-Object { $_.ResourceAppId -eq $GraphAppId } | ForEach-Object { $_.Role }))
        Write-Host ("  [OK]   Graph token carries {0}." -f ((@($RequiredAppRoles | Where-Object { $_.ResourceAppId -eq $GraphAppId } | ForEach-Object { $_.Role })) -join ', ')) -ForegroundColor DarkGray

        $sharePointResource = Get-ResourceRootFromUrl -Url $(if ($SiteUrl) { $SiteUrl } else { $TenantUrl })
        [void](Get-ResourceToken -Resource $sharePointResource -RequiredRoles @('Sites.FullControl.All'))
        Write-Host "  [OK]   SharePoint token carries Sites.FullControl.All." -ForegroundColor DarkGray
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
# ── SHARED BLOCK END ──────────────────────────────────────────────────────────

# ── Writing to SharePoint ─────────────────────────────────────────────────────
$script:FormDigest = @{}   # web URL -> @{ Value; Expiry }

function Get-SPFormDigest {
    # SharePoint accepts an app-only bearer token on POST without a form digest in most cases,
    # but not all — and a refusal looks like a generic 403, which in a revocation script would
    # read as "not allowed to remove this" rather than "missing a header". Cheaper to always
    # send one. Digests last 30 minutes; this refreshes at 25.
    param([Parameter(Mandatory = $true)][string]$WebUrl)

    $cached = $script:FormDigest[$WebUrl]
    if ($cached -and (Get-Date) -lt $cached.Expiry) { return $cached.Value }

    $headers = Get-ResourceToken -Resource (Get-ResourceRootFromUrl -Url $WebUrl)
    $resp = Invoke-RestMethod -Method POST -Uri ("{0}/_api/contextinfo" -f $WebUrl) `
                -Headers $headers -TimeoutSec $GraphTimeoutSec -ErrorAction Stop
    $value = if ($resp.FormDigestValue) { $resp.FormDigestValue } else { $resp.d.GetContextWebInformation.FormDigestValue }
    if (-not $value) { throw "Could not obtain a form digest for $WebUrl." }

    $script:FormDigest[$WebUrl] = @{ Value = $value; Expiry = (Get-Date).AddMinutes(25) }
    return $value
}

function Invoke-GraphDelete {
    # The only call in this script that changes the directory rather than SharePoint. Mirrors
    # Invoke-GraphGet's retry and re-auth so a throttle does not read as a failed removal.
    param([Parameter(Mandatory = $true)][string]$Uri)
    $reauthTried = $false
    for ($attempt = 1; $attempt -le $MaxGraphRetry; $attempt++) {
        try {
            $headers = Get-ResourceToken -Resource $GraphResource
            return Invoke-RestMethod -Method DELETE -Uri $Uri -Headers $headers -TimeoutSec $GraphTimeoutSec -ErrorAction Stop
        } catch {
            $statusCode = Get-ResponseStatusCode -ErrorRecord $_
            if ($statusCode -eq 401 -and -not $reauthTried) {
                $reauthTried = $true
                $script:TokenCache.Remove($GraphResource)
                Start-Sleep -Seconds 2
                continue
            }
            $isRetryable = $statusCode -in @(408, 429, 500, 502, 503, 504)
            if (-not $isRetryable -or $attempt -eq $MaxGraphRetry) { throw }
            $delay = Get-RetryDelaySeconds -Attempt $attempt -ErrorRecord $_
            Write-ProgressHost -Message ("[WAIT] Graph throttled on a write — retry ({0}/{1}) in {2}s" -f $attempt, $MaxGraphRetry, $delay) -ForegroundColor Yellow
            Start-Sleep -Seconds $delay
        }
    }
}

function Invoke-SPPost {
    # The one call in this script that changes anything. Mirrors Invoke-SPGet's retry and 401
    # handling so a throttle does not read as a failed revocation.
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$WebUrl,
        [string]$Body,
        [hashtable]$ExtraHeaders = @{}
    )
    $resourceRoot = Get-ResourceRootFromUrl -Url $Uri
    $reauthTried  = $false

    for ($attempt = 1; $attempt -le $MaxGraphRetry; $attempt++) {
        try {
            $headers = @{} + (Get-ResourceToken -Resource $resourceRoot)
            $headers['X-RequestDigest'] = Get-SPFormDigest -WebUrl $WebUrl
            foreach ($key in $ExtraHeaders.Keys) { $headers[$key] = $ExtraHeaders[$key] }

            $params = @{
                Method      = 'POST'
                Uri         = $Uri
                Headers     = $headers
                TimeoutSec  = $GraphTimeoutSec
                ErrorAction = 'Stop'
            }
            if ($Body) {
                $params['Body']        = $Body
                $params['ContentType'] = 'application/json;odata=nometadata'
            }
            return Invoke-RestMethod @params
        } catch {
            $statusCode = Get-ResponseStatusCode -ErrorRecord $_
            if ($statusCode -eq 401 -and -not $reauthTried) {
                $reauthTried = $true
                $script:TokenCache.Remove($resourceRoot)
                $script:FormDigest.Remove($WebUrl)
                Start-Sleep -Seconds 2
                continue
            }
            $isRetryable = $statusCode -in @(408, 429, 500, 502, 503, 504)
            if (-not $isRetryable -or $attempt -eq $MaxGraphRetry) { throw }
            $delay = Get-RetryDelaySeconds -Attempt $attempt -ErrorRecord $_
            Write-ProgressHost -Message ("[WAIT] SharePoint throttled on a write — retry ({0}/{1}) in {2}s" -f $attempt, $MaxGraphRetry, $delay) -ForegroundColor Yellow
            Start-Sleep -Seconds $delay
        }
    }
}

# ── Finding the user ──────────────────────────────────────────────────────────
function ConvertFrom-GuestLoginName {
    # A guest's identity in SharePoint is not their address. It is stored as
    # jan_partner.com#ext#@contoso.onmicrosoft.com — the local part and the domain of their real
    # address joined by an underscore, with the host tenant appended. Turn that back into
    # jan@partner.com so it can be compared to what the caller typed.
    param([string]$Value)
    if (-not $Value) { return $null }
    $name = ($Value -split '\|')[-1]           # strip the i:0#.f|membership| claim prefix
    if ($name -notmatch '#ext#') { return $null }
    $external = ($name -split '#ext#')[0]
    $cut = $external.LastIndexOf('_')          # last underscore: a local part may contain one
    if ($cut -lt 1) { return $null }
    return ($external.Substring(0, $cut) + '@' + $external.Substring($cut + 1)).ToLowerInvariant()
}

function Test-UserIdentityMatch {
    # Exact comparisons only. A substring test here would be a way to revoke the wrong person —
    # "an@contoso.com" is a substring of "jan@contoso.com", and this script deletes permissions.
    param($SiteUser, [string]$Needle)

    $candidates = [System.Collections.Generic.List[string]]::new()
    foreach ($field in @($SiteUser.UserPrincipalName, $SiteUser.Email)) {
        if ($field) { $candidates.Add(([string]$field).ToLowerInvariant()) | Out-Null }
    }
    if ($SiteUser.LoginName) {
        $login = ([string]$SiteUser.LoginName).ToLowerInvariant()
        # The identity is whatever follows the last pipe of the claim.
        $candidates.Add(($login -split '\|')[-1]) | Out-Null
        $guest = ConvertFrom-GuestLoginName -Value $login
        if ($guest) { $candidates.Add($guest) | Out-Null }
    }
    return ($candidates -contains $Needle)
}

function Get-SiteUserEntry {
    # The user as this site collection knows them, or $null when they have never been given
    # anything here. Returning $null is the fast path: a site the user is unknown in has nothing
    # to revoke, and skipping it is what keeps a tenant-wide run to minutes instead of hours.
    param([Parameter(Mandatory = $true)][string]$WebUrl)

    $select = 'Id,Title,LoginName,Email,UserPrincipalName,IsSiteAdmin,PrincipalType'

    # The membership claim is how a cloud identity is written; try it directly before falling
    # back to reading the whole user list.
    $claim   = "i:0#.f|membership|$UserPrincipalName"
    $encoded = [uri]::EscapeDataString($claim)
    $direct  = Invoke-SPGet -Uri ("{0}/_api/web/siteusers/getByLoginName(@v)?@v='{1}'&`$select={2}" -f $WebUrl, $encoded, $select)
    if ($direct -and $direct.Id) { return $direct }

    # A guest is not addressable that way, so fall back to reading the site's users. Every match
    # is collected rather than the first one returned: two different accounts answering to the
    # same address means the caller has to say which, not that this script picks one and starts
    # deleting.
    $needle  = $UserPrincipalName.ToLowerInvariant()
    $matches = [System.Collections.Generic.List[object]]::new()
    foreach ($user in @(Get-SPCollection -Uri ("{0}/_api/web/siteusers?`$select={1}&`$top=500" -f $WebUrl, $select))) {
        if ([int]$user.PrincipalType -ne 1) { continue }
        if (Test-UserIdentityMatch -SiteUser $user -Needle $needle) { $matches.Add($user) | Out-Null }
    }

    $distinct = @($matches | Group-Object { [string]$_.LoginName })
    if ($distinct.Count -gt 1) {
        throw ("{0} matches {1} different accounts in this site ({2}) — name the exact one to revoke." -f
               $UserPrincipalName, $distinct.Count, (($distinct | ForEach-Object { $_.Name }) -join ' | '))
    }
    if ($matches.Count -gt 0) { return $matches[0] }
    return $null
}

function Get-UserSiteGroups {
    # The SharePoint groups this user belongs to in this site collection. Sharing-link groups
    # are in here too — a link shared with someone puts them in a SharingLinks.* group, which is
    # the only place that membership is recorded.
    param([Parameter(Mandatory = $true)][string]$WebUrl, [Parameter(Mandatory = $true)][int]$UserId)
    return @(Get-SPCollection -Uri ("{0}/_api/web/getUserById({1})/groups?`$select=Id,Title,LoginName,OwnerTitle" -f $WebUrl, $UserId))
}

# ── Recording what happens ────────────────────────────────────────────────────
# Only the Entra groups this run actually caught holding a role assignment on a scope in range.
# Not $userGroupIds, which is every group the user belongs to: a leaver can be in fifty groups,
# and -RemoveFromEntraGroups must touch exactly the ones that were seen granting SharePoint
# access and nothing else.
$script:GrantingEntraGroups = @{}
$script:ActionRows   = [System.Collections.Generic.List[object]]::new()
$script:ActionCsvPath = $actionCsv

function Add-ActionRow {
    param(
        [string]$SiteUrl, [string]$WebUrl, [string]$ScopeType, [string]$ScopeTitle, [string]$ScopeUrl,
        [string]$AccessVia, [string]$PermissionLevels, [string]$Action, [string]$Detail
    )
    $row = [PSCustomObject]@{
        UserPrincipalName = $UserPrincipalName
        SiteUrl           = $SiteUrl
        WebUrl            = $WebUrl
        ScopeType         = $ScopeType
        ScopeTitle        = $ScopeTitle
        ScopeUrl          = $ScopeUrl
        AccessVia         = $AccessVia
        PermissionLevels  = $PermissionLevels
        Action            = $Action
        Detail            = $Detail
        RunUtc            = (Get-Date).ToUniversalTime().ToString('s')
    }
    $script:ActionRows.Add($row) | Out-Null

    # Written as it happens, not at the end. A run that revokes two hundred things and then dies
    # would otherwise leave no record of what it removed, which is the one thing a destructive
    # script must never do. Append-CheckpointRows retries a locked file and throws if it cannot
    # write at all — losing the run is recoverable, losing the audit trail is not.
    Append-CheckpointRows -Path $script:ActionCsvPath -Rows @($row)
}

function Invoke-Revocation {
    # One revocation: decide, do it, record it. Every change in this script goes through here so
    # -Apply, -WhatIf, ShouldProcess and the audit CSV cannot disagree with each other.
    param(
        [Parameter(Mandatory = $true)][string]$Target,
        [Parameter(Mandatory = $true)][string]$Operation,
        [Parameter(Mandatory = $true)][scriptblock]$Do,
        [Parameter(Mandatory = $true)][hashtable]$Row
    )
    # -WhatIf is a dry run by any other name, so it records the same intent rather than looking
    # like someone declined a prompt.
    if (-not $Apply -or $WhatIfPreference) {
        Add-ActionRow @Row -Action 'WouldRevoke' -Detail $Operation
        Write-ProgressHost -Message ("    [DRY ] {0}: {1}" -f $Operation, $Target) -ForegroundColor DarkGray
        return $false
    }
    if (-not $PSCmdlet.ShouldProcess($Target, $Operation)) {
        Add-ActionRow @Row -Action 'Skipped' -Detail 'Declined at the confirmation prompt'
        return $false
    }
    try {
        & $Do
        Add-ActionRow @Row -Action 'Revoked' -Detail $Operation
        Write-ProgressHost -Message ("    [DONE] {0}: {1}" -f $Operation, $Target) -ForegroundColor Green
        return $true
    } catch {
        # Already gone is the outcome this script wants, not a failure. SharePoint answers a
        # removal of something that is no longer there with 404, and on a re-run after a partial
        # pass that is the normal case — counting it as failed would make a clean second run
        # look broken.
        $status = Get-ResponseStatusCode -ErrorRecord $_
        if ($status -eq 404 -or $_.Exception.Message -match 'does not exist|not found|cannot be found') {
            Add-ActionRow @Row -Action 'AlreadyGone' -Detail 'Nothing left to remove'
            Write-ProgressHost -Message ("    [ OK ] {0}: {1} — already gone" -f $Operation, $Target) -ForegroundColor DarkGray
            return $false
        }
        Add-ActionRow @Row -Action 'Failed' -Detail $_.Exception.Message
        Write-ProgressHost -Message ("    [FAIL] {0}: {1} — {2}" -f $Operation, $Target, $_.Exception.Message) -ForegroundColor Red
        return $false
    }
}

# ── Resolve the user in Entra ─────────────────────────────────────────────────
Write-ProgressHost -Message ("Resolving {0}..." -f $UserPrincipalName) -ForegroundColor Cyan
$entraUser = $null
try {
    $encodedUpn = [uri]::EscapeDataString($UserPrincipalName)
    $entraUser = Invoke-GraphGet -Uri ("https://graph.microsoft.com/v1.0/users/{0}?`$select=id,displayName,userPrincipalName,mail,userType,accountEnabled" -f $encodedUpn)
    $entraLookupDenied = $false
} catch {
    # Being refused the directory is not the same fact as the account not existing, and only one
    # of them is safe to shrug at. A 403 here means the app is missing User.Read.All, and reading
    # it as "not found" would hide a real account and silently skip every Entra group it belongs
    # to — which is the half of the report that says what this script cannot revoke.
    $entraLookupDenied = ((Get-ResponseStatusCode -ErrorRecord $_) -in @(401, 403))

    # A guest is often addressable by their real address rather than their tenant UPN.
    try {
        $filter = [uri]::EscapeDataString("mail eq '$UserPrincipalName' or userPrincipalName eq '$UserPrincipalName'")
        $found = Invoke-GraphGet -Uri ("https://graph.microsoft.com/v1.0/users?`$filter={0}&`$select=id,displayName,userPrincipalName,mail,userType,accountEnabled" -f $filter)
        $entraUser = @($found.value) | Select-Object -First 1
        if ($entraUser) { $entraLookupDenied = $false }
    } catch {
        if ((Get-ResponseStatusCode -ErrorRecord $_) -in @(401, 403)) { $entraLookupDenied = $true }
        $entraUser = $null
    }
}

if ($entraUser) {
    Write-Host ("  [OK]   {0} <{1}>{2}{3}" -f $entraUser.displayName, $entraUser.userPrincipalName,
        $(if ($entraUser.userType -eq 'Guest') { ' — guest' } else { '' }),
        $(if ($entraUser.accountEnabled -eq $false) { ' — account disabled' } else { '' })) -ForegroundColor DarkGray
} elseif ($entraLookupDenied) {
    Write-Host ''
    Write-Host "  [ERROR] Entra ID refused the lookup of $UserPrincipalName (403)." -ForegroundColor Red
    Write-Host "  The app is missing Graph User.Read.All, so this run could not tell whether the account" -ForegroundColor Red
    Write-Host "  exists — and without it the Entra groups that also grant access are never listed, which" -ForegroundColor Red
    Write-Host "  is exactly the part of the report saying what this script cannot revoke." -ForegroundColor Red
    Write-Host "  Using your own -ClientId? Grant it User.Read.All. Otherwise re-run and let the script" -ForegroundColor Yellow
    Write-Host "  create its own app, which now asks for it." -ForegroundColor Yellow
    Remove-TempApp; exit 1
} else {
    # Genuinely absent, and not fatal: a user deleted from Entra can still hold SharePoint
    # grants, and those are exactly the ones worth removing.
    Write-Host "  [WARN] No such account in Entra ID — continuing on the SharePoint side only." -ForegroundColor Yellow
    Write-Host "         A deleted account can still hold grants, which is a reason to run this, not to stop." -ForegroundColor Yellow
}

# The groups that will keep letting this user in after every SharePoint grant is gone.
$userGroupIds = @{}
if ($entraUser) {
    try {
        $uri = "https://graph.microsoft.com/v1.0/users/$($entraUser.id)/transitiveMemberOf/microsoft.graph.group?`$select=id,displayName&`$top=999"
        while ($uri) {
            $page = Invoke-GraphGet -Uri $uri
            foreach ($group in @($page.value)) { $userGroupIds[[string]$group.id] = [string]$group.displayName }
            $uri = $page.'@odata.nextLink'
        }
        Write-Host ("  [OK]   Member of {0} Entra group(s) — those grants cannot be revoked from SharePoint." -f $userGroupIds.Count) -ForegroundColor DarkGray
    } catch {
        Write-Host ("  [WARN] Could not read Entra group membership: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}

# ── Reading the permissions report ────────────────────────────────────────────
function Resolve-ReportFile {
    # -FromReport takes whatever is to hand: the detail CSV, the site-access CSV, the Excel
    # workbook, or just the folder they are in. Anything else would mean remembering which of
    # four filenames the report wrote.
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { throw "Report not found: $Path" }
    $item = Get-Item -LiteralPath $Path

    if ($item.PSIsContainer) {
        # Newest detail CSV in the folder — a folder usually holds several runs.
        $candidate = Get-ChildItem -LiteralPath $Path -Filter 'SharePoint_Permissions_Detail_*.csv' |
                     Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $candidate) { throw "No SharePoint_Permissions_Detail_*.csv in $Path — point -FromReport at the report folder or the detail CSV itself." }
        return $candidate.FullName
    }

    if ($item.Name -like 'SharePoint_Permissions_Detail_*.csv') { return $item.FullName }

    # Any other file from the same run: find its detail sibling by the shared timestamp.
    if ($item.Name -match '_(\d{8}_\d{6})\.(csv|xlsx)$') {
        $sibling = Join-Path $item.DirectoryName ("SharePoint_Permissions_Detail_{0}.csv" -f $Matches[1])
        if (Test-Path -LiteralPath $sibling) { return $sibling }
    }
    throw "Could not find the detail CSV belonging to $($item.Name). Point -FromReport at SharePoint_Permissions_Detail_<timestamp>.csv or at the folder."
}

function Get-ReportSitesForUser {
    # Which site collections the report says this user can reach. The report's own site-access
    # view already answers that per person — resolved through groups and sharing links — so it is
    # read in preference to the raw grant list, which only names the group and would force every
    # site holding any SharePoint group to be visited. On a tenant where most sites grant through
    # "Site Members", that is the difference between visiting five sites and visiting all of them.
    param([Parameter(Mandatory = $true)][string]$ReportPath)

    $detail = Resolve-ReportFile -Path $ReportPath
    Write-ProgressHost -Message ("Reading {0}..." -f (Split-Path $detail -Leaf)) -ForegroundColor Cyan

    # A report is a snapshot, and acting on a stale one silently misses everything granted since.
    $age = (Get-Date) - (Get-Item -LiteralPath $detail).LastWriteTime
    if ($age.TotalDays -ge 1) {
        Write-ProgressHost -Message ("  [WARN] This report is {0:N0} day(s) old. Access granted since then is not in it." -f $age.TotalDays) -ForegroundColor Yellow
    }

    $needle = $UserPrincipalName.ToLowerInvariant()
    $sites  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    function Test-ReportIdentity([string]$Login, [string]$Mail, [string]$Upn) {
        foreach ($value in @($Upn, $Mail)) {
            if ($value -and $value.ToLowerInvariant() -eq $needle) { return $true }
        }
        if ($Login) {
            $low = $Login.ToLowerInvariant()
            if (($low -split '\|')[-1] -eq $needle) { return $true }
            if ((ConvertFrom-GuestLoginName -Value $low) -eq $needle) { return $true }
        }
        return $false
    }

    # Preferred source: one row per person per site, groups already resolved to people.
    $siteAccess = $detail -replace '_Detail_', '_SiteAccess_'
    if (Test-Path -LiteralPath $siteAccess) {
        $rows = 0
        Import-Csv -LiteralPath $siteAccess | ForEach-Object {
            $rows++
            if (Test-ReportIdentity -Login $null -Mail ([string]$_.UserEmail) -Upn ([string]$_.UserPrincipalName)) {
                if ($_.SiteUrl) { [void]$sites.Add([string]$_.SiteUrl) }
            }
        }
        Write-ProgressHost -Message ("  {0:N0} access row(s) read; {1} site(s) name this user" -f $rows, $sites.Count) -ForegroundColor DarkGray
    } else {
        # Older report, or one written before the site-access view existed. Fall back to the raw
        # grants: direct ones are exact, and a SharePoint group has to be taken on trust because
        # this file does not say who is in it.
        Write-ProgressHost -Message "  No site-access file beside the report — falling back to the grant list." -ForegroundColor Yellow
        $direct = 0; $viaGroup = 0
        Import-Csv -LiteralPath $detail | ForEach-Object {
            if ($_.ItemType -eq 'Error') { return }
            $site = if ($_.SiteUrl) { [string]$_.SiteUrl } else { [string]$_.WebUrl }
            if (-not $site) { return }
            if (Test-ReportIdentity -Login ([string]$_.PrincipalLogin) -Mail ([string]$_.PrincipalEmail) -Upn $null) {
                [void]$sites.Add($site); $direct++; return
            }
            if ($_.DirectoryObjectId -and $userGroupIds.ContainsKey([string]$_.DirectoryObjectId)) {
                [void]$sites.Add($site); $viaGroup++; return
            }
            if ($_.PrincipalType -in @('SharePointGroup', 'SharingLink')) { [void]$sites.Add($site); $viaGroup++ }
        }
        Write-ProgressHost -Message ("  {0} direct grant(s), {1} through a group or link" -f $direct, $viaGroup) -ForegroundColor DarkGray
    }

    # The report's own gaps are this run's gaps, so they are counted and reported rather than
    # inherited quietly.
    $errorRows = @(Import-Csv -LiteralPath $detail | Where-Object { $_.ItemType -eq 'Error' }).Count
    if ($errorRows -gt 0) {
        Write-ProgressHost -Message ("  [WARN] The report itself recorded {0} scope(s) it could not read — those are blind spots here too." -f $errorRows) -ForegroundColor Yellow
    }
    return @($sites)
}

# ── Site discovery ────────────────────────────────────────────────────────────
# Same three sources as the permissions report, for the same reason: no single one is complete,
# and a sub-web Graph does not know about is exactly where a forgotten grant hides.
Write-ProgressHost -Message "Retrieving sites..." -ForegroundColor Cyan

$targetWebs   = [System.Collections.Generic.List[object]]::new()
$knownWebUrl  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

function Add-TargetWeb {
    param([string]$WebUrl, [string]$Title, [string]$GraphId)
    if ([string]::IsNullOrWhiteSpace($WebUrl)) { return $false }
    $normalized = $WebUrl.TrimEnd('/')
    if (-not $knownWebUrl.Add($normalized)) { return $false }
    $targetWebs.Add([PSCustomObject]@{ WebUrl = $normalized; Title = $Title; GraphId = $GraphId }) | Out-Null
    return $true
}

if ($FromReport) {
    # The permissions report already walked the tenant and wrote down where this user can reach.
    # Reusing that instead of walking it again is the difference between minutes and seconds, and
    # it means the revocation acts on exactly the webs you read in the report rather than on a
    # second, slightly different scan.
    $reportSites = Get-ReportSitesForUser -ReportPath $FromReport
    foreach ($site in $reportSites) { [void](Add-TargetWeb -WebUrl $site -Title $null -GraphId $null) }

    Write-ProgressHost -Message ("From the report: {0} site collection(s) where {1} holds access" -f $targetWebs.Count, $UserPrincipalName) -ForegroundColor Green
    if ($targetWebs.Count -eq 0) {
        Write-Host ''
        Write-Host ("  [OK]   The report lists no access for {0}. Nothing to revoke." -f $UserPrincipalName) -ForegroundColor Green
        Write-Host "         If that is a surprise, check the report covered the ground you expected." -ForegroundColor DarkGray
        Remove-TempApp; exit 0
    }
    # A report is a snapshot. Anything granted after it was written is invisible here, so the
    # webs are still scanned live — the report decides where to look, never what to remove.
    Write-ProgressHost -Message "  Each web is still read live; the report only decides which ones to visit." -ForegroundColor DarkGray
} elseif (-not $allSitesMode) {
    try {
        $uri  = [System.Uri]$SiteUrl.TrimEnd('/')
        $obj  = Invoke-GraphGet -Uri ("https://graph.microsoft.com/v1.0/sites/{0}:{1}?`$select=id,displayName,webUrl" -f $uri.Host, $uri.AbsolutePath.TrimEnd('/'))
        if (-not $obj -or -not $obj.id) {
            Write-Host "  [ERROR] Site not found: $SiteUrl" -ForegroundColor Red
            Remove-TempApp; exit 1
        }
        [void](Add-TargetWeb -WebUrl $obj.webUrl -Title $obj.displayName -GraphId $obj.id)
    } catch {
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        Remove-TempApp; exit 1
    }
} else {
    $next = 'https://graph.microsoft.com/v1.0/sites/getAllSites?$select=id,displayName,webUrl&$top=200'
    while ($next) {
        $page = Invoke-GraphGet -Uri $next
        foreach ($s in @($page.value)) { [void](Add-TargetWeb -WebUrl $s.webUrl -Title $s.displayName -GraphId $s.id) }
        $next = $page.'@odata.nextLink'
    }
    if (-not $IncludeOneDriveSites) {
        $filtered = @($targetWebs | Where-Object { $_.WebUrl -notmatch '-my\.sharepoint\.com/personal/' })
        $targetWebs = [System.Collections.Generic.List[object]]::new($filtered)
        $knownWebUrl.Clear()
        foreach ($w in $targetWebs) { [void]$knownWebUrl.Add($w.WebUrl) }
    }
}

$queue = [System.Collections.Generic.Queue[object]]::new()
foreach ($w in $targetWebs) { $queue.Enqueue($w) }
while ($queue.Count -gt 0) {
    $parent = $queue.Dequeue()
    if ($parent.GraphId) {
        try {
            $subUri = "https://graph.microsoft.com/v1.0/sites/$($parent.GraphId)/sites`?`$select=id,displayName,webUrl&`$top=200"
            do {
                $resp = Invoke-GraphGet -Uri $subUri
                foreach ($s in @($resp.value)) {
                    if (Add-TargetWeb -WebUrl $s.webUrl -Title $s.displayName -GraphId $s.id) { $queue.Enqueue($targetWebs[$targetWebs.Count - 1]) }
                }
                $subUri = $resp.'@odata.nextLink'
            } while ($subUri)
        } catch { }
    }
    try {
        foreach ($web in @(Get-SPCollection -Uri ("{0}/_api/web/webs?`$select=Title,Url" -f $parent.WebUrl))) {
            if (Add-TargetWeb -WebUrl $web.Url -Title $web.Title -GraphId $null) { $queue.Enqueue($targetWebs[$targetWebs.Count - 1]) }
        }
    } catch { }
}

Write-ProgressHost -Message ("Target webs: {0}" -f $targetWebs.Count) -ForegroundColor Green
if ($targetWebs.Count -eq 0) {
    Write-Host '  [ERROR] No sites to search.' -ForegroundColor Red
    Remove-TempApp; exit 1
}

function Get-SiteCollectionUrl {
    # SharePoint users and groups live on the site collection, not the web, so every sub-web of
    # one site shares a single user entry and a single set of groups.
    param([Parameter(Mandatory = $true)][string]$WebUrl)
    $uri  = [System.Uri]$WebUrl
    $path = $uri.AbsolutePath.TrimEnd('/')
    if ($path -match '^(/(?:sites|teams|personal)/[^/]+)') { return ("{0}://{1}{2}" -f $uri.Scheme, $uri.Host, $Matches[1]) }
    return ("{0}://{1}" -f $uri.Scheme, $uri.Host)
}

function Get-ScopeRoleAssignments {
    # -ThrowOnDenied for the same reason as in the report: an empty result from a refused read
    # would read as "this user has nothing here", and acting on that is how a revocation quietly
    # misses a grant.
    # Takes the scope base (.../_api/web, .../lists(guid'..'), .../items(n)) and appends
    # /roleassignments itself, so it matches what Invoke-ScopeRevocation builds its removal URL
    # from. Leaving that to the callers is how this came to read the web object instead of its
    # role assignments: SharePoint answered with the web, there was no value array to find, and
    # a live run reported access on fifteen sites with nothing to revoke and no error at all.
    param([Parameter(Mandatory = $true)][string]$Uri)
    $query = "?`$expand=Member,RoleDefinitionBindings&`$select=PrincipalId,Member/Id,Member/Title,Member/LoginName,Member/PrincipalType,RoleDefinitionBindings/Name"
    return Get-SPCollection -ThrowOnDenied -Uri ($Uri.TrimEnd('/') + '/roleassignments' + $query)
}

# ── Scan and revoke, one site collection at a time ────────────────────────────
$siteCollections = [ordered]@{}
foreach ($web in $targetWebs) {
    $sc = Get-SiteCollectionUrl -WebUrl $web.WebUrl
    if (-not $siteCollections.Contains($sc)) { $siteCollections[$sc] = [System.Collections.Generic.List[object]]::new() }
    $siteCollections[$sc].Add($web) | Out-Null
}

$stats = [PSCustomObject]@{ SitesSearched = 0; SitesWithAccess = 0; Revoked = 0; Failed = 0; Found = 0; GroupOnly = 0; AdminLeft = 0; EntraRemoved = 0 }
# Declared here, filled just before the Entra phase: the summary reads it after the scan's
# finally, and a run that dies early must leave it an empty list rather than undefined.
$scanLimits = @()
$scIndex = 0

try {
    foreach ($siteCollectionUrl in $siteCollections.Keys) {
        $scIndex++
        $webs = $siteCollections[$siteCollectionUrl]
        Set-ScanProgress -Id 1 -Activity 'Sites doorzoeken' -Status ("[{0}/{1}] {2}" -f $scIndex, $siteCollections.Count, $siteCollectionUrl) `
            -PercentComplete ([int](($scIndex / [Math]::Max($siteCollections.Count, 1)) * 100))

        $stats.SitesSearched++
        $rootWeb = ($webs | Sort-Object { $_.WebUrl.Length } | Select-Object -First 1).WebUrl

        try {
            $siteUser = Get-SiteUserEntry -WebUrl $rootWeb
        } catch {
            Write-ProgressHost -Message ("[WARN] {0}: could not look the user up — {1}" -f $siteCollectionUrl, $_.Exception.Message) -ForegroundColor Yellow
            Add-ActionRow -SiteUrl $siteCollectionUrl -WebUrl $rootWeb -ScopeType 'Site' -ScopeTitle $siteCollectionUrl -ScopeUrl $siteCollectionUrl `
                -AccessVia 'n/a' -PermissionLevels $null -Action 'Failed' -Detail "User lookup failed: $($_.Exception.Message)"
            $stats.Failed++
            continue
        }

        # Entra groups that grant access here are reported whether or not the user is known to
        # this site — they are the routes this script cannot close.
        $groupGrants = [System.Collections.Generic.List[object]]::new()

        if (-not $siteUser) {
            if (-not $IncludeGroupAccess) { continue }
        } else {
            $stats.SitesWithAccess++
            Write-ProgressHost -Message ("[{0}/{1}] {2}" -f $scIndex, $siteCollections.Count, $siteCollectionUrl) -ForegroundColor White
        }

        $userId = if ($siteUser) { [int]$siteUser.Id } else { -1 }

        # -- Site collection administrator -----------------------------------
        # First, because it overrides every role assignment below: leaving it in place would
        # make every other removal cosmetic.
        if ($siteUser -and [bool]$siteUser.IsSiteAdmin) {
            $stats.Found++
            $row = @{ SiteUrl = $siteCollectionUrl; WebUrl = $rootWeb; ScopeType = 'SiteCollection'
                      ScopeTitle = $siteCollectionUrl; ScopeUrl = $siteCollectionUrl
                      AccessVia = 'Site collection administrator'; PermissionLevels = 'Full control (administrator)' }
            $done = Invoke-Revocation -Target $siteCollectionUrl -Operation 'Remove site collection administrator' -Row $row -Do {
                Invoke-SPPost -WebUrl $rootWeb -Uri ("{0}/_api/web/getUserById({1})" -f $rootWeb, $userId) `
                    -Body '{"IsSiteAdmin": false}' -ExtraHeaders @{ 'X-HTTP-Method' = 'MERGE'; 'IF-MATCH' = '*' } | Out-Null
            }
            if ($done) {
                $stats.Revoked++
            } elseif ($Apply -and -not $WhatIfPreference) {
                # Everything below this point is cosmetic while the flag is still set: a site
                # collection administrator reaches every scope in the site regardless of role
                # assignments. Saying so here beats a summary that reads like a success.
                $stats.Failed++
                $stats.AdminLeft++
                Write-ProgressHost -Message ("  [FAIL] Still a site collection administrator on {0} — every other removal here is cosmetic until that is fixed." -f $siteCollectionUrl) -ForegroundColor Red
            }
        }

        # -- SharePoint groups, sharing links included -------------------------
        $sharingLinkGroupIds = @{}
        if ($siteUser) {
            $memberships = @()
            try { $memberships = Get-UserSiteGroups -WebUrl $rootWeb -UserId $userId } catch {
                Write-ProgressHost -Message ("  [WARN] Could not read group membership: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
            }
            foreach ($group in $memberships) {
                $isLink = ([string]$group.LoginName) -like 'SharingLinks.*'
                if ($isLink) { $sharingLinkGroupIds[[string]$group.Id] = $true }
                if ($isLink -and $KeepSharingLinks) {
                    Add-ActionRow -SiteUrl $siteCollectionUrl -WebUrl $rootWeb -ScopeType 'SharingLink' -ScopeTitle $group.Title -ScopeUrl $rootWeb `
                        -AccessVia "Sharing link: $($group.Title)" -PermissionLevels $null -Action 'Kept' -Detail '-KeepSharingLinks was given'
                    continue
                }
                $stats.Found++
                $kind = if ($isLink) { 'SharingLink' } else { 'SharePointGroup' }
                $via  = if ($isLink) { "Sharing link group $($group.Title)" } else { "SharePoint group $($group.Title)" }
                $row  = @{ SiteUrl = $siteCollectionUrl; WebUrl = $rootWeb; ScopeType = $kind
                           ScopeTitle = [string]$group.Title; ScopeUrl = $rootWeb
                           AccessVia = $via; PermissionLevels = $null }
                $gid  = [int]$group.Id
                $done = Invoke-Revocation -Target "$($group.Title) @ $siteCollectionUrl" -Operation "Remove from $kind" -Row $row -Do {
                    Invoke-SPPost -WebUrl $rootWeb -Uri ("{0}/_api/web/sitegroups({1})/users/removeById({2})" -f $rootWeb, $gid, $userId) | Out-Null
                }
                if ($done) { $stats.Revoked++ } elseif ($Apply) { $stats.Failed++ }
            }
        }

        # -- Direct role assignments, scope by scope ---------------------------
        # A grant straight to the person, rather than through a group. Only scopes with unique
        # permissions can carry one; everything else inherits and is handled by its parent.
        foreach ($web in $webs) {
            $webUrl = $web.WebUrl

            function Resolve-UserAssignment {
                # Does this scope grant anything to the user, to a sharing-link group they are
                # in, or to an Entra group they belong to? The last is reported, never removed.
                param([object[]]$Assignments)
                $hits = [System.Collections.Generic.List[object]]::new()
                foreach ($ra in $Assignments) {
                    if (-not $ra.Member) { continue }
                    $levels = @(@($ra.RoleDefinitionBindings) | ForEach-Object { [string]$_.Name } | Where-Object { $_ })
                    if ($levels.Count -eq 0) { continue }
                    $principal = Get-PrincipalInfo -Member $ra.Member
                    # Not $pid: that is a read-only automatic variable holding the process id, and
                    # assigning to it throws. Thrown here it killed the evaluation of every scope,
                    # which is why a live run reported access on 15 sites and nothing to revoke.
                    $principalId = [int]$ra.PrincipalId

                    if ($siteUser -and $principalId -eq $userId) {
                        $hits.Add([PSCustomObject]@{ Kind = 'Direct'; PrincipalId = $principalId; Via = 'Granted directly to the user'; Levels = ($levels -join '; ') }) | Out-Null
                    } elseif ($sharingLinkGroupIds.ContainsKey([string]$principalId)) {
                        $hits.Add([PSCustomObject]@{ Kind = 'SharingLink'; PrincipalId = $principalId; Via = "Sharing link $($principal.Title)"; Levels = ($levels -join '; ') }) | Out-Null
                    } elseif ($principal.DirectoryId -and $userGroupIds.ContainsKey([string]$principal.DirectoryId)) {
                        $hits.Add([PSCustomObject]@{ Kind = 'EntraGroup'; PrincipalId = $principalId; Via = "Entra group $($principal.Title)"; Levels = ($levels -join '; ')
                                                     DirectoryId = $principal.DirectoryId; GroupName = $principal.Title }) | Out-Null
                    } elseif ($principal.Kind -in @('Everyone', 'EveryoneExceptExternalUsers', 'AllAuthenticatedUsers')) {
                        $hits.Add([PSCustomObject]@{ Kind = 'Everyone'; PrincipalId = $principalId; Via = $principal.Title; Levels = ($levels -join '; ') }) | Out-Null
                    }
                }
                return $hits
            }

            function Invoke-ScopeRevocation {
                # Removes the user's own assignment on one scope, and reports the routes that are
                # not this script's to close.
                param([string]$ScopeType, [string]$ScopeTitle, [string]$ScopeUrl, [string]$RoleAssignmentUri, [object[]]$Hits)
                foreach ($hit in $Hits) {
                    $script:stats.Found++
                    $row = @{ SiteUrl = $siteCollectionUrl; WebUrl = $webUrl; ScopeType = $ScopeType
                              ScopeTitle = $ScopeTitle; ScopeUrl = $ScopeUrl
                              AccessVia = $hit.Via; PermissionLevels = $hit.Levels }

                    if ($hit.Kind -eq 'EntraGroup') {
                        # Remembered so -RemoveFromEntraGroups can act on exactly these groups
                        # afterwards, and on no others: the user may be in fifty groups, and only
                        # the ones that actually grant SharePoint access are in scope here.
                        if ($hit.DirectoryId) { $script:GrantingEntraGroups[[string]$hit.DirectoryId] = $hit.GroupName }
                        $detail = if ($RemoveFromEntraGroups) {
                            'Granted through an Entra ID group — handled in the Entra phase below'
                        } else {
                            'Granted through an Entra ID group — remove the user from that group in Entra, or re-run with -RemoveFromEntraGroups'
                        }
                        Add-ActionRow @row -Action 'CannotRevoke' -Detail $detail
                        $script:stats.GroupOnly++
                        continue
                    }
                    if ($hit.Kind -eq 'Everyone') {
                        Add-ActionRow @row -Action 'CannotRevoke' -Detail 'Granted to everyone — removing it would revoke access for the whole tenant, not this user'
                        continue
                    }
                    # A sharing-link group's assignment is left in place: the user was already
                    # taken out of the group above, and the link may still serve other people.
                    if ($hit.Kind -eq 'SharingLink') {
                        Add-ActionRow @row -Action $(if ($KeepSharingLinks) { 'Kept' } else { 'Revoked' }) `
                            -Detail $(if ($KeepSharingLinks) { '-KeepSharingLinks was given' } else { 'Removed by taking the user out of the sharing link group; the link itself still exists for others' })
                        continue
                    }

                    $principalId = $hit.PrincipalId
                    $done = Invoke-Revocation -Target "$ScopeTitle ($ScopeUrl)" -Operation "Remove direct $($hit.Levels)" -Row $row -Do {
                        Invoke-SPPost -WebUrl $webUrl -Uri ("{0}/roleassignments/removeroleassignment(principalid={1})" -f $RoleAssignmentUri, $principalId) | Out-Null
                    }
                    if ($done) { $script:stats.Revoked++ } elseif ($Apply) { $script:stats.Failed++ }
                }
            }

            # Web level. Always checked: a sub-web that inherits is covered by its parent, but a
            # sub-web with its own permissions is a scope of its own.
            try {
                $webInfo = Invoke-SPGet -Uri ("{0}/_api/web?`$select=Title,HasUniqueRoleAssignments" -f $webUrl)
                if ($webInfo -and [bool]$webInfo.HasUniqueRoleAssignments) {
                    $hits = Resolve-UserAssignment -Assignments @(Get-ScopeRoleAssignments -Uri ("{0}/_api/web" -f $webUrl))
                    Invoke-ScopeRevocation -ScopeType 'Web' -ScopeTitle ([string]$webInfo.Title) -ScopeUrl $webUrl `
                        -RoleAssignmentUri ("{0}/_api/web" -f $webUrl) -Hits $hits
                }
            } catch {
                Write-ProgressHost -Message ("  [WARN] {0}: {1}" -f $webUrl, $_.Exception.Message) -ForegroundColor Yellow
                Add-ActionRow -SiteUrl $siteCollectionUrl -WebUrl $webUrl -ScopeType 'Web' -ScopeTitle $webUrl -ScopeUrl $webUrl `
                    -AccessVia 'n/a' -PermissionLevels $null -Action 'Failed' -Detail $_.Exception.Message
                $stats.Failed++
                continue
            }

            if ($Scope -eq 'Site') { continue }

            $lists = @()
            try {
                $lists = @(Get-SPCollection -Uri ("{0}/_api/web/lists?`$select=Id,Title,Hidden,BaseTemplate,ItemCount,HasUniqueRoleAssignments" -f $webUrl))
            } catch {
                Write-ProgressHost -Message ("  [WARN] Could not list {0}: {1}" -f $webUrl, $_.Exception.Message) -ForegroundColor Yellow
            }
            if (-not $IncludeHiddenLists) { $lists = @($lists | Where-Object { -not $_.Hidden }) }

            foreach ($list in $lists) {
                $listId = [string]$list.Id
                try {
                    if ([bool]$list.HasUniqueRoleAssignments) {
                        $hits = Resolve-UserAssignment -Assignments @(Get-ScopeRoleAssignments -Uri ("{0}/_api/web/lists(guid'{1}')" -f $webUrl, $listId))
                        Invoke-ScopeRevocation -ScopeType 'List' -ScopeTitle ([string]$list.Title) -ScopeUrl $webUrl `
                            -RoleAssignmentUri ("{0}/_api/web/lists(guid'{1}')" -f $webUrl, $listId) -Hits $hits
                    }

                    # Template 112 is the User Information List: SharePoint refuses to enumerate
                    # its items at any field width, and they are directory records rather than
                    # content. Same skip as the permissions report.
                    if ($Scope -ne 'Item' -or [int]$list.ItemCount -eq 0 -or [int]$list.BaseTemplate -eq 112) { continue }

                    $unique = [System.Collections.Generic.List[object]]::new()
                    Invoke-SPCollectionPaged -Uri ("{0}/_api/web/lists(guid'{1}')/items?`$select=Id,FileRef,FileLeafRef,HasUniqueRoleAssignments&`$top=2000" -f $webUrl, $listId) -OnPage {
                        param($PageItems)
                        foreach ($it in $PageItems) { if ([bool]$it.HasUniqueRoleAssignments) { $unique.Add($it) | Out-Null } }
                    }
                    foreach ($item in $unique) {
                        $itemId = [int]$item.Id
                        $hits = Resolve-UserAssignment -Assignments @(Get-ScopeRoleAssignments -Uri ("{0}/_api/web/lists(guid'{1}')/items({2})" -f $webUrl, $listId, $itemId))
                        Invoke-ScopeRevocation -ScopeType 'Item' -ScopeTitle ([string]$item.FileLeafRef) -ScopeUrl ([string]$item.FileRef) `
                            -RoleAssignmentUri ("{0}/_api/web/lists(guid'{1}')/items({2})" -f $webUrl, $listId, $itemId) -Hits $hits
                    }
                } catch {
                    if ($_.Exception.Message -match 'SharePoint refused the token \(401\)') { throw }
                    Write-ProgressHost -Message ("  [WARN] {0} > {1}: {2}" -f $webUrl, $list.Title, $_.Exception.Message) -ForegroundColor Yellow
                    Add-ActionRow -SiteUrl $siteCollectionUrl -WebUrl $webUrl -ScopeType 'List' -ScopeTitle ([string]$list.Title) -ScopeUrl $webUrl `
                        -AccessVia 'n/a' -PermissionLevels $null -Action 'Failed' -Detail $_.Exception.Message
                    $stats.Failed++
                }
            }
        }

        # -- Optionally drop the user from the site collection entirely --------
        if ($RemoveFromSite -and $siteUser) {
            # Counted like any other finding, or a dry run whose only action is this one reports
            # "Grants found: 0" while the CSV says something would be removed.
            $stats.Found++
            $row = @{ SiteUrl = $siteCollectionUrl; WebUrl = $rootWeb; ScopeType = 'SiteCollection'
                      ScopeTitle = $siteCollectionUrl; ScopeUrl = $siteCollectionUrl
                      AccessVia = 'Site user list'; PermissionLevels = $null }
            $done = Invoke-Revocation -Target $siteCollectionUrl -Operation 'Remove the user from the site collection' -Row $row -Do {
                Invoke-SPPost -WebUrl $rootWeb -Uri ("{0}/_api/web/siteusers/removeById({1})" -f $rootWeb, $userId) | Out-Null
            }
            if ($done) { $stats.Revoked++ } elseif ($Apply) { $stats.Failed++ }
        }
    }
    # ── Entra ID groups that grant access ─────────────────────────────────
    # Deliberately last, and deliberately narrow. Only the groups this run actually saw granting
    # SharePoint access are touched — never every group the user belongs to. An Entra group is
    # also not a SharePoint object: it can carry Teams, mailboxes, licences and app assignments,
    # so removing someone from one reaches much further than this report can see.
    # What this phase can see is bounded by what the scan found, so say where those bounds are
    # before acting on the result. "Removed every group that grants access" is only true of the
    # ground the scan actually covered, and treating a narrowed or partly failed run as complete
    # is how someone concludes an offboarding is finished when it is not.
    # With -FromReport the ground covered is the report's, not this run's: a web the report never
    # visited is a web this run never saw, whatever flags were passed here.
    if ($FromReport)       { $scanLimits += "only the webs named in $(Split-Path $FromReport -Leaf) were visited, so this inherits whatever that report did not cover" }
    if ($SiteUrl)          { $scanLimits += "only $SiteUrl was searched, not the tenant" }
    if ($Scope -eq 'Site') { $scanLimits += 'only site level was searched, so grants on lists, folders and files were never looked at' }
    if ($Scope -eq 'List') { $scanLimits += 'folders and files were not searched' }
    if (-not $IncludeOneDriveSites -and -not $SiteUrl) { $scanLimits += 'OneDrive sites were excluded' }
    if (-not $IncludeHiddenLists)  { $scanLimits += 'hidden and system lists were skipped' }
    if ($stats.Failed -gt 0)       { $scanLimits += "$($stats.Failed) scope(s) could not be read" }

    if ($RemoveFromEntraGroups -and $script:GrantingEntraGroups.Count -gt 0) {
        Write-Out ''
        Write-ProgressHost -Message ("Entra ID groups that grant access: {0}" -f $script:GrantingEntraGroups.Count) -ForegroundColor Cyan
        Write-ProgressHost -Message "  These grant more than SharePoint — Teams, mailboxes and licences ride on the same membership." -ForegroundColor Yellow
        if ($scanLimits.Count -gt 0) {
            Write-ProgressHost -Message "  [WARN] This list is only as complete as the scan behind it:" -ForegroundColor Yellow
            foreach ($limit in $scanLimits) { Write-ProgressHost -Message ("           - {0}" -f $limit) -ForegroundColor Yellow }
            Write-ProgressHost -Message "         A group granting access somewhere that was not searched is not in this list." -ForegroundColor Yellow
        }

        if (-not $entraUser) {
            Write-ProgressHost -Message "  [SKIP] The user could not be resolved in Entra, so membership cannot be changed." -ForegroundColor Yellow
        }

        foreach ($groupId in $script:GrantingEntraGroups.Keys) {
            $groupName = $script:GrantingEntraGroups[$groupId]
            $row = @{ SiteUrl = 'Entra ID'; WebUrl = $null; ScopeType = 'EntraGroup'
                      ScopeTitle = $groupName; ScopeUrl = "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Members/groupId/$groupId"
                      AccessVia = "Entra group $groupName"; PermissionLevels = $null }

            if (-not $entraUser) {
                Add-ActionRow @row -Action 'CannotRevoke' -Detail 'The user was not resolved in Entra ID'
                continue
            }

            $group = $null
            try {
                $group = Invoke-GraphGet -Uri ("https://graph.microsoft.com/v1.0/groups/{0}?`$select=id,displayName,groupTypes,membershipRule,onPremisesSyncEnabled,mailEnabled,securityEnabled" -f $groupId)
            } catch {
                Add-ActionRow @row -Action 'Failed' -Detail "Could not read the group: $($_.Exception.Message)"
                $stats.Failed++
                continue
            }

            # Membership of a dynamic group is computed from a rule, not stored, so there is
            # nothing to remove — editing the rule or the user's attributes is the only way out.
            if (@($group.groupTypes) -contains 'DynamicMembership') {
                Add-ActionRow @row -Action 'CannotRevoke' -Detail 'Dynamic group — membership follows a rule; change the rule or the attributes it matches'
                Write-ProgressHost -Message ("    [SKIP] {0}: dynamic membership, nothing to remove" -f $groupName) -ForegroundColor Yellow
                continue
            }
            # A group mastered on-premises is read-only in the cloud; the change belongs in AD.
            if ($group.onPremisesSyncEnabled) {
                Add-ActionRow @row -Action 'CannotRevoke' -Detail 'Synced from on-premises Active Directory — remove the membership there, it cannot be changed in the cloud'
                Write-ProgressHost -Message ("    [SKIP] {0}: synced from on-premises AD" -f $groupName) -ForegroundColor Yellow
                continue
            }

            # Only a direct member can be removed. Access through a nested group has to be cut
            # at the group that actually holds the user, and saying which one beats a 404.
            $isDirect = $false
            try {
                $direct = Invoke-GraphGet -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/members/{1}?`$select=id" -f $groupId, $entraUser.id)
                $isDirect = [bool]($direct -and $direct.id)
            } catch { $isDirect = $false }

            if (-not $isDirect) {
                Add-ActionRow @row -Action 'CannotRevoke' -Detail 'Not a direct member — the access comes through a nested group, which is where it has to be cut'
                Write-ProgressHost -Message ("    [SKIP] {0}: membership is inherited from a nested group" -f $groupName) -ForegroundColor Yellow
                continue
            }

            $done = Invoke-Revocation -Target "$groupName (Entra ID)" -Operation 'Remove from Entra ID group' -Row $row -Do {
                Invoke-GraphDelete -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/members/{1}/`$ref" -f $groupId, $entraUser.id)
            }
            if ($done) { $stats.EntraRemoved++; $stats.GroupOnly-- } elseif ($Apply) { $stats.Failed++ }
        }
    }
} finally {
    1, 2, 3 | ForEach-Object { Complete-ScanProgress -Id $_ }
    Remove-TempApp
}

# ── Result ────────────────────────────────────────────────────────────────────
# The CSV was written row by row as the run went, so there is nothing to export here.

Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host '   Summary' -ForegroundColor Cyan
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host ("  User              : {0}" -f $UserPrincipalName) -ForegroundColor Cyan
Write-Host ("  Sites searched    : {0}" -f $stats.SitesSearched) -ForegroundColor Cyan
Write-Host ("  Sites with access : {0}" -f $stats.SitesWithAccess) -ForegroundColor Cyan
Write-Host ("  Grants found      : {0}" -f $stats.Found) -ForegroundColor Cyan

if ($script:ActionRows.Count -gt 0) {
    Write-Host ("  Report            : {0}" -f $actionCsv) -ForegroundColor Green
} else {
    Write-Host "  Report            : (nothing found, no file written)" -ForegroundColor DarkGray
}

if ($Apply) {
    Write-Host ("  Revoked           : {0}" -f $stats.Revoked) -ForegroundColor $(if ($stats.Revoked -gt 0) { 'Magenta' } else { 'DarkGray' })
    if ($stats.Failed -gt 0) {
        Write-Host ("  Failed            : {0} — see the CSV, Action = Failed" -f $stats.Failed) -ForegroundColor Red
    }
    if ($stats.AdminLeft -gt 0) {
        Write-Host ''
        Write-Host ("  [FAIL] Still a site collection administrator on {0} site(s)." -f $stats.AdminLeft) -ForegroundColor Red
        Write-Host "         That role reaches every scope in the site, so the other removals there" -ForegroundColor Red
        Write-Host "         changed nothing in practice. Fix this before treating the user as revoked." -ForegroundColor Red
    }
} else {
    Write-Host ''
    Write-Host "  [NOTE] Nothing was changed. Re-run with -Apply to revoke what is listed above." -ForegroundColor Yellow
}

# The routes this script cannot close are the ones most likely to be assumed closed.
if ($stats.EntraRemoved -gt 0) {
    Write-Host ''
    Write-Host ("  Entra groups      : removed from {0}" -f $stats.EntraRemoved) -ForegroundColor Magenta
    Write-Host "  Those memberships often carried more than SharePoint — check Teams, mailboxes and" -ForegroundColor Yellow
    Write-Host "  licences for this user if that was not intended." -ForegroundColor Yellow
}
# Repeated at the end on purpose: the groups removed are the ones the scan found, and a reader
# who only sees the last screen should not take that for "every group that grants access".
if ($RemoveFromEntraGroups -and $scanLimits.Count -gt 0) {
    Write-Host ''
    Write-Host "  [WARN] The Entra groups handled above are the ones this scan found. It did not cover:" -ForegroundColor Yellow
    foreach ($limit in $scanLimits) { Write-Host ("           - {0}" -f $limit) -ForegroundColor Yellow }
    Write-Host "         Re-run without -SiteUrl and at -Scope Item for the complete picture before" -ForegroundColor Yellow
    Write-Host "         treating this user as fully offboarded." -ForegroundColor Yellow
}
if ($RemoveFromEntraGroups -and $script:GrantingEntraGroups.Count -eq 0) {
    Write-Host ''
    Write-Host "  [NOTE] -RemoveFromEntraGroups was given, but no Entra group was seen granting access." -ForegroundColor DarkGray
    Write-Host "         Only groups this run caught holding a role assignment are touched, never every" -ForegroundColor DarkGray
    Write-Host "         group the user belongs to." -ForegroundColor DarkGray
}
if ($stats.GroupOnly -gt 0) {
    Write-Host ''
    Write-Host ("  [WARN] {0} grant(s) reach this user through an Entra ID group and were NOT revoked." -f $stats.GroupOnly) -ForegroundColor Yellow
    Write-Host "         The group is the grant — removing them from SharePoint does not take it away." -ForegroundColor Yellow
    Write-Host "         Filter the CSV on Action = CannotRevoke for the group names to handle in Entra." -ForegroundColor Yellow
}
$everyoneRows = @($script:ActionRows | Where-Object { $_.Action -eq 'CannotRevoke' -and $_.AccessVia -match 'Everyone|authenticated' })
if ($everyoneRows.Count -gt 0) {
    Write-Host ''
    Write-Host ("  [WARN] {0} scope(s) are granted to everyone in the tenant, so this user reaches them" -f $everyoneRows.Count) -ForegroundColor Yellow
    Write-Host "         regardless. Those grants were left alone: removing one revokes access for all." -ForegroundColor Yellow
}
if ($Apply -and $stats.Revoked -gt 0) {
    Write-Host ''
    Write-Host "  [NOTE] Re-run without -Apply to confirm nothing is left, and check the Entra groups above." -ForegroundColor DarkGray
}
Write-Host ''

exit $(if ($stats.Failed -gt 0) { 1 } else { 0 })
