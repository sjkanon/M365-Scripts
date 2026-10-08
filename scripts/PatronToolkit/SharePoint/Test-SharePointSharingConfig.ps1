#Requires -Version 7.0
<#
.SYNOPSIS
    Report SharePoint Online tenant sharing configuration and (optionally) external users.

.DESCRIPTION
    Reports the tenant-wide external sharing settings that matter most for a security
    review, read from Microsoft Graph (GET /admin/sharepoint/settings): sharing
    capability, re-sharing by external users, whether the accepting account must match
    the invited one, the sharing domain allow/block mode, legacy auth and idle session
    sign-out.

    Three things have no Graph API and come from PnP.PowerShell on the SharePoint admin
    URL, only when asked for:
      - -IncludeLinkSettings: default sharing link type, default link permission and
        anonymous link expiry (Get-PnPTenant)
      - -IncludeSiteOverrides: sites whose own SharingCapability is broader than the
        tenant default (Get-PnPTenantSite)
      - -IncludeExternalUsers: every external user per site collection
        (Get-PnPExternalUser)
    If the PnP sign-in fails, those parts are skipped with a warning and the Graph part
    is still reported.

.PARAMETER TenantName
    SharePoint tenant name (the part before "-admin.sharepoint.com"), e.g. "contoso" for
    https://contoso-admin.sharepoint.com. Only used for the PnP part; when neither this
    nor -AdminUrl is given the admin URL is derived from Graph (/sites/root). Passing it
    (or -AdminUrl) also turns on -IncludeLinkSettings, as the earlier SPO-shell version
    always reported those settings.

.PARAMETER AdminUrl
    Full SharePoint admin center URL. Alternative to -TenantName.

.PARAMETER IncludeLinkSettings
    Also report default link type, default link permission and anonymous link expiry
    (PnP; no Graph API).

.PARAMETER IncludeExternalUsers
    Also enumerate external (guest) users across all site collections (PnP). Slower on
    tenants with many sites.

.PARAMETER IncludeSiteOverrides
    Also report individual sites whose SharingCapability is broader than the tenant
    default (PnP).

.PARAMETER OutputPath
    Folder for the CSV report(s). Defaults to C:\Temp (Windows) / ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Tenant ID or domain. Defaults to the GDAP customer (load.config.ps1) or your own
    tenant. Required for app-only sign-in.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint), used for both
    Graph and PnP. Without it the script signs in delegated, as you.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json (for
    both Graph and PnP).

.PARAMETER PnPClientId
    The PnP app registration for the delegated PnP sign-in. Defaults to the tenant's
    entry in pnp.appid.json (repo root). Not needed for app-only.

.EXAMPLE
    .\Test-SharePointSharingConfig.ps1

.EXAMPLE
    .\Test-SharePointSharingConfig.ps1 -TenantName "contoso"

.EXAMPLE
    .\Test-SharePointSharingConfig.ps1 -IncludeLinkSettings -IncludeExternalUsers -IncludeSiteOverrides

.EXAMPLE
    .\Test-SharePointSharingConfig.ps1 -TenantId contoso.onmicrosoft.com -AppOnly -IncludeSiteOverrides

.NOTES
    Inspired by the capability list of the retired directorcia/patron toolkit
    (o365-spo-orgconf-get.ps1, o365-spo-extuser-csv.ps1, o365-spo-extavail-csv.ps1,
    o365-spo-user-csv.ps1), consolidated and rewritten from scratch — the originals were
    four separate scripts, one of which compared settings against a hardcoded external
    "best practices" JSON endpoint that no longer belongs to this repo.

    The legacy SharePoint Online Management Shell (Connect-SPOService, Get-SPOTenant,
    Get-SPOSite, Get-SPOExternalUser) is no longer used.

    Sign-in, through scripts\Startup\Connect-M365.ps1:
      Graph - delegated by default (scope SharePointTenantSettings.Read.All, plus
              Sites.Read.All only when the admin URL must be looked up; a SharePoint
              admin role; device code / GDAP customer per load.config.ps1), app-only with
              -ClientId/-CertificateThumbprint or -AppOnly (application permission
              SharePointTenantSettings.Read.All, Sites.Read.All).
      PnP   - only for -IncludeLinkSettings / -IncludeSiteOverrides / -IncludeExternalUsers
              (or -TenantName/-AdminUrl). Delegated with -PnPClientId or pnp.appid.json,
              app-only with the same app as Graph (SharePoint application permission
              Sites.FullControl.All for tenant admin calls).

    Required modules: Microsoft.Graph.Authentication; PnP.PowerShell for the PnP part.
#>
[CmdletBinding()]
param(
    [string] $TenantName,
    [string] $AdminUrl,
    [switch] $IncludeLinkSettings,
    [switch] $IncludeExternalUsers,
    [switch] $IncludeSiteOverrides,
    [string] $OutputPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly,
    [string] $PnPClientId
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($OutputPath) { $OutputPath } elseif ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

if (-not $AdminUrl -and $TenantName) { $AdminUrl = "https://$TenantName-admin.sharepoint.com" }
# Backward compatible: naming the tenant used to mean the full SPO-shell report.
if ($AdminUrl) { $IncludeLinkSettings = $true }
$needPnP = $IncludeLinkSettings -or $IncludeSiteOverrides -or $IncludeExternalUsers

# ── Connection (Graph) ────────────────────────────────────────────────────────
$scopes = @('SharePointTenantSettings.Read.All')
if ($needPnP -and -not $AdminUrl) { $scopes += 'Sites.Read.All' }
$graph = Connect-M365Graph -Scopes $scopes -TenantId $TenantId `
    -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   SharePoint Online Sharing Configuration" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

$findings = [System.Collections.Generic.List[PSObject]]::new()
function Add-Finding {
    param([string] $Category, [string] $Setting, [string] $Value, [ValidateSet('Pass', 'Warn', 'Fail', 'Info')] [string] $Flag = 'Info')
    $script:findings.Add([PSCustomObject]@{ Category = $Category; Setting = $Setting; Value = $Value; Flag = $Flag })
    $color = switch ($Flag) { 'Pass' { 'Green' }; 'Warn' { 'Yellow' }; 'Fail' { 'Red' }; default { 'DarkGray' } }
    Write-Host ("  [{0,-4}] {1}: {2}" -f $Flag, $Setting, $Value) -ForegroundColor $color
}

# ── Org sharing settings (Graph) ──────────────────────────────────────────────
try {
    $tenant = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/admin/sharepoint/settings' -OutputType Hashtable -ErrorAction Stop
} catch {
    Write-Host "  [ERROR] Could not read the SharePoint tenant settings: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  [HINT] Needs SharePointTenantSettings.Read.All and (delegated) the SharePoint Administrator role." -ForegroundColor Yellow
    Disconnect-M365Graph $graph
    exit 1
}

$tenantCapability = [string]$tenant.sharingCapability
Add-Finding 'Sharing' 'SharingCapability' $tenantCapability $(if ($tenantCapability -eq 'externalUserAndGuestSharing') { 'Warn' } else { 'Info' })
Add-Finding 'Sharing' 'ResharingByExternalUsersEnabled' $tenant.isResharingByExternalUsersEnabled $(if ($tenant.isResharingByExternalUsersEnabled) { 'Warn' } else { 'Pass' })
Add-Finding 'Sharing' 'RequireAcceptingUserToMatchInvitedUser' $tenant.isRequireAcceptingUserToMatchInvitedUserEnabled $(if ($tenant.isRequireAcceptingUserToMatchInvitedUserEnabled) { 'Pass' } else { 'Warn' })
Add-Finding 'Sharing' 'SharingDomainRestrictionMode' $tenant.sharingDomainRestrictionMode 'Info'
if (@($tenant.sharingAllowedDomainList).Count) { Add-Finding 'Sharing' 'SharingAllowedDomainList' (@($tenant.sharingAllowedDomainList) -join '; ') 'Info' }
if (@($tenant.sharingBlockedDomainList).Count) { Add-Finding 'Sharing' 'SharingBlockedDomainList' (@($tenant.sharingBlockedDomainList) -join '; ') 'Info' }
Add-Finding 'Legacy Auth' 'LegacyAuthProtocolsEnabled' $tenant.isLegacyAuthProtocolsEnabled $(if ($tenant.isLegacyAuthProtocolsEnabled) { 'Warn' } else { 'Pass' })
if ($tenant.idleSessionSignOut) {
    Add-Finding 'Session' 'IdleSessionSignOut' $tenant.idleSessionSignOut.isEnabled $(if ($tenant.idleSessionSignOut.isEnabled) { 'Pass' } else { 'Info' })
}

# ── PnP part (no Graph API) ───────────────────────────────────────────────────
$pnp = $null
$externalUsers = [System.Collections.Generic.List[PSObject]]::new()
if ($needPnP) {
    if (-not $AdminUrl) {
        try {
            $root = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl' -OutputType Hashtable -ErrorAction Stop
            if ($root.webUrl -match '^https://([^./]+)\.sharepoint\.com') { $AdminUrl = "https://$($Matches[1])-admin.sharepoint.com" }
        } catch {
            Write-Host "  [WARN] Could not look up the SharePoint URL via Graph: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    if (-not $AdminUrl) {
        Write-Host "  [WARN] No SharePoint admin URL — pass -TenantName or -AdminUrl. Skipping link settings, site overrides and external users." -ForegroundColor Yellow
    } else {
        $pnpParams = @{ Url = $AdminUrl; TenantId = $TenantId }
        if ($AppOnly -and -not $ClientId) {
            try {
                $reg = Get-M365AppRegistration -TenantId (Resolve-M365TenantId -TenantId $TenantId)
                $pnpParams['ClientId'] = $reg.ClientId
                $pnpParams['CertificateThumbprint'] = $reg.CertificateThumbprint
                if (-not $pnpParams['TenantId']) { $pnpParams['TenantId'] = $reg.Tenant }
            } catch { Write-Host "  [WARN] $($_.Exception.Message)" -ForegroundColor Yellow }
        } elseif ($ClientId) {
            $pnpParams['ClientId'] = $ClientId
            $pnpParams['CertificateThumbprint'] = $CertificateThumbprint
        } elseif ($PnPClientId) {
            $pnpParams['ClientId'] = $PnPClientId
        }
        try {
            Write-Host ""
            Write-Host "  Connecting PnP to $AdminUrl..." -ForegroundColor DarkGray
            $pnp = Connect-M365PnP @pnpParams
        } catch {
            Write-Host "  [WARN] PnP sign-in failed, skipping link settings, site overrides and external users: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

if ($pnp -and $IncludeLinkSettings) {
    try {
        $spo = Get-PnPTenant -Connection $pnp -ErrorAction Stop
        Add-Finding 'Links' 'DefaultSharingLinkType' $spo.DefaultSharingLinkType $(if ([string]$spo.DefaultSharingLinkType -eq 'AnonymousAccess') { 'Warn' } else { 'Pass' })
        Add-Finding 'Links' 'DefaultLinkPermission' $spo.DefaultLinkPermission 'Info'
        Add-Finding 'Links' 'RequireAnonymousLinksExpireInDays' $spo.RequireAnonymousLinksExpireInDays $(if ($spo.RequireAnonymousLinksExpireInDays -le 0) { 'Warn' } else { 'Pass' })
    } catch {
        Write-Host "  [WARN] Could not read link settings (Get-PnPTenant): $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

$sites = $null
if ($pnp -and ($IncludeSiteOverrides -or $IncludeExternalUsers)) {
    try {
        $sites = @(Get-PnPTenantSite -Connection $pnp -ErrorAction Stop)
    } catch {
        Write-Host "  [WARN] Could not list sites (Get-PnPTenantSite): $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# ── Site-level overrides ─────────────────────────────────────────────────────
if ($sites -and $IncludeSiteOverrides) {
    Write-Host ""
    Write-Host "  Checking per-site sharing overrides..." -ForegroundColor DarkGray
    $broaderThanDefault = 0
    foreach ($site in $sites) {
        $siteCapability = [string]$site.SharingCapability
        if ($siteCapability -ne $tenantCapability -and $siteCapability -eq 'ExternalUserAndGuestSharing') {
            $broaderThanDefault++
            Add-Finding 'Site override' $site.Url "SharingCapability = $siteCapability (tenant default: $tenantCapability)" 'Warn'
        }
    }
    Write-Host "  $broaderThanDefault site(s) with broader-than-default sharing." -ForegroundColor $(if ($broaderThanDefault -gt 0) { 'Yellow' } else { 'DarkGray' })
}

# ── External users ────────────────────────────────────────────────────────────
if ($sites -and $IncludeExternalUsers) {
    Write-Host ""
    Write-Host "  Enumerating external users across $($sites.Count) site collection(s)..." -ForegroundColor DarkGray
    foreach ($site in $sites) {
        $position = 0
        do {
            try {
                $page = @(Get-PnPExternalUser -SiteUrl $site.Url -PageSize 50 -Position $position -Connection $pnp -ErrorAction Stop)
            } catch {
                Write-Host "  [WARN] $($site.Url): $($_.Exception.Message)" -ForegroundColor Yellow
                break
            }
            foreach ($u in $page) {
                $externalUsers.Add([PSCustomObject]@{
                    SiteUrl     = $site.Url
                    DisplayName = $u.DisplayName
                    Email       = $u.Email
                    AcceptedAs  = $u.AcceptedAs
                    WhenCreated = $u.WhenCreated
                    InvitedBy   = $u.InvitedBy
                })
            }
            $position += 50
        } while ($page.Count -eq 50)
    }
    Write-Host "  Found $($externalUsers.Count) external user membership(s)." -ForegroundColor DarkGray
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$configPath = Join-Path $outputDir "SharePointSharingConfig_$ts.csv"
$findings | Export-Csv -Path $configPath -NoTypeInformation -Encoding UTF8
Write-Host "  Config report saved: $configPath" -ForegroundColor Green

if ($IncludeExternalUsers -and $externalUsers.Count -gt 0) {
    $extPath = Join-Path $outputDir "SharePointExternalUsers_$ts.csv"
    $externalUsers | Export-Csv -Path $extPath -NoTypeInformation -Encoding UTF8
    Write-Host "  External users report saved: $extPath" -ForegroundColor Green
}

$warnCount = @($findings | Where-Object { $_.Flag -eq 'Warn' }).Count
Write-Host ""
Write-Host ("  {0} finding(s) — {1} warning(s)" -f $findings.Count, $warnCount) -ForegroundColor $(if ($warnCount -gt 0) { 'Yellow' } else { 'Green' })
Write-Host ""

# ── Disconnect what we connected ──────────────────────────────────────────────
# The PnP connection object is local to this script (-ReturnConnection); nothing global to close.
Disconnect-M365Graph $graph
