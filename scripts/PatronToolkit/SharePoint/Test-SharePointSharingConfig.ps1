#Requires -Version 5.1
<#
.SYNOPSIS
    Report SharePoint Online tenant sharing configuration and (optionally) external users.

.DESCRIPTION
    Connects to the SharePoint Online Management Shell and reports the tenant-wide
    external sharing settings that matter most for a security review (sharing capability,
    default link type, anonymous link expiry, anonymous link default permission,
    re-sharing by external users, legacy auth). Optionally also enumerates every external
    (guest) user across all site collections and flags any individual site whose own
    sharing capability is broader than the tenant default.

.PARAMETER TenantName
    SharePoint tenant name (the part before "-admin.sharepoint.com"), e.g. "contoso" for
    https://contoso-admin.sharepoint.com. Required unless -AdminUrl is given or already
    connected via Connect-SPOService.

.PARAMETER AdminUrl
    Full SharePoint admin center URL. Alternative to -TenantName.

.PARAMETER IncludeExternalUsers
    Also enumerate external (guest) users across all site collections. Slower on tenants
    with many sites.

.PARAMETER IncludeSiteOverrides
    Also report individual sites whose SharingCapability is broader than the tenant
    default.

.PARAMETER OutputPath
    Folder for the CSV report(s). Defaults to C:\Temp (Windows) / ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Passed through to Connect-SPOService if supported by
    the installed module version.

.EXAMPLE
    .\Test-SharePointSharingConfig.ps1 -TenantName "contoso"

.EXAMPLE
    .\Test-SharePointSharingConfig.ps1 -TenantName "contoso" -IncludeExternalUsers -IncludeSiteOverrides

.NOTES
    Inspired by the capability list of the retired directorcia/patron toolkit
    (o365-spo-orgconf-get.ps1, o365-spo-extuser-csv.ps1, o365-spo-extavail-csv.ps1,
    o365-spo-user-csv.ps1), consolidated and rewritten from scratch — the originals were
    four separate scripts, one of which compared settings against a hardcoded external
    "best practices" JSON endpoint that no longer belongs to this repo.

    Required module: Microsoft.Online.SharePoint.PowerShell
#>
[CmdletBinding()]
param(
    [string] $TenantName,
    [string] $AdminUrl,
    [switch] $IncludeExternalUsers,
    [switch] $IncludeSiteOverrides,
    [string] $OutputPath,
    [string] $TenantId
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($OutputPath) { $OutputPath } elseif ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

if (-not $AdminUrl -and $TenantName) { $AdminUrl = "https://$TenantName-admin.sharepoint.com" }

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    $null = Get-SPOTenant -ErrorAction Stop
} catch {
    if (-not $AdminUrl) {
        Write-Host "  [ERROR] Not connected to SharePoint Online and no -TenantName/-AdminUrl given." -ForegroundColor Red
        exit 1
    }
    Write-Host "  Connecting to $AdminUrl..." -ForegroundColor DarkGray
    Connect-SPOService -Url $AdminUrl -ErrorAction Stop
    $script:ConnectedHere = $true
}

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

# ── Org sharing settings ──────────────────────────────────────────────────────
$tenant = Get-SPOTenant -ErrorAction Stop

Add-Finding 'Sharing' 'SharingCapability' $tenant.SharingCapability $(if ($tenant.SharingCapability -eq 'ExternalUserAndGuestSharing') { 'Warn' } else { 'Info' })
Add-Finding 'Sharing' 'DefaultSharingLinkType' $tenant.DefaultSharingLinkType $(if ($tenant.DefaultSharingLinkType -eq 'AnonymousAccess') { 'Warn' } else { 'Pass' })
Add-Finding 'Sharing' 'DefaultLinkPermission' $tenant.DefaultLinkPermission 'Info'
Add-Finding 'Sharing' 'RequireAnonymousLinksExpireInDays' $tenant.RequireAnonymousLinksExpireInDays $(if ($tenant.RequireAnonymousLinksExpireInDays -le 0) { 'Warn' } else { 'Pass' })
Add-Finding 'Sharing' 'PreventExternalUsersFromResharing' $tenant.PreventExternalUsersFromResharing $(if ($tenant.PreventExternalUsersFromResharing) { 'Pass' } else { 'Warn' })
Add-Finding 'Sharing' 'SharingDomainRestrictionMode' $tenant.SharingDomainRestrictionMode 'Info'
Add-Finding 'Legacy Auth' 'LegacyAuthProtocolsEnabled' $tenant.LegacyAuthProtocolsEnabled $(if ($tenant.LegacyAuthProtocolsEnabled) { 'Warn' } else { 'Pass' })

# ── Site-level overrides ─────────────────────────────────────────────────────
if ($IncludeSiteOverrides) {
    Write-Host ""
    Write-Host "  Checking per-site sharing overrides..." -ForegroundColor DarkGray
    $sites = @(Get-SPOSite -Limit All -ErrorAction Stop)
    $broaderThanDefault = 0
    foreach ($site in $sites) {
        if ($site.SharingCapability -ne $tenant.SharingCapability -and $site.SharingCapability -eq 'ExternalUserAndGuestSharing') {
            $broaderThanDefault++
            Add-Finding 'Site override' $site.Url "SharingCapability = $($site.SharingCapability) (tenant default: $($tenant.SharingCapability))" 'Warn'
        }
    }
    Write-Host "  $broaderThanDefault site(s) with broader-than-default sharing." -ForegroundColor $(if ($broaderThanDefault -gt 0) { 'Yellow' } else { 'DarkGray' })
}

# ── External users ────────────────────────────────────────────────────────────
$externalUsers = [System.Collections.Generic.List[PSObject]]::new()
if ($IncludeExternalUsers) {
    Write-Host ""
    Write-Host "  Enumerating external users across all site collections..." -ForegroundColor DarkGray
    $sites = if ($IncludeSiteOverrides) { $sites } else { @(Get-SPOSite -Limit All -ErrorAction Stop) }
    foreach ($site in $sites) {
        $position = 0
        do {
            $page = @(Get-SPOExternalUser -SiteUrl $site.Url -PageSize 50 -Position $position -ErrorAction SilentlyContinue)
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

if ($IncludeExternalUsers) {
    $extPath = Join-Path $outputDir "SharePointExternalUsers_$ts.csv"
    if ($externalUsers.Count -gt 0) {
        $externalUsers | Export-Csv -Path $extPath -NoTypeInformation -Encoding UTF8
        Write-Host "  External users report saved: $extPath" -ForegroundColor Green
    }
}

$warnCount = @($findings | Where-Object { $_.Flag -eq 'Warn' }).Count
Write-Host ""
Write-Host ("  {0} finding(s) — {1} warning(s)" -f $findings.Count, $warnCount) -ForegroundColor $(if ($warnCount -gt 0) { 'Yellow' } else { 'Green' })
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-SPOService -ErrorAction SilentlyContinue }
