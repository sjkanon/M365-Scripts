#Requires -Version 7.0
<#
.SYNOPSIS
    Audit and optionally revoke OAuth consent grants for an enterprise application.

.DESCRIPTION
    Connects to Microsoft Graph and reports every delegated (OAuth2PermissionGrant)
    and application (AppRoleAssignment) permission held by a specific enterprise
    application (service principal) — the kind of consent screen a user or admin
    accepted when signing in to a third-party or in-house app. Useful for reviewing
    illicit consent grants (a common phishing/BEC technique) or cleaning up an app
    before removing it.

    Default behavior is a safe report-only preview. Pass -Apply to actually revoke
    the grants that were reported.

    Exactly one of -AppId or -AppDisplayName selects the target application, to
    avoid accidentally mass-revoking consent across every enterprise app in the
    tenant in a single run.

    Sign-in: delegated (you sign in as the admin) by default, through
    scripts\Startup\Connect-M365.ps1 — device code and the GDAP customer come
    from load.config.ps1. App-only with -ClientId + -CertificateThumbprint, or
    -AppOnly (graph.appid.json). A preview asks only for read scopes; -Apply
    adds the ReadWrite scopes. An existing Graph session with the right scopes
    is reused and left connected.

.PARAMETER AppId
    The Application (client) ID, or the service principal's Object ID, of the
    enterprise app to audit.

.PARAMETER AppDisplayName
    Display name of the enterprise app to audit (exact match). Errors if more
    than one service principal matches.

.PARAMETER IncludeUserConsent
    Also report/revoke per-user delegated consent grants (ConsentType = Principal).
    Without this switch, only tenant-wide admin consent grants (ConsentType =
    AllPrincipals) and application permissions are processed.

.PARAMETER Apply
    Actually revoke the reported grants. Without this switch, the script only
    reports what would be revoked.

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\ (Windows) or ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Defaults to the GDAP customer (load.config.ps1),
    else the tenant you sign in to.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.EXAMPLE
    # Report only — what does this app currently have consent to?
    .\Remove-EnterpriseAppConsent.ps1 -AppDisplayName "Suspicious Reporting Tool"

.EXAMPLE
    # Revoke tenant-wide admin consent and application permissions
    .\Remove-EnterpriseAppConsent.ps1 -AppDisplayName "Suspicious Reporting Tool" -Apply

.EXAMPLE
    # Also include and revoke individual users' delegated consent grants
    .\Remove-EnterpriseAppConsent.ps1 -AppId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -IncludeUserConsent -Apply

.NOTES
    Capability inspired by graph-adappperm-del.ps1 from the retired
    directorcia/Office365 (CIAOPS) toolkit, which used the retired AzureAD
    module (Get-AzureADServicePrincipal / Remove-AzureADOAuth2PermissionGrant).
    This rewrite uses Microsoft Graph instead.

    Delegated grants are read server-side with $filter=clientId eq '<sp id>'
    (earlier versions pulled every grant in the tenant and filtered locally).
    Application permissions are shown by their role value (e.g. Mail.Read),
    resolved from the resource's appRoles, instead of a bare GUID.

    Supports -WhatIf (SupportsShouldProcess) for the actual revocations.

    Delegated scopes (preview): Application.Read.All, DelegatedPermissionGrant.Read.All,
    User.ReadBasic.All. With -Apply also: DelegatedPermissionGrant.ReadWrite.All,
    AppRoleAssignment.ReadWrite.All. App-only: the matching application permissions.

    Required modules: Microsoft.Graph.Applications, Microsoft.Graph.Identity.SignIns,
    Microsoft.Graph.Users
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [string] $AppId,
    [string] $AppDisplayName,
    [switch] $IncludeUserConsent,
    [switch] $Apply,
    [string] $OutputPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

if (-not $AppId -and -not $AppDisplayName) {
    throw "Specify -AppId or -AppDisplayName to target a single enterprise application."
}
if ($AppId -and $AppDisplayName) {
    throw "Specify only one of -AppId or -AppDisplayName."
}

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
$scopes = @('Application.Read.All', 'DelegatedPermissionGrant.Read.All', 'User.ReadBasic.All')
if ($Apply) { $scopes += 'DelegatedPermissionGrant.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All' }
$graph = Connect-M365Graph -Scopes $scopes `
    -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Enterprise App Consent Audit" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host ("  Mode : {0}" -f $(if ($Apply) { 'Apply (grants will be revoked)' } else { 'Preview only (no changes)' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

# ── Resolve service principal ───────────────────────────────────────────────
try {
    if ($AppDisplayName) {
        $sps = @(Get-MgServicePrincipal -Filter "displayName eq '$($AppDisplayName -replace "'", "''")'" -All -ErrorAction Stop)
        if ($sps.Count -eq 0) { throw "No service principal found with display name '$AppDisplayName'." }
        if ($sps.Count -gt 1) { throw "Multiple service principals match display name '$AppDisplayName'. Use -AppId instead." }
        $sp = $sps[0]
    } else {
        try {
            $sp = Get-MgServicePrincipal -ServicePrincipalId $AppId -ErrorAction Stop
        } catch {
            $sps = @(Get-MgServicePrincipal -Filter "appId eq '$($AppId -replace "'", "''")'" -ErrorAction Stop)
            if ($sps.Count -eq 0) { throw "No service principal found for AppId/ObjectId '$AppId'." }
            $sp = $sps[0]
        }
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-M365Graph $graph
    exit 1
}

Write-Host "  Target app : $($sp.DisplayName)  [$($sp.AppId)]" -ForegroundColor DarkGray
Write-Host ""

# ── Delegated permission grants (OAuth2PermissionGrants) ───────────────────────
Write-Host "  Retrieving delegated permission grants..." -ForegroundColor DarkGray
try {
    # Filter server-side on this app's service principal instead of reading every grant in the tenant.
    $allGrants = @(Get-MgOauth2PermissionGrant -Filter "clientId eq '$($sp.Id)'" -All -ErrorAction Stop)
} catch {
    Write-Host "  [ERROR] Could not read delegated permission grants: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-M365Graph $graph
    exit 1
}
$grants = @($allGrants | Where-Object { $IncludeUserConsent -or $_.ConsentType -eq 'AllPrincipals' })
$skippedUserGrants = $allGrants.Count - $grants.Count

# ── Application permission (app role) assignments ──────────────────────────────
Write-Host "  Retrieving application permission assignments..." -ForegroundColor DarkGray
try {
    $appRoleAssignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All -ErrorAction Stop)
} catch {
    Write-Host "  [ERROR] Could not read application permission assignments: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-M365Graph $graph
    exit 1
}

$results = [System.Collections.Generic.List[PSObject]]::new()

# Resource service principals (Microsoft Graph, SharePoint, ...), looked up once each.
$resourceCache = @{}
function Get-ResourceSp([string] $Id) {
    if (-not $resourceCache.ContainsKey($Id)) {
        $resourceCache[$Id] = try { Get-MgServicePrincipal -ServicePrincipalId $Id -Property Id, DisplayName, AppRoles -ErrorAction Stop } catch { $null }
    }
    $resourceCache[$Id]
}

foreach ($grant in $grants) {
    $resourceSp = Get-ResourceSp $grant.ResourceId
    $principalUpn = $null
    if ($grant.ConsentType -eq 'Principal' -and $grant.PrincipalId) {
        try { $principalUpn = (Get-MgUser -UserId $grant.PrincipalId -ErrorAction SilentlyContinue).UserPrincipalName } catch {}
    }
    $results.Add([PSCustomObject]@{
        GrantType    = 'Delegated'
        ResourceName = $resourceSp.DisplayName
        ConsentType  = $grant.ConsentType
        Principal    = $principalUpn
        ScopeOrRole  = $grant.Scope
        GrantId      = $grant.Id
    })
}

foreach ($assignment in $appRoleAssignments) {
    $role = (Get-ResourceSp $assignment.ResourceId).AppRoles | Where-Object { $_.Id -eq $assignment.AppRoleId } | Select-Object -First 1
    $results.Add([PSCustomObject]@{
        GrantType    = 'Application'
        ResourceName = $assignment.ResourceDisplayName
        ConsentType  = 'AllPrincipals'
        Principal    = $null
        ScopeOrRole  = $(if ($role.Value) { $role.Value } else { $assignment.AppRoleId })
        GrantId      = $assignment.Id
    })
}

if ($skippedUserGrants -gt 0) {
    Write-Host "  $skippedUserGrants per-user delegated grant(s) not included (pass -IncludeUserConsent to report/revoke them)." -ForegroundColor DarkGray
}

if ($results.Count -eq 0) {
    Write-Host ""
    Write-Host "  No consent grants found for this app (with current -IncludeUserConsent setting)." -ForegroundColor DarkGray
    Disconnect-M365Graph $graph
    exit 0
}

Write-Host ""
$results | Format-Table GrantType, ResourceName, ConsentType, Principal, ScopeOrRole -AutoSize

# ── Output ────────────────────────────────────────────────────────────────────
if (-not $OutputPath) {
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path $outputDir "EnterpriseAppConsent_$($sp.DisplayName -replace '[^\w-]','_')_$ts.csv"
}
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "  Report saved: $OutputPath" -ForegroundColor Green

# ── Revoke ────────────────────────────────────────────────────────────────────
if (-not $Apply) {
    Write-Host ""
    Write-Host "  $($results.Count) grant(s) would be revoked. Re-run with -Apply to revoke them." -ForegroundColor Yellow
} else {
    Write-Host ""
    $revoked = 0
    $failed  = 0
    foreach ($grant in $grants) {
        if ($PSCmdlet.ShouldProcess("Delegated grant $($grant.Id) on app $($sp.DisplayName)", "Revoke")) {
            try {
                Remove-MgOauth2PermissionGrant -OAuth2PermissionGrantId $grant.Id -ErrorAction Stop
                Write-Host "  [OK]   Revoked delegated grant $($grant.Id)" -ForegroundColor Green
                $revoked++
            } catch {
                Write-Host "  [WARN] Failed to revoke delegated grant $($grant.Id): $($_.Exception.Message)" -ForegroundColor Yellow
                $failed++
            }
        }
    }
    foreach ($assignment in $appRoleAssignments) {
        if ($PSCmdlet.ShouldProcess("Application permission $($assignment.Id) on app $($sp.DisplayName)", "Revoke")) {
            try {
                Remove-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -AppRoleAssignmentId $assignment.Id -ErrorAction Stop
                Write-Host "  [OK]   Revoked application permission $($assignment.Id)" -ForegroundColor Green
                $revoked++
            } catch {
                Write-Host "  [WARN] Failed to revoke application permission $($assignment.Id): $($_.Exception.Message)" -ForegroundColor Yellow
                $failed++
            }
        }
    }
    Write-Host ""
    Write-Host "  Revoked : $revoked" -ForegroundColor Green
    if ($failed -gt 0) { Write-Host "  Failed  : $failed" -ForegroundColor Yellow }
}

Write-Host ""

# ── Disconnect if we connected ──────────────────────────────────────────────
Disconnect-M365Graph $graph
