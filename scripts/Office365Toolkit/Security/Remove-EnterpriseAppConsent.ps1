#Requires -Version 5.1
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
    Entra ID tenant ID or domain. Optional if already connected.

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
    This rewrite uses Microsoft.Graph.Applications instead.

    Supports -WhatIf (SupportsShouldProcess) for the actual revocations.

    Required scopes: Application.Read.All, DelegatedPermissionGrant.ReadWrite.All,
    AppRoleAssignment.ReadWrite.All

    Required module: Microsoft.Graph.Applications
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [string] $AppId,
    [string] $AppDisplayName,
    [switch] $IncludeUserConsent,
    [switch] $Apply,
    [string] $OutputPath,
    [string] $TenantId
)

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
$script:ConnectedHere = $false
try {
    $null = Get-MgContext -ErrorAction Stop
    if (-not (Get-MgContext)) { throw }
} catch {
    $connectParams = @{
        Scopes = @('Application.Read.All', 'DelegatedPermissionGrant.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
    }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
}

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
        $sps = @(Get-MgServicePrincipal -Filter "displayName eq '$AppDisplayName'" -ErrorAction Stop)
        if ($sps.Count -eq 0) { throw "No service principal found with display name '$AppDisplayName'." }
        if ($sps.Count -gt 1) { throw "Multiple service principals match display name '$AppDisplayName'. Use -AppId instead." }
        $sp = $sps[0]
    } else {
        try {
            $sp = Get-MgServicePrincipal -ServicePrincipalId $AppId -ErrorAction Stop
        } catch {
            $sps = @(Get-MgServicePrincipal -Filter "appId eq '$AppId'" -ErrorAction Stop)
            if ($sps.Count -eq 0) { throw "No service principal found for AppId/ObjectId '$AppId'." }
            $sp = $sps[0]
        }
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    if ($script:ConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    exit 1
}

Write-Host "  Target app : $($sp.DisplayName)  [$($sp.AppId)]" -ForegroundColor DarkGray
Write-Host ""

# ── Delegated permission grants (OAuth2PermissionGrants) ───────────────────────
Write-Host "  Retrieving delegated permission grants..." -ForegroundColor DarkGray
$allGrants = Get-MgOauth2PermissionGrant -All -ErrorAction SilentlyContinue | Where-Object { $_.ClientId -eq $sp.Id }
$grants = $allGrants | Where-Object { $IncludeUserConsent -or $_.ConsentType -eq 'AllPrincipals' }

# ── Application permission (app role) assignments ──────────────────────────────
Write-Host "  Retrieving application permission assignments..." -ForegroundColor DarkGray
$appRoleAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All -ErrorAction SilentlyContinue

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($grant in $grants) {
    $resourceSp = $null
    try { $resourceSp = Get-MgServicePrincipal -ServicePrincipalId $grant.ResourceId -ErrorAction SilentlyContinue } catch {}
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
    $results.Add([PSCustomObject]@{
        GrantType    = 'Application'
        ResourceName = $assignment.ResourceDisplayName
        ConsentType  = 'AllPrincipals'
        Principal    = $null
        ScopeOrRole  = $assignment.AppRoleId
        GrantId      = $assignment.Id
    })
}

if ($results.Count -eq 0) {
    Write-Host ""
    Write-Host "  No consent grants found for this app (with current -IncludeUserConsent setting)." -ForegroundColor DarkGray
    if ($script:ConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
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
if ($script:ConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
