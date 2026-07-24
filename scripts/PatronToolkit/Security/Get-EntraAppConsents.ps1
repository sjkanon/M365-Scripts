#Requires -Version 5.1
<#
.SYNOPSIS
    Audit OAuth app consent grants (delegated and application permissions) across the tenant.

.DESCRIPTION
    Connects to Microsoft Graph and reports every enterprise application (service principal)
    that has been granted access to the tenant, covering both:
      - Delegated permission grants (OAuth2PermissionGrants) — including whether consent
        was given tenant-wide (AllPrincipals) or by a single user
      - Application permission grants (AppRoleAssignments) — app-only permissions, which
        do not require a signed-in user and are typically higher-impact

    Each grant is flagged High risk when the scope/role name matches a list of
    commonly-abused high-privilege permissions (full tenant mail/file/directory access,
    etc.) — this is a heuristic, not a definitive verdict; review flagged entries manually.
    This is the classic "illicit consent grant" / third-party app risk check MSPs run
    against customer tenants. Results are exported to CSV.

.PARAMETER RiskyOnly
    Only include grants flagged as High risk.

.PARAMETER OutputPath
    CSV report path. Defaults to .\EntraAppConsents_<timestamp>.csv.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Get-EntraAppConsents.ps1

.EXAMPLE
    .\Get-EntraAppConsents.ps1 -RiskyOnly

.NOTES
    Inspired by the capability list of the retired directorcia/patron toolkit
    (az-adappperm-get.ps1, o365-oauth-get.ps1, o365-adal-get.ps1), rewritten from scratch
    against Microsoft Graph — the originals used the deprecated AzureAD module
    (Get-AzureADOAuth2PermissionGrant / Connect-AzureAD), which is retired.

    Required scopes: Application.Read.All, Directory.Read.All
#>
[CmdletBinding()]
param(
    [switch] $RiskyOnly,
    [string] $OutputPath,
    [string] $TenantId
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Risky permission heuristics ─────────────────────────────────────────────────
$riskyPatterns = @(
    'full_access_as_app', 'Mail.ReadWrite', 'Mail.Send', 'MailboxSettings.ReadWrite',
    'Files.ReadWrite.All', 'Sites.FullControl.All', 'Sites.ReadWrite.All',
    'Directory.ReadWrite.All', 'RoleManagement.ReadWrite.Directory',
    'User.ReadWrite.All', 'Group.ReadWrite.All', 'Application.ReadWrite.All',
    'AppRoleAssignment.ReadWrite.All', 'Contacts.ReadWrite', 'Calendars.ReadWrite'
)

function Test-RiskyScope {
    param([string] $Scope)
    foreach ($pattern in $riskyPatterns) {
        if ($Scope -like "*$pattern*") { return $true }
    }
    return $false
}

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    $null = Get-MgContext -ErrorAction Stop
    if (-not (Get-MgContext)) { throw }
} catch {
    $connectParams = @{ Scopes = @('Application.Read.All', 'Directory.Read.All') }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Entra App Consent Audit" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "  Retrieving service principals..." -ForegroundColor DarkGray
$servicePrincipals = @(Get-MgServicePrincipal -All -ErrorAction Stop)
$spById = @{}
foreach ($sp in $servicePrincipals) { $spById[$sp.Id] = $sp }

Write-Host "  Retrieving delegated permission grants..." -ForegroundColor DarkGray
$oauthGrants = @(Get-MgOauth2PermissionGrant -All -ErrorAction Stop)

Write-Host "  Auditing $($servicePrincipals.Count) service principal(s)..." -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()
$userCache = @{}

function Get-CachedUserUpn {
    param([string] $UserId)
    if (-not $UserId) { return $null }
    if ($userCache.ContainsKey($UserId)) { return $userCache[$UserId] }
    try {
        $u = Get-MgUser -UserId $UserId -Property UserPrincipalName -ErrorAction Stop
        $userCache[$UserId] = $u.UserPrincipalName
    } catch {
        $userCache[$UserId] = $UserId
    }
    return $userCache[$UserId]
}

# ── Delegated permission grants ─────────────────────────────────────────────────
foreach ($grant in $oauthGrants) {
    $clientSp = $spById[$grant.ClientId]
    $resourceSp = $spById[$grant.ResourceId]
    $grantedTo = if ($grant.ConsentType -eq 'AllPrincipals') { 'All users (tenant-wide)' } else { Get-CachedUserUpn -UserId $grant.PrincipalId }

    foreach ($scope in ($grant.Scope -split ' ' | Where-Object { $_ })) {
        $results.Add([PSCustomObject]@{
            AppDisplayName   = $clientSp.DisplayName
            AppId            = $clientSp.AppId
            PermissionType   = 'Delegated'
            Permission       = $scope
            ResourceApp      = $resourceSp.DisplayName
            ConsentType      = $grant.ConsentType
            GrantedTo        = $grantedTo
            PublisherDomain  = $clientSp.PublisherName
            SignInAudience   = $clientSp.SignInAudience
            RiskFlag         = if (Test-RiskyScope -Scope $scope) { 'High' } else { 'Normal' }
        })
    }
}

# ── Application permission grants (app roles) ──────────────────────────────────
foreach ($sp in $servicePrincipals) {
    try {
        $appRoleAssignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All -ErrorAction Stop)
    } catch {
        continue
    }

    foreach ($assignment in $appRoleAssignments) {
        $resourceSp = $spById[$assignment.ResourceId]
        $roleName = $assignment.AppRoleId
        if ($resourceSp -and $resourceSp.AppRoles) {
            $role = $resourceSp.AppRoles | Where-Object { $_.Id -eq $assignment.AppRoleId } | Select-Object -First 1
            if ($role) { $roleName = $role.Value }
        }

        $results.Add([PSCustomObject]@{
            AppDisplayName   = $sp.DisplayName
            AppId            = $sp.AppId
            PermissionType   = 'Application'
            Permission       = $roleName
            ResourceApp      = $resourceSp.DisplayName
            ConsentType      = 'Admin (app-only)'
            GrantedTo        = 'N/A (app-only)'
            PublisherDomain  = $sp.PublisherName
            SignInAudience   = $sp.SignInAudience
            RiskFlag         = if (Test-RiskyScope -Scope $roleName) { 'High' } else { 'Normal' }
        })
    }
}

if ($RiskyOnly) {
    $results = $results | Where-Object { $_.RiskFlag -eq 'High' }
}

# ── Output ────────────────────────────────────────────────────────────────────
if ($results.Count -eq 0) {
    Write-Host "  No matching grants found." -ForegroundColor DarkGray
} else {
    $results | Sort-Object RiskFlag -Descending | Format-Table AppDisplayName, PermissionType, Permission, RiskFlag, GrantedTo -AutoSize

    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "EntraAppConsents_$ts.csv"
    }
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

$risky = @($results | Where-Object { $_.RiskFlag -eq 'High' })
Write-Host ""
Write-Host ("  {0} grant(s) reviewed — {1} flagged High risk" -f $results.Count, $risky.Count) -ForegroundColor $(if ($risky.Count -gt 0) { 'Yellow' } else { 'Cyan' })
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
