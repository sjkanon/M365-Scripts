#Requires -Version 5.1
<#
.SYNOPSIS
    Report MFA and SSPR registration status for all (or selected) users via Microsoft Graph.

.DESCRIPTION
    Connects to Microsoft Graph and queries the authentication methods user registration
    details report (reports/authenticationMethods/userRegistrationDetails). For every user
    this reports whether they are registered for MFA, whether MFA is enforced/capable,
    SSPR registration/capability, which authentication methods are registered, and their
    default/system-preferred method.

    Users who are NOT MFA-registered are flagged in the console output, with admin
    accounts (IsAdmin = true) called out separately since an unregistered admin account
    is the highest-priority finding. Results are exported to CSV.

.PARAMETER UserList
    One or more UPNs to report on. If omitted, all users are reported.

.PARAMETER AdminsOnly
    Only report on users flagged by Entra ID as directory role holders (IsAdmin = true).

.PARAMETER NotRegisteredOnly
    Only include users who are not MFA-registered.

.PARAMETER OutputPath
    CSV report path. Defaults to .\MfaRegistrationReport_<timestamp>.csv.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Get-MfaRegistrationReport.ps1

.EXAMPLE
    .\Get-MfaRegistrationReport.ps1 -NotRegisteredOnly

.EXAMPLE
    .\Get-MfaRegistrationReport.ps1 -AdminsOnly -NotRegisteredOnly

.EXAMPLE
    .\Get-MfaRegistrationReport.ps1 -UserList "user1@contoso.com","user2@contoso.com"

.NOTES
    Inspired by the capability list of the retired directorcia/patron toolkit
    (graph-usrreg-read.ps1), rewritten from scratch against the modern Graph reports API —
    the original used the deprecated beta credentialUserRegistrationDetails endpoint and
    stored app credentials in local XML files. This script uses the current v1.0 endpoint
    and an interactive/delegated Graph session (or -TenantId for app-only elsewhere).
#>
[CmdletBinding()]
param(
    [string[]] $UserList,
    [switch] $AdminsOnly,
    [switch] $NotRegisteredOnly,
    [string] $OutputPath,
    [string] $TenantId
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    $null = Get-MgContext -ErrorAction Stop
    if (-not (Get-MgContext)) { throw }
} catch {
    $connectParams = @{ Scopes = @('Reports.Read.All', 'AuditLog.Read.All') }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   MFA / SSPR Registration Report" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# ── Retrieve registration details (paginated) ──────────────────────────────────
Write-Host "  Retrieving user registration details..." -ForegroundColor DarkGray

$filterParts = [System.Collections.Generic.List[string]]::new()
if ($AdminsOnly) { $filterParts.Add('isAdmin eq true') }
$filter = $filterParts -join ' and '

$baseUri = 'https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails'
$url = if ($filter) { "$baseUri`?`$filter=$([uri]::EscapeDataString($filter))&`$top=999" } else { "$baseUri`?`$top=999" }

$records = [System.Collections.Generic.List[PSObject]]::new()
try {
    while ($url) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $url -ErrorAction Stop
        foreach ($r in $resp.value) { $records.Add($r) }
        $url = $resp.'@odata.nextLink'
    }
} catch {
    Write-Host "  [ERROR] Could not retrieve registration details: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  [HINT] Requires an Entra ID P1/P2 license and Reports.Read.All (or AuditLog.Read.All) permission." -ForegroundColor Yellow
    if ($script:ConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    exit 1
}

if ($UserList) {
    $userSet = $UserList | ForEach-Object { $_.ToLowerInvariant() }
    $records = $records | Where-Object { $userSet -contains $_.userPrincipalName.ToLowerInvariant() }
}

Write-Host "  Retrieved $($records.Count) user registration record(s)." -ForegroundColor DarkGray
Write-Host ""

# ── Build results ────────────────────────────────────────────────────────────
$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($r in $records) {
    if ($NotRegisteredOnly -and $r.isMfaRegistered) { continue }

    $results.Add([PSCustomObject]@{
        UserPrincipalName    = $r.userPrincipalName
        DisplayName          = $r.userDisplayName
        UserType             = $r.userType
        IsAdmin              = $r.isAdmin
        IsMfaRegistered      = $r.isMfaRegistered
        IsMfaCapable         = $r.isMfaCapable
        IsSsprRegistered     = $r.isSsprRegistered
        IsSsprCapable        = $r.isSsprCapable
        IsSsprEnabled        = $r.isSsprEnabled
        IsPasswordlessCapable = $r.isPasswordlessCapable
        DefaultMfaMethod     = $r.defaultMfaMethod
        MethodsRegistered    = ($r.methodsRegistered -join '; ')
        LastUpdatedDateTime  = $r.lastUpdatedDateTime
    })

    if (-not $r.isMfaRegistered) {
        $color = if ($r.isAdmin) { 'Red' } else { 'Yellow' }
        $tag = if ($r.isAdmin) { '[ADMIN, NOT MFA-REGISTERED]' } else { '[NOT MFA-REGISTERED]' }
        Write-Host "  $tag $($r.userPrincipalName)" -ForegroundColor $color
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
if ($results.Count -eq 0) {
    Write-Host "  No matching users found." -ForegroundColor DarkGray
} else {
    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "MfaRegistrationReport_$ts.csv"
    }
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

$notRegistered = @($results | Where-Object { -not $_.IsMfaRegistered })
$notRegisteredAdmins = @($notRegistered | Where-Object { $_.IsAdmin })
Write-Host ""
Write-Host ("  {0} user(s) reported — {1} not MFA-registered ({2} admin{3})" -f `
    $results.Count, $notRegistered.Count, $notRegisteredAdmins.Count, $(if ($notRegisteredAdmins.Count -ne 1) { 's' } else { '' })) -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
