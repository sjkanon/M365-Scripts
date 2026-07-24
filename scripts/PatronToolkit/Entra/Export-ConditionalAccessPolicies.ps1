#Requires -Version 5.1
<#
.SYNOPSIS
    Back up all Conditional Access policies and named locations to JSON/CSV.

.DESCRIPTION
    Connects to Microsoft Graph and exports every Conditional Access policy and named
    location currently configured in the tenant — one JSON file per policy/location plus
    a combined JSON snapshot and a flattened CSV summary. Read-only; makes no changes.

    This is a point-in-time backup/change-tracking tool, distinct from
    Import-ConditionalAccessBaseline.ps1 (which imports a specific community baseline
    into a tenant) — use this script to snapshot what's *already* configured, e.g. before
    a change, or to diff two exports over time.

.PARAMETER OutputPath
    Folder to write the export into. Defaults to
    .\CAPolicyBackup_<timestamp>\ under C:\Temp (Windows) / ~/Downloads (macOS/Linux).

.PARAMETER IncludeNamedLocations
    Also export named locations (IP ranges / countries). Default: on. Use
    -IncludeNamedLocations:$false to skip.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Export-ConditionalAccessPolicies.ps1

.EXAMPLE
    .\Export-ConditionalAccessPolicies.ps1 -OutputPath "C:\Backups\ContosoCA"

.NOTES
    Inspired by the capability list of the retired directorcia/patron toolkit
    (ca-policy-get.ps1 / ca-location-get.ps1), rewritten from scratch against the
    Microsoft.Graph.Identity.SignIns module — the originals used the deprecated
    Microsoft.Graph.Intune module and Connect-MSGraph.
#>
[CmdletBinding()]
param(
    [string] $OutputPath,
    [switch] $IncludeNamedLocations = $true,
    [string] $TenantId
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

if (-not $OutputPath) {
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path $outputDir "CAPolicyBackup_$ts"
}
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }

function ConvertTo-SafeFileName {
    param([string] $Name)
    $safe = $Name -replace '[\\/:*?"<>|]', '-'
    return $safe.Trim()
}

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    $null = Get-MgContext -ErrorAction Stop
    if (-not (Get-MgContext)) { throw }
} catch {
    $connectParams = @{ Scopes = @('Policy.Read.All') }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Conditional Access Policy Backup" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Output folder: $OutputPath" -ForegroundColor DarkGray
Write-Host ""

# ── Policies ──────────────────────────────────────────────────────────────────
Write-Host "  Retrieving Conditional Access policies..." -ForegroundColor DarkGray
try {
    $policies = @(Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop)
} catch {
    Write-Host "  [ERROR] Could not retrieve Conditional Access policies: $($_.Exception.Message)" -ForegroundColor Red
    if ($script:ConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    exit 1
}

$policyDir = Join-Path $OutputPath 'Policies'
New-Item -ItemType Directory -Path $policyDir -Force | Out-Null

$summary = [System.Collections.Generic.List[PSObject]]::new()
foreach ($p in $policies) {
    $fileName = "$(ConvertTo-SafeFileName $p.DisplayName)_$($p.Id).json"
    $p | ConvertTo-Json -Depth 20 | Out-File -FilePath (Join-Path $policyDir $fileName) -Encoding utf8

    $summary.Add([PSCustomObject]@{
        DisplayName        = $p.DisplayName
        Id                 = $p.Id
        State              = $p.State
        CreatedDateTime    = $p.CreatedDateTime
        ModifiedDateTime   = $p.ModifiedDateTime
        IncludeUsers       = ($p.Conditions.Users.IncludeUsers -join '; ')
        ExcludeUsers       = ($p.Conditions.Users.ExcludeUsers -join '; ')
        IncludeApplications = ($p.Conditions.Applications.IncludeApplications -join '; ')
        GrantControls      = ($p.GrantControls.BuiltInControls -join '; ')
    })
    Write-Host "  [OK] $($p.DisplayName) — $($p.State)" -ForegroundColor DarkGray
}
$policies | ConvertTo-Json -Depth 20 | Out-File -FilePath (Join-Path $OutputPath 'AllPolicies.json') -Encoding utf8
$summary | Export-Csv -Path (Join-Path $OutputPath 'PolicySummary.csv') -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "  $($policies.Count) polic(y/ies) exported." -ForegroundColor Green

# ── Named locations ──────────────────────────────────────────────────────────
if ($IncludeNamedLocations) {
    Write-Host ""
    Write-Host "  Retrieving named locations..." -ForegroundColor DarkGray
    try {
        $locations = @(Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop)
        $locations | ConvertTo-Json -Depth 20 | Out-File -FilePath (Join-Path $OutputPath 'NamedLocations.json') -Encoding utf8
        Write-Host "  $($locations.Count) named location(s) exported." -ForegroundColor Green
    } catch {
        Write-Host "  [WARN] Could not retrieve named locations: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "  Backup complete: $OutputPath" -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
