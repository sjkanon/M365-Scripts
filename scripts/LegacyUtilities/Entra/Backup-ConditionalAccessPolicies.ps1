#Requires -Version 7.0
<#
.SYNOPSIS
    Export every Conditional Access policy in the tenant to individual JSON files.

.DESCRIPTION
    Modernized replacement for an old script that used the retired AzureADPreview
    module (Get-AzureADMSConditionalAccessPolicy). Uses Microsoft Graph
    (GET /identity/conditionalAccess/policies, paged) instead, writing one JSON
    file per policy (named `<PolicyId>.json`) exactly as Graph returns it — a
    point-in-time backup you can diff against before making changes, or use as a
    reference when rebuilding a policy. Read-only — no -Apply switch needed.

    Sign-in goes through scripts\Startup\Connect-M365.ps1: delegated as the admin by
    default (device code / GDAP customer per load.config.ps1), app-only with -ClientId
    and -CertificateThumbprint or -AppOnly. A Graph session for the right tenant that
    already has the scopes is reused and left connected; only a session this script
    opened is disconnected.
    Delegated scope: Policy.Read.All.

.PARAMETER OutputPath
    Folder to write the JSON files to. Defaults to
    `C:\Temp\CAPolicyBackup_<timestamp>\` (`~/Downloads` on Linux/macOS).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Defaults to the GDAP customer when authMode is GDAP.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint and -TenantId).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.EXAMPLE
    .\Backup-ConditionalAccessPolicies.ps1

.EXAMPLE
    .\Backup-ConditionalAccessPolicies.ps1 -OutputPath "C:\Backups\CA-2026-07-24"

.NOTES
    Required module: Microsoft.Graph.Authentication (Policy.Read.All)
#>
[CmdletBinding()]
param(
    [string] $OutputPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

# ── Output folder ─────────────────────────────────────────────────────────────
$outputRoot = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not $OutputPath) {
    $OutputPath = Join-Path $outputRoot "CAPolicyBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
}
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
$graph = Connect-M365Graph -Scopes 'Policy.Read.All' -TenantId $TenantId `
    -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Backup-ConditionalAccessPolicies" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# The raw Graph JSON rather than the SDK objects: ConvertTo-Json on the SDK
# models adds wrapper properties and cut nested conditions off at -Depth 10.
$policies = [System.Collections.Generic.List[object]]::new()
$next = 'v1.0/identity/conditionalAccess/policies'
while ($next) {
    $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType Hashtable -ErrorAction Stop
    foreach ($p in $page['value']) { $policies.Add($p) }
    $next = $page['@odata.nextLink']
}
Write-Host "  Found $($policies.Count) policy(ies)." -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($policy in $policies) {
    $file = Join-Path $OutputPath "$($policy['id']).json"
    try {
        $policy | ConvertTo-Json -Depth 30 | Out-File -FilePath $file -Encoding UTF8
        Write-Host "  [OK]   $($policy['displayName'])" -ForegroundColor Green
        $results.Add([PSCustomObject]@{ DisplayName = $policy['displayName']; Id = $policy['id']; State = $policy['state']; File = $file; Status = 'Exported' })
    } catch {
        Write-Host "  [WARN] $($policy['displayName']) : $($_.Exception.Message)" -ForegroundColor Yellow
        $results.Add([PSCustomObject]@{ DisplayName = $policy['displayName']; Id = $policy['id']; State = $policy['state']; File = $null; Status = "Error: $($_.Exception.Message)" })
    }
}

# ── Summary ───────────────────────────────────────────────────────────────────
$results | Export-Csv -Path (Join-Path $OutputPath '_index.csv') -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "  Backed up $($results.Count) polic(ies) to: $OutputPath" -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
Disconnect-M365Graph $graph
