#Requires -Version 5.1
<#
.SYNOPSIS
    Export every Conditional Access policy in the tenant to individual JSON files.

.DESCRIPTION
    Modernized replacement for an old script that used the retired AzureADPreview
    module (Get-AzureADMSConditionalAccessPolicy). Uses Microsoft Graph
    (Get-MgIdentityConditionalAccessPolicy) instead, writing one JSON file per
    policy (named `<PolicyId>.json`) — a point-in-time backup you can diff
    against before making changes, or use as a reference when rebuilding a
    policy.

    Connects to Microsoft Graph automatically if no session is active; reuses an
    existing session if already connected. Read-only — no -Apply switch needed.

.PARAMETER OutputPath
    Folder to write the JSON files to. Defaults to
    `C:\Temp\CAPolicyBackup_<timestamp>\` (`~/Downloads` on Linux/macOS).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Backup-ConditionalAccessPolicies.ps1

.EXAMPLE
    .\Backup-ConditionalAccessPolicies.ps1 -OutputPath "C:\Backups\CA-2026-07-24"

.NOTES
    Required module: Microsoft.Graph.Identity.SignIns (Policy.Read.All)
#>
[CmdletBinding()]
param(
    [string] $OutputPath,
    [string] $TenantId
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputRoot = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not $OutputPath) {
    $OutputPath = Join-Path $outputRoot "CAPolicyBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
}
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    $null = Get-MgContext -ErrorAction Stop
    if (-not (Get-MgContext)) { throw }
} catch {
    $connectParams = @{ Scopes = @('Policy.Read.All'); NoWelcome = $true }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Backup-ConditionalAccessPolicies" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

$policies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop
Write-Host "  Found $($policies.Count) policy(ies)." -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($policy in $policies) {
    $file = Join-Path $OutputPath "$($policy.Id).json"
    try {
        $policy | ConvertTo-Json -Depth 10 | Out-File -FilePath $file -Encoding UTF8
        Write-Host "  [OK]   $($policy.DisplayName)" -ForegroundColor Green
        $results.Add([PSCustomObject]@{ DisplayName = $policy.DisplayName; Id = $policy.Id; State = $policy.State; File = $file; Status = 'Exported' })
    } catch {
        Write-Host "  [WARN] $($policy.DisplayName) : $($_.Exception.Message)" -ForegroundColor Yellow
        $results.Add([PSCustomObject]@{ DisplayName = $policy.DisplayName; Id = $policy.Id; State = $policy.State; File = $null; Status = "Error: $($_.Exception.Message)" })
    }
}

# ── Summary ───────────────────────────────────────────────────────────────────
$results | Export-Csv -Path (Join-Path $OutputPath '_index.csv') -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "  Backed up $($results.Count) polic(ies) to: $OutputPath" -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
