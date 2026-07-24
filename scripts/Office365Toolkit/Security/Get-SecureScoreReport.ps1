#Requires -Version 5.1
<#
.SYNOPSIS
    Report Microsoft Secure Score trend and control-level detail for a tenant.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves the tenant's Secure Score history
    (Get-MgSecuritySecureScore). Reports:
      - Current score, max achievable score, and percentage
      - Score trend over the last -HistoryCount snapshots
      - Per-control breakdown for the latest snapshot (weakest controls first),
        including category, current/max points, and whether it counts toward
        the score

    Results are exported to two CSVs (history + control breakdown).

.PARAMETER HistoryCount
    Number of historical Secure Score snapshots to include in the trend report.
    Default 30 (Secure Score snapshots are generated daily, so this is roughly
    the last month).

.PARAMETER OutputPath
    Folder to write the CSV reports to. Defaults to C:\Temp\ (Windows) or
    ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Get-SecureScoreReport.ps1

.EXAMPLE
    .\Get-SecureScoreReport.ps1 -HistoryCount 90 -OutputPath C:\Reports

.NOTES
    Capability inspired by o365-ssdescpt-get.ps1 from the retired
    directorcia/Office365 (CIAOPS) toolkit, which used a hand-rolled app-only
    token against the beta securescores endpoint. This rewrite uses the
    Microsoft.Graph.Security module's Get-MgSecuritySecureScore (v1.0) instead.

    Required scope: SecurityEvents.Read.All
#>
[CmdletBinding()]
param(
    [int]    $HistoryCount = 30,
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
    $connectParams = @{ Scopes = @('SecurityEvents.Read.All') }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Secure Score Report" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# ── Get score history ────────────────────────────────────────────────────────
Write-Host "  Retrieving Secure Score history..." -ForegroundColor DarkGray
try {
    $allScores = Get-MgSecuritySecureScore -Top $HistoryCount -ErrorAction Stop | Sort-Object CreatedDateTime -Descending
} catch {
    Write-Host "  [ERROR] Could not retrieve Secure Score: $($_.Exception.Message)" -ForegroundColor Red
    if ($script:ConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    exit 1
}

if (-not $allScores -or $allScores.Count -eq 0) {
    Write-Host "  No Secure Score data returned. Secure Score may not be enabled for this tenant yet." -ForegroundColor Yellow
    if ($script:ConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    exit 0
}

$latest = $allScores[0]
$percentage = if ($latest.MaxScore) { [math]::Round(($latest.CurrentScore / $latest.MaxScore) * 100, 1) } else { 0 }

Write-Host "  Snapshot date : $($latest.CreatedDateTime)" -ForegroundColor DarkGray
Write-Host ("  Current score : {0} / {1}  ({2}%)" -f $latest.CurrentScore, $latest.MaxScore, $percentage) -ForegroundColor $(if ($percentage -ge 70) { 'Green' } elseif ($percentage -ge 40) { 'Yellow' } else { 'Red' })
Write-Host ""

# ── Trend ─────────────────────────────────────────────────────────────────────
$history = foreach ($snap in $allScores) {
    [PSCustomObject]@{
        Date          = $snap.CreatedDateTime
        CurrentScore  = $snap.CurrentScore
        MaxScore      = $snap.MaxScore
        Percentage    = if ($snap.MaxScore) { [math]::Round(($snap.CurrentScore / $snap.MaxScore) * 100, 1) } else { 0 }
        ActiveUserCount = $snap.ActiveUserCount
    }
}
$history | Select-Object Date, CurrentScore, MaxScore, Percentage | Format-Table -AutoSize

# ── Control breakdown for latest snapshot (weakest first) ──────────────────────
$controls = foreach ($ctrl in $latest.ControlScores) {
    $maxPts = [double]$ctrl.AdditionalProperties['maxScore']
    $curPts = [double]$ctrl.AdditionalProperties['score']
    [PSCustomObject]@{
        ControlName        = $ctrl.AdditionalProperties['controlName']
        ControlCategory    = $ctrl.AdditionalProperties['controlCategory']
        Description        = $ctrl.AdditionalProperties['description']
        CurrentScore       = $curPts
        MaxScore           = $maxPts
        PercentageComplete = if ($maxPts -gt 0) { [math]::Round(($curPts / $maxPts) * 100, 1) } else { $null }
        ImplementationStatus = $ctrl.AdditionalProperties['implementationStatus']
    }
}
$weakest = $controls | Sort-Object { $_.PercentageComplete } | Select-Object -First 15

Write-Host ""
Write-Host "  Top 15 weakest controls (lowest % complete first):" -ForegroundColor Cyan
$weakest | Select-Object ControlName, ControlCategory, CurrentScore, MaxScore, PercentageComplete, ImplementationStatus | Format-Table -AutoSize

# ── Output ────────────────────────────────────────────────────────────────────
if (-not $OutputPath) { $OutputPath = $outputDir }
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$historyPath  = Join-Path $OutputPath "SecureScoreHistory_$ts.csv"
$controlsPath = Join-Path $OutputPath "SecureScoreControls_$ts.csv"

$history  | Export-Csv -Path $historyPath  -NoTypeInformation -Encoding UTF8
$controls | Export-Csv -Path $controlsPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "  History report  : $historyPath" -ForegroundColor Green
Write-Host "  Controls report : $controlsPath" -ForegroundColor Green
Write-Host ""
Write-Host "  $($history.Count) snapshot(s), $($controls.Count) control(s) in latest snapshot." -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ──────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
