#Requires -Version 7.0
<#
.SYNOPSIS
    Report Microsoft Secure Score trend and control-level detail for a tenant.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves the tenant's Secure Score history
    (/security/secureScores) and the control profiles behind it
    (/security/secureScoreControlProfiles). Reports:
      - Current score, max achievable score, and percentage
      - Score trend over the last -HistoryCount snapshots
      - Per-control breakdown for the latest snapshot (weakest controls first),
        including title, category, current/max points and implementation status

    Results are exported to two CSVs (history + control breakdown).

    Sign-in: delegated (you sign in as the admin) by default, through
    scripts\Startup\Connect-M365.ps1 — device code and the GDAP customer come
    from load.config.ps1. App-only with -ClientId + -CertificateThumbprint, or
    -AppOnly (graph.appid.json). An existing Graph session with the right scope
    is reused and left connected.

.PARAMETER HistoryCount
    Number of historical Secure Score snapshots to include in the trend report.
    Default 30 (Secure Score snapshots are generated daily, so this is roughly
    the last month).

.PARAMETER OutputPath
    Folder to write the CSV reports to. Defaults to C:\Temp\ (Windows) or
    ~/Downloads (macOS/Linux).

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
    .\Get-SecureScoreReport.ps1

.EXAMPLE
    .\Get-SecureScoreReport.ps1 -HistoryCount 90 -OutputPath C:\Reports

.NOTES
    Capability inspired by o365-ssdescpt-get.ps1 from the retired
    directorcia/Office365 (CIAOPS) toolkit, which used a hand-rolled app-only
    token against the beta securescores endpoint.

    A controlScore in a snapshot carries the points scored, not the maximum.
    The maximum (and the readable title) come from the matching
    secureScoreControlProfile, joined on controlName = profile id. Earlier
    versions read controlName/maxScore from AdditionalProperties, where the SDK
    never puts typed properties, so the control table came out empty.

    Delegated scope: SecurityEvents.Read.All (app-only: the same as an
    application permission).
    Required module: Microsoft.Graph.Authentication
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 1000)]
    [int]    $HistoryCount = 30,
    [string] $OutputPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
$graph = Connect-M365Graph -Scopes 'SecurityEvents.Read.All' `
    -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

function Get-GraphCollection {
    # GET a collection and follow @odata.nextLink until every page is read.
    param([string] $Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    while ($Uri) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType Hashtable -ErrorAction Stop
        foreach ($v in $page['value']) { $items.Add($v) }
        $Uri = $page['@odata.nextLink']
    }
    , $items
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
    # One page with $top only; no nextLink paging, which would run past -HistoryCount.
    $page = Invoke-MgGraphRequest -Method GET -OutputType Hashtable -ErrorAction Stop `
        -Uri "https://graph.microsoft.com/v1.0/security/secureScores?`$top=$HistoryCount"
    $allScores = @($page['value'] | Sort-Object { [datetime]$_['createdDateTime'] } -Descending | Select-Object -First $HistoryCount)
} catch {
    Write-Host "  [ERROR] Could not retrieve Secure Score: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-M365Graph $graph
    exit 1
}

if ($allScores.Count -eq 0) {
    Write-Host "  No Secure Score data returned. Secure Score may not be enabled for this tenant yet." -ForegroundColor Yellow
    Disconnect-M365Graph $graph
    exit 0
}

$latest = $allScores[0]
$percentage = if ($latest['maxScore']) { [math]::Round(($latest['currentScore'] / $latest['maxScore']) * 100, 1) } else { 0 }

Write-Host "  Snapshot date : $($latest['createdDateTime'])" -ForegroundColor DarkGray
Write-Host ("  Current score : {0} / {1}  ({2}%)" -f $latest['currentScore'], $latest['maxScore'], $percentage) -ForegroundColor $(if ($percentage -ge 70) { 'Green' } elseif ($percentage -ge 40) { 'Yellow' } else { 'Red' })
Write-Host ""

# ── Trend ─────────────────────────────────────────────────────────────────────
$history = foreach ($snap in $allScores) {
    [PSCustomObject]@{
        Date            = $snap['createdDateTime']
        CurrentScore    = $snap['currentScore']
        MaxScore        = $snap['maxScore']
        Percentage      = if ($snap['maxScore']) { [math]::Round(($snap['currentScore'] / $snap['maxScore']) * 100, 1) } else { 0 }
        ActiveUserCount = $snap['activeUserCount']
    }
}
$history | Select-Object Date, CurrentScore, MaxScore, Percentage | Format-Table -AutoSize

# ── Control profiles (max score + title per control) ─────────────────────────
$profiles = @{}
try {
    foreach ($p in (Get-GraphCollection 'https://graph.microsoft.com/v1.0/security/secureScoreControlProfiles')) {
        $profiles[[string]$p['id']] = $p
    }
} catch {
    Write-Host "  [WARN] Could not read control profiles, max points will be empty: $($_.Exception.Message)" -ForegroundColor Yellow
}

# ── Control breakdown for latest snapshot (weakest first) ──────────────────────
$controls = foreach ($ctrl in $latest['controlScores']) {
    $prof    = $profiles[[string]$ctrl['controlName']]
    if (-not $prof) { $prof = @{} }
    $curPts  = [double]$ctrl['score']
    $maxPts  = if ($null -ne $prof['maxScore']) { [double]$prof['maxScore'] } else { $null }
    $pct     = if ($maxPts -gt 0) { [math]::Round(($curPts / $maxPts) * 100, 1) }
               elseif ($null -ne $ctrl['scoreInPercentage']) { [double]$ctrl['scoreInPercentage'] }
               else { $null }
    [PSCustomObject]@{
        ControlName          = $ctrl['controlName']
        Title                = $prof['title']
        ControlCategory      = $ctrl['controlCategory']
        Description          = $ctrl['description']
        CurrentScore         = $curPts
        MaxScore             = $maxPts
        PercentageComplete   = $pct
        ImplementationStatus = $ctrl['implementationStatus']
        Deprecated           = [bool]$prof['deprecated']
    }
}
# Controls that are already complete or carry no points are not "weak".
$weakest = $controls | Where-Object { $null -ne $_.PercentageComplete -and $_.PercentageComplete -lt 100 -and -not $_.Deprecated } |
    Sort-Object PercentageComplete, @{ Expression = 'MaxScore'; Descending = $true } | Select-Object -First 15

Write-Host ""
Write-Host "  Top 15 weakest controls (lowest % complete first):" -ForegroundColor Cyan
$weakest | Select-Object ControlName, ControlCategory, CurrentScore, MaxScore, PercentageComplete | Format-Table -AutoSize

# ── Output ────────────────────────────────────────────────────────────────────
if (-not $OutputPath) { $OutputPath = $outputDir }
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$historyPath  = Join-Path $OutputPath "SecureScoreHistory_$ts.csv"
$controlsPath = Join-Path $OutputPath "SecureScoreControls_$ts.csv"

$history  | Export-Csv -Path $historyPath  -NoTypeInformation -Encoding UTF8
$controls | Export-Csv -Path $controlsPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "  History report  : $historyPath" -ForegroundColor Green
Write-Host "  Controls report : $controlsPath" -ForegroundColor Green
Write-Host ""
Write-Host "  $(@($history).Count) snapshot(s), $(@($controls).Count) control(s) in latest snapshot." -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ──────────────────────────────────────────────
Disconnect-M365Graph $graph
