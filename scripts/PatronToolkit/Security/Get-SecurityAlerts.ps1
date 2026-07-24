#Requires -Version 5.1
<#
.SYNOPSIS
    Report Microsoft Defender / Entra security alerts for a tenant via Microsoft Graph.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves security alerts (the unified alerts_v2 feed
    that spans Defender for Office 365, Defender for Endpoint, Defender for Identity,
    Defender for Cloud Apps and Entra ID Protection) for a given lookback window. Results
    are sorted by severity/creation time and exported to CSV — useful as a recurring
    "what needs attention" check across a managed tenant.

.PARAMETER Days
    How many days back to look. Default: 30.

.PARAMETER Severity
    Filter to one or more severities: informational, low, medium, high.

.PARAMETER Status
    Filter to one or more statuses: new, inProgress, resolved.

.PARAMETER OutputPath
    CSV report path. Defaults to .\SecurityAlerts_<timestamp>.csv.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Get-SecurityAlerts.ps1

.EXAMPLE
    .\Get-SecurityAlerts.ps1 -Days 7 -Severity high,medium -Status new,inProgress

.NOTES
    Inspired by the capability list of the retired directorcia/patron toolkit
    (graph-alerts-get.ps1), rewritten from scratch against the current
    security/alerts_v2 endpoint — the original called the older, now-superseded
    /security/alerts endpoint using a manually-constructed OAuth token read from local
    encrypted XML credential files.

    Required scope: SecurityAlert.Read.All
#>
[CmdletBinding()]
param(
    [int] $Days = 30,
    [ValidateSet('informational', 'low', 'medium', 'high')]
    [string[]] $Severity,
    [ValidateSet('new', 'inProgress', 'resolved')]
    [string[]] $Status,
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
    $connectParams = @{ Scopes = @('SecurityAlert.Read.All') }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Security Alerts Report (last $Days day(s))" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

$since = (Get-Date).ToUniversalTime().AddDays(-$Days).ToString('yyyy-MM-ddTHH:mm:ssZ')
$filterParts = [System.Collections.Generic.List[string]]::new()
$filterParts.Add("createdDateTime ge $since")
if ($Severity) {
    $sevFilter = ($Severity | ForEach-Object { "severity eq '$_'" }) -join ' or '
    $filterParts.Add("($sevFilter)")
}
if ($Status) {
    $statusFilter = ($Status | ForEach-Object { "status eq '$_'" }) -join ' or '
    $filterParts.Add("($statusFilter)")
}
$filter = $filterParts -join ' and '

$url = "https://graph.microsoft.com/v1.0/security/alerts_v2?`$filter=$([uri]::EscapeDataString($filter))&`$top=999"

Write-Host "  Retrieving alerts..." -ForegroundColor DarkGray
$alerts = [System.Collections.Generic.List[PSObject]]::new()
try {
    while ($url) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $url -ErrorAction Stop
        foreach ($a in $resp.value) { $alerts.Add($a) }
        $url = $resp.'@odata.nextLink'
    }
} catch {
    Write-Host "  [ERROR] Could not retrieve security alerts: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  [HINT] Requires SecurityAlert.Read.All and an eligible Defender/Entra ID Protection license." -ForegroundColor Yellow
    if ($script:ConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    exit 1
}

Write-Host "  Retrieved $($alerts.Count) alert(s)." -ForegroundColor DarkGray
Write-Host ""

$results = foreach ($a in $alerts) {
    [PSCustomObject]@{
        Title           = $a.title
        Severity        = $a.severity
        Status          = $a.status
        Category        = $a.category
        ServiceSource   = $a.serviceSource
        DetectionSource = $a.detectionSource
        CreatedDateTime = $a.createdDateTime
        AssignedTo      = $a.assignedTo
        Classification  = $a.classification
        Determination   = $a.determination
        Description     = $a.description
    }
}
$results = $results | Sort-Object @{Expression = 'Severity'; Descending = $true }, CreatedDateTime -Descending

# ── Output ────────────────────────────────────────────────────────────────────
if (-not $results -or @($results).Count -eq 0) {
    Write-Host "  No alerts found for the given window/filters." -ForegroundColor DarkGray
} else {
    $results | Select-Object Title, Severity, Status, ServiceSource, CreatedDateTime | Format-Table -AutoSize

    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "SecurityAlerts_$ts.csv"
    }
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

$highCount = @($results | Where-Object { $_.Severity -eq 'high' }).Count
Write-Host ""
Write-Host ("  {0} alert(s) — {1} high severity" -f @($results).Count, $highCount) -ForegroundColor $(if ($highCount -gt 0) { 'Red' } else { 'Cyan' })
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
