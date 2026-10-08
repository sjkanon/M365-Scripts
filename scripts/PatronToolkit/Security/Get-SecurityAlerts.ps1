#Requires -Version 7.0
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
    Entra ID tenant ID or domain. Defaults to the GDAP customer (load.config.ps1) or
    your own tenant. Required for app-only sign-in.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint). Without it the
    script signs in delegated, as you.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

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

    Sign-in: Microsoft Graph through scripts\Startup\Connect-M365.ps1 - delegated by
    default (scope SecurityAlert.Read.All plus a Security Reader role; device code / GDAP
    customer per load.config.ps1), app-only with -ClientId/-CertificateThumbprint or
    -AppOnly (application permission SecurityAlert.Read.All). An existing fitting Graph
    session is reused and left connected.
#>
[CmdletBinding()]
param(
    [int] $Days = 30,
    [ValidateSet('informational', 'low', 'medium', 'high')]
    [string[]] $Severity,
    [ValidateSet('new', 'inProgress', 'resolved')]
    [string[]] $Status,
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
$graph = Connect-M365Graph -Scopes 'SecurityAlert.Read.All' -TenantId $TenantId `
    -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

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
    Disconnect-M365Graph $graph
    exit 1
}

Write-Host "  Retrieved $($alerts.Count) alert(s)." -ForegroundColor DarkGray
Write-Host ""

$severityRank = @{ high = 4; medium = 3; low = 2; informational = 1 }
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
# Sort on severity rank, not the string: alphabetically 'medium' > 'low' > 'informational' > 'high'.
$results = @($results | Sort-Object @{ Expression = { [int]$severityRank[[string]$_.Severity] }; Descending = $true },
                                    @{ Expression = 'CreatedDateTime'; Descending = $true })

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
Disconnect-M365Graph $graph
