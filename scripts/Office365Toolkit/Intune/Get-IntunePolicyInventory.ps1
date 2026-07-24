#Requires -Version 5.1
<#
.SYNOPSIS
    Inventory all Intune / Endpoint Manager policies in a tenant.

.DESCRIPTION
    Connects to Microsoft Graph and lists every policy across the main Intune
    policy surfaces: device compliance policies, device configuration profiles,
    Settings Catalog configuration policies, app protection policies, and
    Endpoint Security ("intents") policies. Useful as a quick tenant-wide
    inventory/checklist — what policies exist and how many assignments each
    has — without needing a full backup.

    Read-only. Results are exported to CSV.

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\ (Windows) or ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Get-IntunePolicyInventory.ps1

.EXAMPLE
    .\Get-IntunePolicyInventory.ps1 -OutputPath C:\Reports

.NOTES
    Capability inspired by intune-policy-get.ps1 from the retired
    directorcia/Office365 (CIAOPS) toolkit, which used the deprecated
    Microsoft.Graph.Intune ("AzureAD Intune PowerShell SDK") module and
    printed names only, with no assignment info or export. This rewrite uses
    the current Microsoft.Graph.DeviceManagement module and adds assignment
    counts + CSV export.

    For comparing a customer tenant's Intune configuration against an MSP
    reference baseline (drift detection), see
    scripts/Intune/Compare-IntuneConfig.ps1 instead — that script does a full
    backup-based diff; this one is a quick point-in-time inventory of what
    currently exists.

    Required scopes: DeviceManagementConfiguration.Read.All,
    DeviceManagementApps.Read.All
    Required module: Microsoft.Graph.DeviceManagement
#>
[CmdletBinding()]
param(
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
    $connectParams = @{
        Scopes = @('DeviceManagementConfiguration.Read.All', 'DeviceManagementApps.Read.All')
    }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Intune Policy Inventory" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

function Add-PolicyResult {
    param($Items, [string] $Type)
    foreach ($item in $Items) {
        $assignmentCount = $null
        if ($item.Assignments) { $assignmentCount = @($item.Assignments).Count }
        $script:results.Add([PSCustomObject]@{
            Type            = $Type
            DisplayName     = $item.DisplayName
            Id              = $item.Id
            LastModified    = $item.LastModifiedDateTime
            AssignmentCount = $assignmentCount
        })
    }
    Write-Host ("  {0,-28} : {1}" -f $Type, @($Items).Count) -ForegroundColor DarkGray
}

Write-Host "  Retrieving policies..." -ForegroundColor DarkGray
Write-Host ""

try {
    $compliance = Get-MgDeviceManagementDeviceCompliancePolicy -All -ExpandProperty Assignments -ErrorAction Stop
    Add-PolicyResult -Items $compliance -Type 'Compliance Policy'
} catch { Write-Host "  [WARN] Compliance policies: $($_.Exception.Message)" -ForegroundColor Yellow }

try {
    $configuration = Get-MgDeviceManagementDeviceConfiguration -All -ExpandProperty Assignments -ErrorAction Stop
    Add-PolicyResult -Items $configuration -Type 'Device Configuration'
} catch { Write-Host "  [WARN] Device configuration profiles: $($_.Exception.Message)" -ForegroundColor Yellow }

try {
    $settingsCatalog = Get-MgDeviceManagementConfigurationPolicy -All -ErrorAction Stop
    Add-PolicyResult -Items $settingsCatalog -Type 'Settings Catalog'
} catch { Write-Host "  [WARN] Settings Catalog policies: $($_.Exception.Message)" -ForegroundColor Yellow }

try {
    $appProtection = Get-MgDeviceAppManagementManagedAppPolicy -All -ErrorAction Stop
    Add-PolicyResult -Items $appProtection -Type 'App Protection Policy'
} catch { Write-Host "  [WARN] App protection policies: $($_.Exception.Message)" -ForegroundColor Yellow }

try {
    $intents = Get-MgDeviceManagementIntent -All -ErrorAction Stop
    Add-PolicyResult -Items $intents -Type 'Endpoint Security (Intent)'
} catch { Write-Host "  [WARN] Endpoint Security intents: $($_.Exception.Message)" -ForegroundColor Yellow }

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
if ($results.Count -eq 0) {
    Write-Host "  No policies found." -ForegroundColor DarkGray
} else {
    $results | Sort-Object Type, DisplayName | Format-Table Type, DisplayName, AssignmentCount, LastModified -AutoSize

    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "IntunePolicyInventory_$ts.csv"
    }
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "  $($results.Count) polic(y/ies) inventoried." -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ──────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
