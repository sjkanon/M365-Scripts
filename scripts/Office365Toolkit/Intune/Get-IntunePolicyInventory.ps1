#Requires -Version 7.0
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

    Sign-in: delegated (you sign in as the admin) by default, through
    scripts\Startup\Connect-M365.ps1 — device code and the GDAP customer come
    from load.config.ps1. App-only with -ClientId + -CertificateThumbprint, or
    -AppOnly (graph.appid.json). An existing Graph session with the right scopes
    is reused and left connected.

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\ (Windows) or ~/Downloads (macOS/Linux).

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
    .\Get-IntunePolicyInventory.ps1

.EXAMPLE
    .\Get-IntunePolicyInventory.ps1 -OutputPath C:\Reports\intune.csv

.EXAMPLE
    .\Get-IntunePolicyInventory.ps1 -TenantId contoso.onmicrosoft.com -AppOnly

.NOTES
    Capability inspired by intune-policy-get.ps1 from the retired
    directorcia/Office365 (CIAOPS) toolkit, which used the deprecated
    Microsoft.Graph.Intune ("AzureAD Intune PowerShell SDK") module and
    printed names only, with no assignment info or export.

    All five surfaces are read with Invoke-MgGraphRequest and follow
    @odata.nextLink. Settings Catalog (configurationPolicies) and Endpoint
    Security (intents) only exist on the Graph beta endpoint — the v1.0 SDK has
    no cmdlets for them, so earlier versions of this script silently missed
    both. App protection policies have no assignments on the base
    managedAppPolicy type, so their AssignmentCount stays empty.

    For comparing a customer tenant's Intune configuration against an MSP
    reference baseline (drift detection), see
    scripts/Intune/Compare-IntuneConfig.ps1 instead — that script does a full
    backup-based diff; this one is a quick point-in-time inventory of what
    currently exists.

    Delegated scopes: DeviceManagementConfiguration.Read.All,
    DeviceManagementApps.Read.All (app-only: the same as application permissions).
    Required module: Microsoft.Graph.Authentication
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
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
$graph = Connect-M365Graph -Scopes 'DeviceManagementConfiguration.Read.All', 'DeviceManagementApps.Read.All' `
    -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Intune Policy Inventory" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

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

function Add-PolicyResult {
    # $CountAssignments: the call expanded (or fetched) assignments, so an empty
    # list really means 0, not "unknown".
    param($Items, [string] $Type, [switch] $CountAssignments)
    foreach ($item in $Items) {
        $assignmentCount = if ($CountAssignments -and $null -ne $item['assignments']) { @($item['assignments']).Count } else { $null }
        $name = if ($item['displayName']) { $item['displayName'] } else { $item['name'] }   # Settings Catalog uses 'name'
        $script:results.Add([PSCustomObject]@{
            Type            = $Type
            DisplayName     = $name
            Id              = $item['id']
            ODataType       = $item['@odata.type']
            LastModified    = $item['lastModifiedDateTime']
            AssignmentCount = $assignmentCount
        })
    }
    Write-Host ("  {0,-28} : {1}" -f $Type, @($Items).Count) -ForegroundColor DarkGray
}

Write-Host "  Retrieving policies..." -ForegroundColor DarkGray
Write-Host ""

try {
    $compliance = Get-GraphCollection 'https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies?$expand=assignments'
    Add-PolicyResult -Items $compliance -Type 'Compliance Policy' -CountAssignments
} catch { Write-Host "  [WARN] Compliance policies: $($_.Exception.Message)" -ForegroundColor Yellow }

try {
    $configuration = Get-GraphCollection 'https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations?$expand=assignments'
    Add-PolicyResult -Items $configuration -Type 'Device Configuration' -CountAssignments
} catch { Write-Host "  [WARN] Device configuration profiles: $($_.Exception.Message)" -ForegroundColor Yellow }

try {
    # Beta only: Settings Catalog has no v1.0 endpoint.
    $settingsCatalog = Get-GraphCollection 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$expand=assignments'
    Add-PolicyResult -Items $settingsCatalog -Type 'Settings Catalog' -CountAssignments
} catch { Write-Host "  [WARN] Settings Catalog policies: $($_.Exception.Message)" -ForegroundColor Yellow }

try {
    $appProtection = Get-GraphCollection 'https://graph.microsoft.com/v1.0/deviceAppManagement/managedAppPolicies'
    Add-PolicyResult -Items $appProtection -Type 'App Protection Policy'
} catch { Write-Host "  [WARN] App protection policies: $($_.Exception.Message)" -ForegroundColor Yellow }

try {
    # Beta only: Endpoint Security intents have no v1.0 endpoint. Assignments are
    # read per intent from its documented /assignments navigation.
    $intents = Get-GraphCollection 'https://graph.microsoft.com/beta/deviceManagement/intents'
    foreach ($intent in $intents) {
        try {
            $intent['assignments'] = Get-GraphCollection "https://graph.microsoft.com/beta/deviceManagement/intents/$($intent['id'])/assignments"
        } catch { $intent['assignments'] = $null }
    }
    Add-PolicyResult -Items $intents -Type 'Endpoint Security (Intent)' -CountAssignments
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
Disconnect-M365Graph $graph
