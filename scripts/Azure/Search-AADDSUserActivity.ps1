#Requires -Version 5.1
#Requires -Modules Az.Accounts, Az.OperationalInsights
<#
.SYNOPSIS
    Search all Azure AD Domain Services audit tables in Log Analytics for a single user in one query.

.DESCRIPTION
    Runs a union query across the AAD DS diagnostic tables (account management, account logon,
    logon/logoff, directory service access) so you don't have to guess which table an event landed
    in. Matches the given username against every column in each table.

    Connects with Connect-AzAccount if no Az session is active, and resolves the target Log
    Analytics workspace by name or ID (or auto-picks it if the subscription only has one).

.PARAMETER Username
    Username, SamAccountName, or UPN to search for. Matched with `has` across every column, so a
    partial account name is enough — no need to know which column (SamAccountName, ClientUserName,
    ...) it will show up in.

.PARAMETER WorkspaceId
    Log Analytics workspace ID (the `CustomerId` GUID). Takes precedence over -WorkspaceName.

.PARAMETER WorkspaceName
    Log Analytics workspace name to resolve via Get-AzOperationalInsightsWorkspace. Required together
    with -ResourceGroupName if more than one workspace shares that name across resource groups.

.PARAMETER ResourceGroupName
    Resource group of the workspace, used to disambiguate -WorkspaceName.

.PARAMETER HoursBack
    How many hours to look back from now. Ignored if -StartTime is supplied. Default: 2.

.PARAMETER StartTime
    Explicit start of the search window (local time). Overrides -HoursBack.

.PARAMETER EndTime
    Explicit end of the search window (local time). Default: now.

.PARAMETER Table
    AAD DS tables to include in the union. Default covers the four standard AAD DS diagnostic
    categories. Override to add or narrow down tables (e.g. AADDomainServicesDNSAuditsGeneral).

.PARAMETER MaxRows
    Maximum number of rows to return, most recent first. Default: 5000.

.PARAMETER ExportPath
    Folder where the CSV report is saved. Default: C:\Temp.

.EXAMPLE
    .\Search-AADDSUserActivity.ps1 -Username "jdoe" -WorkspaceId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

.EXAMPLE
    # Resolve workspace by name, look back 24 hours
    .\Search-AADDSUserActivity.ps1 -Username "jdoe" -WorkspaceName "log-aadds-prod" -HoursBack 24

.EXAMPLE
    # Explicit time window
    .\Search-AADDSUserActivity.ps1 -Username "jdoe" -WorkspaceId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -StartTime "2026-07-25 00:00" -EndTime "2026-07-27 00:00"
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]   $Username,

    [string]   $WorkspaceId,
    [string]   $WorkspaceName,
    [string]   $ResourceGroupName,

    [int]      $HoursBack = 2,
    [datetime] $StartTime,
    [datetime] $EndTime = (Get-Date),

    [string[]] $Table = @(
        'AADDomainServicesAccountManagement',
        'AADDomainServicesAccountLogon',
        'AADDomainServicesLogonLogoff',
        'AADDomainServicesDirectoryServiceAccess'
    ),

    [int]      $MaxRows = 5000,
    [string]   $ExportPath = 'C:\Temp'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Helpers ────────────────────────────────────────────────────────────────────
function Write-Status {
    param([string]$Msg, [string]$Level = 'INFO')
    $color = switch ($Level) {
        'OK'   { 'Green'  } 'WARN' { 'Yellow' }
        'FAIL' { 'Red'    } 'HEAD' { 'Cyan'   }
        default{ 'Gray'   }
    }
    Write-Host "  [$Level] $Msg" -ForegroundColor $color
}

# ── Banner ─────────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  ════════════════════════════════════════════' -ForegroundColor Cyan
Write-Host '   Search-AADDSUserActivity' -ForegroundColor Cyan
Write-Host '  ════════════════════════════════════════════' -ForegroundColor Cyan
Write-Host ''

# ── Az connection ──────────────────────────────────────────────────────────────
if (-not (Get-AzContext)) {
    Write-Status 'No active Az session — connecting...' 'WARN'
    Connect-AzAccount | Out-Null
}

# ── Resolve workspace ──────────────────────────────────────────────────────────
if (-not $WorkspaceId) {
    Write-Status 'Resolving Log Analytics workspace...' 'INFO'
    $wsParams = @{}
    if ($WorkspaceName)     { $wsParams['Name']              = $WorkspaceName }
    if ($ResourceGroupName) { $wsParams['ResourceGroupName']  = $ResourceGroupName }

    $workspaces = @(Get-AzOperationalInsightsWorkspace @wsParams)

    if ($workspaces.Count -eq 0) {
        throw 'No Log Analytics workspace found. Specify -WorkspaceId, or -WorkspaceName (+ -ResourceGroupName if ambiguous).'
    } elseif ($workspaces.Count -gt 1) {
        $workspaces | Select-Object Name, ResourceGroupName, CustomerId | Format-Table -AutoSize | Out-String | Write-Host
        throw 'Multiple workspaces match. Narrow down with -WorkspaceName + -ResourceGroupName, or pass -WorkspaceId directly.'
    }

    $WorkspaceId = $workspaces[0].CustomerId
    Write-Status "Using workspace '$($workspaces[0].Name)' ($WorkspaceId)" 'OK'
}

# ── Build query ────────────────────────────────────────────────────────────────
$rangeStart  = if ($PSBoundParameters.ContainsKey('StartTime')) { $StartTime } else { $EndTime.AddHours(-$HoursBack) }
$startUtcStr = $rangeStart.ToUniversalTime().ToString('o')
$endUtcStr   = $EndTime.ToUniversalTime().ToString('o')

# Escape for safe interpolation into the KQL string literal below.
$escapedUsername = $Username -replace '\\', '\\\\' -replace '"', '\"'
$tableList        = $Table -join ",`n    "

$query = @"
union
    $tableList
| where TimeGenerated between (datetime($startUtcStr) .. datetime($endUtcStr))
| where * has "$escapedUsername"
| project TimeGenerated, Type, OperationName, SamAccountName, ClientUserName, Workstation, IpAddress, ResultDescription, FailureCode
| order by TimeGenerated desc
| take $MaxRows
"@

Write-Status "Searching $($Table.Count) table(s) for '$Username' between $($rangeStart.ToString('yyyy-MM-dd HH:mm')) and $($EndTime.ToString('yyyy-MM-dd HH:mm'))..." 'HEAD'

# ── Run query ──────────────────────────────────────────────────────────────────
try {
    $result = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkspaceId -Query $query
} catch {
    throw "Query failed: $_"
}

$rows = @($result.Results)

if ($rows.Count -eq 0) {
    Write-Status 'No matching events found in the given time window.' 'WARN'
    exit 0
}

Write-Status "Found $($rows.Count) matching event(s)." 'OK'
Write-Host ''

# ── Export CSV ─────────────────────────────────────────────────────────────────
if (-not (Test-Path $ExportPath)) { New-Item -ItemType Directory -Path $ExportPath | Out-Null }
$ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvFile = Join-Path $ExportPath "AADDSUserActivity_${Username}_$ts.csv"
$rows | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8

$rows | Format-Table -AutoSize TimeGenerated, Type, OperationName, SamAccountName, IpAddress, ResultDescription

Write-Host ''
Write-Status "CSV exported to: $csvFile" 'OK'
if ($rows.Count -eq $MaxRows) {
    Write-Status "Result set was capped at -MaxRows ($MaxRows) — narrow the time window or raise -MaxRows if events may be missing." 'WARN'
}
Write-Host ''
