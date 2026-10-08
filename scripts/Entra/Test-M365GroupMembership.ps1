#Requires -Version 7.0
<#
.SYNOPSIS
    Audit Microsoft 365 group owners and members via Microsoft Graph.

.DESCRIPTION
    Connects to Microsoft Graph and reports on Microsoft 365 Groups (including
    Teams-backed groups). For each group it lists:
      - Owners
      - Members
    Results are exported to CSV. Owners and members come from one paged Graph call
    per group each (with $select), not one Get-MgUser per member.

    Sign-in goes through scripts\Startup\Connect-M365.ps1: delegated as the admin by
    default (device code / GDAP customer per load.config.ps1), app-only with -ClientId
    and -CertificateThumbprint or -AppOnly. A fitting Graph session is reused and left
    connected; only a session this script opened is disconnected.
    Delegated scopes: Group.Read.All, Directory.Read.All.

.PARAMETER Group
    Display name or Object ID of a single group. If omitted, all M365 groups are audited.

.PARAMETER OutputPath
    Path to write the CSV report. Defaults to .\M365GroupMembership_<timestamp>.csv.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Defaults to the GDAP customer when authMode is GDAP.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint and -TenantId).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.EXAMPLE
    .\Test-M365GroupMembership.ps1

.EXAMPLE
    .\Test-M365GroupMembership.ps1 -Group "Team Finance"

.EXAMPLE
    .\Test-M365GroupMembership.ps1 -OutputPath "C:\Reports\groups.csv"
#>
[CmdletBinding()]
param(
    [string] $Group,
    [string] $OutputPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly
)

. (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection (reuses a fitting session) ─────────────────────────────────────
$graph = Connect-M365Graph -Scopes 'Group.Read.All', 'Directory.Read.All' -TenantId $TenantId `
    -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

# Owners/members of a group in pages, with only the properties we report.
function Get-GroupRelation([string] $GroupId, [string] $Relation) {
    $next = "v1.0/groups/$GroupId/$Relation`?`$select=id,displayName,userPrincipalName&`$top=999"
    while ($next) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType Hashtable -ErrorAction Stop
        $page['value']
        $next = $page['@odata.nextLink']
    }
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   M365 Group Membership Audit" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# ── Get groups ────────────────────────────────────────────────────────────────
if ($Group) {
    # Try by Object ID first, then by display name
    try {
        $groups = @(Get-MgGroup -GroupId $Group -ErrorAction Stop)
    } catch {
        $escaped = $Group -replace "'", "''"
        $groups = @(Get-MgGroup -Filter "displayName eq '$escaped' and groupTypes/any(c:c eq 'Unified')" -All -ErrorAction Stop)
    }
} else {
    Write-Host "  Retrieving M365 groups..." -ForegroundColor DarkGray
    # groupTypes/any(c:c eq 'Unified') = Microsoft 365 Groups only
    $groups = @(Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" -All)
}

Write-Host "  Auditing $($groups.Count) group(s)..." -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($grp in $groups) {
    foreach ($rel in @(@{ Name = 'owners'; Role = 'Owner' }, @{ Name = 'members'; Role = 'Member' })) {
        try {
            foreach ($obj in Get-GroupRelation -GroupId $grp.Id -Relation $rel.Name) {
                $results.Add([PSCustomObject]@{
                    GroupName   = $grp.DisplayName
                    GroupEmail  = $grp.Mail
                    GroupId     = $grp.Id
                    Role        = $rel.Role
                    ObjectType  = (([string]$obj['@odata.type']) -replace '^#microsoft\.graph\.', '')
                    DisplayName = $obj['displayName']
                    UPN         = $obj['userPrincipalName']
                })
            }
        } catch {
            Write-Host "  [WARN] $($rel.Role)s — $($grp.DisplayName) : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
if ($results.Count -eq 0) {
    Write-Host "  No groups or members found." -ForegroundColor DarkGray
} else {
    $results | Format-Table -AutoSize

    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "M365GroupMembership_$ts.csv"
    }

    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "  Audited $($groups.Count) group(s) — $($results.Count) entries written." -ForegroundColor Cyan
Write-Host ""

# ── Disconnect only if this script connected ──────────────────────────────────
Disconnect-M365Graph $graph
