#Requires -Version 5.1
<#
.SYNOPSIS
    Audit Microsoft 365 group owners and members via Microsoft Graph.

.DESCRIPTION
    Connects to Microsoft Graph and reports on Microsoft 365 Groups (including
    Teams-backed groups). For each group it lists:
      - Owners
      - Members
    Results are exported to CSV.

.PARAMETER Group
    Display name or Object ID of a single group. If omitted, all M365 groups are audited.

.PARAMETER OutputPath
    Path to write the CSV report. Defaults to .\M365GroupMembership_<timestamp>.csv.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

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
        Scopes = @('Group.Read.All', 'Directory.Read.All')
    }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
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
        $groups = @(Get-MgGroup -Filter "displayName eq '$Group' and groupTypes/any(c:c eq 'Unified')" -ErrorAction Stop)
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
    # Owners
    try {
        $owners = Get-MgGroupOwner -GroupId $grp.Id -All -ErrorAction SilentlyContinue
        foreach ($owner in $owners) {
            $ownerDetail = Get-MgUser -UserId $owner.Id -ErrorAction SilentlyContinue
            $results.Add([PSCustomObject]@{
                GroupName  = $grp.DisplayName
                GroupEmail = $grp.Mail
                GroupId    = $grp.Id
                Role       = 'Owner'
                DisplayName = $ownerDetail.DisplayName
                UPN        = $ownerDetail.UserPrincipalName
            })
        }
    } catch {
        Write-Host "  [WARN] Owners — $($grp.DisplayName) : $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Members
    try {
        $members = Get-MgGroupMember -GroupId $grp.Id -All -ErrorAction SilentlyContinue
        foreach ($member in $members) {
            $memberDetail = Get-MgUser -UserId $member.Id -ErrorAction SilentlyContinue
            $results.Add([PSCustomObject]@{
                GroupName   = $grp.DisplayName
                GroupEmail  = $grp.Mail
                GroupId     = $grp.Id
                Role        = 'Member'
                DisplayName = $memberDetail.DisplayName
                UPN         = $memberDetail.UserPrincipalName
            })
        }
    } catch {
        Write-Host "  [WARN] Members — $($grp.DisplayName) : $($_.Exception.Message)" -ForegroundColor Yellow
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

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
