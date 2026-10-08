#Requires -Version 7.0
<#
.SYNOPSIS
    Add or remove one or more members of a Microsoft 365 / security group via
    Microsoft Graph.

.DESCRIPTION
    Adds (or removes) a single member or a CSV/TXT list of members to/from a
    group, identified by object ID or display name. Defaults to a safe preview —
    pass -Apply to actually change membership.

    Sign-in goes through scripts\Startup\Connect-M365.ps1: delegated as the admin by
    default (device code / GDAP customer per load.config.ps1), app-only with -ClientId
    and -CertificateThumbprint or -AppOnly. A Graph session for the right tenant that
    already has the scopes is reused and left connected; only a session this script
    opened is disconnected.
    Delegated scopes: GroupMember.ReadWrite.All, Group.Read.All, User.Read.All.

.PARAMETER GroupId
    Object ID or exact display name of the target group.

.PARAMETER Member
    UPN or object ID of a single member to add/remove.

.PARAMETER CsvPath
    Path to a CSV with a UserPrincipalName/UPN/Mail column, or a plain TXT file
    with one UPN/object ID per line — for bulk membership changes.

.PARAMETER Action
    "Add" or "Remove". Default: Add.

.PARAMETER Apply
    Actually change membership. Without this switch, the script only reports
    what it would do.

.PARAMETER OutputPath
    CSV report path. Defaults to `C:\Temp\GroupMembershipChanges_<timestamp>.csv`
    (`~/Downloads` on Linux/macOS).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Defaults to the GDAP customer when authMode is GDAP.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint and -TenantId).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.EXAMPLE
    # Preview
    .\Add-M365GroupMember.ps1 -GroupId "Sales Team" -Member "j.doe@contoso.com"

.EXAMPLE
    .\Add-M365GroupMember.ps1 -GroupId "Sales Team" -Member "j.doe@contoso.com" -Apply

.EXAMPLE
    # Bulk add from CSV
    .\Add-M365GroupMember.ps1 -GroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -CsvPath .\newmembers.csv -Apply

.EXAMPLE
    # Bulk remove
    .\Add-M365GroupMember.ps1 -GroupId "Sales Team" -CsvPath .\leavers.csv -Action Remove -Apply

.NOTES
    Required modules: Microsoft.Graph.Groups, Microsoft.Graph.Users
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Single')]
param(
    [Parameter(Mandatory)]
    [string] $GroupId,

    [Parameter(ParameterSetName = 'Single', Mandatory)]
    [string] $Member,

    [Parameter(ParameterSetName = 'Csv', Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $CsvPath,

    [ValidateSet('Add', 'Remove')]
    [string] $Action = 'Add',

    [switch] $Apply,
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
if (-not $OutputPath) { $OutputPath = Join-Path $outputDir "GroupMembershipChanges_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" }

# ── Connection ────────────────────────────────────────────────────────────────
# User.Read.All: members are looked up with Get-MgUser, which the old scope set
# did not cover.
$graph = Connect-M365Graph -Scopes 'GroupMember.ReadWrite.All', 'Group.Read.All', 'User.Read.All' -TenantId $TenantId `
    -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

# ── Resolve group ──────────────────────────────────────────────────────────────
try {
    $group = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
} catch {
    # Quotes escaped for OData; an ambiguous name is an error, not "first match",
    # because writing members into the wrong group is worse than stopping.
    $escaped = $GroupId -replace "'", "''"
    $found = @(Get-MgGroup -Filter "displayName eq '$escaped'" -All -ErrorAction Stop)
    if ($found.Count -ne 1) {
        if ($found.Count -eq 0) { Write-Error "Group '$GroupId' not found by ID or display name." }
        else { Write-Error "Group name '$GroupId' matches $($found.Count) groups. Use the object ID." }
        Disconnect-M365Graph $graph
        exit 1
    }
    $group = $found[0]
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Add-M365GroupMember" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Group  : $($group.DisplayName) ($($group.Id))"
Write-Host "  Action : $Action"
Write-Host ("  Mode   : {0}" -f $(if ($Apply) { 'Apply' } else { 'Preview only' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

# ── Resolve members ────────────────────────────────────────────────────────────
$identities = [System.Collections.Generic.List[string]]::new()
if ($PSCmdlet.ParameterSetName -eq 'Csv') {
    $ext = [System.IO.Path]::GetExtension($CsvPath).ToLower()
    if ($ext -eq '.csv') {
        $raw = Import-Csv -Path $CsvPath
        $col = $raw[0].PSObject.Properties.Name |
            Where-Object { $_ -match 'userprincipalname|^upn$|^mail$' } | Select-Object -First 1
        if (-not $col) {
            Write-Error "CSV must have a UserPrincipalName, UPN, or Mail column. Found: $($raw[0].PSObject.Properties.Name -join ', ')"
            Disconnect-M365Graph $graph
            exit 1
        }
        $raw.$col | Where-Object { $_ } | ForEach-Object { $identities.Add($_.Trim()) }
    } else {
        Get-Content -Path $CsvPath | Where-Object { $_.Trim() -and $_ -notmatch '^\s*#' } | ForEach-Object { $identities.Add($_.Trim()) }
    }
} else {
    $identities.Add($Member)
}

Write-Host "  Member(s): $($identities.Count)" -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($identity in $identities) {
    try {
        $user = Get-MgUser -UserId $identity -ErrorAction Stop
    } catch {
        Write-Host "  [WARN] $identity : not found" -ForegroundColor Yellow
        $results.Add([PSCustomObject]@{ Member = $identity; Action = $Action; Status = 'NotFound' })
        continue
    }

    if (-not $Apply) {
        Write-Host "  [PREVIEW] $Action $($user.UserPrincipalName)" -ForegroundColor DarkGray
        $results.Add([PSCustomObject]@{ Member = $user.UserPrincipalName; Action = $Action; Status = 'Preview' })
        continue
    }

    if (-not $PSCmdlet.ShouldProcess($user.UserPrincipalName, "$Action group membership")) { continue }

    try {
        if ($Action -eq 'Add') {
            New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction Stop
        } else {
            Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction Stop
        }
        Write-Host "  [OK]   $Action $($user.UserPrincipalName)" -ForegroundColor Green
        $results.Add([PSCustomObject]@{ Member = $user.UserPrincipalName; Action = $Action; Status = 'Success' })
    } catch {
        Write-Host "  [WARN] $($user.UserPrincipalName) : $($_.Exception.Message)" -ForegroundColor Yellow
        $results.Add([PSCustomObject]@{ Member = $user.UserPrincipalName; Action = $Action; Status = "Error: $($_.Exception.Message)" })
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
if (-not $Apply) { Write-Host "  Re-run with -Apply to make these changes." -ForegroundColor Yellow }
Write-Host ""

# ── Disconnect only if this script connected ──────────────────────────────────
Disconnect-M365Graph $graph
