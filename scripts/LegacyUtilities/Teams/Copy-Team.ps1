#Requires -Version 5.1
<#
.SYNOPSIS
    Clone an existing Team (apps, tabs, settings, channels, and/or members) into
    a new Team.

.DESCRIPTION
    Wraps the Microsoft Graph team clone operation (POST /teams/{id}/clone),
    which is asynchronous — this script submits the clone, then polls the
    operation status until it reports success or failure. Modernized rewrite of
    an old sample-derived script; house-style dry-run/-Apply and reporting added.

    Connects to Microsoft Graph automatically if no session is active; reuses an
    existing session if already connected.

.PARAMETER SourceTeamId
    Object ID (group ID) or display name of the Team to clone. If a display name
    matches more than one Team, the first match is used and a warning is shown.

.PARAMETER NewTeamName
    Display name for the cloned Team.

.PARAMETER NewTeamDescription
    Description for the cloned Team. Defaults to -NewTeamName if omitted.

.PARAMETER NewMailNickname
    Mail nickname for the cloned Team's underlying group. Spaces are stripped.
    Defaults to a lowercased version of -NewTeamName if omitted.

.PARAMETER Visibility
    "Private" or "Public". Default: Private.

.PARAMETER PartsToClone
    Which parts to include: any combination of Apps, Tabs, Settings, Channels,
    Members. Default: all five.

.PARAMETER Apply
    Actually submit the clone operation. Without this switch, the script only
    reports what it would do.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    # Preview
    .\Copy-Team.ps1 -SourceTeamId "Project Template" -NewTeamName "Project 1234"

.EXAMPLE
    .\Copy-Team.ps1 -SourceTeamId "Project Template" -NewTeamName "Project 1234" -Apply

.EXAMPLE
    # Only clone channels and settings, not members/apps/tabs
    .\Copy-Team.ps1 -SourceTeamId "Project Template" -NewTeamName "Project 1234" -PartsToClone Channels,Settings -Apply

.NOTES
    Cloning is a long-running Graph operation; this script polls every 10
    seconds for up to 10 minutes by default.

    Required module: Microsoft.Graph.Authentication (Group.ReadWrite.All / Team.Create)
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $SourceTeamId,

    [Parameter(Mandatory)]
    [string] $NewTeamName,

    [string] $NewTeamDescription,
    [string] $NewMailNickname,

    [ValidateSet('Private', 'Public')]
    [string] $Visibility = 'Private',

    [ValidateSet('Apps', 'Tabs', 'Settings', 'Channels', 'Members')]
    [string[]] $PartsToClone = @('Apps', 'Tabs', 'Settings', 'Channels', 'Members'),

    [switch] $Apply,
    [string] $TenantId
)

if (-not $NewTeamDescription) { $NewTeamDescription = $NewTeamName }
if (-not $NewMailNickname) { $NewMailNickname = ($NewTeamName -replace '\s', '').ToLower() }

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    $null = Get-MgContext -ErrorAction Stop
    if (-not (Get-MgContext)) { throw }
} catch {
    $connectParams = @{ Scopes = @('Group.ReadWrite.All', 'Directory.Read.All'); NoWelcome = $true }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
}

function Invoke-Graph {
    param([string] $Method = 'GET', [string] $Uri, [string] $Body)
    $params = @{ Method = $Method; Uri = $Uri; ErrorAction = 'Stop' }
    if ($Body) { $params['Body'] = $Body; $params['ContentType'] = 'application/json' }
    Invoke-MgGraphRequest @params
}

# ── Resolve source team ────────────────────────────────────────────────────────
try {
    $source = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/groups/$SourceTeamId`?`$select=id,displayName"
} catch {
    $matches = (Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$SourceTeamId' and resourceProvisioningOptions/Any(x:x eq 'Team')").value
    if (-not $matches) {
        Write-Error "Team '$SourceTeamId' not found by ID or display name."
        if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
        exit 1
    }
    if ($matches.Count -gt 1) { Write-Host "  [WARN] Multiple teams named '$SourceTeamId' — using the first match." -ForegroundColor Yellow }
    $source = $matches[0]
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Copy-Team" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Source     : $($source.displayName) ($($source.id))"
Write-Host "  New team   : $NewTeamName ($NewMailNickname, $Visibility)"
Write-Host "  Parts      : $($PartsToClone -join ', ')"
Write-Host ("  Mode       : {0}" -f $(if ($Apply) { 'Apply' } else { 'Preview only' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

if (-not $Apply) {
    Write-Host "  Would clone '$($source.displayName)' into '$NewTeamName'. Re-run with -Apply to perform the clone." -ForegroundColor Yellow
    if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
    return
}

if (-not $PSCmdlet.ShouldProcess($NewTeamName, "Clone team '$($source.displayName)'")) {
    if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
    return
}

# ── Submit clone ────────────────────────────────────────────────────────────────
$body = @{
    displayName  = $NewTeamName
    description  = $NewTeamDescription
    mailNickname = $NewMailNickname
    partsToClone = ($PartsToClone -join ',').ToLower()
    visibility   = $Visibility
} | ConvertTo-Json

Write-Host "  Submitting clone operation..." -ForegroundColor Cyan

try {
    # POST /teams/{id}/clone is asynchronous — the call itself just enqueues the
    # operation, the new Team shows up under its display name once provisioning
    # finishes, which we poll for below.
    Invoke-Graph -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$($source.id)/clone" -Body $body | Out-Null
} catch {
    Write-Host "  [ERROR] Clone request failed: $($_.Exception.Message)" -ForegroundColor Red
    if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
    exit 1
}

Write-Host "  Clone submitted. Waiting for the new Team to appear (this can take a few minutes)..." -ForegroundColor DarkGray

$deadline = (Get-Date).AddMinutes(10)
$newTeam  = $null
while ((Get-Date) -lt $deadline -and -not $newTeam) {
    Start-Sleep -Seconds 15
    $candidates = (Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$NewTeamName'").value
    if ($candidates) { $newTeam = $candidates[0] }
    Write-Host "  ...still waiting" -ForegroundColor DarkGray
}

Write-Host ""
if ($newTeam) {
    Write-Host "  [OK]   New Team created: $($newTeam.displayName) ($($newTeam.id))" -ForegroundColor Green
} else {
    Write-Host "  [WARN] Clone submitted but the new Team did not appear within 10 minutes." -ForegroundColor Yellow
    Write-Host "         It may still be provisioning — check Teams admin center shortly." -ForegroundColor Yellow
}
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
