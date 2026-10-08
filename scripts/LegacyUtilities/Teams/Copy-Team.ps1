#Requires -Version 7.0
<#
.SYNOPSIS
    Clone an existing Team (apps, tabs, settings, channels, and/or members) into
    a new Team.

.DESCRIPTION
    Wraps the Microsoft Graph team clone operation (POST /teams/{id}/clone),
    which is asynchronous — this script submits the clone, then polls the
    operation status until it reports success or failure. Modernized rewrite of
    an old sample-derived script; house-style dry-run/-Apply and reporting added.

    Sign-in goes through scripts\Startup\Connect-M365.ps1: delegated as the admin by
    default (device code / GDAP customer per load.config.ps1), app-only with -ClientId
    and -CertificateThumbprint or -AppOnly. A Graph session for the right tenant that
    already has the scopes is reused and left connected; only a session this script
    opened is disconnected. Delegated scopes: Group.ReadWrite.All, Directory.Read.All.

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
    Entra ID tenant ID or domain. Defaults to the GDAP customer when authMode is GDAP.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint and -TenantId).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.EXAMPLE
    # Preview
    .\Copy-Team.ps1 -SourceTeamId "Project Template" -NewTeamName "Project 1234"

.EXAMPLE
    .\Copy-Team.ps1 -SourceTeamId "Project Template" -NewTeamName "Project 1234" -Apply

.EXAMPLE
    # Only clone channels and settings, not members/apps/tabs
    .\Copy-Team.ps1 -SourceTeamId "Project Template" -NewTeamName "Project 1234" -PartsToClone Channels,Settings -Apply

.NOTES
    Cloning is a long-running Graph operation; this script polls the
    teamsAsyncOperation from the Location header every 15 seconds for up to 10
    minutes (falling back to a display-name lookup when no Location comes back).

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
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

if (-not $NewTeamDescription) { $NewTeamDescription = $NewTeamName }
if (-not $NewMailNickname) { $NewMailNickname = ($NewTeamName -replace '\s', '').ToLower() }

# ── Connection ────────────────────────────────────────────────────────────────
$graph = Connect-M365Graph -Scopes 'Group.ReadWrite.All', 'Directory.Read.All' -TenantId $TenantId `
    -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

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
    # Quotes escaped for OData; not $matches, that is an automatic variable.
    $escaped = $SourceTeamId -replace "'", "''"
    $found = @((Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$escaped' and resourceProvisioningOptions/Any(x:x eq 'Team')").value)
    if ($found.Count -eq 0) {
        Write-Error "Team '$SourceTeamId' not found by ID or display name."
        Disconnect-M365Graph $graph
        exit 1
    }
    if ($found.Count -gt 1) { Write-Host "  [WARN] Multiple teams named '$SourceTeamId' — using the first match." -ForegroundColor Yellow }
    $source = $found[0]
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
    Disconnect-M365Graph $graph
    return
}

if (-not $PSCmdlet.ShouldProcess($NewTeamName, "Clone team '$($source.displayName)'")) {
    Disconnect-M365Graph $graph
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
    # POST /teams/{id}/clone is asynchronous — it answers 202 with a Location of
    # the teamsAsyncOperation, which we poll below.
    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$($source.id)/clone" `
        -Body $body -ContentType 'application/json' -ResponseHeadersVariable cloneHeaders -ErrorAction Stop | Out-Null
} catch {
    Write-Host "  [ERROR] Clone request failed: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-M365Graph $graph
    exit 1
}
$operationUri = $null
if ($cloneHeaders -and $cloneHeaders['Location']) {
    $loc = @($cloneHeaders['Location'])[0]
    $operationUri = if ($loc -match '^https://') { $loc } else { "https://graph.microsoft.com/v1.0$loc" }
}

Write-Host "  Clone submitted. Waiting for the new Team to appear (this can take a few minutes)..." -ForegroundColor DarkGray

$deadline = (Get-Date).AddMinutes(10)
$newTeam  = $null
$failed   = $null
while ((Get-Date) -lt $deadline -and -not $newTeam -and -not $failed) {
    Start-Sleep -Seconds 15
    if ($operationUri) {
        # The operation says when it is done and which team it made; the old
        # display-name lookup also matched an existing group with the same name.
        try {
            $op = Invoke-Graph -Uri $operationUri
            if ($op.status -eq 'succeeded' -and $op.targetResourceId) {
                $newTeam = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/groups/$($op.targetResourceId)?`$select=id,displayName"
            } elseif ($op.status -eq 'failed') {
                $failed = $op.error | ConvertTo-Json -Compress
            }
        } catch { }
    } else {
        $escapedNew = $NewTeamName -replace "'", "''"
        $candidates = @((Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$escapedNew'").value)
        if ($candidates.Count) { $newTeam = $candidates[0] }
    }
    if (-not $newTeam -and -not $failed) { Write-Host "  ...still waiting" -ForegroundColor DarkGray }
}

Write-Host ""
if ($newTeam) {
    Write-Host "  [OK]   New Team created: $($newTeam.displayName) ($($newTeam.id))" -ForegroundColor Green
} elseif ($failed) {
    Write-Host "  [ERROR] Clone failed: $failed" -ForegroundColor Red
} else {
    Write-Host "  [WARN] Clone submitted but the new Team did not appear within 10 minutes." -ForegroundColor Yellow
    Write-Host "         It may still be provisioning — check Teams admin center shortly." -ForegroundColor Yellow
}
Write-Host ""

# ── Disconnect only if this script connected ──────────────────────────────────
Disconnect-M365Graph $graph
