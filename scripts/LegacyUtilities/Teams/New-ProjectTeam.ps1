#Requires -Version 7.0
<#
.SYNOPSIS
    Bulk-create Teams with a standard set of channels from a CSV-driven template.

.DESCRIPTION
    Generalized replacement for an old script that hardcoded one company's fixed
    project-channel layout and owner UPN. Reads two CSVs — one row per Team to
    create, one row per channel to add to every Team created in this run — so
    the channel template is supplied by the caller instead of baked into the
    script. This is the common "spin up a project Team with our standard
    channel set" MSP pattern, without the original's customer-specific content.

    Defaults to a safe preview — pass -Apply to actually create Teams/channels.

    Runs entirely on Microsoft Graph (the MicrosoftTeams module is no longer used):
    per row it creates the Microsoft 365 group (POST /groups, keeping MailNickname,
    owner as owner and member), turns it into a Team with the standard template
    (POST /teams with group@odata.bind), waits for the async provisioning to finish,
    then adds the channels (POST /teams/{id}/channels). A new group can take a while
    to replicate, so 404s are retried.

    Sign-in goes through scripts\Startup\Connect-M365.ps1: delegated as the admin by
    default (device code / GDAP customer per load.config.ps1), app-only with -ClientId
    and -CertificateThumbprint or -AppOnly. A Graph session for the right tenant that
    already has the scopes is reused and left connected; only a session this script
    opened is disconnected. Delegated scopes: Group.ReadWrite.All, User.Read.All,
    Team.Create, Channel.Create (app-only: Group.ReadWrite.All, User.Read.All).

.PARAMETER TeamsCsvPath
    Path to a CSV with columns: TeamName, MailNickname, Owner (owner UPN),
    optional Visibility (Private/Public, default Private).

.PARAMETER ChannelsCsvPath
    Path to a CSV with columns: ChannelName, optional Description — applied to
    every Team created in this run.

.PARAMETER Apply
    Actually create the Teams and channels. Without this switch, the script only
    reports what it would create.

.PARAMETER OutputPath
    CSV report path. Defaults to `C:\Temp\NewProjectTeams_<timestamp>.csv`
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
    .\New-ProjectTeam.ps1 -TeamsCsvPath .\teams.csv -ChannelsCsvPath .\channels.csv

.EXAMPLE
    .\New-ProjectTeam.ps1 -TeamsCsvPath .\teams.csv -ChannelsCsvPath .\channels.csv -Apply

.NOTES
    Teams CSV example:
        TeamName,MailNickname,Owner,Visibility
        "Project 1001","project-1001","pm@contoso.com","Private"

    Channels CSV example:
        ChannelName,Description
        "Documents","Signed customer documents"
        "Internal",""

    Required module: Microsoft.Graph.Authentication
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $TeamsCsvPath,

    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $ChannelsCsvPath,

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
if (-not $OutputPath) { $OutputPath = Join-Path $outputDir "NewProjectTeams_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" }

$teamRows    = Import-Csv -Path $TeamsCsvPath
$channelRows = Import-Csv -Path $ChannelsCsvPath

# -notcontains: "-not $names -contains 'X'" negated the list first and never fired.
if (-not $teamRows -or $teamRows[0].PSObject.Properties.Name -notcontains 'TeamName') {
    Write-Error "Teams CSV must have at least a 'TeamName' column."
    exit 1
}
if (-not $channelRows -or $channelRows[0].PSObject.Properties.Name -notcontains 'ChannelName') {
    Write-Error "Channels CSV must have at least a 'ChannelName' column."
    exit 1
}

# ── Connection ────────────────────────────────────────────────────────────────
# Everything goes through Graph now; the MicrosoftTeams module (New-Team,
# New-TeamChannel) is no longer needed.
$graph = Connect-M365Graph -Scopes 'Group.ReadWrite.All', 'User.Read.All', 'Team.Create', 'Channel.Create' `
    -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

function Invoke-GraphWithRetry {
    <#
        A new group or team is not visible everywhere at once (Graph documents up
        to 15 minutes of replication); 404/409/429/5xx are retried with a pause.
    #>
    param([string] $Method, [string] $Uri, $Body, [int] $Attempts = 10, [int] $DelaySeconds = 15)
    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            $p = @{ Method = $Method; Uri = $Uri; OutputType = 'Hashtable'; ErrorAction = 'Stop'
                    ResponseHeadersVariable = 'hdr'; StatusCodeVariable = 'code' }
            if ($null -ne $Body) { $p['Body'] = ($Body | ConvertTo-Json -Depth 10); $p['ContentType'] = 'application/json' }
            $resp = Invoke-MgGraphRequest @p
            return [pscustomobject]@{ Body = $resp; Headers = $hdr; StatusCode = $code }
        } catch {
            $status = [int]($_.Exception.Response.StatusCode)
            if ($i -lt $Attempts -and ($status -in 404, 409, 429 -or $status -ge 500)) {
                Write-Host "     ... not ready yet ($status), retry $i/$($Attempts - 1) in $DelaySeconds s" -ForegroundColor DarkGray
                Start-Sleep -Seconds $DelaySeconds
                continue
            }
            throw
        }
    }
}

function Wait-TeamProvisioned {
    # POST /teams answers 202 with a Location of the teamsAsyncOperation; poll it.
    param([string] $GroupId, $Headers, [int] $TimeoutSeconds = 600)
    $location = $null
    if ($Headers) {
        foreach ($k in 'Location', 'Content-Location') {
            if ($Headers[$k]) { $location = @($Headers[$k])[0]; break }
        }
    }
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if ($location) {
            $uri = if ($location -match '^https://') { $location } else { "v1.0$location" }
            try {
                $op = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType Hashtable -ErrorAction Stop
                switch ($op['status']) {
                    'succeeded' { return }
                    'failed'    { throw "Team provisioning failed: $($op['error'] | ConvertTo-Json -Compress)" }
                }
            } catch {
                if ($_.Exception.Message -like 'Team provisioning failed*') { throw }
            }
        } else {
            try { $null = Invoke-MgGraphRequest -Method GET -Uri "v1.0/teams/$GroupId" -ErrorAction Stop; return } catch { }
        }
        Start-Sleep -Seconds 10
    }
    throw "Team for group $GroupId was not provisioned within $TimeoutSeconds seconds."
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   New-ProjectTeam" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Teams to create : $($teamRows.Count)"
Write-Host "  Channels each   : $($channelRows.Count)"
Write-Host ("  Mode            : {0}" -f $(if ($Apply) { 'Apply' } else { 'Preview only' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($row in $teamRows) {
    $teamName     = $row.TeamName
    $mailNickname = if ($row.PSObject.Properties.Name -contains 'MailNickname' -and $row.MailNickname) { $row.MailNickname } else { ($teamName -replace '[^A-Za-z0-9._-]', '').ToLower() }
    $owner        = $row.Owner
    $visibility   = if ($row.PSObject.Properties.Name -contains 'Visibility' -and $row.Visibility) { $row.Visibility } else { 'Private' }

    if (-not $teamName -or -not $owner) {
        Write-Host "  [SKIP] Row missing TeamName or Owner." -ForegroundColor Yellow
        continue
    }

    if (-not $Apply) {
        Write-Host "  [PREVIEW] $teamName (owner: $owner, $visibility) + $($channelRows.Count) channel(s)" -ForegroundColor DarkGray
        $results.Add([PSCustomObject]@{ TeamName = $teamName; Owner = $owner; ChannelsCreated = 0; Status = 'Preview' })
        continue
    }

    if (-not $PSCmdlet.ShouldProcess($teamName, "Create Team")) { continue }

    try {
        $ownerObj = Invoke-MgGraphRequest -Method GET -Uri "v1.0/users/$([uri]::EscapeDataString($owner))?`$select=id" -OutputType Hashtable -ErrorAction Stop
        $ownerRef = "https://graph.microsoft.com/v1.0/users/$($ownerObj['id'])"

        # 1. The Microsoft 365 group first, so the CSV's MailNickname is kept
        #    (POST /teams on its own derives the alias from the display name).
        $group = (Invoke-GraphWithRetry -Method POST -Uri 'v1.0/groups' -Attempts 1 -Body @{
            displayName          = $teamName
            mailNickname         = $mailNickname
            mailEnabled          = $true
            securityEnabled      = $false
            groupTypes           = @('Unified')
            visibility           = $visibility
            'owners@odata.bind'  = @($ownerRef)
            'members@odata.bind' = @($ownerRef)
        }).Body
        $groupId = $group['id']

        # 2. Team on that group, standard template. Retries the 404 a new group gives
        #    until it has replicated.
        $teamResp = Invoke-GraphWithRetry -Method POST -Uri 'v1.0/teams' -Body @{
            'template@odata.bind' = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
            'group@odata.bind'    = "https://graph.microsoft.com/v1.0/groups('$groupId')"
        }
        Wait-TeamProvisioned -GroupId $groupId -Headers $teamResp.Headers
        Write-Host "  [OK]   Created Team: $teamName" -ForegroundColor Green

        # 3. Channels
        $channelCount = 0
        foreach ($ch in $channelRows) {
            if (-not $ch.ChannelName -or $ch.ChannelName -eq 'General') { continue }
            try {
                $desc = if ($ch.PSObject.Properties.Name -contains 'Description') { $ch.Description } else { '' }
                $chBody = @{ displayName = $ch.ChannelName; membershipType = 'standard' }
                if ($desc) { $chBody['description'] = $desc }
                $null = Invoke-GraphWithRetry -Method POST -Uri "v1.0/teams/$groupId/channels" -Body $chBody -Attempts 4 -DelaySeconds 10
                $channelCount++
            } catch {
                Write-Host "     [WARN] Channel '$($ch.ChannelName)': $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }

        Write-Host "     Created $channelCount channel(s)" -ForegroundColor DarkGray
        $results.Add([PSCustomObject]@{ TeamName = $teamName; Owner = $owner; ChannelsCreated = $channelCount; Status = 'Created' })
    } catch {
        Write-Host "  [WARN] $teamName : $($_.Exception.Message)" -ForegroundColor Yellow
        $results.Add([PSCustomObject]@{ TeamName = $teamName; Owner = $owner; ChannelsCreated = 0; Status = "Error: $($_.Exception.Message)" })
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
if (-not $Apply) { Write-Host "  Re-run with -Apply to create these Teams." -ForegroundColor Yellow }
Write-Host ""

# ── Disconnect only if this script connected ──────────────────────────────────
Disconnect-M365Graph $graph
