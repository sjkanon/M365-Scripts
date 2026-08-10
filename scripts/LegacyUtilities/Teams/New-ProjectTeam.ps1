#Requires -Version 5.1
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

    Connects to Microsoft Graph and Microsoft Teams PowerShell automatically if
    no session is active; reuses an existing session if already connected.

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
    Entra ID tenant ID or domain. Optional if already connected.

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

    Required modules: MicrosoftTeams, Microsoft.Graph.Authentication
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
    [string] $TenantId
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
if (-not $OutputPath) { $OutputPath = Join-Path $outputDir "NewProjectTeams_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" }

$teamRows    = Import-Csv -Path $TeamsCsvPath
$channelRows = Import-Csv -Path $ChannelsCsvPath

if (-not $teamRows -or -not $teamRows[0].PSObject.Properties.Name -contains 'TeamName') {
    Write-Error "Teams CSV must have at least a 'TeamName' column."
    exit 1
}
if (-not $channelRows -or -not $channelRows[0].PSObject.Properties.Name -contains 'ChannelName') {
    Write-Error "Channels CSV must have at least a 'ChannelName' column."
    exit 1
}

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    $null = Get-MgContext -ErrorAction Stop
    if (-not (Get-MgContext)) { throw }
} catch {
    $connectParams = @{ Scopes = @('Group.ReadWrite.All'); NoWelcome = $true }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
}

$script:ConnectedTeamsHere = $false
try {
    $null = Get-CsTeamsClientConfiguration -ErrorAction Stop
} catch {
    Connect-MicrosoftTeams | Out-Null
    $script:ConnectedTeamsHere = $true
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
    $mailNickname = if ($row.PSObject.Properties.Name -contains 'MailNickname' -and $row.MailNickname) { $row.MailNickname } else { ($teamName -replace '\s', '').ToLower() }
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
        $team = New-Team -DisplayName $teamName -MailNickName $mailNickname -Owner $owner -Visibility $visibility -ErrorAction Stop
        $groupId = $team.GroupId
        Write-Host "  [OK]   Created Team: $teamName" -ForegroundColor Green

        $channelCount = 0
        foreach ($ch in $channelRows) {
            if (-not $ch.ChannelName -or $ch.ChannelName -eq 'General') { continue }
            try {
                $desc = if ($ch.PSObject.Properties.Name -contains 'Description') { $ch.Description } else { '' }
                New-TeamChannel -GroupId $groupId -DisplayName $ch.ChannelName -Description $desc -ErrorAction Stop | Out-Null
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

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedTeamsHere) { Disconnect-MicrosoftTeams | Out-Null }
if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
