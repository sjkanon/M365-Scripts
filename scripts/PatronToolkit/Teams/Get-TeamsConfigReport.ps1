#Requires -Version 5.1
<#
.SYNOPSIS
    Report Microsoft Teams tenant governance settings and team inventory.

.DESCRIPTION
    Connects to Microsoft Teams PowerShell and reports the tenant-wide policies that
    matter most for a Teams governance/security review:
      - External access (federation) configuration
      - Guest access on/off and the global meeting/messaging policies
      - Global meeting policy (anonymous join, recording, external presenters)
      - Global messaging policy
      - Teams app setup policy (sideloading)
    Also lists every team in the tenant with visibility, archived status, and owner/member
    counts. Read-only. Results are exported to CSV.

.PARAMETER IncludeTeamsInventory
    Also list every team with member/owner counts. Default: on. Use
    -IncludeTeamsInventory:$false to skip (faster on tenants with many teams).

.PARAMETER OutputPath
    Folder for the CSV report(s). Defaults to C:\Temp (Windows) / ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Get-TeamsConfigReport.ps1

.EXAMPLE
    .\Get-TeamsConfigReport.ps1 -IncludeTeamsInventory:$false

.NOTES
    Inspired by the capability list of the retired directorcia/patron toolkit
    (o365-csteams-get.ps1, o365-tms-get.ps1, graph-teams-get.ps1), consolidated and
    rewritten from scratch — the originals used the retired SkypeOnlineConnector module
    and compared settings against a hardcoded external "best practices" JSON endpoint
    that no longer belongs to this repo.

    Required module: MicrosoftTeams
#>
[CmdletBinding()]
param(
    [switch] $IncludeTeamsInventory = $true,
    [string] $OutputPath,
    [string] $TenantId
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($OutputPath) { $OutputPath } elseif ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    $null = Get-CsTenant -ErrorAction Stop
} catch {
    $connectParams = @{}
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MicrosoftTeams @connectParams -ErrorAction Stop | Out-Null
    $script:ConnectedHere = $true
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Teams Configuration Report" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

$findings = [System.Collections.Generic.List[PSObject]]::new()
function Add-Finding {
    param([string] $Category, [string] $Setting, [string] $Value, [ValidateSet('Pass', 'Warn', 'Fail', 'Info')] [string] $Flag = 'Info')
    $script:findings.Add([PSCustomObject]@{ Category = $Category; Setting = $Setting; Value = $Value; Flag = $Flag })
    $color = switch ($Flag) { 'Pass' { 'Green' }; 'Warn' { 'Yellow' }; 'Fail' { 'Red' }; default { 'DarkGray' } }
    Write-Host ("  [{0,-4}] {1}: {2}" -f $Flag, $Setting, $Value) -ForegroundColor $color
}

# ── External access / federation ────────────────────────────────────────────────
try {
    $fed = Get-CsTenantFederationConfiguration -ErrorAction Stop
    Add-Finding 'External Access' 'AllowFederatedUsers' $fed.AllowFederatedUsers $(if ($fed.AllowFederatedUsers) { 'Info' } else { 'Pass' })
    Add-Finding 'External Access' 'AllowTeamsConsumer' $fed.AllowTeamsConsumer $(if ($fed.AllowTeamsConsumer) { 'Warn' } else { 'Pass' })
    Add-Finding 'External Access' 'AllowPublicUsers' $fed.AllowPublicUsers $(if ($fed.AllowPublicUsers) { 'Warn' } else { 'Pass' })
} catch { Add-Finding 'External Access' 'Federation config' "Not available: $($_.Exception.Message)" 'Info' }

# ── Guest access ──────────────────────────────────────────────────────────────
try {
    $guest = Get-CsTeamsGuestMeetingConfiguration -ErrorAction Stop
    Add-Finding 'Guest Access' 'AllowAnonymousUsersToJoinMeeting' $guest.AllowAnonymousUsersToJoinMeeting $(if ($guest.AllowAnonymousUsersToJoinMeeting) { 'Warn' } else { 'Pass' })
} catch { Add-Finding 'Guest Access' 'Guest meeting config' "Not available: $($_.Exception.Message)" 'Info' }

# ── Meeting policy (Global) ──────────────────────────────────────────────────
try {
    $mp = Get-CsTeamsMeetingPolicy -Identity Global -ErrorAction Stop
    Add-Finding 'Meeting Policy (Global)' 'AllowAnonymousUsersToJoinMeeting' $mp.AllowAnonymousUsersToJoinMeeting $(if ($mp.AllowAnonymousUsersToJoinMeeting) { 'Warn' } else { 'Pass' })
    Add-Finding 'Meeting Policy (Global)' 'AllowExternalParticipantGiveRequestControl' $mp.AllowExternalParticipantGiveRequestControl 'Info'
    Add-Finding 'Meeting Policy (Global)' 'AllowCloudRecording' $mp.AllowCloudRecording 'Info'
    Add-Finding 'Meeting Policy (Global)' 'DesignatedPresenterRoleMode' $mp.DesignatedPresenterRoleMode $(if ($mp.DesignatedPresenterRoleMode -eq 'EveryoneUserOverride') { 'Warn' } else { 'Pass' })
} catch { Add-Finding 'Meeting Policy (Global)' 'Policy' "Not available: $($_.Exception.Message)" 'Info' }

# ── Messaging policy (Global) ────────────────────────────────────────────────
try {
    $msgp = Get-CsTeamsMessagingPolicy -Identity Global -ErrorAction Stop
    Add-Finding 'Messaging Policy (Global)' 'AllowUserDeleteMessage' $msgp.AllowUserDeleteMessage 'Info'
    Add-Finding 'Messaging Policy (Global)' 'AllowUserChat' $msgp.AllowUserChat 'Info'
} catch { Add-Finding 'Messaging Policy (Global)' 'Policy' "Not available: $($_.Exception.Message)" 'Info' }

# ── App setup policy (sideloading) ──────────────────────────────────────────────
try {
    $appSetup = Get-CsTeamsAppSetupPolicy -Identity Global -ErrorAction Stop
    Add-Finding 'App Setup Policy (Global)' 'AllowSideloading' $appSetup.AllowSideloading $(if ($appSetup.AllowSideloading) { 'Warn' } else { 'Pass' })
} catch { Add-Finding 'App Setup Policy (Global)' 'Policy' "Not available: $($_.Exception.Message)" 'Info' }

# ── Teams inventory ───────────────────────────────────────────────────────────
$teamsList = [System.Collections.Generic.List[PSObject]]::new()
if ($IncludeTeamsInventory) {
    Write-Host ""
    Write-Host "  Retrieving team inventory..." -ForegroundColor DarkGray
    $teams = @(Get-Team -ErrorAction Stop)
    Write-Host "  Found $($teams.Count) team(s)." -ForegroundColor DarkGray
    foreach ($t in $teams) {
        try {
            $users = @(Get-TeamUser -GroupId $t.GroupId -ErrorAction Stop)
            $ownerCount = @($users | Where-Object { $_.Role -eq 'Owner' }).Count
            $memberCount = $users.Count
        } catch {
            $ownerCount = $null; $memberCount = $null
        }
        $teamsList.Add([PSCustomObject]@{
            DisplayName  = $t.DisplayName
            GroupId      = $t.GroupId
            Visibility   = $t.Visibility
            Archived     = $t.Archived
            OwnerCount   = $ownerCount
            MemberCount  = $memberCount
        })
        if ($ownerCount -eq 0) {
            Write-Host "  [WARN] '$($t.DisplayName)' has no owners" -ForegroundColor Yellow
        }
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$configPath = Join-Path $outputDir "TeamsConfigReport_$ts.csv"
$findings | Export-Csv -Path $configPath -NoTypeInformation -Encoding UTF8
Write-Host "  Config report saved: $configPath" -ForegroundColor Green

if ($IncludeTeamsInventory -and $teamsList.Count -gt 0) {
    $teamsPath = Join-Path $outputDir "TeamsInventory_$ts.csv"
    $teamsList | Export-Csv -Path $teamsPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Teams inventory saved: $teamsPath" -ForegroundColor Green
}

$warnCount = @($findings | Where-Object { $_.Flag -eq 'Warn' }).Count
Write-Host ""
Write-Host ("  {0} finding(s) — {1} warning(s)" -f $findings.Count, $warnCount) -ForegroundColor $(if ($warnCount -gt 0) { 'Yellow' } else { 'Green' })
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue }
