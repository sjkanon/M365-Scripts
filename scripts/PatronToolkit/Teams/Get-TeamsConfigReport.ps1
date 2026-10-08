#Requires -Version 7.0
<#
.SYNOPSIS
    Report Microsoft Teams tenant governance settings and team inventory.

.DESCRIPTION
    Reports the tenant-wide Teams policies that matter most for a Teams
    governance/security review, from Microsoft Teams PowerShell:
      - External access (federation) configuration
      - Guest access on/off (client configuration) and anonymous meeting join
        (meeting configuration)
      - Global meeting policy (anonymous join, recording, external presenters)
      - Global messaging policy
      - Teams app setup policy (sideloading)
    Also lists every team in the tenant with visibility, archived status, and owner/member
    counts, from Microsoft Graph. Read-only. Results are exported to CSV.

.PARAMETER IncludeTeamsInventory
    Also list every team with member/owner counts (Graph). Default: on. Use
    -IncludeTeamsInventory:$false to skip (faster on tenants with many teams, and no
    Graph sign-in).

.PARAMETER OutputPath
    Folder for the CSV report(s). Defaults to C:\Temp (Windows) / ~/Downloads (macOS/Linux).

.PARAMETER TenantId
    Tenant ID or domain. Defaults to the GDAP customer (load.config.ps1) or your own
    tenant. Required for app-only sign-in.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint), used for both
    Teams and Graph. Without it the script signs in delegated, as you.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.EXAMPLE
    .\Get-TeamsConfigReport.ps1

.EXAMPLE
    .\Get-TeamsConfigReport.ps1 -IncludeTeamsInventory:$false

.EXAMPLE
    .\Get-TeamsConfigReport.ps1 -TenantId contoso.onmicrosoft.com -AppOnly

.NOTES
    Inspired by the capability list of the retired directorcia/patron toolkit
    (o365-csteams-get.ps1, o365-tms-get.ps1, graph-teams-get.ps1), consolidated and
    rewritten from scratch — the originals used the retired SkypeOnlineConnector module
    and compared settings against a hardcoded external "best practices" JSON endpoint
    that no longer belongs to this repo.

    Sign-in, through scripts\Startup\Connect-M365.ps1 - delegated by default (device
    code / GDAP customer per load.config.ps1), app-only with
    -ClientId/-CertificateThumbprint or -AppOnly:
      Teams PowerShell - the Cs* tenant policies; Microsoft Graph has no API for them.
                         Needs a Teams Administrator (or Global Reader) role.
      Microsoft Graph  - the team inventory: /groups filtered on Team provisioning,
                         /teams/{id} (isArchived) and /teams/{id}/members (roles).
                         Scopes Group.Read.All, TeamMember.Read.All,
                         TeamSettings.Read.All (delegated) or the same application
                         permissions (app-only).

    Required modules: MicrosoftTeams, Microsoft.Graph.Authentication
#>
[CmdletBinding()]
param(
    [switch] $IncludeTeamsInventory = $true,
    [string] $OutputPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($OutputPath) { $OutputPath } elseif ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
# Teams PowerShell only for the Cs* policies, which Graph has no API for.
$teamsConn = Connect-M365Teams -TenantId $TenantId -ClientId $ClientId `
    -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

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
# Guest access on/off lives in the client configuration and anonymous meeting join in
# the meeting configuration; the guest *meeting* configuration the earlier version read
# has no AllowAnonymousUsersToJoinMeeting property, so it always reported an empty value.
try {
    $client = Get-CsTeamsClientConfiguration -Identity Global -ErrorAction Stop
    Add-Finding 'Guest Access' 'AllowGuestUser' $client.AllowGuestUser 'Info'
} catch { Add-Finding 'Guest Access' 'Client config' "Not available: $($_.Exception.Message)" 'Info' }
try {
    $meetingConfig = Get-CsTeamsMeetingConfiguration -Identity Global -ErrorAction Stop
    Add-Finding 'Guest Access' 'DisableAnonymousJoin' $meetingConfig.DisableAnonymousJoin $(if ($meetingConfig.DisableAnonymousJoin) { 'Pass' } else { 'Warn' })
} catch { Add-Finding 'Guest Access' 'Meeting config' "Not available: $($_.Exception.Message)" 'Info' }

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

# ── Teams inventory (Graph) ───────────────────────────────────────────────────
function Get-GraphPaged {
    param([string] $Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    while ($Uri) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType Hashtable -ErrorAction Stop
        foreach ($v in @($resp.value)) { if ($null -ne $v) { $items.Add($v) } }
        $Uri = $resp.'@odata.nextLink'
    }
    $items
}

$graph = $null
$teamsList = [System.Collections.Generic.List[PSObject]]::new()
if ($IncludeTeamsInventory) {
    Write-Host ""
    Write-Host "  Retrieving team inventory..." -ForegroundColor DarkGray
    try {
        $graph = Connect-M365Graph -Scopes 'Group.Read.All', 'TeamMember.Read.All', 'TeamSettings.Read.All' -TenantId $TenantId `
            -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly
        $teams = @(Get-GraphPaged -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&`$select=id,displayName,visibility&`$top=999")
        Write-Host "  Found $($teams.Count) team(s)." -ForegroundColor DarkGray
        foreach ($t in $teams) {
            $archived = $null
            try {
                $archived = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$($t.id)?`$select=isArchived" -OutputType Hashtable -ErrorAction Stop).isArchived
            } catch { }
            try {
                $members = @(Get-GraphPaged -Uri "https://graph.microsoft.com/v1.0/teams/$($t.id)/members")
                $ownerCount = @($members | Where-Object { @($_.roles) -contains 'owner' }).Count
                $memberCount = $members.Count
            } catch {
                $ownerCount = $null; $memberCount = $null
            }
            $teamsList.Add([PSCustomObject]@{
                DisplayName  = $t.displayName
                GroupId      = $t.id
                Visibility   = $t.visibility
                Archived     = $archived
                OwnerCount   = $ownerCount
                MemberCount  = $memberCount
            })
            if ($ownerCount -eq 0) {
                Write-Host "  [WARN] '$($t.displayName)' has no owners" -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Host "  [WARN] Could not retrieve the team inventory via Graph: $($_.Exception.Message)" -ForegroundColor Yellow
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

# ── Disconnect what we connected ──────────────────────────────────────────────
Disconnect-M365Graph $graph
Disconnect-M365Teams $teamsConn
