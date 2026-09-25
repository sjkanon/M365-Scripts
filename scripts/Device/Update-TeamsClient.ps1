#Requires -Version 5.1
<#
.SYNOPSIS
    Update the new Microsoft Teams client, including the Teams Meeting Add-in for
    Outlook, only when Microsoft publishes a newer build. Supports -WhatIf.

.DESCRIPTION
    Endpoint/AVD script that keeps new Teams current:

      1. Preflight  - inventory of every place Teams can live on the device: the
                      MSTeams AppX package (per user and provisioned), classic Teams
                      (machine-wide installer and per-profile installs), the meeting
                      add-in, whether Outlook itself has it registered, any
                      running Teams/Outlook process, and - on a session host - what
                      Teams logged about the media optimization.
      2. Check      - ask the Teams client config service which build is current for
                      this architecture and compare it with what is installed. Up to
                      date and nothing missing? Nothing happens at all.
      3. AVD        - only with -AvdOptimizations: the IsWVDEnvironment media flag
                      and the Remote Desktop WebRTC Redirector Service. With
                      -RemoveWebRtcRedirector instead: take that redirector away.
      4. Classic    - only with -RemoveClassicTeams: uninstall the classic Teams
                      machine-wide installer and clear the per-profile installs.
      5. Download   - fetch teamsbootstrapper.exe and verify its Microsoft signature
                      BEFORE anything is uninstalled, so a failed download can never
                      leave the device without a Teams client.
      6. Uninstall  - the Teams Meeting Add-in MSI, the MSTeams AppX package for all
                      users, and the provisioned package.
      7. Install    - provision new Teams for all users (teamsbootstrapper.exe -p).
      8. Add-in     - install the Teams Meeting Add-in MSI shipped inside the new
                      Teams package (ALLUSERS=1), after clearing every other copy of
                      it - but only once that MSI is in hand and can actually go in.
      9. Verify     - re-check the add-in registration (machine-wide *and* whether
                      Outlook sees it, per signed-in user), the provisioned package,
                      the classic removal and, where applicable, the AVD components.

    Each part is only done when it is actually needed. A current client with a missing
    add-in installs just the add-in; a current client on an AVD host with the WebRTC
    redirector missing installs just that.

    AVD / VDI
    ---------
    -AvdOptimizations adds the two things a session host needs for media optimization:
    HKLM:\SOFTWARE\Microsoft\Teams\IsWVDEnvironment = 1 (set before Teams is
    provisioned, which is why it is step 3) and the Remote Desktop WebRTC Redirector
    Service from https://aka.ms/msrdcwebrtcsvc/msi. Both are installed only when
    missing, so a scheduled run on a session host costs no extra download - which also
    means an already installed redirector is not silently upgraded; -Force is what
    replaces it (older version removed first, because its MSI keeps one ProductCode
    across versions and plain /i then answers 1638). Without the switch the script
    only points out that the device looks like a session host.

    WebRTC is being retired: end of support 1 October 2026, end of availability
    1 April 2027. Its replacement, SlimCore, cannot be installed here at all: the
    plugin bundled with Windows App stages and registers the MSIX on the endpoint the
    user connects from, silently and without admin intervention. The session host only
    has to run a Teams build of 24193.1805.3040.8975 or newer; the endpoint needs
    Windows App 2.0.352.0 or newer (the classic Remote Desktop client is no longer
    supported). Preflight reports whichever side it is standing on, and on an endpoint
    it also checks the policies that block the staging - BlockNonAdminUserInstall,
    AllowAllTrustedApps and AppLocker. AppLocker is read rather than merely detected:
    only the packaged-app (Appx) collection can stop an MSIX, a collection holding
    rules with enforcement "not configured" is enforced all the same, and nothing is
    enforced at all while the Application Identity service is stopped. The registry
    path, the mode per collection, the service state and the rule names are printed,
    and a policy found on a session host is named for reference rather than warned
    about, because the staging it would block happens on the endpoint.

    IsWVDEnvironment stays required either way, and Microsoft still advises keeping
    the redirector as a fallback for endpoints that cannot do SlimCore, so
    -AvdOptimizations keeps installing it. -RemoveWebRtcRedirector is the other
    direction, for a fleet that has finished moving: it uninstalls the
    redirector and leaves everything else alone. Off by default, and it should stay
    off until every endpoint really does run Windows App 2.0.352.0 or newer - an
    endpoint that cannot do SlimCore and no longer finds the redirector renders media
    on the session host instead, which is the outcome both of them exist to avoid.

    Proof, as opposed to inventory
    ------------------------------
    All of the above only shows that the parts are in place. Whether users are really
    optimized is something only the session host can answer, and it answers it in the
    Application event log: Teams writes a "Microsoft Teams VDI" event (ID 0) on every
    connect and disconnect, and the description carries the codes from Microsoft's
    connection error table. Preflight reads the last seven days of it and translates
    the codes it knows - 24002 and 24010 mean the user is on SlimCore, 16002 means the
    endpoint has no plugin, 16026 means a Citrix policy blocks the virtual channels,
    16389 means BlockNonAdminUserInstall stopped the MSIX registration. This is the
    one check here that reports on the endpoints rather than on this machine, which
    makes -CheckOnly on a session host worth running on its own.

    The meeting add-in
    ------------------
    A full reinstall removes every copy before putting the new one back: the MSI,
    the machine-wide folder, the per-profile folders under
    %LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in, and the per-user COM registrations
    in each loaded hive. Leaving any of those behind is what creates a user whose own
    registration shadows the fresh machine-wide one and points at files that are no
    longer there - Outlook then fails to load it and parks LoadBehavior at 2.

    That sweep waits until step 8, with the replacement MSI in hand and its version
    compared against what is still registered. Windows Installer refuses to put an
    older add-in over a newer registered one (1638), and a run that had already
    deleted every working copy by then left a production host with no add-in at all.
    Two failures make that reachable: a -Force run on a host whose Teams build is
    newer than the published one downgrades the client, so the add-in inside the
    package is older than the one registered; and an uninstall that answers 1612
    ("the installation source is not available") leaves the old product registered
    for good. 1612 is retried against Windows Installer's own cached copy of the MSI
    under C:\Windows\Installer, which usually still works; when that is gone too the
    registration cannot be removed by msiexec at all, and the script says so instead
    of destroying what still works.

    Classic Teams
    -------------
    -RemoveClassicTeams removes the old client as well. The machine-wide installer
    goes through msiexec; that one matters most, because as long as it is there
    Windows keeps staging classic Teams into every new profile. Per profile the
    install root, the Run entry (com.squirrel.Teams.Teams) and the stale uninstall
    key are removed. The documented per-user uninstall (Update.exe --uninstall -s)
    has to run as the profile owner, which System cannot do, so the files are removed
    instead. A profile whose files are locked by a running classic Teams finishes on
    a later run; a machine-wide installer that survives is treated as a failure.
    Roaming data in %APPDATA%\Microsoft\Teams is left alone.

    Every state-changing step is wrapped in ShouldProcess, so -WhatIf walks the whole
    flow and reports exactly what would be uninstalled, downloaded, installed and
    provisioned without touching the machine.

    Version check
    -------------
    https://config.teams.microsoft.com/config/v1/MicrosoftTeams/... is the feed the
    Teams client itself uses to decide it is out of date. It returns the current build
    per architecture (BuildSettings.WebView2PreAuth.<arch>.latestVersion). An
    installed build that is equal or newer means there is nothing to do. The installed
    build comes from the per-user AppX packages, falling back to the provisioned
    package - on a pooled session host or a fresh image Teams is often provisioned
    without any user having it yet, and without that fallback every run would call the
    host outdated and reinstall it. If the service cannot be reached the run stops
    rather than reinstalling blindly; -Force overrides that.

    Safety
    ------
      - Nothing is touched until a newer build is confirmed (or -Force is given).
      - The installer is downloaded and signature-checked before the first uninstall.
      - Every msiexec/bootstrapper call runs with a timeout and is killed if it hangs,
        so an RMM job can never block the agent indefinitely.
      - MSI exit code 1618 (another install in progress) is retried; 3010 is treated
        as success with a reboot flagged in the summary.
      - Unexpected errors abort the run instead of continuing half-way.
      - A run that actually changes something writes a transcript to the log folder;
        a check that finds nothing to do leaves no log litter behind.

    Exit codes
    ----------
        0  success, or already up to date
        1  failure
        2  -CheckOnly only: a newer build is available

    Running it by hand
    ------------------
    From an ordinary PowerShell window the script elevates itself (UAC) and continues
    in a new elevated window that stays open, so no "run as administrator" dance is
    needed first. An interactive run that is about to change something asks for
    confirmation once; -Confirm:$false skips that question. A -WhatIf run never asks.

        powershell -ExecutionPolicy Bypass -File .\Update-TeamsClient.ps1 -WhatIf

    RMM / NinjaOne
    --------------
    The script is safe to deploy from NinjaOne (run as System):

      - With -Quiet it prints nothing at all while Teams is up to date, so a scheduled
        run only shows up in the activity feed when it actually found a newer build or
        hit a problem. Combine with -CheckOnly for a pure detection job (exit code 2
        means "update available").
      - Script variables arrive as environment variables, so checkboxes named whatIf,
        quiet, checkOnly, force, avdOptimizations, removeClassicTeams, repairOutlookAddIn,
        skipMeetingAddIn or skipSignatureCheck and text fields named workingDir,
        logPath or ring are picked up when the matching parameter is not passed.
        Capitalisation does not matter - environment lookups are case-insensitive.
      - If the agent starts PowerShell 32-bit, the script relaunches itself 64-bit
        via SysNative first. Without that, registry reads are redirected to
        WOW6432Node and $env:ProgramFiles points at the x86 folder, so the AppX
        package and the add-in MSI are never found.
      - Add -Confirm:$false to the Parameters field so the confirmation question can
        never appear, whatever the agent reports about the session.

.PARAMETER Quiet
    Print nothing unless there is news: a newer build, an action taken, or a failure.
    Intended for scheduled RMM runs.

.PARAMETER CheckOnly
    Only report whether a newer build is available and change nothing. Exit code 2
    means an update is available, 0 means up to date.

.PARAMETER Ring
    Update ring queried at the config service (default: general). Both audienceGroup
    and teamsRing are set to this value.

.PARAMETER AvdOptimizations
    Also enforce the AVD/VDI media optimizations: the IsWVDEnvironment registry flag
    and the Remote Desktop WebRTC Redirector Service. Both only when missing, unless
    -Force is given.

.PARAMETER RemoveWebRtcRedirector
    Remove the Remote Desktop WebRTC Redirector Service - the old media optimization,
    unsupported after 1 October 2026. Off by default and mutually exclusive with
    -AvdOptimizations. Only use it once every endpoint can do SlimCore (Windows App
    2.0.352.0 or newer); endpoints that cannot fall back to rendering media on the
    session host. IsWVDEnvironment is deliberately left alone, because SlimCore needs
    that flag too.

.PARAMETER RemoveClassicTeams
    Also remove the classic Teams client: uninstall the Teams Machine-Wide Installer
    and clear the per-profile installs (install folder, autostart entry and the stale
    uninstall key). Off by default - taking an app away from users is not something an
    update job should decide by itself.

.PARAMETER RepairOutlookAddIn
    Clear a per-user Outlook registration that points at an add-in DLL which no
    longer exists, so the machine-wide registration takes over again. Only acts when
    that machine-wide registration is healthy. Off by default: it writes into another
    user's hive, which is not something an update job should do unasked.

.PARAMETER WebRtcUrl
    Download URL for the Remote Desktop WebRTC Redirector MSI (default: the Microsoft
    aka.ms link). Must be https. Only used with -AvdOptimizations.

.EXAMPLE
    .\Update-TeamsClient.ps1 -CheckOnly

    Full read-only health report of every place Teams lives on this machine,
    including what the session host logged about the media optimization. Changes
    nothing; exit code 2 means there is work to do.

.EXAMPLE
    .\Update-TeamsClient.ps1 -RemoveWebRtcRedirector -WhatIf

    Show what removing the old WebRTC optimization would do. Drop -WhatIf to apply.

.PARAMETER WorkingDir
    Folder used for the installer downloads (default: C:\IT\AVD\Teams).

.PARAMETER LogPath
    Folder for the transcript of a run that changes something (default: C:\Temp).
    A -WhatIf, -CheckOnly or up-to-date run writes no transcript.

.PARAMETER BootstrapperUrl
    Download URL for teamsbootstrapper.exe (default: the Microsoft fwlink for new
    Teams). Must be https.

.PARAMETER SkipMeetingAddIn
    Leave the Teams Meeting Add-in alone - do not uninstall it up front, do not
    install it afterwards and do not treat a missing add-in as work to do. Reasonable
    on ordinary endpoints, where the Teams client installs and updates the add-in per
    user by itself; the machine-wide install this script performs is what a shared
    machine or session host needs.

.PARAMETER SkipSignatureCheck
    Accept the downloaded bootstrapper without verifying its Authenticode signature.
    Only for an air-gapped or internally hosted -BootstrapperUrl that is not signed
    by Microsoft.

.PARAMETER TimeoutSeconds
    Per-process timeout for msiexec and the bootstrapper (default: 900). A process
    that outlives it is killed and the step is reported as failed.

.PARAMETER Force
    Reinstall even when Teams is already current, and continue when no Teams
    installation or no version information is found at all.

.EXAMPLE
    # Dry run - check for a newer build and show what an update would do
    .\Update-TeamsClient.ps1 -WhatIf

.EXAMPLE
    # Update only if Microsoft published a newer build
    .\Update-TeamsClient.ps1

.EXAMPLE
    # NinjaOne, scheduled: silent unless there is a newer build or a problem
    .\Update-TeamsClient.ps1 -Quiet -Confirm:$false

.EXAMPLE
    # NinjaOne, detection only: exit code 2 when an update is available
    .\Update-TeamsClient.ps1 -CheckOnly -Quiet

.EXAMPLE
    # AVD session host: Teams plus the media flag and the WebRTC redirector
    .\Update-TeamsClient.ps1 -AvdOptimizations -Quiet -Confirm:$false

.EXAMPLE
    # Clean up the old client while keeping new Teams current
    .\Update-TeamsClient.ps1 -RemoveClassicTeams -Quiet -Confirm:$false

.EXAMPLE
    # Repair: reinstall the current build regardless of the version check
    .\Update-TeamsClient.ps1 -Force -Confirm:$false

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows only (AppX + MSI), run as administrator or as System
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [switch] $Quiet,
    [switch] $CheckOnly,
    [switch] $AvdOptimizations,
    [switch] $RemoveClassicTeams,
    [switch] $RepairOutlookAddIn,
    [switch] $RemoveWebRtcRedirector,
    [string] $Ring            = 'general',
    [string] $WorkingDir      = 'C:\IT\AVD\Teams',
    [string] $LogPath         = 'C:\Temp',
    [string] $BootstrapperUrl = 'https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409',
    [string] $WebRtcUrl       = 'https://aka.ms/msrdcwebrtcsvc/msi',
    [switch] $SkipMeetingAddIn,
    [switch] $SkipSignatureCheck,
    [ValidateRange(60, 7200)]
    [int]    $TimeoutSeconds  = 900,
    [switch] $Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-ForwardedArgument {
    <# Rebuild the caller's own parameters as a command line for a relaunch. #>
    param([Parameter(Mandatory)] $Bound)

    $list = @()
    foreach ($entry in $Bound.GetEnumerator()) {
        if ($entry.Value -is [switch] -or $entry.Value -is [bool]) {
            # -Confirm:$false and friends must survive as an explicit :$false.
            if ($entry.Value) { $list += "-$($entry.Key)" } else { $list += "-$($entry.Key):`$false" }
        } else {
            $list += "-$($entry.Key)"
            $list += [string] $entry.Value
        }
    }
    return $list
}

function Test-Elevated {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    return ([Security.Principal.WindowsPrincipal] $identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# -- RMM: run 64-bit -----------------------------------------------------------
# An RMM agent (NinjaOne among them) can start PowerShell 32-bit. Under WOW64 the
# HKLM reads land in WOW6432Node and $env:ProgramFiles points at the x86 folder,
# which hides both the AppX package and the add-in MSI. Relaunch ourselves 64-bit
# and hand the parameters over unchanged.
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $nativeShell = Join-Path $env:WINDIR 'SysNative\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path $nativeShell) {
        $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath) +
                   (Get-ForwardedArgument -Bound $PSBoundParameters)
        & $nativeShell @argList
        exit $LASTEXITCODE
    }
    Write-Warning 'Running 32-bit and SysNative is unavailable - AppX and registry lookups may fail.'
}

# A combination that can never work is caught before the UAC prompt: nobody should
# have to approve elevation for a run that is only going to refuse itself. The same
# check runs again below, once the RMM variables have been folded in.
if ($AvdOptimizations -and $RemoveWebRtcRedirector) {
    Write-Error 'Use either -AvdOptimizations (which installs the WebRTC redirector) or -RemoveWebRtcRedirector, not both.'
    exit 1
}

# -- Elevation -----------------------------------------------------------------
# AppX enumeration, the MSI calls and the bootstrapper all need administrator rights.
# Run as System from an RMM this is already true; started by hand it is usually not,
# so ask for elevation instead of failing on a #Requires line.
if (-not (Test-Elevated)) {
    if (-not [Environment]::UserInteractive) {
        Write-Error 'Administrator rights are required. Run the script as System or from an elevated session.'
        exit 1
    }

    Write-Host ''
    Write-Host '  Not running elevated - asking for administrator rights...' -ForegroundColor Yellow
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-NoExit', '-File', $PSCommandPath) +
               (Get-ForwardedArgument -Bound $PSBoundParameters)
    try {
        Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList $argList -Verb RunAs | Out-Null
        Write-Host '  Continued in an elevated window.' -ForegroundColor Cyan
        exit 0
    } catch {
        Write-Error "Elevation was declined or failed: $($_.Exception.Message)"
        exit 1
    }
}

# -- RMM: script variables -----------------------------------------------------
# NinjaOne exposes script variables as environment variables. Honour them only
# when the matching parameter was not passed on the command line, so a plain
# .\Update-TeamsClient.ps1 -WhatIf keeps working exactly as before. Environment
# lookups are case-insensitive, so a variable named Quiet or QUIET works too.
$rmmTrue = @('true', '1', 'yes')
if (-not $PSBoundParameters.ContainsKey('WhatIf')             -and $env:whatIf             -in $rmmTrue) { $WhatIfPreference   = $true }
if (-not $PSBoundParameters.ContainsKey('Quiet')              -and $env:quiet              -in $rmmTrue) { $Quiet              = $true }
if (-not $PSBoundParameters.ContainsKey('CheckOnly')          -and $env:checkOnly          -in $rmmTrue) { $CheckOnly          = $true }
if (-not $PSBoundParameters.ContainsKey('Force')              -and $env:force              -in $rmmTrue) { $Force              = $true }
if (-not $PSBoundParameters.ContainsKey('SkipMeetingAddIn')   -and $env:skipMeetingAddIn   -in $rmmTrue) { $SkipMeetingAddIn   = $true }
if (-not $PSBoundParameters.ContainsKey('SkipSignatureCheck') -and $env:skipSignatureCheck -in $rmmTrue) { $SkipSignatureCheck = $true }
if (-not $PSBoundParameters.ContainsKey('AvdOptimizations')   -and $env:avdOptimizations   -in $rmmTrue) { $AvdOptimizations   = $true }
if (-not $PSBoundParameters.ContainsKey('RemoveClassicTeams') -and $env:removeClassicTeams -in $rmmTrue) { $RemoveClassicTeams = $true }
if (-not $PSBoundParameters.ContainsKey('RepairOutlookAddIn') -and $env:repairOutlookAddIn -in $rmmTrue) { $RepairOutlookAddIn = $true }
if (-not $PSBoundParameters.ContainsKey('RemoveWebRtcRedirector') -and $env:removeWebRtcRedirector -in $rmmTrue) { $RemoveWebRtcRedirector = $true }
if (-not $PSBoundParameters.ContainsKey('WorkingDir')         -and $env:workingDir)                      { $WorkingDir         = $env:workingDir }
if (-not $PSBoundParameters.ContainsKey('LogPath')            -and $env:logPath)                         { $LogPath            = $env:logPath }
if (-not $PSBoundParameters.ContainsKey('Ring')               -and $env:ring)                            { $Ring               = $env:ring }
if (-not $PSBoundParameters.ContainsKey('WebRtcUrl')          -and $env:webRtcUrl)                       { $WebRtcUrl          = $env:webRtcUrl }

if ($AvdOptimizations -and $RemoveWebRtcRedirector) {
    Write-Error 'Use either -AvdOptimizations (which installs the WebRTC redirector) or -RemoveWebRtcRedirector, not both.'
    exit 1
}

$simulate         = [bool] $WhatIfPreference
$teamsExe         = 'teamsbootstrapper.exe'
$exePath          = Join-Path $WorkingDir $teamsExe
$webRtcMsi        = Join-Path $WorkingDir 'MsRdcWebRTCSvc_x64.msi'
$exitCode         = 0
$rebootRequired   = $false
$transcribing     = $false
$plannedExit      = $null
$workingDirReady  = $false
$addInOrphaned    = $false

# Teams reads this flag to switch to VDI media optimization; it has to be there
# before the client is provisioned.
$AvdRegistryPath = 'HKLM:\SOFTWARE\Microsoft\Teams'

# The meeting add-in's COM class. Stable across versions, and the only way to find
# out which DLL Outlook would actually load for a given user.
$AddInClsid = '{19A6E644-14E6-4A60-B8D7-DD20610A871D}'

# SlimCore, the WebRTC successor, is staged and registered on the *endpoint* the
# user connects from - the plugin does it silently, without admin intervention -
# and never on the session host. These are Microsoft's documented minimums for
# Azure Virtual Desktop and Windows 365.
$SlimCoreMinimumTeamsBuild = [version] '24193.1805.3040.8975'
$SlimCoreMinimumWindowsApp = '2.0.352.0'

# -Confirm:$false means "never ask", for an unattended run from a scheduler or RMM.
$confirmSuppressed = $PSBoundParameters.ContainsKey('Confirm') -and -not $PSBoundParameters['Confirm']

# Which hive the add-in's uninstall entry lands in is not fixed: measured as 64-bit
# for 1.26.21803 on Windows 11, WOW6432Node on other builds. Scanning both is the
# only thing that reliably finds it.
$UninstallRoots = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)

# -- Output --------------------------------------------------------------------
# In -Quiet mode every line is held back until something worth reporting happens.
# A scheduled run on an up-to-date device then produces no output at all.
$script:heldOutput = [System.Collections.Generic.List[object]]::new()
$script:holdOutput = [bool] $Quiet

function Write-Out {
    param([string] $Message = '', [string] $Color = 'Gray')
    if ($script:holdOutput) { $script:heldOutput.Add([PSCustomObject]@{ Message = $Message; Color = $Color }) }
    else { Write-Host $Message -ForegroundColor $Color }
}
function Show-HeldOutput {
    if (-not $script:holdOutput) { return }
    $script:holdOutput = $false
    foreach ($line in $script:heldOutput) { Write-Host $line.Message -ForegroundColor $line.Color }
    $script:heldOutput.Clear()
}

function Write-Step { param([string] $Message) Write-Out "  $Message" 'Cyan' }
function Write-Ok   { param([string] $Message) Write-Out "  [ OK ] $Message" 'Green' }
function Write-Skip { param([string] $Message) Write-Out "  [SKIP] $Message" 'DarkGray' }
function Write-Warn { param([string] $Message) Write-Out "  [WARN] $Message" 'Yellow' }
function Write-Bad  { param([string] $Message) Write-Out "  [FAIL] $Message" 'Red' }
function Write-News { param([string] $Message) Write-Out "  [NEW ] $Message" 'Magenta' }

function Get-PropertyValue {
    <# Property access that returns $null instead of tripping Set-StrictMode. #>
    param($Object, [string] $Name)

    if ($null -eq $Object) { return $null }
    if ($Object.PSObject.Properties.Name -notcontains $Name) { return $null }
    return $Object.$Name
}

function Get-TeamsMeetingAddInEntry {
    <# Uninstall entries for the Teams Meeting Add-in, from both registry views. #>
    foreach ($root in $UninstallRoots) {
        if (-not (Test-Path $root)) { continue }
        foreach ($key in Get-ChildItem $root) {
            $name = $key.GetValue('DisplayName')
            if ($name -like '*Microsoft Teams Meeting Add-in*') {
                [PSCustomObject]@{
                    ProductCode     = $key.PSChildName
                    DisplayName     = $name
                    Version         = $key.GetValue('DisplayVersion')
                    RegistryPath    = $key.PSPath
                    UninstallString = $key.GetValue('UninstallString')
                }
            }
        }
    }
}

function Get-ClassicTeamsEntry {
    <# The classic Teams Machine-Wide Installer MSI, from both registry views. #>
    foreach ($root in $UninstallRoots) {
        if (-not (Test-Path $root)) { continue }
        foreach ($key in Get-ChildItem $root) {
            $name = $key.GetValue('DisplayName')
            if ($name -like '*Teams Machine-Wide Installer*') {
                [PSCustomObject]@{
                    ProductCode     = $key.PSChildName
                    DisplayName     = $name
                    Version         = $key.GetValue('DisplayVersion')
                    RegistryPath    = $key.PSPath
                    UninstallString = $key.GetValue('UninstallString')
                }
            }
        }
    }
}

function Test-UserProfileSid {
    <#
        A real user profile, as opposed to a service account or a helper hive.
        Deliberately not a whitelist on S-1-5-21: an Entra-joined device hands out
        S-1-12-1 SIDs, and matching on the domain format skips every user there.
    #>
    param([string] $Sid)

    if (-not $Sid -or $Sid -notlike 'S-1-*') { return $false }
    if ($Sid -like '*_Classes') { return $false }
    return ($Sid -notin @('S-1-5-18', 'S-1-5-19', 'S-1-5-20'))
}

function Get-ClassicTeamsUserInstall {
    <#
        Classic Teams installs itself into every user profile. Enumerating the profile
        list rather than Get-ChildItem C:\Users keeps it accurate when profiles live
        on another drive or a redirected path.
    #>
    $profileList = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    if (-not (Test-Path $profileList)) { return }

    foreach ($key in Get-ChildItem $profileList) {
        $sid = $key.PSChildName
        if (-not (Test-UserProfileSid $sid)) { continue }

        $profilePath = $key.GetValue('ProfileImagePath')
        if (-not $profilePath) { continue }

        $exe = Join-Path $profilePath 'AppData\Local\Microsoft\Teams\current\Teams.exe'
        if (-not (Test-Path $exe)) { continue }

        [PSCustomObject]@{
            Sid         = $sid
            Account     = Resolve-SidName $sid
            Path        = $exe
            InstallRoot = Join-Path $profilePath 'AppData\Local\Microsoft\Teams'
            Version     = (Get-Item $exe -ErrorAction SilentlyContinue).VersionInfo.FileVersion
        }
    }
}

function Resolve-SidName {
    <# SID to DOMAIN\user, or the SID itself when it cannot be translated. #>
    param([Parameter(Mandatory)] [string] $Sid)

    try { return (New-Object Security.Principal.SecurityIdentifier $Sid).Translate([Security.Principal.NTAccount]).Value }
    catch { return $Sid }
}

function Get-UninspectableProfile {
    <#
        Profiles that exist on this machine but whose hive is not mounted, so their
        Outlook registration cannot be read at all. Reporting them by name beats
        leaving them out of the output: "not listed" and "fine" look identical
        otherwise, which is exactly the question a technician asks when the add-in
        does not appear for everyone.
    #>
    $loaded      = @(Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty PSChildName)
    $profileList = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    if (-not (Test-Path $profileList)) { return }

    foreach ($key in Get-ChildItem $profileList) {
        $sid = $key.PSChildName
        if (-not (Test-UserProfileSid $sid)) { continue }
        if ($loaded -contains $sid) { continue }

        [PSCustomObject]@{
            Sid         = $sid
            Account     = Resolve-SidName $sid
            ProfilePath = $key.GetValue('ProfileImagePath')
        }
    }
}

function Get-TeamsAddInFolder {
    <#
        Every folder a meeting add-in copy can sit in: the machine-wide one this
        script installs, and the per-profile one the Teams client installs by itself
        (note the hyphen - the two spellings are not a typo, Microsoft uses both).
    #>
    $machineWide = Join-Path ${env:ProgramFiles(x86)} 'Microsoft\TeamsMeetingAddin'
    if (Test-Path $machineWide) {
        [PSCustomObject]@{ Account = 'all users (machine-wide)'; Path = $machineWide; Sid = $null }
    }

    $profileList = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    if (-not (Test-Path $profileList)) { return }

    foreach ($key in Get-ChildItem $profileList) {
        $sid = $key.PSChildName
        if (-not (Test-UserProfileSid $sid)) { continue }

        $profilePath = $key.GetValue('ProfileImagePath')
        if (-not $profilePath) { continue }

        $perUser = Join-Path $profilePath 'AppData\Local\Microsoft\TeamsMeetingAdd-in'
        if (Test-Path $perUser) {
            [PSCustomObject]@{ Account = (Resolve-SidName $sid); Path = $perUser; Sid = $sid }
        }
    }
}

function Get-AddInDllPath {
    <#
        Where a hive's COM registration says the add-in loader lives. Outlook resolves
        the add-in through this CLSID, so a path that no longer exists is the real
        reason behind a LoadBehavior that keeps dropping back to 2.
    #>
    param([Parameter(Mandatory)] [string] $ClassesRoot)

    foreach ($candidate in "$ClassesRoot\CLSID\$AddInClsid\InprocServer32",
                           "$ClassesRoot\WOW6432Node\CLSID\$AddInClsid\InprocServer32") {
        if (-not (Test-Path $candidate)) { continue }
        $dll = Get-PropertyValue (Get-ItemProperty $candidate -ErrorAction SilentlyContinue) '(default)'
        if ($dll) { return $dll }
    }
    return $null
}

function Get-AddInClsidKey {
    <# The CLSID keys a hive holds for the add-in, in either registry view. #>
    param([Parameter(Mandatory)] [string] $ClassesRoot)

    foreach ($candidate in "$ClassesRoot\CLSID\$AddInClsid", "$ClassesRoot\WOW6432Node\CLSID\$AddInClsid") {
        if (Test-Path $candidate) { $candidate }
    }
}

function Get-OutlookAddInRegistration {
    <#
        Whether Outlook itself sees the add-in. The MSI in Program Files proves the
        machine-wide install; Outlook loads COM add-ins from its own key, so that is
        what decides whether the meeting button is actually there. LoadBehavior 3 =
        loaded at startup, 2 or 0 = Outlook switched it off.

        Run as System only the hives of signed-in users are mounted under HKEY_USERS,
        so a profile nobody is logged into is invisible here - not broken, just unseen.
    #>
    $suffix  = 'Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect'
    $results = [System.Collections.Generic.List[object]]::new()

    # Both registry views are read, and both can hold an entry, so they are labelled
    # apart - two identical "all users" lines only look like a bug in the output.
    $machineViews = [ordered] @{
        "HKLM:\SOFTWARE\$suffix"                = 'all users (machine-wide, 64-bit)'
        "HKLM:\SOFTWARE\WOW6432Node\$suffix"    = 'all users (machine-wide, 32-bit)'
    }
    foreach ($machinePath in $machineViews.Keys) {
        if (-not (Test-Path $machinePath)) { continue }
        $dll = Get-AddInDllPath 'HKLM:\SOFTWARE\Classes'
        $results.Add([PSCustomObject]@{
            Sid          = $null
            Account      = $machineViews[$machinePath]
            LoadBehavior = Get-PropertyValue (Get-ItemProperty $machinePath -ErrorAction SilentlyContinue) 'LoadBehavior'
            AddInKeyPath = $machinePath
            ClassesRoot  = 'HKLM:\SOFTWARE\Classes'
            DllPath      = $dll
            DllMissing   = ($dll -and -not (Test-Path $dll))
        })
    }

    # Addressed through the provider directly instead of New-PSDrive: that cmdlet
    # supports ShouldProcess, so under -WhatIf the drive would not be created and this
    # read-only check would wrongly report that nobody has the add-in registered.
    foreach ($hive in (Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue)) {
        $sid = $hive.PSChildName
        if (-not (Test-UserProfileSid $sid)) { continue }

        $userPath = "Registry::HKEY_USERS\$sid\SOFTWARE\$suffix"
        if (-not (Test-Path $userPath)) { continue }
        $dll = Get-AddInDllPath "Registry::HKEY_USERS\$sid\SOFTWARE\Classes"
        $results.Add([PSCustomObject]@{
            Sid          = $sid
            Account      = Resolve-SidName $sid
            LoadBehavior = Get-PropertyValue (Get-ItemProperty $userPath -ErrorAction SilentlyContinue) 'LoadBehavior'
            AddInKeyPath = $userPath
            ClassesRoot  = "Registry::HKEY_USERS\$sid\SOFTWARE\Classes"
            DllPath      = $dll
            DllMissing   = ($dll -and -not (Test-Path $dll))
        })
    }

    return $results
}

function Get-OfficeBitness {
    <# x64 or x86, from the Click-to-Run configuration. $null when it cannot be read. #>
    $cfg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
    return (Get-PropertyValue $cfg 'Platform')
}

function Get-AddInLoadBlocker {
    <#
        Why Outlook would refuse to load the add-in, beyond the registration being
        present. These are the three that leave no trace in LoadBehavior alone:
        a bitness mismatch, Outlook's own resiliency lists, and a policy-managed
        load behaviour that overrides whatever the user has.
    #>
    param([Parameter(Mandatory)] $Registration)

    $reasons = [System.Collections.Generic.List[string]]::new()

    # Outlook can only load a loader of its own bitness, and the add-in ships both.
    $bitness = Get-OfficeBitness
    if ($bitness -and $Registration.DllPath) {
        $dllArch = if ($Registration.DllPath -match '\\x86\\') { 'x86' }
                   elseif ($Registration.DllPath -match '\\x64\\') { 'x64' }
                   else { $null }
        if ($dllArch -and $dllArch -ne $bitness) {
            $reasons.Add("Outlook is $bitness but this registration points at the $dllArch loader ($($Registration.DllPath)) - a mismatched DLL cannot load")
        }
    }

    # Outlook parks add-ins that crashed or started slowly, and keeps them parked.
    if ($Registration.Sid) {
        foreach ($office in '16.0', '15.0') {
            foreach ($leaf in 'DisabledItems', 'CrashedAddins') {
                $key = "Registry::HKEY_USERS\$($Registration.Sid)\Software\Microsoft\Office\$office\Outlook\Resiliency\$leaf"
                if (-not (Test-Path $key)) { continue }

                $entries = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
                foreach ($property in $entries.PSObject.Properties) {
                    if ($property.Name -like 'PS*') { continue }
                    $text = ''
                    try { $text = [Text.Encoding]::Unicode.GetString([byte[]] $property.Value) } catch { $text = '' }
                    if ($text -match 'TeamsMeetingAdd|TeamsAddin|AddinLoader') {
                        $reasons.Add("Outlook parked the add-in in its $leaf list (Office $office) after a crash or a slow start - clear it under File > Options > Add-ins > Manage: Disabled Items, and keep it out with the DoNotDisableAddinList policy")
                    }
                }
            }
        }
    }

    # A policy-set load behaviour wins over whatever the user has.
    foreach ($office in '16.0', '15.0') {
        $policy = Get-ItemProperty "HKLM:\SOFTWARE\Policies\Microsoft\Office\$office\Outlook\Addins\TeamsAddin.FastConnect" -ErrorAction SilentlyContinue
        $policyBehavior = Get-PropertyValue $policy 'LoadBehavior'
        if ($null -ne $policyBehavior -and $policyBehavior -ne 3) {
            $reasons.Add("Group policy sets LoadBehavior $policyBehavior for this add-in (Office $office) - the policy overrides the user's own setting")
        }
    }

    return ($reasons | Select-Object -Unique)
}

function Write-OutlookAddInStatus {
    <# Report the Outlook-side registration. Informational: a missing per-user entry
       appears by itself at the next Outlook start, and a profile that is not signed
       in cannot be inspected at all, so neither is treated as a failure. #>
    param($Registrations)

    $registrations = if ($Registrations) { @($Registrations) } else { @(Get-OutlookAddInRegistration) }
    $machineWideOk = [bool] ($registrations | Where-Object { $_.Account -like 'all users*' -and -not $_.DllMissing })

    # Profiles nobody is signed into cannot be read. Say so, with names: a profile
    # that is simply absent from the output is indistinguishable from a healthy one.
    $unseen = @(Get-UninspectableProfile)
    if ($unseen.Count -gt 0) {
        $consequence = if ($machineWideOk) {
            'they pick up the machine-wide registration the first time that user starts Outlook'
        } else {
            'and there is no machine-wide registration for them to fall back on'
        }
        Write-Skip ("{0} profile(s) are not signed in, so their Outlook registration cannot be read ({1}) - {2}" -f
                    $unseen.Count, (($unseen | Select-Object -ExpandProperty Account) -join ', '), $consequence)
    }

    if ($registrations.Count -eq 0) {
        Write-Warn 'Outlook has not registered the add-in for any signed-in user yet - it registers itself at the next Outlook start'
        return
    }

    foreach ($reg in $registrations) {
        # A registration pointing at a DLL that is gone is the reason a LoadBehavior
        # keeps falling back to 2: Outlook tries, fails and switches the add-in off
        # again. Ticking the box back on does not survive that, a reinstall for that
        # user does, so the two cases are reported apart.
        if ($reg.DllMissing) {
            Write-Warn "The add-in is registered for $($reg.Account) but its DLL is gone ($($reg.DllPath)) - re-enabling it in Outlook will not stick; the add-in has to be installed again for that user"
        } elseif ($reg.LoadBehavior -eq 3) {
            Write-Ok "Outlook loads the add-in for $($reg.Account)"
        } elseif ($null -eq $reg.LoadBehavior) {
            Write-Warn "Outlook knows the add-in for $($reg.Account) but has no LoadBehavior set - it should appear at the next Outlook start"
        } else {
            Write-Warn "Outlook has the add-in switched off for $($reg.Account) (LoadBehavior $($reg.LoadBehavior))"
        }

        # Anything that is not plainly loading gets a reason, or an explicit "no
        # reason found here" - "it does not work" without a cause is what turns a
        # ticket into three more.
        if ($reg.DllMissing -or $reg.LoadBehavior -ne 3) {
            $blockers = @(Get-AddInLoadBlocker -Registration $reg)
            if ($blockers.Count -gt 0) {
                foreach ($blocker in $blockers) { Write-Warn "  why: $blocker" }
            } elseif (-not $reg.DllMissing) {
                Write-Skip '  why: nothing on this machine blocks it - Outlook has to be fully closed and started once, and the user must have signed in to Teams at least once'
            }
        }
    }
}

function Get-WebRtcRedirectorEntry {
    <# Uninstall entry for the Remote Desktop WebRTC Redirector Service, if any. #>
    foreach ($root in $UninstallRoots) {
        if (-not (Test-Path $root)) { continue }
        foreach ($key in Get-ChildItem $root) {
            $name = $key.GetValue('DisplayName')
            if ($name -like '*Remote Desktop WebRTC Redirector Service*') {
                [PSCustomObject]@{
                    ProductCode     = $key.PSChildName
                    DisplayName     = $name
                    Version         = $key.GetValue('DisplayVersion')
                    RegistryPath    = $key.PSPath
                    UninstallString = $key.GetValue('UninstallString')
                }
            }
        }
    }
}

function Get-WvdEnvironmentFlag {
    <# Current value of IsWVDEnvironment, or $null when it is not set at all. #>
    if (-not (Test-Path $AvdRegistryPath)) { return $null }
    return Get-PropertyValue (Get-ItemProperty -Path $AvdRegistryPath -ErrorAction SilentlyContinue) 'IsWVDEnvironment'
}

function Test-AvdSessionHost {
    <# The AVD agent registers itself here - good enough to recognise a session host. #>
    return (Test-Path 'HKLM:\SOFTWARE\Microsoft\RDInfraAgent')
}

# Microsoft publishes one table for the VDI optimization errors with two columns of
# numbers - a standard error code and the loadErrc/deployErrc the plugin logs - and
# an event does not say which of the two it printed. Both are keyed here.
$TeamsVdiEventSource = 'Microsoft Teams VDI'
$TeamsVdiErrorCodes  = @{
    '3000'  = 'SlimCore deployment not needed - the user is on the new architecture'
    '24002' = 'SlimCore deployment not needed - the user is on the new architecture'
    '3001'  = 'SlimCore already loaded - the user is on the new architecture'
    '24010' = 'SlimCore already loaded - the user is on the new architecture'
    '5'     = 'access denied starting MsTeamsVdi.exe - BlockNonAdminUserInstall, or AppX was still registering packages'
    '43'    = 'access denied starting MsTeamsVdi.exe - BlockNonAdminUserInstall, or AppX was still registering packages'
    '110'   = 'a background access policy blocks BITS on the endpoint, so SlimCore cannot download'
    '403'   = 'the SlimCore download was refused (HTTP 403) - proxy or content filtering'
    '3227'  = 'the SlimCore download was refused (HTTP 403) - proxy or content filtering'
    '404'   = 'the SlimCore package was not found on the CDN (HTTP 404)'
    '3235'  = 'the SlimCore package was not found on the CDN (HTTP 404)'
    '1260'  = 'blocked by policy - AppLocker or WDAC without an exception for the SlimCoreVdi packages'
    '10083' = 'blocked by policy - AppLocker or WDAC without an exception for the SlimCoreVdi packages'
    '1460'  = 'MsTeamsVdi.exe timed out starting (60 seconds) - usually MSIX registration being slow'
    '11683' = 'MsTeamsVdi.exe timed out starting (60 seconds) - usually MSIX registration being slow'
    '1722'  = 'RPC to MsTeamsVdi.exe was unavailable - transient if optimization comes up anyway'
    '13779' = 'RPC to MsTeamsVdi.exe was unavailable - transient if optimization comes up anyway'
    '2000'  = 'no plugin on the endpoint - the remote desktop client has none, or it did not load'
    '16002' = 'no plugin on the endpoint - the remote desktop client has none, or it did not load'
    '2001'  = 'the virtual channel is not available on the VDA agent'
    '16008' = 'the virtual channel is not available on the VDA agent'
    '2003'  = 'the MSTEAMS/MSTEAM1/MSTEAM2 virtual channels are blocked by a Citrix policy'
    '16026' = 'the MSTEAMS/MSTEAM1/MSTEAM2 virtual channels are blocked by a Citrix policy'
    '2004'  = 'the Citrix Workspace App on the endpoint is too old'
    '16032' = 'the Citrix Workspace App on the endpoint is too old'
    '2005'  = 'Teams runs as a published app or RemoteApp - those sessions stay on WebRTC, never SlimCore'
    '16043' = 'Teams runs as a published app or RemoteApp - those sessions stay on WebRTC, never SlimCore'
    '2008'  = 'the Mac endpoint uses the Store build of Windows App - only the non-store build is supported'
    '16066' = 'the Mac endpoint uses the Store build of Windows App - only the non-store build is supported'
    '3002'  = 'SlimCore was never downloaded on the endpoint'
    '24018' = 'SlimCore was never downloaded on the endpoint'
    '3004'  = 'the plugin did not answer - usually transient while it stages the media engine'
    '24035' = 'the plugin did not answer - usually transient while it stages the media engine'
    '3005'  = 'the plugin timed out downloading the MSIX (2 minutes)'
    '24043' = 'the plugin timed out downloading the MSIX (2 minutes)'
    '3007'  = 'the SlimCore download or installation timed out'
    '24058' = 'the SlimCore download or installation timed out'
    '3021'  = 'no usable SlimCore package on the endpoint - the Host MSIX is missing'
    '24170' = 'no usable SlimCore package on the endpoint - the Host MSIX is missing'
    '4390'  = 'reparse point error - a thin client with a write filter or RAM disk; point MSTEAMSVDI_BITS_TMP_PATH at a real disk'
    '12002' = 'the endpoint timed out reaching the internet'
    '12030' = 'the connection to the Microsoft CDN was aborted'
    '1951'  = 'sideloading is off on the endpoint (AllowAllTrustedApps is 0)'
    '15615' = 'sideloading is off on the endpoint (AllowAllTrustedApps is 0)'
    '15616' = 'the SlimCore package was updating - retryable, not a deployment failure'
    '15618' = 'the SlimCore package is in use and could not be replaced'
    '15700' = 'no package identity - the MsTeamsVdi alias is missing from %LOCALAPPDATA%\Microsoft\WindowsApps'
    '16389' = 'Package Manager returned E_FAIL - most often BlockNonAdminUserInstall on the endpoint'
}

function Get-TeamsVdiEvent {
    <#
        Teams VDI events out of the Application log. Filtering on the provider has to
        go through FilterXPath: the FilterHashtable form throws outright when the
        provider has never written an event, which is exactly the case this has to
        survive. Teams writes these from build 24123 onwards.
    #>
    param([int] $Days = 7, [int] $MaxEvents = 200)

    $since = (Get-Date).AddDays(-$Days)
    return @(Get-WinEvent -LogName Application -MaxEvents $MaxEvents -ErrorAction SilentlyContinue `
                          -FilterXPath "*[System[Provider[@Name='$TeamsVdiEventSource']]]" |
             Where-Object { $_.TimeCreated -ge $since })
}

function Get-TeamsVdiErrorMeaning {
    <#
        Known codes out of an event description. Only numbers written as an error
        field are read: the same message carries version strings full of digits that
        would otherwise pass for error codes. A zero errc is deliberately not in the
        table - it only means that phase raised no error, and printing "OK" beside a
        real failure in the other phase would be a lie.
    #>
    param([string] $Message)

    $meanings = [System.Collections.Generic.List[string]]::new()
    if ($Message) {
        foreach ($hit in [regex]::Matches($Message, '(?i)(?:errc|_error|"val")\D{0,4}(\d+)')) {
            $code = $hit.Groups[1].Value
            if ($TeamsVdiErrorCodes.ContainsKey($code)) {
                $meanings.Add("code $code - $($TeamsVdiErrorCodes[$code])")
            }
        }
    }
    return ($meanings | Select-Object -Unique)
}

function Write-TeamsVdiEventStatus {
    <#
        What this machine itself logged about the media optimization. On a session
        host it is the only first-hand evidence that users really are optimized -
        every other check only proves the parts are in place.
    #>
    param([int] $Days = 7)

    $events = @(Get-TeamsVdiEvent -Days $Days)
    if ($events.Count -eq 0) {
        Write-Skip "No '$TeamsVdiEventSource' events in the last $Days days - no optimized session was logged here"
        return
    }

    $last     = $events | Sort-Object TimeCreated -Descending | Select-Object -First 1
    $failures = @($events | Where-Object { $_.LevelDisplayName -in @('Error', 'Critical') })
    Write-Ok ('Teams VDI logged {0} event(s) in the last {1} days, last one {2:yyyy-MM-dd HH:mm}' -f
              $events.Count, $Days, $last.TimeCreated)

    foreach ($entry in ($failures | Sort-Object TimeCreated -Descending | Select-Object -First 3)) {
        $text = (($entry.Message -replace '\s+', ' ')).Trim()
        if ($text.Length -gt 200) { $text = $text.Substring(0, 200) + '...' }
        Write-Warn ('Teams VDI error {0:yyyy-MM-dd HH:mm}: {1}' -f $entry.TimeCreated, $text)
        foreach ($meaning in Get-TeamsVdiErrorMeaning -Message $entry.Message) { Write-Warn "         $meaning" }
    }

    # A success code in the newest event says as much as the error count does: it is
    # the difference between "nobody connected" and "everybody is optimized".
    if ($failures.Count -eq 0) {
        foreach ($meaning in Get-TeamsVdiErrorMeaning -Message $last.Message) { Write-Ok "Last event: $meaning" }
    }
}

$AppLockerPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\SrpV2'

# What in a rule makes it relevant to SlimCore: the packages by name, or a publisher
# rule broad enough to cover them (they are signed by Microsoft like everything else).
$SlimCoreRulePattern         = '(?i)SlimCore|Microsoft\.Teams'
$MicrosoftPublisherPattern   = '(?i)O=MICROSOFT CORPORATION'

function Get-AppLockerDetail {
    <#
        What AppLocker actually says, instead of only whether it exists. Three things
        decide whether it can stop the SlimCore MSIX, and each is easy to get wrong:

          - only the packaged-app collection (Appx) applies to an MSIX at all. An
            enforced Exe or Msi collection has nothing to do with it;
          - "not configured" does not mean harmless. Microsoft: a collection that
            holds at least one rule and has enforcement not configured is enforced.
            Only an explicit 0 (audit only) lets everything through;
          - nothing is enforced at all while the Application Identity service is not
            running, which is worth saying out loud rather than assuming either way.

        Returns $null when AppLocker is not configured on this machine. -Root exists
        so the reader can be pointed at a rebuilt policy in a test.
    #>
    param([string] $Root = $AppLockerPolicyPath)

    if (-not (Test-Path $Root)) { return $null }

    $collections = [System.Collections.Generic.List[object]]::new()
    foreach ($key in @(Get-ChildItem -Path $Root -ErrorAction SilentlyContinue)) {
        $mode  = Get-PropertyValue (Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue) 'EnforcementMode'
        $rules = [System.Collections.Generic.List[object]]::new()

        foreach ($ruleKey in @(Get-ChildItem -Path $key.PSPath -ErrorAction SilentlyContinue)) {
            $xml = Get-PropertyValue (Get-ItemProperty -Path $ruleKey.PSPath -ErrorAction SilentlyContinue) 'Value'
            if (-not $xml) { continue }
            $name   = if ($xml -match 'Name="([^"]*)"')   { $Matches[1] } else { $ruleKey.PSChildName }
            $action = if ($xml -match 'Action="([^"]*)"') { $Matches[1] } else { 'unknown' }
            $rules.Add([PSCustomObject]@{ Name = $name; Action = $action; Xml = [string] $xml })
        }

        $enforced = if ($mode -eq 1) { $true } elseif ($mode -eq 0) { $false } else { $rules.Count -gt 0 }
        $label    = if ($mode -eq 1) { 'enforced' } elseif ($mode -eq 0) { 'audit only' }
                    elseif ($rules.Count -gt 0) { 'enforced (enforcement not configured, which still enforces)' }
                    else { 'no rules' }

        $collections.Add([PSCustomObject]@{
            Name      = $key.PSChildName
            RuleCount = $rules.Count
            Rules     = @($rules)
            Enforced  = $enforced
            Mode      = $label
            Path      = "$Root\$($key.PSChildName)"
        })
    }

    $service = Get-Service -Name 'AppIDSvc' -ErrorAction SilentlyContinue
    return [PSCustomObject]@{
        Path          = $Root
        ServiceStatus = if ($service) { [string] $service.Status } else { 'not installed' }
        Enforcing     = [bool] ($service -and $service.Status -eq 'Running')
        Collections   = @($collections)
    }
}

function Get-AppLockerAppxRule {
    <# The packaged-app collection, the only one an MSIX ever meets. #>
    param($Detail)
    if (-not $Detail) { return $null }
    return @($Detail.Collections | Where-Object { $_.Name -eq 'Appx' }) | Select-Object -First 1
}

function Write-AppLockerStatus {
    <#
        Where the policy lives and what it says. "AppLocker policy is present" is not
        something a technician can act on; a registry path, an enforcement mode per
        collection and the rule names are.
    #>
    param([Parameter(Mandatory)] $Detail)

    # The key exists with no collections under it on plenty of machines - a leftover
    # from a policy that was removed. Saying "-" there reads like a broken script.
    $summary = @(foreach ($collection in ($Detail.Collections | Sort-Object Name)) {
        '{0} {1}' -f $collection.Name, $collection.Mode
    })
    $what = if ($summary.Count -gt 0) { $summary -join '; ' } else { 'no rule collections configured, so it blocks nothing' }
    Write-Skip "AppLocker policy: $($Detail.Path) - $what (Application Identity service: $($Detail.ServiceStatus))"

    $appx = Get-AppLockerAppxRule -Detail $Detail
    if (-not $appx -or $appx.RuleCount -eq 0) { return }

    # Relevant rules first, then denies, then the rest: truncating the one rule that
    # decides this away in favour of filler would defeat the point of printing them.
    $ordered = $appx.Rules | Sort-Object -Property @{ Expression = {
        if ($_.Xml -match $SlimCoreRulePattern -or $_.Xml -match $MicrosoftPublisherPattern) { 0 }
        elseif ($_.Action -eq 'Deny') { 1 } else { 2 }
    } }, Name

    foreach ($rule in ($ordered | Select-Object -First 5)) {
        Write-Skip "  Appx rule: $($rule.Action) - $($rule.Name)"
    }
    if ($appx.RuleCount -gt 5) { Write-Skip "  ... and $($appx.RuleCount - 5) more Appx rule(s)" }

    if (-not $appx.Enforced) { return }
    if (@($appx.Rules | Where-Object { $_.Action -eq 'Allow' -and $_.Xml -match $SlimCoreRulePattern })) {
        Write-Ok 'An AppLocker rule already allows the SlimCore packages by name'
    } elseif (@($appx.Rules | Where-Object { $_.Action -eq 'Allow' -and $_.Xml -match $MicrosoftPublisherPattern })) {
        Write-Skip '  Rules allowing anything signed by Microsoft Corporation should cover the SlimCore packages - check they are not narrowed to one product name'
    }
}

function Get-SlimCoreBlocker {
    <#
        Registry policies Microsoft documents as stopping the SlimCore MSIX from
        staging, with the error code Teams reports for each. They apply wherever the
        staging happens, which is the endpoint.
    #>
    param($AppLocker)

    $blockers = [System.Collections.Generic.List[string]]::new()

    foreach ($path in 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\ApplicationManagement',
                      'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Appx') {
        $policy = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
        if ((Get-PropertyValue $policy 'BlockNonAdminUserInstall') -eq 1) {
            $blockers.Add('BlockNonAdminUserInstall is 1 - a non-admin cannot register the SlimCore packages (Teams reports error 16389)')
        }
        if ((Get-PropertyValue $policy 'AllowAllTrustedApps') -eq 0) {
            $blockers.Add('AllowAllTrustedApps is 0 - sideloading is off, so the SlimCore MSIX cannot install (Teams reports error 15615)')
        }
    }

    # Only the packaged-app collection, only when it is actually enforced, and only
    # when nothing in it lets the packages through. Warning on the mere presence of
    # an AppLocker policy cried wolf at every Exe rule ever written.
    $appx = Get-AppLockerAppxRule -Detail $AppLocker
    if ($appx -and $appx.Enforced) {
        $allows = @($appx.Rules | Where-Object {
            $_.Action -eq 'Allow' -and ($_.Xml -match $SlimCoreRulePattern -or $_.Xml -match $MicrosoftPublisherPattern)
        })
        if ($allows.Count -eq 0) {
            $note = if ($AppLocker.Enforcing) { '' }
                    else { " (the Application Identity service is $($AppLocker.ServiceStatus), so nothing is enforced at this moment)" }
            $blockers.Add(("AppLocker enforces packaged apps at {0} with {1} rule(s), none of which allows the SlimCoreVdi packages - Teams reports error 10083{2}" -f
                           $appx.Path, $appx.RuleCount, $note))
        }
    }

    return ($blockers | Select-Object -Unique)
}

function Write-SlimCoreStatus {
    <#
        Where SlimCore stands, from the point of view of the machine this runs on.

        On a session host the packages are never expected: the plugin stages them on
        the endpoint. All the host contributes is a recent enough Teams build. On an
        endpoint the packages and the blocking policies are the real story.
    #>
    param([version] $TeamsVersion)

    $appLocker = Get-AppLockerDetail
    $blockers  = @(Get-SlimCoreBlocker -AppLocker $appLocker)

    if (Test-AvdSessionHost) {
        if ($TeamsVersion -and $TeamsVersion -lt $SlimCoreMinimumTeamsBuild) {
            Write-Warn "Teams $TeamsVersion is below $SlimCoreMinimumTeamsBuild, the minimum build that can use SlimCore"
        } else {
            Write-Ok "Teams build supports SlimCore - whether it is used depends on the endpoint's Windows App ($SlimCoreMinimumWindowsApp or newer); SlimCore itself is staged there, never here"
        }

        # Same reason the packages are not looked for here: the staging happens on the
        # endpoint, so a policy on this host is not what blocks it. It is usually the
        # same GPO though, which is why it is named rather than hidden.
        foreach ($blocker in $blockers) {
            Write-Skip "Policy on this host, for reference - it blocks nothing here because SlimCore stages on the endpoint, but check whether the same policy is linked there: $blocker"
        }
    } else {
        $staged = @(Get-AppxPackage -AllUsers -Name 'Microsoft.Teams.SlimCoreVdiHost*' -ErrorAction SilentlyContinue) |
                  Select-Object -First 1
        if (-not $staged) {
            $staged = @(Get-AppxPackage -Name 'Microsoft.Teams.SlimCoreVdiHost*' -ErrorAction SilentlyContinue) |
                      Select-Object -First 1
        }
        if ($staged) {
            Write-Ok "SlimCore is staged on this endpoint ($($staged.Version))"
        } else {
            Write-Skip 'SlimCore is not staged on this endpoint yet - the plugin downloads it on the first optimized connection'
        }

        foreach ($blocker in $blockers) { Write-Warn "SlimCore blocker: $blocker" }
    }

    if ($appLocker) { Write-AppLockerStatus -Detail $appLocker }
}

function Save-VerifiedDownload {
    <#
        Download a Microsoft installer and refuse anything that is not what it claims
        to be: a proxy error page, a truncated file, or a binary that is not signed by
        Microsoft. Throws on all three.
    #>
    param(
        [Parameter(Mandatory)] [string] $Uri,
        [Parameter(Mandatory)] [string] $Path,
        [int] $MinimumBytes = 100KB
    )

    if ($Uri -notmatch '^https://') { throw "Download URL must be https: $Uri" }

    # PowerShell 5.1 on older builds still defaults to TLS 1.0, which the CDN refuses.
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $progressBackup     = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    try {
        Invoke-WebRequest -Uri $Uri -OutFile $Path -UseBasicParsing
    } catch {
        throw "Download failed ($Uri): $($_.Exception.Message)"
    } finally {
        $ProgressPreference = $progressBackup
    }

    $file = Get-Item $Path
    if ($file.Length -lt $MinimumBytes) {
        throw "Downloaded file is only $($file.Length) bytes - not a valid installer ($Uri)"
    }

    if ($SkipSignatureCheck) {
        Write-Warn "Signature check skipped for $($file.Name) (-SkipSignatureCheck)"
    } else {
        $signature = Get-AuthenticodeSignature -FilePath $Path
        if ($signature.Status -ne 'Valid') {
            throw "$($file.Name) signature is $($signature.Status) - refusing to run it"
        }
        if ($signature.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
            throw "$($file.Name) is not signed by Microsoft: $($signature.SignerCertificate.Subject)"
        }
    }
    return $file
}

function Get-LatestTeamsBuild {
    <#
        The build Microsoft currently publishes for this architecture, from the same
        config service the Teams client uses to decide it is out of date. Returns
        $null when the service cannot be reached or has no build for this platform.
    #>
    param([string] $UpdateRing = 'general')

    $arch = switch ($env:PROCESSOR_ARCHITECTURE) {
        'ARM64' { 'arm64' }
        'AMD64' { 'x64' }
        default { if ($env:PROCESSOR_ARCHITEW6432 -eq 'AMD64') { 'x64' } else { 'x86' } }
    }

    $uri = 'https://config.teams.microsoft.com/config/v1/MicrosoftTeams/0.0.0.0' +
           "?environment=prod&audienceGroup=$UpdateRing&teamsRing=$UpdateRing&agent=TeamsBuilds"
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $config = Invoke-RestMethod -Uri $uri -UseBasicParsing -TimeoutSec 30
    } catch {
        Write-Warn "Could not reach the Teams config service: $($_.Exception.Message)"
        return $null
    }

    $builds = Get-PropertyValue $config 'BuildSettings'
    # WebView2PreAuth carries the Windows builds; WebView2 is the older key and still
    # holds macOS, so both are checked before giving up.
    foreach ($channel in 'WebView2PreAuth', 'WebView2') {
        $node = Get-PropertyValue (Get-PropertyValue $builds $channel) $arch
        $latest = Get-PropertyValue $node 'latestVersion'
        if ($latest) {
            return [PSCustomObject]@{
                Version      = [version] $latest
                Link         = Get-PropertyValue $node 'buildLink'
                Architecture = $arch
                Channel      = $channel
            }
        }
    }

    Write-Warn "The Teams config service returned no build for $arch."
    return $null
}

function Get-TeamsAddInInstaller {
    <#
        The meeting add-in MSI inside the staged new Teams package, newest first.

        Deliberately found by looking in WindowsApps rather than by asking
        Get-AppxPackage for a version: teamsbootstrapper.exe -p *provisions* the
        package, which stages it for future sign-ins without installing it for
        whoever ran the script. On a session host Get-AppxPackage then returns
        nothing and the add-in step fails even though the install worked.
    #>
    $pattern = Join-Path $env:ProgramFiles 'WindowsApps\MSTeams_*_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi'

    return Get-Item -Path $pattern -ErrorAction SilentlyContinue |
           Sort-Object -Property @{ Expression = {
               $raw = $_.Directory.Name -replace '^MSTeams_', '' -replace '_x64__8wekyb3d8bbwe$', ''
               try { [version] $raw } catch { [version] '0.0.0.0' }
           } } |
           Select-Object -Last 1
}

function Get-MsiCachedPackage {
    <#
        Windows Installer's own cached copy of an installed product's MSI, under
        C:\Windows\Installer. It needs that file to uninstall; when it has gone
        missing msiexec /x answers 1612 ("the installation source is not available")
        and the product stays registered for good - which is what later makes an
        install of a different version answer 1638.
    #>
    param([Parameter(Mandatory)] [string] $DisplayNameLike)

    $root = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products'
    if (-not (Test-Path $root)) { return }

    foreach ($key in @(Get-ChildItem "$root\*\InstallProperties" -ErrorAction SilentlyContinue)) {
        $name = $key.GetValue('DisplayName')
        if ($name -notlike $DisplayNameLike) { continue }
        $package = $key.GetValue('LocalPackage')
        [PSCustomObject]@{
            DisplayName  = $name
            Version      = $key.GetValue('DisplayVersion')
            LocalPackage = $package
            Cached       = [bool] ($package -and (Test-Path $package))
        }
    }
}

function Get-MsiProductVersion {
    <#
        ProductVersion straight from the MSI property table. Microsoft's own sample
        reads it with Get-AppLockerFileInformation, but that module is absent on some
        editions and under PowerShell 7 it drags in the Windows PowerShell
        compatibility layer, which fails and floods a -WhatIf run with its own output.
    #>
    param([Parameter(Mandatory)] [string] $Path)

    $installer = $null; $database = $null; $view = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $database  = $installer.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $installer, @($Path, 0))
        $view      = $database.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $database, @("SELECT Value FROM Property WHERE Property = 'ProductVersion'"))
        $null      = $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, @($null))
        $record    = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
        if (-not $record) { return $null }
        return $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, @(1))
    } catch {
        return $null
    } finally {
        foreach ($com in @($view, $database, $installer)) {
            if ($com) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($com) }
        }
    }
}

function Invoke-Installer {
    <#
        Run an installer and never hang the caller: the process is killed when it
        outlives the timeout, 1618 (another install in progress) is retried, and
        3010 counts as success with a reboot pending.
    #>
    param(
        [Parameter(Mandatory)] [string] $FilePath,
        [string] $Arguments = '',
        [int]    $Attempts  = 3
    )

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        $startArgs = @{ FilePath = $FilePath; PassThru = $true; WindowStyle = 'Hidden' }
        if ($Arguments) { $startArgs['ArgumentList'] = $Arguments }

        $proc = Start-Process @startArgs
        if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
            try { $proc.Kill() } catch { Write-Warn "Could not kill PID $($proc.Id): $($_.Exception.Message)" }
            return [PSCustomObject]@{ Success = $false; ExitCode = $null; RebootRequired = $false
                                      Message = "timed out after $TimeoutSeconds seconds and was killed" }
        }

        $code = $proc.ExitCode
        if ($code -eq 1618 -and $attempt -lt $Attempts) {
            Write-Warn "Another installation is in progress (1618) - retrying in 30 seconds ($attempt/$($Attempts - 1))"
            Start-Sleep -Seconds 30
            continue
        }

        # 3010 = reboot required, 1641 = reboot already initiated by the installer.
        # Both are successes that happen to need a restart; failing on them would
        # abort a run that actually worked.
        return [PSCustomObject]@{
            Success        = ($code -in @(0, 3010, 1641))
            ExitCode       = $code
            RebootRequired = ($code -in @(3010, 1641))
            Message        = "exit code $code"
        }
    }
}

# -- Run -----------------------------------------------------------------------
try {
    Write-Out ''
    Write-Out "  Mode: $(if ($CheckOnly) { 'CHECK ONLY - reporting, never installing' } elseif ($simulate) { '-WhatIf - nothing will be changed' } else { 'APPLY - Teams is updated when a newer build exists' })" 'Cyan'
    Write-Out ''

    # -- 1. Preflight ----------------------------------------------------------
    Write-Step '1. Preflight'
    try {
        $existingTeams = @(Get-AppxPackage -AllUsers -Name '*MSTEAMS*' -ErrorAction SilentlyContinue)
    } catch {
        throw "Could not enumerate AppX packages ($($_.Exception.Message)). This needs administrator or System rights."
    }

    $installedVersion = $null
    foreach ($pkg in $existingTeams) {
        $users = $null
        try { $users = @($pkg.PackageUserInformation).Count } catch { $users = $null }
        $forWhom = if ($users) { " - installed for $users user profile(s)" } else { '' }
        Write-Ok "New Teams (AppX) $($pkg.Name) $($pkg.Version)$forWhom"

        $candidate = [version] $pkg.Version
        if ($null -eq $installedVersion -or $candidate -gt $installedVersion) { $installedVersion = $candidate }
    }

    # A pooled session host or a fresh image often has Teams provisioned without any
    # user having it installed yet. Without this fallback the version check has
    # nothing to compare, calls the host outdated and reinstalls on every single run.
    $provisionedVersion = $null
    try {
        $provisionedPkg = @(Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -eq 'MSTeams' }) |
                          Select-Object -First 1
        if ($provisionedPkg) { $provisionedVersion = [version] $provisionedPkg.Version }
    } catch {
        Write-Warn "Could not read the provisioned packages: $($_.Exception.Message)"
    }

    if ($null -eq $installedVersion -and $provisionedVersion) {
        $installedVersion = $provisionedVersion
        Write-Ok "Provisioned MSTeams $provisionedVersion (not installed for any user yet - normal on a session host or a fresh image)"
    }

    if ($null -eq $installedVersion) {
        if ($Force) { Write-Skip 'No Teams installation detected - continuing because -Force was given' }
        else        { throw 'No Teams installation detected. Use -Force to install new Teams anyway.' }
    }

    # Everywhere else Teams can still live. Classic Teams is only removed when asked
    # for: it shares the October 2026 end-of-support date and a leftover machine-wide
    # installer keeps restaging it into new profiles, but taking an app away from a
    # user is not something an update job should decide by itself.
    $classicMachineWide = @(Get-ClassicTeamsEntry)
    $classicUserInstall = @(Get-ClassicTeamsUserInstall)
    $classicNote = if ($RemoveClassicTeams) { 'will be removed' } else { 'left alone - use -RemoveClassicTeams' }

    foreach ($classic in $classicMachineWide) {
        Write-Warn "Classic Teams machine-wide installer present ($($classic.Version)) - $classicNote"
    }
    foreach ($classic in $classicUserInstall) {
        Write-Warn "Classic Teams installed for $($classic.Account) ($($classic.Version)) - $classicNote"
    }

    # The registered version is kept: it is what decides whether the add-in inside
    # the package about to be staged can be installed over it at all.
    $addInEntries          = @(Get-TeamsMeetingAddInEntry)
    $addInInstalled        = $addInEntries.Count -gt 0
    $installedAddInVersion = $null
    foreach ($entry in $addInEntries) {
        try { $candidate = [version] $entry.Version } catch { continue }
        if ($null -eq $installedAddInVersion -or $candidate -gt $installedAddInVersion) {
            $installedAddInVersion = $candidate
        }
    }
    $machineWideAddInOk     = $false
    $machineWideAddInBroken = $false
    $brokenAddInUsers       = @()

    if (-not $SkipMeetingAddIn) {
        if ($addInInstalled) {
            foreach ($entry in $addInEntries) { Write-Ok "Teams Meeting Add-in is installed ($($entry.Version))" }
        } else {
            Write-Warn 'Teams Meeting Add-in is not installed'
        }

        $addInRegistrations = @(Get-OutlookAddInRegistration)
        $machineWideAddInOk = [bool] ($addInRegistrations | Where-Object { $_.Account -like 'all users*' -and -not $_.DllMissing })
        # Registered but pointing at files that are gone is not "installed". Counting
        # it as installed is how a device whose add-in was deleted gets told there is
        # nothing to do.
        $machineWideAddInBroken = [bool] ($addInRegistrations | Where-Object { $_.Account -like 'all users*' -and $_.DllMissing })
        $brokenAddInUsers   = @($addInRegistrations | Where-Object { $_.Sid -and $_.DllMissing })
        Write-OutlookAddInStatus -Registrations $addInRegistrations
    }

    # SlimCore is reported separately from -AvdOptimizations: it is an endpoint-side
    # story, and on an endpoint nobody passes that switch.
    if ($AvdOptimizations -or (Test-AvdSessionHost) -or
        (Get-AppxPackage -Name 'Microsoft.Teams.SlimCoreVdiHost*' -ErrorAction SilentlyContinue)) {
        Write-SlimCoreStatus -TeamsVersion $installedVersion
    }

    # Read the optimization state whether or not this run may change it: on a session
    # host it is half the answer to "is Teams healthy here".
    $avdFlagSet  = ((Get-WvdEnvironmentFlag) -eq 1)
    $webRtcEntry = @(Get-WebRtcRedirectorEntry) | Select-Object -First 1
    $sessionHost = Test-AvdSessionHost

    if ($AvdOptimizations -or $RemoveWebRtcRedirector -or $sessionHost) {
        if ($avdFlagSet) { Write-Ok 'AVD media flag IsWVDEnvironment is set' }
        else             { Write-Warn 'AVD media flag IsWVDEnvironment is not set - neither WebRTC nor SlimCore optimizes a session host without it' }

        if ($webRtcEntry)          { Write-Ok "WebRTC Redirector is installed ($($webRtcEntry.Version)) - the old optimization, unsupported after 1 October 2026" }
        elseif ($AvdOptimizations) { Write-Warn 'Remote Desktop WebRTC Redirector is not installed' }
        else                       { Write-Skip 'Remote Desktop WebRTC Redirector is not installed - optimization depends on SlimCore from the endpoint' }

        Write-TeamsVdiEventStatus
    }

    if ($sessionHost -and -not ($AvdOptimizations -or $RemoveWebRtcRedirector)) {
        Write-Skip 'This looks like an AVD session host - -AvdOptimizations sets the media flag and installs WebRTC, -RemoveWebRtcRedirector drops the old stack'
    }

    # -- 2. Version check ------------------------------------------------------
    Write-Out ''
    Write-Step '2. Version check'
    $latest = Get-LatestTeamsBuild -UpdateRing $Ring

    if ($latest) {
        Write-Ok "Latest published build ($($latest.Architecture), ring $Ring): $($latest.Version)"
    } elseif ($Force) {
        Write-Skip 'No version information - continuing because -Force was given'
    } else {
        throw 'Could not determine the latest published build. Use -Force to reinstall without the check.'
    }

    $clientOutdated = $false
    if ($null -eq $installedVersion) {
        $clientOutdated = $true
    } elseif ($latest) {
        if ($installedVersion -lt $latest.Version) {
            $clientOutdated = $true
            Write-News "Update available: $installedVersion -> $($latest.Version)"
        } elseif ($installedVersion -gt $latest.Version) {
            Write-Ok "Installed build is newer than the published one ($installedVersion) - nothing to update"
            if ($Force) {
                Write-Warn "-Force replaces it with the published $($latest.Version), which is a downgrade - the meeting add-in inside that older package can be older than the $installedAddInVersion registered now, and Windows Installer refuses that with 1638"
            }
        } else {
            Write-Ok "Installed build is current ($installedVersion)"
        }
    }

    $addInMissing  = (-not $SkipMeetingAddIn) -and ((-not $addInInstalled) -or $machineWideAddInBroken)
    $avdWork       = $AvdOptimizations -and ((-not $avdFlagSet) -or (-not $webRtcEntry) -or $Force)
    $webRtcRemoval = $RemoveWebRtcRedirector -and [bool] $webRtcEntry
    $classicWork   = $RemoveClassicTeams -and (($classicMachineWide.Count + $classicUserInstall.Count) -gt 0)
    $repairWork    = $RepairOutlookAddIn -and $brokenAddInUsers.Count -gt 0
    $fullReinstall = $clientOutdated -or $Force

    if (-not $fullReinstall -and -not $addInMissing -and -not $avdWork -and -not $classicWork -and
        -not $repairWork -and -not $webRtcRemoval) {
        $plannedExit = 0
        throw 'Teams is up to date - nothing to do.'
    }

    $reasons = @()
    if ($clientOutdated) { $reasons += 'a newer build is available' }
    if ($addInMissing) {
        $reasons += if ($machineWideAddInBroken) { 'the machine-wide add-in registration points at files that are gone' }
                    else { 'the Teams Meeting Add-in is missing' }
    }
    if ($avdWork)        { $reasons += 'the AVD optimizations are incomplete' }
    if ($webRtcRemoval)  { $reasons += 'the old WebRTC optimization is still installed' }
    if ($classicWork)    { $reasons += 'classic Teams is still installed' }
    if ($repairWork)     { $reasons += 'a user registration points at a removed add-in DLL' }
    if ($Force)          { $reasons += '-Force was given' }

    if ($CheckOnly) {
        Show-HeldOutput
        Write-News ('Work is due: {0} (exit code 2)' -f ($reasons -join ', '))
        $plannedExit = 2
        throw 'Check only - nothing was changed.'
    }

    if (-not $fullReinstall) {
        $parts = @()
        if ($avdWork)      { $parts += 'the AVD optimizations' }
        if ($classicWork)  { $parts += 'the classic Teams removal' }
        if ($repairWork)   { $parts += 'the per-user registration repair' }
        if ($addInMissing) { $parts += 'the meeting add-in' }
        Write-Skip ('Client is current - only {0} will be handled' -f ($parts -join ' and '))
    }

    # From here the run intends to change something, so it is worth reporting and
    # worth a transcript.
    Show-HeldOutput
    if (-not $simulate) {
        if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }
        $logFile = Join-Path $LogPath ("Update-TeamsClient_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        Start-Transcript -Path $logFile | Out-Null
        $transcribing = $true
    }

    if ($fullReinstall -and (Get-Process -Name 'ms-teams', 'Teams' -ErrorAction SilentlyContinue)) {
        Write-Warn 'Teams is running - it will be closed by the removal'
    }
    if (-not $SkipMeetingAddIn -and (Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)) {
        Write-Warn 'Outlook is running - the meeting add-in registers reliably only after Outlook restarts'
    }

    # Last exit before anything is touched. Only for a hands-on run: an RMM or
    # scheduled session is not interactive, and -Confirm:$false silences it outright.
    if (-not $simulate -and -not $confirmSuppressed -and [Environment]::UserInteractive) {
        Write-Out ''
        Write-Warn 'Teams and the meeting add-in will be uninstalled and reinstalled on this device.'
        $answer = Read-Host '  Continue? [y/N]'
        if ($answer -notmatch '^[Yy]') {
            $plannedExit = 0
            throw 'Cancelled - nothing was changed.'
        }
    }

    # -- 3. AVD / VDI optimizations --------------------------------------------
    # Before the client is provisioned: Teams reads IsWVDEnvironment at startup to
    # decide whether to hand media off to the redirector.
    Write-Out ''
    Write-Step '3. AVD optimizations'
    if ($RemoveWebRtcRedirector) {
        # Taking the old stack away only leaves users optimized when their endpoint
        # can do SlimCore. Anyone on an older Windows App falls back to server-side
        # rendering on this host, so this is deliberately never the default.
        if (-not $webRtcEntry) {
            Write-Skip 'WebRTC Redirector is not installed - nothing to remove'
        } elseif ($PSCmdlet.ShouldProcess("Remote Desktop WebRTC Redirector Service $($webRtcEntry.Version)",
                                          'msiexec /x /qn (remove the old optimization)')) {
            $removal = Invoke-Installer -FilePath 'msiexec.exe' -Arguments "/x $($webRtcEntry.ProductCode) /qn /norestart"
            if ($removal.ExitCode -eq 1605) {
                # Same as the classic Teams case: msiexec never heard of the product,
                # so what is left is an orphaned uninstall entry, not an install.
                Write-Warn 'msiexec does not know this product (1605) - the uninstall entry is stale, removing it'
                Remove-Item -Path $webRtcEntry.RegistryPath -Recurse -Force -ErrorAction SilentlyContinue
            } elseif (-not $removal.Success) {
                throw "Could not remove the WebRTC Redirector ($($removal.Message))"
            } else {
                if ($removal.RebootRequired) { $rebootRequired = $true }
                Write-Ok "WebRTC Redirector $($webRtcEntry.Version) removed - media now depends on SlimCore from the endpoint"
            }
        }
        if ($avdFlagSet) { Write-Skip 'IsWVDEnvironment left at 1 - SlimCore needs that flag just as much' }
    } elseif (-not $AvdOptimizations) {
        Write-Skip 'Skipped - use -AvdOptimizations on an AVD/VDI session host'
    } else {
        if ($WebRtcUrl -notmatch '^https://') { throw "WebRtcUrl must be https: $WebRtcUrl" }

        if ($avdFlagSet) {
            Write-Skip 'IsWVDEnvironment is already 1'
        } elseif ($PSCmdlet.ShouldProcess("$AvdRegistryPath\IsWVDEnvironment", 'Set to 1 (DWORD)')) {
            if (-not (Test-Path $AvdRegistryPath)) { New-Item -Path $AvdRegistryPath -Force | Out-Null }
            New-ItemProperty -Path $AvdRegistryPath -Name 'IsWVDEnvironment' -PropertyType DWORD -Value 1 -Force | Out-Null
            Write-Ok 'IsWVDEnvironment set to 1'
        }

        if ($webRtcEntry -and -not $Force) {
            Write-Skip "WebRTC Redirector already installed ($($webRtcEntry.Version))"
        } else {
            if (-not $workingDirReady) {
                if ($PSCmdlet.ShouldProcess($WorkingDir, 'Create working directory')) {
                    New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
                    Write-Ok "Working directory ready: $WorkingDir"
                }
                $workingDirReady = $true
            }

            if ($PSCmdlet.ShouldProcess($webRtcMsi, "Download $WebRtcUrl")) {
                $file = Save-VerifiedDownload -Uri $WebRtcUrl -Path $webRtcMsi
                Write-Ok "Downloaded $($file.Name) ($([math]::Round($file.Length / 1MB, 1)) MB), signature verified"
            }

            # The redirector MSI keeps its ProductCode across versions, so msiexec /i
            # over an existing install returns 1638 ("another version of this product
            # is already installed") instead of upgrading. Same version: repair in
            # place. Different version: take the old one off first.
            $msiVersion = if (Test-Path $webRtcMsi) { Get-MsiProductVersion -Path $webRtcMsi } else { $null }
            $msiArgs    = '/i "{0}" /qn /norestart' -f $webRtcMsi

            if ($webRtcEntry -and -not $msiVersion) {
                Write-Skip 'On a real run the downloaded version decides whether the installed redirector is removed first'
            } elseif ($webRtcEntry -and $msiVersion) {
                if ($webRtcEntry.Version -eq $msiVersion) {
                    Write-Skip "The download is the installed version ($msiVersion) - repairing in place"
                    $msiArgs = '/i "{0}" REINSTALL=ALL REINSTALLMODE=vomus /qn /norestart' -f $webRtcMsi
                } else {
                    Write-Ok "Replacing WebRTC Redirector $($webRtcEntry.Version) with $msiVersion"
                    if ($PSCmdlet.ShouldProcess("Remote Desktop WebRTC Redirector Service $($webRtcEntry.Version)", 'msiexec /x /qn (uninstall the old version)')) {
                        $removal = Invoke-Installer -FilePath 'msiexec.exe' -Arguments "/x $($webRtcEntry.ProductCode) /qn /norestart"
                        if (-not $removal.Success) {
                            throw "Could not remove WebRTC Redirector $($webRtcEntry.Version) ($($removal.Message))"
                        }
                        if ($removal.RebootRequired) { $rebootRequired = $true }
                    }
                }
            }

            if ($PSCmdlet.ShouldProcess('Remote Desktop WebRTC Redirector Service', "msiexec.exe $msiArgs")) {
                $result = Invoke-Installer -FilePath 'msiexec.exe' -Arguments $msiArgs
                if ($result.ExitCode -eq 1638) {
                    Write-Warn 'msiexec reports another version of the redirector is already installed (1638) - leaving the existing one in place'
                } elseif (-not $result.Success) {
                    throw "WebRTC Redirector install failed ($($result.Message))"
                } else {
                    if ($result.RebootRequired) { $rebootRequired = $true }
                    Write-Ok 'WebRTC Redirector installed'
                }
            }
        }
    }

    # -- 4. Remove classic Teams -----------------------------------------------
    Write-Out ''
    Write-Step '4. Classic Teams'
    if (-not $RemoveClassicTeams) {
        if ($classicMachineWide.Count + $classicUserInstall.Count -gt 0) {
            Write-Skip 'Found but left alone - use -RemoveClassicTeams to remove it'
        } else {
            Write-Skip 'Not installed'
        }
    } else {
        if ($classicMachineWide.Count + $classicUserInstall.Count -eq 0) {
            Write-Skip 'Nothing to remove'
        }

        if (Get-Process -Name 'Teams' -ErrorAction SilentlyContinue) {
            Write-Warn 'Classic Teams is running - its files are locked, so a per-user removal may only finish after the user signs out'
        }

        foreach ($entry in $classicMachineWide) {
            $target = "$($entry.DisplayName) $($entry.Version) [$($entry.ProductCode)]"
            if ($PSCmdlet.ShouldProcess($target, 'msiexec /x /qn (uninstall)')) {
                $result = Invoke-Installer -FilePath 'msiexec.exe' -Arguments "/x $($entry.ProductCode) /qn /norestart"

                if ($result.Success) {
                    Write-Ok "Uninstalled $($entry.DisplayName)"
                    if ($result.RebootRequired) { $rebootRequired = $true }
                } elseif ($result.ExitCode -eq 1605) {
                    # 1605 is "this action is only valid for products that are
                    # currently installed": the entry in Programs and Features
                    # outlived the product itself, which is common once the new Teams
                    # bootstrapper has been over the machine. Nothing to uninstall,
                    # but the stale entry keeps this script reporting classic Teams
                    # forever, so it goes too.
                    Write-Warn "$($entry.DisplayName) is no longer registered with Windows Installer (1605) - the entry in Programs and Features is stale"
                    if ($PSCmdlet.ShouldProcess($entry.RegistryPath, 'Remove the stale uninstall entry')) {
                        Remove-Item -Path $entry.RegistryPath -Recurse -Force -ErrorAction SilentlyContinue
                        if (Test-Path $entry.RegistryPath) {
                            Write-Warn "Could not remove the stale entry at $($entry.RegistryPath)"
                        } else {
                            Write-Ok 'Removed the stale uninstall entry'
                        }
                    }
                } else {
                    throw "Uninstall of $($entry.DisplayName) failed ($($result.Message))"
                }
            }
        }

        # The documented per-user uninstall is Update.exe --uninstall -s, but that
        # has to run as the user who owns the profile. From System the install root,
        # the autostart entry and the stale uninstall key are removed instead.
        foreach ($user in $classicUserInstall) {
            if ($PSCmdlet.ShouldProcess("$($user.Account): $($user.InstallRoot)", 'Remove classic Teams install folder')) {
                Remove-Item -LiteralPath $user.InstallRoot -Recurse -Force -ErrorAction SilentlyContinue
                if (Test-Path $user.InstallRoot) {
                    Write-Warn "Could not fully remove $($user.InstallRoot) - files in use; it will finish on a later run"
                } else {
                    Write-Ok "Removed classic Teams for $($user.Account)"
                }
            }

            $hive   = "Registry::HKEY_USERS\$($user.Sid)"
            $runKey = "$hive\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
            if ((Test-Path $runKey) -and (Get-PropertyValue (Get-ItemProperty $runKey -ErrorAction SilentlyContinue) 'com.squirrel.Teams.Teams')) {
                if ($PSCmdlet.ShouldProcess("$($user.Account): Run\com.squirrel.Teams.Teams", 'Remove autostart entry')) {
                    Remove-ItemProperty -Path $runKey -Name 'com.squirrel.Teams.Teams' -Force -ErrorAction SilentlyContinue
                    Write-Ok "Removed the classic Teams autostart entry for $($user.Account)"
                }
            }

            $uninstallKey = "$hive\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Teams"
            if (Test-Path $uninstallKey) {
                if ($PSCmdlet.ShouldProcess("$($user.Account): Uninstall\Teams", 'Remove stale uninstall entry')) {
                    Remove-Item -Path $uninstallKey -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Ok "Removed the stale uninstall entry for $($user.Account)"
                }
            }
        }
    }

    # -- 5. Download and verify the bootstrapper -------------------------------
    # Deliberately before any uninstall: a failed download must never leave the
    # device without a Teams client.
    Write-Out ''
    Write-Step '5. Download new Teams bootstrapper'
    if (-not $fullReinstall) {
        Write-Skip 'Not needed - the client stays as it is'
    } else {
        if ($BootstrapperUrl -notmatch '^https://') {
            throw "BootstrapperUrl must be https: $BootstrapperUrl"
        }

        if (-not $workingDirReady) {
            if ($PSCmdlet.ShouldProcess($WorkingDir, 'Create working directory')) {
                New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
                Write-Ok "Working directory ready: $WorkingDir"
            }
            $workingDirReady = $true
        }

        if ($PSCmdlet.ShouldProcess($exePath, "Download $BootstrapperUrl")) {
            $file = Save-VerifiedDownload -Uri $BootstrapperUrl -Path $exePath
            Write-Ok "Downloaded $($file.Name) ($([math]::Round($file.Length / 1MB, 1)) MB), signature verified"
        }
    }

    # -- 6. Uninstall the add-in and the current package -----------------------
    Write-Out ''
    Write-Step '6. Uninstall current Teams'
    if (-not $fullReinstall) {
        Write-Skip 'Not needed - the client stays as it is'
    } else {
        if ($SkipMeetingAddIn) {
            Write-Skip 'Meeting add-in left alone (-SkipMeetingAddIn)'
        } else {
            $addInEntries = @(Get-TeamsMeetingAddInEntry)
            if ($addInEntries.Count -eq 0) { Write-Skip 'Microsoft Teams Meeting Add-in is not installed' }

            foreach ($entry in $addInEntries) {
                $target = "$($entry.DisplayName) $($entry.Version) [$($entry.ProductCode)]"
                if ($PSCmdlet.ShouldProcess($target, 'msiexec /x /qn (uninstall)')) {
                    $result = Invoke-Installer -FilePath 'msiexec.exe' -Arguments "/x $($entry.ProductCode) /qn /norestart"
                    if ($result.Success) {
                        Write-Ok "Uninstalled $($entry.DisplayName)"
                        if ($result.RebootRequired) { $rebootRequired = $true }
                    } elseif ($result.ExitCode -eq 1612) {
                        # 1612 means Windows Installer cannot find the cached MSI it
                        # needs. Its own copy under C:\Windows\Installer often is still
                        # there and works; when it is not, nothing can uninstall this
                        # product and it keeps blocking other versions with 1638.
                        $cached = @(Get-MsiCachedPackage -DisplayNameLike '*Teams Meeting Add-in*' |
                                    Where-Object { $_.Cached }) | Select-Object -First 1
                        if ($cached -and $PSCmdlet.ShouldProcess($cached.LocalPackage, 'msiexec /x /qn (uninstall from the cached package)')) {
                            Write-Warn "Windows Installer lost the source for $($entry.DisplayName) (1612) - retrying with its cached copy $($cached.LocalPackage)"
                            $retry = Invoke-Installer -FilePath 'msiexec.exe' -Arguments "/x `"$($cached.LocalPackage)`" /qn /norestart"
                            if ($retry.Success) {
                                Write-Ok "Uninstalled $($entry.DisplayName) from the cached package"
                                if ($retry.RebootRequired) { $rebootRequired = $true }
                            } else {
                                $addInOrphaned = $true
                                Write-Warn "The cached package did not uninstall either ($($retry.Message)) - $($entry.Version) stays registered"
                            }
                        } else {
                            $addInOrphaned = $true
                            Write-Warn "Windows Installer lost the source for $($entry.DisplayName) (1612) and its cached copy is gone too, so $($entry.Version) cannot be uninstalled - it stays registered and will refuse a different version with 1638"
                        }
                    } else {
                        Write-Warn "Uninstall of $($entry.DisplayName) failed ($($result.Message))"
                    }
                }
            }

            # The rest of the sweep - the folders and the per-user COM registrations -
            # deliberately waits until step 8, where the replacement MSI is in hand.
            # Removing every copy here and only then discovering the new package
            # carries an older add-in leaves the device with no add-in at all, which
            # is exactly what happened on a production host.
            Write-Skip 'The other add-in copies are removed in step 8, once the replacement MSI is known to be installable'
        }

        if ($existingTeams.Count -eq 0) { Write-Skip 'No AppX package to remove' }
        foreach ($pkg in $existingTeams) {
            if ($PSCmdlet.ShouldProcess("$($pkg.Name) $($pkg.Version)", 'Remove-AppxPackage -AllUsers')) {
                $pkg | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
                if (Get-AppxPackage -AllUsers -Name $pkg.Name -ErrorAction SilentlyContinue) {
                    Write-Warn "$($pkg.Name) is still present after removal - the reinstall will upgrade it in place"
                } else {
                    Write-Ok "Removed $($pkg.Name)"
                }
            }
        }

        # Without dropping the provisioned copy, new user profiles keep getting the
        # old version staged from the image.
        $provisioned = @(Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like 'MSTeams*' })
        if ($provisioned.Count -eq 0) { Write-Skip 'No provisioned MSTeams package to remove' }
        foreach ($prov in $provisioned) {
            if ($PSCmdlet.ShouldProcess($prov.PackageName, 'Remove-AppxProvisionedPackage -Online')) {
                try {
                    Remove-AppxProvisionedPackage -Online -PackageName $prov.PackageName | Out-Null
                    Write-Ok "Deprovisioned $($prov.PackageName)"
                } catch {
                    Write-Warn "Could not deprovision $($prov.PackageName): $($_.Exception.Message)"
                }
            }
        }
    }

    # -- 7. Install / provision new Teams --------------------------------------
    Write-Out ''
    Write-Step '7. Install new Teams'
    if (-not $fullReinstall) {
        Write-Skip 'Not needed - the client stays as it is'
    } elseif ($PSCmdlet.ShouldProcess($exePath, 'Provision new Teams for all users (-p)')) {
        if (-not (Test-Path $exePath)) { throw "Bootstrapper not found at $exePath" }

        $result = Invoke-Installer -FilePath $exePath -Arguments '-p'
        if (-not $result.Success) { throw "Bootstrapper failed ($($result.Message))" }
        if ($result.RebootRequired) { $rebootRequired = $true }
        Write-Ok 'Bootstrapper completed'
    }

    # -- 8. Install the Teams Meeting Add-in for all users ---------------------
    Write-Out ''
    Write-Step '8. Teams Meeting Add-in (install)'
    if ($SkipMeetingAddIn) {
        Write-Skip 'Skipped (-SkipMeetingAddIn)'
    } elseif (-not $fullReinstall -and -not $addInMissing) {
        Write-Skip 'Already installed and the client was not replaced'
    } else {
        $tmaMsi = Get-TeamsAddInInstaller

        if (-not $tmaMsi -and $simulate) {
            Write-Skip 'Nothing was installed in this dry run, so the staged package is not there either - a real run takes the MSI from it'
            Write-Skip 'Would run: msiexec.exe /i "<ProgramFiles>\WindowsApps\MSTeams_<version>_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi" TARGETDIR="<ProgramFiles(x86)>\Microsoft\TeamsMeetingAddin\<version>\" /qn ALLUSERS=1'
        } elseif (-not $tmaMsi) {
            throw "No add-in MSI found under $env:ProgramFiles\WindowsApps - did the bootstrapper stage the package?"
        } else {
            Write-Ok "Add-in MSI from staged package $($tmaMsi.Directory.Name)"
            $tmaVersion = Get-MsiProductVersion -Path $tmaMsi.FullName
            if (-not $tmaVersion) { throw "Could not read the product version from $($tmaMsi.FullName)" }

            Write-Ok "Found Teams Meeting Add-in version: $tmaVersion"

            # Windows Installer refuses to put an older version over a newer one that
            # is still registered (1638). Finding that out before touching the copies
            # that do work is the difference between "nothing changed" and "the device
            # has no add-in".
            $stillRegistered = @(Get-TeamsMeetingAddInEntry) | Select-Object -First 1
            $msiOlder        = $false
            if ($stillRegistered) {
                try { $msiOlder = ([version] $tmaVersion) -lt ([version] $stillRegistered.Version) } catch { $msiOlder = $false }
            }

            if ($msiOlder) {
                Write-Warn "The staged package carries add-in $tmaVersion, older than the $($stillRegistered.Version) still registered - Windows Installer would answer 1638, so the working add-in is left exactly as it is"
                if ($addInOrphaned) {
                    Write-Warn 'That registration cannot be uninstalled either, because Windows Installer has lost its cached MSI - the way out is a newer Teams build whose add-in is at least this version, or removing the product registration by hand'
                }
            } else {
                # Now that a usable MSI is in hand, clear every other copy: a per-user
                # folder or COM registration that survives shadows the fresh
                # machine-wide one and is how a LoadBehavior 2 is born.
                foreach ($folder in @(Get-TeamsAddInFolder)) {
                    if ($PSCmdlet.ShouldProcess("$($folder.Account): $($folder.Path)", 'Remove leftover add-in folder')) {
                        Remove-Item -LiteralPath $folder.Path -Recurse -Force -ErrorAction SilentlyContinue
                        if (Test-Path $folder.Path) {
                            Write-Warn "Could not fully remove $($folder.Path) - files in use; Outlook or Teams may still hold them"
                        } else {
                            Write-Ok "Removed the add-in copy for $($folder.Account)"
                        }
                    }
                }

                foreach ($reg in @(Get-OutlookAddInRegistration | Where-Object { $_.Sid })) {
                    foreach ($clsidKey in @(Get-AddInClsidKey $reg.ClassesRoot)) {
                        if ($PSCmdlet.ShouldProcess("$($reg.Account): $clsidKey", 'Remove the per-user COM registration')) {
                            Remove-Item -Path $clsidKey -Recurse -Force -ErrorAction SilentlyContinue
                            if (-not (Test-Path $clsidKey)) {
                                Write-Ok "Cleared the per-user COM registration for $($reg.Account)"
                            }
                        }
                    }

                    if ($reg.LoadBehavior -ne 3 -and
                        $PSCmdlet.ShouldProcess("$($reg.Account): LoadBehavior", 'Set to 3 (load at startup)')) {
                        Set-ItemProperty -Path $reg.AddInKeyPath -Name 'LoadBehavior' -Value 3 -Type DWord -ErrorAction SilentlyContinue
                        Write-Ok "Set LoadBehavior back to 3 for $($reg.Account)"
                    }
                }

                $targetDir = '{0}\Microsoft\TeamsMeetingAddin\{1}\' -f ${env:ProgramFiles(x86)}, $tmaVersion
                $params    = '/i "{0}" TARGETDIR="{1}" /qn /norestart ALLUSERS=1' -f $tmaMsi.FullName, $targetDir

                if ($PSCmdlet.ShouldProcess("Teams Meeting Add-in $tmaVersion", "msiexec.exe $params")) {
                    $result = Invoke-Installer -FilePath 'msiexec.exe' -Arguments $params
                    if ($result.ExitCode -eq 1638) {
                        # Not fatal any more: aborting here used to skip verification,
                        # which is the one thing that says whether Outlook still has a
                        # working add-in after all this.
                        Write-Warn "Windows Installer reports another version of the add-in is already installed (1638) - $tmaVersion was not installed; verification below reports what Outlook is left with"
                    } elseif (-not $result.Success) {
                        throw "Teams Meeting Add-in install failed ($($result.Message))"
                    } else {
                        if ($result.RebootRequired) { $rebootRequired = $true }
                        Write-Ok "Installed Teams Meeting Add-in to $targetDir"
                    }
                }
            }
        }
    }

    # -- 8b. Repair per-user registrations that point at a removed DLL ---------
    # A user whose own HKCU registration survives an older add-in keeps shadowing
    # the machine-wide one: HKCU\SOFTWARE\Classes wins, Outlook loads nothing and
    # switches the add-in off. Dropping that stale class lets the machine-wide
    # registration take over, which is why this only runs when that one is healthy.
    if (-not $SkipMeetingAddIn -and $brokenAddInUsers.Count -gt 0) {
        Write-Out ''
        Write-Step '8b. Per-user Outlook registration'

        if (-not $RepairOutlookAddIn) {
            Write-Skip ('{0} user(s) point at a removed DLL - use -RepairOutlookAddIn to clear those registrations' -f $brokenAddInUsers.Count)
        } elseif (-not $machineWideAddInOk) {
            Write-Warn 'Not repairing: there is no healthy machine-wide registration for these users to fall back on'
        } else {
            foreach ($reg in $brokenAddInUsers) {
                foreach ($clsidKey in @(Get-AddInClsidKey $reg.ClassesRoot)) {
                    if ($PSCmdlet.ShouldProcess("$($reg.Account): $clsidKey", 'Remove the stale COM registration')) {
                        Remove-Item -Path $clsidKey -Recurse -Force -ErrorAction SilentlyContinue
                        if (Test-Path $clsidKey) {
                            Write-Warn "Could not remove $clsidKey"
                        } else {
                            Write-Ok "Cleared the stale COM registration for $($reg.Account)"
                        }
                    }
                }

                # Outlook parked the add-in at 2 (or 0) when the load failed; with the
                # class gone it would otherwise stay parked.
                if ($reg.LoadBehavior -ne 3 -and
                    $PSCmdlet.ShouldProcess("$($reg.Account): LoadBehavior", 'Set to 3 (load at startup)')) {
                    Set-ItemProperty -Path $reg.AddInKeyPath -Name 'LoadBehavior' -Value 3 -Type DWord -ErrorAction SilentlyContinue
                    Write-Ok "Set LoadBehavior back to 3 for $($reg.Account)"
                }
            }
            Write-Skip 'Those users get the add-in from the machine-wide registration at their next Outlook start'
        }
    }

    # -- 9. Verify -------------------------------------------------------------
    Write-Out ''
    Write-Step '9. Verification'
    if ($simulate) {
        Write-Skip 'Skipped - nothing was changed, so there is nothing to verify (-WhatIf)'
        Write-Out ''
        Write-Out '  Dry run only - rerun without -WhatIf to apply these changes.' 'Yellow'
    } else {
        if ($AvdOptimizations) {
            if ((Get-WvdEnvironmentFlag) -eq 1) { Write-Ok 'AVD media flag IsWVDEnvironment is 1' }
            else { Write-Bad 'AVD media flag IsWVDEnvironment is not set'; $exitCode = 1 }

            $webRtcNow = @(Get-WebRtcRedirectorEntry) | Select-Object -First 1
            if ($webRtcNow) { Write-Ok "WebRTC Redirector installed ($($webRtcNow.Version))" }
            else { Write-Bad 'WebRTC Redirector installation failed'; $exitCode = 1 }
        }

        if ($RemoveWebRtcRedirector) {
            if (@(Get-WebRtcRedirectorEntry).Count -eq 0) { Write-Ok 'WebRTC Redirector is gone' }
            else { Write-Bad 'The WebRTC Redirector is still installed'; $exitCode = 1 }
        }

        if ($RemoveClassicTeams) {
            # A machine-wide installer that survives is a genuine failure. A per-user
            # folder that survives is usually a file lock, which the next run clears
            # once the user has signed out.
            if (@(Get-ClassicTeamsEntry).Count -gt 0) {
                Write-Bad 'The classic Teams machine-wide installer is still present'
                $exitCode = 1
            } else {
                Write-Ok 'No classic Teams machine-wide installer'
            }

            $classicLeft = @(Get-ClassicTeamsUserInstall)
            if ($classicLeft.Count -eq 0) {
                Write-Ok 'No per-user classic Teams left'
            } else {
                foreach ($left in $classicLeft) {
                    Write-Warn "Classic Teams still present for $($left.Account) - files in use; it clears on a later run"
                }
            }
        }

        if (-not $SkipMeetingAddIn) {
            if (Get-TeamsMeetingAddInEntry) { Write-Ok 'Teams Meeting Add-in installed (machine-wide)' }
            else { Write-Bad 'Teams Meeting Add-in installation failed'; $exitCode = 1 }

            # Does Outlook itself see it? Not a pass/fail: a profile that is not
            # signed in cannot be read, and a fresh registration appears at the next
            # Outlook start. Reported so a technician knows what to check with.
            Write-OutlookAddInStatus
        }

        $nowInstalled = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -eq 'MSTeams' } | Select-Object -First 1
        if ($nowInstalled) {
            Write-Ok "Teams provisioned for all users ($($nowInstalled.Version))"
            if ($latest -and [version] $nowInstalled.Version -lt $latest.Version) {
                Write-Warn "Installed build $($nowInstalled.Version) is still older than the published $($latest.Version)"
            }
        } else {
            Write-Bad 'Teams installation failed'
            $exitCode = 1
        }

        Write-Out ''
        if ($rebootRequired) { Write-Warn 'A reboot is required to complete the installation (MSI returned 3010)' }
        if ($exitCode -eq 0) { Write-Ok 'Done' }
    }
} catch {
    if ($null -ne $plannedExit) {
        # A deliberate stop: up to date, check-only, or cancelled by the operator.
        Write-Out ''
        Write-Skip $_.Exception.Message
        $exitCode = $plannedExit
    } else {
        Show-HeldOutput
        Write-Out ''
        Write-Bad "Aborted: $($_.Exception.Message)"
        $exitCode = 1
    }
} finally {
    if ($transcribing) {
        Write-Out ''
        try { Stop-Transcript | Out-Null } catch { Write-Warning "Transcript not closed cleanly: $($_.Exception.Message)" }
    }
}

if (-not $script:holdOutput) { Write-Host '' }
exit $exitCode















