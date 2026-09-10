#Requires -Version 5.1
<#
.SYNOPSIS
    Update the new Microsoft Teams client, including the Teams Meeting Add-in for
    Outlook, only when Microsoft publishes a newer build. Supports -WhatIf.

.DESCRIPTION
    Endpoint/AVD script that keeps new Teams current:

      1. Preflight  - the installed MSTeams AppX package, the meeting add-in and any
                      running Teams/Outlook process.
      2. Check      - ask the Teams client config service which build is current for
                      this architecture and compare it with what is installed. Up to
                      date and nothing missing? Nothing happens at all.
      3. AVD        - only with -AvdOptimizations: the IsWVDEnvironment media flag and
                      the Remote Desktop WebRTC Redirector Service.
      4. Download   - fetch teamsbootstrapper.exe and verify its Microsoft signature
                      BEFORE anything is uninstalled, so a failed download can never
                      leave the device without a Teams client.
      5. Uninstall  - Teams Meeting Add-in (MSI), the MSTeams AppX package for all
                      users, and the provisioned package.
      6. Install    - provision new Teams for all users (teamsbootstrapper.exe -p).
      7. Add-in     - install the Teams Meeting Add-in MSI shipped inside the new
                      Teams package (ALLUSERS=1).
      8. Verify     - re-check the add-in registration, the provisioned package and,
                      where applicable, the AVD components.

    Each part is only done when it is actually needed. A current client with a missing
    add-in installs just the add-in; a current client on an AVD host with the WebRTC
    redirector missing installs just that.

    AVD / VDI
    ---------
    -AvdOptimizations adds the two things a session host needs for media optimization:
    HKLM:\SOFTWARE\Microsoft\Teams\IsWVDEnvironment = 1 (set before Teams is
    provisioned, which is why it is step 3) and the Remote Desktop WebRTC Redirector
    Service from https://aka.ms/msrdcwebrtcsvc/msi. Both are installed only when
    missing, so a scheduled run on a session host costs no extra download. -Force
    reinstalls the redirector as well. Without the switch the script only points out
    that the device looks like a session host.

    WebRTC is being retired: end of support 1 October 2026, end of availability
    1 April 2027. Its replacement, SlimCore, needs nothing installed here - it ships
    inside new Teams (the Microsoft.Teams.SlimCore* packages, reported in preflight)
    and inside Windows App on the endpoint the user connects from. The endpoint's
    Windows App version is what decides which path is used, and this script cannot see
    that from the session host. IsWVDEnvironment stays required either way, and
    Microsoft still advises keeping the redirector as a fallback for endpoints that
    cannot do SlimCore, so the switch keeps installing it. Revisit before April 2027.

    Every state-changing step is wrapped in ShouldProcess, so -WhatIf walks the whole
    flow and reports exactly what would be uninstalled, downloaded, installed and
    provisioned without touching the machine.

    Version check
    -------------
    https://config.teams.microsoft.com/config/v1/MicrosoftTeams/... is the feed the
    Teams client itself uses to decide it is out of date. It returns the current build
    per architecture (BuildSettings.WebView2PreAuth.<arch>.latestVersion). An
    installed build that is equal or newer means there is nothing to do. If the
    service cannot be reached the run stops rather than reinstalling blindly; -Force
    overrides that.

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
        quiet, checkOnly, force, avdOptimizations, skipMeetingAddIn or
        skipSignatureCheck and text fields named workingDir, logPath or ring are
        picked up when the matching parameter is not passed on the command line.
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

.PARAMETER WebRtcUrl
    Download URL for the Remote Desktop WebRTC Redirector MSI (default: the Microsoft
    aka.ms link). Must be https. Only used with -AvdOptimizations.

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
    # Repair: reinstall the current build regardless of the version check
    .\Update-TeamsClient.ps1 -Force

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows only (AppX + MSI), run as administrator or as System
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [switch] $Quiet,
    [switch] $CheckOnly,
    [switch] $AvdOptimizations,
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
if (-not $PSBoundParameters.ContainsKey('WorkingDir')         -and $env:workingDir)                      { $WorkingDir         = $env:workingDir }
if (-not $PSBoundParameters.ContainsKey('LogPath')            -and $env:logPath)                         { $LogPath            = $env:logPath }
if (-not $PSBoundParameters.ContainsKey('Ring')               -and $env:ring)                            { $Ring               = $env:ring }
if (-not $PSBoundParameters.ContainsKey('WebRtcUrl')          -and $env:webRtcUrl)                       { $WebRtcUrl          = $env:webRtcUrl }

$simulate         = [bool] $WhatIfPreference
$teamsExe         = 'teamsbootstrapper.exe'
$exePath          = Join-Path $WorkingDir $teamsExe
$webRtcMsi        = Join-Path $WorkingDir 'MsRdcWebRTCSvc_x64.msi'
$exitCode         = 0
$rebootRequired   = $false
$transcribing     = $false
$plannedExit      = $null
$workingDirReady  = $false

# Teams reads this flag to switch to VDI media optimization; it has to be there
# before the client is provisioned.
$AvdRegistryPath = 'HKLM:\SOFTWARE\Microsoft\Teams'

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
                    ProductCode = $key.PSChildName
                    DisplayName = $name
                    Version     = $key.GetValue('DisplayVersion')
                }
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
                    ProductCode = $key.PSChildName
                    DisplayName = $name
                    Version     = $key.GetValue('DisplayVersion')
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

        return [PSCustomObject]@{
            Success        = ($code -in @(0, 3010))
            ExitCode       = $code
            RebootRequired = ($code -eq 3010)
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
        Write-Ok "Found $($pkg.Name) $($pkg.Version)"
        $candidate = [version] $pkg.Version
        if ($null -eq $installedVersion -or $candidate -gt $installedVersion) { $installedVersion = $candidate }
    }
    if ($existingTeams.Count -eq 0) {
        if ($Force) { Write-Skip 'No Teams installation detected - continuing because -Force was given' }
        else        { throw 'No Teams installation detected. Use -Force to install new Teams anyway.' }
    }

    $addInInstalled = [bool] (Get-TeamsMeetingAddInEntry)
    if (-not $SkipMeetingAddIn) {
        if ($addInInstalled) { Write-Ok 'Teams Meeting Add-in is installed' }
        else                 { Write-Warn 'Teams Meeting Add-in is not installed' }
    }

    $avdFlagSet  = $false
    $webRtcEntry = $null
    if ($AvdOptimizations) {
        $avdFlagSet  = ((Get-WvdEnvironmentFlag) -eq 1)
        $webRtcEntry = @(Get-WebRtcRedirectorEntry) | Select-Object -First 1

        if ($avdFlagSet) { Write-Ok 'AVD media flag IsWVDEnvironment is set' }
        else             { Write-Warn 'AVD media flag IsWVDEnvironment is not set' }

        if ($webRtcEntry) { Write-Ok "WebRTC Redirector is installed ($($webRtcEntry.Version))" }
        else              { Write-Warn 'Remote Desktop WebRTC Redirector is not installed' }

        # Informational only. SlimCore replaces WebRTC (support ends 1 October 2026)
        # and ships with new Teams, but whether it is actually used depends on the
        # Windows App version on the endpoint, which is invisible from here.
        $slimCoreHost = @(Get-AppxPackage -Name 'Microsoft.Teams.SlimCoreVdiHost*' -ErrorAction SilentlyContinue) |
                        Select-Object -First 1
        if ($slimCoreHost) {
            Write-Ok "SlimCore is present ($($slimCoreHost.Version)) - the endpoint's Windows App version decides whether it is used"
        } else {
            Write-Warn 'SlimCore packages not found - this host still depends on WebRTC, which loses support on 1 October 2026'
        }
    } elseif (Test-AvdSessionHost) {
        Write-Skip 'This looks like an AVD session host - consider -AvdOptimizations for the media flag and the WebRTC redirector'
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
        } else {
            Write-Ok "Installed build is current ($installedVersion)"
        }
    }

    $addInMissing  = (-not $SkipMeetingAddIn) -and (-not $addInInstalled)
    $avdWork       = $AvdOptimizations -and ((-not $avdFlagSet) -or (-not $webRtcEntry) -or $Force)
    $fullReinstall = $clientOutdated -or $Force

    if (-not $fullReinstall -and -not $addInMissing -and -not $avdWork) {
        $plannedExit = 0
        throw 'Teams is up to date - nothing to do.'
    }

    $reasons = @()
    if ($clientOutdated) { $reasons += 'a newer build is available' }
    if ($addInMissing)   { $reasons += 'the Teams Meeting Add-in is missing' }
    if ($avdWork)        { $reasons += 'the AVD optimizations are incomplete' }
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
    if (-not $AvdOptimizations) {
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

            if ($PSCmdlet.ShouldProcess('Remote Desktop WebRTC Redirector Service', 'msiexec /i /qn (install)')) {
                $result = Invoke-Installer -FilePath 'msiexec.exe' -Arguments "/i `"$webRtcMsi`" /qn /norestart"
                if (-not $result.Success) { throw "WebRTC Redirector install failed ($($result.Message))" }
                if ($result.RebootRequired) { $rebootRequired = $true }
                Write-Ok 'WebRTC Redirector installed'
            }
        }
    }

    # -- 4. Download and verify the bootstrapper -------------------------------
    # Deliberately before any uninstall: a failed download must never leave the
    # device without a Teams client.
    Write-Out ''
    Write-Step '4. Download new Teams bootstrapper'
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

    # -- 5. Uninstall the add-in and the current package -----------------------
    Write-Out ''
    Write-Step '5. Uninstall current Teams'
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
                    } else {
                        # Not fatal: the reinstall below replaces the add-in anyway.
                        Write-Warn "Uninstall of $($entry.DisplayName) failed ($($result.Message))"
                    }
                }
            }
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

    # -- 6. Install / provision new Teams --------------------------------------
    Write-Out ''
    Write-Step '6. Install new Teams'
    if (-not $fullReinstall) {
        Write-Skip 'Not needed - the client stays as it is'
    } elseif ($PSCmdlet.ShouldProcess($exePath, 'Provision new Teams for all users (-p)')) {
        if (-not (Test-Path $exePath)) { throw "Bootstrapper not found at $exePath" }

        $result = Invoke-Installer -FilePath $exePath -Arguments '-p'
        if (-not $result.Success) { throw "Bootstrapper failed ($($result.Message))" }
        if ($result.RebootRequired) { $rebootRequired = $true }
        Write-Ok 'Bootstrapper completed'
    }

    # -- 7. Install the Teams Meeting Add-in for all users ---------------------
    Write-Out ''
    Write-Step '7. Teams Meeting Add-in (install)'
    if ($SkipMeetingAddIn) {
        Write-Skip 'Skipped (-SkipMeetingAddIn)'
    } elseif (-not $fullReinstall -and -not $addInMissing) {
        Write-Skip 'Already installed and the client was not replaced'
    } else {
        $newTeams        = Get-AppxPackage -Name 'MSTeams' -ErrorAction SilentlyContinue
        $newTeamsVersion = if ($newTeams) { $newTeams.Version } else { $null }

        if (-not $newTeamsVersion -and $simulate) {
            Write-Skip 'New Teams is not installed yet - on a real run the add-in MSI comes from the freshly installed package folder'
            Write-Skip 'Would run: msiexec.exe /i "<ProgramFiles>\WindowsApps\MSTeams_<version>_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi" TARGETDIR="<ProgramFiles(x86)>\Microsoft\TeamsMeetingAddin\<version>\" /qn ALLUSERS=1'
        } elseif (-not $newTeamsVersion) {
            throw 'New Teams package not found after install. Check the bootstrapper output.'
        } else {
            Write-Ok "Found new Teams version: $newTeamsVersion"

            $tmaPath    = '{0}\WindowsApps\MSTeams_{1}_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi' -f $env:ProgramFiles, $newTeamsVersion
            $tmaVersion = if (Test-Path $tmaPath) { Get-MsiProductVersion -Path $tmaPath } else { $null }

            if (-not $tmaVersion -and $simulate) {
                Write-Skip "Add-in MSI not present in the current package - a real run takes it from $tmaPath"
            } elseif (-not $tmaVersion) {
                throw "Teams Meeting Add-in installer not found at $tmaPath"
            } else {
                Write-Ok "Found Teams Meeting Add-in version: $tmaVersion"
                $targetDir = '{0}\Microsoft\TeamsMeetingAddin\{1}\' -f ${env:ProgramFiles(x86)}, $tmaVersion
                $params    = '/i "{0}" TARGETDIR="{1}" /qn /norestart ALLUSERS=1' -f $tmaPath, $targetDir

                if ($PSCmdlet.ShouldProcess("Teams Meeting Add-in $tmaVersion", "msiexec.exe $params")) {
                    $result = Invoke-Installer -FilePath 'msiexec.exe' -Arguments $params
                    if (-not $result.Success) { throw "Teams Meeting Add-in install failed ($($result.Message))" }
                    if ($result.RebootRequired) { $rebootRequired = $true }
                    Write-Ok "Installed Teams Meeting Add-in to $targetDir"
                }
            }
        }
    }

    # -- 8. Verify -------------------------------------------------------------
    Write-Out ''
    Write-Step '8. Verification'
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

        if (-not $SkipMeetingAddIn) {
            if (Get-TeamsMeetingAddInEntry) { Write-Ok 'Teams Meeting Add-in installed' }
            else { Write-Bad 'Teams Meeting Add-in installation failed'; $exitCode = 1 }
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
