#Requires -Version 5.1
<#
.SYNOPSIS
    Bring the ephemeral temp disk back as D: and keep the pagefile on it. Meant to
    run at every boot from a scheduled task. Supports -WhatIf.

.DESCRIPTION
    An Azure VM's temp disk - the resource disk, or the local NVMe disk on the newer
    sizes - is wiped whenever the VM is deallocated, resized or moved to another
    host. It comes back empty, sometimes RAW, sometimes offline, sometimes without
    its drive letter. A pagefile configured on a drive letter that is not there at
    boot is simply not created, so Windows falls back to a pagefile on C: or runs
    with none at all. That is the failure this script exists to prevent.

    Per run:

      1. Preflight - the disks, what holds the target drive letter, the pagefile as
                     configured in the registry, and the pagefile actually in use in
                     this session.
      2. Letter    - an optical drive sitting on D: is moved out of the way, because
                     Windows hands D: to the DVD on an image that has no temp disk
                     yet and then never gives it back.
      3. Disk      - a volume that is already the temp disk (recognised by its label
                     or by the DataLoss_Warning_Readme.txt that Azure writes on it)
                     gets the drive letter back. Otherwise a RAW, non-boot,
                     non-system disk is brought online, initialised GPT, partitioned
                     and formatted NTFS.
      4. Pagefile  - automatic management off, the pagefile pointed at D:\pagefile.sys
                     and the entry for every other drive removed.
      5. Verify    - everything read back, saying what is in effect now and what is
                     waiting for the next restart.
      6. Restart   - only with -RestartIfNeeded: restart the machine when that is the
                     one thing left between the configuration and a pagefile that is
                     actually in use.

    Safety
    ------
    Nothing that is not RAW is ever initialised. A disk that already carries
    partitions is reported and left alone: an empty temp disk and an unformatted data
    disk look identical from the outside, and only one of the two may be formatted.
    With more than one RAW candidate the script refuses to guess and asks for
    -DiskNumber; -Force plus -DiskNumber is the only way to format a disk that still
    has partitions on it.

    Why a reboot is still in the story
    ----------------------------------
    Windows reads the pagefile configuration at boot and never re-reads it. A
    pagefile written by this run therefore appears at the next restart, which is not
    a failure and is reported as such rather than counted as success. Because the
    task runs at every boot the device heals itself: the boot that recreates D:
    configures the pagefile, the next boot puts it in use.

    -RestartIfNeeded closes that gap instead of waiting for it. A script that runs at
    every boot and may restart the machine is a reboot loop waiting to happen, so it
    only fires when a restart demonstrably fixes something and cannot repeat:

      - the run finished clean, the temp disk is there, the pagefile is configured on
        it, and the only thing missing is that this session is not using it yet. A
        run that failed never restarts - that would hide the failure behind a reboot;
      - nobody is signed in, connected or disconnected. A session host restarts when
        it is drained, not while someone is in a meeting. -RestartEvenIfUsersSignedIn
        overrides that for a fleet where the countdown is warning enough;
      - at most one restart per -RestartCooldownMinutes, remembered in the registry
        under -RestartMarkerPath. A second restart for the same
        thing means the first one did not help, and repeating it forever is worse
        than saying so.

    The restart is announced through shutdown.exe. The countdown exists to warn
    people, so it only applies when there are people: signed in, it is
    -RestartDelaySeconds (60 by default) and shutdown /a stops it; nobody signed in -
    the normal case at boot, and the guaranteed one on a session host whose pool is
    drained - and the restart happens within seconds, because waiting a minute for an
    audience of nobody only costs availability.

    Exit codes
    ----------
        0  the temp disk and the pagefile are as they should be
        1  failure
        2  -CheckOnly only: work is due

.PARAMETER DriveLetter
    Drive letter for the temp disk (default: D).

.PARAMETER Label
    Volume label written when the disk is formatted, and the label an existing temp
    volume is recognised by (default: Temporary Storage - what Azure itself uses).

.PARAMETER DiskNumber
    Format this disk instead of letting the script pick one. Needed when more than
    one RAW disk is present, because guessing there can cost a data disk.

.PARAMETER InitialSizeMB
    Initial pagefile size in MB. 0 (the default) means system managed.

.PARAMETER MaximumSizeMB
    Maximum pagefile size in MB. 0 (the default) means system managed.

.PARAMETER KeepSystemDrivePagefile
    Leave an existing pagefile on C: (or any other drive) in place instead of
    removing it once the temp disk has one.

.PARAMETER SkipPagefile
    Only bring the temp disk back; do not touch the pagefile configuration at all.

.PARAMETER OpticalDriveLetter
    Letter an optical drive is moved to when it holds the target letter
    (default: Z). Any other free letter is used when that one is taken as well.

.PARAMETER RestartIfNeeded
    Restart the machine when that is the only thing left between the configuration
    and a pagefile that is actually in use. Off by default. Never restarts after a
    failed run, never while a user is signed in, and at most once per
    -RestartCooldownMinutes.

.PARAMETER RestartDelaySeconds
    Countdown before a restart that happens while someone is signed in (default: 60),
    so they see the notice and can stop it with shutdown /a. With nobody signed in -
    the normal case at boot - the restart happens within seconds instead, because
    there is nobody to warn and the wait only costs availability.

.PARAMETER RestartCooldownMinutes
    Shortest interval between two restarts triggered by this script (default: 60).
    A second restart for the same thing means the first one did not help.

.PARAMETER RestartMarkerPath
    Registry key where the last self-triggered restart is remembered
    (default: HKLM:\SOFTWARE\M365-Scripts\InitTempDisk). In the registry rather
    than next to the log, because it has to survive the restart it records.

.PARAMETER RestartEvenIfUsersSignedIn
    Restart even when someone is signed in. Only sensible where the countdown is
    warning enough - on a session host, drain it instead.

.PARAMETER CheckOnly
    Report only, change nothing. Exit code 2 means work is due.

.PARAMETER Quiet
    Print nothing unless there is news: a change, or a failure. For scheduled runs.
    The log file always gets the full story.

.PARAMETER LogPath
    Folder for Init-TempDisk.log (default: C:\Temp). The log is appended to and
    rotated once past 1 MB. A -WhatIf or -CheckOnly run writes no log.

.PARAMETER Force
    Together with -DiskNumber: format that disk even though it still has partitions.
    Everything on it is lost.

.EXAMPLE
    .\Init-TempDisk.ps1 -CheckOnly

    Read-only report: where the temp disk is, what holds D:, how the pagefile is
    configured and which pagefile this session is really using.

.EXAMPLE
    .\Init-TempDisk.ps1 -WhatIf

    Walk the whole flow and show what would be initialised, moved and configured.

.EXAMPLE
    .\Init-TempDisk.ps1 -Quiet

    How the scheduled task runs it: silent while everything is in order.

.EXAMPLE
    .\Init-TempDisk.ps1 -Quiet -RestartIfNeeded

    The same, but a boot that had to rebuild the temp disk restarts the machine so
    the pagefile is in use straight away - if nobody is signed in and no restart was
    triggered in the last hour.

.EXAMPLE
    .\Init-TempDisk.ps1 -InitialSizeMB 16384 -MaximumSizeMB 16384

    Fixed 16 GB pagefile on the temp disk instead of a system managed one.

.EXAMPLE
    .\Init-TempDisk.ps1 -DiskNumber 2 -Force

    Format disk 2 as the temp disk even though it still carries partitions.

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows only, run as administrator or as System
    Register it for every boot with Register-InitTempDiskTask.ps1 in this folder.
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [ValidatePattern('^[D-Zd-z]:?$')]
    [string] $DriveLetter = 'D',
    [string] $Label       = 'Temporary Storage',
    [int]    $DiskNumber  = -1,
    [ValidateRange(0, 1048576)]
    [int]    $InitialSizeMB = 0,
    [ValidateRange(0, 1048576)]
    [int]    $MaximumSizeMB = 0,
    [switch] $KeepSystemDrivePagefile,
    [switch] $SkipPagefile,
    [ValidatePattern('^[D-Zd-z]:?$')]
    [string] $OpticalDriveLetter = 'Z',
    [switch] $RestartIfNeeded,
    [ValidateRange(0, 3600)]
    [int]    $RestartDelaySeconds = 60,
    [ValidateRange(0, 10080)]
    [int]    $RestartCooldownMinutes = 60,
    [ValidatePattern('^HK(LM|CU):\\')]
    [string] $RestartMarkerPath = 'HKLM:\SOFTWARE\M365-Scripts\InitTempDisk',
    [switch] $RestartEvenIfUsersSignedIn,
    [switch] $CheckOnly,
    [switch] $Quiet,
    [string] $LogPath = 'C:\Temp',
    [switch] $Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$simulate      = [bool] $WhatIfPreference
$targetLetter  = $DriveLetter.TrimEnd(':').ToUpper()
$targetRoot    = "${targetLetter}:"
$pagefilePath  = "$targetRoot\pagefile.sys"
$exitCode         = 0
$changed          = $false
$rebootPending    = $false
$restartTriggered = $false

function Test-Elevated {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    return ([Security.Principal.WindowsPrincipal] $identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# -- Output --------------------------------------------------------------------
# -Quiet holds every line back until something worth reporting happens, so a boot on
# a healthy machine produces no console output at all. The log file is written either
# way: a task that runs unattended at boot is worth being able to read back.
$script:heldOutput = [System.Collections.Generic.List[object]]::new()
$script:holdOutput = [bool] $Quiet
$script:logFile    = $null

function Write-Log {
    param([string] $Message)
    if (-not $script:logFile) { return }
    try { Add-Content -LiteralPath $script:logFile -Value ('{0:yyyy-MM-dd HH:mm:ss}  {1}' -f (Get-Date), $Message) }
    catch { $script:logFile = $null }   # a log that cannot be written must not abort the run
}

function Write-Out {
    param([string] $Message = '', [string] $Color = 'Gray')
    Write-Log $Message
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

function Start-LogFile {
    <# One appended log, rotated at 1 MB, so a machine that reboots daily keeps a
       readable history instead of an ever-growing file or a folder full of them. #>
    param([string] $Folder)

    if ($simulate -or $CheckOnly) { return }
    try {
        if (-not (Test-Path -LiteralPath $Folder)) { New-Item -ItemType Directory -Path $Folder -Force | Out-Null }
        $file = Join-Path $Folder 'Init-TempDisk.log'
        if ((Test-Path -LiteralPath $file) -and (Get-Item -LiteralPath $file).Length -gt 1MB) {
            Move-Item -LiteralPath $file -Destination "$file.old" -Force
        }
        $script:logFile = $file
        Write-Log "--- Init-TempDisk started (user $env:USERNAME, drive $targetRoot) ---"
    } catch {
        Write-Warning "Could not open the log in ${Folder}: $($_.Exception.Message)"
    }
}

function Get-FreeDriveLetter {
    <# A letter no volume and no PSDrive is using, preferring the one asked for. #>
    param([string] $Preferred)

    $used = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($volume in @(Get-CimInstance -ClassName Win32_Volume -ErrorAction SilentlyContinue)) {
        if ($volume.DriveLetter) { [void] $used.Add($volume.DriveLetter.TrimEnd(':').ToUpper()) }
    }
    foreach ($drive in @(Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue)) {
        if ($drive.Name.Length -eq 1) { [void] $used.Add($drive.Name.ToUpper()) }
    }

    $wanted = $Preferred.TrimEnd(':').ToUpper()
    if ($wanted -and -not $used.Contains($wanted)) { return $wanted }
    foreach ($letter in [char[]] 'ZYXWVUTSRQPONMLKJIHGFE') {
        if (-not $used.Contains([string] $letter)) { return [string] $letter }
    }
    return $null
}

function Get-OpticalDriveOnLetter {
    <# The CD/DVD drive holding a letter, if that is what is holding it (type 5). #>
    param([string] $Root)
    return @(Get-CimInstance -ClassName Win32_Volume -Filter "DriveType = 5 AND DriveLetter = '$Root'" -ErrorAction SilentlyContinue) |
           Select-Object -First 1
}

function Get-TempDiskPartition {
    <#
        The partition that already is the temp disk: the label this script writes (and
        Azure writes itself), or the readme Azure drops on the resource disk. That
        second check is what recognises a disk the Azure agent formatted with a label
        in another language, or one an operator renamed.
    #>
    param([string] $VolumeLabel)

    foreach ($partition in @(Get-Partition -ErrorAction SilentlyContinue)) {
        $volume = $partition | Get-Volume -ErrorAction SilentlyContinue
        if (-not $volume -or $volume.DriveType -ne 'Fixed') { continue }

        if ($volume.FileSystemLabel -eq $VolumeLabel) { return $partition }
        if ($volume.DriveLetter -and (Test-Path -LiteralPath "$($volume.DriveLetter):\DataLoss_Warning_Readme.txt")) {
            return $partition
        }
    }
    return $null
}

function Get-TempDiskCandidate {
    <#
        The disks that may be formatted as the temp disk. Only RAW, never the boot or
        the system disk. More than one candidate is not a choice this script is
        entitled to make on its own - a data disk waiting to be formatted looks
        exactly the same - so they are all returned and the caller refuses.
    #>
    $disks = @(Get-Disk -ErrorAction SilentlyContinue | Where-Object {
        -not $_.IsBoot -and -not $_.IsSystem -and $_.PartitionStyle -eq 'RAW'
    })

    # A local NVMe disk on the newer VM sizes is the temp disk by definition, so it
    # wins when several RAW disks are present.
    $local = @($disks | Where-Object { $_.BusType -eq 'NVMe' })
    if ($local.Count -eq 1) { return $local }
    return $disks
}

function Get-PagefileState {
    <#
        The pagefile as configured (registry, read at boot) and as in use (this
        session). The two disagreeing is the normal state right after a change, and
        telling them apart is the difference between "done" and "done at next boot".
    #>
    $system  = Get-CimInstance -ClassName Win32_ComputerSystem
    $setting = @(Get-CimInstance -ClassName Win32_PageFileSetting -ErrorAction SilentlyContinue)
    $usage   = @(Get-CimInstance -ClassName Win32_PageFileUsage -ErrorAction SilentlyContinue)

    return [PSCustomObject]@{
        AutomaticManaged = [bool] $system.AutomaticManagedPagefile
        Configured       = @($setting | ForEach-Object {
                               [PSCustomObject]@{ Name = $_.Name; Initial = [int] $_.InitialSize; Maximum = [int] $_.MaximumSize }
                           })
        Active           = @($usage | ForEach-Object { $_.Name })
    }
}

function Format-PagefileEntry {
    param($Entry)
    if ($Entry.Initial -eq 0 -and $Entry.Maximum -eq 0) { return "$($Entry.Name) (system managed)" }
    return "$($Entry.Name) ($($Entry.Initial)-$($Entry.Maximum) MB)"
}

function Set-PagefileEntry {
    <#
        Point the pagefile at one file. New-CimInstance is the documented route; it is
        refused on some builds, where writing the registry value Windows reads at boot
        is the same change through the other door.
    #>
    param([string] $Path, [int] $Initial, [int] $Maximum)

    $existing = @(Get-CimInstance -ClassName Win32_PageFileSetting -ErrorAction SilentlyContinue |
                  Where-Object { $_.Name -eq $Path }) | Select-Object -First 1
    if ($existing) {
        Set-CimInstance -InputObject $existing -Property @{ InitialSize = [uint32] $Initial; MaximumSize = [uint32] $Maximum }
        return
    }

    try {
        $null = New-CimInstance -ClassName Win32_PageFileSetting -ErrorAction Stop -Property @{
            Name = $Path; InitialSize = [uint32] $Initial; MaximumSize = [uint32] $Maximum
        }
    } catch {
        Write-Warn "WMI refused to create the pagefile entry ($($_.Exception.Message)) - writing the registry value instead"
        $memory  = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management'
        $current = @(Get-ItemProperty -Path $memory -Name 'PagingFiles' -ErrorAction SilentlyContinue |
                     ForEach-Object { $_.PagingFiles })
        $kept    = @($current | Where-Object { $_ -and $_ -notmatch [regex]::Escape($Path) })
        Set-ItemProperty -Path $memory -Name 'PagingFiles' -Type MultiString `
                         -Value (@($kept) + ('{0} {1} {2}' -f $Path, $Initial, $Maximum))
    }
}

function Get-InteractiveSession {
    <#
        Who is on this machine, connected or disconnected. One explorer.exe runs per
        interactive desktop, which is language independent - unlike parsing query.exe,
        whose column headers follow the display language and would read an empty list
        out of a Dutch session host.
    #>
    $sessions = [System.Collections.Generic.List[string]]::new()
    foreach ($proc in @(Get-CimInstance -ClassName Win32_Process -Filter "Name = 'explorer.exe'" -ErrorAction SilentlyContinue)) {
        $who = $null
        try {
            $owner = Invoke-CimMethod -InputObject $proc -MethodName GetOwner -ErrorAction Stop
            if ($owner.User) {
                $who = if ($owner.Domain) { "$($owner.Domain)\$($owner.User)" } else { $owner.User }
            }
        } catch {
            $who = $null   # a session we cannot name still counts as a session
        }
        $sessions.Add($(if ($who) { $who } else { "session $($proc.SessionId)" }))
    }
    return @($sessions | Select-Object -Unique)
}

function Get-RestartMarker {
    <# When this script last restarted the machine, or $null if it never did. #>
    $item = Get-ItemProperty -Path $RestartMarkerPath -Name 'LastRestart' -ErrorAction SilentlyContinue
    if (-not $item) { return $null }
    try { return [datetime]::Parse($item.LastRestart, [Globalization.CultureInfo]::InvariantCulture) }
    catch { return $null }
}

function Set-RestartMarker {
    <# Record a restart so the next run can refuse to repeat it. Round-trip format
       on purpose: a locale-formatted timestamp written by one run and read by
       another under a different locale is how a cooldown quietly stops working. #>
    if (-not (Test-Path $RestartMarkerPath)) { New-Item -Path $RestartMarkerPath -Force | Out-Null }
    Set-ItemProperty -Path $RestartMarkerPath -Name 'LastRestart' -Value ((Get-Date).ToString('o')) -Force
}

# -- Elevation -----------------------------------------------------------------
# Disk and pagefile changes need administrator rights. Run by the scheduled task this
# is already true; started by hand it usually is not, so ask for elevation instead of
# failing halfway through.
if (-not (Test-Elevated)) {
    if (-not [Environment]::UserInteractive) {
        Write-Error 'Administrator rights are required. Run the script as System or from an elevated session.'
        exit 1
    }

    Write-Host ''
    Write-Host '  Not running elevated - asking for administrator rights...' -ForegroundColor Yellow
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-NoExit', '-File', $PSCommandPath)
    foreach ($entry in $PSBoundParameters.GetEnumerator()) {
        if ($entry.Value -is [switch] -or $entry.Value -is [bool]) {
            if ($entry.Value) { $argList += "-$($entry.Key)" } else { $argList += "-$($entry.Key):`$false" }
        } else {
            $argList += "-$($entry.Key)"
            $argList += [string] $entry.Value
        }
    }
    try {
        Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList $argList -Verb RunAs | Out-Null
        Write-Host '  Continued in an elevated window.' -ForegroundColor Cyan
        exit 0
    } catch {
        Write-Error "Elevation was declined or failed: $($_.Exception.Message)"
        exit 1
    }
}

Start-LogFile -Folder $LogPath

# -- Run -----------------------------------------------------------------------
try {
    Write-Out ''
    Write-Out ("  Mode: {0}" -f $(if ($CheckOnly)     { 'CHECK ONLY - reporting, never changing' }
                                  elseif ($simulate)  { '-WhatIf - nothing will be changed' }
                                  else                { "APPLY - $targetRoot is restored and the pagefile put on it" })) 'Cyan'
    Write-Out ''

    # -- 1. Preflight ----------------------------------------------------------
    Write-Step '1. Preflight'

    foreach ($disk in @(Get-Disk -ErrorAction SilentlyContinue | Sort-Object Number)) {
        $role = @()
        if ($disk.IsBoot)     { $role += 'boot' }
        if ($disk.IsSystem)   { $role += 'system' }
        if ($disk.IsOffline)  { $role += 'offline' }
        if ($disk.IsReadOnly) { $role += 'read-only' }
        $note = if ($role.Count -gt 0) { ' - ' + ($role -join ', ') } else { '' }
        Write-Skip ('Disk {0}: {1}, {2} GB, {3}, bus {4}{5}' -f
                    $disk.Number, $disk.FriendlyName, [math]::Round($disk.Size / 1GB, 0),
                    $disk.PartitionStyle, $disk.BusType, $note)
    }

    $targetVolume  = Get-Volume -DriveLetter $targetLetter -ErrorAction SilentlyContinue
    $tempPartition = Get-TempDiskPartition -VolumeLabel $Label
    $diskReady     = [bool] ($targetVolume -and $targetVolume.DriveType -eq 'Fixed')

    if ($diskReady) {
        Write-Ok ('{0} is a fixed volume: label "{1}", {2} GB free of {3} GB' -f
                  $targetRoot, $targetVolume.FileSystemLabel,
                  [math]::Round($targetVolume.SizeRemaining / 1GB, 1), [math]::Round($targetVolume.Size / 1GB, 1))
    } elseif ($targetVolume) {
        Write-Warn "$targetRoot exists but is a $($targetVolume.DriveType) drive, not a disk - it is in the way of the temp disk"
    } else {
        Write-Warn "$targetRoot does not exist - the temp disk is gone, offline or has lost its drive letter"
    }

    $pagefile = Get-PagefileState
    if ($SkipPagefile) {
        Write-Skip 'Pagefile left alone (-SkipPagefile)'
    } else {
        if ($pagefile.AutomaticManaged) {
            Write-Warn 'Windows manages the pagefile automatically - that is what puts it back on C: whenever the temp disk is missing'
        }
        if ($pagefile.Configured.Count -eq 0) {
            Write-Warn 'No pagefile is configured at all'
        } else {
            foreach ($entry in $pagefile.Configured) { Write-Ok "Configured pagefile: $(Format-PagefileEntry $entry)" }
        }
        if ($pagefile.Active.Count -eq 0) {
            Write-Warn 'This session is running without a pagefile'
        } else {
            foreach ($active in $pagefile.Active) { Write-Ok "Pagefile in use this session: $active" }
        }
    }

    # -- 2. What is due --------------------------------------------------------
    $wantedEntry = [PSCustomObject]@{ Name = $pagefilePath; Initial = $InitialSizeMB; Maximum = $MaximumSizeMB }
    $pagefileOk  = $SkipPagefile -or (
        -not $pagefile.AutomaticManaged -and
        @($pagefile.Configured | Where-Object {
            $_.Name -eq $pagefilePath -and $_.Initial -eq $InitialSizeMB -and $_.Maximum -eq $MaximumSizeMB
        }).Count -eq 1 -and
        ($KeepSystemDrivePagefile -or @($pagefile.Configured | Where-Object { $_.Name -ne $pagefilePath }).Count -eq 0)
    )

    $reasons = @()
    if (-not $diskReady)  { $reasons += "$targetRoot is not a fixed volume" }
    if (-not $pagefileOk) { $reasons += 'the pagefile is not configured on it' }

    # With -RestartIfNeeded a correctly configured pagefile that this session is not
    # using is work too: the restart is what puts it in use. Without the switch it is
    # only worth a warning, and the run below stops at "nothing to do".
    if ($RestartIfNeeded -and -not $SkipPagefile -and $pagefile.Active -notcontains $pagefilePath) {
        $reasons += "$pagefilePath is configured but not in use in this session"
    }

    if ($reasons.Count -eq 0) {
        Write-Out ''
        Write-Skip 'Nothing to do - the temp disk is in place and the pagefile is configured on it'
        if (-not $SkipPagefile -and $pagefile.Active -notcontains $pagefilePath) {
            # Worth saying out loud even on an otherwise silent run: the machine is
            # configured correctly and still paging somewhere else until it restarts.
            Show-HeldOutput
            Write-Warn "The configured pagefile $pagefilePath is not in use in this session - it is created at the next restart"
        }
        if (-not $script:holdOutput) { Write-Host '' }
        Write-Log '--- nothing to do ---'
        exit 0
    }

    if ($CheckOnly) {
        Show-HeldOutput
        Write-Out ''
        Write-News ('Work is due: {0} (exit code 2)' -f ($reasons -join ', '))
        Write-Host ''
        exit 2
    }

    Show-HeldOutput

    # -- 3. Free the drive letter ----------------------------------------------
    Write-Out ''
    Write-Step "2. Drive letter $targetRoot"
    if ($diskReady) {
        Write-Skip 'Already held by the temp disk'
    } else {
        $optical = Get-OpticalDriveOnLetter -Root $targetRoot
        if (-not $optical) {
            Write-Skip 'Free - nothing to move out of the way'
        } else {
            $newLetter = Get-FreeDriveLetter -Preferred $OpticalDriveLetter
            if (-not $newLetter) {
                throw "An optical drive holds $targetRoot and there is no free letter left to move it to."
            }
            if ($PSCmdlet.ShouldProcess("optical drive $targetRoot", "Move to ${newLetter}:")) {
                Set-CimInstance -InputObject $optical -Property @{ DriveLetter = "${newLetter}:" }
                if (Get-OpticalDriveOnLetter -Root $targetRoot) {
                    throw "Could not move the optical drive off $targetRoot."
                }
                $changed = $true
                Write-Ok "Moved the optical drive from $targetRoot to ${newLetter}:"
            }
        }
    }

    # -- 4. The disk itself ----------------------------------------------------
    Write-Out ''
    Write-Step '3. Temp disk'
    if ($diskReady) {
        Write-Skip "$targetRoot is already a formatted fixed volume"
    } elseif ($tempPartition) {
        # The disk survived, only its drive letter did not. Formatting here would
        # throw away a perfectly good volume, so only the letter is put back.
        if ($PSCmdlet.ShouldProcess("disk $($tempPartition.DiskNumber) partition $($tempPartition.PartitionNumber)",
                                    "Assign drive letter $targetRoot")) {
            Set-Partition -InputObject $tempPartition -NewDriveLetter $targetLetter
            $changed = $true
            Write-Ok "Gave the existing temp volume its drive letter back ($targetRoot)"
        }
    } else {
        $candidates = @(Get-TempDiskCandidate)
        $chosen     = $null

        if ($DiskNumber -ge 0) {
            $chosen = Get-Disk -Number $DiskNumber -ErrorAction SilentlyContinue
            if (-not $chosen) { throw "Disk $DiskNumber does not exist." }
            if ($chosen.IsBoot -or $chosen.IsSystem) {
                throw "Disk $DiskNumber is the boot or system disk - refusing to format it."
            }
            if ($chosen.PartitionStyle -ne 'RAW' -and -not $Force) {
                throw "Disk $DiskNumber is $($chosen.PartitionStyle), not RAW - it already carries partitions. Add -Force if everything on it may be lost."
            }
        } elseif ($candidates.Count -eq 1) {
            $chosen = $candidates[0]
        } elseif ($candidates.Count -eq 0) {
            throw "No RAW disk to format as the temp disk, and no existing volume labelled '$Label'. On an Azure VM that means the resource disk was not attached - check the VM size."
        } else {
            $list = ($candidates | ForEach-Object { "disk $($_.Number) ($([math]::Round($_.Size / 1GB, 0)) GB, bus $($_.BusType))" }) -join ', '
            throw "More than one RAW disk could be the temp disk: $list. Pass -DiskNumber to say which one - formatting the wrong one costs a data disk."
        }

        $target = "disk $($chosen.Number) ($($chosen.FriendlyName), $([math]::Round($chosen.Size / 1GB, 0)) GB, bus $($chosen.BusType))"
        if ($PSCmdlet.ShouldProcess($target, "Initialise GPT, partition and format NTFS as $targetRoot")) {
            if ($chosen.IsOffline) {
                Set-Disk -Number $chosen.Number -IsOffline $false
                Write-Ok "Brought disk $($chosen.Number) online"
            }
            if ($chosen.IsReadOnly) {
                Set-Disk -Number $chosen.Number -IsReadOnly $false
                Write-Ok "Cleared the read-only flag on disk $($chosen.Number)"
            }

            $chosen = Get-Disk -Number $chosen.Number
            if ($chosen.PartitionStyle -ne 'RAW') {
                # Only reachable through -DiskNumber -Force; the automatic path never
                # picks a disk that has partitions on it.
                Clear-Disk -Number $chosen.Number -RemoveData -RemoveOEM -Confirm:$false
            }
            Initialize-Disk -Number $chosen.Number -PartitionStyle GPT -Confirm:$false

            $partition = New-Partition -DiskNumber $chosen.Number -UseMaximumSize -DriveLetter $targetLetter
            Format-Volume -Partition $partition -FileSystem NTFS -NewFileSystemLabel $Label -Confirm:$false -Force | Out-Null
            $changed = $true
            Write-Ok "Formatted disk $($chosen.Number) as $targetRoot with label '$Label'"
        }
    }

    # -- 5. Pagefile -----------------------------------------------------------
    Write-Out ''
    Write-Step '4. Pagefile'
    if ($SkipPagefile) {
        Write-Skip 'Skipped (-SkipPagefile)'
    } elseif ($pagefileOk) {
        Write-Skip "Already configured as $(Format-PagefileEntry $wantedEntry)"
    } elseif (-not $simulate -and -not (Test-Path -LiteralPath "$targetRoot\")) {
        # Configuring a pagefile on a drive that is not there writes a setting Windows
        # ignores, which reads as success and is not one.
        throw "$targetRoot is still not available, so the pagefile cannot be configured on it."
    } else {
        $system = Get-CimInstance -ClassName Win32_ComputerSystem
        if ($system.AutomaticManagedPagefile) {
            if ($PSCmdlet.ShouldProcess('Win32_ComputerSystem', 'Turn automatic pagefile management off')) {
                Set-CimInstance -InputObject $system -Property @{ AutomaticManagedPagefile = $false }
                $changed = $true
                Write-Ok 'Automatic pagefile management turned off'
            }
        }

        if (-not $KeepSystemDrivePagefile) {
            foreach ($entry in @(Get-CimInstance -ClassName Win32_PageFileSetting -ErrorAction SilentlyContinue |
                                 Where-Object { $_.Name -ne $pagefilePath })) {
                if ($PSCmdlet.ShouldProcess($entry.Name, 'Remove the pagefile entry')) {
                    Remove-CimInstance -InputObject $entry
                    $changed = $true
                    Write-Ok "Removed the pagefile entry for $($entry.Name)"
                }
            }
        }

        if ($PSCmdlet.ShouldProcess($pagefilePath, "Configure as $(Format-PagefileEntry $wantedEntry)")) {
            Set-PagefileEntry -Path $pagefilePath -Initial $InitialSizeMB -Maximum $MaximumSizeMB
            $changed = $true
            Write-Ok "Pagefile configured: $(Format-PagefileEntry $wantedEntry)"
        }
    }

    # -- 6. Verify -------------------------------------------------------------
    Write-Out ''
    Write-Step '5. Verification'
    if ($simulate) {
        Write-Skip 'Skipped - nothing was changed, so there is nothing to verify (-WhatIf)'
        Write-Out ''
        Write-Out '  Dry run only - rerun without -WhatIf to apply these changes.' 'Yellow'
    } else {
        $volumeNow = Get-Volume -DriveLetter $targetLetter -ErrorAction SilentlyContinue
        if ($volumeNow -and $volumeNow.DriveType -eq 'Fixed') {
            Write-Ok ('{0} is available: label "{1}", {2} GB' -f $targetRoot, $volumeNow.FileSystemLabel,
                      [math]::Round($volumeNow.Size / 1GB, 1))
        } else {
            Write-Bad "$targetRoot is still not a fixed volume"
            $exitCode = 1
        }

        if (-not $SkipPagefile) {
            $after = Get-PagefileState
            if ($after.AutomaticManaged) {
                Write-Bad 'Windows still manages the pagefile automatically'
                $exitCode = 1
            }

            $configured = @($after.Configured | Where-Object { $_.Name -eq $pagefilePath })
            if ($configured.Count -eq 1) {
                Write-Ok "Pagefile configured on the temp disk: $(Format-PagefileEntry $configured[0])"
            } else {
                Write-Bad "The pagefile is not configured as $pagefilePath"
                $exitCode = 1
            }

            foreach ($entry in @($after.Configured | Where-Object { $_.Name -ne $pagefilePath })) {
                if ($KeepSystemDrivePagefile) { Write-Skip "Left in place: $(Format-PagefileEntry $entry)" }
                else                          { Write-Warn "Another pagefile is still configured: $(Format-PagefileEntry $entry)" }
            }

            if ($after.Active -contains $pagefilePath) {
                Write-Ok 'The pagefile on the temp disk is in use in this session'
            } else {
                $rebootPending = $true
                Write-Warn "$pagefilePath is configured but not in use yet - Windows reads the pagefile at boot, so it is created at the next restart"
            }
        }

        Write-Out ''
        if ($exitCode -eq 0) {
            if ($changed) { Write-News 'Done - the temp disk and the pagefile were repaired' }
            else          { Write-Ok 'Done' }
        }
    }

    # -- 7. Restart ------------------------------------------------------------
    # Everything above is idempotent and can be repeated for free. This cannot, so
    # it is the one step that has to prove it will help before it acts.
    Write-Out ''
    Write-Step '6. Restart'
    if (-not $RestartIfNeeded) {
        if ($rebootPending) { Write-Skip 'Not restarting - use -RestartIfNeeded to have the pagefile put in use without waiting for the next boot' }
        else                { Write-Skip 'Not needed' }
    } elseif ($simulate) {
        Write-Skip 'Skipped - a -WhatIf run never restarts anything'
    } elseif (-not $rebootPending) {
        Write-Skip 'Not needed - the pagefile on the temp disk is already in use'
    } elseif ($exitCode -ne 0) {
        # A restart here would replace a visible failure with a reboot and a device
        # that comes back just as broken.
        Write-Warn 'Not restarting: something above failed, and a restart would only hide it'
    } else {
        $signedIn = @(Get-InteractiveSession)
        $marker   = Get-RestartMarker
        $since    = if ($marker) { (Get-Date) - $marker } else { $null }

        if ($signedIn.Count -gt 0 -and -not $RestartEvenIfUsersSignedIn) {
            Write-Warn ("Not restarting: {0} signed in ({1}) - drain the machine, or pass -RestartEvenIfUsersSignedIn. The pagefile lands at their next restart either way" -f
                        $(if ($signedIn.Count -eq 1) { 'someone is' } else { "$($signedIn.Count) users are" }), ($signedIn -join ', '))
        } elseif ($since -and $since.TotalMinutes -lt $RestartCooldownMinutes) {
            # The second restart for the same thing is the first symptom of a loop.
            Write-Warn ('Not restarting: this script already restarted the machine {0:N0} minute(s) ago ({1:yyyy-MM-dd HH:mm}) and the pagefile is still not in use - restarting again would only repeat that. Check the log at {2}' -f
                        $since.TotalMinutes, $marker, $LogPath)
        } else {
            # The countdown is there to warn people. Nobody signed in - the normal
            # case at boot, and the guaranteed one on a host whose pool is drained -
            # means there is nobody to warn, and a minute of waiting is a minute the
            # host is not available. A few seconds is kept so this run's own log line
            # is written before the shutdown starts.
            $delay = if ($signedIn.Count -gt 0) { $RestartDelaySeconds } else { [math]::Min(5, $RestartDelaySeconds) }
            if ($signedIn.Count -gt 0) {
                Write-Warn ('Restarting with {0} signed in (-RestartEvenIfUsersSignedIn), {1} second countdown: {2}' -f
                            $signedIn.Count, $delay, ($signedIn -join ', '))
            }
            $notice = "The temp disk was restored; this computer restarts in order to put the pagefile on $targetRoot."
            if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Restart in $delay seconds so $pagefilePath is created")) {
                # The guard goes down before the thing it guards, not after: a restart
                # that happens without its marker having been written is a restart
                # nothing can stop from happening again at the next boot. A marker
                # left behind by a restart that then failed to schedule only costs one
                # cooldown, which is the cheaper of the two mistakes.
                $guarded = $false
                try { Set-RestartMarker; $guarded = $true }
                catch { Write-Bad "Not restarting: the restart marker could not be written ($($_.Exception.Message)), and without it nothing limits a restart loop"; $exitCode = 1 }

                if ($guarded) {
                    # p:2:4 is the planned "Operating System: Reconfiguration" reason,
                    # so the restart shows up as intended rather than unexpected in the
                    # System log and in Azure's own restart reporting.
                    $scheduled = $false
                    try {
                        $null = & shutdown.exe /r /t $delay /c $notice /d p:2:4
                        $scheduled = ($LASTEXITCODE -eq 0)
                    } catch {
                        # From PowerShell 7.4 a non-zero exit code from a native
                        # command can surface as a terminating error instead, through
                        # $PSNativeCommandUseErrorActionPreference. Measured on 7.6 it
                        # does not, and $LASTEXITCODE carries the failure as it always
                        # has - but the same failure arriving by the other door must
                        # not read as a restart that was scheduled.
                        $scheduled = $false
                    }

                    if ($scheduled) {
                        $restartTriggered = $true
                        Write-News "Restarting in $delay seconds so $pagefilePath is created - 'shutdown /a' cancels it"
                    } else {
                        Write-Bad "shutdown.exe refused to schedule the restart (exit code $LASTEXITCODE)"
                        $exitCode = 1
                    }
                }
            }
        }
    }

    if ($changed -or $restartTriggered -or $exitCode -ne 0) { Show-HeldOutput }
} catch {
    Show-HeldOutput
    Write-Out ''
    Write-Bad "Aborted: $($_.Exception.Message)"
    $exitCode = 1
} finally {
    $tail = if ($restartTriggered) { ', restart triggered' } elseif ($rebootPending) { ', restart pending' } else { '' }
    Write-Log ('--- finished, exit code {0}{1} ---' -f $exitCode, $tail)
}

if (-not $script:holdOutput) { Write-Host '' }
exit $exitCode
