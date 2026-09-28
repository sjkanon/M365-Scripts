#Requires -Version 5.1
<#
.SYNOPSIS
    Install Init-TempDisk.ps1 on the device and register it as a scheduled task that
    runs at every boot as SYSTEM. Supports -WhatIf.

.DESCRIPTION
    Run once per device - by hand, from Tactical RMM / NinjaOne, or as an Intune
    platform script. It copies Init-TempDisk.ps1 to a local folder (the repository is
    not there at boot time) and registers a task that runs it at every startup, as
    SYSTEM, with the highest privileges and without needing anyone to sign in.

    The task exists because the ephemeral temp disk is wiped on every deallocate,
    resize or host move: D: has to be re-created and the pagefile pointed back at it
    before the machine is in the state it was deployed in. Init-TempDisk.ps1 does
    that work and is silent when there is nothing to do, so a boot on a healthy
    machine leaves no trace beyond its log line.

    The task runs it with -Quiet -RestartIfNeeded by default. Windows reads the
    pagefile configuration at boot, so a boot on which the temp disk had to be
    rebuilt runs that whole session without a pagefile on D: unless the machine
    restarts once. Init-TempDisk.ps1 only restarts when it demonstrably helps: the
    run finished clean, the disk is there, the pagefile is configured on it and only
    this session is not using it - and never while someone is signed in, and at most
    once an hour. Pass -ScriptArguments '-Quiet' to leave the restart out.

    Re-running this script is safe: an existing task with the same name is replaced,
    so it is also the way to change the arguments the task runs with.

    -Unregister is the other direction: the task goes, the copied script goes with it
    and the pagefile configuration is deliberately left exactly as it is - removing
    the task must not take a machine's pagefile away with it.

.PARAMETER ScriptSourcePath
    Init-TempDisk.ps1 to install (default: the copy next to this script).

.PARAMETER ScriptTargetDir
    Folder it is copied to on the device (default: C:\Scripts). The task runs the
    copy, never the source, so the repository or the RMM staging folder can go away.

.PARAMETER TaskName
    Name of the scheduled task (default: InitTempDisk).

.PARAMETER ScriptArguments
    Arguments for Init-TempDisk.ps1 (default: '-Quiet -RestartIfNeeded' - silent on a
    healthy boot, and one restart when that is what the pagefile is waiting for).
    Pass '-Quiet' to leave the restart out, or e.g.
    '-Quiet -RestartIfNeeded -InitialSizeMB 16384 -MaximumSizeMB 16384' for a fixed
    pagefile size.

.PARAMETER DelaySeconds
    Delay between boot and the task starting (default: 30). The storage stack does
    not always have the temp disk enumerated the instant the task engine is up.

.PARAMETER RunNow
    Also start the task once, right after registering it, so the device is in the
    intended state without waiting for a reboot.

.PARAMETER Unregister
    Remove the task and the installed copy of the script instead of registering them.

.EXAMPLE
    .\Register-InitTempDiskTask.ps1 -WhatIf

    Show what would be installed and registered, without touching the device.

.EXAMPLE
    .\Register-InitTempDiskTask.ps1 -RunNow

    Install the script, register the boot task and run it once immediately.

.EXAMPLE
    .\Register-InitTempDiskTask.ps1 -ScriptArguments '-Quiet -RestartIfNeeded -InitialSizeMB 16384 -MaximumSizeMB 16384'

    Same, but with a fixed 16 GB pagefile instead of a system managed one.

.EXAMPLE
    .\Register-InitTempDiskTask.ps1 -ScriptArguments '-Quiet'

    Without the restart: the task repairs the disk and configures the pagefile, and
    the pagefile is put in use at whatever restart happens next.

.EXAMPLE
    .\Register-InitTempDiskTask.ps1 -Unregister

    Take the task and the installed script off the device again.

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows only, run as administrator or as System
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [string] $ScriptTargetDir  = 'C:\Scripts',
    [string] $ScriptSourcePath = (Join-Path $PSScriptRoot 'Init-TempDisk.ps1'),
    [string] $TaskName         = 'InitTempDisk',
    [string] $ScriptArguments  = '-Quiet -RestartIfNeeded',
    [ValidateRange(0, 3600)]
    [int]    $DelaySeconds     = 30,
    [switch] $RunNow,
    [switch] $Unregister
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-Elevated {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    return ([Security.Principal.WindowsPrincipal] $identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Write-Ok   { param([string] $Message) Write-Host "  [ OK ] $Message" -ForegroundColor Green }
function Write-Skip { param([string] $Message) Write-Host "  [SKIP] $Message" -ForegroundColor DarkGray }
function Write-Warn { param([string] $Message) Write-Host "  [WARN] $Message" -ForegroundColor Yellow }

if (-not (Test-Elevated)) {
    Write-Error 'Administrator rights are required. Run this from an elevated session or as System.'
    exit 1
}

$targetScript = Join-Path $ScriptTargetDir 'Init-TempDisk.ps1'
$exitCode     = 0

try {
    Write-Host ''

    # -- Remove ----------------------------------------------------------------
    if ($Unregister) {
        $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if (-not $existing) {
            Write-Skip "No scheduled task named '$TaskName'"
        } elseif ($PSCmdlet.ShouldProcess($TaskName, 'Unregister the scheduled task')) {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
            Write-Ok "Removed the scheduled task '$TaskName'"
        }

        if (-not (Test-Path -LiteralPath $targetScript)) {
            Write-Skip "No installed script at $targetScript"
        } elseif ($PSCmdlet.ShouldProcess($targetScript, 'Remove the installed script')) {
            Remove-Item -LiteralPath $targetScript -Force
            Write-Ok "Removed $targetScript"
        }

        Write-Skip 'The pagefile configuration is left as it is - removing the task must not take the pagefile with it'
        Write-Host ''
        exit 0
    }

    # -- Install the script ----------------------------------------------------
    # Copy failures were silent in the first version of this script, which is how a
    # task ends up registered against a file that is not there: it runs at every boot
    # and fails at every boot, without anyone noticing until the pagefile is gone.
    if (-not (Test-Path -LiteralPath $ScriptSourcePath)) {
        throw "Init-TempDisk.ps1 not found at $ScriptSourcePath - point -ScriptSourcePath at it."
    }

    $source = (Get-Item -LiteralPath $ScriptSourcePath).FullName
    if ($source -ieq $targetScript) {
        Write-Skip "Running from $ScriptTargetDir already - nothing to copy"
    } else {
        if (-not (Test-Path -LiteralPath $ScriptTargetDir)) {
            if ($PSCmdlet.ShouldProcess($ScriptTargetDir, 'Create the folder')) {
                New-Item -ItemType Directory -Path $ScriptTargetDir -Force | Out-Null
                Write-Ok "Created $ScriptTargetDir"
            }
        }
        if ($PSCmdlet.ShouldProcess($targetScript, "Copy from $source")) {
            Copy-Item -LiteralPath $source -Destination $targetScript -Force
            Write-Ok "Installed $targetScript"
        }
    }

    # -- Register the task -----------------------------------------------------
    $argument = '-NoProfile -ExecutionPolicy Bypass -File "{0}"' -f $targetScript
    if ($ScriptArguments) { $argument += ' ' + $ScriptArguments }

    $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argument
    $trigger   = New-ScheduledTaskTrigger -AtStartup
    if ($DelaySeconds -gt 0) { $trigger.Delay = 'PT{0}S' -f $DelaySeconds }

    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd `
                                              -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                                              -MultipleInstances IgnoreNew `
                                              -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        if ($PSCmdlet.ShouldProcess($TaskName, 'Unregister the existing task first')) {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
            Write-Skip "Replaced the existing task '$TaskName'"
        }
    }

    if ($PSCmdlet.ShouldProcess($TaskName, "Register a startup task running $targetScript as SYSTEM")) {
        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
                               -Principal $principal -Settings $settings `
                               -Description 'Restores the ephemeral temp disk (D:) and keeps the pagefile on it at every boot.' | Out-Null

        $registered = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if (-not $registered) { throw "The task '$TaskName' is not there after registering it." }
        Write-Ok "Registered '$TaskName' - runs at every boot as SYSTEM$(if ($DelaySeconds -gt 0) { ", $DelaySeconds seconds after startup" })"
    }

    # -- Run it once -----------------------------------------------------------
    if ($RunNow) {
        if ($PSCmdlet.ShouldProcess($TaskName, 'Start the task once now')) {
            Start-ScheduledTask -TaskName $TaskName
            Write-Ok 'Started the task - check C:\Temp\Init-TempDisk.log for what it did'
        }
    } else {
        Write-Skip 'Not run yet - it runs at the next boot, or add -RunNow to do it immediately'
    }

    # What the task may do to a machine is worth saying at registration time, not
    # only in the help of the script it runs.
    if ($ScriptArguments -match '(?i)-RestartIfNeeded') {
        Write-Warn 'A pagefile is read at boot, so this task may restart the machine once - only when the disk is in order, the pagefile is configured on it, nobody is signed in, and no restart was triggered in the last hour'
    } else {
        Write-Warn 'A pagefile is read at boot, so a pagefile this task configures is only in use after the next restart (-RestartIfNeeded in -ScriptArguments makes the task handle that itself)'
    }
} catch {
    Write-Host ''
    Write-Host "  [FAIL] Aborted: $($_.Exception.Message)" -ForegroundColor Red
    $exitCode = 1
}

Write-Host ''
exit $exitCode
