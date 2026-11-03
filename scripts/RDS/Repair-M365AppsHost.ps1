#Requires -Version 5.1
<#
.SYNOPSIS
    One go on a session host: repair everything the M365 app watchdog keeps reporting
    (Teams, new Outlook, Photos, Snipping Tool, Copilot), test it, and say whether the
    host needs a restart.

.DESCRIPTION
    For a host where Watch-M365Apps.ps1 is installed (C:\IT\AppWatchdog). Uses the
    scripts installed there - the locked folder System runs from - and changes nothing
    the watchdog itself would not, apart from step 2. In order:

      1. Before     A snapshot: FSLogix InstallAppxPackages, the provisioned build of
                    each app, per app how many of the signed-in users have it
                    installed (and who does not), failed package registrations in the
                    last 24 hours, and the signals that a restart is due.
      2. FSLogix    InstallAppxPackages = 0 under HKLM\SOFTWARE\FSLogix\Profiles, after
                    a .reg backup of that key. FSLogix then stops replaying the
                    packages saved in the profile at sign-in - the replay behind the
                    daily "FSLogix could not register Microsoft.OutlookForWindows"
                    (0x80073CF9 / 0x80070490) while Windows registers the provisioned
                    build anyway. Store apps users installed themselves no longer
                    follow them; provisioned apps are not affected. Takes effect at the
                    next sign-in. -KeepAppxReplay leaves it as it is.
      3. Repair     The watchdog task is held (disabled) so it cannot run a repair at
                    the same time, then Repair-AppxPackageStore.ps1: Teams and Outlook
                    with -Latest -Provision (the newest build, signature-checked), and
                    Photos and Snipping Tool without -Provision - their store is cleaned,
                    no Store download (that needs an Entra admin). No -RemoveOld:
                    nothing is taken away from a user who still has an older build open.
      4. Test       The watchdog task itself is started - as System, with its own
                    config.json, so it is exactly the run that happens every 30 minutes:
                    every signed-in user is checked, repaired in their session where
                    needed, and our own account (itceadmin) has each app started. Its
                    exit code and what it still reports are shown.
      5. After      The snapshot of step 1 again, side by side with before, and whether
                    a restart is due: Repair-AppxPackageStore edited the package store
                    since the last boot (it asks for a restart so the deployment engine
                    rereads it), or Windows has a restart pending (servicing, Windows
                    Update, file renames). -Restart restarts the host when that is so and
                    nobody is signed in; with users signed in it only says so - drain
                    the host first.

    -CheckOnly does steps 1 and 5 and a watchdog run with -NoRepair: nothing is changed.

.PARAMETER CheckOnly
    Only look: the snapshot, a watchdog run that repairs nothing, and the restart
    verdict.

.PARAMETER KeepAppxReplay
    Leave FSLogix InstallAppxPackages as it is (step 2 skipped).

.PARAMETER Restart
    Restart the host at the end when a restart is due and no user is signed in.

.PARAMETER WaitMinutes
    How long to wait for a watchdog run that is busy, and for the test run. Default 20.

.PARAMETER WorkingDir
    The watchdog's folder. Default C:\IT\AppWatchdog.

.EXAMPLE
    # What would be repaired, and is a restart due? Changes nothing.
    .\Repair-M365AppsHost.ps1 -CheckOnly

.EXAMPLE
    # Repair and test everything, without asking
    .\Repair-M365AppsHost.ps1 -Confirm:$false

.EXAMPLE
    # The same on a drained host, and restart it when that is due
    .\Repair-M365AppsHost.ps1 -Restart -Confirm:$false

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows 11 Enterprise multi-session / Windows Server with FSLogix, with
              Watch-M365Apps.ps1 installed; run elevated or as System
    Exit    : 0 healthy and no restart due, 1 something still broken or the run failed,
              2 healthy but a restart is due (and not done)
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param (
    [switch] $CheckOnly,
    [switch] $KeepAppxReplay,
    [switch] $Restart,
    [ValidateRange(5, 120)]
    [int]    $WaitMinutes = 20,
    [string] $WorkingDir = 'C:\IT\AppWatchdog'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -- Constants -------------------------------------------------------------------
$Apps = @(
    [PSCustomObject]@{ App = 'Teams';        Package = 'MSTeams' }
    [PSCustomObject]@{ App = 'Outlook';      Package = 'Microsoft.OutlookForWindows' }
    [PSCustomObject]@{ App = 'Photos';       Package = 'Microsoft.Windows.Photos' }
    [PSCustomObject]@{ App = 'SnippingTool'; Package = 'Microsoft.ScreenSketch' }
)
$TaskName         = 'M365 App Watchdog'
$FslogixProfiles  = 'HKLM:\SOFTWARE\FSLogix\Profiles'
$FslogixPolicy    = 'HKLM:\SOFTWARE\Policies\FSLogix\Profiles'
# Deployment results that are not a fault, as in the watchdog.
$BenignAppxCodes  = @('0x80073D02', '0x80073CFB', '0x80073D06')

# -- Output ------------------------------------------------------------------------
function Write-Step { param([string] $Message) Write-Host ''; Write-Host "  $Message" -ForegroundColor Cyan }
function Write-Ok   { param([string] $Message) Write-Host "  [ OK ] $Message" -ForegroundColor Green }
function Write-Skip { param([string] $Message) Write-Host "  [SKIP] $Message" -ForegroundColor DarkGray }
function Write-Warn { param([string] $Message) Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
function Write-Bad  { param([string] $Message) Write-Host "  [FAIL] $Message" -ForegroundColor Red }

function Get-PropertyValue {
    <# Property access that returns $null instead of tripping Set-StrictMode. #>
    param($Object, [string] $Name)
    if ($null -eq $Object) { return $null }
    $prop = $Object.PSObject.Properties[$Name]
    if ($prop) { return $prop.Value }
    return $null
}

function ConvertTo-ScriptCommand {
    <#
        A script call with its parameters as one -Command string. Not -File: that
        hands every argument over as text, so -Confirm:$false arrives as the string
        '$false'. Here PowerShell parses the line itself; strings are single-quoted,
        and the script's exit code is passed on.
    #>
    param([Parameter(Mandatory)] [string] $Path, [System.Collections.IDictionary] $Parameters = @{})
    $quote = { param($s) "'" + ([string] $s -replace "'", "''") + "'" }
    $line  = '& ' + (& $quote $Path)
    foreach ($entry in $Parameters.GetEnumerator()) {
        $value = $entry.Value
        if ($value -is [switch] -or $value -is [bool]) {
            $line += ' -{0}:${1}' -f $entry.Key, ([bool] $value).ToString().ToLower()
        } elseif ($value -is [array]) {
            $line += ' -{0} {1}' -f $entry.Key, (($value | ForEach-Object { & $quote $_ }) -join ',')
        } else {
            $line += ' -{0} {1}' -f $entry.Key, (& $quote $value)
        }
    }
    return $line + '; exit $LASTEXITCODE'
}

# -- Windows PowerShell, 64-bit, elevated --------------------------------------------
# The AppX cmdlets need Windows PowerShell; under WOW64 HKLM lands in WOW6432Node.
$nativeShell = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $nativeShell = Join-Path $env:WINDIR 'SysNative\WindowsPowerShell\v1.0\powershell.exe'
}
if ($PSVersionTable.PSEdition -eq 'Core' -or ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess)) {
    & $nativeShell -NoProfile -ExecutionPolicy Bypass -Command (ConvertTo-ScriptCommand -Path $PSCommandPath -Parameters $PSBoundParameters)
    exit $LASTEXITCODE
}
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not ([Security.Principal.WindowsPrincipal] $identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error 'Administrator rights are required. Run the script as System or from an elevated session.'
    exit 1
}

# -- Readers -------------------------------------------------------------------------
function Get-SignedInUser {
    <# Everyone with an explorer.exe on this host: SID and DOMAIN\user. #>
    $users = @{}
    foreach ($process in @(Get-CimInstance Win32_Process -Filter "Name = 'explorer.exe'")) {
        $sid = (Invoke-CimMethod -InputObject $process -MethodName GetOwnerSid -ErrorAction SilentlyContinue).Sid
        if (-not $sid -or $users.ContainsKey($sid)) { continue }
        $owner = Invoke-CimMethod -InputObject $process -MethodName GetOwner -ErrorAction SilentlyContinue
        $users[$sid] = if ($owner -and $owner.User) { '{0}\{1}' -f $owner.Domain, $owner.User } else { $sid }
    }
    return $users
}

function Get-AppxReplay {
    <# FSLogix InstallAppxPackages as set, and what a policy key says about it. #>
    $value  = Get-PropertyValue (Get-ItemProperty $FslogixProfiles -ErrorAction SilentlyContinue) 'InstallAppxPackages'
    $policy = Get-PropertyValue (Get-ItemProperty $FslogixPolicy -ErrorAction SilentlyContinue) 'InstallAppxPackages'
    [PSCustomObject]@{ Value = $value; Policy = $policy
                       Text  = if ($null -eq $value) { 'not set (on)' } elseif ($value -eq 0) { '0 (off)' } else { "$value (on)" } }
}

function Get-RegisterFailure {
    <#
        Failed registrations of the watched packages since $Since (AppXDeploymentServer
        401/404), one per user, package and second - Windows writes both events for one
        failure - per app.
    #>
    param([datetime] $Since)
    $seen = @{}
    foreach ($record in @(Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-AppXDeploymentServer/Operational'; Id = 401, 404; StartTime = $Since } -ErrorAction SilentlyContinue)) {
        $values = @($record.Properties | ForEach-Object { $_.Value })
        $codeAt = if ($record.Id -eq 401) { 3 } else { 2 }
        if ($values.Count -le $codeAt) { continue }
        $code = '0x{0:X8}' -f ([long] $values[$codeAt] -band [long] 4294967295)
        if ($code -in $BenignAppxCodes) { continue }
        $entry = @($Apps | Where-Object { [string] $values[1] -like "$($_.Package)_*" }) | Select-Object -First 1
        if (-not $entry) { continue }
        $seen['{0}|{1}|{2:yyyyMMddHHmmss}' -f $record.UserId, $values[1], $record.TimeCreated] = $entry.App
    }
    $count = @{}
    foreach ($app in $seen.Values) { $count[$app] = 1 + [int] $count[$app] }
    return $count
}

function Get-HostSnapshot {
    <# Per app: the provisioned build, and how many signed-in users have it installed. #>
    param([hashtable] $Users)
    $provisioned = @(Get-AppxProvisionedPackage -Online)
    $failures    = Get-RegisterFailure -Since (Get-Date).AddHours(-24)
    foreach ($entry in $Apps) {
        $build = @($provisioned | Where-Object { $_.DisplayName -eq $entry.Package } | Sort-Object { [version] $_.Version } -Descending) | Select-Object -First 1
        $known = @(Get-AppxPackage -AllUsers -Name $entry.Package -ErrorAction SilentlyContinue | ForEach-Object { $_.PackageUserInformation } |
                   Where-Object { [string] $_.InstallState -eq 'Installed' } | ForEach-Object { $_.UserSecurityId.Sid })
        $missing = @($Users.Keys | Where-Object { $_ -notin $known } | ForEach-Object { $Users[$_] } | Sort-Object)
        [PSCustomObject]@{
            App         = $entry.App
            Provisioned = if ($build) { $build.Version } else { 'NOT provisioned' }
            Users       = '{0}/{1}' -f ($Users.Count - $missing.Count), $Users.Count
            Missing     = $missing -join ', '
            Failed24h   = [int] $failures[$entry.App]
        }
    }
}

function Get-RestartReason {
    <# Why this host should restart, one line per reason; nothing when it need not. #>
    $boot    = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
    $reasons = [System.Collections.Generic.List[string]]::new()
    # Repair-AppxPackageStore backs up every store key it removes, in a folder per run.
    $edits = @(Get-ChildItem -Path (Join-Path $WorkingDir 'Repair') -Directory -Filter 'Repair-AppxPackageStore_*' -ErrorAction SilentlyContinue |
               Where-Object { @(Get-ChildItem -Path $_.FullName -Filter '*.reg' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -gt $boot }).Count -gt 0 })
    if ($edits.Count -gt 0) {
        $reasons.Add(('the package store was edited {0}x since the last boot ({1:yyyy-MM-dd HH:mm}), last at {2:yyyy-MM-dd HH:mm} - the deployment engine rereads it at a restart' -f
                      $edits.Count, $boot, ($edits | Sort-Object LastWriteTime | Select-Object -Last 1).LastWriteTime))
    }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $reasons.Add('Windows servicing has a restart pending') }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $reasons.Add('Windows Update has a restart pending') }
    return $reasons
}

function Write-Snapshot {
    param([string] $Title, $Snapshot, $Replay)
    Write-Step $Title
    Write-Host "  FSLogix InstallAppxPackages: $($Replay.Text)"
    $Snapshot | Format-Table App, Provisioned, Users, Failed24h, Missing -AutoSize -Wrap | Out-String -Width 200 | Write-Host
}

# -- The watchdog --------------------------------------------------------------------
function Wait-WatchdogIdle {
    <# $true when the watchdog task is not running (any more), $false after -WaitMinutes. #>
    $until = (Get-Date).AddMinutes($WaitMinutes)
    while ((Get-ScheduledTask -TaskName $TaskName).State -eq 'Running') {
        if ((Get-Date) -gt $until) { return $false }
        Start-Sleep -Seconds 10
    }
    return $true
}

function Invoke-WatchdogRun {
    <#
        The watchdog task itself, started now: as System, with its own config.json.
        Returns its exit code (the task's last result) and the run's lines from its log.
    #>
    $started = Get-Date
    Start-ScheduledTask -TaskName $TaskName
    $until = $started.AddMinutes($WaitMinutes)
    do {
        Start-Sleep -Seconds 10
        $info = Get-ScheduledTaskInfo -TaskName $TaskName
        $done = $info.LastRunTime -ge $started.AddSeconds(-5) -and (Get-ScheduledTask -TaskName $TaskName).State -ne 'Running'
    } until ($done -or (Get-Date) -gt $until)
    if (-not $done) { return [PSCustomObject]@{ Code = $null; Lines = @() } }
    return [PSCustomObject]@{ Code = [int] $info.LastTaskResult; Lines = @(Get-WatchdogLog -Since $started.AddSeconds(-5)) }
}

function Get-WatchdogLog {
    <# The lines of the last watchdog run in its log, when it started after $Since. #>
    param([datetime] $Since)
    $log = Join-Path $WorkingDir ('Logs\Watch-M365Apps_{0:yyyyMMdd}.log' -f $Since)
    if (-not (Test-Path $log)) { return }
    $text  = Get-Content -Path $log -Raw
    $start = $text.LastIndexOf('Windows PowerShell transcript start')
    if ($start -lt 0) { return }
    $run = $text.Substring($start)
    if ($run -match 'Start time: (\d{14})' -and [datetime]::ParseExact($Matches[1], 'yyyyMMddHHmmss', $null) -lt $Since) { return }
    return $run -split '\r?\n'
}

# -- One run -------------------------------------------------------------------------
$watchdog = Join-Path $WorkingDir 'Watch-M365Apps.ps1'
$repair   = Join-Path $WorkingDir 'Repair-AppxPackageStore.ps1'
foreach ($path in $watchdog, $repair) {
    if (-not (Test-Path $path)) { Write-Error "$path not found - install the watchdog on this host first (Watch-M365Apps.ps1 -Install)"; exit 1 }
}
$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if (-not $task) { Write-Error "Scheduled task '$TaskName' not found - install the watchdog on this host first (Watch-M365Apps.ps1 -Install)"; exit 1 }
# Only the apps the watchdog was installed to watch: one left out on purpose is not a fault.
$configPath = Join-Path $WorkingDir 'config.json'
if (Test-Path $configPath) {
    $watched = @(Get-PropertyValue (Get-Content -Path $configPath -Raw | ConvertFrom-Json) 'App')
    if ($watched.Count -gt 0) { $Apps = @($Apps | Where-Object { $_.App -in $watched }) }
}
# -WhatIf: the test run would be the task's own, which repairs - look only.
if ($WhatIfPreference) { $CheckOnly = [switch] $true }

$logDir = Join-Path $WorkingDir 'Logs'
New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false | Out-Null
$transcribing = $false
try {
    Start-Transcript -Path (Join-Path $logDir ('Repair-M365AppsHost_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))) -WhatIf:$false | Out-Null
    $transcribing = $true
} catch {
    Write-Warn "No transcript: $($_.Exception.Message)"
}

$exitCode = 0
$held     = $false
try {
    Write-Host ''
    Write-Host ("  M365 apps - repair and test - {0} - {1:yyyy-MM-dd HH:mm}{2}" -f $env:COMPUTERNAME, (Get-Date), $(if ($CheckOnly) { ' - check only' } else { '' })) -ForegroundColor Cyan

    # -- 1. Before ---------------------------------------------------------------
    $users        = Get-SignedInUser
    $replayBefore = Get-AppxReplay
    $before       = @(Get-HostSnapshot -Users $users)
    Write-Snapshot -Title ('1. Before - {0} user(s) signed in' -f $users.Count) -Snapshot $before -Replay $replayBefore
    foreach ($reason in @(Get-RestartReason)) { Write-Warn "Restart due: $reason" }

    if (-not (Wait-WatchdogIdle)) { throw "the watchdog task is still running after $WaitMinutes minutes - try again later" }

    if ($CheckOnly) {
        Write-Step '4. Test - the watchdog, repairing nothing'
        & $nativeShell -NoProfile -ExecutionPolicy Bypass -Command (ConvertTo-ScriptCommand -Path $watchdog -Parameters ([ordered]@{ WorkingDir = $WorkingDir; NoRepair = $true })) | Out-Host
        $test = [PSCustomObject]@{ Code = $LASTEXITCODE; Lines = @() }
    } else {
        # -- 2. FSLogix ----------------------------------------------------------
        Write-Step '2. FSLogix - stop replaying the packages saved in the profile'
        if ($KeepAppxReplay) {
            Write-Skip "-KeepAppxReplay: InstallAppxPackages stays $($replayBefore.Text)"
        } elseif ($replayBefore.Value -eq 0) {
            Write-Ok 'InstallAppxPackages is already 0'
        } elseif ($PSCmdlet.ShouldProcess("$env:COMPUTERNAME - $FslogixProfiles\InstallAppxPackages = 0", 'Set')) {
            $backupDir = Join-Path $WorkingDir 'Repair'
            New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
            $backup = Join-Path $backupDir ('FSLogix-Profiles_{0:yyyyMMdd_HHmmss}.reg' -f (Get-Date))
            if (Test-Path $FslogixProfiles) {
                & reg.exe export 'HKLM\SOFTWARE\FSLogix\Profiles' $backup /y | Out-Null
                if ($LASTEXITCODE -ne 0) { throw "$FslogixProfiles not backed up - not changed" }
            } else {
                New-Item -Path $FslogixProfiles -Force | Out-Null
            }
            New-ItemProperty -Path $FslogixProfiles -Name 'InstallAppxPackages' -Value 0 -PropertyType DWord -Force | Out-Null
            Write-Ok "InstallAppxPackages = 0 - from the next sign-in (backup $backup)"
        }
        if ($null -ne $replayBefore.Policy -and $replayBefore.Policy -ne 0) {
            Write-Warn "$FslogixPolicy sets InstallAppxPackages = $($replayBefore.Policy) - change it in that policy (GPO / Intune), or it may come back"
        }

        # -- 3. Repair -----------------------------------------------------------
        Write-Step '3. Repair - the package store, Teams and Outlook at the newest build'
        if ($PSCmdlet.ShouldProcess("$env:COMPUTERNAME - Repair-AppxPackageStore for Teams, Outlook, Photos, Snipping Tool", 'Repair')) {
            # Held, so the task cannot start a repair of its own halfway through ours.
            Disable-ScheduledTask -TaskName $TaskName | Out-Null
            $held = $true
            # A trigger may have started it just before: disabling does not stop a run.
            if (-not (Wait-WatchdogIdle)) { throw "the watchdog task is still running after $WaitMinutes minutes - try again later" }
            $repairDir = Join-Path $WorkingDir 'Repair'
            $latest    = @(@{ Teams = 'teams'; Outlook = 'outlook' }.GetEnumerator() | Where-Object { $_.Key -in $Apps.App } | ForEach-Object { $_.Value })
            $storeOnly = @(@{ Photos = 'photos'; SnippingTool = 'snippingtool' }.GetEnumerator() | Where-Object { $_.Key -in $Apps.App } | ForEach-Object { $_.Value })
            foreach ($call in @(
                if ($latest.Count) { [ordered]@{ Name = $latest; Latest = $true; Provision = $true; Days = 1; WorkingDir = $repairDir; LogPath = $repairDir; Confirm = $false } }
                if ($storeOnly.Count) { [ordered]@{ Name = $storeOnly; Days = 1; WorkingDir = $repairDir; LogPath = $repairDir; Confirm = $false } }
            )) {
                $command = ConvertTo-ScriptCommand -Path $repair -Parameters $call
                Write-Step $command
                & $nativeShell -NoProfile -ExecutionPolicy Bypass -Command $command | Out-Host
                if ($LASTEXITCODE -eq 0) { Write-Ok "Repair-AppxPackageStore $($call.Name -join ',') - clean" }
                else { Write-Warn "Repair-AppxPackageStore $($call.Name -join ',') exited with $LASTEXITCODE - see its output above" }
            }
            Enable-ScheduledTask -TaskName $TaskName | Out-Null
            $held = $false
        }

        # -- 4. Test -------------------------------------------------------------
        Write-Step '4. Test - one watchdog run, as the task runs it'
        $test = Invoke-WatchdogRun
    }

    if ($null -eq $test.Code) {
        Write-Bad "The watchdog run did not finish within $WaitMinutes minutes - see $logDir"
        $exitCode = 1
    } else {
        foreach ($line in @($test.Lines | Where-Object { $_ -match '\[(FAIL|WARN)\]|Not signed in on this host|Everything (healthy|repaired)|problem\(s\) left' })) {
            Write-Host "    $($line.Trim())"
        }
        if ($test.Lines -match 'Not signed in on this host') {
            Write-Warn 'Our own account is not signed in on this host: the apps were not started, only registrations checked - keep a session of it open'
        }
        if ($test.Code -eq 0) { Write-Ok 'Watchdog: healthy (or repaired)' } else { Write-Bad "Watchdog: exit code $($test.Code) - something is still broken"; $exitCode = 1 }
    }

    # -- 5. After ----------------------------------------------------------------
    $users = Get-SignedInUser
    $after = @(Get-HostSnapshot -Users $users)
    Write-Snapshot -Title ('5. After - {0} user(s) signed in' -f $users.Count) -Snapshot $after -Replay (Get-AppxReplay)
    foreach ($row in $after) {
        if ($row.Provisioned -eq 'NOT provisioned') { Write-Bad "$($row.App) is not provisioned - new profiles do not get it"; $exitCode = 1 }
        if ($row.Missing) { Write-Bad "$($row.App) not installed for: $($row.Missing)"; $exitCode = 1 }
    }
    Write-Skip 'Failed24h counts the last 24 hours, so it only drops as sign-ins without failures come in - compare it tomorrow'

    $reasons = @(Get-RestartReason)
    Write-Step 'Restart'
    if ($reasons.Count -eq 0) {
        Write-Ok 'Not needed'
    } else {
        foreach ($reason in $reasons) { Write-Warn "Due: $reason" }
        if (-not $Restart -or $CheckOnly) {
            Write-Skip 'Not restarted - drain the host and run again with -Restart, or restart it out of hours'
            if ($exitCode -eq 0) { $exitCode = 2 }
        } elseif ($users.Count -gt 0) {
            Write-Warn ('Not restarted: {0} user(s) signed in ({1}) - drain the host first' -f $users.Count, (@($users.Values | Sort-Object) -join ', '))
            if ($exitCode -eq 0) { $exitCode = 2 }
        } elseif ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, 'Restart')) {
            Write-Ok 'Restarting in 30 seconds'
            if ($transcribing) { try { Stop-Transcript | Out-Null } catch { }; $transcribing = $false }
            & shutdown.exe /r /t 30 /c 'Repair-M365AppsHost: package store edited, restart so the deployment engine rereads it' /d p:4:1
        } else {
            if ($exitCode -eq 0) { $exitCode = 2 }
        }
    }
} catch {
    Write-Bad ('Aborted at line {0}: {1}' -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message)
    $exitCode = 1
} finally {
    # Never leave the watchdog switched off.
    if ($held) { try { Enable-ScheduledTask -TaskName $TaskName | Out-Null } catch { Write-Bad "Task '$TaskName' could not be enabled again - enable it by hand" } }
    if ($transcribing) { try { Stop-Transcript | Out-Null } catch { } }
}
exit $exitCode
