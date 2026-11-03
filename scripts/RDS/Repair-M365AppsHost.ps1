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
                    Teams, Outlook, Photos and Snipping Tool, per app how many of the
                    signed-in users have it installed (and who does not), failed
                    package registrations in the last 24 hours; for Copilot the unified
                    app Edge Update installs machine-wide (version, copilotapp.exe, a
                    Start menu entry for all users); and the signals that a restart is
                    due.
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
                    with -Latest -Provision (the newest build, signature-checked);
                    Photos and Snipping Tool without -Provision - their store is cleaned,
                    no Store download (that needs an Entra admin) - or, when one is not
                    provisioned, -Provision -FromHost: provisioned again from the
                    Microsoft-signed build already on the host; Copilot, when the
                    unified app is not on the host, -Name copilot -Provision (Edge
                    Update installs it). No -RemoveOld: nothing is taken away from a
                    user who still has an older build open.
      4. Test       The watchdog task itself is started - as System, with its own
                    config.json, so it is exactly the run that happens every 30 minutes:
                    every signed-in user is checked, repaired in their session where
                    needed, and an app that someone tried to open in the last 30
                    minutes and Windows refused is opened for them once it works -
                    nothing else is started. Its exit code and what it still reports
                    are shown.
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

.PARAMETER KeepStoreUpdates
    Leave the Microsoft Store's automatic app updates as they are (step 2b skipped).

.PARAMETER NoUserRepair
    Do not repair Photos and Snipping Tool in anyone's session (step 3b skipped).

.PARAMETER NoReport
    Do not send the run's report to n8n (the diagnose webhook next to the watchdog's).

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
    [switch] $KeepStoreUpdates,
    [switch] $NoUserRepair,
    [switch] $NoReport,
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
    # The unified app, installed machine-wide by Edge Update: no package of its own.
    [PSCustomObject]@{ App = 'Copilot';      Package = '' }
)
$CopilotGuid = '{C50565E9-CCCF-44B4-BA15-5AC5C6569197}'
$CopilotExe  = 'copilotapp.exe'
$TaskName         = 'M365 App Watchdog'
$FslogixProfiles  = 'HKLM:\SOFTWARE\FSLogix\Profiles'
$FslogixPolicy    = 'HKLM:\SOFTWARE\Policies\FSLogix\Profiles'
# Deployment results that are not a fault, as in the watchdog.
$BenignAppxCodes  = @('0x80073D02', '0x80073CFB', '0x80073D06')
$StorePolicy      = 'HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore'
# Repaired per user in step 3b: no installer, so the Store updates them per user.
$UserApps = @(
    [PSCustomObject]@{ App = 'Photos';       Package = 'Microsoft.Windows.Photos'; Exe = 'Photos' }
    [PSCustomObject]@{ App = 'SnippingTool'; Package = 'Microsoft.ScreenSketch';   Exe = 'SnippingTool' }
)
# What this run did and found, sent to n8n at the end.
$Report = [ordered]@{ source = 'Repair-M365AppsHost'; host = $env:COMPUTERNAME; time = (Get-Date).ToString('o'); actions = [System.Collections.Generic.List[string]]::new() }
function Add-Action { param([string] $Line) $Report.actions.Add($Line) }

function Send-Report {
    <#
        The run's report to n8n: the diagnose webhook next to the watchdog's (its URL with
        app-watchdog replaced by diag), with the watchdog's token, both from config.json.
        n8n keeps it in its log table, so a host where something went wrong shows there.
    #>
    param([int] $ExitCode)
    if ($NoReport) { return }
    try {
        $config = Get-Content -Path (Join-Path $WorkingDir 'config.json') -Raw | ConvertFrom-Json
        $url    = [string] (Get-PropertyValue $config 'WebhookUrl')
        if (-not $url -or $url -notmatch 'app-watchdog') { Write-Skip 'No watchdog webhook in config.json to send the report next to - not sent'; return }
        $Report.exitCode = $ExitCode
        $Report.status   = switch ($ExitCode) { 0 { 'ok' } 2 { 'restart due' } default { 'problem' } }
        $body = [Text.Encoding]::UTF8.GetBytes(($Report | ConvertTo-Json -Depth 6 -Compress))
        Invoke-RestMethod -Uri ($url -replace 'app-watchdog', 'diag') -Method Post -Headers @{ 'X-Watchdog-Token' = [string] (Get-PropertyValue $config 'WebhookToken') } `
            -ContentType 'application/json; charset=utf-8' -Body $body -TimeoutSec 30 | Out-Null
        Write-Ok "Report sent to n8n ($('{0:N0}' -f ($body.Length / 1KB)) KB)"
    } catch {
        Write-Warn "Report not sent to n8n: $($_.Exception.Message)"
    }
}

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
        $entry = @($Apps | Where-Object { $_.Package -and [string] $values[1] -like "$($_.Package)_*" }) | Select-Object -First 1
        if (-not $entry) { continue }
        $seen['{0}|{1}|{2:yyyyMMddHHmmss}' -f $record.UserId, $values[1], $record.TimeCreated] = $entry.App
    }
    $count = @{}
    foreach ($app in $seen.Values) { $count[$app] = 1 + [int] $count[$app] }
    return $count
}

function Get-UnifiedCopilot {
    <#
        The unified Copilot app as the watchdog finds it: the version Edge Update
        reports, copilotapp.exe (App Paths, or Microsoft\<app>\Application under Program
        Files), and a Start menu entry for all users pointing at it. $null when Edge
        Update reports no version.
    #>
    $version = $null
    foreach ($root in 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients', 'HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients') {
        $pv = Get-PropertyValue (Get-ItemProperty (Join-Path $root $CopilotGuid) -ErrorAction SilentlyContinue) 'pv'
        if ($pv -and $pv -ne '0.0.0.0') { $version = [string] $pv; break }
    }
    if (-not $version) { return $null }
    $candidates = @()
    $appPath = Get-PropertyValue (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$CopilotExe" -ErrorAction SilentlyContinue) '(default)'
    if ($appPath) { $candidates += ([string] $appPath).Trim('"') }
    foreach ($programs in @(${env:ProgramFiles(x86)}, $env:ProgramFiles) | Where-Object { $_ }) {
        $candidates += @(Get-ChildItem (Join-Path $programs 'Microsoft') -Directory -ErrorAction SilentlyContinue | ForEach-Object { Join-Path $_.FullName "Application\$CopilotExe" })
    }
    $path = @($candidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) }) | Select-Object -First 1
    $shortcut = $null
    try {
        $shell = New-Object -ComObject WScript.Shell
        foreach ($lnk in @(Get-ChildItem -Path (Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs') -Filter '*.lnk' -Recurse -ErrorAction SilentlyContinue)) {
            if ($shell.CreateShortcut($lnk.FullName).TargetPath -like "*\$CopilotExe") { $shortcut = $lnk.FullName; break }
        }
    } catch { }
    [PSCustomObject]@{ Version = $version; Path = $path; Shortcut = $shortcut }
}

function Get-HostSnapshot {
    <#
        Per app: the provisioned build, and how many signed-in users have it installed.
        Copilot: the unified app on the host - for everyone, so no per-user count; Note
        says what is missing.
    #>
    param([hashtable] $Users)
    $provisioned = @(Get-AppxProvisionedPackage -Online)
    $failures    = Get-RegisterFailure -Since (Get-Date).AddHours(-24)
    foreach ($entry in $Apps) {
        if ($entry.App -eq 'Copilot') {
            $copilot = Get-UnifiedCopilot
            $note    = if (-not $copilot) { 'Edge Update reports no unified Copilot app' }
                       elseif (-not $copilot.Path) { "$CopilotExe is nowhere on the host" }
                       elseif (-not $copilot.Shortcut) { 'no Start menu entry for all users - users cannot find it' }
                       else { '' }
            [PSCustomObject]@{
                App         = 'Copilot'
                Provisioned = if ($copilot -and $copilot.Path) { "unified app $($copilot.Version)" } else { 'NOT installed' }
                Users       = 'machine-wide'
                Missing     = ''
                Failed24h   = 0
                Note        = $note
            }
            continue
        }
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
            Note        = if ($build) { '' } else { 'new profiles do not get it' }
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
    $Snapshot | Format-Table App, Provisioned, Users, Failed24h, Missing, Note -AutoSize -Wrap | Out-String -Width 220 | Write-Host
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

# -- Photos and Snipping Tool per user ---------------------------------------------------
function Get-VersionTail {
    <#
        The last three parts of a version. A bundle and its main package differ in the
        first (Snipping Tool: bundle 2022.2608.41.0, main package 11.2608.41.0), so
        builds are compared on these.
    #>
    param($Version)
    $v = [version] [string] $Version
    return [version] ('{0}.{1}.{2}' -f $v.Minor, [Math]::Max(0, $v.Build), [Math]::Max(0, $v.Revision))
}

function Get-SessionId {
    <# The session of this user's explorer.exe, $null when they have none. #>
    param([string] $Sid)
    foreach ($process in @(Get-CimInstance Win32_Process -Filter "Name = 'explorer.exe'")) {
        if ((Invoke-CimMethod -InputObject $process -MethodName GetOwnerSid -ErrorAction SilentlyContinue).Sid -eq $Sid) { return [int] $process.SessionId }
    }
    return $null
}

function Start-UserTask {
    <#
        A program run as this user in their own session: a one-off task with their SID
        and an interactive token - no password, only while they are signed in - as the
        watchdog does it. Waits up to $WaitSeconds; the task is removed either way.
    #>
    param([string] $Sid, [string] $Execute, [string] $Argument, [int] $WaitSeconds = 180)
    $name = 'Repair-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
    $path = '\M365AppsRepair\'
    try {
        $principal = New-ScheduledTaskPrincipal -UserId $Sid -LogonType Interactive -RunLevel Limited
        $action    = New-ScheduledTaskAction -Execute $Execute -Argument $Argument
        $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
        Register-ScheduledTask -TaskName $name -TaskPath $path -Action $action -Principal $principal -Settings $settings -Force | Out-Null
        $started = Get-Date
        Start-ScheduledTask -TaskName $name -TaskPath $path
        do {
            Start-Sleep -Seconds 2
            $info = Get-ScheduledTaskInfo -TaskName $name -TaskPath $path
            $done = (Get-ScheduledTask -TaskName $name -TaskPath $path).State -ne 'Running' -and $info.LastRunTime -ge $started.AddSeconds(-1)
        } until ($done -or (Get-Date) -gt $started.AddSeconds($WaitSeconds))
        return $done
    } catch {
        Write-Warn "No task in that session: $($_.Exception.Message)"
        return $false
    } finally {
        Unregister-ScheduledTask -TaskName $name -TaskPath $path -Confirm:$false -ErrorAction SilentlyContinue
    }
}

function Invoke-AsSessionUser {
    <#
        PowerShell run as this user in their session, without a window (conhost
        --headless), passed encoded so no quoting can break it. A task passes no output
        on, so what it writes goes to a file in their own profile and comes back here.
    #>
    param([string] $Sid, [string] $Script)
    $profilePath = [string] (Get-CimInstance Win32_UserProfile -Filter "SID = '$Sid'").LocalPath
    if (-not $profilePath) { return @('no profile folder for this user') }
    $out = Join-Path $profilePath ('AppData\Local\Temp\M365AppsRepair_{0}.txt' -f [guid]::NewGuid().ToString('N').Substring(0, 8))
    $wrapped = "& { `$ErrorActionPreference = 'Continue'; $Script } *>&1 | ForEach-Object { [string] `$_ } | Out-File -LiteralPath '$out' -Encoding UTF8"
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($wrapped))
    $done = Start-UserTask -Sid $Sid -Execute (Join-Path $env:WINDIR 'System32\conhost.exe') -Argument ('--headless "{0}" -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand {1}' -f $nativeShell, $encoded)
    $lines = @(if (Test-Path -LiteralPath $out) { Get-Content -LiteralPath $out -Encoding UTF8 })
    Remove-Item -LiteralPath $out -Force -ErrorAction SilentlyContinue
    if (-not $done) { $lines += 'did not finish within 3 minutes' }
    return $lines
}

function Get-MainBuild {
    <#
        Every main package of this app on disk, newest first, and which of them is the
        provisioned build (matched on the version tail: the provisioned entry can be the
        bundle).
    #>
    param($Entry)
    $provisioned = @(Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -eq $Entry.Package } | Sort-Object { [version] $_.Version } -Descending) | Select-Object -First 1
    $tail = if ($provisioned) { Get-VersionTail $provisioned.Version } else { $null }
    @(Get-AppxPackage -AllUsers -Name $Entry.Package -PackageTypeFilter Main -ErrorAction SilentlyContinue |
      Where-Object { $_.InstallLocation -and (Test-Path -LiteralPath (Join-Path $_.InstallLocation 'AppxManifest.xml')) } |
      Sort-Object { [version] $_.Version } -Descending | ForEach-Object {
        [PSCustomObject]@{ Package = $_; Manifest = Join-Path $_.InstallLocation 'AppxManifest.xml'; Provisioned = [bool] ($tail -and (Get-VersionTail $_.Version) -eq $tail) }
    })
}

function Get-UserAppState {
    <#
        Per signed-in user, for one app: the build registered for them, and why it needs
        a repair ('' when it does not): not registered, its files gone, a status other
        than Ok, a Store build newer than the host provisions whose bundle is no longer
        on the host (registering it again fails with 0x80070490 - itceadmin's Photos on
        LEM-AVD-6), or a registration of it that failed for them since $Since.
    #>
    param([hashtable] $Users, $Entry, [datetime] $Since)
    $builds  = @(Get-MainBuild $Entry)
    $provTail = @($builds | Where-Object { $_.Provisioned } | ForEach-Object { Get-VersionTail $_.Package.Version }) | Select-Object -First 1
    $bundles = @(Get-AppxPackage -AllUsers -Name $Entry.Package -PackageTypeFilter Bundle -ErrorAction SilentlyContinue |
                 Where-Object { $_.InstallLocation -and (Test-Path -LiteralPath $_.InstallLocation) } | ForEach-Object { Get-VersionTail $_.Version })
    $failed  = @{}
    foreach ($record in @(Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-AppXDeploymentServer/Operational'; Id = 401, 404; StartTime = $Since } -ErrorAction SilentlyContinue)) {
        $values = @($record.Properties | ForEach-Object { $_.Value })
        $codeAt = if ($record.Id -eq 401) { 3 } else { 2 }
        if ($values.Count -le $codeAt -or [string] $values[1] -notlike "$($Entry.Package)_*") { continue }
        if (('0x{0:X8}' -f ([long] $values[$codeAt] -band [long] 4294967295)) -in $BenignAppxCodes) { continue }
        $failed[[string] $record.UserId] = $true
    }
    $all = @(Get-AppxPackage -AllUsers -Name $Entry.Package -PackageTypeFilter Main -ErrorAction SilentlyContinue)
    foreach ($sid in @($Users.Keys | Sort-Object { $Users[$_] })) {
        $mine = @($all | Where-Object { @($_.PackageUserInformation | Where-Object { $_.UserSecurityId.Sid -eq $sid -and [string] $_.InstallState -eq 'Installed' }).Count -gt 0 } |
                  Sort-Object { [version] $_.Version } -Descending) | Select-Object -First 1
        $why = ''
        if (-not $mine) { $why = 'not registered' }
        elseif (-not ($mine.InstallLocation -and (Test-Path -LiteralPath $mine.InstallLocation))) { $why = 'its files are gone' }
        elseif ([string] $mine.Status -and [string] $mine.Status -ne 'Ok') { $why = "status $($mine.Status)" }
        elseif ($provTail -and (Get-VersionTail $mine.Version) -gt $provTail -and (Get-VersionTail $mine.Version) -notin $bundles) { $why = "Store build $($mine.Version), newer than the host provisions, its bundle gone from the host" }
        elseif ($provTail -and (Get-VersionTail $mine.Version) -lt $provTail) { $why = "$($mine.Version), older than the build the host provisions" }
        elseif ($failed.ContainsKey($sid)) { $why = 'a registration of it failed for them' }
        [PSCustomObject]@{ Sid = $sid; User = $Users[$sid]; App = $Entry.App; Version = if ($mine) { [string] $mine.Version } else { '' }; Why = $why }
    }
}

function Remove-OlderBuild {
    <#
        Every build of this app older than the one the host provisions, gone: for every
        user who has it (Remove-AppxPackage -AllUsers, else per user) and as an older
        provisioned copy. Each registration of the app made Windows update the
        signed-off users still on an older build too, failed on them with 0x80070490
        ("Failure in the Resolved state handler") and took the whole registration down -
        new Outlook would not start for itceadmin. Main package against main package:
        the repair helper's -RemoveOld compares with the provisioned entry, which for
        Snipping Tool is the bundle (2022.2607 against main packages 11.x), and would
        take every build. A build someone signed in is still on is left: it may be open
        for them, and a later run moves them on. Returns $true when no older build is left.
    #>
    param($Entry, [hashtable] $Users)
    $target = @(Get-MainBuild $Entry | Where-Object { $_.Provisioned }) | Select-Object -First 1
    if (-not $target) { Write-Skip "$($Entry.App): no provisioned build on disk - nothing removed"; return $true }
    $keep  = [version] $target.Package.Version
    $inUse = @()
    foreach ($pkg in @(Get-AppxPackage -AllUsers -Name $Entry.Package -PackageTypeFilter Main -ErrorAction SilentlyContinue | Where-Object { [version] $_.Version -lt $keep })) {
        $online = @($pkg.PackageUserInformation | Where-Object { $Users.ContainsKey([string] $_.UserSecurityId.Sid) -and [string] $_.InstallState -eq 'Installed' } |
                    ForEach-Object { $Users[[string] $_.UserSecurityId.Sid] })
        if ($online.Count -gt 0) {
            $inUse += $pkg.PackageFullName
            Write-Skip "$($pkg.PackageFullName): still in use by $($online -join ', ') - left, a later run moves them on"
            Add-Action "$($Entry.App): $($pkg.PackageFullName) left, still in use by $($online -join ', ')"
            continue
        }
        if (-not $PSCmdlet.ShouldProcess($pkg.PackageFullName, 'Remove-AppxPackage -AllUsers (older build)')) { continue }
        try {
            Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop
            Write-Ok "Removed the older $($pkg.PackageFullName) for every user"
            Add-Action "$($Entry.App): older $($pkg.PackageFullName) removed for every user"
        } catch {
            Write-Warn ("{0} for all users: {1} - per user" -f $pkg.PackageFullName, (($_.Exception.Message -split '\r?\n' | Where-Object { $_.Trim() }) -join ' '))
            foreach ($holder in @($pkg.PackageUserInformation)) {
                $who = if ($holder.UserSecurityId.Username) { [string] $holder.UserSecurityId.Username } else { [string] $holder.UserSecurityId.Sid }
                try {
                    Remove-AppxPackage -Package $pkg.PackageFullName -User ([string] $holder.UserSecurityId.Sid) -ErrorAction Stop
                    Write-Ok "  removed for $who"; Add-Action "$($Entry.App): older $($pkg.PackageFullName) removed for $who"
                } catch { Write-Warn "  still there for ${who}: $($_.Exception.Message)" }
            }
        }
    }
    # Older provisioned copies - the provisioned entry can be the bundle, so on the version tail.
    $keepTail = Get-VersionTail $keep
    foreach ($prov in @(Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -eq $Entry.Package -and (Get-VersionTail $_.Version) -lt $keepTail })) {
        if (-not $PSCmdlet.ShouldProcess($prov.PackageName, 'Remove-AppxProvisionedPackage -Online (older build)')) { continue }
        try {
            Remove-AppxProvisionedPackage -Online -PackageName $prov.PackageName -ErrorAction Stop | Out-Null
            Write-Ok "Deprovisioned the older $($prov.PackageName)"; Add-Action "$($Entry.App): older provisioned $($prov.PackageName) removed"
        } catch { Write-Warn "Could not deprovision $($prov.PackageName): $($_.Exception.Message)" }
    }
    # Read back. Left for someone signed in is expected, not a failure.
    $left = @(Get-AppxPackage -AllUsers -Name $Entry.Package -PackageTypeFilter Main -ErrorAction SilentlyContinue |
              Where-Object { [version] $_.Version -lt $keep -and $_.PackageFullName -notin $inUse })
    if ($left.Count -eq 0) { Write-Ok "$($Entry.App): nothing older than $keep left$(if ($inUse) { ' but what signed-in users are still on' })"; return $true }
    Write-Warn "$($Entry.App): still older on the host: $(@($left | ForEach-Object { $_.PackageFullName }) -join ', ')"
    return $false
}

function Repair-UserApp {
    <#
        One app for one user, in their session, nothing closed: a build newer than the
        provisioned one is removed for them only, then the provisioned main package is
        registered from its own manifest - not by family, which takes the newest build
        Windows lists and fails on its missing bundle. Checked again afterwards; still
        broken, every other build on the host is tried, newest first, each checked.
        Returns $true when the app is registered and Ok for them at the end.
    #>
    param($State, $Entry, [hashtable] $Users, [datetime] $Since)
    $builds = @(Get-MainBuild $Entry)
    $order  = @(@($builds | Where-Object { $_.Provisioned }) + @($builds | Where-Object { -not $_.Provisioned }))
    if ($order.Count -eq 0) { Write-Bad "$($Entry.App): no build with a manifest on this host"; Add-Action "$($State.User) $($Entry.App): no build on the host to register"; return $false }
    $attempt = 0
    foreach ($build in $order) {
        $attempt++
        $remove = if ($State.Version -and (Get-VersionTail $State.Version) -gt (Get-VersionTail $build.Package.Version)) { 'yes' } else { '' }
        $script = @"
`$name = '$($Entry.Package)'
`$mine = @(Get-AppxPackage -Name `$name | Sort-Object { [version] `$_.Version } -Descending) | Select-Object -First 1
'before: ' + `$(if (`$mine) { "`$(`$mine.PackageFullName) `$(`$mine.Status)" } else { 'not registered' })
if ('$remove' -and `$mine) { try { Remove-AppxPackage -Package `$mine.PackageFullName -ErrorAction Stop; "removed `$(`$mine.PackageFullName)" } catch { "remove failed: `$(`$_.Exception.Message)" } }
try { Add-AppxPackage -DisableDevelopmentMode -Register '$($build.Manifest)' -ErrorAction Stop; 'registered $($build.Package.PackageFullName)' } catch { "register failed: `$(`$_.Exception.Message)" }
`$now = @(Get-AppxPackage -Name `$name | Sort-Object { [version] `$_.Version } -Descending) | Select-Object -First 1
'after: ' + `$(if (`$now) { "`$(`$now.PackageFullName) `$(`$now.Status)" } else { 'not registered' })
"@
        Write-Skip "attempt $attempt - $($build.Package.PackageFullName)$(if ($remove) { ', newer build removed for them first' })"
        foreach ($line in @(Invoke-AsSessionUser -Sid $State.Sid -Script $script)) { Write-Host "         $line" }
        Add-Action "$($State.User) $($Entry.App): attempt $attempt with $($build.Package.PackageFullName)"
        # Checked again: registered and Ok, and no new failed registration for them.
        $check = @(Get-UserAppState -Users @{ $State.Sid = $State.User } -Entry $Entry -Since (Get-Date).AddSeconds(-5)) | Select-Object -First 1
        if ($check -and -not $check.Why) { Write-Ok "$($Entry.App) for $($State.User): $($check.Version) registered and Ok"; return $true }
        Write-Warn "$($Entry.App) for $($State.User): still $(if ($check) { $check.Why } else { 'unknown' }) after attempt $attempt"
        $State = if ($check) { $check } else { $State }
    }
    return $false
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
$Report.log   = Join-Path $logDir ('Repair-M365AppsHost_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
try {
    Start-Transcript -Path $Report.log -WhatIf:$false | Out-Null
    $transcribing = $true
} catch {
    Write-Warn "No transcript: $($_.Exception.Message)"
}

$exitCode   = 0
$held       = $false
$restartNow = $false
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

        # -- 2b. Store -----------------------------------------------------------
        # The Store updated Photos and Snipping Tool per user to builds whose bundle the
        # host does not keep; registering them again then fails with 0x80070490.
        Write-Step '2b. Store - no automatic app updates on a shared host'
        $autoDownload = Get-PropertyValue (Get-ItemProperty $StorePolicy -ErrorAction SilentlyContinue) 'AutoDownload'
        if ($KeepStoreUpdates) {
            Write-Skip "-KeepStoreUpdates: AutoDownload stays $(if ($null -eq $autoDownload) { 'not set' } else { $autoDownload })"
        } elseif ($autoDownload -eq 2) {
            Write-Ok 'Automatic Store app updates are already off (AutoDownload = 2)'
        } elseif ($PSCmdlet.ShouldProcess("$env:COMPUTERNAME - $StorePolicy\AutoDownload = 2", 'Set')) {
            $backupDir = Join-Path $WorkingDir 'Repair'
            New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
            if (Test-Path $StorePolicy) {
                $backup = Join-Path $backupDir ('WindowsStore-Policy_{0:yyyyMMdd_HHmmss}.reg' -f (Get-Date))
                & reg.exe export 'HKLM\SOFTWARE\Policies\Microsoft\WindowsStore' $backup /y | Out-Null
                if ($LASTEXITCODE -ne 0) { throw "$StorePolicy not backed up - not changed" }
            } else {
                New-Item -Path $StorePolicy -Force | Out-Null
            }
            New-ItemProperty -Path $StorePolicy -Name 'AutoDownload' -Value 2 -PropertyType DWord -Force | Out-Null
            Write-Ok 'AutoDownload = 2 - users keep the build the host provisions'
            Add-Action 'Store: AutoDownload = 2 (automatic app updates off)'
        }

        # -- 3. Repair -----------------------------------------------------------
        Write-Step '3. Repair - the package store, Teams and Outlook at the newest build, what is not provisioned'
        if ($PSCmdlet.ShouldProcess("$env:COMPUTERNAME - Repair-AppxPackageStore for $(@($Apps.App) -join ', ')", 'Repair')) {
            # Held, so the task cannot start a repair of its own halfway through ours.
            Disable-ScheduledTask -TaskName $TaskName | Out-Null
            $held = $true
            # A trigger may have started it just before: disabling does not stop a run.
            if (-not (Wait-WatchdogIdle)) { throw "the watchdog task is still running after $WaitMinutes minutes - try again later" }
            $repairDir = Join-Path $WorkingDir 'Repair'
            $missingOnHost = @($before | Where-Object { $_.Provisioned -like 'NOT *' } | ForEach-Object { $_.App })
            $latest    = @(@{ Teams = 'teams'; Outlook = 'outlook' }.GetEnumerator() | Where-Object { $_.Key -in $Apps.App } | ForEach-Object { $_.Value })
            # No installer: provisioned from the newest Microsoft-signed build on the host
            # when it is not provisioned or a newer build than the provisioned one is
            # there (a Store update of a user's) - so everyone ends on the newest - and
            # otherwise only their store cleaned.
            $newest    = @(foreach ($e in @($UserApps | Where-Object { $_.App -in $Apps.App })) {
                             $b = @(Get-MainBuild $e)
                             if ($e.App -in $missingOnHost -or ($b.Count -gt 0 -and -not $b[0].Provisioned)) { $e.App }
                         })
            $fromHost  = @(@{ Photos = 'photos'; SnippingTool = 'snippingtool' }.GetEnumerator() | Where-Object { $_.Key -in $Apps.App -and $_.Key -in $newest } | ForEach-Object { $_.Value })
            $storeOnly = @(@{ Photos = 'photos'; SnippingTool = 'snippingtool' }.GetEnumerator() | Where-Object { $_.Key -in $Apps.App -and $_.Key -notin $newest } | ForEach-Object { $_.Value })
            foreach ($call in @(
                if ($latest.Count) { [ordered]@{ Name = $latest; Latest = $true; Provision = $true; Days = 1; WorkingDir = $repairDir; LogPath = $repairDir; Confirm = $false } }
                if ($fromHost.Count) { [ordered]@{ Name = $fromHost; Provision = $true; FromHost = $true; Days = 1; WorkingDir = $repairDir; LogPath = $repairDir; Confirm = $false } }
                if ($storeOnly.Count) { [ordered]@{ Name = $storeOnly; Days = 1; WorkingDir = $repairDir; LogPath = $repairDir; Confirm = $false } }
                # Only when the unified app is not there: -Provision writes Edge Update's install policy.
                if ('Copilot' -in $Apps.App -and 'Copilot' -in $missingOnHost) { [ordered]@{ Name = @('copilot'); Provision = $true; Days = 1; WorkingDir = $repairDir; LogPath = $repairDir; Confirm = $false } }
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

        # -- 3b. Users -----------------------------------------------------------
        Write-Step '3b. Users - Photos and Snipping Tool in each session: check, repair, check again'
        $watchedAccounts = @(if (Test-Path $configPath) { Get-PropertyValue (Get-Content -Path $configPath -Raw | ConvertFrom-Json) 'Account' })
        $isOurs = { param($user) @($watchedAccounts | Where-Object { $_ -and ($user -ieq $_ -or $user -like "*\$_") }).Count -gt 0 }
        $Report.users = [System.Collections.Generic.List[object]]::new()
        if ($NoUserRepair) {
            Write-Skip "-NoUserRepair: no session touched"
        } else {
            $usersNow = Get-SignedInUser
            foreach ($entry in @($UserApps | Where-Object { $_.App -in $Apps.App })) {
                $states = @(Get-UserAppState -Users $usersNow -Entry $entry -Since (Get-Date).AddHours(-24))
                # Our own account first: a repair that does not work shows on it, not on a customer.
                $states = @(@($states | Where-Object { & $isOurs $_.User }) + @($states | Where-Object { -not (& $isOurs $_.User) }))
                $oursFailed = $false
                foreach ($state in $states) {
                    $ours = & $isOurs $state.User
                    $row  = [ordered]@{ user = $state.User; app = $entry.App; before = $state.Version; why = $state.Why; result = '' }
                    # Registered and Ok: left alone, not started - closed or not in use is their choice.
                    if (-not $state.Why) {
                        Write-Ok "$($entry.App) for $($state.User): $($state.Version)"
                        $row.result = 'ok'; $Report.users.Add($row); continue
                    }
                    Write-Warn "$($entry.App) for $($state.User): $($state.Why)"
                    if (-not $ours -and $oursFailed) {
                        Write-Skip 'not touched - the same repair did not work for our own account'
                        $row.result = 'not touched, the repair failed for our own account'; $Report.users.Add($row); $exitCode = 1; continue
                    }
                    $session = Get-SessionId $state.Sid
                    # Open for them - our own account included, someone may work in it.
                    if (@(Get-Process -Name $entry.Exe -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $session }).Count -gt 0) {
                        Write-Skip 'open for them right now - not touched, the watchdog tries again later'
                        $row.result = 'open for them, not touched'; $Report.users.Add($row); continue
                    }
                    if (-not $PSCmdlet.ShouldProcess("$($state.User) - $($entry.App)", 'Repair in their session')) { continue }
                    # Not started afterwards: the watchdog run in step 4 opens it for whoever
                    # tried to open it in the last 30 minutes, and nothing for anyone else.
                    $fixed = Repair-UserApp -State $state -Entry $entry -Users $usersNow -Since (Get-Date).AddHours(-24)
                    if (-not $fixed) { $exitCode = 1; if ($ours) { $oursFailed = $true } }
                    $now = @(Get-UserAppState -Users @{ $state.Sid = $state.User } -Entry $entry -Since (Get-Date).AddSeconds(-5)) | Select-Object -First 1
                    $row.after  = if ($now) { $now.Version } else { '' }
                    $row.result = if ($fixed) { 'repaired' } else { 'still broken' }
                    Add-Action "$($state.User) $($entry.App): $($row.result) ($($row.before) -> $($row.after))"
                    $Report.users.Add($row)
                }
            }
        }

        # -- 3c. Older builds -----------------------------------------------------
        # After 3b moved the signed-in users on: what is left of older builds belongs to
        # signed-off users, and makes every registration of the app fail for everyone.
        Write-Step '3c. Older builds - every app on one build, the newest the host provisions'
        $Report.olderBuilds = [ordered]@{}
        if ($NoUserRepair) {
            Write-Skip '-NoUserRepair: older builds left, users were not moved on'
        } else {
            $usersNow = Get-SignedInUser
            foreach ($entry in @($Apps | Where-Object { $_.Package })) {
                $clean = Remove-OlderBuild -Entry $entry -Users $usersNow
                $Report.olderBuilds[$entry.App] = if ($clean) { 'none left' } else { 'some left' }
                if (-not $clean) { $exitCode = 1 }
            }
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
        if ($row.Provisioned -like 'NOT *') { Write-Bad "$($row.App) - $($row.Provisioned.ToLower()): $($row.Note)"; $exitCode = 1 }
        elseif ($row.Note) { Write-Bad "$($row.App) - $($row.Note)"; $exitCode = 1 }
        if ($row.Missing) { Write-Bad "$($row.App) not installed for: $($row.Missing)"; $exitCode = 1 }
        if (-not $row.Note -and -not $row.Missing -and $row.Provisioned -notlike 'NOT *') { Write-Ok "$($row.App) - $($row.Provisioned), $($row.Users)" }
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
            Write-Ok 'Restarting in 30 seconds, once the report is sent'
            $restartNow = $true
        } else {
            if ($exitCode -eq 0) { $exitCode = 2 }
        }
    }
    $Report.before   = $before
    $Report.after    = $after
    $Report.restart  = @($reasons)
    $Report.restarted = [bool] $restartNow
    $Report.watchdog = [ordered]@{ exitCode = $test.Code; lines = @($test.Lines | Where-Object { $_ -match '\[(FAIL|WARN)\]|problem\(s\) left|Everything (healthy|repaired)' } | ForEach-Object { $_.Trim() } | Select-Object -First 40) }
} catch {
    Write-Bad ('Aborted at line {0}: {1}' -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message)
    $Report.error = 'line {0}: {1}' -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message
    $exitCode = 1
} finally {
    # Never leave the watchdog switched off.
    if ($held) { try { Enable-ScheduledTask -TaskName $TaskName | Out-Null } catch { Write-Bad "Task '$TaskName' could not be enabled again - enable it by hand" } }
    if (-not $CheckOnly) { Send-Report -ExitCode $exitCode }
    if ($transcribing) { try { Stop-Transcript | Out-Null } catch { } }
}
if ($restartNow) { & shutdown.exe /r /t 30 /c 'Repair-M365AppsHost: package store edited, restart so the deployment engine rereads it' /d p:4:1 }
exit $exitCode
