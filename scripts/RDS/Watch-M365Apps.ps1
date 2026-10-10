#Requires -Version 5.1
<#
.SYNOPSIS
    Watchdog for new Teams, new Outlook and Copilot on a session host: tests them with
    our own account (itceadmin) and for every signed-in user, repairs what is broken -
    on the host and in that user's own session - and reports to an n8n webhook.

.DESCRIPTION
    Runs as a scheduled task under System (install it with -Install). Each run:

      0. Crashes     Every crash (Application Error 1000) and every hang that ended in
                     a close (Application Hang 1002) of Teams, new Outlook or Copilot
                     since the last run, from the Application log - for every user on
                     the host, customers included. Grouped per app, module and
                     exception code. Reported, not repaired: a customer's app is never
                     restarted for them. Our own account's app is started again in
                     step 2 when it is no longer running.
      1. Host        Teams and new Outlook are provisioned for all users, and Copilot
                     is there - provisioned (MicrosoftOfficeHub / Copilot) or the
                     unified app Edge Update installs. Without that no new profile,
                     customer or ours, gets the app.
      1a. Update     Every -UpdateHours: the newest Teams and new Outlook are
                     provisioned on the host (Repair-AppxPackageStore.ps1 -Latest
                     -Provision: Teams from Microsoft's config service, Outlook the
                     newest build seen on this host or in a profile - only ever newer,
                     signature-checked), and Edge Update is asked to check now for the
                     unified Copilot app. Every run: when our own account has an older
                     build than the host provisions, the app is closed in our session
                     and registered again from the provisioned build, so step 2 starts
                     the new build - a build that does not run is found by us, not by a
                     customer at their next sign-in. Customers are never updated by the
                     watchdog: Windows gives them the new build at their next sign-in.
                     A new build on the host - from this step or on its own - is
                     reported once, with what it was before.
      2. Accounts    For each watched account signed in on this host (the owner of an
                     explorer.exe), per app: the package is registered for that user,
                     its files are there and its status is Ok. Then the app has to be
                     running in that session; when it is not, it is started there
                     (shell:AppsFolder\<AUMID>, through a one-off task in the user's
                     own session) and must still be running 15 seconds later.
      2b. Users      Every other signed-in user (customers), once signed in for 10
                     minutes: the same registration check, without starting anything.
                     And every attempt since the last run, by any user, to open one of
                     the apps that Windows refused (TWinUI 5961), and every failed
                     registration of their packages (AppXDeploymentServer 401/404,
                     minus "close the app first" and "already installed"), with the
                     user it happened to.
      2c. Announce   Something new is wrong: Get-M365AppsLog.ps1 collects the evidence for
                     the users and apps concerned into Diag\ (a zip, kept $LogDays days),
                     and a "repairing" report goes to n8n naming every user, before
                     anything is changed. The same problem again is not announced or
                     collected again within -RenotifyHours.
      3. Repair      A problem of one customer only in their session (below). The host
                     only when it is more than that one user - the host itself, our own
                     account (the test for the whole host), or the same app at two or
                     more customers: Repair-AppxPackageStore.ps1 -Provision (Microsoft's
                     installers, signature-checked), at most once per
                     -RepairCooldownHours. Then per user, in that user's own session
                     (a one-off task running a headless console, so nothing appears):
                     a package that is not registered, is broken or would not open is
                     registered again by family name - for a customer only while the
                     app is not running for them, at most once per -RepairCooldownHours
                     per user and app, and never reset. In our own account an app that
                     does not start is reset (Reset-AppxPackage). Nothing is closed or
                     removed for a customer - no -RemoveOld, no -Latest. A user who tried to
                     open the app themselves (TWinUI 5961) in the last 30 minutes gets it
                     opened in their session once it is registered again, unless
                     -NoUserLaunch. Nothing else is ever started for a customer; an
                     attempt in the first 2 minutes after sign-in is autostart and does
                     not count.
      4. Read back   Steps 1, 2 and 2b's registration check again; a user's failure
                     counts as repaired when the package is registered and Ok for them
                     afterwards.
      5. Report      A JSON POST to -WebhookUrl when something is wrong, was repaired,
                     recovered on its own, or crashed at least -CrashThreshold times -
                     not on every healthy run. A problem that stays is reported again
                     after -RenotifyHours.

    An account that is not signed in on this host is skipped: its packages live in its
    FSLogix container and cannot be tested without it. The watched accounts are
    canaries - keep a session of each open (disconnected is fine) on every host. Every
    report names the user each finding and each repair belongs to.

    -Install copies this script, Repair-AppxPackageStore.ps1 and Get-M365AppsLog.ps1 (from
    the repo checkout, or from GitHub at the commit and SHA-256 pinned below) to
    -WorkingDir, locks that folder to
    System and Administrators, writes the settings to config.json there, and registers
    the task "M365 App Watchdog". The webhook URL and token live only in that file.

.PARAMETER Account
    Accounts to test with, by user name, UPN or DOMAIN\user. Default: itceadmin.
    More than one tests each of them, e.g. -Account itceadmin,itce.user.

.PARAMETER App
    Apps to watch: Teams, Outlook, Copilot. Default: all three.

.PARAMETER WebhookUrl
    n8n webhook (production URL) the report is POSTed to. Without it the run only logs.

.PARAMETER WebhookToken
    Sent as header X-Watchdog-Token. Match it with Header Auth on the n8n Webhook node.

.PARAMETER IntervalMinutes
    How often the task runs. Default 30.

.PARAMETER RepairCooldownHours
    Minimum time between two host repairs (Repair-AppxPackageStore.ps1), so a problem
    it cannot fix is not retried every run. Default 4. Fixes in our own account are
    not limited.

.PARAMETER RenotifyHours
    A problem that stays the same is reported again after this many hours. Default 12.

.PARAMETER UpdateHours
    How often the newest Teams and Outlook are provisioned on the host and Edge Update
    is asked to check for Copilot. Default 6; 0 turns updating off, for our own account
    as well. -NoRepair also leaves everything at the build it is.

.PARAMETER CrashThreshold
    Report crashes and hangs of one app once it reaches this many since the last run.
    Default 1 (every crash); 0 turns crash reporting off. Raise it on a busy pool where
    the odd crash is noise.

.PARAMETER NoRepair
    Test and report only, change nothing.

.PARAMETER NoUserRepair
    Repair the host and our own account, but never run anything in a customer's
    session; their problems are still reported, with their name.

.PARAMETER SkipLaunchTest
    Do not start an app that is not running; only check its registration.

.PARAMETER NoUserLaunch
    After repairing an app a user could not open, do not open it for them.

.PARAMETER NoDiagnostics
    Do not run Get-M365AppsLog.ps1 before a repair.

.PARAMETER Install
    Copy the watchdog to -WorkingDir and register the scheduled task, with the other
    parameters of this call as its settings.

.PARAMETER Uninstall
    Remove the scheduled task and -WorkingDir.

.PARAMETER TestNotification
    Send one test message to the webhook and stop.

.PARAMETER WorkingDir
    Where the watchdog, its settings, state and logs live. Default C:\IT\AppWatchdog.

.EXAMPLE
    # One run now, in this console: what does it find? Repairs only with -NoRepair off.
    .\Watch-M365Apps.ps1 -NoRepair

.EXAMPLE
    # Install on a session host, reporting to n8n
    .\Watch-M365Apps.ps1 -Install -WebhookUrl 'https://n8n.example.com/webhook/m365-apps' -WebhookToken '<token>' -Confirm:$false

.EXAMPLE
    # Check the webhook end to end with the installed settings
    .\Watch-M365Apps.ps1 -TestNotification

.EXAMPLE
    .\Watch-M365Apps.ps1 -Uninstall -Confirm:$false

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows 11 Enterprise multi-session / Windows Server with FSLogix, run as
              administrator or as System
    Exit    : 0 healthy (or repaired), 1 something still broken or the run failed
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param (
    [string[]] $Account = @('itceadmin'),
    [ValidateSet('Teams', 'Outlook', 'Copilot')]
    [string[]] $App = @('Teams', 'Outlook', 'Copilot'),
    [string]   $WebhookUrl,
    [string]   $WebhookToken,
    [ValidateRange(5, 1440)]
    [int]      $IntervalMinutes = 30,
    [ValidateRange(1, 168)]
    [int]      $RepairCooldownHours = 4,
    [ValidateRange(1, 168)]
    [int]      $RenotifyHours = 12,
    [ValidateRange(0, 168)]
    [int]      $UpdateHours = 6,
    [ValidateRange(0, 1000)]
    [int]      $CrashThreshold = 1,
    [switch]   $NoRepair,
    [switch]   $NoUserRepair,
    [switch]   $SkipLaunchTest,
    [switch]   $NoUserLaunch,
    [switch]   $NoDiagnostics,
    [switch]   $Install,
    [switch]   $Uninstall,
    [switch]   $TestNotification,
    [string]   $WorkingDir = 'C:\IT\AppWatchdog'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -- Constants -------------------------------------------------------------------
# Exe: the process name in a crash event. Packaged apps are also matched on the
# package name the event carries, which covers a Copilot executable not listed here.
$Apps = @(
    [PSCustomObject]@{ App = 'Teams';   Packages = @('MSTeams');                                         Exe = @('ms-teams.exe');    Repair = 'teams' }
    [PSCustomObject]@{ App = 'Outlook'; Packages = @('Microsoft.OutlookForWindows');                     Exe = @('olk.exe');         Repair = 'outlook' }
    # copilotapp.exe is the unified Copilot app (Edge Update, 152.x and up): not a package, so only its name finds it.
    [PSCustomObject]@{ App = 'Copilot'; Packages = @('Microsoft.MicrosoftOfficeHub', 'Microsoft.Copilot'); Exe = @('M365Copilot.exe', 'copilotapp.exe'); Repair = 'copilot' }
)
$PublisherId   = '8wekyb3d8bbwe'   # Microsoft's Store publisher id, the same for all four packages
$CopilotGuid   = '{C50565E9-CCCF-44B4-BA15-5AC5C6569197}'
$TaskName      = 'M365 App Watchdog'
$ProbeTaskPath = '\M365AppWatchdog\'
$HostAccount   = '(host)'
$LaunchSeconds = 60
$UserGraceMinutes = 10   # a fresh sign-in is still registering its apps; leave it alone until then
# An app is only opened for a user after a repair when they tried it themselves, recently:
# their last refused open is at most $UserLaunchMinutes old, and not within $AutoStartSeconds
# of signing in - that is the app's own autostart, not the user.
$UserLaunchMinutes = 30
$AutoStartSeconds  = 120
# Deployment results that are not a fault: an update waiting for the app to close
# (0x80073D02), and asking for what is already there (0x80073CFB, 0x80073D06).
$BenignAppxCodes = @('0x80073D02', '0x80073CFB', '0x80073D06')
$LogDays       = 14
# Set before anything can report: Send-Notification reads it, and -TestNotification
# reports before a run has opened its log.
$script:logFile = $null
# The repair helper, when it is not next to this script: fetched from this repo at a
# pinned commit of main and refused unless the SHA-256 matches - the same pin as
# Update-SessionHostImage.ps1. Move both together.
$Helpers = @{
    # main, 2026-10-08
    'Repair-AppxPackageStore.ps1' = @{ Folder = 'Device'; Commit = '048cf967673f02131ddf4efe8a3a8fe72e55f90c'; Hash = '88A9CEAB9F801CF60BFD25AEB69320A97AECD9F19323BC6E354A60118A04C45C' }
    # devel, 2026-10-10 - collects the evidence before a repair (-NoDiagnostics turns it off)
    'Get-M365AppsLog.ps1'         = @{ Folder = 'RDS';    Commit = '6a94769cf7b81c66251e589cac76818410f1b5b0'; Hash = '38D052E1DA99CC8A37162A030B377450B0524687F5841102CDDDF8C5BA2C70CB' }
}
$DiagDir = Join-Path $WorkingDir 'Diag'

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

# -- Settings ------------------------------------------------------------------------
# The task passes nothing but -WorkingDir: everything else comes from config.json,
# and a parameter given on the command line wins over it.
$ConfigPath = Join-Path $WorkingDir 'config.json'
$StatePath  = Join-Path $WorkingDir 'state.json'
$Installed  = (Test-Path $PSScriptRoot) -and ((Resolve-Path $PSScriptRoot).Path.TrimEnd('\') -eq ([IO.Path]::GetFullPath($WorkingDir)).TrimEnd('\'))
if (-not $Install -and (Test-Path $ConfigPath)) {
    $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
    foreach ($name in 'Account', 'App', 'WebhookUrl', 'WebhookToken', 'RepairCooldownHours', 'RenotifyHours', 'UpdateHours', 'CrashThreshold', 'NoRepair', 'NoUserRepair', 'SkipLaunchTest', 'NoUserLaunch', 'NoDiagnostics') {
        if ($PSBoundParameters.ContainsKey($name)) { continue }
        $value = Get-PropertyValue $config $name
        if ($null -ne $value) { Set-Variable -Name $name -Value $value }
    }
}
$Apps = @($Apps | Where-Object { $_.App -in $App })

# -- Helper --------------------------------------------------------------------------
function Protect-WorkingDir {
    <# System runs what is in -WorkingDir, so only System and Administrators may write to it. #>
    New-Item -ItemType Directory -Path $WorkingDir -Force -WhatIf:$false | Out-Null
    & icacls.exe $WorkingDir /inheritance:r /grant:r '*S-1-5-18:(OI)(CI)F' '*S-1-5-32-544:(OI)(CI)F' | Out-Null
    return $LASTEXITCODE -eq 0
}

function Find-Helper {
    <#
        Next to this script in the installed copy (its locked folder), its own folder
        (..\Device or ..\RDS) in a checkout of the repo, and otherwise GitHub at the
        commit pinned in $Helpers when its SHA-256 matches. Not "next to this script"
        anywhere else: downloaded to a folder such as C:\IT\Setup, whatever sits beside
        it - or in C:\IT\Device - could have been put there by any user, and -Install
        would hand it to System.
    #>
    param([string] $Name)
    if (-not $Helpers.ContainsKey($Name)) { return $null }
    $helper     = $Helpers[$Name]
    $candidates = @()
    if ($Installed) {
        $candidates += Join-Path $PSScriptRoot $Name
    } elseif (Test-Path (Join-Path $PSScriptRoot '..\..\menu.ps1')) {
        $candidates += Join-Path (Join-Path (Split-Path $PSScriptRoot) $helper.Folder) $Name
    }
    foreach ($path in $candidates) {
        if (Test-Path $path) { return $path }
    }

    $url    = 'https://raw.githubusercontent.com/sjkanon/M365-Scripts/{0}/scripts/{1}/{2}' -f $helper.Commit, $helper.Folder, $Name
    $target = Join-Path $WorkingDir $Name
    try {
        New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $url -OutFile $target -UseBasicParsing
        $hash = (Get-FileHash -Path $target -Algorithm SHA256).Hash
        if ($hash -ne $helper.Hash) {
            throw "SHA-256 is $hash, expected $($helper.Hash) - not used"
        }
        Write-Ok "$Name fetched from GitHub at $($helper.Commit.Substring(0, 7)), SHA-256 verified"
        return $target
    } catch {
        Remove-Item $target -Force -ErrorAction SilentlyContinue
        Write-Bad "$Name from $url - $($_.Exception.Message)"
        return $null
    }
}

function Invoke-Helper {
    <# Run a repo script in its own process, so its exit does not end this one. #>
    param([string] $Name, [System.Collections.IDictionary] $Parameters, [switch] $Quiet)
    $path = Find-Helper $Name
    if (-not $path) { Write-Bad "$Name not found locally and not verified from GitHub - skipped"; return 1 }
    $command = ConvertTo-ScriptCommand -Path $path -Parameters $Parameters
    Write-Step $command
    # -Quiet: the helper keeps its own summary, the watchdog log does not need it twice.
    if ($Quiet) { & $nativeShell -NoProfile -ExecutionPolicy Bypass -Command $command | Out-Null }
    else { & $nativeShell -NoProfile -ExecutionPolicy Bypass -Command $command | Out-Host }
    $code = $LASTEXITCODE
    if ($code -eq 0) { Write-Ok "$Name finished" } else { Write-Warn "$Name exited with $code" }
    return $code
}

# -- Readers -------------------------------------------------------------------------
function Get-EdgeUpdateClientVersion {
    param([string] $Guid)
    foreach ($root in 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients', 'HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients') {
        $pv = Get-PropertyValue (Get-ItemProperty (Join-Path $root $Guid) -ErrorAction SilentlyContinue) 'pv'
        if ($pv -and $pv -ne '0.0.0.0') { return $pv }
    }
    return $null
}

function Get-UserSession {
    <#
        Everyone signed in on this host: the owner of each explorer.exe, one row per
        user. Watched marks our own accounts, matched on the user name alone, so AD,
        Entra (AzureAD\) and local accounts all count; their Account is the name from
        -Account, a customer's is DOMAIN\user. Since is when that explorer started.
    #>
    $seen = @{}
    foreach ($proc in @(Get-CimInstance -ClassName Win32_Process -Filter "Name = 'explorer.exe'" -ErrorAction SilentlyContinue)) {
        $owner = Invoke-CimMethod -InputObject $proc -MethodName GetOwner -ErrorAction SilentlyContinue
        if (-not $owner -or $owner.ReturnValue -ne 0) { continue }
        $sid = (Invoke-CimMethod -InputObject $proc -MethodName GetOwnerSid -ErrorAction SilentlyContinue).Sid
        if (-not $sid -or $seen.ContainsKey($sid)) { continue }
        $seen[$sid] = $true
        $user    = '{0}\{1}' -f $owner.Domain, $owner.User
        $watched = @($Account | Where-Object { (($_ -split '\\')[-1] -split '@')[0] -ieq $owner.User }) | Select-Object -First 1
        [PSCustomObject]@{
            Account   = if ($watched) { $watched } else { $user }
            User      = $user
            Sid       = $sid
            SessionId = $proc.SessionId
            Since     = $proc.CreationDate
            Watched   = [bool] $watched
        }
    }
}

function Get-AccountSession {
    <# Where a watched account is signed in on this host. #>
    return @(Get-UserSession | Where-Object { $_.Watched })
}

function Resolve-UserName {
    <# A SID from an event as a name: a signed-in user, else the profile folder, else the SID. #>
    param([string] $Sid, $Sessions)
    $session = @($Sessions | Where-Object { $_.Sid -eq $Sid }) | Select-Object -First 1
    if ($session) { return $session.Account }
    $profilePath = Get-PropertyValue (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$Sid" -ErrorAction SilentlyContinue) 'ProfileImagePath'
    if ($profilePath) { return Split-Path $profilePath -Leaf }
    try { return ([Security.Principal.SecurityIdentifier] $Sid).Translate([Security.Principal.NTAccount]).Value } catch { return $Sid }
}

function Get-AppEntry {
    <#
        The AUMID to start and the process name to look for, from the package manifest:
        the first application that shows in Start. Not simply the first one - Teams
        lists MSTeamsRemoteModuleContainer (AppListEntry="none") before MSTeams.
    #>
    param($Package)
    try {
        $manifest = Get-AppxPackageManifest -Package $Package -ErrorAction Stop
        $entry    = @($manifest.Package.Applications.Application | Where-Object {
                        $visual = @($_.ChildNodes | Where-Object { $_.LocalName -eq 'VisualElements' }) | Select-Object -First 1
                        -not $visual -or $visual.GetAttribute('AppListEntry') -ne 'none'
                    }) | Select-Object -First 1
        if (-not $entry) { return $null }
        $exe      = ([string] $entry.GetAttribute('Executable') -split '[\\/]')[-1]
        if (-not $exe) { return $null }
        return [PSCustomObject]@{
            Aumid = '{0}!{1}' -f $Package.PackageFamilyName, $entry.GetAttribute('Id')
            Exe   = [IO.Path]::GetFileNameWithoutExtension($exe)
        }
    } catch {
        return $null
    }
}

function Get-SessionProcess {
    param($Session, [string] $Exe)
    return @(Get-Process -Name $Exe -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $Session.SessionId })
}

function Invoke-AsUser {
    <#
        Run a command in the user's own session: a one-off scheduled task with that
        user's SID and an interactive token - no password, and it only runs while the
        user is signed in. Waits up to $WaitSeconds for it to finish and returns its
        result code; the task is removed either way.
    #>
    param($Session, [string] $Execute, [string] $Argument, [int] $WaitSeconds = 120)
    $name = 'Probe-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
    try {
        $principal = New-ScheduledTaskPrincipal -UserId $Session.Sid -LogonType Interactive -RunLevel Limited
        $action    = New-ScheduledTaskAction -Execute $Execute -Argument $Argument
        $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
        Register-ScheduledTask -TaskName $name -TaskPath $ProbeTaskPath -Action $action -Principal $principal -Settings $settings -Force | Out-Null
        $started = Get-Date
        Start-ScheduledTask -TaskName $name -TaskPath $ProbeTaskPath
        $deadline = $started.AddSeconds($WaitSeconds)
        do {
            Start-Sleep -Seconds 2
            $task = Get-ScheduledTask -TaskName $name -TaskPath $ProbeTaskPath
            $info = Get-ScheduledTaskInfo -TaskName $name -TaskPath $ProbeTaskPath
            $done = $task.State -ne 'Running' -and $info.LastRunTime -ge $started.AddSeconds(-1)
        } until ($done -or (Get-Date) -gt $deadline)
        if (-not $done) { return -1 }
        return [long] $info.LastTaskResult
    } catch {
        Write-Warn "  Could not run a task in $($Session.User)'s session: $($_.Exception.Message)"
        return -1
    } finally {
        Unregister-ScheduledTask -TaskName $name -TaskPath $ProbeTaskPath -Confirm:$false -ErrorAction SilentlyContinue
    }
}

function Invoke-PowerShellAsUser {
    <#
        A PowerShell command in the user's session without a window: through conhost
        --headless, which on Windows 11 also keeps Windows Terminal from opening. It
        waits for the command, but does not pass its exit code on - check the result
        afterwards instead.
    #>
    param($Session, [string] $Command, [int] $WaitSeconds = 300)
    $argument = '--headless "{0}" -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "{1}"' -f $nativeShell, $Command
    return Invoke-AsUser -Session $Session -Execute (Join-Path $env:WINDIR 'System32\conhost.exe') -Argument $argument -WaitSeconds $WaitSeconds
}

function Test-UserPackage {
    <# Whether one of the app's packages is registered, present and Ok for this user (DOMAIN\user, not a SID - see Test-UserApp). #>
    param([string] $User, [string[]] $Packages)
    foreach ($name in $Packages) {
        foreach ($pkg in @(Get-AppxPackage -User $User -Name $name -ErrorAction SilentlyContinue)) {
            $location = Get-PropertyValue $pkg 'InstallLocation'
            $status   = [string] (Get-PropertyValue $pkg 'Status')
            if ($location -and (Test-Path -LiteralPath $location) -and (-not $status -or $status -eq 'Ok')) { return $true }
        }
    }
    return $false
}

function Test-AppStart {
    <# Start the app in the user's session; it has to come up and still be there 15 seconds later. #>
    param($Session, $Entry)
    Invoke-AsUser -Session $Session -Execute 'explorer.exe' -Argument "shell:AppsFolder\$($Entry.Aumid)" -WaitSeconds 30 | Out-Null
    $deadline = (Get-Date).AddSeconds($LaunchSeconds)
    while ((Get-Date) -lt $deadline -and @(Get-SessionProcess $Session $Entry.Exe).Count -eq 0) { Start-Sleep -Seconds 3 }
    if (@(Get-SessionProcess $Session $Entry.Exe).Count -eq 0) { return $false }
    Start-Sleep -Seconds 15
    return @(Get-SessionProcess $Session $Entry.Exe).Count -gt 0
}

# -- Updates -------------------------------------------------------------------------
function Get-ProvisionedBuild {
    <# Per package name the newest provisioned build: its version and the manifest it was provisioned from. #>
    $builds = @{}
    foreach ($p in @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue)) {
        $v = try { [version] $p.Version } catch { $null }
        if (-not $v) { continue }
        if ($builds.ContainsKey($p.DisplayName) -and $builds[$p.DisplayName].Version -ge $v) { continue }
        $builds[$p.DisplayName] = [PSCustomObject]@{
            Version  = $v
            Manifest = [Environment]::ExpandEnvironmentVariables([string] (Get-PropertyValue $p 'InstallLocation'))
        }
    }
    return $builds
}

function Get-HostVersion {
    <# Per app the build a sign-in gets on this host: the unified Copilot app (Edge Update) first, else the newest provisioned package. #>
    $builds   = Get-ProvisionedBuild
    $versions = [ordered]@{}
    foreach ($entry in $Apps) {
        $v = $null
        if ($entry.App -eq 'Copilot') { $v = Get-EdgeUpdateClientVersion $CopilotGuid }
        if (-not $v) {
            $v = @($entry.Packages | Where-Object { $builds.ContainsKey($_) } | ForEach-Object { $builds[$_].Version } | Sort-Object -Descending) | Select-Object -First 1
        }
        if ($v) { $versions[$entry.App] = [string] $v }
    }
    return $versions
}

function Update-HostApp {
    <#
        The newest builds on the host: Teams and new Outlook through
        Repair-AppxPackageStore.ps1 -Latest -Provision, which only provisions a build
        newer than what is there and checks its signature, and Edge Update asked to
        check now for the unified Copilot app - what it installs shows at the next run.
        Returns one line per action for the report.
    #>
    param($State)
    $actions   = [System.Collections.Generic.List[string]]::new()
    $repairDir = Join-Path $WorkingDir 'Repair'
    $names     = @($Apps | Where-Object { $_.App -ne 'Copilot' } | ForEach-Object { $_.Repair })
    if ($names.Count -gt 0) {
        $code = Invoke-Helper 'Repair-AppxPackageStore.ps1' ([ordered]@{ Name = $names; Latest = $true; Provision = $true; Days = 1; WorkingDir = $repairDir; LogPath = $repairDir; Confirm = $false })
        $actions.Add("Update: Repair-AppxPackageStore -Name $($names -join ',') -Latest -Provision: exit code $code")
    }
    if (@($Apps | Where-Object { $_.App -eq 'Copilot' }).Count -gt 0 -and (Get-EdgeUpdateClientVersion $CopilotGuid)) {
        $task    = @(Get-ScheduledTask -TaskName 'MicrosoftEdgeUpdateTaskMachineUA*' -ErrorAction SilentlyContinue) | Select-Object -First 1
        $updater = Join-Path ${env:ProgramFiles(x86)} 'Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe'
        try {
            if ($task) { Start-ScheduledTask -InputObject $task }
            elseif (Test-Path $updater) { Start-Process -FilePath $updater -ArgumentList '/ua /installsource scheduler' -WindowStyle Hidden | Out-Null }
            else { throw 'Edge Update is not installed' }
            Write-Ok 'Edge Update asked to check now (unified Copilot app)'
            $actions.Add('Update: Edge Update asked to check for Copilot - a new build shows at the next run')
        } catch {
            Write-Warn "Edge Update check not started: $($_.Exception.Message)"
            $actions.Add("Update: Edge Update check for Copilot not started - $($_.Exception.Message)")
        }
    }
    $State.LastUpdate = (Get-Date).ToString('o')
    return $actions
}

function Update-AccountApp {
    <#
        Our own accounts onto the build the host provisions, when theirs is older: the
        app closed in that session - ours, never a customer's - and registered again
        from the provisioned manifest, then the version read back. The launch test in
        Get-Finding then starts the new build. A package that is not registered at all
        is left to Test-UserApp, which reports and repairs it.
    #>
    param($Sessions)
    $actions = [System.Collections.Generic.List[string]]::new()
    $builds  = Get-ProvisionedBuild
    foreach ($session in @($Sessions | Where-Object { $_.Watched })) {
        foreach ($entry in $Apps) {
            foreach ($name in @($entry.Packages | Where-Object { $builds.ContainsKey($_) })) {
                $mine = @(Get-AppxPackage -User $session.User -Name $name -ErrorAction SilentlyContinue) |
                        Sort-Object { [version] $_.Version } -Descending | Select-Object -First 1
                if (-not $mine) { continue }
                $have = [version] $mine.Version
                $want = $builds[$name].Version
                if ($have -ge $want) { continue }

                Write-Step "$($entry.App) for $($session.User): $have, the host provisions $want - updating"
                $exes = @($entry.Exe | ForEach-Object { [IO.Path]::GetFileNameWithoutExtension($_) })
                $appEntry = Get-AppEntry $mine
                if ($appEntry) { $exes += $appEntry.Exe }
                foreach ($exe in @($exes | Sort-Object -Unique)) {
                    Get-SessionProcess $session $exe | Stop-Process -Force -ErrorAction SilentlyContinue
                }
                $manifest = $builds[$name].Manifest
                $command  = if ($manifest -and (Test-Path -LiteralPath $manifest)) {
                                "Add-AppxPackage -Register '$manifest' -DisableDevelopmentMode -ForceApplicationShutdown"
                            } else {
                                "Add-AppxPackage -RegisterByFamilyName -MainPackage '${name}_$PublisherId' -ForceApplicationShutdown"
                            }
                Invoke-PowerShellAsUser -Session $session -Command $command | Out-Null

                $now = @(Get-AppxPackage -User $session.User -Name $name -ErrorAction SilentlyContinue) |
                       ForEach-Object { [version] $_.Version } | Sort-Object -Descending | Select-Object -First 1
                if ($now -and $now -ge $want) {
                    Write-Ok "$($entry.App) $now registered for $($session.Account) - started by the test below"
                    $actions.Add("Update: $($session.Account) $($entry.App) $have -> $now")
                } else {
                    Write-Warn "$($entry.App) for $($session.Account) is still $now after registering $want"
                    $actions.Add("Update: $($session.Account) $($entry.App) still $now after registering $want")
                }
            }
        }
    }
    return $actions
}

# -- Crashes -------------------------------------------------------------------------
function Get-AppCrash {
    <#
        Crashes (Application Error 1000) and hangs that ended in a close (Application
        Hang 1002) of the watched apps since $Since, grouped per app, kind, module and
        exception code. Field positions as Windows writes them: 1000 has the app at 0,
        its version at 1, the module at 3, the exception code at 6 and the package at
        13; 1002 has the package at 7 and the hang type at 9. The events name no user.
    #>
    param([datetime] $Since)
    $events = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; ProviderName = 'Application Error', 'Application Hang'; Id = 1000, 1002; StartTime = $Since } -ErrorAction SilentlyContinue)
    $rows = @(foreach ($record in $events) {
        $values = @($record.Properties | ForEach-Object { [string] $_.Value })
        $field  = { param($i) if ($i -lt $values.Count) { $values[$i] } else { '' } }
        $exe    = & $field 0
        $crash  = $record.Id -eq 1000
        $pkg    = if ($crash) { & $field 13 } else { & $field 7 }
        $entry  = @($Apps | Where-Object {
                      $exe -in $_.Exe -or ($pkg -and @($_.Packages | Where-Object { $pkg -like "$($_)_*" }).Count -gt 0)
                  }) | Select-Object -First 1
        if (-not $entry) { continue }
        [PSCustomObject]@{
            App     = $entry.App
            Kind    = if ($crash) { 'Crash' } else { 'Hang' }
            Time    = $record.TimeCreated
            Exe     = $exe
            Version = & $field 1
            Module  = if ($crash) { & $field 3 } else { 'hang: ' + (& $field 9) }
            Code    = if ($crash) { '0x' + (& $field 6) } else { '' }
        }
    })
    if ($rows.Count -eq 0) { return }
    foreach ($group in @($rows | Group-Object App, Kind, Module, Code)) {
        $newest = $group.Group | Sort-Object Time -Descending | Select-Object -First 1
        [PSCustomObject]@{
            App     = $newest.App
            Kind    = $newest.Kind
            Count   = $group.Count
            Last    = $newest.Time.ToString('o')
            Exe     = $newest.Exe
            Version = $newest.Version
            Module  = $newest.Module
            Code    = $newest.Code
        }
    }
}

# -- Apps users could not open ---------------------------------------------------------
function Format-HResult {
    <# An HRESULT as 0x80070490, whether the event stored it signed or unsigned. #>
    param($Value)
    return '0x{0:X8}' -f ([long] $Value -band [long] 4294967295)
}

function Get-OpenFailure {
    <#
        Since $Since, for any user: Windows refusing to open one of the apps (TWinUI
        5961 - AUMID at 0, HRESULT at 1), and a registration of one of their packages
        that failed (AppXDeploymentServer 401 - package at 1, HRESULT at 3; 404 -
        package at 1, HRESULT at 2), minus $BenignAppxCodes. Fields, not message text:
        the text is translated on a Dutch or French Windows. One failure writes both
        401 and 404, so they are counted once per user, package, code and second. The
        user is the event's own SID; System (FSLogix registering at sign-in) becomes
        the host. One finding per user, app, kind and code.
    #>
    param([datetime] $Since, $Sessions)
    $rows = @(
        foreach ($record in @(Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-TWinUI/Operational'; Id = 5961; StartTime = $Since } -ErrorAction SilentlyContinue)) {
            $values = @($record.Properties | ForEach-Object { $_.Value })
            if ($values.Count -lt 2) { continue }
            [PSCustomObject]@{ Kind = 'WontOpen'; Target = [string] $values[0]; Code = Format-HResult $values[1]; Record = $record }
        }
        foreach ($record in @(Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-AppXDeploymentServer/Operational'; Id = 401, 404; StartTime = $Since } -ErrorAction SilentlyContinue)) {
            $values = @($record.Properties | ForEach-Object { $_.Value })
            $codeAt = if ($record.Id -eq 401) { 3 } else { 2 }
            if ($values.Count -le $codeAt) { continue }
            [PSCustomObject]@{ Kind = 'RegisterFailed'; Target = [string] $values[1]; Code = Format-HResult $values[$codeAt]; Record = $record }
        }
    )
    $seen  = @{}
    $found = @(foreach ($row in $rows) {
        if ($row.Code -in $BenignAppxCodes) { continue }
        $entry = $null; $package = $null
        foreach ($candidate in $Apps) {
            $package = @($candidate.Packages | Where-Object { $row.Target -like "$($_)_*" }) | Select-Object -First 1
            if ($package) { $entry = $candidate; break }
        }
        if (-not $entry) { continue }
        $sid = [string] $row.Record.UserId
        $key = '{0}|{1}|{2}|{3:yyyyMMddHHmmss}' -f $sid, $row.Target, $row.Code, $row.Record.TimeCreated
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        [PSCustomObject]@{ Sid = $sid; App = $entry.App; Package = $package; Kind = $row.Kind; Code = $row.Code; Time = $row.Record.TimeCreated }
    })
    if ($found.Count -eq 0) { return }
    foreach ($group in @($found | Group-Object Sid, App, Kind, Code)) {
        $first   = $group.Group[0]
        $last    = ($group.Group | Sort-Object Time -Descending | Select-Object -First 1).Time
        $system  = -not $first.Sid -or $first.Sid -eq 'S-1-5-18'
        $session = @($Sessions | Where-Object { $_.Sid -eq $first.Sid }) | Select-Object -First 1
        $name    = if ($system) { $HostAccount } else { Resolve-UserName $first.Sid $Sessions }
        $what    = if ($first.Kind -eq 'WontOpen') { 'Windows could not open it' } else { 'its registration failed' }
        $detail  = '{0} {1}x since {2:dd-MM HH:mm}, last at {3:HH:mm}, error {4}' -f $what, $group.Count, $Since, $last, $first.Code
        New-Finding -Account $name -App $first.App -Problem $first.Kind -Detail $detail -Package $first.Package `
            -Sid $(if ($system) { '' } else { $first.Sid }) -Customer:(-not $system -and -not ($session -and $session.Watched)) -FromEvent -LastAttempt $last
    }
}

# -- Checks --------------------------------------------------------------------------
function New-Finding {
    param([string] $Account, [string] $App, [string] $Problem, [string] $Detail, [string] $Package,
          [string] $Sid = '', [switch] $Customer, [switch] $FromEvent, $LastAttempt = $null)
    $who = if ($Account -eq $HostAccount) { '' } else { "$Account - " }
    Write-Bad "$who$App - $Detail"
    [PSCustomObject]@{ Account = $Account; App = $App; Problem = $Problem; Detail = $Detail; Package = $Package
                       Sid = $Sid; Customer = [bool] $Customer; FromEvent = [bool] $FromEvent; Fixed = $false
                       LastAttempt = $LastAttempt }
}

function Test-HostApp {
    param($Entry, [string[]] $Provisioned)
    $have = @($Entry.Packages | Where-Object { $_ -in $Provisioned })
    if ($have.Count -gt 0) { Write-Ok "$($Entry.App) provisioned ($($have -join ', '))"; return }
    if ($Entry.App -eq 'Copilot') {
        $unified = Get-EdgeUpdateClientVersion $CopilotGuid
        if ($unified) { Write-Ok "Copilot: unified app $unified (Edge Update)"; return }
    }
    New-Finding $HostAccount $Entry.App 'NotProvisioned' "not provisioned for new profiles ($($Entry.Packages -join ' / '))" $Entry.Packages[0]
}

function Test-UserApp {
    <# -NoLaunch for customers: registration only, nothing is started in their session, and no OK lines. #>
    param($Session, $Entry, [string[]] $Provisioned, [switch] $NoLaunch)
    $customer = -not $Session.Watched
    $pkg = $null
    foreach ($name in $Entry.Packages) {
        # By name, not SID: Get-AppxPackage answers "No valid SID could be determined"
        # for an Entra ID SID (S-1-12-1-...), in Windows PowerShell 5.1 as well.
        $pkg = @(Get-AppxPackage -User $Session.User -Name $name -ErrorAction SilentlyContinue) |
               Sort-Object { [version] $_.Version } -Descending | Select-Object -First 1
        if ($pkg) { break }
    }
    if (-not $pkg) {
        if ($Entry.App -eq 'Copilot' -and (Get-EdgeUpdateClientVersion $CopilotGuid)) {
            if (-not $NoLaunch) { Write-Ok 'Copilot: unified app (Edge Update, machine-wide) - not started by this test' }
            return
        }
        $target = @($Entry.Packages | Where-Object { $_ -in $Provisioned }) | Select-Object -First 1
        if (-not $target) { $target = $Entry.Packages[0] }
        New-Finding $Session.Account $Entry.App 'NotRegistered' 'not registered for this user' $target -Sid $Session.Sid -Customer:$customer
        return
    }

    $location = Get-PropertyValue $pkg 'InstallLocation'
    if (-not $location -or -not (Test-Path -LiteralPath $location)) {
        New-Finding $Session.Account $Entry.App 'Broken' "$($pkg.PackageFullName): its files are gone" $pkg.Name -Sid $Session.Sid -Customer:$customer
        return
    }
    $status = [string] (Get-PropertyValue $pkg 'Status')
    if ($status -and $status -ne 'Ok') {
        New-Finding $Session.Account $Entry.App 'Broken' "$($pkg.PackageFullName): status is $status" $pkg.Name -Sid $Session.Sid -Customer:$customer
        return
    }
    if ($NoLaunch) { return }

    $appEntry = Get-AppEntry $pkg
    if (-not $appEntry) { Write-Warn "$($Entry.App) $($pkg.Version) registered, but its manifest names no executable - not started"; return }
    if (@(Get-SessionProcess $Session $appEntry.Exe).Count -gt 0) { Write-Ok "$($Entry.App) $($pkg.Version) running"; return }
    if ($SkipLaunchTest) { Write-Ok "$($Entry.App) $($pkg.Version) registered (not running, launch test off)"; return }
    if (Test-AppStart $Session $appEntry) { Write-Ok "$($Entry.App) $($pkg.Version) started"; return }
    New-Finding $Session.Account $Entry.App 'WontStart' "$($pkg.Version) is registered but $($appEntry.Exe).exe did not start or did not stay up" $pkg.Name -Sid $Session.Sid
}

function Get-Finding {
    <# Every check, read only - apart from starting an app that is not running in our own session. #>
    param($Sessions)
    $provisioned = @(Get-AppxProvisionedPackage -Online | ForEach-Object { $_.DisplayName })
    Write-Step 'Host'
    foreach ($entry in $Apps) { Test-HostApp $entry $provisioned }
    foreach ($name in $Account) {
        $session = @($Sessions | Where-Object { $_.Watched -and $_.Account -eq $name }) | Select-Object -First 1
        if (-not $session) { Write-Step $name; Write-Skip 'Not signed in on this host - nothing to test'; continue }
        Write-Step ('{0} (session {1})' -f $session.User, $session.SessionId)
        foreach ($entry in $Apps) { Test-UserApp $session $entry $provisioned }
    }
    $customers = @($Sessions | Where-Object { -not $_.Watched })
    $settled   = @($customers | Where-Object { -not $_.Since -or $_.Since -lt (Get-Date).AddMinutes(-$UserGraceMinutes) })
    Write-Step ('Users: {0} signed in, {1} checked (registration only)' -f $customers.Count, $settled.Count)
    if ($customers.Count -gt $settled.Count) { Write-Skip ('{0} signed in less than {1} minutes ago - next run' -f ($customers.Count - $settled.Count), $UserGraceMinutes) }
    foreach ($session in $settled) {
        foreach ($entry in $Apps) { Test-UserApp $session $entry $provisioned -NoLaunch }
    }
}

# -- Repair --------------------------------------------------------------------------
function Invoke-Repair {
    <#
        Host first - a user's fix needs the package on the machine - then each user
        in their own session. Returns one line per action for the report, naming the
        user, and marks a finding Fixed when the read-back says it is.
    #>
    param($Found, $State)
    $actions = [System.Collections.Generic.List[string]]::new()

    # A problem of one customer is fixed in their session only. The host is repaired for
    # an app when it is not just that one user: the host itself, our own account (the
    # test for everyone on the host), or the same app at two or more customers. The host
    # repair provisions and starts nothing in anyone's session.
    $customersPerApp = @{}
    foreach ($group in @($Found | Where-Object { $_.Customer } | Group-Object App)) {
        $customersPerApp[$group.Name] = @($group.Group | ForEach-Object { $_.Account } | Sort-Object -Unique).Count
    }
    $hostApps = @($Found | Where-Object { -not $_.Customer -or $customersPerApp[$_.App] -ge 2 } | ForEach-Object { $_.App } | Sort-Object -Unique)
    if ($hostApps.Count -gt 0) {
        $last = $null
        if ($State.LastRepair) { $last = [datetime]::Parse($State.LastRepair, $null, [Globalization.DateTimeStyles]::RoundtripKind) }
        if ($last -and $last.AddHours($RepairCooldownHours) -gt (Get-Date)) {
            $actions.Add(('Host repair skipped: the last one ran at {0:yyyy-MM-dd HH:mm}, cooldown {1} h' -f $last, $RepairCooldownHours))
        } else {
            $repairDir = Join-Path $WorkingDir 'Repair'
            $names = @($Apps | Where-Object { $_.App -in $hostApps -and $_.App -ne 'Copilot' } | ForEach-Object { $_.Repair })
            if ($names.Count -gt 0) {
                $code = Invoke-Helper 'Repair-AppxPackageStore.ps1' ([ordered]@{ Name = $names; Provision = $true; Days = 1; WorkingDir = $repairDir; LogPath = $repairDir; Confirm = $false })
                $actions.Add("Repair-AppxPackageStore -Name $($names -join ',') -Provision: exit code $code")
            }
            if ('Copilot' -in $hostApps) {
                $code = Invoke-Helper 'Repair-AppxPackageStore.ps1' ([ordered]@{ Name = @('copilot'); Provision = $true; Days = 1; WorkingDir = $repairDir; LogPath = $repairDir; Confirm = $false })
                $actions.Add("Repair-AppxPackageStore -Name copilot -Provision: exit code $code")
            }
            $State.LastRepair = (Get-Date).ToString('o')
        }
    }

    $sessions = @(Get-UserSession)
    $done     = @{}
    foreach ($finding in @($Found | Where-Object { $_.Sid })) {
        $who   = $finding.Account
        $key   = '{0}|{1}' -f $finding.Sid, $finding.App
        $entry = @($Apps | Where-Object { $_.App -eq $finding.App }) | Select-Object -First 1
        if ($done.ContainsKey($key)) { $finding.Fixed = $done[$key]; continue }

        $session = @($sessions | Where-Object { $_.Sid -eq $finding.Sid }) | Select-Object -First 1
        if (-not $session) {
            $actions.Add("$who $($finding.App): signed off - the host repair covers the next sign-in")
            continue
        }
        if ($finding.Customer -and $NoUserRepair) {
            $actions.Add("$who $($finding.App): not repaired in their session (-NoUserRepair)")
            continue
        }
        if ($finding.Customer) {
            $running = @($entry.Exe | ForEach-Object { Get-SessionProcess $session ([IO.Path]::GetFileNameWithoutExtension($_)) })
            if ($running.Count -gt 0) {
                # It is open for them now, so whatever failed earlier has passed - and
                # re-registering under a running app would close it.
                $finding.Fixed = $true
                $done[$key]    = $true
                $actions.Add("$who $($finding.App): running for them now - nothing to do")
                continue
            }
            $lastFix = $State.UserRepairs[$key]
            if ($lastFix -and ([datetime]::Parse($lastFix, $null, [Globalization.DateTimeStyles]::RoundtripKind)).AddHours($RepairCooldownHours) -gt (Get-Date)) {
                $actions.Add("$who $($finding.App): already re-registered at $(([datetime] $lastFix).ToString('HH:mm')) - not again within $RepairCooldownHours h")
                continue
            }
        }

        $family = '{0}_{1}' -f $finding.Package, $PublisherId
        if ($finding.Problem -eq 'WontStart' -and -not $finding.Customer) {
            $verb    = 'reset'
            $command = "Get-AppxPackage -Name '$($finding.Package)' | Reset-AppxPackage"
        } else {
            $verb    = 're-registered'
            $command = "Add-AppxPackage -RegisterByFamilyName -MainPackage '$family'"
        }
        Write-Step "$($finding.App) for $($session.User): $verb in their session"
        Invoke-PowerShellAsUser -Session $session -Command $command | Out-Null
        if ($finding.Customer) { $State.UserRepairs[$key] = (Get-Date).ToString('o') }

        # The headless console does not pass the exit code on: the package itself says
        # whether it worked. A reset is judged by the launch test in the read-back.
        $ok = ($verb -eq 'reset') -or (Test-UserPackage -User $session.User -Packages $entry.Packages)
        $finding.Fixed = $ok
        $done[$key]    = $ok
        if ($ok) { Write-Ok "$($finding.App) $verb for $who" } else { Write-Warn "$($finding.App) for $who - still not registered and Ok after $verb" }
        $actions.Add(('{0} {1}: {2} in their session - {3}' -f $who, $finding.App, $verb, $(if ($ok) { 'OK' } else { 'still not registered and Ok' })))

        # They tried to open it and Windows refused: now that it is registered again,
        # open it for them, so they do not have to try once more - or call.
        if ($ok -and $finding.Problem -eq 'WontOpen' -and -not $NoUserLaunch) {
            $notNow = Test-UserAttempt -Finding $finding -Session $session
            if ($notNow) {
                Write-Skip "$($finding.App) for $who not opened - $notNow"
                $actions.Add("$who $($finding.App): not opened for them - $notNow")
            } else {
                $result = Start-AppForUser -Session $session -Entry $entry -Who $who
                $actions.Add($result.Action)
                if (-not $result.Ok) { $finding.Fixed = $false; $done[$key] = $false }
            }
        }
    }
    return $actions
}

function Test-UserAttempt {
    <#
        Why an app should not be opened for this user now, or '' when it may: only
        right after they tried it themselves. An old attempt (they gave up, or are
        doing something else) or one in the first minutes after sign-in (autostart
        of Teams or Outlook, which Windows logs the same way) does not count.
    #>
    param($Finding, $Session)
    $last = $Finding.LastAttempt
    if (-not $last) { return 'no time of their own attempt' }
    $last = [datetime] $last
    if ($last -lt (Get-Date).AddMinutes(-$UserLaunchMinutes)) {
        return ('their last attempt was at {0:HH:mm}, more than {1} minutes ago' -f $last, $UserLaunchMinutes)
    }
    if ($Session.Since -and $last -lt ([datetime] $Session.Since).AddSeconds($AutoStartSeconds)) {
        return 'the attempt came right after sign-in - the app''s autostart, not the user'
    }
    return ''
}

function Start-AppForUser {
    <#
        Open the app in the user's own session, as Test-AppStart does for our account:
        after a repair, for a user whose open Windows refused. Not when it is already
        running for them.
    #>
    param($Session, $Entry, [string] $Who)
    $running = @($Entry.Exe | ForEach-Object { Get-SessionProcess $Session ([IO.Path]::GetFileNameWithoutExtension($_)) })
    if ($running.Count -gt 0) { return @{ Ok = $true; Action = "$Who $($Entry.App): already open for them" } }
    $pkg = $null
    foreach ($name in $Entry.Packages) {
        $pkg = @(Get-AppxPackage -User $Session.User -Name $name -ErrorAction SilentlyContinue) |
               Sort-Object { [version] $_.Version } -Descending | Select-Object -First 1
        if ($pkg) { break }
    }
    $appEntry = if ($pkg) { Get-AppEntry $pkg } else { $null }
    if (-not $appEntry) { return @{ Ok = $true; Action = "$Who $($Entry.App): repaired, but its manifest names nothing to open - not opened for them" } }
    Write-Step "$($Entry.App) for $($Session.User): opening it for them - they tried before"
    if (Test-AppStart $Session $appEntry) {
        Write-Ok "$($Entry.App) open for $Who"
        return @{ Ok = $true; Action = "$Who $($Entry.App): opened for them - running" }
    }
    Write-Warn "$($Entry.App) for $Who - opened, but $($appEntry.Exe).exe did not start or stay up"
    return @{ Ok = $false; Action = "$Who $($Entry.App): opened for them, but it did not start or stay up" }
}

# -- Report --------------------------------------------------------------------------
function Send-Notification {
    param([string] $Kind, [string] $Summary, $Found = @(), $Before = @(), $Actions = @(), $Crashes = @(), [string] $Diagnostics,
          $Updates = @(), $Versions = $null)
    if (-not $WebhookUrl) { Write-Skip "No -WebhookUrl - not reported ($Kind)"; return $false }
    $payload = [ordered]@{
        source   = 'Watch-M365Apps'
        event    = $Kind
        host     = $env:COMPUTERNAME
        time     = (Get-Date).ToString('o')
        summary  = $Summary
        accounts = @($Account)
        findings = @($Found  | Select-Object Account, App, Problem, Detail, Customer)
        before   = @($Before | Select-Object Account, App, Problem, Detail, Customer, Fixed)
        actions  = @($Actions)
        crashes  = @($Crashes | Select-Object App, Kind, Count, Last, Exe, Version, Module, Code)
        log      = $script:logFile
        diagnostics = $Diagnostics
        updates  = @($Updates)
        versions = $Versions
    }
    $headers = @{}
    if ($WebhookToken) { $headers['X-Watchdog-Token'] = $WebhookToken }
    $body = [Text.Encoding]::UTF8.GetBytes(($payload | ConvertTo-Json -Depth 5))
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    foreach ($try in 1..3) {
        try {
            Invoke-RestMethod -Uri $WebhookUrl -Method Post -ContentType 'application/json; charset=utf-8' -Body $body -Headers $headers -TimeoutSec 30 | Out-Null
            Write-Ok "Reported to n8n ($Kind)"
            return $true
        } catch {
            Write-Warn "Webhook attempt $try failed: $($_.Exception.Message)"
            if ($try -lt 3) { Start-Sleep -Seconds (10 * $try) }
        }
    }
    return $false
}

function Invoke-Diagnostics {
    <#
        Get-M365AppsLog.ps1 for the users and apps of these findings, into $DiagDir,
        before anything is repaired - afterwards the evidence is gone. Returns the zip,
        or '' when there is no collector or it wrote nothing. Keeps $LogDays of zips.
    #>
    param($Found, [datetime] $Since)
    New-Item -ItemType Directory -Path $DiagDir -Force | Out-Null
    Get-ChildItem -Path $DiagDir -Filter 'M365AppsLog_*.zip' -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogDays) } | Remove-Item -Force -ErrorAction SilentlyContinue
    $users = @($Found | Where-Object { $_.Account -ne $HostAccount } | ForEach-Object { $_.Account } | Sort-Object -Unique)
    $apps  = @($Found | ForEach-Object { $_.App } | Sort-Object -Unique)
    $hours = [Math]::Max(1, [Math]::Min(336, [int] [Math]::Ceiling(((Get-Date) - $Since).TotalHours) + 1))
    $parameters = [ordered]@{ App = $apps; Hours = $hours; OutputPath = $DiagDir; WorkingDir = $WorkingDir }
    if ($users.Count -gt 0) { $parameters['User'] = $users }
    $started = Get-Date
    Invoke-Helper 'Get-M365AppsLog.ps1' $parameters -Quiet | Out-Null
    # The zip is what travels; the folder next to it is the same content.
    Get-ChildItem -Path $DiagDir -Directory -Filter 'M365AppsLog_*' -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    $zip = @(Get-ChildItem -Path $DiagDir -Filter 'M365AppsLog_*.zip' -ErrorAction SilentlyContinue |
             Where-Object { $_.LastWriteTime -ge $started.AddSeconds(-5) } | Sort-Object LastWriteTime -Descending) | Select-Object -First 1
    if ($zip) { Write-Ok "Evidence collected: $($zip.FullName)"; return $zip.FullName }
    Write-Warn 'Evidence not collected'
    return ''
}

function Get-Signature {
    param($Found)
    return (@($Found | ForEach-Object { '{0}|{1}|{2}' -f $_.Account, $_.App, $_.Problem } | Sort-Object) -join ';')
}

# =====================================================================================
if ($Uninstall) {
    if ($PSCmdlet.ShouldProcess("$TaskName and $WorkingDir", 'Remove')) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        Get-ScheduledTask -TaskPath $ProbeTaskPath -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
        if (Test-Path $WorkingDir) { Remove-Item -Path $WorkingDir -Recurse -Force }
        Write-Ok "Task '$TaskName' and $WorkingDir removed"
    }
    exit 0
}

if ($Install) {
    if (-not $WebhookUrl) { Write-Warn 'No -WebhookUrl: the watchdog will repair and log, but report nothing' }
    if (-not $PSCmdlet.ShouldProcess("$env:COMPUTERNAME - task '$TaskName' every $IntervalMinutes min as System, in $WorkingDir", 'Install')) { exit 0 }

    if (-not (Protect-WorkingDir)) { Write-Bad "Could not lock $WorkingDir - not installed"; exit 1 }
    Write-Ok "$WorkingDir - System and Administrators only"

    $self = Join-Path $WorkingDir 'Watch-M365Apps.ps1'
    if ($PSCommandPath -ne $self) { Copy-Item -Path $PSCommandPath -Destination $self -Force }
    $helper = Find-Helper 'Repair-AppxPackageStore.ps1'
    if (-not $helper) { Write-Bad 'Repair-AppxPackageStore.ps1 not available - not installed'; exit 1 }
    $helperTarget = Join-Path $WorkingDir 'Repair-AppxPackageStore.ps1'
    if ($helper -ne $helperTarget) { Copy-Item -Path $helper -Destination $helperTarget -Force }
    Write-Ok 'Watch-M365Apps.ps1 and Repair-AppxPackageStore.ps1 in place'
    # The collector is a nice-to-have: without it the watchdog still repairs and reports.
    $collector = Find-Helper 'Get-M365AppsLog.ps1'
    if ($collector) {
        $collectorTarget = Join-Path $WorkingDir 'Get-M365AppsLog.ps1'
        if ($collector -ne $collectorTarget) { Copy-Item -Path $collector -Destination $collectorTarget -Force }
        Write-Ok 'Get-M365AppsLog.ps1 in place - evidence is collected before a repair'
    } else {
        Write-Warn 'Get-M365AppsLog.ps1 not available - the watchdog runs without collecting evidence'
    }

    [ordered]@{
        Account             = @($Account)
        App                 = @($App)
        WebhookUrl          = $WebhookUrl
        WebhookToken        = $WebhookToken
        RepairCooldownHours = $RepairCooldownHours
        RenotifyHours       = $RenotifyHours
        UpdateHours         = $UpdateHours
        CrashThreshold      = $CrashThreshold
        NoRepair            = [bool] $NoRepair
        NoUserRepair        = [bool] $NoUserRepair
        SkipLaunchTest      = [bool] $SkipLaunchTest
        NoUserLaunch        = [bool] $NoUserLaunch
        NoDiagnostics       = [bool] $NoDiagnostics
    } | ConvertTo-Json | Set-Content -Path $ConfigPath -Encoding UTF8
    Write-Ok "Settings in $ConfigPath"

    $action    = New-ScheduledTaskAction -Execute $nativeShell -Argument ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" -WorkingDir "{1}"' -f $self, $WorkingDir)
    $trigger   = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(2) -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes)
    $principal = New-ScheduledTaskPrincipal -UserId 'S-1-5-18' -LogonType ServiceAccount -RunLevel Highest
    $settings  = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 1)
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings `
        -Description 'Tests Teams, new Outlook and Copilot with the IT accounts, repairs them and reports to n8n (M365-Scripts, scripts/RDS/Watch-M365Apps.ps1)' -Force | Out-Null
    Write-Ok "Task '$TaskName' registered - first run in 2 minutes, then every $IntervalMinutes minutes"
    exit 0
}

if ($TestNotification) {
    $sent = Send-Notification -Kind 'test' -Summary "Test message from Watch-M365Apps on $env:COMPUTERNAME"
    if ($sent) { exit 0 } else { exit 1 }
}

# -- One run -------------------------------------------------------------------------
$transcribing   = $false
$logDir = Join-Path $WorkingDir 'Logs'
# A run before -Install would otherwise create the folder - where the helper is
# downloaded and run from - with whatever C:\IT lets users do.
if (-not (Protect-WorkingDir)) { Write-Bad "Could not lock $WorkingDir - stopped"; exit 1 }
try {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    Get-ChildItem -Path $logDir -Filter '*.log' | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogDays) } | Remove-Item -Force
    $script:logFile = Join-Path $logDir ('Watch-M365Apps_{0}.log' -f (Get-Date -Format 'yyyyMMdd'))
    Start-Transcript -Path $script:logFile -Append | Out-Null
    $transcribing = $true
} catch {
    Write-Warn "No transcript: $($_.Exception.Message)"
}

$exitCode = 0
try {
    Write-Host ''
    Write-Host ("  M365 app watchdog - {0} - {1:yyyy-MM-dd HH:mm}" -f $env:COMPUTERNAME, (Get-Date)) -ForegroundColor Cyan

    $runStart = Get-Date
    # RepairingSignature / LastRepairing: what the last "repairing" report was about, so a
    # problem the repair cannot fix does not announce itself again every run.
    # Versions: the build per app on the host at the last run, so a new one is reported once.
    $state = [PSCustomObject]@{ Signature = ''; LastNotified = $null; LastRepair = $null; LastRun = $null; UserRepairs = @{}
                                RepairingSignature = ''; LastRepairing = $null; LastUpdate = $null; Versions = $null }
    if (Test-Path $StatePath) {
        $saved = Get-Content -Path $StatePath -Raw | ConvertFrom-Json
        foreach ($name in 'Signature', 'LastNotified', 'LastRepair', 'LastRun', 'RepairingSignature', 'LastRepairing', 'LastUpdate', 'Versions') { $state.$name = Get-PropertyValue $saved $name }
        if ($null -eq $state.Signature) { $state.Signature = '' }
        if ($null -eq $state.RepairingSignature) { $state.RepairingSignature = '' }
        # Per user and app, when their app was last re-registered; older than a day is forgotten.
        $savedRepairs = Get-PropertyValue $saved 'UserRepairs'
        if ($savedRepairs) {
            foreach ($prop in $savedRepairs.PSObject.Properties) {
                $when = [datetime]::Parse([string] $prop.Value, $null, [Globalization.DateTimeStyles]::RoundtripKind)
                if ($when -gt $runStart.AddDays(-1)) { $state.UserRepairs[$prop.Name] = [string] $prop.Value }
            }
        }
    }

    # Events since the previous run; the first run looks back one hour, and a long gap
    # (host off, task disabled) no more than a day.
    $since = $runStart.AddHours(-1)
    if ($state.LastRun) { $since = [datetime]::Parse($state.LastRun, $null, [Globalization.DateTimeStyles]::RoundtripKind) }
    if ($since -lt $runStart.AddDays(-1)) { $since = $runStart.AddDays(-1) }

    $sessions = @(Get-UserSession)
    Write-Step ('Apps users could not open since {0:yyyy-MM-dd HH:mm}' -f $since)
    $openFailures = @(Get-OpenFailure -Since $since -Sessions $sessions)
    if ($openFailures.Count -eq 0) { Write-Ok 'None' }

    $crashes = @()
    if ($CrashThreshold -gt 0) {
        Write-Step ('Crashes and hangs since {0:yyyy-MM-dd HH:mm}' -f $since)
        $crashes = @(Get-AppCrash -Since $since | Where-Object { $_.Count -ge $CrashThreshold })
        if ($crashes.Count -eq 0) { Write-Ok 'None' }
        foreach ($c in $crashes) {
            Write-Warn ('{0} {1} {2}x ({3} {4}, {5}{6}) - last at {7:HH:mm}' -f $c.App, $c.Kind.ToLower(), $c.Count, $c.Exe, $c.Version, $c.Module, $(if ($c.Code) { ', ' + $c.Code } else { '' }), [datetime] $c.Last)
        }
    }

    # Updates before the checks, so the checks test the newest build: the host every
    # -UpdateHours, our own account every run it is behind what the host provisions.
    $updateActions = @()
    if ($UpdateHours -gt 0 -and -not $NoRepair -and -not $WhatIfPreference) {
        $lastUpdate = $null
        if ($state.LastUpdate) { $lastUpdate = [datetime]::Parse($state.LastUpdate, $null, [Globalization.DateTimeStyles]::RoundtripKind) }
        if (-not $lastUpdate -or $lastUpdate.AddHours($UpdateHours) -lt (Get-Date)) {
            Write-Host ''
            Write-Host '  ==== Updating '.PadRight(80, '=') -ForegroundColor Cyan
            $updateActions += @(Update-HostApp -State $state)
        }
        $updateActions += @(Update-AccountApp -Sessions $sessions)
    }
    # A new build on the host, from the update above or on its own (Edge Update, an
    # admin): reported once, with what it was before. The first run only records.
    $hostVersions = Get-HostVersion
    $newBuilds    = @(foreach ($name in @($hostVersions.Keys)) {
        $was = [string] (Get-PropertyValue $state.Versions $name)
        if ($was -and $was -ne $hostVersions[$name]) { '{0} {1} -> {2}' -f $name, $was, $hostVersions[$name] }
    })
    foreach ($line in $newBuilds) { Write-Ok "New build on the host: $line" }
    $state.Versions = $hostVersions

    # Events are what happened since the last run; the read-back can only say whether
    # what they point at is still wrong, so an event finding stays unless it was fixed.
    $before  = @(@(Get-Finding -Sessions $sessions) + $openFailures)
    $after   = $before
    $actions = @()
    $repair  = $before.Count -gt 0 -and -not $NoRepair -and -not $WhatIfPreference

    # Something new is wrong: first the evidence, then word to n8n that it is broken and
    # being repaired - for every user it concerns - and only then the repair. The same
    # problem again within -RenotifyHours (a repair that did not take) is not announced
    # or collected twice.
    $diagnostics = ''
    if ($before.Count -gt 0) {
        $beforeSignature = Get-Signature $before
        $lastRepairing   = $null
        if ($state.LastRepairing) { $lastRepairing = [datetime]::Parse($state.LastRepairing, $null, [Globalization.DateTimeStyles]::RoundtripKind) }
        $isNew = $beforeSignature -ne $state.RepairingSignature -or -not $lastRepairing -or $lastRepairing.AddHours($RenotifyHours) -lt (Get-Date)
        if ($isNew) {
            if (-not $NoDiagnostics) {
                Write-Host ''
                Write-Host '  ==== Collecting evidence '.PadRight(80, '=') -ForegroundColor Cyan
                try { $diagnostics = Invoke-Diagnostics -Found $before -Since $since }
                catch { Write-Warn "Evidence not collected: $($_.Exception.Message)" }
            }
            if ($repair) {
                $summary = '{0}: {1} problem(s) found, repairing now - {2}' -f $env:COMPUTERNAME, $before.Count, ((@($before | ForEach-Object { "$($_.Account): $($_.App) $($_.Problem)" })) -join ', ')
                Send-Notification -Kind 'repairing' -Summary $summary -Found $before -Diagnostics $diagnostics | Out-Null
            }
            $state.RepairingSignature = $beforeSignature
            $state.LastRepairing      = (Get-Date).ToString('o')
        }
    }

    if ($repair) {
        Write-Host ''
        Write-Host '  ==== Repairing '.PadRight(80, '=') -ForegroundColor Cyan
        $actions = @(Invoke-Repair -Found $before -State $state)
        Write-Host ''
        Write-Host '  ==== Read back '.PadRight(80, '=') -ForegroundColor Cyan
        $after = @(@(Get-Finding -Sessions @(Get-UserSession)) + @($openFailures | Where-Object { -not $_.Fixed }))
    } elseif ($before.Count -gt 0) {
        Write-Skip 'Not repaired (-NoRepair)'
    }

    $signature = Get-Signature $after
    $lastSent  = $null
    if ($state.LastNotified) { $lastSent = [datetime]::Parse($state.LastNotified, $null, [Globalization.DateTimeStyles]::RoundtripKind) }
    $kind = $null
    if ($before.Count -gt 0 -and $after.Count -eq 0) {
        $kind   = 'repaired'
        $summary = '{0}: {1} problem(s) found and repaired - {2}' -f $env:COMPUTERNAME, $before.Count, ((@($before | ForEach-Object { "$($_.Account): $($_.App) $($_.Problem)" })) -join ', ')
    } elseif ($after.Count -gt 0) {
        if ($signature -ne $state.Signature -or -not $lastSent -or $lastSent.AddHours($RenotifyHours) -lt (Get-Date) -or $actions.Count -gt 0) {
            $kind   = if ($repair) { 'repair-failed' } else { 'failing' }
            $summary = '{0}: {1} problem(s) {2} - {3}' -f $env:COMPUTERNAME, $after.Count, $(if ($repair) { 'left after repair' } else { 'found' }), ((@($after | ForEach-Object { "$($_.Account): $($_.App) $($_.Problem)" })) -join ', ')
        }
    } elseif ($state.Signature) {
        $kind   = 'recovered'
        $summary = '{0}: healthy again without a repair (was: {1})' -f $env:COMPUTERNAME, $state.Signature
    }

    # A new build, like a crash, rides along with any report of this run or is a report
    # of its own - a healthy run after it means our account started it.
    if ($newBuilds.Count -gt 0) {
        $buildText = $newBuilds -join ', '
        if ($kind) { $summary += " - also new: $buildText" }
        else {
            $kind    = 'updated'
            $tested  = if ($after.Count -gt 0) { "$($after.Count) problem(s) still open, see the findings" }
                       elseif (@($sessions | Where-Object { $_.Watched }).Count -gt 0 -and -not $SkipLaunchTest) { 'running with our own account' }
                       else { 'not started - no watched account signed in, or the launch test is off' }
            $summary = '{0}: new build(s) {1} - {2}' -f $env:COMPUTERNAME, $buildText, $tested
        }
    }

    # Crashes are events, not a state: they ride along with any report of this run,
    # or are a report of their own.
    if ($crashes.Count -gt 0) {
        $crashText = (@($crashes | ForEach-Object { '{0} {1} {2}x' -f $_.App, $_.Kind.ToLower(), $_.Count })) -join ', '
        if ($kind) { $summary += " - also: $crashText" }
        else {
            $kind    = 'crashed'
            $summary = '{0}: {1} since the last check' -f $env:COMPUTERNAME, $crashText
        }
    }

    if ($kind) {
        if (Send-Notification -Kind $kind -Summary $summary -Found $after -Before $before -Actions @($updateActions + $actions) -Crashes $crashes -Diagnostics $diagnostics -Updates $newBuilds -Versions $hostVersions) {
            if ($kind -notin 'crashed', 'updated') { $state.LastNotified = (Get-Date).ToString('o') }
        }
    }
    $state.Signature = $signature
    $state.LastRun   = $runStart.ToString('o')
    if (Test-Path $WorkingDir) { $state | ConvertTo-Json | Set-Content -Path $StatePath -Encoding UTF8 }

    Write-Host ''
    if ($after.Count -eq 0) {
        Write-Ok $(if ($before.Count -gt 0) { 'Everything repaired' } else { 'Everything healthy' })
    } else {
        Write-Host ("  {0} problem(s) left:" -f $after.Count) -ForegroundColor Yellow
        $after | Format-Table Account, App, Problem, Detail -AutoSize -Wrap | Out-String -Width 200 | Write-Host
        $exitCode = 1
    }
} catch {
    Write-Bad "Aborted: $($_.Exception.Message)"
    Send-Notification -Kind 'error' -Summary "$($env:COMPUTERNAME): the watchdog run failed - $($_.Exception.Message)" | Out-Null
    $exitCode = 1
} finally {
    if ($transcribing) { try { Stop-Transcript | Out-Null } catch { } }
}
exit $exitCode
