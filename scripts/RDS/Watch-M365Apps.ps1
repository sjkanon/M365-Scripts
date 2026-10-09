#Requires -Version 5.1
<#
.SYNOPSIS
    Watchdog for new Teams, new Outlook and Copilot on a session host: tests them with
    our own accounts (itceadmin, itce.user), repairs what is broken before a customer
    runs into it, and reports to an n8n webhook.

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
      2. Accounts    For each watched account signed in on this host (the owner of an
                     explorer.exe), per app: the package is registered for that user,
                     its files are there and its status is Ok. Then the app has to be
                     running in that session; when it is not, it is started there
                     (shell:AppsFolder\<AUMID>, through a one-off task in the user's
                     own session) and must still be running 15 seconds later.
      3. Repair      Host problems and packages whose files are gone: Repair-AppxPackageStore.ps1
                     -Provision (Microsoft's installers, signature-checked), at most once
                     per -RepairCooldownHours. Then, in our own account only: a package
                     that is not registered is registered by family name, an app that
                     does not start is reset (Reset-AppxPackage). Nothing is closed,
                     removed or reset for any other user - no -RemoveOld, no -Latest.
      4. Read back   Steps 1 and 2 again.
      5. Report      A JSON POST to -WebhookUrl when something is wrong, was repaired,
                     recovered on its own, or crashed at least -CrashThreshold times -
                     not on every healthy run. A problem that stays is reported again
                     after -RenotifyHours.

    An account that is not signed in on this host is skipped: its packages live in its
    FSLogix container and cannot be tested without it. The watched accounts are
    canaries - keep a session of each open (disconnected is fine) on every host.

    -Install copies this script and Repair-AppxPackageStore.ps1 (from ..\Device, or from
    GitHub at the commit and SHA-256 pinned below) to -WorkingDir, locks that folder to
    System and Administrators, writes the settings to config.json there, and registers
    the task "M365 App Watchdog". The webhook URL and token live only in that file.

.PARAMETER Account
    Accounts to test with, by user name, UPN or DOMAIN\user. Default: itceadmin, itce.user.

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

.PARAMETER CrashThreshold
    Report crashes and hangs of one app once it reaches this many since the last run.
    Default 1 (every crash); 0 turns crash reporting off. Raise it on a busy pool where
    the odd crash is noise.

.PARAMETER NoRepair
    Test and report only, change nothing.

.PARAMETER SkipLaunchTest
    Do not start an app that is not running; only check its registration.

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
    [string[]] $Account = @('itceadmin', 'itce.user'),
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
    [ValidateRange(0, 1000)]
    [int]      $CrashThreshold = 1,
    [switch]   $NoRepair,
    [switch]   $SkipLaunchTest,
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
    [PSCustomObject]@{ App = 'Copilot'; Packages = @('Microsoft.MicrosoftOfficeHub', 'Microsoft.Copilot'); Exe = @('M365Copilot.exe'); Repair = 'copilot' }
)
$PublisherId   = '8wekyb3d8bbwe'   # Microsoft's Store publisher id, the same for all four packages
$CopilotGuid   = '{C50565E9-CCCF-44B4-BA15-5AC5C6569197}'
$TaskName      = 'M365 App Watchdog'
$ProbeTaskPath = '\M365AppWatchdog\'
$HostAccount   = '(host)'
$LaunchSeconds = 60
$LogDays       = 14
# The repair helper, when it is not next to this script: fetched from this repo at a
# pinned commit of main and refused unless the SHA-256 matches - the same pin as
# Update-SessionHostImage.ps1. Move both together.
$HelperCommit = '048cf967673f02131ddf4efe8a3a8fe72e55f90c'   # main, 2026-10-08
$HelperHashes = @{
    'Repair-AppxPackageStore.ps1' = '88A9CEAB9F801CF60BFD25AEB69320A97AECD9F19323BC6E354A60118A04C45C'
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

# -- Settings ------------------------------------------------------------------------
# The task passes nothing but -WorkingDir: everything else comes from config.json,
# and a parameter given on the command line wins over it.
$ConfigPath = Join-Path $WorkingDir 'config.json'
$StatePath  = Join-Path $WorkingDir 'state.json'
$Installed  = (Test-Path $PSScriptRoot) -and ((Resolve-Path $PSScriptRoot).Path.TrimEnd('\') -eq ([IO.Path]::GetFullPath($WorkingDir)).TrimEnd('\'))
if (-not $Install -and (Test-Path $ConfigPath)) {
    $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
    foreach ($name in 'Account', 'App', 'WebhookUrl', 'WebhookToken', 'RepairCooldownHours', 'RenotifyHours', 'CrashThreshold', 'NoRepair', 'SkipLaunchTest') {
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
        Next to this script in the installed copy (its locked folder), ..\Device in a
        checkout of the repo, and otherwise GitHub at $HelperCommit when its SHA-256
        matches. Not "next to this script" anywhere else: downloaded to a folder such
        as C:\IT\Setup, whatever sits beside it - or in C:\IT\Device - could have been
        put there by any user, and -Install would hand it to System.
    #>
    param([string] $Name)
    $candidates = @()
    if ($Installed) {
        $candidates += Join-Path $PSScriptRoot $Name
    } elseif (Test-Path (Join-Path $PSScriptRoot '..\..\menu.ps1')) {
        $candidates += Join-Path (Join-Path (Split-Path $PSScriptRoot) 'Device') $Name
    }
    foreach ($path in $candidates) {
        if (Test-Path $path) { return $path }
    }
    if (-not $HelperHashes.ContainsKey($Name)) { return $null }

    $url    = 'https://raw.githubusercontent.com/sjkanon/M365-Scripts/{0}/scripts/Device/{1}' -f $HelperCommit, $Name
    $target = Join-Path $WorkingDir $Name
    try {
        New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $url -OutFile $target -UseBasicParsing
        $hash = (Get-FileHash -Path $target -Algorithm SHA256).Hash
        if ($hash -ne $HelperHashes[$Name]) {
            throw "SHA-256 is $hash, expected $($HelperHashes[$Name]) - not used"
        }
        Write-Ok "$Name fetched from GitHub at $($HelperCommit.Substring(0, 7)), SHA-256 verified"
        return $target
    } catch {
        Remove-Item $target -Force -ErrorAction SilentlyContinue
        Write-Bad "$Name from $url - $($_.Exception.Message)"
        return $null
    }
}

function Invoke-Helper {
    <# Run a repo script in its own process, so its exit does not end this one. #>
    param([string] $Name, [System.Collections.IDictionary] $Parameters)
    $path = Find-Helper $Name
    if (-not $path) { Write-Bad "$Name not found locally and not verified from GitHub - skipped"; return 1 }
    $command = ConvertTo-ScriptCommand -Path $path -Parameters $Parameters
    Write-Step $command
    & $nativeShell -NoProfile -ExecutionPolicy Bypass -Command $command | Out-Host
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

function Get-AccountSession {
    <#
        Where a watched account is signed in on this host: the owner of each
        explorer.exe, matched on the user name alone, so AD, Entra (AzureAD\) and
        local accounts all count. One row per account.
    #>
    $seen = @{}
    foreach ($proc in @(Get-CimInstance -ClassName Win32_Process -Filter "Name = 'explorer.exe'" -ErrorAction SilentlyContinue)) {
        $owner = Invoke-CimMethod -InputObject $proc -MethodName GetOwner -ErrorAction SilentlyContinue
        if (-not $owner -or $owner.ReturnValue -ne 0) { continue }
        foreach ($name in $Account) {
            $short = (($name -split '\\')[-1] -split '@')[0]
            if ($owner.User -ine $short -or $seen.ContainsKey($name)) { continue }
            $sid = (Invoke-CimMethod -InputObject $proc -MethodName GetOwnerSid -ErrorAction SilentlyContinue).Sid
            if (-not $sid) { continue }
            $seen[$name] = $true
            [PSCustomObject]@{ Account = $name; User = '{0}\{1}' -f $owner.Domain, $owner.User; Sid = $sid; SessionId = $proc.SessionId }
        }
    }
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

# -- Checks --------------------------------------------------------------------------
function New-Finding {
    param([string] $Account, [string] $App, [string] $Problem, [string] $Detail, [string] $Package)
    Write-Bad "$App - $Detail"
    [PSCustomObject]@{ Account = $Account; App = $App; Problem = $Problem; Detail = $Detail; Package = $Package }
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
    param($Session, $Entry, [string[]] $Provisioned)
    $pkg = $null
    foreach ($name in $Entry.Packages) {
        $pkg = @(Get-AppxPackage -User $Session.Sid -Name $name -ErrorAction SilentlyContinue) |
               Sort-Object { [version] $_.Version } -Descending | Select-Object -First 1
        if ($pkg) { break }
    }
    if (-not $pkg) {
        if ($Entry.App -eq 'Copilot' -and (Get-EdgeUpdateClientVersion $CopilotGuid)) {
            Write-Ok 'Copilot: unified app (Edge Update, machine-wide) - not started by this test'
            return
        }
        $target = @($Entry.Packages | Where-Object { $_ -in $Provisioned }) | Select-Object -First 1
        if (-not $target) { $target = $Entry.Packages[0] }
        New-Finding $Session.Account $Entry.App 'NotRegistered' 'not registered for this user' $target
        return
    }

    $location = Get-PropertyValue $pkg 'InstallLocation'
    if (-not $location -or -not (Test-Path -LiteralPath $location)) {
        New-Finding $Session.Account $Entry.App 'Broken' "$($pkg.PackageFullName): its files are gone" $pkg.Name
        return
    }
    $status = [string] (Get-PropertyValue $pkg 'Status')
    if ($status -and $status -ne 'Ok') {
        New-Finding $Session.Account $Entry.App 'Broken' "$($pkg.PackageFullName): status is $status" $pkg.Name
        return
    }

    $appEntry = Get-AppEntry $pkg
    if (-not $appEntry) { Write-Warn "$($Entry.App) $($pkg.Version) registered, but its manifest names no executable - not started"; return }
    if (@(Get-SessionProcess $Session $appEntry.Exe).Count -gt 0) { Write-Ok "$($Entry.App) $($pkg.Version) running"; return }
    if ($SkipLaunchTest) { Write-Ok "$($Entry.App) $($pkg.Version) registered (not running, launch test off)"; return }
    if (Test-AppStart $Session $appEntry) { Write-Ok "$($Entry.App) $($pkg.Version) started"; return }
    New-Finding $Session.Account $Entry.App 'WontStart' "$($pkg.Version) is registered but $($appEntry.Exe).exe did not start or did not stay up" $pkg.Name
}

function Get-Finding {
    <# Every check, read only - apart from starting an app that is not running in our own session. #>
    $provisioned = @(Get-AppxProvisionedPackage -Online | ForEach-Object { $_.DisplayName })
    $sessions    = @(Get-AccountSession)
    Write-Step 'Host'
    foreach ($entry in $Apps) { Test-HostApp $entry $provisioned }
    foreach ($name in $Account) {
        $session = @($sessions | Where-Object { $_.Account -eq $name }) | Select-Object -First 1
        if (-not $session) { Write-Step $name; Write-Skip 'Not signed in on this host - nothing to test'; continue }
        Write-Step ('{0} (session {1})' -f $session.User, $session.SessionId)
        foreach ($entry in $Apps) { Test-UserApp $session $entry $provisioned }
    }
}

# -- Repair --------------------------------------------------------------------------
function Invoke-Repair {
    <#
        Host first - our own account's fixes need the package on the machine - then our
        account. Returns one line per action for the report.
    #>
    param($Found, $State)
    $actions = [System.Collections.Generic.List[string]]::new()

    $hostApps = @($Found | Where-Object { $_.Account -eq $HostAccount -or $_.Problem -eq 'Broken' } | ForEach-Object { $_.App } | Sort-Object -Unique)
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

    $sessions = @(Get-AccountSession)
    foreach ($finding in @($Found | Where-Object { $_.Account -ne $HostAccount })) {
        $session = @($sessions | Where-Object { $_.Account -eq $finding.Account }) | Select-Object -First 1
        if (-not $session) { continue }
        $family = '{0}_{1}' -f $finding.Package, $PublisherId
        if ($finding.Problem -eq 'WontStart') {
            $verb    = 'reset'
            $command = "try { Get-AppxPackage -Name '$($finding.Package)' | Reset-AppxPackage -ErrorAction Stop; exit 0 } catch { exit 1 }"
        } else {
            $verb    = 're-registered'
            $command = "try { Add-AppxPackage -RegisterByFamilyName -MainPackage '$family' -ErrorAction Stop; exit 0 } catch { exit 1 }"
        }
        Write-Step "$($finding.App) for $($session.User): $verb in that session"
        $code = Invoke-AsUser -Session $session -Execute $nativeShell -Argument "-NoProfile -NonInteractive -WindowStyle Hidden -Command `"$command`"" -WaitSeconds 300
        if ($code -eq 0) { Write-Ok "$($finding.App) $verb" } else { Write-Warn "$($finding.App) not $verb (result $code)" }
        $actions.Add("$($finding.Account) $($finding.App): $verb as the user - result $code")
    }
    return $actions
}

# -- Report --------------------------------------------------------------------------
function Send-Notification {
    param([string] $Kind, [string] $Summary, $Found = @(), $Before = @(), $Actions = @(), $Crashes = @())
    if (-not $WebhookUrl) { Write-Skip "No -WebhookUrl - not reported ($Kind)"; return $false }
    $payload = [ordered]@{
        source   = 'Watch-M365Apps'
        event    = $Kind
        host     = $env:COMPUTERNAME
        time     = (Get-Date).ToString('o')
        summary  = $Summary
        accounts = @($Account)
        findings = @($Found  | Select-Object Account, App, Problem, Detail)
        before   = @($Before | Select-Object Account, App, Problem, Detail)
        actions  = @($Actions)
        crashes  = @($Crashes | Select-Object App, Kind, Count, Last, Exe, Version, Module, Code)
        log      = $script:logFile
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

    [ordered]@{
        Account             = @($Account)
        App                 = @($App)
        WebhookUrl          = $WebhookUrl
        WebhookToken        = $WebhookToken
        RepairCooldownHours = $RepairCooldownHours
        RenotifyHours       = $RenotifyHours
        CrashThreshold      = $CrashThreshold
        NoRepair            = [bool] $NoRepair
        SkipLaunchTest      = [bool] $SkipLaunchTest
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
$script:logFile = $null
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
    $state = [PSCustomObject]@{ Signature = ''; LastNotified = $null; LastRepair = $null; LastRun = $null }
    if (Test-Path $StatePath) {
        $saved = Get-Content -Path $StatePath -Raw | ConvertFrom-Json
        foreach ($name in 'Signature', 'LastNotified', 'LastRepair', 'LastRun') { $state.$name = Get-PropertyValue $saved $name }
        if ($null -eq $state.Signature) { $state.Signature = '' }
    }

    # Crashes since the previous run; the first run looks back one hour, and a long
    # gap (host off, task disabled) no more than a day.
    $crashes = @()
    if ($CrashThreshold -gt 0) {
        $since = $runStart.AddHours(-1)
        if ($state.LastRun) { $since = [datetime]::Parse($state.LastRun, $null, [Globalization.DateTimeStyles]::RoundtripKind) }
        if ($since -lt $runStart.AddDays(-1)) { $since = $runStart.AddDays(-1) }
        Write-Step ('Crashes and hangs since {0:yyyy-MM-dd HH:mm}' -f $since)
        $crashes = @(Get-AppCrash -Since $since | Where-Object { $_.Count -ge $CrashThreshold })
        if ($crashes.Count -eq 0) { Write-Ok 'None' }
        foreach ($c in $crashes) {
            Write-Warn ('{0} {1} {2}x ({3} {4}, {5}{6}) - last at {7:HH:mm}' -f $c.App, $c.Kind.ToLower(), $c.Count, $c.Exe, $c.Version, $c.Module, $(if ($c.Code) { ', ' + $c.Code } else { '' }), [datetime] $c.Last)
        }
    }

    $before  = @(Get-Finding)
    $after   = $before
    $actions = @()
    $repair  = $before.Count -gt 0 -and -not $NoRepair -and -not $WhatIfPreference
    if ($repair) {
        Write-Host ''
        Write-Host '  ==== Repairing '.PadRight(80, '=') -ForegroundColor Cyan
        $actions = @(Invoke-Repair -Found $before -State $state)
        Write-Host ''
        Write-Host '  ==== Read back '.PadRight(80, '=') -ForegroundColor Cyan
        $after = @(Get-Finding)
    } elseif ($before.Count -gt 0) {
        Write-Skip 'Not repaired (-NoRepair)'
    }

    $signature = Get-Signature $after
    $lastSent  = $null
    if ($state.LastNotified) { $lastSent = [datetime]::Parse($state.LastNotified, $null, [Globalization.DateTimeStyles]::RoundtripKind) }
    $kind = $null
    if ($before.Count -gt 0 -and $after.Count -eq 0) {
        $kind   = 'repaired'
        $summary = '{0}: {1} problem(s) found and repaired - {2}' -f $env:COMPUTERNAME, $before.Count, ((@($before | ForEach-Object { "$($_.Account) $($_.App) $($_.Problem)" })) -join ', ')
    } elseif ($after.Count -gt 0) {
        if ($signature -ne $state.Signature -or -not $lastSent -or $lastSent.AddHours($RenotifyHours) -lt (Get-Date) -or $actions.Count -gt 0) {
            $kind   = if ($repair) { 'repair-failed' } else { 'failing' }
            $summary = '{0}: {1} problem(s) {2} - {3}' -f $env:COMPUTERNAME, $after.Count, $(if ($repair) { 'left after repair' } else { 'found' }), ((@($after | ForEach-Object { "$($_.Account) $($_.App) $($_.Problem)" })) -join ', ')
        }
    } elseif ($state.Signature) {
        $kind   = 'recovered'
        $summary = '{0}: healthy again without a repair (was: {1})' -f $env:COMPUTERNAME, $state.Signature
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
        if (Send-Notification -Kind $kind -Summary $summary -Found $after -Before $before -Actions $actions -Crashes $crashes) {
            if ($kind -ne 'crashed') { $state.LastNotified = (Get-Date).ToString('o') }
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
