#Requires -Version 5.1
<#
.SYNOPSIS
    Install or update the M365 app watchdog on one or more session hosts from this
    checkout - keeping each host's settings - and optionally check or repair them
    straight after.

.DESCRIPTION
    Puts four scripts in -WorkingDir (C:\IT\AppWatchdog) on each host and registers the
    watchdog task with Watch-M365Apps.ps1 -Install. From the checkout this script is
    in; downloaded on its own to a session host - no checkout - from GitHub at the
    commit pinned below, each file refused unless its SHA-256 matches:

      Watch-M365Apps.ps1           the watchdog (scheduled task, as System)
      Repair-AppxPackageStore.ps1  its host repair (scripts\Device)
      Get-M365AppsLog.ps1          its evidence collector
      Repair-M365AppsHost.ps1      repair and test the host in one go, by hand

    Per host, in order:

      1. Lock      -WorkingDir is locked to System and Administrators before anything
                   is copied into it - System runs what is there.
      2. Copy      The four scripts, over PowerShell remoting with -ComputerName, and
                   read back by SHA-256 against this checkout.
      3. Install   Watch-M365Apps.ps1 -Install, from the installed copy. -Install does
                   not read the existing config.json - run without -WebhookUrl, it
                   would write a config without the webhook - so the host's own
                   config.json is read first, on the host (the token never leaves
                   it), and every setting in it is passed on, as well as the interval
                   of the existing task. A parameter given here wins. The watchdog's
                   state (what it reported, its cooldowns) is kept.
      4. Verify    The task is there and enabled, and runs the installed copy.
      5. Then      -Then Check: Repair-M365AppsHost.ps1 -CheckOnly. -Then Repair: the
                   full repair and test (FSLogix InstallAppxPackages = 0, package
                   store, Teams and Outlook at the newest build, a watchdog run), with
                   -Restart when given. One host after the other, so a restart never
                   takes the whole pool down at once.

    Ends with one table across the hosts.

.PARAMETER ComputerName
    Session hosts to install on, e.g. lem-avd-4,lem-avd-5,lem-avd-6, over PowerShell
    remoting (WinRM). Without it: this machine.

.PARAMETER Credential
    Credential for the remoting sessions, when the current account is not an
    administrator on the hosts.

.PARAMETER WebhookUrl
    n8n webhook. Needed on a host without a config.json; on one with it, replaces
    the URL there.

.PARAMETER WebhookToken
    X-Watchdog-Token for the webhook; replaces the one in config.json when given.

.PARAMETER Account
    Accounts the watchdog tests with; replaces the list in config.json when given.

.PARAMETER App
    Apps to watch (Teams, Outlook, Copilot, SnippingTool, Photos); replaces the list
    in config.json when given.

.PARAMETER IntervalMinutes
    How often the task runs; default the interval of the existing task, else 30.

.PARAMETER Then
    After installing: None (default), Check (Repair-M365AppsHost.ps1 -CheckOnly) or
    Repair (Repair-M365AppsHost.ps1, repairs and tests).

.PARAMETER Restart
    With -Then Repair: restart a host when a restart is due and nobody is signed in.

.PARAMETER WorkingDir
    The watchdog's folder on the hosts. Default C:\IT\AppWatchdog.

.EXAMPLE
    # Update the whole pool, keeping every host's webhook and settings, and look
    .\Install-M365AppsWatchdog.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Then Check -Confirm:$false

.EXAMPLE
    # A new host
    .\Install-M365AppsWatchdog.ps1 -ComputerName lem-avd-7 -WebhookUrl 'https://n8n.example.com/webhook/m365-apps' -WebhookToken '<token>' -Confirm:$false

.EXAMPLE
    # Update and repair the drained pool, restarting each host when that is due
    .\Install-M365AppsWatchdog.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Then Repair -Restart -Confirm:$false

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows PowerShell 5.1 or PowerShell 7, from a checkout of this repo or on
              its own on the host (GitHub reachable); the hosts: Windows 11 Enterprise
              multi-session / Windows Server with FSLogix
    Exit    : 0 installed everywhere (and Then healthy), 1 a host failed, 2 installed
              and healthy but a host needs a restart
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param (
    [string[]]     $ComputerName,
    [pscredential] $Credential,
    [string]       $WebhookUrl,
    [string]       $WebhookToken,
    [string[]]     $Account,
    [ValidateSet('Teams', 'Outlook', 'Copilot', 'SnippingTool', 'Photos')]
    [string[]]     $App,
    [ValidateRange(5, 1440)]
    [int]          $IntervalMinutes,
    [ValidateSet('None', 'Check', 'Repair')]
    [string]       $Then = 'None',
    [switch]       $Restart,
    [string]       $WorkingDir = 'C:\IT\AppWatchdog'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -- Output ------------------------------------------------------------------------
function Write-Step { param([string] $Message) Write-Host ''; Write-Host "  $Message" -ForegroundColor Cyan }
function Write-Ok   { param([string] $Message) Write-Host "  [ OK ] $Message" -ForegroundColor Green }
function Write-Skip { param([string] $Message) Write-Host "  [SKIP] $Message" -ForegroundColor DarkGray }
function Write-Warn { param([string] $Message) Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
function Write-Bad  { param([string] $Message) Write-Host "  [FAIL] $Message" -ForegroundColor Red }

function Get-Sha256 {
    <# Not Get-FileHash: under -WhatIf, Windows PowerShell 5.1's returns nothing. #>
    param([string] $Path)
    $stream = [IO.File]::OpenRead((Resolve-Path -LiteralPath $Path).ProviderPath)
    try {
        $sha = [Security.Cryptography.SHA256]::Create()
        return -join @($sha.ComputeHash($stream) | ForEach-Object { $_.ToString('X2') })
    } finally {
        $stream.Dispose()
    }
}

# -- What is installed ---------------------------------------------------------------
# Without a checkout - this script downloaded on its own to a session host - the four
# scripts come from GitHub at this commit, each refused unless its SHA-256 matches.
# Move the commit and the hashes together.
$Pin = @{
    Commit = 'c331d22d69bd8b83f90e612039fe4062b9d645ae'
    Files  = [ordered]@{
        'Watch-M365Apps.ps1'          = @{ Path = 'scripts/RDS/Watch-M365Apps.ps1';             Hash = '371207FD6C9641AE4ADBEA8E29453584CF55F77DBE0E96216CA76051116AF10C' }
        'Repair-AppxPackageStore.ps1' = @{ Path = 'scripts/Device/Repair-AppxPackageStore.ps1'; Hash = 'B512C5EE1B438AD1BE6D9CF6817F098C41FBEFD1B12EC8810D502D0679126897' }
        'Get-M365AppsLog.ps1'         = @{ Path = 'scripts/RDS/Get-M365AppsLog.ps1';            Hash = '7A77015B171FFFBBDCA60BACE23DD017DE1ED337D8AA2311F5E84CB8F359FA4F' }
        'Repair-M365AppsHost.ps1'     = @{ Path = 'scripts/RDS/Repair-M365AppsHost.ps1';        Hash = 'A512EC7C828DC86DD8AFADFEEC165AE4FFDEE3B7E76FD73F658E957D0C57604E' }
    }
}
$Sources = [ordered]@{}
$Hashes  = @{}
$staging = $null
if (Test-Path (Join-Path $PSScriptRoot '..\..\menu.ps1')) {
    foreach ($name in $Pin.Files.Keys) {
        $path = Join-Path $PSScriptRoot ($Pin.Files[$name].Path -replace '^scripts/RDS/', '' -replace '^scripts/Device/', '..\Device\')
        if (-not (Test-Path $path)) { Write-Error "$path not found in this checkout"; exit 1 }
        $Sources[$name] = $path
        $Hashes[$name]  = Get-Sha256 $path
    }
    $Commit = try { (& git -C $PSScriptRoot log -1 --format='%h %cs' 2>$null) } catch { '' }
} else {
    $staging = Join-Path $env:TEMP ('M365AppsWatchdog_{0}' -f [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $staging -Force -WhatIf:$false | Out-Null
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    foreach ($name in $Pin.Files.Keys) {
        $url    = 'https://raw.githubusercontent.com/sjkanon/M365-Scripts/{0}/{1}' -f $Pin.Commit, $Pin.Files[$name].Path
        $target = Join-Path $staging $name
        try {
            Invoke-WebRequest -Uri $url -OutFile $target -UseBasicParsing
        } catch {
            Remove-Item $staging -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false
            Write-Error "$name could not be downloaded from $url - $($_.Exception.Message)"; exit 1
        }
        $hash = Get-Sha256 $target
        if ($hash -ne $Pin.Files[$name].Hash) {
            Remove-Item $staging -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false
            Write-Error "$name from GitHub has SHA-256 $hash, expected $($Pin.Files[$name].Hash) - nothing installed"; exit 1
        }
        $Sources[$name] = $target
        # The pinned hash, not the staged file's: the host checks against what was pinned.
        $Hashes[$name]  = $Pin.Files[$name].Hash
    }
    $Commit = '{0} from GitHub, SHA-256 verified' -f $Pin.Commit.Substring(0, 7)
}

# Settings given here, passed on to -Install over what config.json says.
$Override = @{}
foreach ($name in 'WebhookUrl', 'WebhookToken', 'Account', 'App', 'IntervalMinutes') {
    if ($PSBoundParameters.ContainsKey($name)) { $Override[$name] = $PSBoundParameters[$name] }
}

# -- On the host -----------------------------------------------------------------------
# Script blocks, so they run the same here and over remoting.
$LockBlock = {
    param([string] $WorkingDir)
    New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
    & icacls.exe $WorkingDir /inheritance:r /grant:r '*S-1-5-18:(OI)(CI)F' '*S-1-5-32-544:(OI)(CI)F' | Out-Null
    return $LASTEXITCODE -eq 0
}

$InstallBlock = {
    param([string] $WorkingDir, [hashtable] $Override, [hashtable] $Hashes)
    $ErrorActionPreference = 'Stop'
    $taskName = 'M365 App Watchdog'
    $result = [ordered]@{ Host = $env:COMPUTERNAME; Ok = $false; Settings = 'new'; Webhook = 'none'; Interval = 30; Detail = '' }
    try {
        foreach ($name in $Hashes.Keys) {
            $hash = (Get-FileHash -Path (Join-Path $WorkingDir $name) -Algorithm SHA256).Hash
            if ($hash -ne $Hashes[$name]) { throw "$name on the host does not match this checkout (SHA-256 $hash)" }
        }

        # Everything config.json holds, as -Install's parameters: -Install itself does not read it.
        $params     = @{}
        $configPath = Join-Path $WorkingDir 'config.json'
        if (Test-Path $configPath) {
            $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
            foreach ($name in 'Account', 'App', 'WebhookUrl', 'WebhookToken', 'RepairCooldownHours', 'RenotifyHours', 'UpdateHours', 'CrashThreshold',
                              'NoRepair', 'NoUserRepair', 'SkipLaunchTest', 'NoUserLaunch', 'NoDiagnostics') {
                $prop = $config.PSObject.Properties[$name]
                if (-not $prop -or $null -eq $prop.Value -or [string] $prop.Value -eq '') { continue }
                if ($prop.Value -is [bool]) { if ($prop.Value) { $params[$name] = $true }; continue }
                $params[$name] = $prop.Value
            }
            $result.Settings = 'kept'
        }
        $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($task) {
            $every = @($task.Triggers | ForEach-Object { $_.Repetition } | Where-Object { $_ -and $_.Interval } | ForEach-Object { $_.Interval }) | Select-Object -First 1
            if ($every) { $params['IntervalMinutes'] = [int] [Xml.XmlConvert]::ToTimeSpan($every).TotalMinutes }
        }
        foreach ($name in $Override.Keys) { $params[$name] = $Override[$name] }
        if ($Override.Count -gt 0 -and $result.Settings -eq 'kept') { $result.Settings = 'kept, ' + (@($Override.Keys | Sort-Object) -join ', ') + ' replaced' }
        if ($params.ContainsKey('WebhookUrl') -and $params['WebhookUrl']) { $result.Webhook = 'set' }
        if ($params.ContainsKey('IntervalMinutes')) { $result.Interval = $params['IntervalMinutes'] }

        $params['WorkingDir'] = $WorkingDir
        $params['Install']    = $true
        $params['Confirm']    = $false
        & (Join-Path $WorkingDir 'Watch-M365Apps.ps1') @params | Out-Host
        if ($LASTEXITCODE -ne 0) { throw "Watch-M365Apps.ps1 -Install exited with $LASTEXITCODE" }

        $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if (-not $task) { throw "task '$taskName' not there after -Install" }
        if ($task.State -eq 'Disabled') { throw "task '$taskName' is disabled" }
        $runs = [string] (@($task.Actions | ForEach-Object { $_.Arguments }) -join ' ')
        if ($runs -notlike "*$(Join-Path $WorkingDir 'Watch-M365Apps.ps1')*") { throw "task '$taskName' does not run the installed copy: $runs" }
        $result.Ok = $true
    } catch {
        $result.Detail = $_.Exception.Message
    }
    return [PSCustomObject] $result
}

$ThenBlock = {
    param([string] $WorkingDir, [string] $Then, [bool] $Restart)
    $params = @{}
    if ($Then -eq 'Check') { $params['CheckOnly'] = $true } else { $params['Confirm'] = $false; if ($Restart) { $params['Restart'] = $true } }
    & (Join-Path $WorkingDir 'Repair-M365AppsHost.ps1') @params | Out-Host
    return $LASTEXITCODE
}

# -- One run -------------------------------------------------------------------------
$targets = if ($ComputerName) { @($ComputerName) } else { @($env:COMPUTERNAME) }
Write-Host ''
Write-Host ("  M365 app watchdog - install from {0}{1} - {2}" -f $PSScriptRoot, $(if ($Commit) { " ($Commit)" } else { '' }), ($targets -join ', ')) -ForegroundColor Cyan
if ($Restart -and $Then -ne 'Repair') { Write-Warn '-Restart only counts with -Then Repair - ignored' }

if (-not $ComputerName) {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not ([Security.Principal.WindowsPrincipal] $identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        if ($staging) { Remove-Item $staging -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false }
        Write-Error 'Administrator rights are required. Run the script from an elevated session.'
        exit 1
    }
}
if (-not $PSCmdlet.ShouldProcess(($targets -join ', '), "Install the watchdog in $WorkingDir$(if ($Then -ne 'None') { ", then $Then" })")) {
    if ($staging) { Remove-Item $staging -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false }
    exit 0
}

$results = [System.Collections.Generic.List[object]]::new()
foreach ($target in $targets) {
    Write-Step "==== $target $('=' * [Math]::Max(0, 70 - $target.Length))"
    $session = $null
    $row     = [PSCustomObject]@{ Host = $target; Installed = 'no'; Settings = ''; Webhook = ''; Interval = ''; Then = ''; Detail = '' }
    try {
        if ($ComputerName) {
            $sessionArgs = @{ ComputerName = $target }
            if ($Credential) { $sessionArgs['Credential'] = $Credential }
            $session = New-PSSession @sessionArgs
            $run = { param($Block, $Arguments) Invoke-Command -Session $session -ScriptBlock $Block -ArgumentList $Arguments }
        } else {
            $run = { param($Block, $Arguments) & $Block @Arguments }
        }

        Write-Step '1. Lock'
        if (-not (& $run $LockBlock @(, $WorkingDir))) { throw "could not lock $WorkingDir" }
        Write-Ok "$WorkingDir - System and Administrators only"

        Write-Step '2. Copy'
        foreach ($entry in $Sources.GetEnumerator()) {
            $destination = Join-Path $WorkingDir $entry.Key
            if ($session) { Copy-Item -Path $entry.Value -Destination $destination -ToSession $session -Force }
            else { Copy-Item -Path $entry.Value -Destination $destination -Force }
            Write-Ok $entry.Key
        }

        Write-Step '3. Install'
        $installed = & $run $InstallBlock @($WorkingDir, $Override, $Hashes)
        $row.Settings = $installed.Settings
        $row.Webhook  = $installed.Webhook
        $row.Interval = $installed.Interval
        if (-not $installed.Ok) { throw $installed.Detail }
        $row.Installed = 'yes'
        Write-Step '4. Verify'
        Write-Ok ('Installed, files match this checkout, task enabled - settings {0}, webhook {1}, every {2} min' -f $installed.Settings, $installed.Webhook, $installed.Interval)
        if ($installed.Webhook -eq 'none') { Write-Warn 'No webhook: the watchdog repairs and logs, but reports nothing - give -WebhookUrl' }

        if ($Then -ne 'None') {
            Write-Step "5. Then - $Then"
            $code = & $run $ThenBlock @($WorkingDir, $Then, [bool] $Restart)
            $row.Then = switch ($code) { 0 { 'healthy' } 2 { 'healthy, restart due' } default { "exit $code" } }
        }
    } catch {
        $row.Detail = $_.Exception.Message
        Write-Bad "$target - $($_.Exception.Message)"
    } finally {
        if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
    }
    $results.Add($row)
}
if ($staging) { Remove-Item $staging -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false }

Write-Step 'Summary'
$results | Format-Table Host, Installed, Settings, Webhook, Interval, Then, Detail -AutoSize -Wrap | Out-String -Width 200 | Write-Host

if (@($results | Where-Object { $_.Installed -ne 'yes' -or ($_.Then -and $_.Then -notlike 'healthy*') }).Count -gt 0) { exit 1 }
if (@($results | Where-Object { $_.Then -eq 'healthy, restart due' }).Count -gt 0) { exit 2 }
exit 0
