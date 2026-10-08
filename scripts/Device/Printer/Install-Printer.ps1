#Requires -Version 5.1
<#
.SYNOPSIS
    Install printer drivers and printers as described in a JSON file, with the
    drivers downloaded from a GitHub repository. Supports -WhatIf and -CheckOnly.

.DESCRIPTION
    One JSON file says which drivers exist, where each one is downloaded from and
    which printers use them. This script makes the device match it:

      1. Config     - read the JSON, from a local path, a UNC path or an https URL
                      (a GitHub raw link to the same repository works), validate
                      every field of it, and pick the printers asked for with -Printer.
      2. Preflight  - wait for the Print Spooler, then for every driver those
                      printers need: is it installed for this architecture, and in
                      which version. For every printer: does it exist, on which port,
                      with which driver and settings.
      3. Download   - only for a driver that is missing or older than the JSON asks
                      for. Sources: a GitHub release asset, a file or folder in a
                      GitHub repository, any https URL, or a local/UNC path. An
                      optional sha256 is checked, a .zip or .cab is extracted.
      4. Driver     - find the INF (named in the JSON, or the one that declares the
                      driver name), check that it is a printer INF for this
                      architecture that declares that exact name, check the catalog
                      signature, add it to the driver store with pnputil and install
                      it with Add-PrinterDriver.
      5. Printer    - create the TCP/IP port, add the printer or correct its driver
                      and port, then apply location, comment, sharing and the
                      print defaults (duplex, colour, paper size) from the JSON.
      6. Verify     - read everything back.

    Nothing is downloaded for a driver that is already there in the right version,
    and nothing is changed for a printer that already matches. A scheduled run on a
    device that is in order costs one read of the JSON.

    The JSON
    --------
        {
          "drivers": [
            {
              "name":    "HP Universal Printing PCL 6",
              "version": "7.2.0.25780",
              "inf":     "hpcu270u.inf",
              "source":  { "type": "githubRelease", "repository": "contoso/printer-drivers",
                           "tag": "latest", "asset": "hp-upd-pcl6-x64-*.zip" }
            }
          ],
          "printers": [
            { "name": "Office 1st floor", "driver": "HP Universal Printing PCL 6",
              "address": "10.0.5.20", "location": "1st floor", "duplex": "TwoSidedLongEdge" }
          ]
        }

    The driver "name" is the exact name inside the INF - the one Get-PrinterDriver
    shows after installing it by hand. That name is also what a printer's "driver"
    refers to. When it does not match, the run names the models the INF does
    declare. printers.example.json beside this script shows every field.

    Printer fields beyond name/driver/address: portName (default IP_<address>),
    portNumber (default 9100), lprQueue (LPR instead of RAW), snmp (true turns SNMP
    status on - off by default, so a printer that does not answer SNMP is not shown
    as Offline), location, comment, shared, shareName, duplex, color, paperSize,
    and "ensure": "absent" to remove a printer (its port too when nothing else uses
    it; the driver always stays).

    After a golden image
    --------------------
    Made to run unattended on a server or session host that was just provisioned
    from an image: as System, from a Custom Script Extension, a first-boot or
    startup task, Intune or an RMM. On that first boot Windows is still settling,
    so a Print Spooler that has not started yet is started and waited for, and
    downloads that fail on DNS, TLS or a timeout are retried - both for up to
    -WaitSeconds. A 401/403/404 is an answer rather than a hiccup and fails at once.
    The run is idempotent, so the same command can be scheduled at every startup:
    a host that is in order costs one read of the JSON. Printers installed this way
    are machine-wide, so on a session host every user sees them.

    What keeps a run from going wrong
    ---------------------------------
      - The whole JSON is validated before anything is downloaded: required and
        unknown fields (a typo such as "adress"), duplicate names, printer names
        Windows refuses, addresses, port numbers, enum values, two printers on one
        port with different addresses. A broken file changes nothing.
      - One run at a time: a startup task and an RMM job that overlap queue on a
        machine-wide lock instead of installing the same driver twice.
      - A JSON fetched from a URL is cached. When the URL cannot be reached, the
        last good copy is used, with a warning, rather than leaving a host without
        printers because GitHub was briefly unreachable.
      - Downloads are checked for what they are: an HTML page from a proxy or a
        login portal is refused instead of being "extracted", and there has to be
        room on the disk before anything is fetched.
      - The INF is checked before pnputil sees it: printer class, a section for
        this architecture, the exact driver name among its models, a signed catalog.
      - A spooler that stops or stalls during an install - common right after a
        driver lands - is restarted and the step tried once more.
      - Get-/Set-PrintConfiguration run the vendor's driver code, which hangs on
        some universal drivers, so both run in a job with a timeout.
      - A driver that fails takes only its own printers down; the rest still go in,
        and the exit code reports the failure.
      - Every real run is logged to one appended, rotated file, so an unattended
        first boot can be read back afterwards.
      - A clean run records the SHA256 of the JSON it applied under
        HKLM:\SOFTWARE\M365-Scripts\InstallPrinter (ConfigSha256, LastSuccess,
        Printers), for an Intune detection rule or RMM condition.

    Driver sources
    --------------
      githubRelease  repository, tag (default: latest), asset (name, wildcards allowed)
      github         repository, path (a .zip or a folder holding the INF), ref
                     (branch/tag/commit, default: the repository's default branch)
      url            url (https), downloads one file - .zip, .cab or a lone file
      path           path to a folder or .zip; relative paths resolve against the
                     folder of the JSON file

    A private repository needs a token with read access to its contents:
    -GitHubToken, or the environment variable GITHUB_TOKEN. The token is only ever
    sent to api.github.com and GitHub's own download hosts, never to a "url"
    source. It is deliberately not something the JSON can carry.

    RMM / NinjaOne
    --------------
    Script variables arrive as environment variables and are used when the matching
    parameter is not passed: configPath, printer (comma separated), githubToken,
    workingDir, logPath, waitSeconds, proxy, and the checkboxes checkOnly, whatIf,
    quiet, force and skipSignatureCheck. Started 32-bit, the script relaunches
    itself 64-bit - pnputil does not exist under SysWOW64. Started by hand without
    elevation, it asks for it.

    Exit codes
    ----------
        0  success, or everything already in order
        1  failure
        2  -CheckOnly only: work is due

.PARAMETER ConfigPath
    The JSON file: a local or UNC path, or an https URL. Mandatory unless the RMM
    variable configPath supplies it.

.PARAMETER Printer
    Only these printers from the JSON (names, wildcards allowed). Default: all of
    them. Drivers that none of the selected printers use are left alone, unless
    the JSON marks them "install": true.

.PARAMETER GitHubToken
    Token for a private GitHub repository (falls back to $env:GITHUB_TOKEN).

.PARAMETER Proxy
    Proxy for every download, e.g. http://proxy.contoso.local:8080. Running as
    System there is no user proxy setting to inherit, so a network that only lets
    traffic out through a proxy needs it here. Uses the machine's own credentials.

.PARAMETER WorkingDir
    Folder for downloads, extracted drivers and the cached JSON
    (default: C:\IT\Printers). A driver's files are removed again once it is
    installed; the driver store keeps its own copy.

.PARAMETER LogPath
    Folder for Install-Printer.log (default: C:\Temp), appended by every run that
    may change something and rotated past 1 MB.

.PARAMETER CheckOnly
    Report what would be done and change nothing. Exit code 2 when work is due.

.PARAMETER Quiet
    Print nothing unless there is work to do or something failed. The log file
    gets the full story either way.

.PARAMETER Force
    Reinstall the drivers even when the installed version already matches.

.PARAMETER SkipSignatureCheck
    Do not require a valid signature on the driver's catalog file. Windows itself
    still refuses an unsigned package driver; this only skips the script's own check.

.PARAMETER WaitSeconds
    How long a freshly provisioned machine is given to get ready: the Print Spooler
    to start, the network to answer downloads, and another run of this script to
    finish (default: 300, 0 = no waiting).

.EXAMPLE
    .\Install-Printer.ps1 -ConfigPath .\printers.json -CheckOnly

    What would be installed, nothing changed. Exit code 2 when work is due.

.EXAMPLE
    .\Install-Printer.ps1 -ConfigPath https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json -Printer 'Office*' -WhatIf

    The JSON from GitHub, only the printers whose name starts with Office, as a dry run.

.EXAMPLE
    .\Install-Printer.ps1 -ConfigPath \\fs01\it$\printers.json -Quiet -Confirm:$false

    NinjaOne/Intune: silent unless something was installed or failed.

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File C:\IT\Install-Printer.ps1 -ConfigPath https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json -Quiet -Confirm:$false

    First boot of a server built from a golden image (Custom Script Extension or a
    startup task as System): waits for the spooler and the network, then installs.

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows 10 / Server 2016 or newer (PrintManagement + pnputil),
              run as administrator or as System
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [string]   $ConfigPath,
    [string[]] $Printer,
    [string]   $GitHubToken,
    [string]   $Proxy,
    [string]   $WorkingDir = 'C:\IT\Printers',
    [string]   $LogPath    = 'C:\Temp',
    [switch]   $CheckOnly,
    [switch]   $Quiet,
    [switch]   $Force,
    [switch]   $SkipSignatureCheck,
    [ValidateRange(0, 3600)]
    [int]      $WaitSeconds = 300
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-ForwardedArgument {
    <# Rebuild the caller's own parameters as a command line for a relaunch. #>
    param([Parameter(Mandatory)] $Bound)

    $list = @()
    foreach ($entry in $Bound.GetEnumerator()) {
        if ($entry.Value -is [switch] -or $entry.Value -is [bool]) {
            if ($entry.Value) { $list += "-$($entry.Key)" } else { $list += "-$($entry.Key):`$false" }
        } elseif ($entry.Value -is [array]) {
            $list += "-$($entry.Key)"
            $list += ($entry.Value -join ',')
        } else {
            $list += "-$($entry.Key)"
            $list += [string] $entry.Value
        }
    }
    return $list
}

# -- RMM: run 64-bit -----------------------------------------------------------
# pnputil only exists in System32. A 32-bit PowerShell is redirected to SysWOW64,
# where it is not, and the 64-bit driver environment is not the one it sees.
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $nativeShell = Join-Path $env:WINDIR 'SysNative\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path $nativeShell) {
        $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath) +
                   (Get-ForwardedArgument -Bound $PSBoundParameters)
        & $nativeShell @argList
        exit $LASTEXITCODE
    }
    Write-Error 'Running 32-bit and SysNative is unavailable - pnputil cannot be reached. Start the script from 64-bit PowerShell.'
    exit 1
}

# -- Elevation -----------------------------------------------------------------
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not ([Security.Principal.WindowsPrincipal] $identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
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
$rmmTrue = @('true', '1', 'yes')
if (-not $PSBoundParameters.ContainsKey('WhatIf')             -and $env:whatIf             -in $rmmTrue) { $WhatIfPreference   = $true }
if (-not $PSBoundParameters.ContainsKey('CheckOnly')          -and $env:checkOnly          -in $rmmTrue) { $CheckOnly          = $true }
if (-not $PSBoundParameters.ContainsKey('Quiet')              -and $env:quiet              -in $rmmTrue) { $Quiet              = $true }
if (-not $PSBoundParameters.ContainsKey('Force')              -and $env:force              -in $rmmTrue) { $Force              = $true }
if (-not $PSBoundParameters.ContainsKey('SkipSignatureCheck') -and $env:skipSignatureCheck -in $rmmTrue) { $SkipSignatureCheck = $true }
if (-not $PSBoundParameters.ContainsKey('ConfigPath')         -and $env:configPath)  { $ConfigPath  = $env:configPath }
if (-not $PSBoundParameters.ContainsKey('Printer')            -and $env:printer)     { $Printer     = @($env:printer -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
if (-not $PSBoundParameters.ContainsKey('WorkingDir')         -and $env:workingDir)  { $WorkingDir  = $env:workingDir }
if (-not $PSBoundParameters.ContainsKey('LogPath')            -and $env:logPath)     { $LogPath     = $env:logPath }
if (-not $PSBoundParameters.ContainsKey('Proxy')              -and $env:proxy)       { $Proxy       = $env:proxy }
if (-not $PSBoundParameters.ContainsKey('WaitSeconds')        -and $env:waitSeconds -match '^\d+$') { $WaitSeconds = [math]::Min([int] $env:waitSeconds, 3600) }
if (-not $GitHubToken) { $GitHubToken = if ($env:githubToken) { $env:githubToken } else { $env:GITHUB_TOKEN } }

if (-not $ConfigPath) {
    Write-Error 'No configuration given. Pass -ConfigPath (a JSON file or an https URL), or set the configPath script variable.'
    exit 1
}
if ($Proxy -and $Proxy -notmatch '^https?://[^\s/]+') {
    Write-Error "Proxy must look like http://host:port, got '$Proxy'"
    exit 1
}

$simulate          = [bool] $WhatIfPreference
$confirmSuppressed = $PSBoundParameters.ContainsKey('Confirm') -and -not $PSBoundParameters['Confirm']
$exitCode          = 0
$plannedExit       = $null
$rebootRequired    = $false
$mutex             = $null
$mutexOwned        = $false

# The hosts a GitHub token may be sent to. A "url" source pointing anywhere else
# is downloaded without it - a token in a JSON-chosen request is a token leaked.
$GitHubHosts = @('api.github.com', 'github.com', 'raw.githubusercontent.com', 'codeload.github.com', 'objects.githubusercontent.com')

# The driver environment and INF decoration this machine needs. An x64 host
# cannot use an x86-only package, and Get-PrinterDriver lists the x86 copy of a
# driver (installed for 32-bit clients) under the same name.
$NativeArch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') {
    [PSCustomObject]@{ Environment = 'Windows ARM64'; Decoration = 'NTarm64'; Folder = 'arm64' }
} else {
    [PSCustomObject]@{ Environment = 'Windows x64';   Decoration = 'NTamd64'; Folder = 'x64|amd64|win64|64bit|64-bit' }
}

# Every field the JSON may carry. Anything else is a typo until proven otherwise,
# and a typo in an optional field ("loaction") would otherwise be silently ignored.
$KnownRootFields    = @('drivers', 'printers', '$schema', 'description')
$KnownDriverFields  = @('name', 'version', 'inf', 'install', 'source', 'description')
$KnownSourceFields  = @{
    githubRelease = @('type', 'repository', 'tag', 'asset', 'sha256')
    github        = @('type', 'repository', 'ref', 'path', 'sha256')
    url           = @('type', 'url', 'sha256')
    path          = @('type', 'path', 'sha256')
}
$KnownPrinterFields = @('name', 'driver', 'address', 'portName', 'portNumber', 'lprQueue', 'snmp', 'location',
                        'comment', 'shared', 'shareName', 'duplex', 'color', 'paperSize', 'ensure', 'description')
$DuplexValues       = @('OneSided', 'TwoSidedLongEdge', 'TwoSidedShortEdge')

# Written after every clean run, for detection rules - see Set-SuccessMarker.
$MarkerPath = 'HKLM:\SOFTWARE\M365-Scripts\InstallPrinter'

# -- Output --------------------------------------------------------------------
# -Quiet holds every line back until something worth reporting happens. The log
# file is written either way: a first boot nobody watched is worth reading back.
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
    <# One appended log, rotated at 1 MB, so a host that runs this at every boot
       keeps a readable history instead of an ever-growing file. #>
    if ($simulate -or $CheckOnly) { return }
    try {
        if (-not (Test-Path -LiteralPath $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }
        $file = Join-Path $LogPath 'Install-Printer.log'
        if ((Test-Path -LiteralPath $file) -and (Get-Item -LiteralPath $file).Length -gt 1MB) {
            Move-Item -LiteralPath $file -Destination "$file.old" -Force
        }
        $script:logFile = $file
        Write-Log "--- Install-Printer started (user $env:USERNAME, config $ConfigPath$(if ($Printer) { ", printers $($Printer -join ', ')" })) ---"
    } catch {
        Write-Warning "Could not open the log in ${LogPath}: $($_.Exception.Message)"
    }
}

function Get-PropertyValue {
    <# Property access that returns $null instead of tripping Set-StrictMode. #>
    param($Object, [string] $Name)

    if ($null -eq $Object) { return $null }
    if ($Object.PSObject.Properties.Name -notcontains $Name) { return $null }
    return $Object.$Name
}

function Test-JsonObject {
    <#
        Whether a value is a JSON object. Not "-is [PSCustomObject]": that type
        accelerator stands for PSObject, which every wrapped value is, so a list
        or a string passes it too.
    #>
    param($Value)
    return ($null -ne $Value -and $Value.GetType().FullName -eq 'System.Management.Automation.PSCustomObject')
}

function Get-FieldName {
    <# The names a JSON object carries, for the unknown-field check. #>
    param($Object)
    if ($null -eq $Object -or -not (Test-JsonObject $Object)) { return @() }
    return @($Object.PSObject.Properties | ForEach-Object { $_.Name })
}

# -- Readiness -----------------------------------------------------------------
# On the first boot of a server provisioned from a golden image this runs while
# Windows is still coming up: DNS that does not resolve yet, a clock that has not
# synced (so TLS fails), a spooler that is set to Automatic but not started. All
# of that settles within minutes, so it is waited for (up to -WaitSeconds) rather
# than reported as a failure that leaves a brand-new host without its printers.

function Invoke-WithRetry {
    <#
        Retry a network call while the error looks transient. A 400/401/403/404 is
        an answer, not a hiccup - retrying it only delays the message that says so.
        A 403 that is a rate limit is the exception, but waiting out GitHub's hour
        is longer than any first boot should, so it fails with that explanation.
    #>
    param([Parameter(Mandatory)] [scriptblock] $Action, [Parameter(Mandatory)] [string] $What)

    $deadline = (Get-Date).AddSeconds($WaitSeconds)
    $delay    = 10
    while ($true) {
        try { return (& $Action) } catch {
            $message = "$($_.Exception.Message)"
            if ($message -match '\b(400|401|403|404|410|422)\b' -or (Get-Date).AddSeconds($delay) -gt $deadline) { throw }
            Write-Warn "$What not reachable yet ($message) - retrying in $delay seconds"
            Start-Sleep -Seconds $delay
            $delay = [math]::Min($delay * 2, 60)
        }
    }
}

function Wait-PrintSpooler {
    <# Start the spooler when it is not running, and give it -WaitSeconds to come up. #>
    $deadline = (Get-Date).AddSeconds([math]::Max($WaitSeconds, 30))
    $service  = Get-Service -Name Spooler -ErrorAction SilentlyContinue
    if (-not $service) { throw 'The Print Spooler service does not exist on this machine (Server Core without the print feature?).' }
    if ("$($service.StartType)" -eq 'Disabled') {
        throw 'The Print Spooler service is disabled - a hardened image often does this (PrintNightmare). Set it to Automatic before installing printers; this script does not undo a hardening decision by itself.'
    }
    $warned = $false
    while ($true) {
        $service.Refresh()
        if ($service.Status -eq 'Running') {
            # Running is not the same as answering: right after a start the RPC
            # endpoint takes a moment, and the first cmdlet then fails with 0x800706BA.
            try { $null = @(Get-PrinterDriver -ErrorAction Stop); return } catch {
                if ((Get-Date) -gt $deadline) { throw "The Print Spooler is running but does not answer: $($_.Exception.Message)" }
            }
        } elseif ($service.Status -eq 'Stopped') {
            try { Start-Service -Name Spooler -ErrorAction Stop } catch { Write-Warn "Could not start the Print Spooler yet: $($_.Exception.Message)" }
        }
        if ((Get-Date) -gt $deadline) { throw "The Print Spooler service is still $($service.Status) after $WaitSeconds seconds." }
        if (-not $warned) { Write-Warn "Waiting for the Print Spooler service ($($service.Status))"; $warned = $true }
        Start-Sleep -Seconds 3
    }
}

function Invoke-SpoolerAction {
    <#
        One spooler call, retried once after a spooler restart when the spooler is
        what failed: it died (a vendor driver crashing it right after install is a
        classic), or its RPC endpoint stopped answering. Any other error - a wrong
        name, a port in use - is real and goes straight back to the caller.
    #>
    param([Parameter(Mandatory)] [scriptblock] $Action, [Parameter(Mandatory)] [string] $What)

    try { return (& $Action) } catch {
        $message = "$($_.Exception.Message)"
        $stopped = (Get-Service -Name Spooler -ErrorAction SilentlyContinue).Status -ne 'Running'
        $rpc     = $message -match '0x800706BA|0x800706BE|0x800706BF|RPC server|remote procedure call'
        if (-not ($stopped -or $rpc)) { throw }

        Write-Warn "$What failed ($message) - the Print Spooler $(if ($stopped) { 'stopped' } else { 'stopped answering' }); restarting it and trying once more"
        try { Restart-Service -Name Spooler -Force -ErrorAction Stop } catch { Write-Warn "Restarting the Print Spooler: $($_.Exception.Message)" }
        Wait-PrintSpooler
        return (& $Action)
    }
}

function Invoke-WithTimeout {
    <#
        Run a scriptblock in a job and give up after -Seconds. For calls into vendor
        driver code (print configuration), which hangs on some universal drivers -
        a hang there must cost one setting, not the whole provisioning run.
    #>
    param([Parameter(Mandatory)] [scriptblock] $ScriptBlock, [object[]] $ArgumentList = @(), [int] $Seconds = 120)

    $job = Start-Job -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList
    try {
        if (-not (Wait-Job -Job $job -Timeout $Seconds)) {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
            throw "did not finish within $Seconds seconds - the printer driver hangs on it"
        }
        return (Receive-Job -Job $job -ErrorAction Stop)
    } finally {
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }
}

function Enter-SingleRun {
    <#
        One run at a time per machine. A startup task and an RMM job that fire
        together would otherwise both see the same driver missing and both run
        pnputil on it. The second one waits its turn, up to -WaitSeconds.
    #>
    $script:mutex = [Threading.Mutex]::new($false, 'Global\M365Scripts-Install-Printer')
    try {
        $script:mutexOwned = $script:mutex.WaitOne(0)
        if (-not $script:mutexOwned) {
            Write-Warn 'Another Install-Printer run is busy on this machine - waiting for it to finish'
            $script:mutexOwned = $script:mutex.WaitOne([math]::Max($WaitSeconds, 1) * 1000)
        }
    } catch [Threading.AbandonedMutexException] {
        # The previous run was killed mid-way. The lock is ours now; the state it
        # left behind is exactly what the preflight is about to read.
        $script:mutexOwned = $true
        Write-Warn 'A previous run ended without cleaning up (killed?) - continuing; the preflight below shows what it left'
    }
    if (-not $script:mutexOwned) { throw "Another Install-Printer run is still busy after $WaitSeconds seconds." }
}

# -- Downloads -----------------------------------------------------------------

function Get-RequestParameter {
    <# Headers and proxy for one request; the token only travels to GitHub itself. #>
    param([Parameter(Mandatory)] [string] $Uri, [string] $Accept)

    $headers = @{ 'User-Agent' = 'M365-Scripts-Install-Printer' }
    if ($Accept) { $headers['Accept'] = $Accept }
    if ($GitHubToken -and (([uri] $Uri).Host -in $GitHubHosts)) {
        $headers['Authorization'] = "Bearer $GitHubToken"
    }
    $request = @{ Uri = $Uri; Headers = $headers; UseBasicParsing = $true; TimeoutSec = 300 }
    if ($Proxy) { $request['Proxy'] = $Proxy; $request['ProxyUseDefaultCredentials'] = $true }
    return $request
}

function Enable-Tls12 {
    # Added to what is enabled, not replacing it: replacing would switch TLS 1.3
    # off on a machine that has it.
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
}

function Invoke-Download {
    <# One file to disk over https, with the right headers and without a progress bar. #>
    param([Parameter(Mandatory)] [string] $Uri, [Parameter(Mandatory)] [string] $Path, [string] $Accept)

    if ($Uri -notmatch '^https://') { throw "Download URL must be https: $Uri" }
    Enable-Tls12
    $progressBackup     = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    try {
        $parent = Split-Path $Path -Parent
        if (-not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        $request = Get-RequestParameter -Uri $Uri -Accept $Accept
        Invoke-WithRetry -What $Uri -Action {
            # A partial file from a failed attempt must not be mistaken for a download.
            Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
            Invoke-WebRequest @request -OutFile $Path
        }
        if (-not (Test-Path -LiteralPath $Path) -or (Get-Item -LiteralPath $Path).Length -eq 0) {
            throw 'the server answered with an empty file'
        }
    } catch {
        # A private repository answers 404, not 403, to a request without a token -
        # which reads like a typo in the path when it is really a missing token.
        $message = "$($_.Exception.Message)"
        $hint = if (-not $GitHubToken -and (([uri] $Uri).Host -in $GitHubHosts) -and $message -match '404') {
            ' - for a private repository, pass -GitHubToken or set GITHUB_TOKEN'
        } elseif (([uri] $Uri).Host -eq 'api.github.com' -and $message -match '403|429') {
            ' - GitHub rate limited this IP (60 API requests an hour without a token); a token raises it to 5000'
        } elseif ($message -match '407') {
            ' - the proxy wants authentication; check -Proxy and that the computer account may use it'
        } else { '' }
        throw "Download failed ($Uri): $message$hint"
    } finally {
        $ProgressPreference = $progressBackup
    }
}

function Invoke-GitHubApi {
    param([Parameter(Mandatory)] [string] $Path)

    $uri = "https://api.github.com/$($Path.TrimStart('/'))"
    Enable-Tls12
    $request = Get-RequestParameter -Uri $uri -Accept 'application/vnd.github+json'
    try {
        return Invoke-WithRetry -What $uri -Action { Invoke-RestMethod @request }
    } catch {
        $message = "$($_.Exception.Message)"
        $hint = if (-not $GitHubToken -and $message -match '404') {
            ' - check the name, or pass -GitHubToken for a private repository'
        } elseif ($message -match '403|429') {
            ' - rate limited or no access. Without a token GitHub allows 60 API requests an hour per public IP, which a pool of hosts behind one NAT uses up quickly; a token raises it to 5000'
        } else { '' }
        throw "GitHub API $uri failed: $message$hint"
    }
}

function Assert-FreeSpace {
    <# Room for a driver package before it is fetched, rather than a half-written zip. #>
    param([Parameter(Mandatory)] [string] $Folder, [long] $Bytes = 1GB)
    try {
        $root = [IO.Path]::GetPathRoot([IO.Path]::GetFullPath($Folder))
        if ($root -notmatch '^[A-Za-z]:\\$') { return }   # UNC: the share reports its own space
        $free = ([IO.DriveInfo]::new($root)).AvailableFreeSpace
    } catch { return }
    if ($free -lt $Bytes) {
        throw ("Only {0:N0} MB free on {1} - at least {2:N0} MB is needed to download and extract a driver package" -f ($free / 1MB), $root, ($Bytes / 1MB))
    }
}

function Test-DownloadedFile {
    <#
        A download that is not what it claims to be: a proxy block page, a login
        portal or a GitHub error page saved under the name of a zip. Caught here,
        it is a clear message; caught by Expand-Archive, it is "not a zip file".
    #>
    param([Parameter(Mandatory)] [string] $Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return }
    $stream = [IO.File]::OpenRead($Path)
    try {
        $buffer = New-Object byte[] 512
        $read   = $stream.Read($buffer, 0, $buffer.Length)
    } finally { $stream.Dispose() }
    $head = [Text.Encoding]::ASCII.GetString($buffer, 0, $read).TrimStart([char] 0xFEFF, ' ', "`r", "`n", "`t")

    if ($head -match '^(?i)(<!doctype html|<html|<\?xml|\{\s*"message")') {
        throw "$(Split-Path $Path -Leaf) is a web page or an error message, not a driver package - a proxy, a login portal or a wrong URL answered instead of the file"
    }
    $extension = [IO.Path]::GetExtension($Path).ToLower()
    if ($extension -eq '.zip' -and -not $head.StartsWith('PK')) { throw "$(Split-Path $Path -Leaf) is not a zip file (wrong URL, or a damaged download)" }
    if ($extension -eq '.cab' -and -not $head.StartsWith('MSCF')) { throw "$(Split-Path $Path -Leaf) is not a cabinet file (wrong URL, or a damaged download)" }
}

function Test-RepositoryName {
    param([string] $Repository)
    return ($Repository -match '^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$')
}

function Save-GitHubReleaseAsset {
    <#
        A release asset by name. Downloaded through the API's asset URL rather than
        browser_download_url: only the API one honours a token, so the same code
        works for a public and a private repository.
    #>
    param([Parameter(Mandatory)] $Source, [Parameter(Mandatory)] [string] $Destination)

    $repo = Get-PropertyValue $Source 'repository'
    $tag  = Get-PropertyValue $Source 'tag'
    $want = Get-PropertyValue $Source 'asset'

    $release = if (-not $tag -or $tag -eq 'latest') { Invoke-GitHubApi "repos/$repo/releases/latest" }
               else { Invoke-GitHubApi "repos/$repo/releases/tags/$([uri]::EscapeDataString($tag))" }

    $assets = @(@($release.assets) | Where-Object { $_.name -like $want })
    if ($assets.Count -eq 0) {
        throw "Release $($release.tag_name) of $repo has no asset matching '$want' (it has: $((@($release.assets) | ForEach-Object { $_.name }) -join ', '))"
    }
    if ($assets.Count -gt 1) {
        throw "Release $($release.tag_name) of $repo has $($assets.Count) assets matching '$want' ($(($assets | ForEach-Object { $_.name }) -join ', ')) - make the pattern specific"
    }

    # With a token the API's asset URL, the only one that honours it (a private
    # repository). Without one the plain download link: it does not count toward
    # the 60 API requests an hour GitHub allows per public IP, which a pool of
    # hosts booting behind one NAT would otherwise use up.
    $file = Join-Path $Destination $assets[0].name
    if ($GitHubToken) { Invoke-Download -Uri $assets[0].url -Path $file -Accept 'application/octet-stream' }
    else              { Invoke-Download -Uri $assets[0].browser_download_url -Path $file }
    Write-Ok ("Downloaded {0} ({1:N1} MB) from release {2} of {3}" -f $assets[0].name, ((Get-Item -LiteralPath $file).Length / 1MB), $release.tag_name, $repo)
    return $file
}

function Save-GitHubPath {
    <#
        A file or a whole folder out of a repository. A folder is what a driver
        usually is - INF, CAT and a handful of DLLs side by side - so the tree is
        listed once and every file under the path is fetched through the contents
        API, which serves raw bytes to a token as well as without one.
    #>
    param([Parameter(Mandatory)] $Source, [Parameter(Mandatory)] [string] $Destination)

    $repo = Get-PropertyValue $Source 'repository'
    $path = ([string] (Get-PropertyValue $Source 'path')).Trim('/').Replace('\', '/')

    $ref = Get-PropertyValue $Source 'ref'
    if (-not $ref) { $ref = (Invoke-GitHubApi "repos/$repo").default_branch }

    $tree = Invoke-GitHubApi "repos/$repo/git/trees/$([uri]::EscapeDataString($ref))?recursive=1"
    if (Get-PropertyValue $tree 'truncated') {
        # Silently installing half a driver is worse than not installing it.
        throw "The tree of $repo is too large for one GitHub listing, so files under '$path' could be missing - publish the driver as a release asset or a .zip instead"
    }
    $files = @(@($tree.tree) | Where-Object { $_.type -eq 'blob' -and ($_.path -eq $path -or $_.path -like "$path/*") })
    if ($files.Count -eq 0) { throw "Nothing at '$path' in $repo ($ref)" }

    # Git LFS keeps a pointer in the tree, not the file; the contents API then
    # returns the 130-byte pointer, which would install as a corrupt driver.
    $prefix = if ($files.Count -eq 1 -and $files[0].path -eq $path) { ($path -replace '[^/]+$', '') } else { "$path/" }
    foreach ($item in $files) {
        $relative = $item.path.Substring($prefix.Length)
        $target   = Join-Path $Destination ($relative -replace '/', '\')
        $escaped  = ($item.path -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
        # Same reasoning as for release assets: the contents API for a token,
        # raw.githubusercontent.com - outside the API rate limit - without one.
        # A folder costs one request per file, so this is what keeps a driver of
        # forty files from exhausting the limit on its own.
        if ($GitHubToken) {
            Invoke-Download -Uri "https://api.github.com/repos/$repo/contents/$escaped`?ref=$([uri]::EscapeDataString($ref))" `
                            -Path $target -Accept 'application/vnd.github.raw'
        } else {
            $refPath = ($ref -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
            Invoke-Download -Uri "https://raw.githubusercontent.com/$repo/$refPath/$escaped" -Path $target
        }
        if ((Get-Item -LiteralPath $target).Length -lt 200 -and
            (Get-Content -LiteralPath $target -Raw -ErrorAction SilentlyContinue) -match '^version https://git-lfs') {
            throw "$($item.path) is a Git LFS pointer, not the file - publish the driver as a release asset instead of through LFS"
        }
    }
    Write-Ok "Downloaded $($files.Count) file(s) from $repo/$path ($ref)"

    if ($files.Count -eq 1) { return (Join-Path $Destination ($files[0].path.Substring($prefix.Length) -replace '/', '\')) }
    return $Destination
}

function Test-FileHash {
    param([Parameter(Mandatory)] [string] $Path, [string] $Expected)

    if (-not $Expected) { return }
    if (Test-Path $Path -PathType Container) {
        Write-Warn 'sha256 is set but the source is a folder - a hash only applies to a single file, so it was not checked'
        return
    }
    $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    if ($actual -ne $Expected.ToUpper()) {
        throw "SHA256 of $(Split-Path $Path -Leaf) is $actual, the JSON expects $($Expected.ToUpper()) - refusing to install it"
    }
    Write-Ok "SHA256 matches ($actual)"
}

function Expand-Package {
    <# A zip or cab into a folder. Expand-Archive first, the .NET extractor when the
       module is unavailable or chokes, which happens on stripped-down images. #>
    param([Parameter(Mandatory)] [string] $File, [Parameter(Mandatory)] [string] $Destination)

    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    if ([IO.Path]::GetExtension($File) -eq '.cab') {
        $output = & "$env:WINDIR\System32\expand.exe" -F:* $File $Destination 2>&1
        if ($LASTEXITCODE -ne 0) { throw "expand.exe could not extract $(Split-Path $File -Leaf) (exit code $LASTEXITCODE): $($output -join ' ')" }
        return
    }
    try {
        Expand-Archive -LiteralPath $File -DestinationPath $Destination -Force
    } catch {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        Remove-Item -LiteralPath $Destination -Recurse -Force -ErrorAction SilentlyContinue
        [IO.Compression.ZipFile]::ExtractToDirectory($File, $Destination)
    }
}

function New-DriverFolder {
    <#
        A clean folder for one driver. When last run's folder cannot be removed -
        a file still locked by an antivirus scan, say - a fresh one beside it is
        used instead of mixing two versions' files in one place.
    #>
    param([Parameter(Mandatory)] [string] $Name)

    $safe   = ($Name -replace '[^\w.-]+', '_').Trim('_')
    if ($safe.Length -gt 60) { $safe = $safe.Substring(0, 60) }
    $folder = Join-Path $WorkingDir $safe
    if (Test-Path -LiteralPath $folder) {
        try { Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction Stop }
        catch { $folder = '{0}_{1:yyyyMMddHHmmss}' -f $folder, (Get-Date) }
    }
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    return $folder
}

function Save-DriverSource {
    <#
        Bring one driver's files onto this machine and return the folder that
        holds them, extracted.
    #>
    param([Parameter(Mandatory)] $Driver, [string] $ConfigFolder, [Parameter(Mandatory)] [string] $Folder)

    $source   = $Driver.source
    $type     = [string] $source.type
    $download = Join-Path $Folder 'download'
    New-Item -ItemType Directory -Path $download -Force | Out-Null

    $fetched = switch ($type) {
        'githubRelease' { Save-GitHubReleaseAsset -Source $source -Destination $download }
        'github'        { Save-GitHubPath -Source $source -Destination $download }
        'url' {
            $url  = [string] $source.url
            $name = [uri]::UnescapeDataString((Split-Path ([uri] $url).AbsolutePath -Leaf))
            if (-not $name -or $name.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) { $name = 'driver.zip' }
            $file = Join-Path $download $name
            Invoke-Download -Uri $url -Path $file
            Write-Ok ("Downloaded {0} ({1:N1} MB)" -f $name, ((Get-Item -LiteralPath $file).Length / 1MB))
            $file
        }
        'path' {
            $local = [string] $source.path
            if (-not [IO.Path]::IsPathRooted($local)) {
                if (-not $ConfigFolder) { throw "Relative path '$local' needs a JSON file on disk to resolve against, not a URL" }
                $local = Join-Path $ConfigFolder $local
            }
            if (-not (Test-Path -LiteralPath $local)) { throw "Driver path not found (or not reachable from this account): $local" }
            Copy-Item -LiteralPath $local -Destination $download -Recurse -Force
            Write-Ok "Copied $local"
            Join-Path $download (Split-Path $local -Leaf)
        }
    }

    Test-FileHash -Path $fetched -Expected (Get-PropertyValue $source 'sha256')
    Test-DownloadedFile -Path $fetched

    if ((Test-Path -LiteralPath $fetched -PathType Leaf) -and ([IO.Path]::GetExtension($fetched) -in @('.zip', '.cab'))) {
        $extracted = Join-Path $Folder 'extracted'
        Expand-Package -File $fetched -Destination $extracted
        Write-Ok "Extracted $(Split-Path $fetched -Leaf)"
        return $extracted
    }
    if (Test-Path -LiteralPath $fetched -PathType Leaf) { return (Split-Path $fetched -Parent) }
    return $fetched
}

# -- Drivers -------------------------------------------------------------------

function ConvertFrom-DriverVersion {
    <# Get-PrinterDriver reports the version as one 64-bit number: four 16-bit parts. #>
    param($Value)

    if ($null -eq $Value) { return $null }
    $v = [uint64] $Value
    return [version] ('{0}.{1}.{2}.{3}' -f (($v -shr 48) -band 0xFFFF), (($v -shr 32) -band 0xFFFF),
                                           (($v -shr 16) -band 0xFFFF), ($v -band 0xFFFF))
}

function Get-InstalledPrinterDriver {
    <# The driver for this machine's own architecture - not the x86 copy a print
       server keeps for 32-bit clients under the very same name. #>
    param([Parameter(Mandatory)] [string] $Name)

    $driver = @(Get-PrinterDriver -Name $Name -ErrorAction SilentlyContinue |
                Where-Object { -not (Get-PropertyValue $_ 'PrinterEnvironment') -or $_.PrinterEnvironment -eq $NativeArch.Environment }) |
              Select-Object -First 1
    if (-not $driver) { return $null }
    return [PSCustomObject]@{
        Name    = $driver.Name
        Version = ConvertFrom-DriverVersion (Get-PropertyValue $driver 'DriverVersion')
        Inf     = Get-PropertyValue $driver 'InfPath'
    }
}

function Get-InfSection {
    <# The lines of one INF section, comments stripped. #>
    param([Parameter(Mandatory)] [string] $Text, [Parameter(Mandatory)] [string] $Name)

    # Concatenated rather than interpolated: inside double quotes the "$(" of the
    # end-of-line anchor would be read as a PowerShell subexpression.
    $pattern = '(?ims)^\s*\[' + [regex]::Escape($Name) + '\][ \t]*\r?$(.*?)(?=^\s*\[|\z)'
    $match   = [regex]::Match($Text, $pattern)
    if (-not $match.Success) { return @() }
    return @($match.Groups[1].Value -split "`r?`n" | ForEach-Object { ($_ -replace ';.*$', '').Trim() } | Where-Object { $_ })
}

function Get-InfInfo {
    <#
        What an INF says about itself: its class, which architectures its models
        are decorated for, the driver names it declares (with %token% strings
        resolved), and the catalog it is signed through. INFs are UTF-16 as often
        as ANSI; Get-Content reads both by their BOM.
    #>
    param([Parameter(Mandatory)] $Inf)

    $text = Get-Content -LiteralPath $Inf.FullName -Raw

    $strings = @{}
    foreach ($line in Get-InfSection -Text $text -Name 'Strings') {
        if ($line -match '^([^=]+?)\s*=\s*"?(.*?)"?$') { $strings[$Matches[1].Trim()] = $Matches[2] }
    }
    $resolve = {
        param([string] $Value)
        $Value = $Value.Trim().Trim('"')
        if ($Value -match '^%(.+)%$' -and $strings.ContainsKey($Matches[1])) { return $strings[$Matches[1]] }
        return $Value
    }

    $class = if ($text -match '(?im)^\s*Class\s*=\s*([^\s;]+)') { $Matches[1] } else { $null }

    # [Manufacturer] lines look like  %HP% = HP, NTamd64, NTarm64  - the names
    # after the first comma are the decorations, and each one names a models
    # section "HP.NTamd64" whose lines start with the driver names.
    $decorations = [System.Collections.Generic.List[string]]::new()
    $models      = [System.Collections.Generic.List[string]]::new()
    foreach ($line in Get-InfSection -Text $text -Name 'Manufacturer') {
        if ($line -notmatch '=\s*(.+)$') { continue }
        $parts   = @($Matches[1] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $section = $parts[0]
        $decor   = @($parts | Select-Object -Skip 1)
        foreach ($d in $decor) { $decorations.Add(($d -split '\.')[0]) }
        $sections = if ($decor.Count -gt 0) { @($decor | ForEach-Object { "$section.$_" }) } else { @($section) }
        foreach ($modelSection in $sections) {
            foreach ($modelLine in Get-InfSection -Text $text -Name $modelSection) {
                if ($modelLine -match '^("[^"]+"|%[^%]+%)\s*=') { $models.Add((& $resolve $Matches[1])) }
            }
        }
    }

    $catalogs = @{}
    foreach ($m in [regex]::Matches($text, '(?im)^\s*CatalogFile(\.[\w]+)?\s*=\s*([^\s;]+)')) {
        $catalogs[$m.Groups[1].Value.TrimStart('.')] = $m.Groups[2].Value
    }

    return [PSCustomObject]@{
        Class       = $class
        Decorations = @($decorations | Select-Object -Unique)
        Models      = @($models | Select-Object -Unique)
        Catalogs    = $catalogs
    }
}

function Test-InfForThisMachine {
    <# Whether an INF can install on this architecture at all. #>
    param([Parameter(Mandatory)] $Info)
    return [bool] (@($Info.Decorations | Where-Object { $_ -like "$($NativeArch.Decoration)*" }).Count -gt 0)
}

function Find-DriverInf {
    <#
        The INF to install. Named in the JSON, it is taken as given (wildcards
        allowed). Otherwise every INF in the package is read and the one that
        declares this exact driver name for this architecture is it - a universal
        driver package carries a dozen INFs, for other models and other platforms.
    #>
    param([Parameter(Mandatory)] [string] $Folder, [Parameter(Mandatory)] $Driver)

    $infName = Get-PropertyValue $Driver 'inf'
    $all     = @(Get-ChildItem -LiteralPath $Folder -Recurse -File -Filter '*.inf' -ErrorAction SilentlyContinue)
    if ($all.Count -eq 0) {
        $exe = @(Get-ChildItem -LiteralPath $Folder -Recurse -File -Include '*.exe', '*.msi' -ErrorAction SilentlyContinue) | Select-Object -First 1
        $why = if ($exe) { " - it holds $($exe.Name), an installer; extract the driver package from it (most vendors offer a 'driver only' zip)" } else { '' }
        throw "No .inf file in the files for '$($Driver.name)'$why"
    }

    # The @() around the whole if: a package with one INF otherwise yields a bare
    # FileInfo here, and its .Count throws under Set-StrictMode.
    $pool = @(if ($infName) {
        $pattern = $infName.Replace('/', '\')
        $all | Where-Object { $_.Name -like $pattern -or $_.FullName -like "*\$pattern" }
    } else { $all })
    if ($pool.Count -eq 0) { throw "No INF named '$infName' among: $(($all | ForEach-Object { $_.Name } | Select-Object -Unique) -join ', ')" }

    $declared = [System.Collections.Generic.List[string]]::new()
    $matching = [System.Collections.Generic.List[object]]::new()
    $otherArch = $false
    foreach ($inf in $pool) {
        $info = Get-InfInfo -Inf $inf
        foreach ($model in $info.Models) { $declared.Add($model) }
        if (@($info.Models | Where-Object { $_ -eq $Driver.name }).Count -eq 0) { continue }
        if (-not (Test-InfForThisMachine $info)) { $otherArch = $true; continue }
        $matching.Add([PSCustomObject]@{ File = $inf; Info = $info })
    }

    if ($matching.Count -eq 0) {
        if ($otherArch) {
            throw "The package declares '$($Driver.name)', but not for $($NativeArch.Environment) - it has no $($NativeArch.Decoration) section. Use the package for this architecture"
        }
        # The most common mistake by far: a name that is almost right. Name the
        # ones that are there, closest first, so the fix is a copy and paste.
        $words   = @($Driver.name -split '\W+' | Where-Object { $_.Length -gt 2 })
        $ranked  = @($declared | Select-Object -Unique | Sort-Object -Property @{ Expression = {
                        $model = $_; -(@($words | Where-Object { $model -like "*$_*" }).Count) } }, @{ Expression = { $_ } })
        $suggest = if ($ranked.Count -gt 0) { " Names it does declare: $((@($ranked | Select-Object -First 8) | ForEach-Object { "'$_'" }) -join ', ')$(if ($ranked.Count -gt 8) { ", ... ($($ranked.Count) in all)" })" } else { '' }
        throw "No INF in the package declares the driver name '$($Driver.name)' - it has to match the INF exactly.$suggest"
    }

    if ($matching.Count -gt 1) {
        # Prefer the one in an architecture folder, then the shortest path: vendor
        # packages nest the real driver deepest under language and OS folders.
        $byFolder = @($matching | Where-Object { $_.File.DirectoryName -match "(?i)\\($($NativeArch.Folder))(\\|$)" })
        $pick     = if ($byFolder.Count -gt 0) { $byFolder[0] } else { @($matching | Sort-Object { $_.File.FullName.Length })[0] }
        Write-Skip "  $($matching.Count) INFs declare it - using $($pick.File.FullName.Substring($Folder.Length).TrimStart('\')); set 'inf' in the JSON to choose"
        return $pick
    }
    return $matching[0]
}

function Test-DriverCatalog {
    <#
        The INF must be a printer INF, and the catalog it names for this
        architecture must be there and validly signed. Refused here with a clear
        message rather than by pnputil with an exit code.
    #>
    param([Parameter(Mandatory)] $Candidate)

    $inf  = $Candidate.File
    $info = $Candidate.Info
    if ($info.Class -and $info.Class -ne 'Printer') {
        throw "$($inf.Name) is a '$($info.Class)' driver, not a printer driver"
    }

    $catName = $null
    foreach ($key in $info.Catalogs.Keys) {
        if ($key -like "$($NativeArch.Decoration)*") { $catName = $info.Catalogs[$key] }
    }
    if (-not $catName -and $info.Catalogs.ContainsKey(''))   { $catName = $info.Catalogs[''] }
    if (-not $catName -and $info.Catalogs.ContainsKey('NT')) { $catName = $info.Catalogs['NT'] }
    if (-not $catName) { throw "$($inf.Name) names no CatalogFile for $($NativeArch.Environment) - an unsigned driver package, which Windows will not install" }

    $cat = Join-Path $inf.DirectoryName $catName
    if (-not (Test-Path -LiteralPath $cat)) { throw "$($inf.Name) refers to $catName, which is not in the package - it was extracted or copied incompletely" }

    if ($SkipSignatureCheck) {
        Write-Warn "Signature check skipped for $catName (-SkipSignatureCheck)"
        return
    }
    $signature = Get-AuthenticodeSignature -LiteralPath $cat
    if ($signature.Status -ne 'Valid') {
        throw "$catName signature is $($signature.Status) - refusing to install it"
    }
    Write-Ok "$catName is signed: $($signature.SignerCertificate.Subject -replace ',.*$', '')"
}

function Install-DriverPackage {
    <#
        Add the package to the driver store, then install the printer driver from
        it. pnputil is the supported route for a driver that is not in Windows'
        own store; Add-PrinterDriver alone only finds drivers that already are.
    #>
    param([Parameter(Mandatory)] $Inf, [Parameter(Mandatory)] [string] $DriverName)

    $pnputil = Join-Path $env:WINDIR 'System32\pnputil.exe'
    if (-not (Test-Path -LiteralPath $pnputil)) { throw "pnputil.exe is not at $pnputil" }

    $output = & $pnputil /add-driver $Inf.FullName /install 2>&1
    $code   = $LASTEXITCODE
    foreach ($line in @($output)) {
        $text = "$line".Trim()
        if ($text) { Write-Skip "  pnputil: $text" }
    }
    # 259 = no device was waiting for it, which for a printer driver is the normal
    # case: the package is in the store all the same. 3010/1641 = reboot needed or
    # already started - Microsoft's documented pnputil return values.
    if ($code -notin @(0, 259, 3010, 1641)) {
        throw "pnputil could not add $($Inf.Name) (exit code $code) - $env:WINDIR\INF\setupapi.dev.log says why"
    }
    if ($code -in @(3010, 1641)) { $script:rebootRequired = $true }

    # A vendor driver landing in the store is a classic moment for the spooler to
    # stall, so Add-PrinterDriver gets the restart-and-retry. When the spooler
    # does not find the driver by name, the INF in the store is named explicitly.
    try {
        Invoke-SpoolerAction -What 'Add-PrinterDriver' -Action { Add-PrinterDriver -Name $DriverName -ErrorAction Stop }
    } catch {
        $first     = $_.Exception.Message
        $published = @(Get-WindowsDriver -Online -ErrorAction SilentlyContinue |
                       Where-Object { $_.OriginalFileName -like "*\$($Inf.Name)" } |
                       Sort-Object { try { [version] $_.Version } catch { [version] '0.0' } } -Descending) | Select-Object -First 1
        if (-not $published) { throw "Add-PrinterDriver '$DriverName' failed: $first" }
        Write-Warn "Add-PrinterDriver by name failed ($first) - installing from the driver store copy $($published.OriginalFileName)"
        Invoke-SpoolerAction -What 'Add-PrinterDriver' -Action {
            Add-PrinterDriver -Name $DriverName -InfPath $published.OriginalFileName -ErrorAction Stop
        }
    }
}

# -- Printers ------------------------------------------------------------------

function Get-PortName {
    param([Parameter(Mandatory)] $Spec)
    $name = Get-PropertyValue $Spec 'portName'
    if ($name) { return [string] $name }
    # An "absent" entry needs no address; without one there is no port to name.
    return "IP_$(Get-PropertyValue $Spec 'address')"
}

function Test-PrinterAbsent {
    param([Parameter(Mandatory)] $Spec)
    return ((Get-PropertyValue $Spec 'ensure') -eq 'absent')
}

function Get-PortAddress {
    param($Port)
    foreach ($field in 'PrinterHostAddress', 'LprHostAddress', 'HostName') {
        $value = Get-PropertyValue $Port $field
        if ($value) { return [string] $value }
    }
    return $null
}

function Get-PrintSetting {
    <# The printer's current print defaults, read in a job: this is driver code too. #>
    param([Parameter(Mandatory)] [string] $Name)
    return Invoke-WithTimeout -Seconds 60 -ArgumentList $Name -ScriptBlock {
        param($PrinterName)
        $c = Get-PrintConfiguration -PrinterName $PrinterName -ErrorAction Stop
        [PSCustomObject]@{ Duplex = "$($c.DuplexingMode)"; Color = [bool] $c.Color; Paper = "$($c.PaperSize)" }
    }
}

function Get-PrinterState {
    <# What differs between the printer as it is and as the JSON wants it. #>
    param([Parameter(Mandatory)] $Spec)

    $portName = Get-PortName $Spec
    $existing = Get-Printer -Name $Spec.name -ErrorAction SilentlyContinue
    $port     = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
    $changes  = [System.Collections.Generic.List[string]]::new()
    $state    = [PSCustomObject]@{
        Exists      = [bool] $existing
        PortExists  = [bool] $port
        PortDrift   = $false
        PortName    = $portName
        CurrentPort = if ($existing) { $existing.PortName } else { $null }
        Changes     = $changes
    }

    if (Test-PrinterAbsent $Spec) {
        if ($existing) { $changes.Add('printer exists but the JSON says "ensure": "absent"') }
        return $state
    }

    $portAddress = Get-PortAddress $port
    if (-not $port) { $changes.Add("port $portName -> $($Spec.address)") }
    elseif ($portAddress -and $portAddress -ne [string] $Spec.address) {
        $state.PortDrift = $true
        $changes.Add("port $portName points at $portAddress, not $($Spec.address)")
    }

    if (-not $existing) {
        $changes.Add('printer does not exist')
        return $state
    }

    if ($existing.DriverName -ne $Spec.driver) { $changes.Add("driver is '$($existing.DriverName)'") }
    if ($existing.PortName -ne $portName)      { $changes.Add("port is '$($existing.PortName)'") }
    foreach ($field in 'Location', 'Comment') {
        $want = Get-PropertyValue $Spec $field.ToLower()
        # Windows reports an empty field as $null or '' depending on the build;
        # both mean "nothing", and treating them apart is a change on every run.
        if ($null -ne $want -and [string] $want -ne [string] (Get-PropertyValue $existing $field)) {
            $changes.Add("$($field.ToLower()) is '$(Get-PropertyValue $existing $field)'")
        }
    }
    $shared = Get-PropertyValue $Spec 'shared'
    if ($null -ne $shared -and [bool] $shared -ne [bool] $existing.Shared) { $changes.Add("shared is $($existing.Shared)") }
    if ($shared -eq $true) {
        $shareName = if (Get-PropertyValue $Spec 'shareName') { [string] $Spec.shareName } else { [string] $Spec.name }
        if ($existing.Shared -and [string] $existing.ShareName -ne $shareName) { $changes.Add("share name is '$($existing.ShareName)'") }
    }

    $duplex = Get-PropertyValue $Spec 'duplex'
    $color  = Get-PropertyValue $Spec 'color'
    $paper  = Get-PropertyValue $Spec 'paperSize'
    if ($duplex -or $null -ne $color -or $paper) {
        try {
            $current = Get-PrintSetting -Name $Spec.name
            if ($duplex -and $current.Duplex -ne $duplex)                     { $changes.Add("duplex is $($current.Duplex)") }
            if ($null -ne $color -and [bool] $color -ne [bool] $current.Color) { $changes.Add("color is $($current.Color)") }
            if ($paper -and $current.Paper -ne $paper)                         { $changes.Add("paper size is $($current.Paper)") }
        } catch {
            # Unreadable is not the same as wrong: no change is claimed, so a
            # driver that cannot report its settings does not loop every run.
            Write-Warn "Could not read the print defaults of '$($Spec.name)': $($_.Exception.Message)"
        }
    }
    return $state
}

function Add-PortFromSpec {
    param([Parameter(Mandatory)] $Spec, [Parameter(Mandatory)] [string] $PortName)

    $lprQueue = Get-PropertyValue $Spec 'lprQueue'
    if ($lprQueue) {
        Invoke-SpoolerAction -What 'Add-PrinterPort' -Action {
            Add-PrinterPort -Name $PortName -LprHostAddress $Spec.address -LprQueueName $lprQueue -ErrorAction Stop
        }
    } else {
        $number = Get-PropertyValue $Spec 'portNumber'
        if (-not $number) { $number = 9100 }
        # Without -SNMP the port is created with SNMP off - unlike the GUI, which
        # turns it on. Deliberate: a printer that does not answer SNMP with
        # community "public" otherwise shows as Offline to every user.
        $snmp = @{}
        if ((Get-PropertyValue $Spec 'snmp') -eq $true) { $snmp = @{ SNMP = 1; SNMPCommunity = 'public' } }
        Invoke-SpoolerAction -What 'Add-PrinterPort' -Action {
            Add-PrinterPort -Name $PortName -PrinterHostAddress $Spec.address -PortNumber ([int] $number) @snmp -ErrorAction Stop
        }
    }
}

function Test-PortInUse {
    param([Parameter(Mandatory)] [string] $PortName, [string] $Except)
    return [bool] @(Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.PortName -eq $PortName -and $_.Name -ne $Except }).Count
}

function Remove-UnusedPort {
    <# A port this script created (IP_*) that no printer uses any more. Ports
       named otherwise were made by somebody else and are left to them. #>
    param([string] $PortName)

    if (-not $PortName -or $PortName -notlike 'IP_*') { return }
    if (-not (Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue)) { return }
    if (Test-PortInUse -PortName $PortName) { return }
    if (-not $PSCmdlet.ShouldProcess($PortName, 'Remove-PrinterPort (no printer uses it any more)')) { return }
    try {
        Remove-PrinterPort -Name $PortName -ErrorAction Stop
        Write-Ok "Port $PortName removed - no printer uses it any more"
    } catch {
        # The spooler can hold on to a port for a moment after its printer moved;
        # an unused port is harmless and a later run removes it.
        Write-Skip "  Port $PortName is still held by the spooler - it goes on a later run"
    }
}

function Remove-PrinterFromSpec {
    <#
        Take a printer away. The driver stays: other printers may use it, and a
        driver in use cannot be removed anyway. The port goes only when this
        printer was the last one on it.
    #>
    param([Parameter(Mandatory)] $Spec, [Parameter(Mandatory)] $State)

    if ($PSCmdlet.ShouldProcess($Spec.name, 'Remove-Printer')) {
        # Jobs still in the queue keep a printer in "pending deletion" for as long
        # as they stay there; clearing them first makes the removal immediate.
        Get-PrintJob -PrinterName $Spec.name -ErrorAction SilentlyContinue | Remove-PrintJob -ErrorAction SilentlyContinue
        Invoke-SpoolerAction -What 'Remove-Printer' -Action { Remove-Printer -Name $Spec.name -ErrorAction Stop }
        Write-Ok "Printer '$($Spec.name)' removed"
    }
    Remove-UnusedPort -PortName $State.CurrentPort
}

function Set-PrinterFromSpec {
    param([Parameter(Mandatory)] $Spec, [Parameter(Mandatory)] $State)

    $portName = $State.PortName
    if (-not $State.PortExists) {
        if ($PSCmdlet.ShouldProcess("$portName -> $($Spec.address)", 'Add-PrinterPort')) {
            Add-PortFromSpec -Spec $Spec -PortName $portName
            Write-Ok "Port $portName created ($($Spec.address))"
        }
    } elseif ($State.PortDrift) {
        # There is no Set-PrinterPort. A port only this printer uses is rebuilt in
        # place: the printer moves to a temporary port with the new address, the
        # old port is recreated, the printer moves back. If any step fails, the
        # printer is left on the temporary port - which points at the right
        # address and prints. A port other printers share is not this one's to change.
        if (Test-PortInUse -PortName $portName -Except $Spec.name) {
            throw "Port $portName is shared with other printers and points at another address - give this printer its own 'portName' in the JSON"
        }
        if ($PSCmdlet.ShouldProcess($portName, "Rebuild the port for $($Spec.address)")) {
            $temporary = "$portName`_new"
            if (-not (Get-PrinterPort -Name $temporary -ErrorAction SilentlyContinue)) { Add-PortFromSpec -Spec $Spec -PortName $temporary }
            if ($State.Exists) { Invoke-SpoolerAction -What 'Set-Printer' -Action { Set-Printer -Name $Spec.name -PortName $temporary -ErrorAction Stop } }
            Invoke-SpoolerAction -What 'Remove-PrinterPort' -Action { Remove-PrinterPort -Name $portName -ErrorAction Stop }
            Add-PortFromSpec -Spec $Spec -PortName $portName
            if ($State.Exists) { Invoke-SpoolerAction -What 'Set-Printer' -Action { Set-Printer -Name $Spec.name -PortName $portName -ErrorAction Stop } }
            try { Remove-PrinterPort -Name $temporary -ErrorAction Stop } catch { Write-Skip "  Temporary port $temporary is removed on a later run" }
            Write-Ok "Port $portName now points at $($Spec.address)"
        }
    }

    $extra = @{}
    foreach ($field in 'Location', 'Comment') {
        $value = Get-PropertyValue $Spec $field.ToLower()
        if ($null -ne $value) { $extra[$field] = [string] $value }
    }
    $shared = Get-PropertyValue $Spec 'shared'
    if ($null -ne $shared) {
        $extra['Shared'] = [bool] $shared
        if ($shared) { $extra['ShareName'] = if (Get-PropertyValue $Spec 'shareName') { [string] $Spec.shareName } else { [string] $Spec.name } }
    }

    if (-not $State.Exists) {
        if ($PSCmdlet.ShouldProcess($Spec.name, "Add-Printer (driver '$($Spec.driver)', port $portName)")) {
            Invoke-SpoolerAction -What 'Add-Printer' -Action {
                Add-Printer -Name $Spec.name -DriverName $Spec.driver -PortName $portName @extra -ErrorAction Stop
            }
            Write-Ok "Printer '$($Spec.name)' added"
        }
    } elseif ($PSCmdlet.ShouldProcess($Spec.name, "Set-Printer ($($State.Changes -join '; '))")) {
        Invoke-SpoolerAction -What 'Set-Printer' -Action {
            Set-Printer -Name $Spec.name -DriverName $Spec.driver -PortName $portName @extra -ErrorAction Stop
        }
        Write-Ok "Printer '$($Spec.name)' corrected"
        # With default port names a new address is a new port; the old one is
        # left without a printer and would otherwise linger forever.
        if ($State.CurrentPort -and $State.CurrentPort -ne $portName) { Remove-UnusedPort -PortName $State.CurrentPort }
    }

    # Print defaults for every user of this printer. Only what the JSON sets is
    # touched; what the driver does not support or hangs on is a warning, because
    # the printer itself is installed by now and prints.
    $config = @{}
    if (Get-PropertyValue $Spec 'duplex')            { $config['DuplexingMode'] = [string] $Spec.duplex }
    if ($null -ne (Get-PropertyValue $Spec 'color')) { $config['Color'] = [bool] $Spec.color }
    if (Get-PropertyValue $Spec 'paperSize')         { $config['PaperSize'] = [string] $Spec.paperSize }
    if ($config.Count -gt 0 -and $PSCmdlet.ShouldProcess($Spec.name, "Set-PrintConfiguration ($(($config.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', '))")) {
        try {
            $null = Invoke-WithTimeout -Seconds 120 -ArgumentList $Spec.name, $config -ScriptBlock {
                param($Name, $Settings)
                Set-PrintConfiguration -PrinterName $Name @Settings -ErrorAction Stop
            }
            Write-Ok "Print defaults set for '$($Spec.name)'"
        } catch {
            Write-Warn "Could not set the print defaults for '$($Spec.name)': $($_.Exception.Message) - the printer itself is installed"
        }
    }
}

function Set-SuccessMarker {
    <#
        What the last clean run applied, for an Intune detection rule or an RMM
        condition to compare against: the SHA256 of the JSON it read, when, and
        which printers. A changed JSON gives a different hash, so detection can
        tell "installed, but with last month's configuration" apart from current.
    #>
    param([Parameter(Mandatory)] [string] $ConfigText, $Printers)

    try {
        $sha  = [Security.Cryptography.SHA256]::Create()
        $hash = -join ($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($ConfigText)) | ForEach-Object { $_.ToString('X2') })
        $sha.Dispose()
        if (-not (Test-Path $MarkerPath)) { New-Item -Path $MarkerPath -Force | Out-Null }
        Set-ItemProperty -Path $MarkerPath -Name 'ConfigSha256' -Value $hash
        Set-ItemProperty -Path $MarkerPath -Name 'ConfigPath'   -Value $ConfigPath
        Set-ItemProperty -Path $MarkerPath -Name 'LastSuccess'  -Value (Get-Date).ToString('s')
        Set-ItemProperty -Path $MarkerPath -Name 'Printers'     -Value (@($Printers | ForEach-Object { $_.name }) -join '; ')
    } catch {
        Write-Warn "Could not write the success marker under ${MarkerPath}: $($_.Exception.Message)"
    }
}

# -- Configuration -------------------------------------------------------------

function Read-Configuration {
    <#
        The JSON as text. From a URL it is cached after every successful read, and
        the cache is the fallback when the URL cannot be reached: a host that boots
        while GitHub is unreachable keeps the printers it had, instead of failing.
    #>
    $cache = Join-Path $WorkingDir 'printers.cache.json'

    if ($ConfigPath -notmatch '^[a-z]+://') {
        if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { throw "Configuration not found (or not reachable from this account): $ConfigPath" }
        return [PSCustomObject]@{
            Text   = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8
            Folder = Split-Path (Resolve-Path -LiteralPath $ConfigPath).ProviderPath -Parent
        }
    }
    if ($ConfigPath -notmatch '^https://') { throw "ConfigPath URL must be https: $ConfigPath" }

    $temp = Join-Path ([IO.Path]::GetTempPath()) ('printers_{0}.json' -f [guid]::NewGuid().ToString('N'))
    try {
        Invoke-Download -Uri $ConfigPath -Path $temp -Accept 'application/vnd.github.raw'
        $text = Get-Content -LiteralPath $temp -Raw -Encoding UTF8
        if (-not ($simulate -or $CheckOnly)) {
            # Only a file that parses is worth falling back on later.
            try {
                $null = $text | ConvertFrom-Json
                if (-not (Test-Path -LiteralPath $WorkingDir)) { New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null }
                Copy-Item -LiteralPath $temp -Destination $cache -Force
            } catch { }
        }
        return [PSCustomObject]@{ Text = $text; Folder = $null }
    } catch {
        if (-not (Test-Path -LiteralPath $cache)) { throw }
        Write-Warn "$($_.Exception.Message)"
        Write-Warn "Using the copy cached at $((Get-Item -LiteralPath $cache).LastWriteTime.ToString('yyyy-MM-dd HH:mm')) ($cache) - changes made to the JSON since then are not applied by this run"
        return [PSCustomObject]@{ Text = Get-Content -LiteralPath $cache -Raw -Encoding UTF8; Folder = $null }
    } finally {
        Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
    }
}

function Test-Configuration {
    <#
        Every problem in the JSON at once, before anything is touched. One message
        listing them all beats fixing one typo per run on a host that reboots for
        each attempt.
    #>
    param([Parameter(Mandatory)] $Config)

    $problems = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()

    if (-not (Test-JsonObject $Config)) {
        $problems.Add('the file must be one JSON object with "drivers" and "printers"')
        return [PSCustomObject]@{ Problems = $problems; Warnings = $warnings }
    }
    foreach ($field in Get-FieldName $Config) {
        if ($field -notin $KnownRootFields) { $warnings.Add("unknown top-level field '$field' is ignored") }
    }
    foreach ($field in 'drivers', 'printers') {
        $value = Get-PropertyValue $Config $field
        if ($null -ne $value -and $value -isnot [array] -and -not (Test-JsonObject $value)) { $problems.Add("'$field' must be a list") }
    }

    $driverNames = @{}
    $index = 0
    foreach ($d in @(Get-PropertyValue $Config 'drivers')) {
        $index++
        if ($null -eq $d) { continue }
        $name = [string] (Get-PropertyValue $d 'name')
        $who  = if ($name) { "driver '$name'" } else { "driver #$index" }
        if (-not (Test-JsonObject $d)) { $problems.Add("$who must be an object"); continue }
        foreach ($field in Get-FieldName $d) { if ($field -notin $KnownDriverFields) { $problems.Add("$who has an unknown field '$field' (typo?)") } }
        if (-not $name) { $problems.Add("$who has no 'name'") }
        elseif ($driverNames.ContainsKey($name)) { $problems.Add("$who is listed twice") }
        else { $driverNames[$name] = $true }

        $version = Get-PropertyValue $d 'version'
        if ($null -ne $version) { try { $null = [version] [string] $version } catch { $problems.Add("$who has version '$version', which is not a.b.c.d") } }
        $install = Get-PropertyValue $d 'install'
        if ($null -ne $install -and $install -isnot [bool]) { $problems.Add("$who 'install' must be true or false") }

        $source = Get-PropertyValue $d 'source'
        if (-not (Test-JsonObject $source)) { $problems.Add("$who has no 'source' object"); continue }
        $type = [string] (Get-PropertyValue $source 'type')
        if (-not $KnownSourceFields.ContainsKey($type)) {
            $problems.Add("$who has source type '$type' - use githubRelease, github, url or path"); continue
        }
        foreach ($field in Get-FieldName $source) {
            if ($field -notin $KnownSourceFields[$type]) { $problems.Add("$who source has an unknown field '$field' for type $type") }
        }
        switch ($type) {
            'githubRelease' {
                if (-not (Test-RepositoryName (Get-PropertyValue $source 'repository'))) { $problems.Add("$who 'repository' must be 'owner/name'") }
                if (-not (Get-PropertyValue $source 'asset')) { $problems.Add("$who has no 'asset'") }
            }
            'github' {
                if (-not (Test-RepositoryName (Get-PropertyValue $source 'repository'))) { $problems.Add("$who 'repository' must be 'owner/name'") }
                if (-not ([string] (Get-PropertyValue $source 'path')).Trim('/\')) { $problems.Add("$who has no 'path'") }
            }
            'url'  { if ([string] (Get-PropertyValue $source 'url') -notmatch '^https://[^\s/]+/\S*') { $problems.Add("$who 'url' must be an https URL") } }
            'path' { if (-not (Get-PropertyValue $source 'path')) { $problems.Add("$who has no 'path'") } }
        }
        $hash = Get-PropertyValue $source 'sha256'
        if ($hash -and [string] $hash -notmatch '^[0-9A-Fa-f]{64}$') { $problems.Add("$who 'sha256' must be 64 hexadecimal characters") }
    }

    $printerNames = @{}
    $ports        = @{}
    $index = 0
    foreach ($p in @(Get-PropertyValue $Config 'printers')) {
        $index++
        if ($null -eq $p) { continue }
        $name = [string] (Get-PropertyValue $p 'name')
        $who  = if ($name) { "printer '$name'" } else { "printer #$index" }
        if (-not (Test-JsonObject $p)) { $problems.Add("$who must be an object"); continue }
        foreach ($field in Get-FieldName $p) { if ($field -notin $KnownPrinterFields) { $problems.Add("$who has an unknown field '$field' (typo?)") } }

        if (-not $name) { $problems.Add("$who has no 'name'") }
        elseif ($printerNames.ContainsKey($name)) { $problems.Add("$who is listed twice") }
        else {
            $printerNames[$name] = $true
            # The spooler refuses these, and a printer name of more than 220
            # characters is cut off in places Windows does not warn about.
            if ($name -match '[\\,!]') { $problems.Add("$who contains \ , or ! - Windows does not allow those in a printer name") }
            if ($name.Length -gt 220)  { $problems.Add("$who is longer than 220 characters") }
        }

        $ensure = Get-PropertyValue $p 'ensure'
        if ($null -ne $ensure -and $ensure -notin @('present', 'absent')) { $problems.Add("$who 'ensure' must be present or absent") }
        if ($ensure -eq 'absent') { continue }

        $driver  = [string] (Get-PropertyValue $p 'driver')
        $address = [string] (Get-PropertyValue $p 'address')
        if (-not $driver) { $problems.Add("$who has no 'driver'") }
        if (-not $address) { $problems.Add("$who has no 'address'") }
        else {
            $ip = $null
            $isIp       = [Net.IPAddress]::TryParse($address, [ref] $ip)
            $isHostname = $address -match '^(?=.{1,253}$)[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?)*$'
            if (-not ($isIp -or $isHostname)) { $problems.Add("$who address '$address' is neither an IP address nor a host name") }
        }

        $portNumber = Get-PropertyValue $p 'portNumber'
        if ($null -ne $portNumber -and -not ("$portNumber" -match '^\d+$' -and [int] "$portNumber" -ge 1 -and [int] "$portNumber" -le 65535)) {
            $problems.Add("$who portNumber must be 1-65535")
        }
        $portName = Get-PropertyValue $p 'portName'
        if ($portName -and [string] $portName -match '[\\,]') { $problems.Add("$who portName may not contain \ or ,") }

        $duplex = Get-PropertyValue $p 'duplex'
        if ($null -ne $duplex -and $duplex -notin $DuplexValues) { $problems.Add("$who duplex must be one of $($DuplexValues -join ', ')") }
        foreach ($flag in 'color', 'shared', 'snmp') {
            $value = Get-PropertyValue $p $flag
            if ($null -ne $value -and $value -isnot [bool]) { $problems.Add("$who '$flag' must be true or false (without quotes)") }
        }
        foreach ($text in 'location', 'comment', 'shareName', 'paperSize', 'lprQueue') {
            $value = Get-PropertyValue $p $text
            if ($null -ne $value -and $value -isnot [string]) { $problems.Add("$who '$text' must be text") }
        }
        if ((Get-PropertyValue $p 'lprQueue') -and $null -ne $portNumber) { $warnings.Add("$who has lprQueue, so portNumber is ignored (LPR always uses 515)") }

        # Two printers may share a port - but not with two different addresses,
        # because the second would silently print to the first one's printer.
        $effectivePort = if ($portName) { [string] $portName } else { "IP_$address" }
        if ($ports.ContainsKey($effectivePort) -and $ports[$effectivePort] -ne $address) {
            $problems.Add("$who and another printer both use port '$effectivePort' with different addresses")
        } else { $ports[$effectivePort] = $address }
    }

    return [PSCustomObject]@{ Problems = $problems; Warnings = $warnings }
}

# -- Run -----------------------------------------------------------------------
try {
    Start-LogFile
    Write-Out ''
    Write-Out "  Mode: $(if ($CheckOnly) { 'CHECK ONLY - reporting, never installing' } elseif ($simulate) { '-WhatIf - nothing will be changed' } else { 'APPLY - drivers and printers are installed where needed' })" 'Cyan'
    Write-Out ''

    # A second run waits here, before it reads anything: what it reads after the
    # first one finished is what it should act on.
    if (-not ($simulate -or $CheckOnly)) { Enter-SingleRun }

    # -- 1. Config -------------------------------------------------------------
    Write-Step '1. Configuration'
    $loaded       = Read-Configuration
    $configFolder = $loaded.Folder
    if (-not "$($loaded.Text)".Trim()) { throw "$ConfigPath is empty" }
    try { $config = $loaded.Text | ConvertFrom-Json } catch { throw "$ConfigPath is not valid JSON: $($_.Exception.Message)" }

    $check = Test-Configuration -Config $config
    foreach ($warning in $check.Warnings) { Write-Warn $warning }
    if ($check.Problems.Count -gt 0) {
        foreach ($problem in $check.Problems) { Write-Bad $problem }
        throw "The configuration has $($check.Problems.Count) problem(s), listed above - nothing was changed."
    }

    $allDrivers  = @(@(Get-PropertyValue $config 'drivers')  | Where-Object { $_ })
    $allPrinters = @(@(Get-PropertyValue $config 'printers') | Where-Object { $_ })
    Write-Ok "$ConfigPath - $($allDrivers.Count) driver(s), $($allPrinters.Count) printer(s)"

    # The @() goes around the whole if: one match inside the branch alone comes
    # out as a bare object, and its .Count throws under Set-StrictMode.
    $printers = @(if ($Printer) {
        $allPrinters | Where-Object { $printerName = $_.name; @($Printer | Where-Object { $printerName -like $_ }).Count -gt 0 }
    } else { $allPrinters })
    if ($Printer -and $printers.Count -eq 0) {
        throw "No printer in the JSON matches $($Printer -join ', ') (it has: $(($allPrinters | ForEach-Object { $_.name }) -join ', '))"
    }
    if ($printers.Count -eq 0 -and @($allDrivers | Where-Object { (Get-PropertyValue $_ 'install') -eq $true }).Count -eq 0) {
        Write-Warn 'The configuration lists no printers - there is nothing to install'
    }

    # -- 2. Preflight ----------------------------------------------------------
    # The spooler first: every question below is a question to it, and on a first
    # boot asking too early answers "not installed" for everything.
    Write-Out ''
    Write-Step '2. Preflight'
    Wait-PrintSpooler

    $neededNames = @($printers | Where-Object { -not (Test-PrinterAbsent $_) } | ForEach-Object { [string] $_.driver } | Select-Object -Unique)
    $drivers     = @($allDrivers | Where-Object { $_.name -in $neededNames -or (Get-PropertyValue $_ 'install') -eq $true })
    foreach ($name in $neededNames) {
        if ($name -notin @($allDrivers | ForEach-Object { $_.name }) -and -not (Get-InstalledPrinterDriver -Name $name)) {
            throw "Printer driver '$name' is neither in the JSON's drivers nor installed on this machine"
        }
    }

    $driverWork = [System.Collections.Generic.List[object]]::new()
    foreach ($driver in $drivers) {
        $installed = Get-InstalledPrinterDriver -Name $driver.name
        $wanted    = Get-PropertyValue $driver 'version'
        if (-not $installed) {
            Write-News "Driver '$($driver.name)' is not installed"
            $driverWork.Add($driver)
        } elseif ($wanted -and $installed.Version -and $installed.Version -lt [version] [string] $wanted) {
            Write-News "Driver '$($driver.name)' is $($installed.Version), the JSON asks for $wanted"
            $driverWork.Add($driver)
        } elseif ($Force) {
            Write-Ok "Driver '$($driver.name)' $($installed.Version) is installed - reinstalling because -Force was given"
            $driverWork.Add($driver)
        } else {
            Write-Ok "Driver '$($driver.name)' $($installed.Version) is installed"
        }
    }

    $printerWork = [System.Collections.Generic.List[object]]::new()
    foreach ($spec in $printers) {
        $state = Get-PrinterState -Spec $spec
        if ($state.Changes.Count -eq 0) {
            if (Test-PrinterAbsent $spec) { Write-Ok "Printer '$($spec.name)' is absent, as configured" }
            else { Write-Ok "Printer '$($spec.name)' is as configured" }
        } else {
            Write-News "Printer '$($spec.name)': $($state.Changes -join '; ')"
            $printerWork.Add([PSCustomObject]@{ Spec = $spec; State = $state })
        }
    }

    if ($driverWork.Count -eq 0 -and $printerWork.Count -eq 0) {
        $plannedExit = 0
        throw 'Every driver and printer is as configured - nothing to do.'
    }

    if ($CheckOnly) {
        Show-HeldOutput
        Write-News ("Work is due: {0} driver(s), {1} printer(s) (exit code 2)" -f $driverWork.Count, $printerWork.Count)
        $plannedExit = 2
        throw 'Check only - nothing was changed.'
    }

    Show-HeldOutput
    if (-not $simulate -and -not $confirmSuppressed -and [Environment]::UserInteractive) {
        Write-Out ''
        $answer = Read-Host ("  Install {0} driver(s) and set up {1} printer(s)? [y/N]" -f $driverWork.Count, $printerWork.Count)
        if ($answer -notmatch '^[Yy]') {
            $plannedExit = 0
            throw 'Cancelled - nothing was changed.'
        }
    }

    # -- 3/4. Download and install the drivers ---------------------------------
    # A driver that fails takes only its own printers down: the others still get
    # what they need, and the exit code says something went wrong.
    $failedDrivers = [System.Collections.Generic.List[string]]::new()
    if ($driverWork.Count -gt 0 -and -not $simulate) { Assert-FreeSpace -Folder $WorkingDir }
    foreach ($driver in $driverWork) {
        Write-Out ''
        Write-Step "3. Driver '$($driver.name)'"
        if (-not $PSCmdlet.ShouldProcess($driver.name, "Download from $($driver.source.type) source and install")) { continue }
        $folder = $null
        try {
            $folder    = New-DriverFolder -Name $driver.name
            $files     = Save-DriverSource -Driver $driver -ConfigFolder $configFolder -Folder $folder
            $candidate = Find-DriverInf -Folder $files -Driver $driver
            Write-Ok "INF: $($candidate.File.FullName.Substring($files.Length).TrimStart('\'))"
            Test-DriverCatalog -Candidate $candidate
            Install-DriverPackage -Inf $candidate.File -DriverName $driver.name

            $now = Get-InstalledPrinterDriver -Name $driver.name
            if (-not $now) { throw "pnputil and Add-PrinterDriver reported success, yet '$($driver.name)' is not installed for $($NativeArch.Environment)" }
            $wanted = Get-PropertyValue $driver 'version'
            if ($wanted -and $now.Version -and $now.Version -lt [version] [string] $wanted) {
                Write-Warn "Driver '$($driver.name)' is $($now.Version) after the install, the JSON asks for $wanted - the package in the source is older than its 'version' says"
            }
            Write-Ok "Driver '$($driver.name)' $($now.Version) installed"

            # The driver store has its own copy now; the download is only clutter.
            Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction SilentlyContinue
        } catch {
            $failedDrivers.Add($driver.name)
            Write-Bad "Driver '$($driver.name)': $($_.Exception.Message)"
            if ($folder) { Write-Skip "  The downloaded files are kept for inspection in $folder" }
            $exitCode = 1
        }
    }

    # -- 5. Printers -----------------------------------------------------------
    Write-Out ''
    Write-Step '5. Printers'
    if ($printerWork.Count -eq 0) { Write-Skip 'Every printer is already as configured' }
    foreach ($item in $printerWork) {
        $spec = $item.Spec
        try {
            if (Test-PrinterAbsent $spec) {
                Remove-PrinterFromSpec -Spec $spec -State $item.State
                continue
            }
            if ($spec.driver -in $failedDrivers) {
                Write-Bad "Printer '$($spec.name)' skipped - its driver failed to install"
                $exitCode = 1
                continue
            }
            if (-not $simulate -and -not (Get-InstalledPrinterDriver -Name $spec.driver)) {
                Write-Bad "Printer '$($spec.name)' skipped - driver '$($spec.driver)' is not installed"
                $exitCode = 1
                continue
            }
            Set-PrinterFromSpec -Spec $spec -State $item.State
        } catch {
            Write-Bad "Printer '$($spec.name)': $($_.Exception.Message)"
            $exitCode = 1
        }
    }

    # -- 6. Verify -------------------------------------------------------------
    Write-Out ''
    Write-Step '6. Verification'
    if ($simulate) {
        Write-Skip 'Skipped - nothing was changed (-WhatIf)'
        Write-Out ''
        Write-Out '  Dry run only - rerun without -WhatIf to apply these changes.' 'Yellow'
    } else {
        foreach ($spec in $printers) {
            $state = Get-PrinterState -Spec $spec
            if (Test-PrinterAbsent $spec) {
                if ($state.Exists) { Write-Bad "Printer '$($spec.name)' is still there"; $exitCode = 1 }
                else { Write-Ok "Printer '$($spec.name)' is gone" }
            } elseif ($state.Changes.Count -eq 0) { Write-Ok "Printer '$($spec.name)' is as configured" }
            elseif ($state.Exists) {
                # Installed and printing, but a setting did not take - usually a
                # print default the driver does not support. Worth knowing, not a failure.
                Write-Warn "Printer '$($spec.name)' still differs: $($state.Changes -join '; ')"
            } else { Write-Bad "Printer '$($spec.name)' is not installed"; $exitCode = 1 }
        }
        Write-Out ''
        if ($rebootRequired) { Write-Warn 'pnputil asked for a reboot to finish the driver installation' }
        if ($exitCode -eq 0) {
            Write-Ok 'Done'
            Set-SuccessMarker -ConfigText $loaded.Text -Printers $printers
        } elseif ($script:logFile) { Write-Out "  Log: $script:logFile" 'Yellow' }
    }
} catch {
    if ($null -ne $plannedExit) {
        Write-Out ''
        Write-Skip $_.Exception.Message
        $exitCode = $plannedExit
    } else {
        Show-HeldOutput
        Write-Out ''
        Write-Bad "Aborted: $($_.Exception.Message)"
        if ($script:logFile) { Write-Out "  Log: $script:logFile" 'Yellow' }
        $exitCode = 1
    }
} finally {
    if ($mutexOwned) { try { $mutex.ReleaseMutex() } catch { } }
    if ($mutex) { $mutex.Dispose() }
    Write-Log "--- Install-Printer finished with exit code $exitCode ---"
}

if (-not $script:holdOutput) { Write-Host '' }
exit $exitCode
