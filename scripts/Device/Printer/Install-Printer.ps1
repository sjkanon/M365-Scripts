#Requires -Version 5.1
<#
.SYNOPSIS
    Install printer drivers and printers as described in a JSON file, with the
    drivers downloaded from a GitHub repository. Supports -WhatIf and -CheckOnly.

.DESCRIPTION
    One JSON file says which drivers exist, where each one is downloaded from and
    which printers use them. This script makes the device match it:

      1. Config     - read the JSON, from a local path, a UNC path or an https URL
                      (a GitHub raw link to the same repository works), and pick
                      the printers asked for with -Printer.
      2. Preflight  - for every driver those printers need: is it installed, and in
                      which version. For every printer: does it exist, on which port
                      and with which driver.
      3. Download   - only for a driver that is missing or older than the JSON asks
                      for. Sources: a GitHub release asset, a file or folder in a
                      GitHub repository, any https URL, or a local/UNC path. A .zip
                      is extracted; an optional sha256 in the JSON is checked first.
      4. Driver     - find the INF (named in the JSON, or the one that mentions the
                      driver name), check the catalog signature, add it to the
                      driver store with pnputil and install it with Add-PrinterDriver.
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
    refers to. printers.example.json beside this script shows every field.

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
    downloads that fail on DNS or a timeout are retried - both for up to
    -WaitSeconds. A 401/403/404 is an answer rather than a hiccup and fails at once.
    The run is idempotent, so the same command can be scheduled at every startup:
    a host that is in order costs one read of the JSON. Printers installed this way
    are machine-wide, so on a session host every user sees them.

    Driver sources
    --------------
      githubRelease  repository, tag (default: latest), asset (name, wildcards allowed)
      github         repository, path (a .zip or a folder holding the INF), ref
                     (branch/tag/commit, default: the repository's default branch)
      url            url (https), downloads one file - .zip or a lone .inf/.cab
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
    workingDir, logPath, waitSeconds, and the checkboxes checkOnly, whatIf, quiet, force and
    skipSignatureCheck. Started 32-bit, the script relaunches itself 64-bit -
    pnputil does not exist under SysWOW64. Started by hand without elevation, it
    asks for it.

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

.PARAMETER WorkingDir
    Folder for downloads and extracted drivers (default: C:\IT\Printers).

.PARAMETER LogPath
    Folder for the transcript of a run that changes something (default: C:\Temp).

.PARAMETER CheckOnly
    Report what would be done and change nothing. Exit code 2 when work is due.

.PARAMETER Quiet
    Print nothing unless there is work to do or something failed.

.PARAMETER Force
    Reinstall the drivers even when the installed version already matches.

.PARAMETER SkipSignatureCheck
    Do not require a valid signature on the driver's catalog file. Windows itself
    still refuses an unsigned package driver; this only skips the script's own check.

.PARAMETER WaitSeconds
    How long a freshly provisioned machine is given to get ready: the Print Spooler
    to start, and the network to answer downloads (default: 300, 0 = no waiting).

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
    Platform: Windows only (PrintManagement + pnputil), run as administrator or as System
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [string]   $ConfigPath,
    [string[]] $Printer,
    [string]   $GitHubToken,
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
    Write-Warning 'Running 32-bit and SysNative is unavailable - pnputil will not be found.'
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
if (-not $PSBoundParameters.ContainsKey('WaitSeconds')        -and $env:waitSeconds -match '^\d+$') { $WaitSeconds = [int] $env:waitSeconds }
if (-not $GitHubToken) { $GitHubToken = if ($env:githubToken) { $env:githubToken } else { $env:GITHUB_TOKEN } }

if (-not $ConfigPath) {
    Write-Error 'No configuration given. Pass -ConfigPath (a JSON file or an https URL), or set the configPath script variable.'
    exit 1
}

$simulate          = [bool] $WhatIfPreference
$confirmSuppressed = $PSBoundParameters.ContainsKey('Confirm') -and -not $PSBoundParameters['Confirm']
$exitCode          = 0
$plannedExit       = $null
$transcribing      = $false
$rebootRequired    = $false

# The hosts a GitHub token may be sent to. A "url" source pointing anywhere else
# is downloaded without it - a token in a JSON-chosen request is a token leaked.
$GitHubHosts = @('api.github.com', 'github.com', 'raw.githubusercontent.com', 'codeload.github.com', 'objects.githubusercontent.com')

# -- Output --------------------------------------------------------------------
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

# -- Readiness -----------------------------------------------------------------
# On the first boot of a server provisioned from a golden image this runs while
# Windows is still coming up: DNS that does not resolve yet, a spooler that is
# set to Automatic but not started. Both settle within minutes, so they are
# waited for (up to -WaitSeconds) rather than reported as a failure that leaves
# a brand-new host without its printers.

function Invoke-WithRetry {
    <#
        Retry a network call while the error looks transient. A 401/403/404 is an
        answer, not a hiccup - retrying it only delays the message that says so.
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
    $deadline = (Get-Date).AddSeconds($WaitSeconds)
    $service  = Get-Service -Name Spooler -ErrorAction SilentlyContinue
    if (-not $service) { throw 'The Print Spooler service does not exist on this machine.' }
    if ($service.StartType -eq 'Disabled') {
        throw 'The Print Spooler service is disabled - a hardened image often does this (PrintNightmare). Enable it before installing printers.'
    }
    while ($true) {
        $service.Refresh()
        if ($service.Status -eq 'Running') { return }
        if ($service.Status -eq 'Stopped') {
            try { Start-Service -Name Spooler -ErrorAction Stop } catch { Write-Warn "Could not start the Print Spooler yet: $($_.Exception.Message)" }
        }
        if ((Get-Date) -gt $deadline) { throw "The Print Spooler service is still $($service.Status) after $WaitSeconds seconds." }
        Write-Warn "Waiting for the Print Spooler service ($($service.Status))"
        Start-Sleep -Seconds 5
    }
}

# -- Downloads -----------------------------------------------------------------

function Get-RequestHeader {
    <# Headers for one request; the token only travels to GitHub itself. #>
    param([Parameter(Mandatory)] [string] $Uri, [string] $Accept)

    $headers = @{ 'User-Agent' = 'M365-Scripts-Install-Printer' }
    if ($Accept) { $headers['Accept'] = $Accept }
    if ($GitHubToken -and (([uri] $Uri).Host -in $GitHubHosts)) {
        $headers['Authorization'] = "Bearer $GitHubToken"
    }
    return $headers
}

function Invoke-Download {
    <# One file to disk over https, with the right headers and without a progress bar. #>
    param([Parameter(Mandatory)] [string] $Uri, [Parameter(Mandatory)] [string] $Path, [string] $Accept)

    if ($Uri -notmatch '^https://') { throw "Download URL must be https: $Uri" }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $progressBackup     = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    try {
        $parent = Split-Path $Path -Parent
        if (-not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        Invoke-WithRetry -What $Uri -Action {
            Invoke-WebRequest -Uri $Uri -OutFile $Path -Headers (Get-RequestHeader -Uri $Uri -Accept $Accept) -UseBasicParsing
        }
    } catch {
        # A private repository answers 404, not 403, to a request without a token -
        # which reads like a typo in the path when it is really a missing token.
        $hint = if (-not $GitHubToken -and (([uri] $Uri).Host -in $GitHubHosts) -and "$($_.Exception.Message)" -match '404') {
            ' - for a private repository, pass -GitHubToken or set GITHUB_TOKEN'
        } else { '' }
        throw "Download failed ($Uri): $($_.Exception.Message)$hint"
    } finally {
        $ProgressPreference = $progressBackup
    }
}

function Invoke-GitHubApi {
    param([Parameter(Mandatory)] [string] $Path)

    $uri = "https://api.github.com/$($Path.TrimStart('/'))"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    try {
        return Invoke-WithRetry -What $uri -Action {
            Invoke-RestMethod -Uri $uri -Headers (Get-RequestHeader -Uri $uri -Accept 'application/vnd.github+json') -UseBasicParsing
        }
    } catch {
        $hint = if (-not $GitHubToken -and "$($_.Exception.Message)" -match '404') {
            ' - check the name, or pass -GitHubToken for a private repository'
        } elseif ("$($_.Exception.Message)" -match '403') {
            ' - rate limited or no access; a token raises the limit from 60 to 5000 requests an hour'
        } else { '' }
        throw "GitHub API $uri failed: $($_.Exception.Message)$hint"
    }
}

function Test-RepositoryName {
    param([string] $Repository)
    if ($Repository -notmatch '^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$') {
        throw "repository must be 'owner/name', got '$Repository'"
    }
}

function Save-GitHubReleaseAsset {
    <#
        A release asset by name. Downloaded through the API's asset URL rather than
        browser_download_url: only the API one honours a token, so the same code
        works for a public and a private repository.
    #>
    param([Parameter(Mandatory)] $Source, [Parameter(Mandatory)] [string] $Destination)

    $repo = Get-PropertyValue $Source 'repository'; Test-RepositoryName $repo
    $tag  = Get-PropertyValue $Source 'tag'
    $want = Get-PropertyValue $Source 'asset'
    if (-not $want) { throw "githubRelease source for $repo has no 'asset'" }

    $release = if (-not $tag -or $tag -eq 'latest') { Invoke-GitHubApi "repos/$repo/releases/latest" }
               else { Invoke-GitHubApi "repos/$repo/releases/tags/$([uri]::EscapeDataString($tag))" }

    $assets = @($release.assets | Where-Object { $_.name -like $want })
    if ($assets.Count -eq 0) {
        throw "Release $($release.tag_name) of $repo has no asset matching '$want' (it has: $((@($release.assets) | ForEach-Object { $_.name }) -join ', '))"
    }
    if ($assets.Count -gt 1) {
        throw "Release $($release.tag_name) of $repo has $($assets.Count) assets matching '$want' ($(($assets | ForEach-Object { $_.name }) -join ', ')) - make the pattern specific"
    }

    $file = Join-Path $Destination $assets[0].name
    Invoke-Download -Uri $assets[0].url -Path $file -Accept 'application/octet-stream'
    Write-Ok "Downloaded $($assets[0].name) from release $($release.tag_name) of $repo"
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

    $repo = Get-PropertyValue $Source 'repository'; Test-RepositoryName $repo
    $path = ([string] (Get-PropertyValue $Source 'path')).Trim('/').Replace('\', '/')
    if (-not $path) { throw "github source for $repo has no 'path'" }

    $ref = Get-PropertyValue $Source 'ref'
    if (-not $ref) { $ref = (Invoke-GitHubApi "repos/$repo").default_branch }

    $tree  = Invoke-GitHubApi "repos/$repo/git/trees/$([uri]::EscapeDataString($ref))?recursive=1"
    if (Get-PropertyValue $tree 'truncated') {
        Write-Warn "The tree of $repo is too large for one listing - files beyond GitHub's limit may be missing"
    }
    $files = @($tree.tree | Where-Object { $_.type -eq 'blob' -and ($_.path -eq $path -or $_.path -like "$path/*") })
    if ($files.Count -eq 0) { throw "Nothing at '$path' in $repo ($ref)" }

    # A single file lands in the destination as is; a folder keeps its layout
    # under it, because an INF refers to its files by relative path.
    $prefix = if ($files.Count -eq 1 -and $files[0].path -eq $path) { ($path -replace '[^/]+$', '') } else { "$path/" }
    foreach ($item in $files) {
        $relative = $item.path.Substring($prefix.Length)
        $target   = Join-Path $Destination ($relative -replace '/', '\')
        $escaped  = ($item.path -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
        Invoke-Download -Uri "https://api.github.com/repos/$repo/contents/$escaped`?ref=$([uri]::EscapeDataString($ref))" `
                        -Path $target -Accept 'application/vnd.github.raw'
    }
    Write-Ok "Downloaded $($files.Count) file(s) from $repo/$path ($ref)"

    if ($files.Count -eq 1) { return (Join-Path $Destination ($files[0].path.Substring($prefix.Length) -replace '/', '\')) }
    return $Destination
}

function Test-FileHash {
    param([Parameter(Mandatory)] [string] $Path, [string] $Expected)

    if (-not $Expected) { return }
    if (Test-Path $Path -PathType Container) {
        Write-Warn "sha256 is set but the source is a folder - a hash only applies to a single file, so it was not checked"
        return
    }
    $actual = (Get-FileHash -Path $Path -Algorithm SHA256).Hash
    if ($actual -ne $Expected.ToUpper()) {
        throw "SHA256 of $(Split-Path $Path -Leaf) is $actual, the JSON expects $($Expected.ToUpper()) - refusing to install it"
    }
    Write-Ok "SHA256 matches ($actual)"
}

function Save-DriverSource {
    <#
        Bring one driver's files onto this machine and return the folder holding
        them, extracted. Every download goes into its own folder, emptied first, so
        files left from last month's version can never end up in this install.
    #>
    param([Parameter(Mandatory)] $Driver, [Parameter(Mandatory)] [string] $ConfigFolder)

    $source = Get-PropertyValue $Driver 'source'
    if (-not $source) { throw "Driver '$($Driver.name)' has no 'source'" }
    $type = [string] (Get-PropertyValue $source 'type')

    $safe   = ($Driver.name -replace '[^\w.-]+', '_').Trim('_')
    $folder = Join-Path $WorkingDir $safe
    if (Test-Path $folder) { Remove-Item -LiteralPath $folder -Recurse -Force }
    $download = Join-Path $folder 'download'
    New-Item -ItemType Directory -Path $download -Force | Out-Null

    $fetched = switch ($type) {
        'githubRelease' { Save-GitHubReleaseAsset -Source $source -Destination $download }
        'github'        { Save-GitHubPath -Source $source -Destination $download }
        'url' {
            $url  = Get-PropertyValue $source 'url'
            $name = Split-Path ([uri] $url).AbsolutePath -Leaf
            if (-not $name) { $name = 'driver.zip' }
            $file = Join-Path $download ([uri]::UnescapeDataString($name))
            Invoke-Download -Uri $url -Path $file
            Write-Ok "Downloaded $name"
            $file
        }
        'path' {
            $local = [string] (Get-PropertyValue $source 'path')
            if (-not [IO.Path]::IsPathRooted($local)) {
                if (-not $ConfigFolder) { throw "Relative path '$local' needs a JSON file on disk to resolve against, not a URL" }
                $local = Join-Path $ConfigFolder $local
            }
            if (-not (Test-Path $local)) { throw "Driver path not found: $local" }
            Copy-Item -Path $local -Destination $download -Recurse -Force
            Join-Path $download (Split-Path $local -Leaf)
        }
        default { throw "Driver '$($Driver.name)' has an unknown source type '$type' (githubRelease, github, url or path)" }
    }

    Test-FileHash -Path $fetched -Expected (Get-PropertyValue $source 'sha256')

    if ((Test-Path $fetched -PathType Leaf) -and ([IO.Path]::GetExtension($fetched) -eq '.zip')) {
        $extracted = Join-Path $folder 'extracted'
        Expand-Archive -LiteralPath $fetched -DestinationPath $extracted -Force
        Write-Ok "Extracted $(Split-Path $fetched -Leaf)"
        return $extracted
    }
    if ((Test-Path $fetched -PathType Leaf) -and ([IO.Path]::GetExtension($fetched) -eq '.cab')) {
        $extracted = Join-Path $folder 'extracted'
        New-Item -ItemType Directory -Path $extracted -Force | Out-Null
        & "$env:WINDIR\System32\expand.exe" -F:* $fetched $extracted | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "expand.exe could not extract $(Split-Path $fetched -Leaf) (exit code $LASTEXITCODE)" }
        Write-Ok "Extracted $(Split-Path $fetched -Leaf)"
        return $extracted
    }
    if (Test-Path $fetched -PathType Leaf) { return (Split-Path $fetched -Parent) }
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
    param([Parameter(Mandatory)] [string] $Name)

    $driver = @(Get-PrinterDriver -Name $Name -ErrorAction SilentlyContinue) | Select-Object -First 1
    if (-not $driver) { return $null }
    return [PSCustomObject]@{
        Name    = $driver.Name
        Version = ConvertFrom-DriverVersion (Get-PropertyValue $driver 'DriverVersion')
        Inf     = Get-PropertyValue $driver 'InfPath'
    }
}

function Find-DriverInf {
    <#
        The INF to install. Named in the JSON, it is taken as given (wildcards
        allowed). Otherwise the INF that contains the driver's name in quotes is the
        one - a universal driver package carries several INFs, and only the one that
        declares this model installs it. With more than one candidate the folder
        matching this machine's architecture wins.
    #>
    param([Parameter(Mandatory)] [string] $Folder, [Parameter(Mandatory)] $Driver)

    $infName = Get-PropertyValue $Driver 'inf'
    $all     = @(Get-ChildItem -Path $Folder -Recurse -File -Filter '*.inf' -ErrorAction SilentlyContinue)
    if ($all.Count -eq 0) { throw "No .inf file in the files for '$($Driver.name)' - is it an installer (.exe) rather than a driver package?" }

    $candidates = if ($infName) {
        $pattern = $infName.Replace('/', '\')
        @($all | Where-Object { $_.Name -like $pattern -or $_.FullName -like "*\$pattern" })
    } else {
        # Get-Content detects the BOM, so this reads the UTF-16 INFs many vendors
        # ship as well as the ANSI ones; a byte-level search would miss the first.
        $quoted = '"' + $Driver.name + '"'
        @($all | Where-Object { (Get-Content -LiteralPath $_.FullName -Raw) -like "*$([WildcardPattern]::Escape($quoted))*" })
    }
    $candidates = @($candidates | Where-Object { $_ })
    if ($candidates.Count -eq 0) {
        $what = if ($infName) { "named '$infName'" } else { "mentioning `"$($Driver.name)`"" }
        throw "No INF $what among: $(($all | ForEach-Object { $_.Name }) -join ', ')"
    }
    if ($candidates.Count -eq 1) { return $candidates[0] }

    $arch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'arm64' } else { 'x64|amd64|win64|64' }
    $preferred = @($candidates | Where-Object { $_.DirectoryName -match "(?i)\\($arch)(\\|$)" -or $_.Name -match "(?i)($arch)" })
    if ($preferred.Count -ge 1) { return $preferred[0] }

    Write-Warn "Several INFs match ($(($candidates | ForEach-Object { $_.Name }) -join ', ')) - using $($candidates[0].Name); set 'inf' in the JSON to choose"
    return $candidates[0]
}

function Test-DriverCatalog {
    <#
        The catalog named in the INF is what Windows checks the driver against. A
        package whose catalog is unsigned, or signed by somebody unexpected, is
        better refused here with a clear message than by pnputil with a code.
    #>
    param([Parameter(Mandatory)] $Inf)

    $text    = Get-Content -LiteralPath $Inf.FullName -Raw
    $catName = if ($text -match '(?im)^\s*CatalogFile(?:\.NTAMD64|\.NTARM64)?\s*=\s*([^\s;]+)') { $Matches[1] } else { $null }
    if (-not $catName) { throw "$($Inf.Name) names no CatalogFile - an unsigned driver package, which Windows will not install" }

    $cat = Join-Path $Inf.DirectoryName $catName
    if (-not (Test-Path $cat)) { throw "$($Inf.Name) refers to $catName, which is not in the package" }

    if ($SkipSignatureCheck) {
        Write-Warn "Signature check skipped for $catName (-SkipSignatureCheck)"
        return
    }
    $signature = Get-AuthenticodeSignature -FilePath $cat
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
    $output  = & $pnputil /add-driver $Inf.FullName /install 2>&1
    $code    = $LASTEXITCODE
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

    try {
        Add-PrinterDriver -Name $DriverName -ErrorAction Stop
    } catch {
        # A driver store entry the spooler does not pick up by name is installed
        # from the INF in the store directly - same package, explicit path.
        $published = @(Get-WindowsDriver -Online -ErrorAction SilentlyContinue |
                       Where-Object { $_.OriginalFileName -like "*\$($Inf.Name)" } |
                       Sort-Object { [version] $_.Version } -Descending) | Select-Object -First 1
        if (-not $published) { throw "Add-PrinterDriver '$DriverName' failed: $($_.Exception.Message)" }
        Add-PrinterDriver -Name $DriverName -InfPath $published.OriginalFileName -ErrorAction Stop
    }
}

# -- Printers ------------------------------------------------------------------

function Get-PortName {
    param([Parameter(Mandatory)] $Spec)
    $name = Get-PropertyValue $Spec 'portName'
    if ($name) { return $name }
    # An "absent" entry needs no address; without one there is no port to name.
    return "IP_$(Get-PropertyValue $Spec 'address')"
}

function Test-PrinterAbsent {
    param([Parameter(Mandatory)] $Spec)
    return ((Get-PropertyValue $Spec 'ensure') -eq 'absent')
}

function Remove-PrinterFromSpec {
    <#
        Take a printer away. The driver stays: other printers may use it, and a
        driver in use cannot be removed anyway. The port goes only when this
        printer was the last one on it.
    #>
    param([Parameter(Mandatory)] $Spec, [Parameter(Mandatory)] $State)

    if ($PSCmdlet.ShouldProcess($Spec.name, 'Remove-Printer')) {
        Remove-Printer -Name $Spec.name -ErrorAction Stop
        Write-Ok "Printer '$($Spec.name)' removed"
    }
    $portName = $State.PortName
    $others   = @(Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.PortName -eq $portName -and $_.Name -ne $Spec.name })
    if ((Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue) -and $others.Count -eq 0 -and
        $PSCmdlet.ShouldProcess($portName, 'Remove-PrinterPort (no other printer uses it)')) {
        try {
            Remove-PrinterPort -Name $portName -ErrorAction Stop
            Write-Ok "Port $portName removed"
        } catch {
            # The spooler can hold on to a port for a moment after its printer is
            # gone; an unused port is harmless and the next run removes it.
            Write-Warn "Port $portName is still held by the spooler - it goes on a later run"
        }
    }
}

function Get-PrinterState {
    <# What differs between the printer as it is and as the JSON wants it. #>
    param([Parameter(Mandatory)] $Spec)

    $portName = Get-PortName $Spec
    $existing = Get-Printer -Name $Spec.name -ErrorAction SilentlyContinue
    $port     = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
    $changes  = [System.Collections.Generic.List[string]]::new()

    if (Test-PrinterAbsent $Spec) {
        if ($existing) { $changes.Add('printer exists but the JSON says "ensure": "absent"') }
        return [PSCustomObject]@{ Exists = [bool] $existing; PortExists = [bool] $port; PortName = $portName; Changes = @($changes) }
    }

    if (-not $port) { $changes.Add("port $portName -> $($Spec.address)") }
    elseif ((Get-PropertyValue $port 'PrinterHostAddress') -and $port.PrinterHostAddress -ne $Spec.address) {
        $changes.Add("port $portName points at $($port.PrinterHostAddress), not $($Spec.address)")
    }

    if (-not $existing) {
        $changes.Add('printer does not exist')
    } else {
        if ($existing.DriverName -ne $Spec.driver) { $changes.Add("driver is '$($existing.DriverName)'") }
        if ($existing.PortName -ne $portName)      { $changes.Add("port is '$($existing.PortName)'") }
        foreach ($field in 'location', 'comment') {
            $want = Get-PropertyValue $Spec $field
            $have = Get-PropertyValue $existing ($field.Substring(0, 1).ToUpper() + $field.Substring(1))
            if ($null -ne $want -and $want -ne $have) { $changes.Add("$field is '$have'") }
        }
        $shared = Get-PropertyValue $Spec 'shared'
        if ($null -ne $shared -and [bool] $shared -ne [bool] $existing.Shared) { $changes.Add("shared is $($existing.Shared)") }

        $config = Get-PrintConfiguration -PrinterName $Spec.name -ErrorAction SilentlyContinue
        if ($config) {
            $duplex = Get-PropertyValue $Spec 'duplex'
            $color  = Get-PropertyValue $Spec 'color'
            $paper  = Get-PropertyValue $Spec 'paperSize'
            if ($duplex -and "$($config.DuplexingMode)" -ne $duplex)               { $changes.Add("duplex is $($config.DuplexingMode)") }
            if ($null -ne $color -and [bool] $color -ne [bool] $config.Color)      { $changes.Add("color is $($config.Color)") }
            if ($paper -and "$($config.PaperSize)" -ne $paper)                     { $changes.Add("paper size is $($config.PaperSize)") }
        }
    }

    return [PSCustomObject]@{ Exists = [bool] $existing; PortExists = [bool] $port; PortName = $portName; Changes = @($changes) }
}

function Set-PrinterFromSpec {
    param([Parameter(Mandatory)] $Spec, [Parameter(Mandatory)] $State)

    $portName = $State.PortName
    if (-not $State.PortExists) {
        if ($PSCmdlet.ShouldProcess("$portName -> $($Spec.address)", 'Add-PrinterPort')) {
            $lprQueue = Get-PropertyValue $Spec 'lprQueue'
            if ($lprQueue) {
                Add-PrinterPort -Name $portName -LprHostAddress $Spec.address -LprQueueName $lprQueue -ErrorAction Stop
            } else {
                $number = Get-PropertyValue $Spec 'portNumber'
                if (-not $number) { $number = 9100 }
                # Without -SNMP the port is created with SNMP off - unlike the GUI,
                # which turns it on. Deliberate: a printer that does not answer SNMP
                # with community "public" otherwise shows as Offline to every user.
                $snmp = @{}
                if ((Get-PropertyValue $Spec 'snmp') -eq $true) { $snmp = @{ SNMP = 1; SNMPCommunity = 'public' } }
                Add-PrinterPort -Name $portName -PrinterHostAddress $Spec.address -PortNumber $number @snmp -ErrorAction Stop
            }
            Write-Ok "Port $portName created ($($Spec.address))"
        }
    } elseif ($State.Changes -match '^port .* points at') {
        # There is no Set-PrinterPort. Re-creating the port means taking it away
        # from every printer on it first, which is not this printer's call to make.
        Write-Warn "Port $portName already exists for another address - give this printer its own 'portName' in the JSON"
    }

    $extra = @{}
    foreach ($field in 'location', 'comment') {
        $value = Get-PropertyValue $Spec $field
        if ($null -ne $value) { $extra[$field.Substring(0, 1).ToUpper() + $field.Substring(1)] = [string] $value }
    }
    $shared = Get-PropertyValue $Spec 'shared'
    if ($null -ne $shared) {
        $extra['Shared'] = [bool] $shared
        if ($shared) { $extra['ShareName'] = if (Get-PropertyValue $Spec 'shareName') { $Spec.shareName } else { $Spec.name } }
    }

    if (-not $State.Exists) {
        if ($PSCmdlet.ShouldProcess($Spec.name, "Add-Printer (driver '$($Spec.driver)', port $portName)")) {
            Add-Printer -Name $Spec.name -DriverName $Spec.driver -PortName $portName @extra -ErrorAction Stop
            Write-Ok "Printer '$($Spec.name)' added"
        }
    } elseif ($PSCmdlet.ShouldProcess($Spec.name, "Set-Printer ($($State.Changes -join '; '))")) {
        Set-Printer -Name $Spec.name -DriverName $Spec.driver -PortName $portName @extra -ErrorAction Stop
        Write-Ok "Printer '$($Spec.name)' corrected"
    }

    # Print defaults for every user of this printer. Only what the JSON sets is
    # touched; what the driver does not support is reported rather than fatal.
    $config = @{}
    if (Get-PropertyValue $Spec 'duplex')        { $config['DuplexingMode'] = $Spec.duplex }
    if ($null -ne (Get-PropertyValue $Spec 'color')) { $config['Color'] = [bool] $Spec.color }
    if (Get-PropertyValue $Spec 'paperSize')     { $config['PaperSize'] = $Spec.paperSize }
    if ($config.Count -gt 0 -and $PSCmdlet.ShouldProcess($Spec.name, "Set-PrintConfiguration ($(($config.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', '))")) {
        # Set-PrintConfiguration runs the vendor driver's own configuration code,
        # and some universal drivers hang in it. In a job it gets two minutes; a
        # hang there costs the print defaults, not the whole provisioning run.
        $job = Start-Job -ScriptBlock {
            param($Name, $Settings)
            Set-PrintConfiguration -PrinterName $Name @Settings -ErrorAction Stop
        } -ArgumentList $Spec.name, $config
        try {
            if (-not (Wait-Job -Job $job -Timeout 120)) {
                Write-Warn "Setting the print defaults for '$($Spec.name)' did not finish within 2 minutes - the driver hangs on it; the printer itself is installed"
            } else {
                Receive-Job -Job $job -ErrorAction Stop | Out-Null
                Write-Ok "Print defaults set for '$($Spec.name)'"
            }
        } catch {
            Write-Warn "Could not set the print defaults for '$($Spec.name)': $($_.Exception.Message)"
        } finally {
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }
    }
}

# -- Run -----------------------------------------------------------------------
try {
    Write-Out ''
    Write-Out "  Mode: $(if ($CheckOnly) { 'CHECK ONLY - reporting, never installing' } elseif ($simulate) { '-WhatIf - nothing will be changed' } else { 'APPLY - drivers and printers are installed where needed' })" 'Cyan'
    Write-Out ''

    # -- 1. Config -------------------------------------------------------------
    Write-Step '1. Configuration'
    $configFolder = $null
    if ($ConfigPath -match '^https?://') {
        if ($ConfigPath -notmatch '^https://') { throw "ConfigPath URL must be https: $ConfigPath" }
        $configFile = Join-Path ([IO.Path]::GetTempPath()) ("printers_{0}.json" -f [guid]::NewGuid().ToString('N'))
        try {
            Invoke-Download -Uri $ConfigPath -Path $configFile -Accept 'application/vnd.github.raw'
            $raw = Get-Content -LiteralPath $configFile -Raw -Encoding UTF8
        } finally {
            Remove-Item -LiteralPath $configFile -Force -ErrorAction SilentlyContinue
        }
    } else {
        if (-not (Test-Path $ConfigPath)) { throw "Configuration not found: $ConfigPath" }
        $raw          = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8
        $configFolder = Split-Path (Resolve-Path $ConfigPath).ProviderPath -Parent
    }
    try { $config = $raw | ConvertFrom-Json } catch { throw "$ConfigPath is not valid JSON: $($_.Exception.Message)" }

    $allDrivers  = @(@(Get-PropertyValue $config 'drivers')  | Where-Object { $_ })
    $allPrinters = @(@(Get-PropertyValue $config 'printers') | Where-Object { $_ })
    Write-Ok "$ConfigPath - $(@($allDrivers).Count) driver(s), $(@($allPrinters).Count) printer(s)"

    # Validate everything before anything is downloaded: a typo in printer 7 should
    # not surface after printer 1 to 6 have been installed.
    $problems = [System.Collections.Generic.List[string]]::new()
    foreach ($d in $allDrivers) {
        if (-not (Get-PropertyValue $d 'name'))   { $problems.Add('a driver without a name') }
        if (-not (Get-PropertyValue $d 'source')) { $problems.Add("driver '$(Get-PropertyValue $d 'name')' has no source") }
    }
    foreach ($p in $allPrinters) {
        $required = if (Test-PrinterAbsent $p) { @('name') } else { @('name', 'driver', 'address') }
        foreach ($field in $required) {
            if (-not (Get-PropertyValue $p $field)) { $problems.Add("printer '$(Get-PropertyValue $p 'name')' has no '$field'") }
        }
    }
    if ($problems.Count -gt 0) { throw "The configuration is incomplete: $($problems -join '; ')" }

    # The @() goes around the whole if: one match inside the branch alone comes
    # out as a bare object, and its .Count throws under Set-StrictMode.
    $printers = @(if ($Printer) {
        $allPrinters | Where-Object { $name = $_.name; @($Printer | Where-Object { $name -like $_ }).Count -gt 0 }
    } else { $allPrinters })
    if ($Printer -and $printers.Count -eq 0) {
        throw "No printer in the JSON matches $($Printer -join ', ') (it has: $(($allPrinters | ForEach-Object { $_.name }) -join ', '))"
    }

    $neededNames = @($printers | Where-Object { -not (Test-PrinterAbsent $_) } | ForEach-Object { $_.driver } | Select-Object -Unique)
    $drivers     = @($allDrivers | Where-Object { $_.name -in $neededNames -or (Get-PropertyValue $_ 'install') -eq $true })
    foreach ($name in $neededNames) {
        if ($name -notin @($allDrivers | ForEach-Object { $_.name }) -and -not (Get-InstalledPrinterDriver -Name $name)) {
            throw "Printer driver '$name' is neither in the JSON's drivers nor installed on this machine"
        }
    }

    # -- 2. Preflight ----------------------------------------------------------
    Write-Out ''
    Write-Step '2. Preflight'
    Wait-PrintSpooler

    $driverWork = [System.Collections.Generic.List[object]]::new()
    foreach ($driver in $drivers) {
        $installed = Get-InstalledPrinterDriver -Name $driver.name
        $wanted    = Get-PropertyValue $driver 'version'
        if (-not $installed) {
            Write-News "Driver '$($driver.name)' is not installed"
            $driverWork.Add($driver)
        } elseif ($wanted -and $installed.Version -lt [version] $wanted) {
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
    if (-not $simulate) {
        if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }
        $logFile = Join-Path $LogPath ("Install-Printer_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        Start-Transcript -Path $logFile | Out-Null
        $transcribing = $true
    }

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
    foreach ($driver in $driverWork) {
        Write-Out ''
        Write-Step "3. Driver '$($driver.name)'"
        if (-not $PSCmdlet.ShouldProcess($driver.name, "Download from $((Get-PropertyValue $driver.source 'type')) source and install")) { continue }
        try {
            $folder = Save-DriverSource -Driver $driver -ConfigFolder $configFolder
            $inf    = Find-DriverInf -Folder $folder -Driver $driver
            Write-Ok "INF: $($inf.FullName.Substring($folder.Length).TrimStart('\'))"
            Test-DriverCatalog -Inf $inf
            Install-DriverPackage -Inf $inf -DriverName $driver.name

            $now = Get-InstalledPrinterDriver -Name $driver.name
            if (-not $now) { throw "pnputil and Add-PrinterDriver reported success, yet '$($driver.name)' is not installed - is that the exact name inside the INF?" }
            Write-Ok "Driver '$($driver.name)' $($now.Version) installed"
        } catch {
            $failedDrivers.Add($driver.name)
            Write-Bad "Driver '$($driver.name)': $($_.Exception.Message)"
            $exitCode = 1
        }
    }

    # -- 5. Printers -----------------------------------------------------------
    Write-Out ''
    Write-Step '5. Printers'
    if ($printerWork.Count -eq 0) { Write-Skip 'Every printer is already as configured' }
    foreach ($item in $printerWork) {
        $spec = $item.Spec
        if (Test-PrinterAbsent $spec) {
            try { Remove-PrinterFromSpec -Spec $spec -State $item.State }
            catch { Write-Bad "Printer '$($spec.name)': $($_.Exception.Message)"; $exitCode = 1 }
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
        try {
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
            if ((Test-PrinterAbsent $spec) -and $state.Exists) { Write-Bad "Printer '$($spec.name)' is still there"; $exitCode = 1 }
            elseif (Test-PrinterAbsent $spec) { Write-Ok "Printer '$($spec.name)' is gone" }
            elseif ($state.Changes.Count -eq 0) { Write-Ok "Printer '$($spec.name)' is as configured" }
            elseif ($state.Exists)          { Write-Warn "Printer '$($spec.name)' still differs: $($state.Changes -join '; ')" }
            else                            { Write-Bad "Printer '$($spec.name)' is not installed"; $exitCode = 1 }
        }
        Write-Out ''
        if ($rebootRequired) { Write-Warn 'pnputil asked for a reboot to finish the driver installation' }
        if ($exitCode -eq 0) { Write-Ok 'Done' }
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
