#Requires -Version 5.1
<#
.SYNOPSIS
    Test file I/O on a path, diagnose failures, and monitor the directory in real time.

.DESCRIPTION
    Runs a write / append / read / delete loop on a target path and on every failure:

      - Classifies the error: AUTH, NETWORK, TIMEOUT, DISK, PATH, or IO
      - Captures FileSystemWatcher events leading up to the failure
      - Compares NTFS permissions (icacls) against the baseline taken at startup
      - Lists processes with an open handle on the directory via Sysinternals Handle.exe
        (downloaded automatically to C:\Temp\handle\ if not already present)
      - Snapshots new processes that appeared since the script started
      - Checks the Windows Security event log for auth failures and object access events
      - Stops after 3 failures to preserve the diagnostic data

    Sysinternals Handle.exe is downloaded automatically from Microsoft unless
    -SkipHandleDownload is specified or the download fails.

.PARAMETER TestPath
    Folder to test. E.g. G:\sas\work  or  \\server\share\folder

.PARAMETER Iterations
    Number of write/delete cycles. Default: 5000.

.PARAMETER DelayMs
    Milliseconds between iterations. Default: 50.

.PARAMETER StopOnFirstError
    Stop after the first failure instead of continuing to 3.

.PARAMETER HandleExe
    Path to an existing Handle.exe. If not set the script downloads it automatically.

.PARAMETER SkipHandleDownload
    Do not attempt to download Handle.exe (e.g. no internet access on the server).

.PARAMETER LogPath
    Output folder for the log file. Default: C:\Temp\

.EXAMPLE
    # Basic run — downloads Handle.exe automatically
    .\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work

.EXAMPLE
    # Fast stress test, stop on first failure
    .\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work -Iterations 10000 -DelayMs 0 -StopOnFirstError

.EXAMPLE
    # Use an existing Handle.exe, no download
    .\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work -HandleExe C:\Tools\handle.exe

.EXAMPLE
    # Air-gapped server — skip the download
    .\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work -SkipHandleDownload
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string] $TestPath,

    [int]    $Iterations         = 5000,
    [int]    $DelayMs            = 50,
    [switch] $StopOnFirstError,
    [string] $HandleExe          = '',
    [switch] $SkipHandleDownload,
    [string] $LogPath            = 'C:\Temp'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'SilentlyContinue'

# ── Log setup ──────────────────────────────────────────────────────────────────
if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath | Out-Null }
$ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile = Join-Path $LogPath "FileIO_Diagnostics_$ts.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line  = '[{0}] [{1,-5}] {2}' -f (Get-Date -Format 'HH:mm:ss.fff'), $Level, $Message
    $color = switch ($Level) {
        'OK'    { 'Green'   }
        'WARN'  { 'Yellow'  }
        'FAIL'  { 'Red'     }
        'ALERT' { 'Magenta' }
        'AUTH'  { 'Magenta' }
        default { 'Gray'    }
    }
    Write-Host $line -ForegroundColor $color
    Add-Content -Path $logFile -Value $line -Encoding UTF8
}

function Write-Section {
    param([string]$Title)
    $sep   = '─' * 56
    $block = "`n  $sep`n   $Title`n  $sep"
    Write-Host $block -ForegroundColor Cyan
    Add-Content -Path $logFile -Value $block -Encoding UTF8
}

# ── Error classifier ───────────────────────────────────────────────────────────
function Get-ErrorCategory {
    param([System.Management.Automation.ErrorRecord]$Err)

    $msg       = $Err.Exception.Message
    $type      = $Err.Exception.GetType().Name
    $hres      = $Err.Exception.HResult
    $innerType = if ($Err.Exception.InnerException) { $Err.Exception.InnerException.GetType().Name } else { '' }
    $innerHRes = if ($Err.Exception.InnerException) { $Err.Exception.InnerException.HResult } else { 0 }

    if ($type -eq 'UnauthorizedAccessException' -or $innerType -eq 'UnauthorizedAccessException') { return 'AUTH' }
    if ($msg -imatch 'access.*denied|is denied|logon.?fail|not have.?access|credentials|password|privilege') { return 'AUTH' }
    if ($hres -in @(0x80070005, 0x8007052e, 0x8007052f, 0x80070569) -or
        $innerHRes -in @(0x80070005, 0x8007052e, 0x8007052f, 0x80070569)) { return 'AUTH' }

    if ($msg -imatch 'network.?path|network.?name|could not find|server.?not.?found|host.?unreachable') { return 'NETWORK' }
    if ($hres -in @(0x80070035, 0x80070040, 0x800704cf, 0x80070043) -or
        $innerHRes -in @(0x80070035, 0x80070040, 0x800704cf, 0x80070043)) { return 'NETWORK' }

    if ($msg -imatch 'timed? out|timeout|semaphore') { return 'TIMEOUT' }
    if ($hres -in @(0x800705b4, 0x80070079) -or
        $innerHRes -in @(0x800705b4, 0x80070079)) { return 'TIMEOUT' }

    if ($msg -imatch 'disk.?full|not enough.?space|quota|no space') { return 'DISK' }
    if ($hres -in @(0x80070070, 0x80070522) -or
        $innerHRes -in @(0x80070070, 0x80070522)) { return 'DISK' }

    if ($msg -imatch 'path.?not.?found|file.?not.?found|directory.?not') { return 'PATH' }
    if ($hres -in @(0x80070002, 0x80070003) -or
        $innerHRes -in @(0x80070002, 0x80070003)) { return 'PATH' }

    return 'IO'
}

# ── Handle.exe — auto-download ────────────────────────────────────────────────
function Get-HandleExe {
    param([string]$ExplicitPath, [switch]$Skip)

    if ($ExplicitPath -and (Test-Path $ExplicitPath)) {
        Write-Log "Handle.exe  : $ExplicitPath" 'OK'
        return $ExplicitPath
    }

    # Check default install location
    $defaultPath = 'C:\Temp\handle\handle64.exe'
    $fallback    = 'C:\Temp\handle\handle.exe'
    foreach ($p in @($defaultPath, $fallback)) {
        if (Test-Path $p) {
            Write-Log "Handle.exe  : $p (cached)" 'OK'
            return $p
        }
    }

    if ($Skip) {
        Write-Log "Handle.exe  : skipped (-SkipHandleDownload)" 'WARN'
        return $null
    }

    # Download from Sysinternals
    Write-Log "Handle.exe  : not found — downloading from Sysinternals..." 'WARN'
    $zipUrl   = 'https://download.sysinternals.com/files/Handle.zip'
    $zipDest  = 'C:\Temp\handle\Handle.zip'
    $extractTo = 'C:\Temp\handle'

    try {
        if (-not (Test-Path $extractTo)) { New-Item -ItemType Directory -Path $extractTo | Out-Null }
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $zipUrl -OutFile $zipDest -UseBasicParsing -ErrorAction Stop
        Expand-Archive -Path $zipDest -DestinationPath $extractTo -Force -ErrorAction Stop
        Remove-Item $zipDest -Force -ErrorAction SilentlyContinue

        $found = @($defaultPath, $fallback) | Where-Object { Test-Path $_ } | Select-Object -First 1
        if ($found) {
            Write-Log "Handle.exe  : downloaded to $found" 'OK'
            return $found
        }
        Write-Log "Handle.exe  : download succeeded but exe not found in $extractTo" 'WARN'
        return $null
    } catch {
        Write-Log "Handle.exe  : download failed — $($_.Exception.Message)" 'WARN'
        Write-Log "             Process handle detection will be skipped" 'WARN'
        return $null
    }
}

# ── Diagnostic helpers ─────────────────────────────────────────────────────────
function Get-Icacls {
    param([string]$Path)
    try { (& icacls $Path 2>&1 | Out-String).Trim() } catch { "(icacls failed)" }
}

function Get-ProcessHandles {
    param([string]$Path, [string]$HandlePath)
    if (-not $HandlePath -or -not (Test-Path $HandlePath)) { return $null }
    try { (& $HandlePath $Path -nobanner -accepteula 2>&1 | Out-String).Trim() } catch { $null }
}

function Get-ProcessSnapshot {
    try {
        Get-Process | Select-Object Id, Name, StartTime |
            Sort-Object Name |
            ForEach-Object { '{0,6}  {1,-32}  started: {2}' -f $_.Id, $_.Name, $_.StartTime }
    } catch { @() }
}

# ── FileSystemWatcher ──────────────────────────────────────────────────────────
$fsEvents = [System.Collections.Generic.List[string]]::new()
$fsLock   = [System.Object]::new()

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path                  = $TestPath
$watcher.NotifyFilter          = [System.IO.NotifyFilters]'FileName,DirectoryName,Attributes,Security,LastWrite'
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents   = $true

$addFsEvent = {
    $e    = $Event.SourceEventArgs
    $verb = switch ($Event.SourceEventArgs.GetType().Name) {
        'RenamedEventArgs' { "RENAMED : $($e.OldFullPath) → $($e.FullPath)" }
        default            { "$($e.ChangeType.ToString().ToUpper().PadRight(8)): $($e.FullPath)" }
    }
    [System.Threading.Monitor]::Enter($fsLock)
    try { $fsEvents.Add("[$(Get-Date -Format 'HH:mm:ss.fff')] $verb") }
    finally { [System.Threading.Monitor]::Exit($fsLock) }
}

$evChanged = Register-ObjectEvent $watcher 'Changed' -Action $addFsEvent
$evCreated = Register-ObjectEvent $watcher 'Created' -Action $addFsEvent
$evDeleted = Register-ObjectEvent $watcher 'Deleted' -Action $addFsEvent
$evRenamed = Register-ObjectEvent $watcher 'Renamed' -Action $addFsEvent

function Get-AndClearFsEvents {
    [System.Threading.Monitor]::Enter($fsLock)
    try {
        $copy = @($fsEvents)
        $fsEvents.Clear()
        return $copy
    } finally { [System.Threading.Monitor]::Exit($fsLock) }
}

# ── Pre-flight ─────────────────────────────────────────────────────────────────
Write-Section 'Pre-flight checks'

# Drive/UNC detection
$uncServer = $null
if ($TestPath -match '^\\\\([^\\]+)') {
    $uncServer = $Matches[1]
    Write-Log "UNC server  : $uncServer"
} elseif ($TestPath -match '^([A-Za-z]):') {
    $driveLetter = $Matches[1].ToUpper()
    Write-Log "Drive letter: ${driveLetter}:"
    $mapping = net use 2>$null | Select-String "${driveLetter}:"
    if ($mapping) {
        Write-Log "Drive mapping: $mapping" 'INFO'
        if ($mapping -match '\\\\([^\\]+)') { $uncServer = $Matches[1] }
    } else {
        Write-Log "Drive ${driveLetter}: not in 'net use' — local drive or SAN" 'WARN'
    }
}

# Path check
if (Test-Path $TestPath) {
    Write-Log "Path exists : $TestPath" 'OK'
} else {
    Write-Log "Path not found — attempting to create: $TestPath" 'WARN'
    try {
        New-Item -ItemType Directory -Path $TestPath -Force | Out-Null
        Write-Log "Path created" 'OK'
    } catch {
        Write-Log "Cannot create path: $($_.Exception.Message)" 'FAIL'
        exit 1
    }
}

# SMB check for network paths
if ($uncServer) {
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $ar  = $tcp.BeginConnect($uncServer, 445, $null, $null)
        $ok  = $ar.AsyncWaitHandle.WaitOne(3000)
        $tcp.Close()
        $lvl = if ($ok) { 'OK' } else { 'FAIL' }
        Write-Log "SMB port 445 on '$uncServer': $(if ($ok) { 'OPEN' } else { 'NOT reachable' })" $lvl
    } catch {
        Write-Log "SMB port check failed: $($_.Exception.Message)" 'WARN'
    }
}

Write-Log "Running as  : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)" 'INFO'

# Kerberos
try {
    $klistOut    = & klist 2>&1 | Out-String
    $ticketCount = ([regex]::Matches($klistOut, 'Client:')).Count
    $lvl = if ($ticketCount -gt 0) { 'OK' } else { 'WARN' }
    Write-Log "Kerberos    : $ticketCount ticket(s) cached" $lvl
} catch {
    Write-Log "Kerberos    : klist unavailable" 'WARN'
}

# Handle.exe
$handleExePath = Get-HandleExe -ExplicitPath $HandleExe -Skip:$SkipHandleDownload

# Baseline NTFS permissions
$baselineAcl   = Get-Icacls $TestPath
Write-Log "Baseline NTFS permissions:" 'INFO'
$baselineAcl -split "`n" | ForEach-Object { Write-Log "  $_" 'INFO' }

# Baseline process list
$baselineProcs = Get-ProcessSnapshot

# ── Main I/O loop ──────────────────────────────────────────────────────────────
Write-Section "File I/O Test — $Iterations iterations on $TestPath"

$testFile   = Join-Path $TestPath "fileio_test_$pid.tmp"
$stats      = @{ OK = 0; Fail = 0 }
$failCount  = 0
$errorLog   = [System.Collections.Generic.List[object]]::new()
$maxFails   = if ($StopOnFirstError) { 1 } else { 3 }

try {
    for ($i = 1; $i -le $Iterations; $i++) {

        # Periodically report FS events (not failure context)
        if ($i % 200 -eq 0) {
            $pending = Get-AndClearFsEvents
            if ($pending.Count -gt 0) {
                Write-Log "FileSystemWatcher — $($pending.Count) event(s) in last 200 iterations:" 'WARN'
                $pending | ForEach-Object { Write-Log "  $_" 'WARN' }
            }
        }

        try {
            [System.IO.File]::WriteAllText($testFile, "Hello,`n")
            [System.IO.File]::AppendAllText($testFile, "`n")
            [System.IO.File]::AppendAllText($testFile, "world`n")
            $content = [System.IO.File]::ReadAllText($testFile)
            if ($content -notmatch 'Hello') { throw [System.IO.IOException]::new("Content verification failed") }
            [System.IO.File]::Delete($testFile)

            $stats.OK++
            if ($i % 500 -eq 0 -or $i -eq 1) {
                Write-Log "Iteration $i / $Iterations — OK ($($stats.OK) passed, $($stats.Fail) failed)" 'OK'
            }
        } catch {
            $stats.Fail++
            $failCount++
            $captured = $_
            $category = Get-ErrorCategory $captured

            $levelMap = @{ AUTH='AUTH'; NETWORK='FAIL'; TIMEOUT='WARN'; DISK='FAIL'; PATH='FAIL'; IO='FAIL' }
            Write-Log "Iteration $i — [$category] $($captured.Exception.Message)" $levelMap[$category]

            $errorLog.Add([PSCustomObject]@{
                Iteration = $i
                Category  = $category
                Message   = $captured.Exception.Message
                HResult   = '0x{0:X8}' -f $captured.Exception.HResult
                Time      = Get-Date
            })

            # ── Failure diagnostics ────────────────────────────────────────────
            Write-Section "Failure #$failCount diagnostics (iteration $i)"

            # 1. Cause explanation
            switch ($category) {
                'AUTH' {
                    Write-Log "CAUSE: Authentication / permission failure" 'AUTH'
                    Write-Log "  Common causes:" 'INFO'
                    Write-Log "    - A process temporarily reset NTFS permissions on the directory" 'INFO'
                    Write-Log "    - Kerberos ticket expired mid-run (check ticket end times below)" 'INFO'
                    Write-Log "    - SAS WORK directory cleanup between job phases" 'INFO'
                    Write-Log "    - Antivirus / EDR quarantine action" 'INFO'
                }
                'NETWORK' {
                    Write-Log "CAUSE: Network path unreachable" 'FAIL'
                    Write-Log "  Server offline, firewall, DNS failure, or VPN disconnected" 'INFO'
                    if ($uncServer) {
                        try {
                            $ips = [System.Net.Dns]::GetHostAddresses($uncServer) | ForEach-Object { $_.IPAddressToString }
                            Write-Log "  DNS resolves '$uncServer' to: $($ips -join ', ')" 'INFO'
                        } catch { Write-Log "  DNS resolution failed for '$uncServer'" 'FAIL' }
                    }
                }
                'TIMEOUT' {
                    Write-Log "CAUSE: Operation timed out — server load or network congestion" 'WARN'
                }
                'DISK' {
                    Write-Log "CAUSE: Disk full or quota exceeded" 'FAIL'
                    try {
                        $drv = Split-Path $TestPath -Qualifier
                        $d   = Get-PSDrive $drv.TrimEnd(':') -ErrorAction SilentlyContinue
                        if ($d) {
                            Write-Log ("  Drive {0}: Used = {1:N1} GB | Free = {2:N1} GB" -f $drv, ($d.Used/1GB), ($d.Free/1GB)) 'INFO'
                        }
                    } catch {}
                }
                'PATH' {
                    Write-Log "CAUSE: Directory disappeared during the test" 'FAIL'
                    Write-Log "  Share disconnected, drive remapped, or directory deleted by another process" 'INFO'
                }
                default {
                    Write-Log ("CAUSE: General I/O error — HResult {0}" -f ('0x{0:X8}' -f $captured.Exception.HResult)) 'FAIL'
                }
            }

            # 2. FileSystemWatcher events captured before the failure
            $fsNow = Get-AndClearFsEvents
            if ($fsNow.Count -gt 0) {
                Write-Log "FileSystemWatcher events leading up to failure:" 'ALERT'
                $fsNow | ForEach-Object { Write-Log "  $_" 'ALERT' }
            } else {
                Write-Log "FileSystemWatcher: no directory events captured before this failure" 'INFO'
            }

            # 3. NTFS permission diff
            $currentAcl = Get-Icacls $TestPath
            if ($currentAcl -ne $baselineAcl) {
                Write-Log "NTFS PERMISSIONS CHANGED since baseline!" 'ALERT'
                $currentAcl -split "`n" | ForEach-Object { Write-Log "  $_" 'ALERT' }
            } else {
                Write-Log "NTFS permissions: unchanged from baseline" 'INFO'
            }

            # 4. Process handles on the directory
            $handles = Get-ProcessHandles $TestPath $handleExePath
            if ($handles) {
                Write-Log "Processes with open handles on '$TestPath':" 'ALERT'
                $handles -split "`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log "  $_" 'ALERT' }
            } elseif (-not $handleExePath) {
                Write-Log "Process handles: Handle.exe not available (download failed or skipped)" 'WARN'
            }

            # 5. New processes since script start
            $currentProcs = Get-ProcessSnapshot
            $newProcs = $currentProcs | Where-Object { $_ -notin $baselineProcs }
            if ($newProcs) {
                Write-Log "New processes since script started:" 'ALERT'
                $newProcs | ForEach-Object { Write-Log "  $_" 'ALERT' }
            } else {
                Write-Log "No new processes since script started" 'INFO'
            }

            # 6. Kerberos tickets at failure time
            try {
                $klNow = & klist 2>&1 | Out-String
                $klCount = ([regex]::Matches($klNow, 'Client:')).Count
                Write-Log "Kerberos tickets at failure time: $klCount cached" 'INFO'
                # Show only ticket 0 (TGT) for brevity
                if ($klNow -match '(?s)(#0>.*?Kdc Called:[^\r\n]+)') {
                    $Matches[1] -split "`n" | ForEach-Object { Write-Log "  $_" 'INFO' }
                }
            } catch {}

            # 7. Security event log — auth failures + object access (last 30 seconds)
            try {
                $since = (Get-Date).AddSeconds(-30)
                $evts  = Get-WinEvent -FilterHashtable @{
                    LogName   = 'Security'
                    Id        = @(4625, 4740, 4656, 4663, 4670)
                    StartTime = $since
                } -MaxEvents 10 -ErrorAction Stop
                if ($evts -and $evts.Count -gt 0) {
                    Write-Log "$($evts.Count) Security event(s) in the last 30 seconds:" 'ALERT'
                    $evts | ForEach-Object {
                        Write-Log ("  [{0}] Event {1}: {2}" -f $_.TimeCreated.ToString('HH:mm:ss.fff'), $_.Id,
                            ($_.Message -replace '\s+', ' ').Substring(0, [Math]::Min(140, $_.Message.Length))) 'ALERT'
                    }
                } else {
                    Write-Log "Security log: no relevant events in the last 30 seconds" 'INFO'
                }
            } catch {
                Write-Log "Security log: could not read (run as admin for full event access)" 'WARN'
            }

            Write-Log "End of failure #$failCount diagnostics" 'INFO'

            # Clean up
            try { if (Test-Path $testFile) { Remove-Item $testFile -Force } } catch {}

            if ($failCount -ge $maxFails) {
                $stopReason = if ($StopOnFirstError) { "first error (-StopOnFirstError)" } else { "3 failures captured" }
                Write-Log "Stopping after $stopReason" 'WARN'
                break
            }
        }

        if ($DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
    }
} finally {
    $watcher.EnableRaisingEvents = $false
    Unregister-Event $evChanged.Id -ErrorAction SilentlyContinue
    Unregister-Event $evCreated.Id -ErrorAction SilentlyContinue
    Unregister-Event $evDeleted.Id -ErrorAction SilentlyContinue
    Unregister-Event $evRenamed.Id -ErrorAction SilentlyContinue
    $watcher.Dispose()
    if (Test-Path $testFile) { Remove-Item $testFile -Force -ErrorAction SilentlyContinue }
}

# ── Summary ────────────────────────────────────────────────────────────────────
Write-Section 'Summary'

$total = $stats.OK + $stats.Fail
$pct   = if ($total -gt 0) { [math]::Round($stats.OK / $total * 100, 1) } else { 0 }
Write-Log "Iterations  : $total" 'INFO'
$pLevel = if ($stats.Fail -eq 0) { 'OK' } else { 'INFO' }
$fLevel = if ($stats.Fail -eq 0) { 'OK' } else { 'FAIL' }
Write-Log "Passed      : $($stats.OK) ($pct%)" $pLevel
Write-Log "Failed      : $($stats.Fail)" $fLevel

if ($errorLog.Count -gt 0) {
    $grouped = $errorLog | Group-Object Category | Sort-Object Count -Descending
    Write-Log "Failure breakdown:" 'INFO'
    foreach ($g in $grouped) { Write-Log "  [$($g.Name)] $($g.Count) failure(s)" 'FAIL' }

    Write-Log "First failure : iteration $($errorLog[0].Iteration) at $($errorLog[0].Time.ToString('HH:mm:ss.fff'))" 'INFO'
    if ($errorLog.Count -gt 1) {
        Write-Log "Last failure  : iteration $($errorLog[-1].Iteration) at $($errorLog[-1].Time.ToString('HH:mm:ss.fff'))" 'INFO'
    }

    Write-Log '' 'INFO'
    Write-Log "What to look for in the ALERT lines above:" 'INFO'
    Write-Log "  - NTFS PERMISSIONS CHANGED → a process reset the folder ACL" 'INFO'
    Write-Log "  - FileSystemWatcher events → what file/folder operations occurred" 'INFO'
    Write-Log "  - Processes with handles  → which process locked the directory" 'INFO'
    Write-Log "  - New processes           → scheduler or cleanup job started before failure" 'INFO'

    $top = $grouped[0].Name
    Write-Log '' 'INFO'
    switch ($top) {
        'AUTH'    { Write-Log "Recommendation: check NTFS permission changes and process handles above — most likely a cleanup process or AV" 'WARN' }
        'NETWORK' { Write-Log "Recommendation: check server availability, SMB port 445, and DNS resolution" 'WARN' }
        'TIMEOUT' { Write-Log "Recommendation: check server/storage load and network latency" 'WARN' }
        'DISK'    { Write-Log "Recommendation: free up disk space or increase user quota" 'WARN' }
        'PATH'    { Write-Log "Recommendation: verify the share is still mounted and the directory exists" 'WARN' }
        default   { Write-Log "Recommendation: review the detailed log for specific error codes" 'WARN' }
    }
} else {
    Write-Log "No failures detected — issue may be timing-dependent" 'OK'
    Write-Log "Try with -DelayMs 0 and more iterations if the problem recurs" 'WARN'
}

Write-Log "Log saved to : $logFile" 'INFO'
Write-Host ''
