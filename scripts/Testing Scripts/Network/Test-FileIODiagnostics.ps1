#Requires -Version 5.1
<#
.SYNOPSIS
    Test file I/O on a path and diagnose the cause of any failures.

.DESCRIPTION
    Runs a write / append / read / delete loop on a target path and
    categorises every failure: authentication error, access denied,
    network unreachable, disk full, timeout, or other I/O error.

    On the first failure the script automatically collects extra
    diagnostics (drive mapping, Kerberos tickets, SMB port, event log)
    to help pinpoint the root cause.

.PARAMETER TestPath
    Folder to test. E.g. G:\sas\work  or  \\server\share\folder

.PARAMETER Iterations
    Number of write/delete cycles to run. Default: 100.

.PARAMETER DelayMs
    Milliseconds to wait between iterations. Default: 0.

.PARAMETER StopOnFirstError
    Stop the loop on the first failure instead of continuing.

.PARAMETER LogPath
    Folder for the output log file. Default: C:\Temp\

.EXAMPLE
    # Quick test of a mapped SAS work drive
    .\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work

.EXAMPLE
    # 1000 iterations, stop on first failure, detailed log
    .\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work -Iterations 1000 -StopOnFirstError

.EXAMPLE
    # Test a UNC path with a small delay between writes
    .\Test-FileIODiagnostics.ps1 -TestPath \\fileserver\data\testfolder -Iterations 500 -DelayMs 100
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string] $TestPath,

    [int]    $Iterations      = 100,
    [int]    $DelayMs         = 0,
    [switch] $StopOnFirstError,
    [string] $LogPath         = 'C:\Temp'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Log setup ──────────────────────────────────────────────────────────────────
if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath | Out-Null }
$ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile = Join-Path $LogPath "FileIO_Diagnostics_$ts.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line  = '[{0}] [{1,-5}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    $color = switch ($Level) {
        'OK'    { 'Green'   }
        'WARN'  { 'Yellow'  }
        'FAIL'  { 'Red'     }
        'HEAD'  { 'Cyan'    }
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
    # PowerShell wraps .NET exceptions in MethodInvocationException — check inner too
    $innerType = if ($Err.Exception.InnerException) { $Err.Exception.InnerException.GetType().Name } else { '' }
    $innerHRes = if ($Err.Exception.InnerException) { $Err.Exception.InnerException.HResult } else { 0 }

    # Auth / credential failures
    if ($type -eq 'UnauthorizedAccessException' -or $innerType -eq 'UnauthorizedAccessException') { return 'AUTH' }
    if ($msg -imatch 'access.*denied|is denied|logon.?fail|not have.?access|credentials|password|privilege') { return 'AUTH' }
    if ($hres -in @(0x80070005, 0x8007052e, 0x8007052f, 0x80070569))   { return 'AUTH'    }
    if ($innerHRes -in @(0x80070005, 0x8007052e, 0x8007052f, 0x80070569)) { return 'AUTH' }

    # Network path / server unreachable
    if ($msg -imatch 'network.?path|network.?name|could not find|server.?not.?found|host.?unreachable') { return 'NETWORK' }
    if ($hres -in @(0x80070035, 0x80070040, 0x800704cf, 0x80070043) -or
        $innerHRes -in @(0x80070035, 0x80070040, 0x800704cf, 0x80070043)) { return 'NETWORK' }

    # Timeout
    if ($msg -imatch 'timed? out|timeout|semaphore')                    { return 'TIMEOUT' }
    if ($hres -in @(0x800705b4, 0x80070079) -or
        $innerHRes -in @(0x800705b4, 0x80070079))                       { return 'TIMEOUT' }

    # Disk / quota
    if ($msg -imatch 'disk.?full|not enough.?space|quota|no space')     { return 'DISK'    }
    if ($hres -in @(0x80070070, 0x80070522) -or
        $innerHRes -in @(0x80070070, 0x80070522))                       { return 'DISK'    }

    # Path / file not found
    if ($msg -imatch 'path.?not.?found|file.?not.?found|directory.?not') { return 'PATH'  }
    if ($hres -in @(0x80070002, 0x80070003) -or
        $innerHRes -in @(0x80070002, 0x80070003))                       { return 'PATH'    }

    return 'IO'
}

# ── Pre-flight diagnostics ─────────────────────────────────────────────────────
Write-Section 'Pre-flight checks'

# Resolve UNC server or drive letter
$uncServer = $null
if ($TestPath -match '^\\\\([^\\]+)') {
    $uncServer = $Matches[1]
    Write-Log "UNC server  : $uncServer"
} elseif ($TestPath -match '^([A-Za-z]):') {
    $driveLetter = $Matches[1].ToUpper()
    Write-Log "Drive letter: ${driveLetter}:"
    $mapping = net use 2>$null | Where-Object { $_ -match "^.*\s+${driveLetter}:\s+(.+)\s+Microsoft" }
    if (-not $mapping) {
        $mapping = net use 2>$null | Select-String "${driveLetter}:"
    }
    if ($mapping) {
        Write-Log "Drive mapping: $mapping" 'INFO'
        if ($mapping -match '\\\\([^\\]+)') { $uncServer = $Matches[1] }
    } else {
        Write-Log "Drive ${driveLetter}: not found in 'net use' — may be a local drive" 'WARN'
    }
}

# Test folder exists
if (Test-Path $TestPath) {
    Write-Log "Path exists : $TestPath" 'OK'
} else {
    Write-Log "Path not found: $TestPath — will attempt to create it" 'WARN'
    try {
        New-Item -ItemType Directory -Path $TestPath -Force | Out-Null
        Write-Log "Path created: $TestPath" 'OK'
    } catch {
        Write-Log "Cannot create path: $($_.Exception.Message)" 'FAIL'
        Write-Log "Aborting — target path is inaccessible" 'FAIL'
        exit 1
    }
}

# SMB port check for UNC / mapped drives
if ($uncServer) {
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $ar  = $tcp.BeginConnect($uncServer, 445, $null, $null)
        $ok  = $ar.AsyncWaitHandle.WaitOne(3000)
        $tcp.Close()
        if ($ok) { Write-Log "SMB port 445 on '$uncServer': OPEN" 'OK' }
        else      { Write-Log "SMB port 445 on '$uncServer': NOT reachable" 'FAIL' }
    } catch {
        Write-Log "SMB port check failed: $($_.Exception.Message)" 'WARN'
    }

    $ping = Test-Connection -ComputerName $uncServer -Count 1 -Quiet -ErrorAction SilentlyContinue
    if ($ping) { Write-Log "Ping '$uncServer': OK" 'OK' }
    else        { Write-Log "Ping '$uncServer': failed (may be blocked)" 'WARN' }
}

# Current identity
Write-Log "Running as  : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)" 'INFO'

# Kerberos tickets
try {
    $klist = & klist 2>&1 | Out-String
    $ticketCount = ([regex]::Matches($klist, 'Client:')).Count
    if ($ticketCount -gt 0) {
        Write-Log "Kerberos    : $ticketCount ticket(s) cached" 'OK'
    } else {
        Write-Log "Kerberos    : no tickets cached — authentication may fail" 'WARN'
    }
} catch {
    Write-Log "Kerberos    : klist unavailable" 'WARN'
}

# ── Main I/O loop ──────────────────────────────────────────────────────────────
Write-Section "File I/O Test — $Iterations iterations on $TestPath"

$testFile  = Join-Path $TestPath "fileio_test_$pid.tmp"
$stats     = @{ OK = 0; Fail = 0 }
$errorLog  = [System.Collections.Generic.List[object]]::new()
$diagnosed = $false   # only collect deep diagnostics once

for ($i = 1; $i -le $Iterations; $i++) {

    $iterError = $null

    try {
        # Step 1 — create / overwrite
        [System.IO.File]::WriteAllText($testFile, "Hello,`n")

        # Step 2 — append
        [System.IO.File]::AppendAllText($testFile, "`n")

        # Step 3 — append more
        [System.IO.File]::AppendAllText($testFile, "world`n")

        # Step 4 — read back (verify write)
        $content = [System.IO.File]::ReadAllText($testFile)
        if ($content -notmatch 'Hello') { throw [System.IO.IOException]::new("Content verification failed") }

        # Step 5 — delete
        [System.IO.File]::Delete($testFile)

        $stats.OK++
        if ($i % 50 -eq 0 -or $i -eq 1) {
            Write-Log "Iteration $i / $Iterations — OK ($($stats.OK) passed, $($stats.Fail) failed)" 'OK'
        }
    } catch {
        $stats.Fail++
        $iterError = $_
        $category  = Get-ErrorCategory $_

        $levelMap  = @{ AUTH = 'AUTH'; NETWORK = 'FAIL'; TIMEOUT = 'WARN'; DISK = 'FAIL'; PATH = 'FAIL'; IO = 'FAIL' }
        $level     = $levelMap[$category]

        Write-Log "Iteration $i — [$category] $($_.Exception.Message)" $level

        $errorLog.Add([PSCustomObject]@{
            Iteration = $i
            Category  = $category
            Message   = $_.Exception.Message
            HResult   = '0x{0:X8}' -f $_.Exception.HResult
            Time      = Get-Date
        })

        # Deep diagnostics — only on first failure
        if (-not $diagnosed) {
            $diagnosed = $true
            $diagErr   = $iterError   # capture before switch overwrites $_
            Write-Section "Failure Diagnostics (first error at iteration $i)"

            switch ($category) {
                'AUTH' {
                    Write-Log "CAUSE: Authentication / permission failure" 'AUTH'
                    Write-Log "  The current user does not have write access to this path." 'AUTH'
                    Write-Log "  Common causes:" 'INFO'
                    Write-Log "    - Kerberos ticket expired or missing (run: klist purge, then re-authenticate)" 'INFO'
                    Write-Log "    - NTFS permissions deny write access for this account" 'INFO'
                    Write-Log "    - Share permissions do not allow write" 'INFO'
                    Write-Log "    - Password changed but cached credentials not updated" 'INFO'
                    Write-Log "    - Account locked out or disabled" 'INFO'

                    # Current tickets
                    try {
                        $klist = & klist 2>&1
                        Write-Log "Kerberos tickets at time of failure:" 'INFO'
                        $klist | ForEach-Object { Write-Log "  $_" 'INFO' }
                    } catch {}

                    # Net use
                    try {
                        $netuse = net use 2>&1
                        Write-Log "Net use (drive mappings):" 'INFO'
                        $netuse | ForEach-Object { Write-Log "  $_" 'INFO' }
                    } catch {}

                    # Recent auth failures in event log
                    try {
                        $since = (Get-Date).AddHours(-1)
                        $evts  = Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = @(4625,4740); StartTime = $since } `
                            -MaxEvents 5 -ErrorAction SilentlyContinue
                        if ($evts) {
                            Write-Log "$($evts.Count) recent auth failure event(s) in Security log:" 'WARN'
                            $evts | ForEach-Object {
                                Write-Log "  [$($_.TimeCreated.ToString('HH:mm:ss'))] Event $($_.Id): $(($_.Message -replace '\s+',' ').Substring(0,[Math]::Min(120,$_.Message.Length)))" 'WARN'
                            }
                        } else {
                            Write-Log "No recent auth failure events in Security log" 'OK'
                        }
                    } catch {
                        Write-Log "Could not read Security log (run as admin for event log access)" 'WARN'
                    }
                }

                'NETWORK' {
                    Write-Log "CAUSE: Network path unreachable" 'FAIL'
                    Write-Log "  The server or share is not reachable." 'INFO'
                    Write-Log "  Common causes:" 'INFO'
                    Write-Log "    - Server offline or network cable unplugged" 'INFO'
                    Write-Log "    - Firewall blocking SMB (port 445)" 'INFO'
                    Write-Log "    - DNS failure — server name does not resolve" 'INFO'
                    Write-Log "    - VPN disconnected" 'INFO'

                    if ($uncServer) {
                        try {
                            $ips = [System.Net.Dns]::GetHostAddresses($uncServer) | ForEach-Object { $_.IPAddressToString }
                            Write-Log "DNS resolves '$uncServer' to: $($ips -join ', ')" 'INFO'
                        } catch {
                            Write-Log "DNS resolution failed for '$uncServer'" 'FAIL'
                        }
                    }
                }

                'TIMEOUT' {
                    Write-Log "CAUSE: Operation timed out" 'WARN'
                    Write-Log "  The server is responding too slowly." 'INFO'
                    Write-Log "  Common causes:" 'INFO'
                    Write-Log "    - Server under heavy load" 'INFO'
                    Write-Log "    - Network congestion or high latency" 'INFO'
                    Write-Log "    - Storage I/O bottleneck on the server" 'INFO'
                }

                'DISK' {
                    Write-Log "CAUSE: Disk full or quota exceeded" 'FAIL'
                    try {
                        $drive = Split-Path $TestPath -Qualifier
                        $disk  = Get-PSDrive $drive.TrimEnd(':') -ErrorAction SilentlyContinue
                        if ($disk) {
                            $usedGB = [math]::Round($disk.Used / 1GB, 1)
                            $freeGB = [math]::Round($disk.Free / 1GB, 1)
                            Write-Log "Drive ${drive}: Used = ${usedGB} GB | Free = ${freeGB} GB" 'INFO'
                        }
                    } catch {}
                }

                'PATH' {
                    Write-Log "CAUSE: Path not found" 'FAIL'
                    Write-Log "  The target directory disappeared during the test." 'INFO'
                    Write-Log "  Common causes:" 'INFO'
                    Write-Log "    - Network share disconnected mid-test" 'INFO'
                    Write-Log "    - Drive remapped to different path" 'INFO'
                    Write-Log "    - Directory deleted by another process" 'INFO'
                }

                default {
                    Write-Log "CAUSE: General I/O error — HResult $('0x{0:X8}' -f $diagErr.Exception.HResult)" 'FAIL'
                }
            }
        }

        # Clean up test file if still present
        try { if (Test-Path $testFile) { Remove-Item $testFile -Force } } catch {}

        if ($StopOnFirstError) {
            Write-Log "Stopping after first error (-StopOnFirstError)" 'WARN'
            break
        }
    }

    if ($DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
}

# ── Summary ────────────────────────────────────────────────────────────────────
Write-Section 'Summary'

$total     = $stats.OK + $stats.Fail
$pct       = if ($total -gt 0) { [math]::Round($stats.OK / $total * 100, 1) } else { 0 }
Write-Log "Iterations  : $total" 'INFO'
$passedLevel = if ($stats.Fail -eq 0) { 'OK' } else { 'INFO' }
$failedLevel = if ($stats.Fail -eq 0) { 'OK' } else { 'FAIL' }
Write-Log "Passed      : $($stats.OK) ($pct%)" $passedLevel
Write-Log "Failed      : $($stats.Fail)" $failedLevel

if ($errorLog.Count -gt 0) {
    $grouped = $errorLog | Group-Object Category | Sort-Object Count -Descending
    Write-Log "Failure breakdown:" 'INFO'
    foreach ($g in $grouped) {
        Write-Log "  [$($g.Name)] $($g.Count) failure(s)" 'FAIL'
    }

    $firstErr = $errorLog[0]
    $lastErr  = $errorLog[-1]
    Write-Log "First failure : iteration $($firstErr.Iteration) at $($firstErr.Time.ToString('HH:mm:ss'))" 'INFO'
    if ($errorLog.Count -gt 1) {
        Write-Log "Last failure  : iteration $($lastErr.Iteration) at $($lastErr.Time.ToString('HH:mm:ss'))" 'INFO'
    }

    # Category-specific recommendation
    $topCategory = $grouped[0].Name
    Write-Log '' 'INFO'
    switch ($topCategory) {
        'AUTH'    { Write-Log "Recommendation: Run 'klist purge', re-enter credentials (net use * /delete), then retest" 'WARN' }
        'NETWORK' { Write-Log "Recommendation: Check server availability, SMB port 445, and DNS resolution" 'WARN' }
        'TIMEOUT' { Write-Log "Recommendation: Check server/storage load and network latency" 'WARN' }
        'DISK'    { Write-Log "Recommendation: Free up disk space or increase user quota on the target drive" 'WARN' }
        'PATH'    { Write-Log "Recommendation: Verify the share is still mounted and the directory exists" 'WARN' }
        default   { Write-Log "Recommendation: Review the detailed log for specific error codes" 'WARN' }
    }
} else {
    Write-Log "All iterations completed without errors" 'OK'
}

Write-Log "Log saved to : $logFile" 'INFO'
Write-Host ''
