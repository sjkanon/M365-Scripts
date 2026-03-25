#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Monitor a work directory for intermittent access-denied failures and capture the cause.

.DESCRIPTION
    Designed as a follow-up to Test-FileIODiagnostics.ps1 when the cause of intermittent
    write failures is unclear. Runs a continuous write/delete loop and on every failure:

      - Captures NTFS permissions (icacls) before and after to detect permission changes
      - Records which processes have a handle open on the directory (requires Sysinternals
        Handle.exe in PATH or -HandleExe path; skipped if not available)
      - Uses FileSystemWatcher to log every change to the directory (created, deleted,
        renamed, changed) with millisecond timestamps
      - Snapshots the running process list around the failure to spot new processes
      - Checks the Windows Security event log for 4656/4663 (object access) if auditing
        is enabled on the folder
      - Saves everything to a timestamped log file

    Run this while SAS (or any other application) is actively using the directory to
    correlate the exact moment write access is lost with what caused it.

.PARAMETER TestPath
    Directory to monitor and test. E.g. G:\sas\work

.PARAMETER Iterations
    Maximum number of write/delete cycles. Default: 5000 (runs until a failure or Ctrl+C).

.PARAMETER DelayMs
    Milliseconds between iterations. Default: 50 — fast enough to catch intermittent issues.

.PARAMETER HandleExe
    Path to Sysinternals Handle.exe. Optional — skipped if not found.
    Download from: https://learn.microsoft.com/sysinternals/downloads/handle

.PARAMETER LogPath
    Output folder. Default: C:\Temp\

.EXAMPLE
    # Basic monitor — runs until failure or Ctrl+C
    .\Test-WorkDirMonitor.ps1 -TestPath G:\sas\work

.EXAMPLE
    # With Sysinternals Handle.exe for process handle detection
    .\Test-WorkDirMonitor.ps1 -TestPath G:\sas\work -HandleExe C:\Tools\handle.exe

.EXAMPLE
    # Slow down the loop to reduce noise, run 10000 cycles
    .\Test-WorkDirMonitor.ps1 -TestPath G:\sas\work -Iterations 10000 -DelayMs 200
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string] $TestPath,

    [int]    $Iterations = 5000,
    [int]    $DelayMs    = 50,
    [string] $HandleExe  = '',
    [string] $LogPath    = 'C:\Temp'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'SilentlyContinue'

# ── Log setup ──────────────────────────────────────────────────────────────────
if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath | Out-Null }
$ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile = Join-Path $LogPath "WorkDirMonitor_$ts.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line  = '[{0}] [{1,-5}] {2}' -f (Get-Date -Format 'HH:mm:ss.fff'), $Level, $Message
    $color = switch ($Level) {
        'OK'    { 'Green'   }
        'WARN'  { 'Yellow'  }
        'FAIL'  { 'Red'     }
        'ALERT' { 'Magenta' }
        'HEAD'  { 'Cyan'    }
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

# ── Capture helpers ────────────────────────────────────────────────────────────

function Get-DirPermissions {
    param([string]$Path)
    try {
        $acl = Get-Acl -Path $Path -ErrorAction Stop
        return $acl.AccessToString
    } catch {
        return "(could not read ACL: $_)"
    }
}

function Get-Icacls {
    param([string]$Path)
    try {
        $out = & icacls $Path 2>&1 | Out-String
        return $out.Trim()
    } catch {
        return "(icacls failed)"
    }
}

function Get-ProcessesWithHandle {
    param([string]$Path, [string]$HandleExePath)
    if (-not $HandleExePath -or -not (Test-Path $HandleExePath)) { return $null }
    try {
        # Accept EULA silently on first run
        $out = & $HandleExePath $Path -nobanner -accepteula 2>&1 | Out-String
        return $out.Trim()
    } catch {
        return "(handle.exe failed)"
    }
}

function Get-ProcessSnapshot {
    try {
        return Get-Process | Select-Object Id, Name, CPU, StartTime |
            Sort-Object Name |
            ForEach-Object { "$($_.Id.ToString().PadLeft(6))  $($_.Name.PadRight(30))  started: $($_.StartTime)" }
    } catch { return @() }
}

function Get-RecentObjectAccessEvents {
    param([string]$Path, [int]$Seconds = 10)
    try {
        $since = (Get-Date).AddSeconds(-$Seconds)
        $evts  = Get-WinEvent -FilterHashtable @{
            LogName   = 'Security'
            Id        = @(4656, 4663, 4670)   # handle request, object access, permissions changed
            StartTime = $since
        } -MaxEvents 20 -ErrorAction Stop |
            Where-Object { $_.Message -imatch [regex]::Escape($Path) }
        return $evts
    } catch {
        return @()
    }
}

# ── FileSystemWatcher setup ────────────────────────────────────────────────────

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path                = $TestPath
$watcher.NotifyFilter        = [System.IO.NotifyFilters]'FileName,DirectoryName,Attributes,Security,LastWrite'
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true

$fsEvents = [System.Collections.Generic.List[string]]::new()
$fsLock   = [System.Object]::new()

$onChange = Register-ObjectEvent $watcher 'Changed' -Action {
    $e = $Event.SourceEventArgs
    [System.Threading.Monitor]::Enter($fsLock)
    try { $fsEvents.Add("[$(Get-Date -Format 'HH:mm:ss.fff')] CHANGED  : $($e.FullPath)") }
    finally { [System.Threading.Monitor]::Exit($fsLock) }
}
$onCreate = Register-ObjectEvent $watcher 'Created' -Action {
    $e = $Event.SourceEventArgs
    [System.Threading.Monitor]::Enter($fsLock)
    try { $fsEvents.Add("[$(Get-Date -Format 'HH:mm:ss.fff')] CREATED  : $($e.FullPath)") }
    finally { [System.Threading.Monitor]::Exit($fsLock) }
}
$onDelete = Register-ObjectEvent $watcher 'Deleted' -Action {
    $e = $Event.SourceEventArgs
    [System.Threading.Monitor]::Enter($fsLock)
    try { $fsEvents.Add("[$(Get-Date -Format 'HH:mm:ss.fff')] DELETED  : $($e.FullPath)") }
    finally { [System.Threading.Monitor]::Exit($fsLock) }
}
$onRename = Register-ObjectEvent $watcher 'Renamed' -Action {
    $e = $Event.SourceEventArgs
    [System.Threading.Monitor]::Enter($fsLock)
    try { $fsEvents.Add("[$(Get-Date -Format 'HH:mm:ss.fff')] RENAMED  : $($e.OldFullPath) → $($e.FullPath)") }
    finally { [System.Threading.Monitor]::Exit($fsLock) }
}

# ── Pre-flight ─────────────────────────────────────────────────────────────────
Write-Section 'Pre-flight'

Write-Log "Test path   : $TestPath"
Write-Log "Running as  : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
Write-Log "Log file    : $logFile"

if ($HandleExe -and (Test-Path $HandleExe)) {
    Write-Log "Handle.exe  : $HandleExe (process handle detection enabled)" 'OK'
} else {
    Write-Log "Handle.exe  : not configured — process handle detection disabled" 'WARN'
    Write-Log "             Download from Sysinternals and pass -HandleExe to enable" 'WARN'
}

# Baseline permissions
$baselineAcl = Get-Icacls $TestPath
Write-Log "Baseline NTFS permissions:" 'INFO'
$baselineAcl -split "`n" | ForEach-Object { Write-Log "  $_" 'INFO' }

# Baseline process list
$baselineProcs = Get-ProcessSnapshot

Write-Section "Monitoring $TestPath — $Iterations iterations (Ctrl+C to stop)"

# ── Main loop ──────────────────────────────────────────────────────────────────
$testFile   = Join-Path $TestPath "workmon_test_$pid.tmp"
$stats      = @{ OK = 0; Fail = 0 }
$failCount  = 0

try {
    for ($i = 1; $i -le $Iterations; $i++) {

        # Drain FS watcher events every 100 iterations (non-failure context)
        if ($i % 100 -eq 0) {
            [System.Threading.Monitor]::Enter($fsLock)
            try {
                if ($fsEvents.Count -gt 0) {
                    Write-Log "FileSystemWatcher — $($fsEvents.Count) directory event(s) in last 100 iterations:" 'WARN'
                    $fsEvents | ForEach-Object { Write-Log "  $_" 'WARN' }
                    $fsEvents.Clear()
                }
            } finally { [System.Threading.Monitor]::Exit($fsLock) }
        }

        try {
            [System.IO.File]::WriteAllText($testFile, "test`n")
            [System.IO.File]::AppendAllText($testFile, "data`n")
            [System.IO.File]::Delete($testFile)

            $stats.OK++
            if ($i % 500 -eq 0 -or $i -eq 1) {
                Write-Log "Iteration $i / $Iterations — OK ($($stats.OK) passed, $($stats.Fail) failed)" 'OK'
            }
        } catch {
            $stats.Fail++
            $failCount++
            $failMsg = $_.Exception.Message

            Write-Log "═══ FAILURE #$failCount at iteration $i ═══" 'ALERT'
            Write-Log "Error: $failMsg" 'FAIL'

            # 1. Flush all FS watcher events up to this moment
            [System.Threading.Monitor]::Enter($fsLock)
            try {
                if ($fsEvents.Count -gt 0) {
                    Write-Log "FileSystemWatcher events leading up to failure:" 'ALERT'
                    $fsEvents | ForEach-Object { Write-Log "  $_" 'ALERT' }
                    $fsEvents.Clear()
                } else {
                    Write-Log "FileSystemWatcher: no directory events captured before failure" 'INFO'
                }
            } finally { [System.Threading.Monitor]::Exit($fsLock) }

            # 2. Current NTFS permissions — compare to baseline
            $currentAcl = Get-Icacls $TestPath
            if ($currentAcl -ne $baselineAcl) {
                Write-Log "NTFS PERMISSIONS CHANGED since baseline!" 'ALERT'
                Write-Log "Current permissions:" 'ALERT'
                $currentAcl -split "`n" | ForEach-Object { Write-Log "  $_" 'ALERT' }
            } else {
                Write-Log "NTFS permissions: unchanged from baseline" 'INFO'
            }

            # 3. Who has a handle on the directory
            $handles = Get-ProcessesWithHandle $TestPath $HandleExe
            if ($handles) {
                Write-Log "Processes with open handles on '$TestPath':" 'ALERT'
                $handles -split "`n" | ForEach-Object { Write-Log "  $_" 'ALERT' }
            }

            # 4. New processes since baseline
            $currentProcs = Get-ProcessSnapshot
            $newProcs = $currentProcs | Where-Object { $_ -notin $baselineProcs }
            if ($newProcs) {
                Write-Log "New processes since start:" 'ALERT'
                $newProcs | ForEach-Object { Write-Log "  $_" 'ALERT' }
            }

            # 5. Security audit events (4656/4663/4670) — only if object access auditing is on
            $auditEvts = Get-RecentObjectAccessEvents $TestPath 15
            if ($auditEvts -and $auditEvts.Count -gt 0) {
                Write-Log "$($auditEvts.Count) Security audit event(s) for this path in the last 15 seconds:" 'ALERT'
                $auditEvts | ForEach-Object {
                    Write-Log "  [$($_.TimeCreated.ToString('HH:mm:ss.fff'))] Event $($_.Id): $(($_.Message -replace '\s+',' ').Substring(0, [Math]::Min(160, $_.Message.Length)))" 'ALERT'
                }
            }

            # 6. icacls on the test file itself if it still exists
            if (Test-Path $testFile) {
                $fileAcl = Get-Icacls $testFile
                Write-Log "Permissions on test file:" 'INFO'
                $fileAcl -split "`n" | ForEach-Object { Write-Log "  $_" 'INFO' }
                try { Remove-Item $testFile -Force } catch {}
            }

            Write-Log "═══ End of failure #$failCount diagnostics ═══" 'ALERT'

            # Stop after 3 failures — enough data to diagnose
            if ($failCount -ge 3) {
                Write-Log "3 failures captured — stopping to preserve diagnostic data" 'WARN'
                break
            }
        }

        if ($DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
    }
} finally {
    # Clean up watcher
    $watcher.EnableRaisingEvents = $false
    Unregister-Event $onChange.Id -ErrorAction SilentlyContinue
    Unregister-Event $onCreate.Id -ErrorAction SilentlyContinue
    Unregister-Event $onDelete.Id -ErrorAction SilentlyContinue
    Unregister-Event $onRename.Id -ErrorAction SilentlyContinue
    $watcher.Dispose()
    if (Test-Path $testFile) { Remove-Item $testFile -Force -ErrorAction SilentlyContinue }
}

# ── Summary ────────────────────────────────────────────────────────────────────
Write-Section 'Summary'

Write-Log "Iterations  : $($stats.OK + $stats.Fail)" 'INFO'
Write-Log "Passed      : $($stats.OK)" 'OK'
$fLevel = if ($stats.Fail -eq 0) { 'OK' } else { 'FAIL' }
Write-Log "Failed      : $($stats.Fail)" $fLevel

if ($stats.Fail -eq 0) {
    Write-Log "No failures detected during this run" 'OK'
    Write-Log "The issue may be timing-dependent — try running with -DelayMs 0 and more iterations" 'WARN'
} else {
    Write-Log "Review ALERT lines above for the cause. Key things to look for:" 'INFO'
    Write-Log "  - NTFS PERMISSIONS CHANGED: a process reset the ACL on the directory" 'INFO'
    Write-Log "  - FileSystemWatcher events: shows what file/folder operations occurred" 'INFO'
    Write-Log "  - Processes with handles: identifies which process had the directory locked" 'INFO'
    Write-Log "  - New processes: a scheduler or cleanup job started just before the failure" 'INFO'
    Write-Log "  If no process handle data: install Sysinternals Handle.exe and pass -HandleExe" 'WARN'
}

Write-Log "Log saved to : $logFile" 'INFO'
Write-Host ''
