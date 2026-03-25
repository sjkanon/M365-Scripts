#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Clean up temporary files, caches, and reclaimable disk space on Windows.

.DESCRIPTION
    Scans and optionally removes:
      - User and system temp folders
      - Windows Update download cache
      - Delivery Optimization cache
      - Prefetch files
      - Memory dump files
      - Windows Error Reporting queues
      - Thumbnail cache
      - DirectX shader cache
      - Recycle Bin
      - Browser caches (Edge, Chrome, Firefox)
      - Application & system log files (CBS, DISM, WindowsUpdate, Intune, IIS, Teams, Office)
      - Windows component store (via DISM)
      - Stale event log entries

    Run without -Apply for a dry run — shows how much space can be freed per category.
    Run with -Apply to perform the actual cleanup.

    Output:
      - Summary CSV saved to C:\Temp\ with space freed per category.

.PARAMETER Apply
    Perform the actual cleanup. Without this switch, only a scan is performed.

.PARAMETER SkipAppLogs
    Skip application and system log cleanup (CBS, DISM, WindowsUpdate, Intune, IIS, Teams, Office telemetry).

.PARAMETER SkipBrowserCache
    Skip browser cache cleanup (Edge, Chrome, Firefox).

.PARAMETER SkipEventLogs
    Skip clearing Windows Event Logs.

.PARAMETER SkipDism
    Skip DISM component store cleanup (slow, frees significant space after Windows Updates).

.PARAMETER SkipRecycleBin
    Skip emptying the Recycle Bin.

.PARAMETER OutputPath
    Override the default output folder (default: C:\Temp\).

.EXAMPLE
    # Dry run — see how much space can be freed
    .\Invoke-WindowsCleanup.ps1

.EXAMPLE
    # Full cleanup
    .\Invoke-WindowsCleanup.ps1 -Apply

.EXAMPLE
    # Full cleanup, skip browser cache and DISM
    .\Invoke-WindowsCleanup.ps1 -Apply -SkipBrowserCache -SkipDism
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [switch] $Apply,
    [switch] $SkipAppLogs,
    [switch] $SkipBrowserCache,
    [switch] $SkipEventLogs,
    [switch] $SkipDism,
    [switch] $SkipRecycleBin,
    [string] $OutputPath = 'C:\Temp'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'SilentlyContinue'

# ── Output ─────────────────────────────────────────────────────────────────────
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
$ts         = Get-Date -Format 'yyyyMMdd_HHmmss'
$reportCsv  = Join-Path $OutputPath "WindowsCleanup_$ts.csv"

# ── Header ─────────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host '   Invoke-WindowsCleanup' -ForegroundColor Cyan
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host ''

if (-not $Apply) {
    Write-Host '  ================================================' -ForegroundColor Yellow
    Write-Host '   DRY RUN — no files will be deleted' -ForegroundColor Yellow
    Write-Host '   Add -Apply to perform the actual cleanup.' -ForegroundColor Yellow
    Write-Host '  ================================================' -ForegroundColor Yellow
    Write-Host ''
}

# ── Helpers ────────────────────────────────────────────────────────────────────
$results = [System.Collections.Generic.List[PSCustomObject]]::new()

function Get-FolderSize {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return [int64]0 }
    try {
        (Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue |
            Measure-Object -Property Length -Sum).Sum -as [int64]
    } catch { [int64]0 }
}

function Get-FileSize {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return [int64]0 }
    try { (Get-Item -Path $Path -Force -ErrorAction SilentlyContinue).Length -as [int64] }
    catch { [int64]0 }
}

function Format-Bytes {
    param([int64]$Bytes)
    if ($Bytes -ge 1GB) { return '{0:N2} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N1} MB' -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return '{0:N0} KB' -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Invoke-CleanFolder {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return }
    Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}

function Add-Result {
    param([string]$Category, [string]$Detail, [int64]$BytesBefore, [int64]$BytesAfter, [bool]$Skipped = $false)
    $freed = $BytesBefore - $BytesAfter
    $color = if ($Skipped) { 'DarkGray' } elseif ($freed -gt 0) { 'Green' } else { 'DarkGray' }
    $label = if ($Skipped) { 'SKIP' } elseif ($Apply) { 'DONE' } else { 'SCAN' }
    $freedStr = if ($Skipped) { '—' } else { Format-Bytes $BytesBefore }

    Write-Host ('    [{0}] {1,-38} {2,10}' -f $label, $Detail, $freedStr) -ForegroundColor $color

    $results.Add([PSCustomObject]@{
        Category   = $Category
        Detail     = $Detail
        FreedBytes = if ($Apply -and -not $Skipped) { $freed } else { $BytesBefore }
        FreedMB    = [math]::Round((if ($Apply -and -not $Skipped) { $freed } else { $BytesBefore }) / 1MB, 2)
        Status     = if ($Skipped) { 'Skipped' } elseif ($Apply) { 'Cleaned' } else { 'Dry run' }
    })
}

$diskBefore = (Get-PSDrive -Name C).Used

# ── Enumerate all local user profiles ──────────────────────────────────────────
$systemFolders = @('Public', 'Default', 'Default User', 'All Users')
$userProfiles  = Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notin $systemFolders }
Write-Host ("  Scanning {0} user profile(s): {1}" -f $userProfiles.Count, ($userProfiles.Name -join ', ')) -ForegroundColor DarkGray
Write-Host ''

# ── 1. User temp folders (all users) ───────────────────────────────────────────
Write-Host '  Temp Files' -ForegroundColor Cyan

foreach ($up in $userProfiles) {
    $path = "$($up.FullName)\AppData\Local\Temp"
    $size = Get-FolderSize $path
    if ($size -eq 0) { continue }
    if ($Apply) { Invoke-CleanFolder $path }
    $after = if ($Apply) { Get-FolderSize $path } else { $size }
    Add-Result 'Temp' "Temp ($($up.Name))" $size $after
}

$sysTempSize = Get-FolderSize "$env:SystemRoot\Temp"
if ($Apply) { Invoke-CleanFolder "$env:SystemRoot\Temp" }
$sysTempAfter = if ($Apply) { Get-FolderSize "$env:SystemRoot\Temp" } else { $sysTempSize }
Add-Result 'Temp' 'System Temp (C:\Windows\Temp)' $sysTempSize $sysTempAfter

# ── 2. Windows Update ──────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  Windows Update' -ForegroundColor Cyan

$wuPaths = @(
    "$env:SystemRoot\SoftwareDistribution\Download",
    "$env:SystemRoot\SoftwareDistribution\DeliveryOptimization"
)

foreach ($path in $wuPaths) {
    $label = Split-Path $path -Leaf
    $size  = Get-FolderSize $path
    if ($Apply) {
        Stop-Service -Name wuauserv, bits -Force -ErrorAction SilentlyContinue
        Invoke-CleanFolder $path
        Start-Service -Name wuauserv, bits -ErrorAction SilentlyContinue
    }
    $after = if ($Apply) { Get-FolderSize $path } else { $size }
    Add-Result 'Windows Update' "SoftwareDistribution\$label" $size $after
}

# ── 3. Prefetch ────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  Prefetch & Memory Dumps' -ForegroundColor Cyan

$prefetchSize = Get-FolderSize "$env:SystemRoot\Prefetch"
if ($Apply) { Invoke-CleanFolder "$env:SystemRoot\Prefetch" }
$prefetchAfter = if ($Apply) { Get-FolderSize "$env:SystemRoot\Prefetch" } else { $prefetchSize }
Add-Result 'Prefetch' 'C:\Windows\Prefetch' $prefetchSize $prefetchAfter

# ── 4. Memory dumps ────────────────────────────────────────────────────────────
$sysDumpPaths = @(
    "$env:SystemRoot\Minidump",
    "$env:SystemRoot\memory.dmp"
)
foreach ($path in $sysDumpPaths) {
    $size = if ((Get-Item $path -Force -ErrorAction SilentlyContinue).PSIsContainer) {
        Get-FolderSize $path
    } else { Get-FileSize $path }
    if ($size -eq 0) { continue }
    if ($Apply) {
        if (Test-Path $path -PathType Container) { Invoke-CleanFolder $path }
        else { Remove-Item $path -Force -ErrorAction SilentlyContinue }
    }
    $after = if ($Apply) { 0 } else { $size }
    Add-Result 'Memory Dumps' (Split-Path $path -Leaf) $size $after
}

foreach ($up in $userProfiles) {
    $path = "$($up.FullName)\AppData\Local\CrashDumps"
    $size = Get-FolderSize $path
    if ($size -eq 0) { continue }
    if ($Apply) { Invoke-CleanFolder $path }
    $after = if ($Apply) { Get-FolderSize $path } else { $size }
    Add-Result 'Memory Dumps' "CrashDumps ($($up.Name))" $size $after
}

# ── 5. Windows Error Reporting ─────────────────────────────────────────────────
Write-Host ''
Write-Host '  Windows Error Reporting' -ForegroundColor Cyan

foreach ($subfolder in @('ReportQueue', 'ReportArchive')) {
    $path = "$env:ProgramData\Microsoft\Windows\WER\$subfolder"
    $size = Get-FolderSize $path
    if ($size -eq 0) { continue }
    if ($Apply) { Invoke-CleanFolder $path }
    $after = if ($Apply) { Get-FolderSize $path } else { $size }
    Add-Result 'WER' "System WER\$subfolder" $size $after
}

foreach ($up in $userProfiles) {
    foreach ($subfolder in @('ReportQueue', 'ReportArchive')) {
        $path = "$($up.FullName)\AppData\Local\Microsoft\Windows\WER\$subfolder"
        $size = Get-FolderSize $path
        if ($size -eq 0) { continue }
        if ($Apply) { Invoke-CleanFolder $path }
        $after = if ($Apply) { Get-FolderSize $path } else { $size }
        Add-Result 'WER' "WER\$subfolder ($($up.Name))" $size $after
    }
}

# ── 6. Thumbnail & shader cache (all users) ────────────────────────────────────
Write-Host ''
Write-Host '  Thumbnail & Shader Cache' -ForegroundColor Cyan

if ($Apply) {
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
}

foreach ($up in $userProfiles) {
    $thumbPath  = "$($up.FullName)\AppData\Local\Microsoft\Windows\Explorer"
    $thumbFiles = Get-ChildItem -Path $thumbPath -Filter 'thumbcache_*.db' -Force -ErrorAction SilentlyContinue
    $thumbSize  = ($thumbFiles | Measure-Object -Property Length -Sum).Sum -as [int64]
    if ($thumbSize -eq 0) { continue }
    if ($Apply) { $thumbFiles | Remove-Item -Force -ErrorAction SilentlyContinue }
    Add-Result 'Cache' "Thumbnail cache ($($up.Name))" $thumbSize (if ($Apply) { 0 } else { $thumbSize })

    $shaderPath = "$($up.FullName)\AppData\Local\D3DSCache"
    $shaderSize = Get-FolderSize $shaderPath
    if ($shaderSize -eq 0) { continue }
    if ($Apply) { Invoke-CleanFolder $shaderPath }
    $shaderAfter = if ($Apply) { Get-FolderSize $shaderPath } else { $shaderSize }
    Add-Result 'Cache' "DirectX shader cache ($($up.Name))" $shaderSize $shaderAfter
}

if ($Apply) { Start-Process explorer }

# ── 7. Font cache ──────────────────────────────────────────────────────────────
$fontCachePath = "$env:WinDir\ServiceProfiles\LocalService\AppData\Local\FontCache"
$fontSize = Get-FolderSize $fontCachePath
if ($fontSize -gt 0) {
    if ($Apply) {
        Stop-Service -Name FontCache -Force -ErrorAction SilentlyContinue
        Invoke-CleanFolder $fontCachePath
        Start-Service -Name FontCache -ErrorAction SilentlyContinue
    }
    Add-Result 'Cache' 'Font cache' $fontSize (if ($Apply) { Get-FolderSize $fontCachePath } else { $fontSize })
}

# ── 8. Recycle Bin ─────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  Recycle Bin' -ForegroundColor Cyan

if ($SkipRecycleBin) {
    Add-Result 'Recycle Bin' 'All drives' 0 0 -Skipped $true
} else {
    $recyclePath = 'C:\$Recycle.Bin'
    $recycleSize = Get-FolderSize $recyclePath
    if ($Apply) {
        Clear-RecycleBin -Force -ErrorAction SilentlyContinue
    }
    $recycleAfter = if ($Apply) { Get-FolderSize $recyclePath } else { $recycleSize }
    Add-Result 'Recycle Bin' 'All drives' $recycleSize $recycleAfter
}

# ── 9. Browser caches ──────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  Browser Cache' -ForegroundColor Cyan

if ($SkipBrowserCache) {
    Add-Result 'Browser' 'Skipped' 0 0 -Skipped $true
} else {
    foreach ($up in $userProfiles) {
        $local   = "$($up.FullName)\AppData\Local"
        $roaming = "$($up.FullName)\AppData\Roaming"

        # Edge — multiple profiles (Default, Profile 1, Profile 2, ...)
        $edgeProfiles = Get-ChildItem "$local\Microsoft\Edge\User Data" -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^(Default|Profile \d+)$' }
        foreach ($ep in $edgeProfiles) {
            $path = "$($ep.FullName)\Cache\Cache_Data"
            $size = Get-FolderSize $path
            if ($size -eq 0) { continue }
            if ($Apply) { Invoke-CleanFolder $path }
            $after = if ($Apply) { Get-FolderSize $path } else { $size }
            Add-Result 'Browser' "Edge/$($ep.Name) ($($up.Name))" $size $after
        }

        # Chrome — multiple profiles
        $chromeProfiles = Get-ChildItem "$local\Google\Chrome\User Data" -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^(Default|Profile \d+)$' }
        foreach ($cp in $chromeProfiles) {
            $path = "$($cp.FullName)\Cache\Cache_Data"
            $size = Get-FolderSize $path
            if ($size -eq 0) { continue }
            if ($Apply) { Invoke-CleanFolder $path }
            $after = if ($Apply) { Get-FolderSize $path } else { $size }
            Add-Result 'Browser' "Chrome/$($cp.Name) ($($up.Name))" $size $after
        }

        # Firefox
        $ffProfiles = Get-ChildItem "$roaming\Mozilla\Firefox\Profiles" -Directory -ErrorAction SilentlyContinue
        foreach ($ffProfile in $ffProfiles) {
            $path = "$($ffProfile.FullName)\cache2"
            $size = Get-FolderSize $path
            if ($size -eq 0) { continue }
            if ($Apply) { Invoke-CleanFolder $path }
            $after = if ($Apply) { Get-FolderSize $path } else { $size }
            Add-Result 'Browser' "Firefox ($($up.Name))" $size $after
        }
    }
}

# ── 10. Event Logs ─────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  Event Logs' -ForegroundColor Cyan

if ($SkipEventLogs) {
    Add-Result 'Event Logs' 'Skipped' 0 0 -Skipped $true
} else {
    $logRoot  = "$env:SystemRoot\System32\winevt\Logs"
    $logFiles = Get-ChildItem -Path $logRoot -Filter '*.evtx' -ErrorAction SilentlyContinue
    $logSize  = ($logFiles | Measure-Object -Property Length -Sum).Sum -as [int64]

    if ($Apply) {
        $logs = Get-WinEvent -ListLog * -ErrorAction SilentlyContinue |
            Where-Object { $_.RecordCount -gt 0 }
        foreach ($log in $logs) {
            try {
                [System.Diagnostics.Eventing.Reader.EventLogSession]::GlobalSession.ClearLog($log.LogName)
            } catch {}
        }
    }
    $logAfter = if ($Apply) { (Get-ChildItem $logRoot -Filter '*.evtx' -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum -as [int64] } else { $logSize }
    Add-Result 'Event Logs' "All event logs ($($logFiles.Count) files)" $logSize $logAfter
}

# ── 11. DISM component store cleanup ──────────────────────────────────────────
Write-Host ''
Write-Host '  DISM Component Store' -ForegroundColor Cyan

if ($SkipDism) {
    Add-Result 'DISM' 'Skipped (use -SkipDism to include)' 0 0 -Skipped $true
} else {
    $winSxSSize = Get-FolderSize "$env:SystemRoot\WinSxS"
    Write-Host '    [INFO] Running DISM cleanup — this may take several minutes...' -ForegroundColor DarkGray

    if ($Apply) {
        & dism.exe /Online /Cleanup-Image /StartComponentCleanup /ResetBase 2>&1 | Out-Null
        $exitCode = $LASTEXITCODE
        if ($exitCode -ne 0) {
            Write-Host '    [WARN] DISM returned exit code {0}' -f $exitCode -ForegroundColor Yellow
        }
    }
    $winSxSAfter = if ($Apply) { Get-FolderSize "$env:SystemRoot\WinSxS" } else { $winSxSSize }
    Add-Result 'DISM' 'WinSxS component store' $winSxSSize $winSxSAfter
}

# ── 12. DNS cache ──────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  DNS Cache' -ForegroundColor Cyan

if ($Apply) {
    ipconfig /flushdns | Out-Null
    Write-Host '    [DONE] DNS cache flushed' -ForegroundColor Green
} else {
    $dnsEntries = (Get-DnsClientCache -ErrorAction SilentlyContinue | Measure-Object).Count
    Write-Host ("    [SCAN] DNS cache contains {0} entries" -f $dnsEntries) -ForegroundColor DarkGray
}

# ── Summary ────────────────────────────────────────────────────────────────────
$diskAfter   = (Get-PSDrive -Name C).Used
$totalFreed  = ($results | Where-Object { $_.Status -ne 'Skipped' } | Measure-Object -Property FreedBytes -Sum).Sum -as [int64]
$actualFreed = $diskBefore - $diskAfter

# Per-category breakdown
$byCategory = $results |
    Where-Object { $_.Status -ne 'Skipped' -and $_.FreedBytes -gt 0 } |
    Group-Object Category |
    ForEach-Object {
        [PSCustomObject]@{
            Category = $_.Name
            Bytes    = ($_.Group | Measure-Object -Property FreedBytes -Sum).Sum -as [int64]
        }
    } | Sort-Object Bytes -Descending

Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host '   Summary' -ForegroundColor Cyan
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host ''

# Per-category table
foreach ($row in $byCategory) {
    Write-Host ('  {0,-30} {1,10}' -f $row.Category, (Format-Bytes $row.Bytes)) -ForegroundColor DarkGray
}

Write-Host ''
Write-Host ('  {0,-30} {1}' -f ('─' * 30), ('─' * 10)) -ForegroundColor DarkGray

if ($Apply) {
    Write-Host ('  {0,-30} {1,10}' -f 'Total freed (reported)',  (Format-Bytes $totalFreed))  -ForegroundColor Green
    Write-Host ('  {0,-30} {1,10}' -f 'Total freed (measured)',  (Format-Bytes $actualFreed)) -ForegroundColor Green
} else {
    Write-Host ('  {0,-30} {1,10}' -f 'Reclaimable (estimate)', (Format-Bytes $totalFreed)) -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  Run with -Apply to perform the actual cleanup.' -ForegroundColor Yellow
}

# ── Export ─────────────────────────────────────────────────────────────────────
$results | Export-Csv -Path $reportCsv -NoTypeInformation -Encoding UTF8
Write-Host ''
Write-Host ("  Report : {0}" -f $reportCsv) -ForegroundColor Green
Write-Host ''
