#Requires -Version 7.0
<#
.SYNOPSIS
    Clears the shared script temp folder.

.DESCRIPTION
    Runs in dry-run mode by default and reports what would be removed.
    Use -Apply to actually delete files.

    Default target folder is platform-specific:
      - Windows: C:\Temp
      - Linux/macOS: /tmp

.PARAMETER Apply
    Perform the cleanup. Without this switch, no files are deleted.

.PARAMETER TempPath
    Temp folder to clean. Defaults to C:\Temp on Windows, /tmp on Linux/macOS.

.PARAMETER OlderThanDays
    Only target items older than this number of days.
    Default is 1 day to reduce the chance of removing actively used temp files.

.EXAMPLE
    .\Clear-TempFiles.ps1

.EXAMPLE
    .\Clear-TempFiles.ps1 -Apply -OlderThanDays 7

.EXAMPLE
    .\Clear-TempFiles.ps1 -Apply -TempPath '/var/tmp'
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$Apply,
    [string]$TempPath,
    [ValidateRange(0, 3650)]
    [int]$OlderThanDays = 1
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'SilentlyContinue'

function Get-FolderSizeBytes {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return [int64]0
    }

    try {
        $sum = (Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue |
            Measure-Object -Property Length -Sum).Sum
        return [int64]($sum -as [int64])
    } catch {
        return [int64]0
    }
}

function Format-Bytes {
    param([int64]$Bytes)

    if ($Bytes -ge 1GB) { return '{0:N2} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N1} MB' -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return '{0:N0} KB' -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Get-CleanupCandidates {
    param(
        [string]$Path,
        [datetime]$Cutoff
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return @()
    }

    Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $Cutoff } |
        Sort-Object -Property FullName -Descending
}

function Get-DefaultTempPath {
    if ($IsWindows) { return 'C:\Temp' }
    return '/tmp'
}

$cutoff = (Get-Date).AddDays(-$OlderThanDays)
$resolvedTempPath = if ([string]::IsNullOrWhiteSpace($TempPath)) { Get-DefaultTempPath } else { $TempPath }

Write-Host ''
Write-Host '========================================' -ForegroundColor Cyan
Write-Host ' Clear-TempFiles' -ForegroundColor Cyan
Write-Host '========================================' -ForegroundColor Cyan
Write-Host (" Mode         : {0}" -f ($(if ($Apply) { 'Apply' } else { 'Dry run' })))
Write-Host (" Temp path    : {0}" -f $resolvedTempPath)
Write-Host (" Older than   : {0} day(s)" -f $OlderThanDays)
Write-Host (" Cutoff date  : {0}" -f $cutoff)
Write-Host ''

$totalBefore = [int64]0
$totalAfter = [int64]0

if (-not (Test-Path -LiteralPath $resolvedTempPath)) {
    Write-Host ("[SKIP] Script Temp          {0}" -f '(path not found)') -ForegroundColor DarkGray
    Write-Host ''
    Write-Host 'Summary' -ForegroundColor Cyan
    Write-Host 'Total size scanned : 0 B'
    Write-Host 'Dry run complete. Re-run with -Apply to delete files.' -ForegroundColor Yellow
    return
}

$before = Get-FolderSizeBytes -Path $resolvedTempPath
$totalBefore += $before

$candidates = Get-CleanupCandidates -Path $resolvedTempPath -Cutoff $cutoff
if ($Apply -and $PSCmdlet.ShouldProcess($resolvedTempPath, "Remove temp files older than $OlderThanDays day(s)")) {
    foreach ($item in $candidates) {
        Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }
}

$after = Get-FolderSizeBytes -Path $resolvedTempPath
if (-not $Apply) {
    $after = $before
}
$totalAfter += $after

$wouldFree = if ($Apply) { [Math]::Max(0, $before - $after) } else {
    [int64](($candidates | Where-Object { -not $_.PSIsContainer } | Measure-Object -Property Length -Sum).Sum)
}

$state = if ($Apply) { 'DONE' } else { 'SCAN' }
Write-Host ("[{0}] {1,-20} Found: {2,10}  Freed: {3,10}" -f $state, 'Script Temp', (Format-Bytes $before), (Format-Bytes $wouldFree)) -ForegroundColor Green

$overallFreed = if ($Apply) { [Math]::Max(0, $totalBefore - $totalAfter) } else { $null }

Write-Host ''
Write-Host 'Summary' -ForegroundColor Cyan
Write-Host ("Total size scanned : {0}" -f (Format-Bytes $totalBefore))
if ($Apply) {
    Write-Host ("Total space freed  : {0}" -f (Format-Bytes $overallFreed)) -ForegroundColor Green
} else {
    Write-Host 'Dry run complete. Re-run with -Apply to delete files.' -ForegroundColor Yellow
}
