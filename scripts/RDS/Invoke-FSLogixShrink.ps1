<#
.SYNOPSIS
    Shrinks FSLogix profile / ODFC containers on a share with Invoke-FslShrinkDisk,
    and checks whether FSLogix's own compaction at sign-out is doing that job.

.DESCRIPTION
    Three steps, each of which can be run on its own:

    1. Download   - fetches Invoke-FslShrinkDisk (github.com/FSLogix/Invoke-FslShrinkDisk)
                    at a pinned commit, unpacks and unblocks it, and checks the SHA-256 of
                    the script against the hash recorded here. Skipped when it is already
                    in place, unless -Force is given.
    2. Report     - lists every VHD(X) on the share, largest first, with the total size.
                    -ReportOnly stops here.
    3. Shrink     - runs Invoke-FslShrinkDisk over the share and summarises its CSV log:
                    how many disks shrank, how much space came back, which failed.

    -CheckHost instead checks, on this session host, whether FSLogix's built-in
    VHD Disk Compaction (FSLogix 2210 / 2.9.8361 and later) can run at sign-out:
    FSLogix version, VHDCompactDisk, the Optimize Drives service (defragsvc) and
    dynamic disks. When it can, a manual shrink is only needed to catch up - for
    disks of users who rarely sign out, or below FSLogix's 20% threshold.

    A disk that is attached (user signed in) cannot be shrunk; Invoke-FslShrinkDisk
    reports it in the CSV and moves on. Run it out of hours or with hosts drained.

.PARAMETER Path
    The share (or folder) holding the containers, e.g.
    \\<storageaccount>.file.core.windows.net\<share>\Profiles. Searched recursively.

.PARAMETER ReportOnly
    Only download and list the disks with their size; shrink nothing.

.PARAMETER IgnoreLessThanGB
    Skip disks smaller than this (default: 5).

.PARAMETER RatioFreeSpace
    Only shrink a disk with at least this fraction of free space inside it
    (default: 0.1 = 10%; Invoke-FslShrinkDisk's own default is 0.05).

.PARAMETER ThrottleLimit
    Disks processed at the same time (default: 4). Microsoft advises at most twice
    the number of CPU cores; the load lands mostly on the storage.

.PARAMETER LogFilePath
    CSV log written by Invoke-FslShrinkDisk (default:
    %SystemDrive%\Temp\FslShrink_<timestamp>.csv). The folder is created when missing.

.PARAMETER ToolPath
    Where Invoke-FslShrinkDisk is kept (default: %SystemDrive%\Scripts\Invoke-FslShrinkDisk).

.PARAMETER Force
    Download Invoke-FslShrinkDisk again even when it is already in place.

.PARAMETER CheckHost
    Check whether FSLogix's built-in compaction at sign-out can run on this host.
    Needs no -Path.

.EXAMPLE
    .\Invoke-FSLogixShrink.ps1 -Path \\sa.file.core.windows.net\profiles\Profiles -ReportOnly
    Lists every container on the share, largest first.

.EXAMPLE
    .\Invoke-FSLogixShrink.ps1 -Path \\sa.file.core.windows.net\profiles\Profiles
    Shrinks every container of 5 GB or more with at least 10% free space inside.

.EXAMPLE
    .\Invoke-FSLogixShrink.ps1 -CheckHost
    Tells whether FSLogix compacts the disks itself at sign-out on this host.

.NOTES
    Run elevated: Invoke-FslShrinkDisk mounts each disk. The account needs access to
    the share - on Azure Files through Kerberos (Entra/AD) or the storage account key.
    Hyper-V is not needed. Exit codes: 0 = ok, 1 = something failed or needs attention.
#>

[CmdletBinding(DefaultParameterSetName = 'Shrink')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Shrink', Position = 0)]
    [string]$Path,

    [Parameter(ParameterSetName = 'Shrink')]
    [switch]$ReportOnly,

    [Parameter(ParameterSetName = 'Shrink')]
    [ValidateRange(0, 1024)]
    [int]$IgnoreLessThanGB = 5,

    [Parameter(ParameterSetName = 'Shrink')]
    [ValidateRange(0.0, 0.99)]
    [double]$RatioFreeSpace = 0.1,

    [Parameter(ParameterSetName = 'Shrink')]
    [ValidateRange(1, 64)]
    [int]$ThrottleLimit = 4,

    [Parameter(ParameterSetName = 'Shrink')]
    [string]$LogFilePath = "$env:SystemDrive\Temp\FslShrink_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter(ParameterSetName = 'Shrink')]
    [string]$ToolPath = "$env:SystemDrive\Scripts\Invoke-FslShrinkDisk",

    [Parameter(ParameterSetName = 'Shrink')]
    [switch]$Force,

    [Parameter(Mandatory, ParameterSetName = 'CheckHost')]
    [switch]$CheckHost
)

$ErrorActionPreference = 'Stop'

# Pinned so a change upstream never runs here unseen. To move to a newer version:
# take the commit from github.com/FSLogix/Invoke-FslShrinkDisk, read the diff, and
# record the new SHA-256 of Invoke-FslShrinkDisk.ps1.
$ToolCommit = 'bfe050459e37a095db3a1bc6eea387e318688d3c'   # 2025-06-19
$ToolHash   = '9E28CF9F8D111B642C1D701E452DFD727D8268520B5EFDFE4996D77FE0F813D0'

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host ("========== {0} ==========" -f $Title.ToUpper()) -ForegroundColor Cyan
}

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ---------------------------------------------------------------------------
# -CheckHost: can FSLogix compact the disks itself at sign-out?
# ---------------------------------------------------------------------------
if ($CheckHost) {
    Write-Section "FSLogix compaction at sign-out"
    $problems = 0

    $frx = Join-Path $env:ProgramFiles 'FSLogix\Apps\frx.exe'
    if (-not (Test-Path $frx)) {
        Write-Host "FSLogix is not installed on $env:COMPUTERNAME ($frx not found)." -ForegroundColor Yellow
        exit 1
    }
    $version = [version]((Get-Item $frx).VersionInfo.ProductVersion -replace '[^\d\.].*$')
    if ($version -ge [version]'2.9.8361') {
        Write-Host "FSLogix version    : $version - has VHD Disk Compaction" -ForegroundColor Green
    } else {
        Write-Host "FSLogix version    : $version - older than 2210 (2.9.8361), no compaction at sign-out. Update FSLogix." -ForegroundColor Yellow
        $problems++
    }

    # Absent means enabled: compaction is on by default.
    $compact = (Get-ItemProperty 'HKLM:\SOFTWARE\FSLogix\Apps' -ErrorAction SilentlyContinue).VHDCompactDisk
    if ($null -eq $compact -or $compact -ne 0) {
        Write-Host ("VHDCompactDisk     : {0}" -f $(if ($null -eq $compact) { 'not set (default: on)' } else { "$compact (on)" })) -ForegroundColor Green
    } else {
        Write-Host "VHDCompactDisk     : 0 - compaction switched off (HKLM\SOFTWARE\FSLogix\Apps)" -ForegroundColor Yellow
        $problems++
    }

    # Compaction asks defragsvc for the minimum partition size; Disabled stops it.
    $defrag = Get-CimInstance Win32_Service -Filter "Name='defragsvc'" -ErrorAction SilentlyContinue
    if (-not $defrag) {
        Write-Host "Optimize Drives    : service defragsvc not found - compaction cannot run" -ForegroundColor Yellow
        $problems++
    } elseif ($defrag.StartMode -eq 'Disabled') {
        Write-Host "Optimize Drives    : defragsvc is Disabled - compaction cannot run. Set it to Manual." -ForegroundColor Yellow
        $problems++
    } else {
        Write-Host "Optimize Drives    : defragsvc $($defrag.StartMode) - ok" -ForegroundColor Green
    }

    # Fixed-size disks are never compacted, by FSLogix or by Invoke-FslShrinkDisk.
    foreach ($key in 'HKLM:\SOFTWARE\FSLogix\Profiles', 'HKLM:\SOFTWARE\Policies\FSLogix\ODFC') {
        if (-not (Test-Path $key)) { continue }
        $cfg = Get-ItemProperty $key
        if ($cfg.Enabled -ne 1) { continue }
        $name = Split-Path $key -Leaf
        if ($null -ne $cfg.IsDynamic -and $cfg.IsDynamic -eq 0) {
            Write-Host ("{0,-19}: IsDynamic = 0 - fixed-size disks cannot be compacted" -f $name) -ForegroundColor Yellow
            $problems++
        } else {
            Write-Host ("{0,-19}: dynamic disks - ok" -f $name) -ForegroundColor Green
        }
    }

    Write-Host ""
    if ($problems) {
        Write-Host "$problems issue(s): FSLogix will not compact (all) disks at sign-out on this host." -ForegroundColor Yellow
        exit 1
    }
    Write-Host "FSLogix compacts containers at sign-out on this host when at least 20% can be won." -ForegroundColor Green
    Write-Host "A manual shrink only catches up disks of users who rarely sign out, or below that threshold."
    exit 0
}

# ---------------------------------------------------------------------------
Write-Section "Invoke-FslShrinkDisk"
# ---------------------------------------------------------------------------
$toolScript = Join-Path $ToolPath 'Invoke-FslShrinkDisk.ps1'

if ($Force -or -not (Test-Path $toolScript)) {
    $zip   = Join-Path $env:TEMP "Invoke-FslShrinkDisk-$ToolCommit.zip"
    $unzip = Join-Path $env:TEMP "Invoke-FslShrinkDisk-unpack"
    Write-Host "Downloading Invoke-FslShrinkDisk at commit $($ToolCommit.Substring(0,7))..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri "https://github.com/FSLogix/Invoke-FslShrinkDisk/archive/$ToolCommit.zip" -OutFile $zip -UseBasicParsing

    if (Test-Path $unzip) { Remove-Item $unzip -Recurse -Force }
    Expand-Archive -Path $zip -DestinationPath $unzip -Force
    $src = Join-Path $unzip "Invoke-FslShrinkDisk-$ToolCommit"

    $hash = (Get-FileHash (Join-Path $src 'Invoke-FslShrinkDisk.ps1') -Algorithm SHA256).Hash
    if ($hash -ne $ToolHash) {
        Remove-Item $zip, $unzip -Recurse -Force -ErrorAction SilentlyContinue
        throw "Invoke-FslShrinkDisk.ps1 has SHA-256 $hash, expected $ToolHash. Not used."
    }

    if (Test-Path $ToolPath) { Remove-Item $ToolPath -Recurse -Force }
    New-Item -Path (Split-Path $ToolPath -Parent) -ItemType Directory -Force | Out-Null
    Move-Item -Path $src -Destination $ToolPath
    Get-ChildItem $ToolPath -Recurse -File | Unblock-File
    Remove-Item $zip, $unzip -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Installed in $ToolPath (hash verified)." -ForegroundColor Green
} else {
    $hash = (Get-FileHash $toolScript -Algorithm SHA256).Hash
    if ($hash -ne $ToolHash) {
        throw "$toolScript is not the pinned version (SHA-256 $hash). Run again with -Force."
    }
    Write-Host "Already in $ToolPath (hash verified)." -ForegroundColor Green
}

# ---------------------------------------------------------------------------
Write-Section "Containers"
# ---------------------------------------------------------------------------
Write-Host $Path
if (-not (Test-Path $Path)) {
    throw "$Path cannot be reached. Check the path and that this account has access to the share."
}

$disks = @(Get-ChildItem -Path $Path -Filter '*.vhd*' -Recurse -File |
    Where-Object { $_.Extension -in '.vhd', '.vhdx' } |
    Sort-Object Length -Descending)

if (-not $disks.Count) {
    Write-Host "No .vhd/.vhdx files found under $Path." -ForegroundColor Yellow
    exit 1
}

$disks | Select-Object @{n = 'Folder'; e = { $_.Directory.Name } }, Name,
    @{n = 'SizeGB'; e = { [math]::Round($_.Length / 1GB, 2) } },
    @{n = 'LastWrite'; e = { $_.LastWriteTime.ToString('yyyy-MM-dd') } } |
    Format-Table -AutoSize | Out-Host

$totalGB = [math]::Round(($disks | Measure-Object Length -Sum).Sum / 1GB, 2)
$eligible = @($disks | Where-Object { $_.Length -ge $IgnoreLessThanGB * 1GB }).Count
Write-Host ("{0} disk(s), {1} GB in total; {2} of {3} GB or more." -f $disks.Count, $totalGB, $eligible, $IgnoreLessThanGB)

if ($ReportOnly) { exit 0 }

# ---------------------------------------------------------------------------
Write-Section "Shrink"
# ---------------------------------------------------------------------------
if (-not (Test-IsAdmin)) {
    throw "Shrinking mounts each disk and needs an elevated session. Use -ReportOnly to only list."
}

# Invoke-FslShrinkDisk appends with Export-Csv, which fails on a missing folder.
$logDir = Split-Path $LogFilePath -Parent
if ($logDir -and -not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }

$shrinkArgs = @{
    Path             = $Path
    Recurse          = $true
    IgnoreLessThanGB = $IgnoreLessThanGB
    RatioFreeSpace   = $RatioFreeSpace
    ThrottleLimit    = $ThrottleLimit
    LogFilePath      = $LogFilePath
}
Write-Host "Shrinking $eligible disk(s) with at least $([int]($RatioFreeSpace * 100))% free inside, $ThrottleLimit at a time..."
& $toolScript @shrinkArgs

if (-not (Test-Path $LogFilePath)) {
    Write-Host "No log written to $LogFilePath - nothing was processed." -ForegroundColor Yellow
    exit 1
}

$log = @(Import-Csv $LogFilePath)
$log | Group-Object DiskState | Sort-Object Count -Descending |
    Select-Object Count, @{n = 'DiskState'; e = { $_.Name } } | Format-Table -AutoSize | Out-Host

$saved = [math]::Round(($log | Where-Object DiskState -eq 'Success' |
    ForEach-Object { [double]$_.SpaceSavedGB } | Measure-Object -Sum).Sum, 2)
Write-Host ("{0} disk(s) shrunk, {1} GB recovered. Log: {2}" -f
    @($log | Where-Object DiskState -eq 'Success').Count, $saved, $LogFilePath) -ForegroundColor Green

# Anything not in a normal state: usually a disk in use (user signed in).
$normal = 'Success', 'Ignored', 'SkippedAlreadyMinimum'
$failed = @($log | Where-Object { $_.DiskState -notin $normal -and $_.DiskState -notlike 'LessThan*FreeInsideDisk' })
if ($failed.Count) {
    Write-Host ""
    Write-Host "$($failed.Count) disk(s) not processed - often attached because the user is signed in:" -ForegroundColor Yellow
    $failed | Select-Object Name, DiskState | Format-Table -AutoSize -Wrap | Out-Host
    exit 1
}
exit 0
