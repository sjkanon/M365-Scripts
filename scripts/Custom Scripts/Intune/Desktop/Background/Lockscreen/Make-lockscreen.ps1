#Requires -Version 5.1
# ==============================================================================
# Make-lockscreen.ps1
# Version 2.0 - Uses the same corporate wallpaper source with validated internet download and logging
#
# Usage:
#   Only change the variables in the CONFIGURATION block below.
#   The rest of the script does not need to be modified.
#
# Deploy via Intune:
#   - Type: PowerShell script or Win32 app
#   - Run as: SYSTEM
#   - Run in 64-bit PowerShell: Yes
#
# Changelog:
#   1.0 - Original lockscreen script using direct WebClient download
#   2.0 - Reworked to match corporate wallpaper configuration, added URL normalization,
#         image validation, HTML detection, structured logging, and safer download handling
#   2.1 - Added explicit lockscreen refresh step (ShellExperienceHost/LockApp) for faster apply
# ==============================================================================
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# ==============================================================================
# CONFIGURATION — change these per customer
# ==============================================================================

# URL to the lockscreen image (PNG or JPG) - same as Set-CorporateWallpaper.ps1
$ImageUrl = "https://raw.githubusercontent.com/sjkanon/Wallpapers/refs/heads/main/HRL/LOCKSCREEN.PNG"
$ClientName = "HRL"

# ==============================================================================
# INTERNAL VARIABLES
# ==============================================================================

$RegKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
$LockScreenPath = "LockScreenImagePath"
$LockScreenStatus = "LockScreenImageStatus"
$LockScreenUrl = "LockScreenImageUrl"
$StatusValue = "1"

$LockScreenFolder = "$env:ProgramData\Wallpapers"
$LockScreenBaseName = "corporate-lockscreen-$($ClientName.ToLower())"
$tempLockScreenPath = Join-Path -Path $LockScreenFolder -ChildPath "$LockScreenBaseName.download"
$LogFilePath = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateLockscreen-$($ClientName.ToUpper()).log"

# ==============================================================================
# FUNCTIONS
# ==============================================================================

function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message
    )
    $logDir = Split-Path -Parent $LogFilePath
    if (-not (Test-Path -Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $LogFilePath -Append
}

function Get-ImageTypeFromHeader {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    $bytes = [System.IO.File]::ReadAllBytes($FilePath)
    if ($bytes.Length -lt 8) {
        return $null
    }

    # JPEG: FF D8 FF
    if ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xD8 -and $bytes[2] -eq 0xFF) {
        return "jpg"
    }

    # PNG: 89 50 4E 47 0D 0A 1A 0A
    if ($bytes[0] -eq 0x89 -and $bytes[1] -eq 0x50 -and $bytes[2] -eq 0x4E -and $bytes[3] -eq 0x47 -and
        $bytes[4] -eq 0x0D -and $bytes[5] -eq 0x0A -and $bytes[6] -eq 0x1A -and $bytes[7] -eq 0x0A) {
        return "png"
    }

    # BMP: 42 4D
    if ($bytes[0] -eq 0x42 -and $bytes[1] -eq 0x4D) {
        return "bmp"
    }

    return $null
}

function Test-FileLooksLikeHtml {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    $bytes = [System.IO.File]::ReadAllBytes($FilePath)
    if ($bytes.Length -eq 0) {
        return $true
    }

    $sampleLength = [Math]::Min(512, $bytes.Length)
    $sampleBytes  = $bytes[0..($sampleLength - 1)]
    $sampleText   = [System.Text.Encoding]::ASCII.GetString($sampleBytes)
    return ($sampleText -match '(?i)<!doctype\s+html|<html|<head|<body')
}

function Resolve-DownloadUrl {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    try {
        $uri = [System.Uri]$Url
    } catch {
        return $Url
    }

    if ($uri.Host -ieq "github.com") {
        if ($uri.AbsolutePath -match '^/([^/]+)/([^/]+)/blob/(.+)$') {
            return "https://raw.githubusercontent.com/$($matches[1])/$($matches[2])/$($matches[3])"
        }

        if ($uri.AbsolutePath -match '^/([^/]+)/([^/]+)/raw/(.+)$') {
            return "https://raw.githubusercontent.com/$($matches[1])/$($matches[2])/$($matches[3])"
        }
    }

    return $Url
}

function Invoke-LockscreenRefresh {
    # Restart lockscreen-related shell processes so the new image is picked up faster.
    $processNames = @("ShellExperienceHost", "LockApp")

    foreach ($processName in $processNames) {
        try {
            $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
            if ($processes) {
                $processes | Stop-Process -Force -ErrorAction SilentlyContinue
                Write-Log "Refreshed process: $processName"
            } else {
                Write-Log "Process not running (skip refresh): $processName"
            }
        } catch {
            Write-Log "WARNING: Could not refresh process $processName : $_"
        }
    }
}

# ==============================================================================
# SCRIPT START
# ==============================================================================

Write-Log "====== Start Make-Lockscreen for client: $ClientName ======"
Write-Log "Source URL : $ImageUrl"

# Step 1: Create target folder
if (-not (Test-Path -Path $LockScreenFolder)) {
    try {
        New-Item -ItemType Directory -Path $LockScreenFolder -Force | Out-Null
        Write-Log "Lockscreen folder created: $LockScreenFolder"
    } catch {
        Write-Log "ERROR creating folder: $_"
        exit 1
    }
}

# Step 2: Download image from internet with validation
$resolvedImageUrl = Resolve-DownloadUrl -Url $ImageUrl

if ($resolvedImageUrl -ne $ImageUrl) {
    Write-Log "Source URL normalized for raw download: $resolvedImageUrl"
}

try {
    Write-Log "Downloading lockscreen image from: $resolvedImageUrl"
    $downloadResponse = Invoke-WebRequest -Uri $resolvedImageUrl -OutFile $tempLockScreenPath -UseBasicParsing -MaximumRedirection 10 -Headers @{ "Accept" = "image/*,*/*;q=0.8"; "User-Agent" = "CorporateLockscreenScript/1.0" }
    Write-Log "Image downloaded to temporary path: $tempLockScreenPath"
} catch {
    Write-Log "ERROR downloading image: $_"
    exit 1
}

if (-not (Test-Path -Path $tempLockScreenPath)) {
    Write-Log "ERROR: File not present after download."
    exit 1
}

$downloadedFile = Get-Item -Path $tempLockScreenPath -ErrorAction SilentlyContinue
if (-not $downloadedFile -or $downloadedFile.Length -lt 1024) {
    $fileLength = if ($downloadedFile) { $downloadedFile.Length } else { 0 }
    Write-Log "ERROR: Downloaded file is too small to be a valid lockscreen ($fileLength bytes)."
    exit 1
}

if (Test-FileLooksLikeHtml -FilePath $tempLockScreenPath) {
    Write-Log "ERROR: Downloaded content appears to be HTML instead of an image (common with non-raw GitHub URLs)."
    exit 1
}

$imageType = Get-ImageTypeFromHeader -FilePath $tempLockScreenPath
if (-not $imageType) {
    Write-Log "ERROR: Downloaded file is not a supported image type (expected JPEG/PNG/BMP)."
    exit 1
}

$LockScreenImageValue = Join-Path -Path $LockScreenFolder -ChildPath "$LockScreenBaseName.$imageType"

try {
    Move-Item -Path $tempLockScreenPath -Destination $LockScreenImageValue -Force
    Write-Log "Validated image type: $imageType"
    Write-Log "Final lockscreen path: $LockScreenImageValue"
} catch {
    Write-Log "ERROR finalizing lockscreen file: $_"
    exit 1
}

# Step 3: Apply lockscreen via registry (PersonalizationCSP)
if (!(Test-Path $RegKeyPath)) {
    try {
        Write-Log "Creating registry path $RegKeyPath"
        New-Item -Path $RegKeyPath -Force | Out-Null
    } catch {
        Write-Log "ERROR creating registry path: $_"
        exit 1
    }
}

try {
    New-ItemProperty -Path $RegKeyPath -Name $LockScreenStatus -Value $StatusValue -PropertyType DWORD -Force | Out-Null
    New-ItemProperty -Path $RegKeyPath -Name $LockScreenPath -Value $LockScreenImageValue -PropertyType STRING -Force | Out-Null
    New-ItemProperty -Path $RegKeyPath -Name $LockScreenUrl -Value $resolvedImageUrl -PropertyType STRING -Force | Out-Null
    Write-Log "Registry keys updated successfully"
} catch {
    Write-Log "ERROR updating registry: $_"
    exit 1
}

# Step 4: Apply lockscreen immediately
try {
    RUNDLL32.EXE USER32.DLL, UpdatePerUserSystemParameters 1, True
    Write-Log "Lockscreen apply signal sent successfully"
} catch {
    Write-Log "WARNING: Could not apply immediate update: $_"
}

# Step 5: Force a lockscreen shell refresh for active sessions
Invoke-LockscreenRefresh

Write-Log "====== Make-Lockscreen completed successfully ======"