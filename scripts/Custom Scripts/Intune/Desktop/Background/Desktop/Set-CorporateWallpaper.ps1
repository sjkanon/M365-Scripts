#Requires -Version 5.1
# ==============================================================================
# Set-CorporateWallpaper.ps1
# Version 2.2 - GitHub raw URL handling and stricter download validation
#
# Usage:
#   Only change the variables in the CONFIGURATION block below.
#   The rest of the script does not need to be modified.
#
# Deploy via Intune:
#   - Type: PowerShell script
#   - Run as: SYSTEM
#   - Run in 64-bit PowerShell: Yes
# ==============================================================================
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
# ==============================================================================
# CONFIGURATION — change these per customer
# ==============================================================================

# URL to the wallpaper image (PNG or JPG)
# Use a publicly accessible URL hosted by or on behalf of the customer
$ImageUrl = "https://raw.githubusercontent.com/FirstITHub/Wallpaper/refs/heads/main/Onco3R/Spring.png"

# Display style:
#   10 = Fill    (recommended — fills screen without distortion)
#    6 = Fit     (fits within screen, black borders possible)
#    2 = Stretch (stretches to fill, may distort)
#    0 = Tile
#   22 = Span    (spreads across multiple monitors)
$WallpaperStyle = "10"

# Customer name — used in log filename and local image filename
$ClientName = "Onco3R"

# ==============================================================================
# INTERNAL VARIABLES — do not modify
# ==============================================================================

$WallpaperFolder   = "$env:ProgramData\Wallpapers"
$WallpaperBaseName = "corporate-background-$($ClientName.ToLower())"
$WallpaperPath     = $null
$LogFilePath       = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateWallpaper-$($ClientName.ToUpper()).log"

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

# ==============================================================================
# SCRIPT START
# ==============================================================================

Write-Log "====== Start Set-CorporateWallpaper for client: $ClientName ======"
Write-Log "Source URL : $ImageUrl"
Write-Log "Style      : $WallpaperStyle"

# Step 1: Create target folder
if (-not (Test-Path -Path $WallpaperFolder)) {
    try {
        New-Item -ItemType Directory -Path $WallpaperFolder -Force | Out-Null
        Write-Log "Wallpaper folder created: $WallpaperFolder"
    } catch {
        Write-Log "ERROR creating folder: $_"
        exit 1
    }
}

# Step 2: Download image to temporary file and validate
$tempWallpaperPath = Join-Path -Path $WallpaperFolder -ChildPath "$WallpaperBaseName.download"
$resolvedImageUrl  = Resolve-DownloadUrl -Url $ImageUrl

if ($resolvedImageUrl -ne $ImageUrl) {
    Write-Log "Source URL normalized for raw download: $resolvedImageUrl"
}

try {
    $downloadResponse = Invoke-WebRequest -Uri $resolvedImageUrl -OutFile $tempWallpaperPath -UseBasicParsing -MaximumRedirection 10 -Headers @{ "Accept" = "image/*,*/*;q=0.8"; "User-Agent" = "CorporateWallpaperScript/2.2" }
    Write-Log "Image downloaded to temporary path: $tempWallpaperPath"
    if ($downloadResponse -and $downloadResponse.BaseResponse -and $downloadResponse.BaseResponse.ResponseUri) {
        Write-Log "Final response URI: $($downloadResponse.BaseResponse.ResponseUri.AbsoluteUri)"
    }
} catch {
    Write-Log "ERROR downloading image: $_"
    exit 1
}

if (-not (Test-Path -Path $tempWallpaperPath)) {
    Write-Log "ERROR: File not present after download."
    exit 1
}

$downloadedFile = Get-Item -Path $tempWallpaperPath -ErrorAction SilentlyContinue
if (-not $downloadedFile -or $downloadedFile.Length -lt 1024) {
    $fileLength = if ($downloadedFile) { $downloadedFile.Length } else { 0 }
    Write-Log "ERROR: Downloaded file is too small to be a valid wallpaper ($fileLength bytes)."
    exit 1
}

if (Test-FileLooksLikeHtml -FilePath $tempWallpaperPath) {
    Write-Log "ERROR: Downloaded content appears to be HTML instead of an image (common with non-raw GitHub URLs)."
    exit 1
}

$imageType = Get-ImageTypeFromHeader -FilePath $tempWallpaperPath
if (-not $imageType) {
    Write-Log "ERROR: Downloaded file is not a supported image type (expected JPEG/PNG/BMP)."
    exit 1
}

$WallpaperPath = Join-Path -Path $WallpaperFolder -ChildPath "$WallpaperBaseName.$imageType"

try {
    Get-ChildItem -Path $WallpaperFolder -Filter "$WallpaperBaseName.*" -File -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -ne $WallpaperPath } |
        Remove-Item -Force -ErrorAction SilentlyContinue

    Move-Item -Path $tempWallpaperPath -Destination $WallpaperPath -Force
    Write-Log "Validated image type: $imageType"
    Write-Log "Final wallpaper path: $WallpaperPath"
} catch {
    Write-Log "ERROR finalizing wallpaper file: $_"
    exit 1
}

# Step 3: Load Windows API
$wallpaperApiTypeDefinition = @'
using System;
using System.Runtime.InteropServices;
public class Wallpaper {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
'@
Add-Type -TypeDefinition $wallpaperApiTypeDefinition -ErrorAction SilentlyContinue

$SPI_SETDESKWALLPAPER = 0x0014
$SPIF_UPDATEINIFILE   = 0x01
$SPIF_SENDCHANGE      = 0x02

# Step 4: PersonalizationCSP (MDM/Intune - enforces wallpaper for all users)
try {
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
    if (-not (Test-Path -Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    New-ItemProperty -Path $regPath -Name "DesktopImagePath"   -Value $WallpaperPath -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $regPath -Name "DesktopImageUrl"    -Value $WallpaperPath -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $regPath -Name "DesktopImageStatus" -Value 1              -PropertyType DWord  -Force | Out-Null
    Write-Log "PersonalizationCSP registry keys updated"
} catch {
    Write-Log "ERROR setting PersonalizationCSP: $_"
}

# Step 5: Current user - WinAPI (applies immediately)
try {
    $result = [Wallpaper]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $WallpaperPath, $SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE)
    if ($result) {
        Write-Log "Wallpaper applied via WinAPI for current user"
    } else {
        Write-Log "WinAPI returned false - non-critical, PersonalizationCSP will apply the wallpaper"
    }
} catch {
    Write-Log "ERROR applying wallpaper via WinAPI: $_"
}

# Step 6: Current user - HKCU registry (persists style setting)
try {
    $regPath = "HKCU:\Control Panel\Desktop"
    Set-ItemProperty -Path $regPath -Name "Wallpaper"      -Value $WallpaperPath -Force
    Set-ItemProperty -Path $regPath -Name "WallpaperStyle" -Value $WallpaperStyle -Force
    Set-ItemProperty -Path $regPath -Name "TileWallpaper"  -Value "0"            -Force
    RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters 1, $true
    Write-Log "HKCU registry keys updated (style: $WallpaperStyle)"
} catch {
    Write-Log "ERROR updating HKCU registry: $_"
}

# Step 7: Default User profile (applies to new user accounts created after deployment)
try {
    $defaultUserRegPath = "$env:SystemDrive\Users\Default\NTUSER.DAT"
    $tempRegPath        = "HKLM\DefaultUserTemp"

    if (Test-Path -Path $defaultUserRegPath) {
        reg load $tempRegPath $defaultUserRegPath | Out-Null
        reg add "$tempRegPath\Control Panel\Desktop" /v Wallpaper      /t REG_SZ /d $WallpaperPath  /f | Out-Null
        reg add "$tempRegPath\Control Panel\Desktop" /v WallpaperStyle /t REG_SZ /d $WallpaperStyle /f | Out-Null
        reg add "$tempRegPath\Control Panel\Desktop" /v TileWallpaper  /t REG_SZ /d 0              /f | Out-Null
        [gc]::Collect()
        Start-Sleep -Seconds 1
        reg unload $tempRegPath | Out-Null
        Write-Log "Default User profile updated for new user accounts"
    } else {
        Write-Log "Default User NTUSER.DAT not found - step skipped"
    }
} catch {
    Write-Log "ERROR updating Default User profile: $_"
    try { reg unload $tempRegPath | Out-Null } catch { }
}

Write-Log "====== Script completed for client: $ClientName ======"
exit 0
