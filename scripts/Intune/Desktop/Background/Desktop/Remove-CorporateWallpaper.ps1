#Requires -Version 5.1
<#
.SYNOPSIS
    Undo Set-CorporateWallpaper.ps1: remove the corporate wallpaper for every user,
    and leave the corporate lockscreen alone.

.DESCRIPTION
    Reverses each place Set-CorporateWallpaper.ps1 writes to, and only what it wrote:

      - PersonalizationCSP: the Desktop* values. The key itself stays when other values
        remain - Make-lockscreen.ps1 keeps its LockScreen* values in the same key.
      - Policies\System: Wallpaper and WallpaperStyle, but only when they point at a
        corporate wallpaper file - a policy someone else set is not ours to remove.
      - Every loaded user hive (S-1-5-21-*), the current user's HKCU when not running as
        SYSTEM, and the Default User profile: a wallpaper pointing at a corporate file is
        put back to the Windows default image, and the transcoded wallpaper cache of those
        users is cleared so the old image does not linger.
      - The corporate-background-* files in %ProgramData%\Wallpapers. The folder is
        removed only when nothing else is left in it - the lockscreen image lives there too.

    Explorer is restarted at the end so signed-in users see the change at once.
    Intended for Intune, run as SYSTEM in 64-bit PowerShell, like the Set script.

.PARAMETER WhatIf
    Show what would be removed or reset, change nothing.

.EXAMPLE
    .\Remove-CorporateWallpaper.ps1
    Remove the corporate wallpaper (as SYSTEM, via Intune or an elevated prompt).

.EXAMPLE
    .\Remove-CorporateWallpaper.ps1 -WhatIf
    List every value and file that would be touched.
#>
[CmdletBinding(SupportsShouldProcess)]
param()

$WallpaperFolder  = "$env:ProgramData\Wallpapers"
$CorporatePattern = 'corporate-background-*'
$DefaultWallpaper = "$env:SystemRoot\Web\Wallpaper\Windows\img0.jpg"
$LogFilePath      = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateWallpaper-Remove.log"

function Write-Log {
    param ([Parameter(Mandatory = $true)][string]$Message)
    Write-Output $Message
    if ($WhatIfPreference) { return }
    $logDir = Split-Path -Parent $LogFilePath
    if (-not (Test-Path -Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message" | Out-File -FilePath $LogFilePath -Append
}

# A value is ours when it points at a corporate-background-* file in the wallpaper folder.
function Test-CorporateWallpaperValue {
    param ([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $folder = [System.IO.Path]::GetDirectoryName($Value)
    $leaf   = [System.IO.Path]::GetFileName($Value)
    return ($folder -ieq $WallpaperFolder -and $leaf -like $CorporatePattern)
}

# Put a Control Panel\Desktop key back to the Windows default when it shows our wallpaper.
function Reset-DesktopKey {
    param ([string]$KeyPath, [string]$Label)
    if (-not (Test-Path -Path $KeyPath)) { return $false }
    $current = (Get-ItemProperty -Path $KeyPath -Name Wallpaper -ErrorAction SilentlyContinue).Wallpaper
    if (-not (Test-CorporateWallpaperValue $current)) { return $false }
    $replacement = if (Test-Path -Path $DefaultWallpaper) { $DefaultWallpaper } else { '' }
    if ($PSCmdlet.ShouldProcess("$Label ($KeyPath)", "reset wallpaper to '$replacement'")) {
        Set-ItemProperty -Path $KeyPath -Name Wallpaper      -Value $replacement -Force
        Set-ItemProperty -Path $KeyPath -Name WallpaperStyle -Value '10'         -Force
        Set-ItemProperty -Path $KeyPath -Name TileWallpaper  -Value '0'          -Force
        # Out-Default, or the message would become part of this function's return value
        Write-Log "Wallpaper reset for $Label" | Out-Default
    }
    return $true
}

Write-Log "====== Start Remove-CorporateWallpaper ======"

# 1. PersonalizationCSP - only the desktop values; the lockscreen values share this key.
$cspPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
if (Test-Path -Path $cspPath) {
    foreach ($name in 'DesktopImagePath', 'DesktopImageUrl', 'DesktopImageStatus') {
        if ($null -ne (Get-ItemProperty -Path $cspPath -Name $name -ErrorAction SilentlyContinue)) {
            if ($PSCmdlet.ShouldProcess("$cspPath\$name", 'remove value')) {
                Remove-ItemProperty -Path $cspPath -Name $name -Force
                Write-Log "Removed PersonalizationCSP value $name"
            }
        }
    }
    $remaining = @((Get-Item -Path $cspPath).Property)
    if ($remaining.Count -eq 0 -and @(Get-ChildItem -Path $cspPath).Count -eq 0) {
        if ($PSCmdlet.ShouldProcess($cspPath, 'remove empty key')) {
            Remove-Item -Path $cspPath -Force
            Write-Log "Removed empty PersonalizationCSP key"
        }
    } elseif ($remaining.Count) {
        Write-Log "PersonalizationCSP kept - other values remain: $($remaining -join ', ')"
    }
} else {
    Write-Log "No PersonalizationCSP key found"
}

# 2. Machine policy fallback - only when it points at our file.
$policyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
$policyWallpaper = (Get-ItemProperty -Path $policyPath -Name Wallpaper -ErrorAction SilentlyContinue).Wallpaper
if (Test-CorporateWallpaperValue $policyWallpaper) {
    if ($PSCmdlet.ShouldProcess("$policyPath\Wallpaper, WallpaperStyle", 'remove values')) {
        Remove-ItemProperty -Path $policyPath -Name Wallpaper      -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $policyPath -Name WallpaperStyle -Force -ErrorAction SilentlyContinue
        Write-Log "Removed machine policy wallpaper values"
    }
} elseif ($policyWallpaper) {
    Write-Log "Machine policy wallpaper left alone - it is not the corporate wallpaper: $policyWallpaper"
}

# 3. Loaded user hives, and their transcoded wallpaper cache.
$profileList = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue |
    Where-Object { $_.PSChildName -match '^S-1-5-21-.+-\d+$' } |
    ForEach-Object {
        $sid = $_.PSChildName
        if (-not (Reset-DesktopKey -KeyPath "Registry::HKEY_USERS\$sid\Control Panel\Desktop" -Label "user $sid")) { return }
        $profilePath = (Get-ItemProperty -Path (Join-Path $profileList $sid) -Name ProfileImagePath -ErrorAction SilentlyContinue).ProfileImagePath
        if (-not $profilePath) { return }
        $themes = Join-Path $profilePath 'AppData\Roaming\Microsoft\Windows\Themes'
        $cached = @(Get-ChildItem -Path $themes -Filter 'TranscodedWallpaper*' -File -ErrorAction SilentlyContinue) +
                  @(Get-ChildItem -Path (Join-Path $themes 'CachedFiles') -File -ErrorAction SilentlyContinue)
        if ($cached.Count -and $PSCmdlet.ShouldProcess($themes, "clear $($cached.Count) cached wallpaper file(s)")) {
            $cached | Remove-Item -Force -ErrorAction SilentlyContinue
            Write-Log "Theme cache cleared for $sid"
        }
    }

# 4. The current user's own hive - when run as SYSTEM, HKCU is SYSTEM's and step 3 covered the users.
$isSystem = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).IsSystem
if (-not $isSystem) {
    [void](Reset-DesktopKey -KeyPath 'HKCU:\Control Panel\Desktop' -Label 'current user')
}

# 5. Default User profile - accounts created from now on.
$defaultHive = "$env:SystemDrive\Users\Default\NTUSER.DAT"
$tempHive    = 'HKLM\DefaultUserTemp'
if (Test-Path -Path $defaultHive) {
    reg load $tempHive $defaultHive 2>$null | Out-Null
}
if ((Test-Path -Path $defaultHive) -and $LASTEXITCODE -ne 0) {
    Write-Log "Default User profile skipped - its hive could not be loaded (needs SYSTEM or an elevated prompt)"
} elseif (Test-Path -Path $defaultHive) {
    try {
        [void](Reset-DesktopKey -KeyPath "Registry::$tempHive\Control Panel\Desktop" -Label 'Default User profile')
    } finally {
        [gc]::Collect()
        Start-Sleep -Seconds 1
        reg unload $tempHive | Out-Null
    }
}

# 6. The wallpaper files - not the folder while the lockscreen image is still in it.
if (Test-Path -Path $WallpaperFolder) {
    $files = @(Get-ChildItem -Path $WallpaperFolder -Filter $CorporatePattern -File -ErrorAction SilentlyContinue)
    foreach ($file in $files) {
        if ($PSCmdlet.ShouldProcess($file.FullName, 'delete')) {
            Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
            Write-Log "Removed $($file.Name)"
        }
    }
    $left = @(Get-ChildItem -Path $WallpaperFolder -Force -ErrorAction SilentlyContinue)
    if ($left.Count -eq 0) {
        if ($PSCmdlet.ShouldProcess($WallpaperFolder, 'remove empty folder')) {
            Remove-Item -Path $WallpaperFolder -Force
            Write-Log "Removed empty wallpaper folder"
        }
    } elseif (-not $WhatIfPreference) {
        Write-Log "Wallpaper folder kept - still holds: $(($left | ForEach-Object Name) -join ', ')"
    }
} else {
    Write-Log "No wallpaper folder found"
}

# 7. Restart Explorer so signed-in users see it now.
if ($PSCmdlet.ShouldProcess('explorer.exe', 'restart')) {
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
}

Write-Log "====== Remove-CorporateWallpaper completed ======"
exit 0
