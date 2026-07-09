# Remove-CorporateWallpaper.ps1

Write-Output "=== Removing corporate wallpaper ==="

# 1. Remove PersonalizationCSP (machine-wide enforced wallpaper)
$CSPPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
if (Test-Path $CSPPath) {
    Remove-Item -Path $CSPPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Output "Removed PersonalizationCSP registry keys"
} else {
    Write-Output "No PersonalizationCSP keys found"
}

# 2. Reset current user wallpaper settings
$HKCUPath = "HKCU:\Control Panel\Desktop"

Set-ItemProperty -Path $HKCUPath -Name Wallpaper -Value "" -ErrorAction SilentlyContinue
Set-ItemProperty -Path $HKCUPath -Name WallpaperStyle -Value "10" -ErrorAction SilentlyContinue
Set-ItemProperty -Path $HKCUPath -Name TileWallpaper -Value "0" -ErrorAction SilentlyContinue

Write-Output "Reset HKCU wallpaper settings"

# 3. Remove wallpaper files
$WallpaperFolder = "C:\ProgramData\Wallpapers"
if (Test-Path $WallpaperFolder) {
    Remove-Item -Path $WallpaperFolder -Recurse -Force -ErrorAction SilentlyContinue
    Write-Output "Removed wallpaper files"
} else {
    Write-Output "No wallpaper folder found"
}

# 4. OPTIONAL: Reset Default User (voor nieuwe accounts)
$DefaultUserHive = "C:\Users\Default\NTUSER.DAT"
if (Test-Path $DefaultUserHive) {
    reg load HKU\TempDefault $DefaultUserHive | Out-Null

    Remove-ItemProperty -Path "Registry::HKU\TempDefault\Control Panel\Desktop" -Name Wallpaper -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "Registry::HKU\TempDefault\Control Panel\Desktop" -Name WallpaperStyle -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "Registry::HKU\TempDefault\Control Panel\Desktop" -Name TileWallpaper -ErrorAction SilentlyContinue

    reg unload HKU\TempDefault | Out-Null
    Write-Output "Reset Default User profile"
}

# 5. Restart explorer to apply immediately
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue

Write-Output "=== Wallpaper reset completed ==="
