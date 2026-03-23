$AppInstallerVersion = Get-AppxPackage Microsoft.DesktopAppInstaller | Select Version
$Folder = 'C:\temp'
"Test to see if folder [$Folder]  exists"
if (Test-Path -Path $Folder) {
 Write-Host "Microsoft Desktop App Installer Version:" $AppInstallerVersion

if ($AppInstallerVersion -match "1.1")
{
    Write-Host "Microsoft Desktop App Installer is equal or higher to the version needed for this script to work, continuing the script."
    Remove-Item "C:\temp\deletedefaultapps.ps1" -Force:$false
    Invoke-WebRequest -Uri "https://endpoint.eoo.cloud/apppackages/deletedefaultapps.ps1" -OutFile "C:\temp\deletedefaultapps.ps1"
    Set-Location "C:\temp"
    .\deletedefaultapps.ps1
}
else
{
Write-Host "Microsoft Desktop App Installer does not meet the minimum version to run this script, a window will now appear to install the version required for this script the run. Click on Update to install it.
Once installed, press the Enter key in the script to continue."
mkdir c:\temp\

Invoke-WebRequest -Uri “https://github.com/microsoft/winget-cli/releases/download/v1.1.12653/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle” -OutFile “C:\temp\WinGet.appxbundle”
Remove-Item "C:\temp\deletedefaultapps.ps1" -Force:$false
Add-AppxPackage “C:\temp\WinGet.appxbundle”
 Invoke-WebRequest -Uri "https://endpoint.eoo.cloud/apppackages/deletedefaultapps.ps1" -OutFile "C:\temp\deletedefaultapps.ps1"
 Set-Location C:\temp
 .\deletedefaultapps.ps1}
 }
 else {
  if ($AppInstallerVersion -match "1.1")
{
 Write-Host "Microsoft Desktop App Installer is equal or higher to the version needed for this script to work, continuing the script."
 mkdir C:\temp
 Invoke-WebRequest -Uri "https://endpoint.eoo.cloud/apppackages/deletedefaultapps.ps1" -OutFile "C:\temp\deletedefaultapps.ps1"
 Set-Location C:\temp
 .\deletedefaultapps.ps1
}
else
{
Write-Host "Microsoft Desktop App Installer does not meet the minimum version to run this script, a window will now appear to install the version required for this script the run. Click on Update to install it.
Once installed, press the Enter key in the script to continue."
mkdir c:\temp\

Invoke-WebRequest -Uri “https://github.com/microsoft/winget-cli/releases/download/v1.1.12653/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle” -OutFile “C:\temp\WinGet.appxbundle”

Add-AppxPackage “C:\temp\WinGet.appxbundle”
 Invoke-WebRequest -Uri "https://endpoint.eoo.cloud/apppackages/deletedefaultapps.ps1" -OutFile "C:\temp\deletedefaultapps.ps1"
 Set-Location C:\temp
 .\deletedefaultapps.ps1}
}
Set-Location C:\
rmdir C:\temp -Confirm:$false -Recurse:$true