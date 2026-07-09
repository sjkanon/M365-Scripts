# Office Theme Colors Deployment

$ThemeName = "Test VIAS.xml"

$Source = "https://raw.githubusercontent.com/FirstITHub/M365-Scripts/main/scripts/Custom%20Scripts/Intune/Desktop/Office%20Themes/Test%20VIAS.xml"

$DestinationFolder = Join-Path $env:APPDATA "Microsoft\Templates\Document Themes\Theme Colors"
$DestinationFile = Join-Path $DestinationFolder $ThemeName

if (!(Test-Path $DestinationFolder)) {
    New-Item -ItemType Directory -Path $DestinationFolder -Force | Out-Null
}

Invoke-WebRequest -Uri $Source -OutFile $DestinationFile

Write-Output "Office color palette deployed."
