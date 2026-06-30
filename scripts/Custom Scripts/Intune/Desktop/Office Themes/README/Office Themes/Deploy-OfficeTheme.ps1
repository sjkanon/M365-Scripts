$ThemeName = "2026 Vias institute colours.thmx"

$Source = "C:\ProgramData\FirstIT\OfficeThemes\$ThemeName"

$DestinationFolder = Join-Path $env:APPDATA "Microsoft\Templates\Document Themes"
$DestinationFile = Join-Path $DestinationFolder $ThemeName

if (!(Test-Path $DestinationFolder)) {
    New-Item -ItemType Directory -Path $DestinationFolder -Force | Out-Null
}

if (Test-Path $Source) {
    Copy-Item $Source $DestinationFile -Force
    Write-Output "Office Theme deployed."
}
else {
    throw "Theme not found: $Source"
}
