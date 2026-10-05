<#
.SYNOPSIS
    Download an Office theme (.thmx) and install it for the signed-in user.

.DESCRIPTION
    Fetches the theme from -ThemeUrl into %ProgramData%\OfficeThemes and copies it to
    %APPDATA%\Microsoft\Templates\Document Themes, so it shows up under Design > Themes
    in Word, Excel and PowerPoint. Intended for Intune deployment in the user context.

    Nothing about the theme is built in: the URL is a parameter, and the file name is
    taken from the URL unless -ThemeName is given. Intune platform scripts cannot pass
    parameters, so either deploy this as a Win32 app with the parameters on the install
    command line, or set the defaults in a copy that is uploaded per customer.

.PARAMETER ThemeUrl
    Direct download URL of the .thmx file (for example a raw GitHub or a public blob
    storage link). Required.

.PARAMETER ThemeName
    File name to save the theme as, ending in .thmx. Default: the last segment of
    -ThemeUrl. This is the name Office shows in the theme picker.

.EXAMPLE
    .\Deploy-OfficeTheme.ps1 -ThemeUrl 'https://contoso.blob.core.windows.net/branding/Contoso.thmx'

.EXAMPLE
    # Win32 app install command
    powershell.exe -ExecutionPolicy Bypass -File .\Deploy-OfficeTheme.ps1 -ThemeUrl 'https://example.com/theme.thmx' -ThemeName 'Contoso 2026.thmx'
#>
[CmdletBinding()]
param (
    [string] $ThemeUrl,
    [string] $ThemeName
)

$ErrorActionPreference = 'Stop'

# Not Mandatory: under Intune there is nobody to answer the prompt, and the script
# would sit there until the timeout instead of failing with a reason.
if (-not $ThemeUrl) {
    throw 'No -ThemeUrl given. Pass the download URL of the .thmx file.'
}

if (-not $ThemeName) {
    $ThemeName = [uri]::UnescapeDataString(([uri]$ThemeUrl).Segments[-1])
}
if ([IO.Path]::GetExtension($ThemeName) -ne '.thmx') {
    throw "Theme name '$ThemeName' does not end in .thmx. Pass -ThemeName."
}

# Lokale opslag
$LocalFolder = "$env:ProgramData\OfficeThemes"
$LocalFile = Join-Path $LocalFolder $ThemeName

# Office Theme map
$OfficeFolder = Join-Path $env:APPDATA "Microsoft\Templates\Document Themes"

# Mappen aanmaken
New-Item -ItemType Directory -Path $LocalFolder -Force | Out-Null
New-Item -ItemType Directory -Path $OfficeFolder -Force | Out-Null

# Download theme
Invoke-WebRequest -Uri $ThemeUrl -OutFile $LocalFile -UseBasicParsing

# Kopiëren naar Office
Copy-Item -Path $LocalFile -Destination $OfficeFolder -Force

Write-Output "Office theme '$ThemeName' deployed."
