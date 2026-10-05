<#
.SYNOPSIS
    Download an Office color scheme (.xml) and install it for the signed-in user.

.DESCRIPTION
    Copies a color scheme definition (<a:clrScheme>) into
    %APPDATA%\Microsoft\Templates\Document Themes\Theme Colors, so it shows up under
    Design > Colors in Word, Excel and PowerPoint. Theme colors only - for a full
    .thmx theme use Deploy-OfficeTheme.ps1 in the parent folder.

    Intune platform scripts cannot pass parameters, so either deploy this as a Win32
    app with the parameters on the install command line, or set the defaults in a
    copy that is uploaded per customer.

.PARAMETER ColorsUrl
    Direct download URL of the color scheme .xml. Required.

.PARAMETER ColorsName
    File name to save the scheme as, ending in .xml. Default: the last segment of
    -ColorsUrl.

.EXAMPLE
    .\Deploy-Officecolors.ps1 -ColorsUrl 'https://contoso.blob.core.windows.net/branding/Contoso.xml'
#>
[CmdletBinding()]
param (
    [string] $ColorsUrl,
    [string] $ColorsName
)

$ErrorActionPreference = 'Stop'

# Not Mandatory: under Intune there is nobody to answer the prompt.
if (-not $ColorsUrl) {
    throw 'No -ColorsUrl given. Pass the download URL of the color scheme .xml.'
}

if (-not $ColorsName) {
    $ColorsName = [uri]::UnescapeDataString(([uri]$ColorsUrl).Segments[-1])
}
if ([IO.Path]::GetExtension($ColorsName) -ne '.xml') {
    throw "Color scheme name '$ColorsName' does not end in .xml. Pass -ColorsName."
}

$DestinationFolder = Join-Path $env:APPDATA "Microsoft\Templates\Document Themes\Theme Colors"
$DestinationFile = Join-Path $DestinationFolder $ColorsName

if (!(Test-Path $DestinationFolder)) {
    New-Item -ItemType Directory -Path $DestinationFolder -Force | Out-Null
}

Invoke-WebRequest -Uri $ColorsUrl -OutFile $DestinationFile -UseBasicParsing

Write-Output "Office color palette '$ColorsName' deployed."
