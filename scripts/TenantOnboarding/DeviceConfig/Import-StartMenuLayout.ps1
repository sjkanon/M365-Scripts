#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Apply a Start Menu layout XML to the local device.

.DESCRIPTION
    Downloads (if -LayoutUrl is given) or uses a local -LayoutXmlPath Start Menu
    layout XML file and applies it via Import-StartLayout. Generalized replacement
    for a script that hardcoded an internal download URL for the layout file.

.PARAMETER LayoutXmlPath
    Local path to a Start Menu layout XML file. Use this or -LayoutUrl.

.PARAMETER LayoutUrl
    URL to download the layout XML from before applying it. Use this or -LayoutXmlPath.

.PARAMETER Apply
    Actually apply the layout. Without this switch, the script only reports what it
    would do.

.EXAMPLE
    .\Import-StartMenuLayout.ps1 -LayoutXmlPath "C:\Deploy\startmenu.xml" -Apply

.EXAMPLE
    .\Import-StartMenuLayout.ps1 -LayoutUrl "https://packages.contoso.com/config/startmenu.xml" -Apply
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $LayoutXmlPath,
    [string] $LayoutUrl,
    [switch] $Apply
)

if (-not $LayoutXmlPath -and -not $LayoutUrl) {
    throw "Provide either -LayoutXmlPath or -LayoutUrl."
}

Write-Host ""
Write-Host "  Import-StartMenuLayout" -ForegroundColor Cyan
Write-Host "  Mode : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

if (-not $Apply) {
    if ($LayoutUrl) { Write-Host "  Would download layout from $LayoutUrl" -ForegroundColor Yellow }
    Write-Host "  Would run Import-StartLayout against $(if ($LayoutXmlPath) { $LayoutXmlPath } else { '<downloaded file>' })" -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to apply the layout." -ForegroundColor Yellow
    Write-Host ""
    exit 0
}

if (-not $PSCmdlet.ShouldProcess('Start Menu layout', "Apply")) { exit 0 }

$xmlPath = $LayoutXmlPath
if ($LayoutUrl) {
    $localDir = Join-Path $env:ProgramData 'TenantOnboarding\StartMenu'
    if (-not (Test-Path $localDir)) { New-Item -Path $localDir -ItemType Directory -Force | Out-Null }
    $xmlPath = Join-Path $localDir 'startmenu.xml'
    Invoke-WebRequest -Uri $LayoutUrl -OutFile $xmlPath -ErrorAction Stop
    Write-Host "  [OK]   Downloaded layout to $xmlPath" -ForegroundColor Green
}

Import-StartLayout -LayoutPath $xmlPath -MountPath "$env:SystemDrive\" -ErrorAction Stop
Write-Host "  [OK]   Start Menu layout applied." -ForegroundColor Green
Write-Host ""
