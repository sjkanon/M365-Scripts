#Requires -Version 5.1
<#
.SYNOPSIS
    Create a .url shortcut to a web address on the Public Desktop.

.DESCRIPTION
    Generalized replacement for a script that hardcoded a single customer's web-app
    URL and shortcut name. Supply -Name and -Url for whatever site you need a desktop
    shortcut to (a web app, an internal portal, a SharePoint site, ...).

.PARAMETER Name
    Shortcut file name, without extension.

.PARAMETER Url
    Target URL.

.PARAMETER Apply
    Actually create the shortcut. Without this switch, the script only reports what
    it would do.

.EXAMPLE
    .\New-DesktopUrlShortcut.ps1 -Name "Company Portal" -Url "https://portal.contoso.com/"

.EXAMPLE
    .\New-DesktopUrlShortcut.ps1 -Name "Company Portal" -Url "https://portal.contoso.com/" -Apply
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $Name,

    [Parameter(Mandatory)]
    [string] $Url,

    [switch] $Apply
)

$targetPath = Join-Path "$env:Public\Desktop" "$Name.url"

Write-Host ""
Write-Host "  New-DesktopUrlShortcut : $Name -> $Url" -ForegroundColor Cyan
Write-Host "  Mode : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })

if (-not $Apply) {
    Write-Host "  Would create: $targetPath" -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to create the shortcut." -ForegroundColor Yellow
    exit 0
}

if (-not $PSCmdlet.ShouldProcess($targetPath, "Create URL shortcut")) { exit 0 }

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($targetPath)
$shortcut.TargetPath = $Url
$shortcut.Save()

Write-Host "  [OK]   Created: $targetPath" -ForegroundColor Green
