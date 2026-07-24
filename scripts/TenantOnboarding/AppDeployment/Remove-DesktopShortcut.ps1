#Requires -Version 5.1
<#
.SYNOPSIS
    Remove desktop shortcuts matching a name pattern, from the Public Desktop or the current user's desktop.

.DESCRIPTION
    Generalized replacement for two one-off scripts that each removed a single
    hardcoded shortcut left behind post-install (an installer leftover icon, a stale
    RDP connection shortcut). Supply your own -NamePattern.

.PARAMETER NamePattern
    Wildcard pattern (PowerShell -like syntax) matched against shortcut file names,
    e.g. "Offline*" or "*.rdp".

.PARAMETER Scope
    "PublicDesktop" (default, all users) or "CurrentUserDesktop".

.PARAMETER Apply
    Actually delete matching shortcuts. Without this switch, the script only reports
    what would be removed.

.EXAMPLE
    .\Remove-DesktopShortcut.ps1 -NamePattern "Offline*"

.EXAMPLE
    .\Remove-DesktopShortcut.ps1 -NamePattern "*.rdp" -Apply
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $NamePattern,

    [ValidateSet('PublicDesktop', 'CurrentUserDesktop')]
    [string] $Scope = 'PublicDesktop',

    [switch] $Apply
)

$desktopPath = if ($Scope -eq 'PublicDesktop') { "$env:Public\Desktop" } else { [Environment]::GetFolderPath('Desktop') }

Write-Host ""
Write-Host "  Remove-DesktopShortcut : '$NamePattern' in $desktopPath" -ForegroundColor Cyan
Write-Host "  Mode : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

$matches = Get-ChildItem -Path $desktopPath -Filter $NamePattern -ErrorAction SilentlyContinue

if ($matches.Count -eq 0) {
    Write-Host "  No matching items found." -ForegroundColor DarkGray
    exit 0
}

foreach ($item in $matches) {
    if (-not $Apply) {
        Write-Host "  [PREVIEW] Would remove: $($item.FullName)" -ForegroundColor Yellow
        continue
    }
    if (-not $PSCmdlet.ShouldProcess($item.FullName, "Remove")) { continue }
    Remove-Item -Path $item.FullName -Force
    Write-Host "  [OK]   Removed: $($item.FullName)" -ForegroundColor Green
}

Write-Host ""
if (-not $Apply) { Write-Host "  Re-run with -Apply to remove the item(s) shown above." -ForegroundColor Yellow }
Write-Host ""
