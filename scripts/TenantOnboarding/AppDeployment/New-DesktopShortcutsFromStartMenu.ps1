#Requires -Version 5.1
<#
.SYNOPSIS
    Copy a set of Start Menu shortcuts to the Public Desktop.

.DESCRIPTION
    Generalized replacement for a script that hardcoded a fixed list of app shortcut
    names to copy from the all-users Start Menu to the Public Desktop (so every user
    of a shared/kiosk machine gets the same desktop icons). Supply your own -AppNames
    list.

.PARAMETER AppNames
    Names of the .lnk files (without extension, without path) to copy from the
    all-users Start Menu to the Public Desktop. Searched recursively, so
    subfolder-nested shortcuts (e.g. vendor program groups) are found too.

.PARAMETER Apply
    Actually copy the shortcuts. Without this switch, the script only reports which
    shortcuts it found and would copy.

.EXAMPLE
    .\New-DesktopShortcutsFromStartMenu.ps1 -AppNames "Excel","Word","Outlook"

.EXAMPLE
    .\New-DesktopShortcutsFromStartMenu.ps1 -AppNames "Excel","Word","Outlook","Citrix Workspace" -Apply
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string[]] $AppNames,

    [switch] $Apply
)

$startMenuRoot   = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
$publicDesktop   = "$env:Public\Desktop"

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   New-DesktopShortcutsFromStartMenu" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "  Mode : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

foreach ($name in $AppNames) {
    $shortcut = Get-ChildItem -Path $startMenuRoot -Filter "$name.lnk" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $shortcut) {
        Write-Host "  [WARN] Shortcut not found for '$name'." -ForegroundColor Yellow
        continue
    }

    if (-not $Apply) {
        Write-Host "  [PREVIEW] Would copy '$($shortcut.FullName)' -> Public Desktop" -ForegroundColor Yellow
        continue
    }
    if (-not $PSCmdlet.ShouldProcess($shortcut.FullName, "Copy to Public Desktop")) { continue }

    Copy-Item -Path $shortcut.FullName -Destination $publicDesktop -Force
    Write-Host "  [OK]   Copied '$name' to Public Desktop." -ForegroundColor Green
}

Write-Host ""
if (-not $Apply) { Write-Host "  Re-run with -Apply to copy the shortcuts found above." -ForegroundColor Yellow }
Write-Host ""
