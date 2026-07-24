#Requires -Version 5.1
<#
.SYNOPSIS
    Create a "Lock Workstation" shortcut on the desktop (or Start Menu).

.DESCRIPTION
    Generalized replacement for a script that downloaded a custom icon and batch
    file from an MSP's own internal endpoint and pinned it to the taskbar via an
    undocumented registry hack. This version needs no external download and no
    registry trickery: it creates a plain .lnk shortcut that runs the standard
    Windows lock-workstation command (rundll32.exe user32.dll,LockWorkStation),
    optionally with a custom .ico if you already have one on disk. Users can pin
    it to the taskbar themselves via the normal right-click "Pin to taskbar".

.PARAMETER TargetFolder
    Where to create the shortcut. Default: the Public Desktop (visible to every
    user on the machine). Pass the all-users Start Menu programs folder to place
    it there instead.

.PARAMETER ShortcutName
    Shortcut file name, without extension. Default: "Lock Workstation".

.PARAMETER IconPath
    Optional path to a local .ico file to use for the shortcut. If omitted, the
    shortcut uses rundll32.exe's own default icon.

.PARAMETER Apply
    Actually create the shortcut. Without this switch, the script only reports
    what it would do.

.EXAMPLE
    .\New-LockWorkstationShortcut.ps1

.EXAMPLE
    .\New-LockWorkstationShortcut.ps1 -Apply

.EXAMPLE
    .\New-LockWorkstationShortcut.ps1 -IconPath "C:\Assets\lock.ico" -Apply
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $TargetFolder = "$env:Public\Desktop",
    [string] $ShortcutName = 'Lock Workstation',
    [string] $IconPath,
    [switch] $Apply
)

$shortcutFile = Join-Path $TargetFolder "$ShortcutName.lnk"

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   New-LockWorkstationShortcut" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Shortcut : $shortcutFile"
Write-Host ("  Mode     : {0}" -f $(if ($Apply) { 'Apply' } else { 'Preview only' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

if (-not $Apply) {
    Write-Host "  Would create a shortcut running: rundll32.exe user32.dll,LockWorkStation" -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to create it." -ForegroundColor Yellow
    exit 0
}

if (-not $PSCmdlet.ShouldProcess($shortcutFile, "Create Lock Workstation shortcut")) { exit 0 }

try {
    if (-not (Test-Path $TargetFolder)) { New-Item -ItemType Directory -Path $TargetFolder -Force | Out-Null }

    $wshShell = New-Object -ComObject WScript.Shell
    $shortcut = $wshShell.CreateShortcut($shortcutFile)
    $shortcut.TargetPath = "$env:SystemRoot\System32\rundll32.exe"
    $shortcut.Arguments  = 'user32.dll,LockWorkStation'
    if ($IconPath -and (Test-Path $IconPath)) { $shortcut.IconLocation = $IconPath }
    $shortcut.Description = 'Lock this computer'
    $shortcut.Save()

    Write-Host "  [OK]   Created: $shortcutFile" -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
