#Requires -Version 5.1
<#
.SYNOPSIS
    Startup loader for M365-Scripts.
    Asks for your admin UPN and display name on first run, saves them locally,
    then launches the interactive menu.

.DESCRIPTION
    Configuration is stored in load.config.ps1 (gitignored).
    Delete that file to re-enter your credentials.

    Before the menu opens, every module in scripts\Startup\RequiredModules.psd1 is checked
    with Update-Modules.ps1 — missing, older than its minimum, or behind the PowerShell
    Gallery (asked at most once every 24 hours) — and it offers to install or update them.

.PARAMETER SetupStartup
    Create a shortcut that starts this launcher at Windows sign-in.

.PARAMETER RemoveStartup
    Remove that shortcut again.

.PARAMETER SkipModuleCheck
    Skip the module check for this start (offline, or in a hurry).
#>

[CmdletBinding()]
param(
    [switch]$SetupStartup,
    [switch]$RemoveStartup,
    [switch]$SkipModuleCheck
)

function Set-StartupLauncher {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Enable
    )

    if (-not $IsWindows) {
        Write-Warning 'Startup shortcut setup is only supported on Windows.'
        return
    }

    $startupDir = [Environment]::GetFolderPath('Startup')
    $shortcutPath = Join-Path $startupDir 'M365-Scripts Launcher.lnk'

    if (-not $Enable) {
        if (Test-Path $shortcutPath) {
            Remove-Item -Path $shortcutPath -Force
            Write-Host "  Removed startup shortcut: $shortcutPath" -ForegroundColor Green
        } else {
            Write-Host '  Startup shortcut was not present.' -ForegroundColor DarkGray
        }
        return
    }

    $pwshPath = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
    if (-not $pwshPath) {
        throw 'pwsh.exe not found on PATH. Install PowerShell 7 to enable startup launcher.'
    }

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $pwshPath
    $shortcut.Arguments = "-NoLogo -NoProfile -ExecutionPolicy Bypass -File `"$PSScriptRoot\load.ps1`""
    $shortcut.WorkingDirectory = $PSScriptRoot
    $shortcut.Description = 'Start M365-Scripts launcher at sign-in'
    $shortcut.Save()

    Write-Host "  Startup shortcut created: $shortcutPath" -ForegroundColor Green
}

if ($SetupStartup) {
    Set-StartupLauncher -Enable $true
    if ($RemoveStartup) { Write-Warning 'Both -SetupStartup and -RemoveStartup were provided. Startup is now enabled.' }
}

if ($RemoveStartup -and -not $SetupStartup) {
    Set-StartupLauncher -Enable $false
}

$configFile = Join-Path $PSScriptRoot 'load.config.ps1'

# ── First run: ask and save config ────────────────────────────────────────────
if (-not (Test-Path $configFile)) {
    Write-Host ""
    Write-Host "  First run — enter your admin details." -ForegroundColor Cyan
    Write-Host "  These will be saved in load.config.ps1 (gitignored)."
    Write-Host ""
    $upnInput  = Read-Host "  Admin UPN (e.g. admin@contoso.com)"
    $nameInput = Read-Host "  Display name (e.g. Sjoerd)"
    $gdapInput = Read-Host "  Use delegated GDAP login by default? [Y/n]"
    $authModeInput = if ($gdapInput -match '^[Nn]') { 'Direct' } else { 'GDAP' }

    $defaultDomainInput = ''
    if ($authModeInput -eq 'GDAP') {
        $defaultDomainInput = Read-Host "  Default customer domain for GDAP (optional, e.g. contoso.com)"
    }

    $deviceCodeInput = Read-Host "  Use device code sign-in for Graph? [Y/n]"
    $useDeviceCode = -not ($deviceCodeInput -match '^[Nn]')
    Write-Host ""

    @"
# M365-Scripts local config — do not commit
`$global:upn      = "$upnInput"
`$global:realname = "$nameInput"
`$global:authMode = "$authModeInput"
`$global:defaultCustomerDomain = "$defaultDomainInput"
`$global:useDeviceCodeAuth = `$$useDeviceCode
"@ | Set-Content -Path $configFile -Encoding UTF8

    Write-Host "  Config saved. Delete load.config.ps1 to reset." -ForegroundColor DarkGray
    Write-Host ""
}

# ── Check and import required modules ────────────────────────────────────────
# The list lives in scripts\Startup\RequiredModules.psd1; Update-Modules.ps1 checks each
# module for missing / older than its minimum / behind the PowerShell Gallery and asks
# before changing anything. The gallery is asked at most once a day (cached), so a normal
# start stays quick. The same line works in a PowerShell profile.
$moduleScript = Join-Path $PSScriptRoot 'scripts\Startup\Update-Modules.ps1'
$moduleList   = Join-Path $PSScriptRoot 'scripts\Startup\RequiredModules.psd1'

if (-not $SkipModuleCheck) {
    Write-Host ""
    & $moduleScript -RequiredOnly -Prompt -MaxAgeHours 24
}

Write-Host "  Loading modules..." -ForegroundColor DarkGray
(Import-PowerShellDataFile -Path $moduleList).Modules |
    Where-Object { $_.ImportAtStartup } |
    ForEach-Object { Import-Module $_.Name -ErrorAction SilentlyContinue }

# ── Load config and launch menu ───────────────────────────────────────────────
. $configFile

if (-not $global:authMode) { $global:authMode = 'Direct' }
if ($null -eq $global:useDeviceCodeAuth) { $global:useDeviceCodeAuth = $true }

& "$PSScriptRoot\menu.ps1"
