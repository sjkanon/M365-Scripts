#Requires -Version 5.1
<#
.SYNOPSIS
    Startup loader for M365-Scripts.
    Asks for your admin UPN and display name on first run, saves them locally,
    then launches the interactive menu.

.DESCRIPTION
    Configuration is stored in load.config.ps1 (gitignored).
    Delete that file to re-enter your credentials.
#>

[CmdletBinding()]
param(
    [switch]$SetupStartup,
    [switch]$RemoveStartup
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
$requiredModules = @(
    'ExchangeOnlineManagement'
    'Microsoft.Graph.Authentication'
    'Microsoft.Graph.Sites'
    'Microsoft.Graph.Identity.DirectoryManagement'
    'Microsoft.Graph.Users'
    'Microsoft.Graph.Groups'
    'Microsoft.Graph.Applications'
)

$missing = $requiredModules | Where-Object { -not (Get-Module -ListAvailable -Name $_ ) }

if ($missing) {
    Write-Host ""
    Write-Host "  Missing modules: $($missing -join ', ')" -ForegroundColor Yellow
    $install = Read-Host "  Run Install-Modules.ps1 now? [Y/n]"
    if ($install -notmatch '^[Nn]') {
        & "$PSScriptRoot\scripts\Startup\Install-Modules.ps1"
    } else {
        Write-Host "  M365 functions may not work until modules are installed." -ForegroundColor DarkYellow
    }
} else {
    Write-Host ""
    Write-Host "  Loading modules..." -ForegroundColor DarkGray
    $requiredModules | ForEach-Object { Import-Module $_ -ErrorAction SilentlyContinue }
}

# ── Load config and launch menu ───────────────────────────────────────────────
. $configFile

if (-not $global:authMode) { $global:authMode = 'Direct' }
if ($null -eq $global:useDeviceCodeAuth) { $global:useDeviceCodeAuth = $true }

& "$PSScriptRoot\menu.ps1"
