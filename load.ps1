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

$configFile = Join-Path $PSScriptRoot 'load.config.ps1'

# ── First run: ask and save config ────────────────────────────────────────────
if (-not (Test-Path $configFile)) {
    Write-Host ""
    Write-Host "  First run — enter your admin details." -ForegroundColor Cyan
    Write-Host "  These will be saved in load.config.ps1 (gitignored)."
    Write-Host ""
    $upnInput  = Read-Host "  Admin UPN (e.g. admin@contoso.com)"
    $nameInput = Read-Host "  Display name (e.g. Sjoerd)"
    Write-Host ""

    @"
# M365-Scripts local config — do not commit
`$global:upn      = "$upnInput"
`$global:realname = "$nameInput"
"@ | Set-Content -Path $configFile -Encoding UTF8

    Write-Host "  Config saved. Delete load.config.ps1 to reset." -ForegroundColor DarkGray
    Write-Host ""
}

# ── Check and import required modules ────────────────────────────────────────
$requiredModules = @(
    'ExchangeOnlineManagement'
    'Microsoft.Graph.Authentication'
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
& "$PSScriptRoot\menu.ps1"
