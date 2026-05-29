#Requires -Version 7.0
<#
.SYNOPSIS
    Bootstrap script to install and import all required modules for M365-Scripts.

.DESCRIPTION
    Installs and imports all PowerShell modules used across the M365-Scripts repository.
    Runs cross-platform (Windows, macOS, Linux). Windows-only modules (AzureAD,
    WindowsAutopilotIntune) are skipped automatically on non-Windows platforms.

.PARAMETER Force
    Reinstall modules even if already present.

.PARAMETER Scope
    Installation scope: CurrentUser (default, no elevation needed) or AllUsers (requires elevation).

.EXAMPLE
    .\Install-Modules.ps1

.EXAMPLE
    .\Install-Modules.ps1 -Force -Scope AllUsers

.NOTES
    Author: Sjoerd Kanon
#>

[CmdletBinding()]
param (
    [switch]$Force,

    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]$Scope = 'CurrentUser'
)

#region Helpers

function Write-Step {
    param([string]$Message)
    Write-Host "`n==> $Message" -ForegroundColor Cyan
}

function Write-Ok {
    param([string]$Message)
    Write-Host "    [OK]  $Message" -ForegroundColor Green
}

function Write-Skip {
    param([string]$Message)
    Write-Host "    [--]  $Message" -ForegroundColor DarkGray
}

function Write-Fail {
    param([string]$Message)
    Write-Host "    [!!]  $Message" -ForegroundColor Red
}

function Install-RequiredModule {
    param(
        [string]$Name,
        [string]$MinimumVersion = $null,
        [bool]$WindowsOnly = $false
    )

    if ($WindowsOnly -and -not $IsWindows) {
        Write-Skip "$Name (Windows only — skipped on $($PSVersionTable.OS.Split(' ')[0]))"
        return
    }

    $installParams = @{
        Name               = $Name
        Scope              = $Scope
        Force              = $Force.IsPresent
        AllowClobber       = $true
        SkipPublisherCheck = $true
        ErrorAction        = 'Stop'
    }
    if ($MinimumVersion) {
        $installParams['MinimumVersion'] = $MinimumVersion
    }

    $existing = Get-Module -ListAvailable -Name $Name | Select-Object -First 1
    if ($existing -and -not $Force) {
        Write-Ok "$Name $($existing.Version) already installed"
        return
    }

    try {
        Write-Host "    [ ]  Installing $Name..." -NoNewline -ForegroundColor Yellow
        Install-Module @installParams
        $installed = Get-Module -ListAvailable -Name $Name | Select-Object -First 1
        Write-Host "`r    [OK] Installed  $Name $($installed.Version)  " -ForegroundColor Green
    }
    catch {
        Write-Host ""
        Write-Fail "Failed to install $Name`: $_"
    }
}

function Import-RequiredModule {
    param(
        [string]$Name,
        [bool]$WindowsOnly = $false
    )

    if ($WindowsOnly -and -not $IsWindows) {
        return
    }

    try {
        Import-Module -Name $Name -ErrorAction Stop
        $loaded = Get-Module -Name $Name | Select-Object -First 1
        Write-Ok "Imported $Name $($loaded.Version)"
    }
    catch {
        Write-Fail "Could not import $Name`: $_"
    }
}

#endregion

#region Platform info

Write-Step "Platform"
Write-Host "    PowerShell $($PSVersionTable.PSVersion)  |  $($PSVersionTable.OS)"
Write-Host "    Install scope: $Scope"

#endregion

#region Ensure PSGallery is trusted

Write-Step "Verifying PSGallery"
$gallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
if ($gallery.InstallationPolicy -ne 'Trusted') {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Write-Ok "PSGallery set to Trusted"
} else {
    Write-Ok "PSGallery already trusted"
}

#endregion

#region Module definitions
# Format: Name, MinimumVersion (optional), WindowsOnly
$modules = @(
    @{ Name = 'ExchangeOnlineManagement';   MinimumVersion = '3.0.0'; WindowsOnly = $false }
    @{ Name = 'Microsoft.Graph.Authentication'; MinimumVersion = '2.0.0'; WindowsOnly = $false }
    @{ Name = 'Microsoft.Graph.Identity.SignIns'; MinimumVersion = '2.0.0'; WindowsOnly = $false }
    @{ Name = 'Microsoft.Graph.Identity.Governance'; MinimumVersion = '2.0.0'; WindowsOnly = $false }
    @{ Name = 'Microsoft.Graph.Applications';   MinimumVersion = '2.0.0'; WindowsOnly = $false }
    @{ Name = 'Microsoft.Graph.Calendar';       MinimumVersion = '2.0.0'; WindowsOnly = $false }
    @{ Name = 'Microsoft.Graph.Groups';         MinimumVersion = '2.0.0'; WindowsOnly = $false }
    @{ Name = 'Microsoft.Graph.Users';          MinimumVersion = '2.0.0'; WindowsOnly = $false }
    @{ Name = 'WindowsAutopilotIntune';         MinimumVersion = $null;   WindowsOnly = $true  }
    @{ Name = 'AzureAD';                        MinimumVersion = $null;   WindowsOnly = $true  }
)
#endregion

#region Install

Write-Step "Installing modules"
foreach ($mod in $modules) {
    Install-RequiredModule -Name $mod.Name -MinimumVersion $mod.MinimumVersion -WindowsOnly $mod.WindowsOnly
}

#endregion

#region Import

Write-Step "Importing modules"
foreach ($mod in $modules) {
    Import-RequiredModule -Name $mod.Name -WindowsOnly $mod.WindowsOnly
}

#endregion

Write-Host "`nDone. All available modules are installed and loaded.`n" -ForegroundColor Cyan
