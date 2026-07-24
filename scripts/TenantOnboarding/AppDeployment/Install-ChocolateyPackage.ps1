#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Install, upgrade, or uninstall a package via Chocolatey.

.DESCRIPTION
    Generic wrapper replacing a pair of hardcoded "install git" / "uninstall git"
    scripts. Installs Chocolatey itself if it isn't present, then installs the named
    package (upgrading it if already installed) or uninstalls it.

.PARAMETER PackageName
    Chocolatey package ID, e.g. "git", "7zip", "vscode".

.PARAMETER Uninstall
    Uninstall the package instead of installing/upgrading it.

.PARAMETER Apply
    Actually install/upgrade/uninstall. Without this switch, the script only reports
    what it would do.

.EXAMPLE
    .\Install-ChocolateyPackage.ps1 -PackageName git -Apply

.EXAMPLE
    .\Install-ChocolateyPackage.ps1 -PackageName git -Uninstall -Apply
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $PackageName,

    [switch] $Uninstall,
    [switch] $Apply
)

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Install-ChocolateyPackage : $PackageName" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "  Action : $(if ($Uninstall) { 'Uninstall' } else { 'Install/Upgrade' })"
Write-Host "  Mode   : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

$chocoInstalled = Get-Command choco.exe -ErrorAction SilentlyContinue

if (-not $Apply) {
    if (-not $chocoInstalled) { Write-Host "  Would install Chocolatey first (not currently present)." -ForegroundColor Yellow }
    Write-Host "  Would run: choco $(if ($Uninstall) { 'uninstall' } else { 'install/upgrade' }) $PackageName -y" -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to perform this action." -ForegroundColor Yellow
    Write-Host ""
    return
}

if (-not $PSCmdlet.ShouldProcess($PackageName, "$(if ($Uninstall) { 'Uninstall' } else { 'Install/upgrade' }) via Chocolatey")) { return }

if (-not $chocoInstalled -and -not $Uninstall) {
    Write-Host "  Installing Chocolatey..." -ForegroundColor DarkGray
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
}

if ($Uninstall) {
    choco uninstall $PackageName -y
} else {
    $localPackages = choco list --localonly
    if ($localPackages -match [regex]::Escape($PackageName)) {
        choco upgrade $PackageName -y
    } else {
        choco install $PackageName -y
    }
}

Write-Host ""
Write-Host "  [OK]   Done." -ForegroundColor Green
Write-Host ""
