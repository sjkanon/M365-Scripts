#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Disable monitor/standby sleep and fast startup — for kiosk, reception, or always-on devices.

.DESCRIPTION
    Sets the active power plan's monitor and standby timeouts to "never" (AC and DC)
    and disables Fast Startup (HiberbootEnabled), for devices that must never sleep or
    lock via power management — shared kiosks, reception/lobby PCs, digital signage.

.PARAMETER Apply
    Actually change power settings. Without this switch, the script only reports
    current values.

.EXAMPLE
    .\Set-KioskPowerSettings.ps1

.EXAMPLE
    .\Set-KioskPowerSettings.ps1 -Apply
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch] $Apply
)

Write-Host ""
Write-Host "  Set-KioskPowerSettings" -ForegroundColor Cyan
Write-Host "  Mode : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

$hiberbootPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power'
$currentHiberboot = (Get-ItemProperty -Path $hiberbootPath -Name HiberbootEnabled -ErrorAction SilentlyContinue).HiberbootEnabled

Write-Host "  Current HiberbootEnabled (Fast Startup) : $(if ($null -eq $currentHiberboot) { 'not set (default: on)' } else { $currentHiberboot })"

if (-not $Apply) {
    Write-Host ""
    Write-Host "  Would set monitor/standby timeouts (AC+DC) to 0 (never) via powercfg." -ForegroundColor Yellow
    Write-Host "  Would set HiberbootEnabled = 0 (disable Fast Startup)." -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to perform these changes." -ForegroundColor Yellow
    Write-Host ""
    exit 0
}

if (-not $PSCmdlet.ShouldProcess('Power settings', "Disable sleep/fast startup for kiosk use")) { exit 0 }

powercfg -change -monitor-timeout-ac 0
powercfg -change -monitor-timeout-dc 0
powercfg -change -standby-timeout-ac 0
powercfg -change -standby-timeout-dc 0
Write-Host "  [OK]   Monitor/standby timeouts set to never." -ForegroundColor Green

if (-not (Test-Path $hiberbootPath)) { New-Item -Path $hiberbootPath -Force | Out-Null }
New-ItemProperty -Path $hiberbootPath -Name HiberbootEnabled -Value 0 -PropertyType DWORD -Force | Out-Null
Write-Host "  [OK]   Fast Startup disabled." -ForegroundColor Green

Write-Host ""
