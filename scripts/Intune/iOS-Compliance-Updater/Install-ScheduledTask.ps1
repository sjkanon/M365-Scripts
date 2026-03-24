#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Registers the Intune iOS Compliance Updater as a Windows scheduled task.

.DESCRIPTION
    Creates a Task Scheduler task that runs Update-iOSCompliancePolicy.ps1
    every Monday at 07:00 as SYSTEM.

.PARAMETER ScriptPath
    Full path to Update-iOSCompliancePolicy.ps1 (default: same folder as this script).

.PARAMETER RunAsUser
    Windows account to run the task as (default: SYSTEM).

.PARAMETER TaskName
    Name of the scheduled task (default: "Intune iOS Compliance Updater").

.EXAMPLE
    .\Install-ScheduledTask.ps1

.EXAMPLE
    .\Install-ScheduledTask.ps1 -ScriptPath "C:\Scripts\Update-iOSCompliancePolicy.ps1"
#>
param (
    [string] $ScriptPath = "$PSScriptRoot\Update-iOSCompliancePolicy.ps1",
    [string] $RunAsUser  = 'SYSTEM',
    [string] $TaskName   = 'Intune iOS Compliance Updater'
)

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Install-ScheduledTask" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $ScriptPath)) {
    Write-Host "  [ERROR] Script not found: $ScriptPath" -ForegroundColor Red
    exit 1
}

$action    = New-ScheduledTaskAction `
    -Execute  'powershell.exe' `
    -Argument "-ExecutionPolicy Bypass -NonInteractive -File `"$ScriptPath`""

$trigger   = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At '07:00'

$settings  = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
    -RestartCount 2 `
    -RestartInterval (New-TimeSpan -Minutes 5) `
    -StartWhenAvailable

$principal = New-ScheduledTaskPrincipal `
    -UserId    $RunAsUser `
    -RunLevel  Highest `
    -LogonType ServiceAccount

try {
    $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($existing) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "  [INFO] Existing task removed." -ForegroundColor DarkGray
    }

    Register-ScheduledTask `
        -TaskName    $TaskName `
        -Action      $action `
        -Trigger     $trigger `
        -Settings    $settings `
        -Principal   $principal `
        -Description 'Automatically updates the minimum iOS version in Intune compliance policy via Graph API.' `
        -Force | Out-Null

    Write-Host "  [OK]   Scheduled task registered." -ForegroundColor Green
    Write-Host "  [OK]   Name     : $TaskName" -ForegroundColor Green
    Write-Host "  [OK]   Schedule : Every Monday at 07:00" -ForegroundColor Green
    Write-Host "  [OK]   Script   : $ScriptPath" -ForegroundColor Green
    Write-Host "  [OK]   Run as   : $RunAsUser" -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] Failed to register task: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
