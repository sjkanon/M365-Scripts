#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Registers the Intune iOS Compliance Updater as a scheduled task.

.DESCRIPTION
    Creates a Windows Task Scheduler task that runs the compliance updater
    every Monday at 07:00 AM.

.PARAMETER ScriptPath
    Full path to the Update-iOSCompliancePolicy.ps1 script

.PARAMETER RunAsUser
    The Windows user account to run the task as (default: SYSTEM)

.EXAMPLE
    .\Register-ScheduledTask.ps1
    .\Register-ScheduledTask.ps1 -ScriptPath "C:\Scripts\Update-iOSCompliancePolicy.ps1"
#>

param (
    [string]$ScriptPath = "$PSScriptRoot\Update-iOSCompliancePolicy.ps1",
    [string]$RunAsUser  = "SYSTEM"
)

$taskName        = "BraveHub - Intune iOS Compliance Updater"
$taskDescription = "Automatically updates the minimum iOS version in Intune compliance policy via Graph API."

Write-Host "Registering scheduled task: $taskName" -ForegroundColor Cyan

# Validate script exists
if (-not (Test-Path $ScriptPath)) {
    Write-Host "Script not found at: $ScriptPath" -ForegroundColor Red
    exit 1
}

# Task action
$action  = New-ScheduledTaskAction `
    -Execute  "powershell.exe" `
    -Argument "-ExecutionPolicy Bypass -NonInteractive -File `"$ScriptPath`""

# Trigger: every Monday at 07:00
$trigger = New-ScheduledTaskTrigger `
    -Weekly `
    -DaysOfWeek Monday `
    -At "07:00"

# Settings
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
    -RestartCount 2 `
    -RestartInterval (New-TimeSpan -Minutes 5) `
    -StartWhenAvailable

# Principal (run as SYSTEM)
$principal = New-ScheduledTaskPrincipal `
    -UserId    $RunAsUser `
    -RunLevel  Highest `
    -LogonType ServiceAccount

try {
    # Remove existing task if it exists
    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Host "Removed existing task." -ForegroundColor Yellow
    }

    Register-ScheduledTask `
        -TaskName   $taskName `
        -Action     $action `
        -Trigger    $trigger `
        -Settings   $settings `
        -Principal  $principal `
        -Description $taskDescription `
        -Force | Out-Null

    Write-Host "Scheduled task registered successfully!" -ForegroundColor Green
    Write-Host "Task: $taskName" -ForegroundColor Green
    Write-Host "Schedule: Every Monday at 07:00" -ForegroundColor Green
    Write-Host "Script: $ScriptPath" -ForegroundColor Green
}
catch {
    Write-Host "Failed to register task: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
