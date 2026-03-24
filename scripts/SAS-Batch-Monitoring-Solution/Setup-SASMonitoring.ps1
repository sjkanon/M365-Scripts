<#
.SYNOPSIS
    Setup script for SAS batch error monitoring

.DESCRIPTION
    Installs and configures the SAS batch error monitoring solution:
    - Copies monitoring script to C:\Scripts\
    - Creates scheduled task for daily reports
    - Configures Zabbix integration (optional)
    - Sets up email alerts (optional)

.PARAMETER InstallZabbix
    Install Zabbix UserParameter configuration

.PARAMETER EmailAlerts
    Configure email alerts

.PARAMETER SmtpServer
    SMTP server for email alerts

.PARAMETER EmailTo
    Email recipient for alerts

.PARAMETER EmailFrom
    Email sender address

.EXAMPLE
    .\Setup-SASMonitoring.ps1 -InstallZabbix

.EXAMPLE
    .\Setup-SASMonitoring.ps1 -EmailAlerts -SmtpServer "smtp.bravehub.be" -EmailTo "sjoerd@bravehub.be" -EmailFrom "sas-monitor@bravehub.be"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [switch]$InstallZabbix,
    
    [Parameter(Mandatory=$false)]
    [switch]$EmailAlerts,
    
    [Parameter(Mandatory=$false)]
    [string]$SmtpServer = "smtp.bravehub.be",
    
    [Parameter(Mandatory=$false)]
    [string]$EmailTo = "sjoerd@bravehub.be",
    
    [Parameter(Mandatory=$false)]
    [string]$EmailFrom = "sas-monitor@bravehub.be",
    
    [Parameter(Mandatory=$false)]
    [string]$SASLogPath = "E:\SAS\Logs"
)

# Require administrator privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script must be run as Administrator"
    exit 1
}

Write-Host "=== SAS Batch Error Monitoring Setup ===" -ForegroundColor Cyan
Write-Host ""

# Create directories
$scriptDir = "C:\Scripts"
$logDir = "C:\Scripts\Logs"

if (-not (Test-Path $scriptDir)) {
    Write-Host "Creating directory: $scriptDir" -ForegroundColor Yellow
    New-Item -Path $scriptDir -ItemType Directory -Force | Out-Null
}

if (-not (Test-Path $logDir)) {
    Write-Host "Creating directory: $logDir" -ForegroundColor Yellow
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
}

# Copy monitoring script
Write-Host "Installing monitoring script..." -ForegroundColor Yellow
$monitorScript = Join-Path $scriptDir "Monitor-SASBatchErrors.ps1"

# Check if Monitor-SASBatchErrors.ps1 exists in current directory
if (Test-Path ".\Monitor-SASBatchErrors.ps1") {
    Copy-Item ".\Monitor-SASBatchErrors.ps1" -Destination $monitorScript -Force
    Write-Host "✓ Monitoring script installed" -ForegroundColor Green
} else {
    Write-Error "Monitor-SASBatchErrors.ps1 not found in current directory"
    exit 1
}

# Create scheduled task for daily reports
Write-Host "`nCreating scheduled task for daily monitoring..." -ForegroundColor Yellow

$taskName = "SAS Batch Error Monitoring"
$taskDescription = "Daily monitoring of SAS batch job errors"

# Check if task already exists
$existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

if ($existingTask) {
    Write-Host "Scheduled task already exists. Removing..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

# Create task action
$outputFile = Join-Path $logDir "SAS_Error_Report_$(Get-Date -Format 'yyyyMMdd').txt"
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$monitorScript`" -LogDirectory `"$SASLogPath`" -DaysToCheck 1 -OutputFormat Text -OutputFile `"$outputFile`""

# Create task trigger (daily at 08:00)
$trigger = New-ScheduledTaskTrigger -Daily -At "08:00"

# Create task settings
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

# Register task
Register-ScheduledTask -TaskName $taskName -Description $taskDescription -Action $action -Trigger $trigger -Settings $settings -User "SYSTEM" -RunLevel Highest | Out-Null

Write-Host "✓ Scheduled task created (runs daily at 08:00)" -ForegroundColor Green

# Install Zabbix configuration
if ($InstallZabbix) {
    Write-Host "`nInstalling Zabbix configuration..." -ForegroundColor Yellow
    
    $zabbixConfDir = "C:\Program Files\Zabbix Agent 2\zabbix_agent2.d"
    
    if (-not (Test-Path $zabbixConfDir)) {
        Write-Warning "Zabbix Agent directory not found: $zabbixConfDir"
        Write-Warning "Please install Zabbix Agent first or specify correct path"
    } else {
        $zabbixConf = Join-Path $zabbixConfDir "sas_monitor.conf"
        
        # Create Zabbix configuration
        $confContent = @"
# Zabbix UserParameter configuration for SAS batch error monitoring

# Check for critical SAS batch errors (last 7 days)
UserParameter=sas.batch.errors.critical,powershell -NoProfile -ExecutionPolicy Bypass -File "$monitorScript" -LogDirectory "$SASLogPath" -DaysToCheck 7 -OutputFormat Zabbix

# Check for critical errors in last 24 hours
UserParameter=sas.batch.errors.critical.24h,powershell -NoProfile -ExecutionPolicy Bypass -File "$monitorScript" -LogDirectory "$SASLogPath" -DaysToCheck 1 -OutputFormat Zabbix

# Get full JSON report
UserParameter=sas.batch.errors.json,powershell -NoProfile -ExecutionPolicy Bypass -File "$monitorScript" -LogDirectory "$SASLogPath" -DaysToCheck 7 -OutputFormat JSON
"@
        
        $confContent | Out-File -FilePath $zabbixConf -Encoding ASCII
        Write-Host "✓ Zabbix configuration installed" -ForegroundColor Green
        Write-Host "  Please restart Zabbix Agent service" -ForegroundColor Yellow
    }
}

# Configure email alerts
if ($EmailAlerts) {
    Write-Host "`nConfiguring email alerts..." -ForegroundColor Yellow
    
    $emailScriptPath = Join-Path $scriptDir "Send-SASErrorAlert.ps1"
    
    $emailScript = @"
`$ErrorReport = & "$monitorScript" -LogDirectory "$SASLogPath" -DaysToCheck 1 -OutputFormat Text

if (`$LASTEXITCODE -eq 2) {
    # Critical errors found
    `$subject = "CRITICAL: SAS Batch Errors Detected"
    `$body = `$ErrorReport
    
    Send-MailMessage -SmtpServer "$SmtpServer" ``
        -From "$EmailFrom" ``
        -To "$EmailTo" ``
        -Subject `$subject ``
        -Body `$body ``
        -Priority High
}
"@
    
    $emailScript | Out-File -FilePath $emailScriptPath -Encoding UTF8
    
    # Create scheduled task for email alerts (runs hourly)
    $emailTaskName = "SAS Batch Error Email Alerts"
    $emailAction = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$emailScriptPath`""
    $emailTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 1)
    
    Register-ScheduledTask -TaskName $emailTaskName -Description "Hourly SAS error email alerts" -Action $emailAction -Trigger $emailTrigger -Settings $settings -User "SYSTEM" -RunLevel Highest | Out-Null
    
    Write-Host "✓ Email alerts configured" -ForegroundColor Green
    Write-Host "  Alerts will be sent to: $EmailTo" -ForegroundColor Yellow
}

Write-Host "`n=== Setup Complete ===" -ForegroundColor Green
Write-Host "`nInstalled components:" -ForegroundColor Cyan
Write-Host "  - Monitoring script: $monitorScript"
Write-Host "  - Log directory: $logDir"
Write-Host "  - Scheduled task: $taskName (daily at 08:00)"

if ($InstallZabbix) {
    Write-Host "  - Zabbix monitoring: Configured"
}

if ($EmailAlerts) {
    Write-Host "  - Email alerts: Configured ($EmailTo)"
}

Write-Host "`nTo manually run the monitor:" -ForegroundColor Cyan
Write-Host "  powershell -File `"$monitorScript`" -LogDirectory `"$SASLogPath`" -DaysToCheck 7"

Write-Host "`nTo test Zabbix integration:" -ForegroundColor Cyan
Write-Host "  zabbix_get -s <hostname> -k sas.batch.errors.critical"
