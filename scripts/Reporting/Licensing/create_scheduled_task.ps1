#Requires -Version 5.1
# ================================================
# Create-LicensingReportTask.ps1
# Registers a monthly scheduled task that runs the
# licensing report generator on the 6th of each month.
#
# Run once as Administrator on the target machine.
# ================================================

$ErrorActionPreference = "Stop"

# ── Configuration ────────────────────────────────────────────────────────────
$TaskName  = "Licensing Report Generator"
$TaskDesc  = "Generates the monthly licensing and Azure cost report from Pax8 and Ingram billing data."
$RunAsUser = "$env:USERDOMAIN\sa-halo"   # Change to the service account that should run the task
$RunDay    = 6                            # Day of month to run
$RunTime   = "08:00"
# ─────────────────────────────────────────────────────────────────────────────

# ── Self-elevation check ──────────────────────────────────────────────────────
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "[ERROR] This script must be run as Administrator." -ForegroundColor Red
    Write-Host "        Right-click PowerShell and choose 'Run as administrator'."
    Read-Host  "Press Enter to exit"
    exit 1
}

# ── Validate Python script path ──────────────────────────────────────────────
$ScriptPath = Join-Path $PSScriptRoot "genereer_licentie_overzicht.py"
if (-not (Test-Path $ScriptPath)) {
    Write-Host "[ERROR] Python script not found:" -ForegroundColor Red
    Write-Host "        $ScriptPath"
    Read-Host  "Press Enter to exit"
    exit 2
}

# ── Locate python.exe ─────────────────────────────────────────────────────────
$PythonExe = (Get-Command python.exe -ErrorAction SilentlyContinue)?.Source

if (-not $PythonExe) {
    # Fallback: check common install locations
    $candidates = @(
        "$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe",
        "C:\Python3*\python.exe",
        "$env:ProgramFiles\Python3*\python.exe"
    )
    foreach ($pattern in $candidates) {
        $found = Get-Item $pattern -ErrorAction SilentlyContinue | Select-Object -Last 1
        if ($found) { $PythonExe = $found.FullName; break }
    }
}

if (-not $PythonExe) {
    Write-Host "[ERROR] python.exe not found in PATH or common install locations." -ForegroundColor Red
    Write-Host "        Install Python from https://www.python.org or provide the full path manually."
    Read-Host  "Press Enter to exit"
    exit 3
}

Write-Host "Python found: $PythonExe" -ForegroundColor DarkGray

# ── Calculate next trigger date ───────────────────────────────────────────────
$now   = Get-Date
$next  = Get-Date -Year $now.Year -Month $now.Month -Day $RunDay -Hour 8 -Minute 0 -Second 0
if ($next -le $now) {
    $next = $next.AddMonths(1)
}
$StartBoundary = $next.ToString("yyyy-MM-ddTHH:mm:ss")

# ── Check if task already exists ─────────────────────────────────────────────
$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host ""
    Write-Host "Task '$TaskName' already exists." -ForegroundColor Yellow
    $answer = Read-Host "Update it with the current settings? [Y/N]"
    if ($answer -notmatch '^[Yy]') {
        Write-Host "Aborted. No changes made."
        exit 0
    }
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Existing task removed." -ForegroundColor DarkGray
}

# ── Build task components ─────────────────────────────────────────────────────
$Action = New-ScheduledTaskAction `
    -Execute  $PythonExe `
    -Argument "`"$ScriptPath`"" `
    -WorkingDirectory $PSScriptRoot

$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit   (New-TimeSpan -Hours 2) `
    -RestartCount         1 `
    -RestartInterval      (New-TimeSpan -Minutes 30) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable

# Register with a daily placeholder trigger first (monthly trigger requires XML)
$PlaceholderTrigger = New-ScheduledTaskTrigger -Daily -At $RunTime

try {
    Register-ScheduledTask `
        -TaskName    $TaskName `
        -Description $TaskDesc `
        -Trigger     $PlaceholderTrigger `
        -Action      $Action `
        -Settings    $Settings `
        -RunLevel    Highest `
        -User        $RunAsUser `
        -Force | Out-Null
} catch {
    Write-Host "[ERROR] Failed to register task: $_" -ForegroundColor Red
    Write-Host "        Verify that the service account '$RunAsUser' exists and has 'Log on as batch job' rights."
    Read-Host  "Press Enter to exit"
    exit 4
}

# ── Replace placeholder trigger with monthly XML trigger ──────────────────────
$TriggerXml = @"
<CalendarTrigger xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <StartBoundary>$StartBoundary</StartBoundary>
  <Enabled>true</Enabled>
  <ScheduleByMonth>
    <DaysOfMonth>
      <Day>$RunDay</Day>
    </DaysOfMonth>
    <Months>
      <January/><February/><March/><April/><May/><June/>
      <July/><August/><September/><October/><November/><December/>
    </Months>
  </ScheduleByMonth>
</CalendarTrigger>
"@

$Task        = Get-ScheduledTask -TaskName $TaskName
$TaskXml     = [xml]($Task | Export-ScheduledTask)
$triggersNode = $TaskXml.Task.Triggers

while ($triggersNode.HasChildNodes) {
    $triggersNode.RemoveChild($triggersNode.FirstChild) | Out-Null
}

$fragment          = $TaskXml.CreateDocumentFragment()
$fragment.InnerXml = $TriggerXml
$triggersNode.AppendChild($fragment) | Out-Null

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
Register-ScheduledTask -TaskName $TaskName -Xml $TaskXml.OuterXml -User $RunAsUser -Force | Out-Null

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "================================================" -ForegroundColor Green
Write-Host "  Scheduled task registered successfully" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Name       : $TaskName"
Write-Host "  Account    : $RunAsUser"
Write-Host "  Schedule   : Day $RunDay of every month at $RunTime"
Write-Host "  Next run   : $($next.ToString('dd-MM-yyyy HH:mm'))"
Write-Host "  Python     : $PythonExe"
Write-Host "  Script     : $ScriptPath"
Write-Host ""
Write-Host "To test immediately:" -ForegroundColor Yellow
Write-Host "  Start-ScheduledTask -TaskName '$TaskName'"
Write-Host ""
