#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Register a scheduled task that keeps OneDrive running, restarting it hourly if needed.

.DESCRIPTION
    Consolidates several near-duplicate old scripts (a simple "relaunch every 59
    minutes" task, a more thorough "kill OneDrive/FileCoAuth then relaunch" variant,
    and a "drop a manual reset .bat on the desktop" helper) into one parameterized
    watchdog. Registers a scheduled task, running as the currently logged-on user,
    that on each trigger:
      - relaunches OneDrive.exe /background if it isn't running, or
      - (with -ForceRestartOnEachRun) stops OneDrive/FileCoAuth first, then relaunches.

    -Reset instead performs a one-time `onedrive.exe /reset` (clears local sync state
    — use when a profile's sync is stuck) rather than registering the watchdog task.

.PARAMETER IntervalMinutes
    How often the watchdog re-checks OneDrive. Default: 59.

.PARAMETER ForceRestartOnEachRun
    Stop OneDrive.exe/FileCoAuth.exe on every trigger and relaunch, instead of only
    relaunching when the process isn't running at all.

.PARAMETER Reset
    Run `onedrive.exe /reset` once, immediately, instead of registering the watchdog
    task. Use for a stuck sync profile — OneDrive will re-download after relaunch.

.PARAMETER Apply
    Actually register the task (or perform the reset). Without this switch, the
    script only reports what it would do.

.EXAMPLE
    # Preview the watchdog task
    .\Register-OneDriveWatchdog.ps1

.EXAMPLE
    .\Register-OneDriveWatchdog.ps1 -Apply

.EXAMPLE
    .\Register-OneDriveWatchdog.ps1 -ForceRestartOnEachRun -IntervalMinutes 30 -Apply

.EXAMPLE
    # One-time reset of a stuck sync profile
    .\Register-OneDriveWatchdog.ps1 -Reset -Apply
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [int] $IntervalMinutes = 59,
    [switch] $ForceRestartOnEachRun,
    [switch] $Reset,
    [switch] $Apply
)

$oneDriveExe = "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe"
if (-not (Test-Path $oneDriveExe)) { $oneDriveExe = 'C:\Program Files\Microsoft OneDrive\OneDrive.exe' }

Write-Host ""
Write-Host "  Register-OneDriveWatchdog" -ForegroundColor Cyan
Write-Host "  Mode : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

if ($Reset) {
    if (-not $Apply) {
        Write-Host "  Would run: `"$oneDriveExe`" /reset" -ForegroundColor Yellow
        Write-Host "  Re-run with -Apply to perform the reset." -ForegroundColor Yellow
        Write-Host ""
        exit 0
    }
    if (-not $PSCmdlet.ShouldProcess('OneDrive', 'Reset local sync state (/reset)')) { exit 0 }

    Start-Process -FilePath $oneDriveExe -ArgumentList '/reset'
    Start-Sleep -Seconds 5
    Start-Process -FilePath $oneDriveExe -ArgumentList '/background'
    Write-Host "  [OK]   OneDrive reset triggered and relaunched." -ForegroundColor Green
    Write-Host ""
    exit 0
}

if (-not $Apply) {
    Write-Host "  Would register scheduled task 'OneDrive Watchdog', trigger every $IntervalMinutes minute(s)," -ForegroundColor Yellow
    Write-Host "  running as the current user, $(if ($ForceRestartOnEachRun) { 'force-restarting OneDrive on every run' } else { 'relaunching OneDrive only if not running' })." -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to register this task." -ForegroundColor Yellow
    Write-Host ""
    exit 0
}

if (-not $PSCmdlet.ShouldProcess('OneDrive Watchdog', 'Register scheduled task')) { exit 0 }

$scriptFolder = Join-Path $env:ProgramData 'TenantOnboarding\OneDriveWatchdog'
if (-not (Test-Path $scriptFolder)) { New-Item -Path $scriptFolder -ItemType Directory -Force | Out-Null }
$payloadPath = Join-Path $scriptFolder 'Invoke-OneDriveWatchdog.ps1'

$payload = if ($ForceRestartOnEachRun) {
@"
`$exe = '$oneDriveExe'
Get-Process -Name OneDrive, FileCoAuth -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
Start-Process -FilePath `$exe -ArgumentList '/background'
"@
} else {
@"
`$exe = '$oneDriveExe'
if (-not (Get-Process -Name OneDrive -ErrorAction SilentlyContinue)) {
    Start-Process -FilePath `$exe -ArgumentList '/background'
}
"@
}

Set-Content -Path $payloadPath -Value $payload -Encoding Unicode -Force

$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionDuration ([TimeSpan]::MaxValue) -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes)
$action  = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-WindowStyle Hidden -NonInteractive -ExecutionPolicy Bypass -File `"$payloadPath`""
Register-ScheduledTask -TaskName 'OneDrive Watchdog' -Trigger $trigger -Action $action -User $env:USERNAME -Force | Out-Null

Write-Host "  [OK]   Scheduled task 'OneDrive Watchdog' registered (every $IntervalMinutes minute(s))." -ForegroundColor Green
Write-Host ""
