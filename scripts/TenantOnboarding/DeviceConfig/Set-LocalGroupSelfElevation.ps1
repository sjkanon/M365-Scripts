#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Grant or revoke a logged-on user's membership of a local group at next logon, via a scheduled task.

.DESCRIPTION
    Consolidates four near-identical old scripts (grant/revoke local Administrators,
    grant/revoke local "Network Configuration Operators") into one parameterized
    script. Registers a SYSTEM-run scheduled task, triggered at logon, whose payload
    adds (or removes) whichever user is logged on at that moment to/from the named
    local group — the standard pattern for temporarily self-service-elevating a user
    without a permanent group assignment (e.g. via an Intune Win32 app / Proactive
    Remediation the user runs on demand). Removes any opposing task left behind by a
    previous run (e.g. a pending "Revoke" task when a new "Grant" is registered).

.PARAMETER GroupName
    Local group to grant/revoke membership of, e.g. "Administrators" or
    "Network Configuration Operators".

.PARAMETER Action
    "Grant" registers a task that adds the logged-on user to the group at next logon.
    "Revoke" registers a task that removes the logged-on user from the group at next
    logon (and un-registers any pending opposing task).

.PARAMETER TaskName
    Name for the scheduled task. Default: "<Action>-<GroupName without spaces>".

.PARAMETER Apply
    Actually register the scheduled task (and run it once immediately). Without this
    switch, the script only reports what it would do.

.EXAMPLE
    # Preview
    .\Set-LocalGroupSelfElevation.ps1 -GroupName "Administrators" -Action Grant

.EXAMPLE
    .\Set-LocalGroupSelfElevation.ps1 -GroupName "Administrators" -Action Grant -Apply

.EXAMPLE
    .\Set-LocalGroupSelfElevation.ps1 -GroupName "Administrators" -Action Revoke -Apply

.EXAMPLE
    .\Set-LocalGroupSelfElevation.ps1 -GroupName "Network Configuration Operators" -Action Grant -Apply
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $GroupName,

    [Parameter(Mandatory)]
    [ValidateSet('Grant', 'Revoke')]
    [string] $Action,

    [string] $TaskName,
    [switch] $Apply
)

$safeGroupName = $GroupName -replace '\s', ''
if (-not $TaskName) { $TaskName = "$Action-$safeGroupName" }
$opposingAction = if ($Action -eq 'Grant') { 'Revoke' } else { 'Grant' }
$opposingTaskName = "$opposingAction-$safeGroupName"

$scriptFolder = Join-Path $env:ProgramData 'TenantOnboarding\LocalGroupSelfElevation'
$payloadPath  = Join-Path $scriptFolder "$TaskName.ps1"

$verb = if ($Action -eq 'Grant') { 'Add-LocalGroupMember' } else { 'Remove-LocalGroupMember' }
$payload = @"
`$group = '$GroupName'
`$loggedOnUser = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
if (`$loggedOnUser) {
    $verb -Group `$group -Member `$loggedOnUser -ErrorAction SilentlyContinue
}
"@

Write-Host ""
Write-Host "  Set-LocalGroupSelfElevation : $Action '$GroupName' at next logon" -ForegroundColor Cyan
Write-Host "  Mode : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

if (-not $Apply) {
    Write-Host "  Would write payload script: $payloadPath" -ForegroundColor Yellow
    Write-Host "  Would register scheduled task '$TaskName' (trigger: at logon, run as SYSTEM)" -ForegroundColor Yellow
    Write-Host "  Would remove any existing opposing task '$opposingTaskName'" -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to perform these actions." -ForegroundColor Yellow
    Write-Host ""
    exit 0
}

if (-not $PSCmdlet.ShouldProcess($TaskName, "Register self-elevation scheduled task")) { exit 0 }

if (-not (Test-Path $scriptFolder)) { New-Item -Path $scriptFolder -ItemType Directory -Force | Out-Null }
Set-Content -Path $payloadPath -Value $payload -Encoding Unicode -Force

$existingOpposing = Get-ScheduledTask -TaskName $opposingTaskName -ErrorAction SilentlyContinue
if ($existingOpposing) {
    Unregister-ScheduledTask -TaskName $opposingTaskName -Confirm:$false
    Write-Host "  [OK]   Removed opposing task '$opposingTaskName'." -ForegroundColor Green
}

$trigger = New-ScheduledTaskTrigger -AtLogOn
$action  = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-ExecutionPolicy Bypass -NoProfile -File `"$payloadPath`""
Register-ScheduledTask -TaskName $TaskName -Trigger $trigger -Action $action -User 'SYSTEM' -Force | Out-Null
Start-ScheduledTask -TaskName $TaskName

Write-Host "  [OK]   Task '$TaskName' registered and triggered." -ForegroundColor Green
Write-Host ""
