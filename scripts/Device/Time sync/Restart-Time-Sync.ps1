<#
.SYNOPSIS
    Repair Windows time synchronisation and keep it repaired with a scheduled task.

.DESCRIPTION
    Sets w32time to start automatically, points it at the Dutch NTP pool
    (0/1.nl.pool.ntp.org) with a manual sync flag and forces a resync. The same
    commands are registered as a scheduled task that reruns every 59 minutes, because
    a single resync does not hold on a machine whose clock keeps drifting.
#>

#Set content for taskscheduler
$content = @'
Set-Service 'w32time' -StartupType Automatic
net start W32time
w32tm /config /manualpeerlist:"0.nl.pool.ntp.org 1.nl.pool.ntp.org" /syncfromflags:manual /update
w32tm /resync
w32tm /config /manualpeerlist:"0.nl.pool.ntp.org 1.nl.pool.ntp.org" /syncfromflags:manual /update
w32tm /resync
'@ 
 
#create custom folder and write PS script 
$path = $(Join-Path $env:ProgramFiles TimeSyncIntune) 
if (!(Test-Path $path)) 
{ 
New-Item -Path $path -ItemType Directory -Force -Confirm:$false 
} 
Out-File -FilePath $(Join-Path $env:ProgramFiles TimeSyncIntune\Restart-NTP.ps1) -Encoding unicode -Force -InputObject $content -Confirm:$false 
  
#register script as scheduled task 
$Time = New-ScheduledTaskTrigger -Once -At 8am -RepetitionDuration  (New-TimeSpan -Days 9999)  -RepetitionInterval  (New-TimeSpan -Minutes 59)
$User = "SYSTEM"
$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ex bypass -file `"C:\ProgramFiles\TimeSyncIntune\Restart-NTP.ps1`"" 
Register-ScheduledTask -TaskName "Restart NTP" -Trigger $Time -User $User -Action $Action -Force 
