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
$path = $(Join-Path $env:ProgramFiles EOO) 
if (!(Test-Path $path)) 
{ 
New-Item -Path $path -ItemType Directory -Force -Confirm:$false 
} 
Out-File -FilePath $(Join-Path $env:ProgramFiles EOO\Restart-NTP.ps1) -Encoding unicode -Force -InputObject $content -Confirm:$false 
  
#register script as scheduled task 
$Time = New-ScheduledTaskTrigger -Once -At 8am -RepetitionDuration  (New-TimeSpan -Days 9999)  -RepetitionInterval  (New-TimeSpan -Minutes 59)
$User = "SYSTEM"
$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ex bypass -file `"C:\ProgramFiles\EOO\Restart-NTP.ps1`"" 
Register-ScheduledTask -TaskName "Restart NTP" -Trigger $Time -User $User -Action $Action -Force 
