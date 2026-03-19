$taskscheduled = "Upgrade Applicaties"
$tasktoremove = "Upgrade Applicaties"
$content = @' 
choco upgrade all -y
Sleep -Seconds 30
 winget upgrade --all --accept-package-agreements --accept-source-agreements
'@ 

 # create custom folder and write PS script 
$path = $(Join-Path $env:ProgramData EOO\Scripts) 
if (!(Test-Path $path)) 
{ 
New-Item -Path $path -ItemType Directory -Force -Confirm:$false 
} 
Remove-Item -Path "$env:ProgramData\EOO\Scripts\upgrade-programs.ps1" -Force -Confirm:$false -ErrorAction SilentlyContinue
Out-File -FilePath $(Join-Path $env:ProgramData EOO\Scripts\upgrade-programs.ps1) -Encoding unicode -Force -InputObject $content -Confirm:$false

$admingroup = Get-ScheduledTask -TaskName $tasktoremove
if ($admingroup) {
    Unregister-ScheduledTask -TaskName $tasktoremove -confirm:$false 
}
  
# register script as scheduled task 
$Time = New-ScheduledTaskTrigger -AtLogOn 
$User = "SYSTEM" 
$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ex bypass -file `"$path\upgrade-programs.ps1`"" 
Register-ScheduledTask -TaskName $taskscheduled -Trigger $Time -User $User -Action $Action -Force 
Start-ScheduledTask -TaskName $taskscheduled