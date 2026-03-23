cd $env:USERPROFILE
cd 
"\Desktop\windows spooler"

$uname = (Get-WmiObject -Class Win32_ComputerSystem | Select-Object username)
./subinacl.exe /SERVICE $svc_name /grant=$uname=F
Write-Host = "Succes"