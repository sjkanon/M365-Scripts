Start-Sleep -Seconds 15
Set-Location '.\Scripts (Do not touch)\makegroups'
./make-SG_ALL_Except_EOO.ps1
Write-Host "Security Group SG-EOO is gereed"
./make-SG_LastPass.ps1
Write-Host "Security Group Lastpass is aangemaakt en gereed"
./make-SG_Workspace.ps1
Write-Host "Security Group Workspace is aangemaakt en gereed" 
Start-Sleep -Seconds 2
f-menuinstall
