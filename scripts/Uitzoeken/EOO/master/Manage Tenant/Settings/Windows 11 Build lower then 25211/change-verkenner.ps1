mkdir 'C:\tmp\Verkenner'
Invoke-WebRequest -Uri "https://endpoint.eoo.cloud/Installers/Verkenner/change.reg" -OutFile "C:\tmp\Verkenner\change.reg"
Set-Location C:\tmp\Verkenner
reg import .\change.reg