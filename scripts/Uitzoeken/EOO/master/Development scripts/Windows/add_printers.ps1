Set-Location C:\ 
mkdir tmp
Invoke-WebRequest -Uri "https://endpoint.eoo.cloud/printers/safely/PrintExport.printerExport" -OutFile "C:\tmp\printerexport.PrinterExport"
Invoke-WebRequest -Uri "https://endpoint.eoo.cloud/printers/00_scripts/add_printers.bat" -OutFile "C:\tmp\add_printers.bat"
Set-Location C:\tmp
$out = Start-Process C:\tmp\add_printers.bat
Write-Host $out
