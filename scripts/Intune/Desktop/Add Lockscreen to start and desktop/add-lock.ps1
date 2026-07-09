$folder = 'C:\Program Files\EOO\'

if(-not(Test-Path -Path $folder -PathType Container)){
Set-Location 'C:\Program Files'
mkdir EOO
Set-Location 'C:\Program Files\EOO'
mkdir lockworkstation
Invoke-WebRequest -Uri "https://endpoint.eoo.cloud/lock/lock.txt" -OutFile "C:\Program Files\EOO\lockworkstation\lock.txt"
Invoke-WebRequest -Uri "https://endpoint.eoo.cloud/lock/lock.ico" -OutFile "C:\Program Files\EOO\lockworkstation\lock.ico"
Set-Location 'C:\Program Files\EOO\lockworkstation'
mv lock.txt lock.bat
Set-Location 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs'
mkdir EOO
Set-Location 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\EOO'
mkdir lockworkstation
$TargetFile = "C:\Windows\explorer.exe"
$shortcutFile = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\EOO\lockworkstation\lock.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$shortcut.TargetPath = $TargetFile
$shortcut.IconLocation = "C:\Program Files\EOO\lockworkstation\lock.ico"
$Shortcut.Arguments =
  "C:\Program Files\EOO\lockworkstation\lock.bat"
$shortcut.Save()
}
else{}
