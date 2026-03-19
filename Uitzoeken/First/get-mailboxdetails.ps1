Get-Date
Write-Host "Run Script" 
Get-MailboxStatistics -Identity "Christophe@groupsuerickx.be" -Archive | Select DisplayName, TotalItemSize, ItemCount
Sleep 120