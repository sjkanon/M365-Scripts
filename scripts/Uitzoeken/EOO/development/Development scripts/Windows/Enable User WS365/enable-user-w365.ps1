$lid = Read-Host "Welk lid wil je toevoegen aan Workspace"
Set-Mailbox -Identity $lid -CustomAttribute1 "Workspace365"
Write-Host "Successfully"