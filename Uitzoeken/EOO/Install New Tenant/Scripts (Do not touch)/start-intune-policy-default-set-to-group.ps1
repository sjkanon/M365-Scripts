Connect-AzureAD
$defaultgroup = (Get-AzureADmsGroup -SearchString "SG_All_Exept_EOO").id
$lastpassgroup = (Get-AzureADmsGroup -SearchString "SG - lastpass").id
$coligogroup = (Get-AzureADmsGroup -SearchString "SG - coligo").id
$allusers = (Get-AzureADmsGroup -SearchString "SG_All_Exept_EOO").id
Set-Location "$env:setuptenantdev/Scripts (Do not Touch)"
$setupgroups = Read-Host "Do you want to setup tenant in group variable"
if($setupgroups -eq 'yes'){

./set-intune-default-policy-to-group.ps1 -DefaultAzureadgroup $defaultgroup
./set-intune-default-policy-to-group.ps1 -LastpassAzureadgroup $lastpassgroup
./set-intune-default-policy-to-group.ps1 -ColigoAzureadgroup $coligogroup
}
else {
    ./set-intune-all-default-policy-to-group.ps1 -AllusersAzureADGroup $allusers
}
