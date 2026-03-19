##install script
Write-Host "Welkom bij het installatie script om de tenant in te gaan richten"
Start-Sleep -Seconds 3
Write-Host "Laten we beginnen met de tenant te selecteren"
Start-Sleep -Milliseconds 300
##get tenant
$installtenant = Read-Host "Welk tenant ga je aanmaken? 'klant'.onmicrosoft.com"
$tenant = (Get-MsolPartnerContract -domainname "$installtenant.onmicrosoft.com").tenantid
$klantselect = (Get-MsolPartnerInformation -TenantId $tenant).PartnerCompanyName
Write-Host  "Je hebt nu toegang tot $klantselect "
Start-Sleep -Milliseconds 120

## To Do
Write-Host "Wat wil je gaan doen"
function Show-installmenu
{
    param (
        [string]$Title = 'Modules laden'
    )
    Clear-Host
    Write-Host "================ $Title ================"
    
    Write-Host "1: EOO-ADMIN maken"
    Write-Host "2: Set-Groups"
    Write-Host "3: Microsoft Teams"
    Write-Host "4: MS Graph"
    Write-Host "Q: Press 'Q' to quit."
}
 

function f-menuinstall
 {
    Show-installmenu
 $installselect = Read-Host 'Welke modules wil je laden'
 switch ($installselect)
 {
     '1' {
         'Make eoo-admin en maak hem beheerder'
         $MsolDomains = Get-MsolDomain -TenantId $tenant

        $regex = '^[^.]*\.onmicrosoft\.com$'
        $Domainname = $MsolDomains |
            Where-Object Name -Match $regex |
            Select-Object -ExpandProperty Name |
            Select-Object -First 1

         Connect-ExchangeOnline -UserPrincipalName $upn -DelegatedOrganization $domainname
         Write-Host "We have Connection"
         Set-Location "$env:setuptenantdev\Scripts (Do not touch)"
         ./start-exchange.ps1
     } '2' {
         'You chose option #2'
         Import-Module AzureADpreview
         Connect-AzureAD -AccountId $upn -TenantId $tenant
         Write-Host "We have Connection"
         Set-Location "$env:setuptenantdev\Scripts (Do not touch)"
         ./start-azuread.ps1
     } '3' {
         'You chose option #3'
         Import-Module MicrosoftTeams
         Connect-MicrosoftTeams -tenantid $tenant
     } '4' {
         'You chose option #4'
        Import-Module -Name MSGraphFunctions
        Import-Module -Name IntuneBackupAndRestore
        Update-MSGraphEnvironment -AuthUrl "https://login.microsoftonline.com/$tenant"
        Connect-MSGraph
     } 
     'q' {
         return
     }
 }}
 f-menuinstall
 ## Connect to M365 Exchange
