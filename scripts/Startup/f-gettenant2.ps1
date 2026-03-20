function f-gettenant2
{
#file to open to
$file = "refreshtoken"  + "_" + "$env:computername" + ".txt"
$location = $env:USERPROFILE + "\OneDrive - Easy Office Online B.V\Documenten\Stored credentials\" + $file

$client_id = "3e47c504-4dd1-44a0-a469-cc2bd3c414dd"
$client_secret = "0KnqgMy/rjIgPtnhjcuPB5pDagHkdZPl8YEiyrJM3io="
$tenant_id = "58ce732e-45b5-4b6e-9cd0-780a538aa727"
$refreshTokensecure = Get-Content -Path $location

function Get-GCITSAccessTokenByResource($AppCredential, $tenantid, $Resource) {
$authority = "https://login.microsoftonline.com/$tenant_id"
$tokenEndpointUri = "$authority/oauth2/token"
$content = @{
    grant_type = "refresh_token"
    client_id = $appCredential.appID
    client_secret = $appCredential.secret
    resource = $resource
    refresh_token = $appCredential.refreshToken
}
$tokenEndpointUri = "$authority/oauth2/token"

$response = Invoke-RestMethod -Uri $tokenEndpointUri -Body $content -Method Post -UseBasicParsing
$access_token = $response.access_token
return $access_token
}

$AppCredential = @{
appId        = $client_id
secret       = $client_secret
refreshToken = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( (ConvertTo-SecureString $refreshTokensecure) ))
}

$MSGraphToken = Get-GCITSAccessTokenByResource -Resource "https://graph.microsoft.com" -tenantid $tenant_id -AppCredential $AppCredential
$AadGraphToken = Get-GCITSAccessTokenByResource -Resource "https://graph.windows.net" -tenantid $tenant_id -AppCredential $AppCredential
Connect-MsolService -MsGraphAccessToken $MSGraphToken -AdGraphAccessToken $AadGraphToken

$domainname = Read-Host "Voer de domeinnaam in"
$Customers = @()
$Customers = @(Get-MsolPartnerContract -DomainName $domainname)

$global:cid = $Customers.tenantid
     
Write-Host "$($Customers.name) selected. User the -tenantid `$cid parameter to run MSOL commands for this customer."
}