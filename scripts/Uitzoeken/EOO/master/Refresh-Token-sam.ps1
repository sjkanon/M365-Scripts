### User Input Variables ###

### Enter the details of your Secure Access Model Application below ###

$ApplicationId           = '9f7ccc60-7d10-4670-a595-978ee545e0e2'
$ApplicationSecret       = 'YZQ8Q~PVNvcq-2Q024pYaWCFlW4A2.WuLcy6ac.w'
$TenantId                = '58ce732e-45b5-4b6e-9cd0-780a538aa727'

### STOP EDITING HERE ###

### Create credential object using UserEntered(ApplicationID) and UserEntered(ApplicationSecret) ###

$Credential = New-Object System.Management.Automation.PSCredential($ApplicationId, ($ApplicationSecret | ConvertTo-SecureString -AsPlainText -Force))

### Splat Params required for Updating Refresh Token ###

$UpdateRefreshTokenParamaters = @{
    ApplicationID        = $ApplicationId
    Tenant               = $TenantId
    Scopes               = 'https://api.partnercenter.microsoft.com/user_impersonation'
    Credential           = $Credential
    UseAuthorizationCode = $true
    ServicePrincipal     = $true
}

### Splat Params required for Updating Exchange Refresh Token ###

$UpdateExchangeTokenParamaters = @{
    ApplicationID           = 'a0c73c16-a7e3-4564-9a95-2bdf47383716'
    Scopes                  = 'https://outlook.office365.com/.default'
    Tenant                  = $TenantId
    UseDeviceAuthentication = $true
}

### Create new Refresh Token using previously splatted paramaters ###

$Token = New-PartnerAccessToken @UpdateRefreshTokenParamaters

### Create new Exchange Refresh Token using previously splatted paramaters ###

$Exchangetoken = New-PartnerAccessToken @UpdateExchangeTokenParamaters

### Output Refresh Tokens and Exchange Refresh Tokens ###

Write-Host "================ Secrets ================"
Write-Host "`$ApplicationId         = $($ApplicationId)"
Write-Host "`$ApplicationSecret     = $($ApplicationSecret)"
Write-Host "`$TenantID              = $($TenantId)"
Write-Host "`$RefreshToken          = $($Token.refreshtoken)" -ForegroundColor Blue
Write-Host "`$ExchangeRefreshToken  = $($ExchangeToken.Refreshtoken)" -ForegroundColor Green
Write-Host "================ Secrets ================"
Write-Host "     SAVE THESE IN A SECURE LOCATION     "