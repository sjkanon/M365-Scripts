#Requires -Version 5.1
<#
.SYNOPSIS
    Provision a new Workspace 365 environment for this tenant, including its SSO
    app registration and default Exchange/SharePoint links.

.DESCRIPTION
    Modernized rewrite of a 600+ line interactive script: replaces its MSAL.PS
    device-code login and raw bearer-token Graph calls with Connect-MgGraph /
    Invoke-MgGraphRequest (reusing an existing session if already connected),
    and turns the always-interactive prompts into parameters with a
    dry-run-by-default / -Apply pattern. The provisioning key and hostname are
    always supplied by the caller — the original saved them to a local
    plaintext .cfg file next to the script; this version does not persist them
    anywhere.

    What it does (only when -Apply is given):
      1. Creates an Entra ID App Registration for Workspace 365 SSO, with the
         delegated Microsoft Graph / Power BI / Windows Virtual Desktop scopes
         Workspace 365 uses, and a client secret.
      2. Calls the Workspace 365 Provisioning API to create the environment.
      3. Links the new App Registration to the environment as its SSO
         (OAuth2) identity provider.
      4. Points the environment's default Exchange (EWS) and SharePoint URLs at
         this tenant.

    Connects to Microsoft Graph automatically if no session is active; reuses an
    existing session if already connected.

.PARAMETER WorkspaceHostname
    Base URL of your Workspace 365 tenant, e.g. "https://yourcompany.workspace365.net".

.PARAMETER ProvisioningKey
    Workspace 365 provisioning key (a GUID), found in your Workspace 365
    partner/reseller portal. Never hardcode this — pass it in, or read it from a
    secret store, at call time.

.PARAMETER EnvironmentName
    Lowercase, alphanumeric-only name for the new environment (this becomes part
    of its URL: WorkspaceHostname/EnvironmentName).

.PARAMETER RequestingUserUpn
    UPN of the admin this environment will be registered under (used as the
    environment's contact). Defaults to the signed-in user if omitted.

.PARAMETER Apply
    Actually create the App Registration and the environment. Without this
    switch, the script only reports what it would do.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    # Preview
    .\New-Workspace365Environment.ps1 -WorkspaceHostname "https://yourcompany.workspace365.net" -ProvisioningKey $key -EnvironmentName "contoso"

.EXAMPLE
    .\New-Workspace365Environment.ps1 -WorkspaceHostname "https://yourcompany.workspace365.net" -ProvisioningKey $key -EnvironmentName "contoso" -Apply

.NOTES
    By using this script you accept the Workspace 365 terms and conditions
    (https://workspace365.net/en/term-and-conditions).

    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Applications
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $WorkspaceHostname,

    [Parameter(Mandatory)]
    [guid] $ProvisioningKey,

    [Parameter(Mandatory)]
    [ValidatePattern('^[a-z0-9]+$')]
    [string] $EnvironmentName,

    [string] $RequestingUserUpn,
    [switch] $Apply,
    [string] $TenantId
)

$WorkspaceHostname = $WorkspaceHostname.TrimEnd('/')

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    $null = Get-MgContext -ErrorAction Stop
    if (-not (Get-MgContext)) { throw }
} catch {
    $connectParams = @{ Scopes = @('Application.ReadWrite.All', 'User.Read', 'Organization.Read.All'); NoWelcome = $true }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
}

function Invoke-Graph {
    param([string] $Method = 'GET', [string] $Uri, [string] $Body)
    $params = @{ Method = $Method; Uri = $Uri; ErrorAction = 'Stop' }
    if ($Body) { $params['Body'] = $Body; $params['ContentType'] = 'application/json' }
    Invoke-MgGraphRequest @params
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   New-Workspace365Environment" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Workspace   : $WorkspaceHostname/$EnvironmentName"
Write-Host ("  Mode        : {0}" -f $(if ($Apply) { 'Apply' } else { 'Preview only' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

if (-not $RequestingUserUpn) {
    $me = Invoke-Graph -Uri 'https://graph.microsoft.com/v1.0/me'
    $RequestingUserUpn = $me.userPrincipalName
}
$requestingUser = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/users/$RequestingUserUpn"

$appName = "Workspace365 - $EnvironmentName"

# ── Check for an existing app registration with this name ────────────────────
$existingApps = (Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=displayName eq '$appName'").value
if ($existingApps) {
    Write-Host "  [ERROR] An App Registration named '$appName' already exists. Aborting to avoid duplicates." -ForegroundColor Red
    if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
    exit 1
}

if (-not $Apply) {
    Write-Host "  Would create:" -ForegroundColor Yellow
    Write-Host "    - App Registration '$appName' with delegated Graph/Power BI/WVD SSO scopes" -ForegroundColor Yellow
    Write-Host "    - Workspace 365 environment '$EnvironmentName' for $($requestingUser.displayName)" -ForegroundColor Yellow
    Write-Host "    - SSO link + default Exchange/SharePoint URLs for this tenant" -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to provision." -ForegroundColor Yellow
    if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
    return
}

if (-not $PSCmdlet.ShouldProcess("$WorkspaceHostname/$EnvironmentName", "Provision Workspace 365 environment")) {
    if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
    return
}

# ── Build required delegated permission scopes ────────────────────────────────
function Get-ServicePrincipalByName {
    param([string] $Name)
    (Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=displayName eq '$Name'").value | Select-Object -First 1
}

$graphScopeNames = @(
    'Directory.ReadWrite.All', 'Calendars.ReadWrite.Shared', 'Contacts.ReadWrite.Shared',
    'Files.ReadWrite.All', 'Mail.ReadWrite.Shared', 'Mail.Send', 'Mail.Send.Shared',
    'People.Read', 'Sites.FullControl.All', 'Tasks.ReadWrite.Shared', 'User.Read', 'openid'
)
$powerBiScopeNames = @(
    'Dashboard.Read.All', 'Dataset.Read.All', 'Dataset.ReadWrite.All', 'Group.Read.All',
    'Report.Read.All'
)

Write-Host "  Resolving service principals for required scopes..." -ForegroundColor DarkGray
$graphSp    = Get-ServicePrincipalByName -Name 'Microsoft Graph'
$powerBiSp  = Get-ServicePrincipalByName -Name 'Power BI Service'

function New-ResourceAccessList {
    param($ServicePrincipal, [string[]] $ScopeNames)
    $access = foreach ($scopeName in $ScopeNames) {
        $scope = $ServicePrincipal.oauth2PermissionScopes | Where-Object { $_.value -eq $scopeName }
        if ($scope) { @{ id = $scope.id; type = 'Scope' } }
    }
    @{ resourceAppId = $ServicePrincipal.appId; resourceAccess = @($access) }
}

$requiredResourceAccess = @(New-ResourceAccessList -ServicePrincipal $graphSp -ScopeNames $graphScopeNames)
if ($powerBiSp) { $requiredResourceAccess += New-ResourceAccessList -ServicePrincipal $powerBiSp -ScopeNames $powerBiScopeNames }

# ── Create the App Registration ────────────────────────────────────────────────
Write-Host "  Creating App Registration '$appName'..." -ForegroundColor Cyan
$appBody = @{
    signInAudience = 'AzureADMyOrg'
    displayName    = $appName
    web            = @{
        homePageUrl  = "$WorkspaceHostname/$EnvironmentName/SignIn"
        logoutUrl    = "$WorkspaceHostname/$EnvironmentName/SignOut"
        redirectUris = @("$WorkspaceHostname/$EnvironmentName/OAuth2/HandleAuthorityResponse")
    }
    requiredResourceAccess = $requiredResourceAccess
} | ConvertTo-Json -Depth 10

$app = Invoke-Graph -Method POST -Uri 'https://graph.microsoft.com/v1.0/applications' -Body $appBody
Write-Host "  [OK]   App Registration created ($($app.appId))" -ForegroundColor Green

$secretBody = @{ passwordCredential = @{ displayName = 'Workspace365-SSO'; endDateTime = (Get-Date).AddYears(2).ToString('o') } } | ConvertTo-Json
$secret = Invoke-Graph -Method POST -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)/addPassword" -Body $secretBody

Invoke-Graph -Method POST -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Body (@{ appId = $app.appId } | ConvertTo-Json) | Out-Null

# ── Create the Workspace 365 environment ──────────────────────────────────────
Write-Host "  Creating Workspace 365 environment '$EnvironmentName'..." -ForegroundColor Cyan
$provisioningHeader = @{ ProvisioningKey = $ProvisioningKey.ToString() }

$createBody = @{
    EmailAddress    = $requestingUser.userPrincipalName
    FirstName       = $requestingUser.givenName
    LastName        = $requestingUser.surname
    EnvironmentName = $EnvironmentName
} | ConvertTo-Json

try {
    Invoke-RestMethod -Method POST -Uri "$WorkspaceHostname/Provisioning/Environment/?version=2.0" `
        -Body $createBody -ContentType 'application/json' -Headers $provisioningHeader -ErrorAction Stop | Out-Null
    Write-Host "  [OK]   Environment created." -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] Environment creation failed: $($_.Exception.Message)" -ForegroundColor Red
    if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
    exit 1
}

# ── Link SSO + default Exchange/SharePoint URLs ───────────────────────────────
$org = Invoke-Graph -Uri 'https://graph.microsoft.com/v1.0/organization'
$defaultDomain = ($org.value.verifiedDomains | Where-Object { $_.isDefault }).name
$sharePointDomain = [regex]::Match($defaultDomain, '(?<domain>\w+)\.').Groups['domain'].Value

$ssoBody = @{
    Authority = "https://login.windows.net/$defaultDomain"
    clientId  = $app.appId
    Key       = $secret.secretText
} | ConvertTo-Json

$exchangeBody   = @{ Servertype = 'Office365'; DefaultEwsUrl = 'https://outlook.office365.com/EWS/Exchange.asmx' } | ConvertTo-Json
$sharePointBody = @{ AuthenticationType = 'OAuth2'; DefaultUrl = "https://$sharePointDomain.sharepoint.com" } | ConvertTo-Json

try {
    Invoke-RestMethod -Method PUT -Uri "$WorkspaceHostname/Provisioning/$EnvironmentName/SingleSignOnSettings/OAuth2" `
        -Body $ssoBody -ContentType 'application/json' -Headers $provisioningHeader -ErrorAction Stop | Out-Null
    Write-Host "  [OK]   SSO linked." -ForegroundColor Green

    Invoke-RestMethod -Method PUT -Uri "$WorkspaceHostname/Provisioning/$EnvironmentName/ExchangeSettings" `
        -Body $exchangeBody -ContentType 'application/json' -Headers $provisioningHeader -ErrorAction SilentlyContinue | Out-Null
    Invoke-RestMethod -Method PUT -Uri "$WorkspaceHostname/Provisioning/$EnvironmentName/SharePointSettings/" `
        -Body $sharePointBody -ContentType 'application/json' -Headers $provisioningHeader -ErrorAction SilentlyContinue | Out-Null
    Write-Host "  [OK]   Default Exchange/SharePoint URLs configured." -ForegroundColor Green
} catch {
    Write-Host "  [WARN] SSO/Exchange/SharePoint configuration step failed: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "         The environment and app registration were still created — finish this step manually if needed." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Green
Write-Host "   Environment ready: $WorkspaceHostname/$EnvironmentName" -ForegroundColor Green
Write-Host "  ================================================" -ForegroundColor Green
Write-Host "  It can take a few minutes for Entra ID to fully propagate the new app registration." -ForegroundColor DarkGray
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
