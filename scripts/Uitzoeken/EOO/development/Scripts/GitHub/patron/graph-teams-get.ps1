<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Query the Microsoft Graph for details on Teams in the tenant
Documentation - https://github.com/directorcia/patron/wiki/Get-Teams-details
Source - https://github.com/directorcia/patron/blob/master/graph-teams-get.ps1

Prerequisites = 2
1. Azure AD app setup per - https://blog.ciaops.com/2019/04/17/using-interactive-powershell-to-access-the-microsoft-graph/
2. Save token details via script - https://github.com/directorcia/patron/blob/master/graph-creds-save.ps1

Graph Permissions
Permission type	Permissions (from least to most privileged)
Delegated (work or school account) = Group.Read.All, Group.ReadWrite.All, Directory.Read.All, Directory.ReadWrite.All, Directory.AccessAsUser.All
Delegated (personal Microsoft account) = Not supported.
Application	= Group.Read.All, Directory.Read.All, Group.ReadWrite.All, Directory.ReadWrite.All

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

# Application (client) ID, tenant ID and secret
$clientidcreds = import-clixml -path ..\clientid.xml
$tenantidcreds = import-clixml -path ..\tenantid.xml
$clientsecretcreds = import-clixml -path ..\clientsec.xml

write-host -foregroundcolor $processmessagecolor "Decrypt credentials`n"

$clientid = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($clientIdcreds.password))
$tenantid = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($tenantIdcreds.password))
$clientsecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($clientsecretcreds.password))

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

## start-transcript "..\o365-graph-teams-get $(get-date -f yyyyMMddHHmmss).txt"

write-host -foregroundcolor $systemmessagecolor "Script started`n"

## Script from - https://www.lee-ford.co.uk/getting-started-with-microsoft-graph-with-powershell/

# Azure AD OAuth Application Token for Graph API
# Get OAuth token for a AAD Application (returned as $token)

# Construct URI
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

# Construct Body
$body = @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}

write-host -foregroundcolor $processmessagecolor "Get OAuth 2.0 Token"
# Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

# Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

# Graph API call in PowerShell using obtained OAuth token (see other gists for more details)

# Specify the URI to call and method
$uri = "https://graph.microsoft.com/beta/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&`$top=999"
$method = "GET"

write-host -foregroundcolor $processmessagecolor "Get Teams summary"
# Run Graph API query 
$query = Invoke-WebRequest -Method $method -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token" } -ErrorAction Stop -UseBasicParsing

$ConvertedOutput = $query.content | ConvertFrom-Json
$TeamSummary = @()                 ## Results array
foreach ($control in $convertedoutput.value) {
    if ($control.displayname -match "%20") {
        $control.displayname = $control.displayname.replace("%20", " ")
    }
    $TeamSummary += [pscustomobject]@{        ## Build array item
        Displayname = $control.displayname
        Mail        = $control.mail
        Visibility  = $control.visibility
        Id          = $control.id
    }
}

$TeamSummary | select-object Displayname, Visibility, Mail, Id | Format-Table

write-host -foregroundcolor $processmessagecolor "Get Channel details for Team"

foreach ($team in $TeamSummary) {
    # Specify the URI to call and method
    write-host -ForegroundColor yellow -BackgroundColor darkmagenta "`nTeam ="$team.displayname, "(Summary)"
    $teamidentity = $team.id
    $uri = "https://graph.microsoft.com/beta/teams/$teamidentity/channels"
    $method = "GET"

    # Run Graph API query 
    $query = Invoke-WebRequest -Method $method -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token" } -ErrorAction Stop -UseBasicParsing

    $ConvertedOutput = $query.content | ConvertFrom-Json
    $ChannelSummary = @()                 ## Results array
    foreach ($channel in $convertedoutput.value) {
        if ($channel.displayname -match "%20") {
            $channel.displayname = $channel.displayname.replace("%20", " ")
        }
        $ChannelSummary += [pscustomobject]@{        ## Build array item
            Teamname       = $team.displayname
            TeamID         = $team.id
            Channelname    = $channel.displayname
            Description    = $channel.description
            Email          = $channel.email
            ChannelId      = $channel.id
            WebUrl         = $channel.WebUrl
            Membershiptype = $channel.membershiptype   
        }
    }
    $ChannelSummary | Select-Object Channelname, Membershiptype, ChannelId | format-table

    foreach ($channel in $ChannelSummary) {
        write-host -foregroundcolor yellow -BackgroundColor darkmagenta "Team ="$team.displayname, "" -NoNewline
        write-host -ForegroundColor Yellow -BackgroundColor darkcyan " [Channel] ="$channel.channelname
        $channelidentity = $channel.channelid
        $teamidentity = $channel.TeamID
        $uri = "https://graph.microsoft.com/beta/teams/$teamidentity/channels/$channelidentity/tabs"
        $method = "GET"

        # Run Graph API query 
        $query = Invoke-WebRequest -Method $method -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token" } -ErrorAction Stop -UseBasicParsing

        $ConvertedOutput = $query.content | ConvertFrom-Json
        $TabSummary = @()                 ## Results array
        foreach ($tab in $convertedoutput.value) {
            if ($tab.displayname -match "%20") {
                $tab.displayname = $tab.displayname.replace("%20", " ")
            }
            $tabSummary += [pscustomobject]@{        ## Build array item
                Teamname    = $team.displayname
                TeamID      = $team.id
                Channelname = $channel.channelname
                ChannelId   = $channel.channelid
                Tabname     = $tab.displayname
                TabId       = $tab.id   
            }
        }
        $tabSummary | Select-Object Tabname, TabId | format-table
    }
}

write-host -foregroundcolor $systemmessagecolor "`nScript Completed`n"
