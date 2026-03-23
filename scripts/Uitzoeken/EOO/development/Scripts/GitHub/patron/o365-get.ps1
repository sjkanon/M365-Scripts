param(                        
    [switch]$debug = $false     ## if -debug parameter don't prompt for input
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Provide launch for Patron o365 get scripts

Source - https://github.com/directorcia/patron/blob/master/endpoint.ps1
Documentation - 
Notes - 

Prerequisites - 0

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$publicrepo = "..\Office365\"                   ## Default location on disk of free scripts repository

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

if ($debug) {
    write-host "Script activity logged at ..\o365-get.txt"
    start-transcript "..\o365-get.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

Clear-Host
write-host -foregroundcolor $systemmessagecolor "Script started`n"
write-host -ForegroundColor $processmessagecolor "Debug =",$debug
write-host -ForegroundColor $processmessagecolor "Prompt =",(-not $noprompt)

<# Test for Public repo #>
if (-not (test-path -path ($publicrepo))) {
    do {
        write-host -ForegroundColor yellow -backgroundcolor $errormessagecolor "[001] - Connection file directory", $publicrepo, "does not exist or is not found`n"
        $publicrepo = read-host -Prompt "`nEnter full path to connection files directory or press ENTER to end script"
    } until (([string]::isnullorempty($publicrepo)) -or (test-path -path ($publicrepo)))
    if ([string]::isnullorempty($publicrepo)) {
        Stop-Transcript | Out-Null      ## Terminate transcription
        exit 1                          ## Terminate script
    }
    else {
        write-host -ForegroundColor $processmessagecolor "Connection file directory found at", $publicrepo    
    }
}
else {
    write-host -ForegroundColor $processmessagecolor "Connection file directory found at", $publicrepo
}

$scripts = @()
$scripts += [PSCustomObject]@{
    Name = "o365-alerts-activity-get.ps1";
    Service = "Office 365";
    Context = "Security"
    Module = "ExchangeOnlineManagement";
    Description = "Read (get) tenant Activity alerts"    
}
$scripts += [PSCustomObject]@{
    Name = "o365-adal-get.ps1";
    Service = "Office 365"; 
    Context = "Security";
    Module = "ExchangeOnlineManagement";
    Description = "Read (get) tenant and user authentication policies"   
}
$scripts += [PSCustomObject]@{
    Name = "o365-mx-alert-get.ps1";
    Service = "Office 365";
    Context = "Security";
    Module = "ExchangeOnlineManagement";
    Description = "Read (get) tenant Protection alerts"    
}
$scripts += [PSCustomObject]@{
    Name = "o365-mx-archive-get.ps1";
    Service = "Office 365";
    Context = "Exchange";
    Module = "ExchangeOnlineManagement";
    Description = "Read (get) whether archiving has been enabled for mailbox"
}
$scripts += [PSCustomObject]@{
    Name = "o365-mx-audit-get.ps1";
    Service = "Office 365";
    Context = "Exchange";
    Module = "ExchangeOnlineManagement";
    Description = "Read (get) audit details for all mailboxes"    
}
$scripts += [PSCustomObject]@{
    Name = "o365-mx-auditage-get.ps1";
    Service = "Office 365";
    Context = "Exchange";
    Module = "ExchangeOnlineManagement";
    Description = "Read (get) the audit log length for all mailboxes"
}
$scripts += [PSCustomObject]@{
    Name = "o365-mx-be-get.ps1";
    Service = "Office 365";
    Context = "Exchange";
    Module = "ExchangeOnlineManagement";
    Description = "Read (get) status of Briefing email for all users"
}
$scripts += [PSCustomObject]@{
    Name = "o365-mx-check-exp.ps1";
    Service = "Office 365";
    Context = "Exchange";
    Module = "ExchangeOnlineManagement";
    Description = "Read (get) and export the status of mailboxes and tenant email settings"
}
$scripts += [PSCustomObject]@{
    Name = "o365-mx-connectionpolicy-get.ps1";
    Service = "Office 365";
    Context = "Exchange";
    Module = "ExchangeOnlineManagement";
    Description = "Read (get) the all connection filter policies and settings for each in a tenant"
}
$scripts += [PSCustomObject]@{
    Name = "o365-mx-inboxrules-get.ps1";
    Service = "Office 365";
    Context = "Exchange";
    Module = "ExchangeOnlineManagement";
    Description = "Read (get) which email boxes have forwarding options set"
}
$scripts += [PSCustomObject]@{
    Name = "o365-mx-junk-get.ps1";
    Service = "Office 365";
    Context = "Exchange";
    Module = "ExchangeOnlineManagement";
    Description = "Read (get) Display junk mail details for all mailboxes"
}
$scripts += [PSCustomObject]@{
    Name = "o365-mx-legal-get.ps1";
    Service = "Office 365";
    Context = "Exchange";
    Module = "ExchangeOnlineManagement";
    Description = "Read (get) litigation hold status of mailboxes"
}
$scripts += [PSCustomObject]@{
    Name = "o365-mx-malware-get.ps1";
    Service = "Office 365";
    Context = "Exchange";
    Module = "ExchangeOnlineManagement";
    Description = "Read (get) existing malware policies and checks these against best practices"
}
$scripts += [PSCustomObject]@{
    Name = "o365-mx-org-get.ps1";
    Service = "Office 365";
    Context = "Exchange";
    Module = "ExchangeOnlineManagement";
    Description = "Read (get) Exchange organizational and compare to best practices"
}
$scripts += [PSCustomObject]@{
    Name = "o365-mx-spam-get.ps1";
    Service = "Office 365";
    Context = "Exchange";
    Module = "ExchangeOnlineManagement";
    Description = "Read (get) existing spam policies and checks these against best practices"
}
$scripts += [PSCustomObject]@{
    Name = "o365-oauth-get.ps1";
    Service = "Office 365";
    Context = "Security";
    Module = "azuread";
    Description = "Read (get) Oauth tokens in tenant"
}
$scripts += [PSCustomObject]@{
    Name = "o365-NoSPO-ADAcct.ps1";
    Service = "Office 365";
    Context = "SharePoint";
    Module = "microsoft.online.sharepoint.powershell";
    Description = "Read (get) every current Azure AD user to see whether there is a corresponding SharePoint user"
}
$scripts += [PSCustomObject]@{
    Name = "o365-alerts-protect-get.ps1";
    Service = "Office 365";
    Context = "Security";
    Module = "ExchangeOnlineManagement";
    Description = "Read (get) Protection Alerts in the Security and Compliance Center"
}
$scripts += [PSCustomObject]@{
    Name = "o365-skype-get.ps1";
    Service = "Office 365";
    Context = "Skype for Business";
    Module = "skypeonlineconnector";
    Description = "Read (get) Skype for Business configuration"
}
$scripts += [PSCustomObject]@{
    Name = "o365-spo-extavail.ps1";
    Service = "Office 365";
    Context = "SharePoint";
    Module = "microsoft.online.sharepoint.powershell";
    Description = "Read (get) files that are accessible/have been shared externally"
}
$scripts += [PSCustomObject]@{
    Name = "o365-spo-extuser-30.ps1";
    Service = "Office 365";
    Context = "SharePoint";
    Module = "microsoft.online.sharepoint.powershell";
    Description = "Read (get) external users who have been added to SharePoint and ODFB in last 30 days"
}
$scripts += [PSCustomObject]@{
    Name = "o365-SPO-NoADAcct.ps1";
    Service = "Office 365";
    Context = "SharePoint";
    Module = "microsoft.online.sharepoint.powershell";
    Description = "Read (get) every current SharePoint user to see whether there is a corresponding Azure AD user"
}
$scripts += [PSCustomObject]@{
    Name = "o365-spo-orgconf-get.ps1";
    Service = "Office 365";
    Context = "SharePoint";
    Module = "microsoft.online.sharepoint.powershell";
    Description = "Read (get) SharePoint organisation configuration and compare to best practices"
}
$scripts += [PSCustomObject]@{
    Name = "o365-tms-get.ps1";
    Service = "Office 365";
    Context = "Teams";
    Module = "MicrosoftTeams";
    Description = "Read (get) Microsoft Teams configuration information for a tenant"
}

$results = $scripts | select-object name,description,Context | Sort-Object name | Out-GridView -PassThru -title "Select script(s) to run (Multiple selections permitted) "    

foreach ($result in $results) {
    foreach ($script in $scripts) {
        if ($result.name -eq $script.Name) {
            if (-not [string]::isnullorempty($script.module)) {             ## If a PowerShell module is required to be installed?
                if (get-module -listavailable -name $script.module) {       ## Has the Online PowerShell module been loaded?
                    write-host -ForegroundColor $processmessagecolor "Required",$script.module,"module found"
                }
                else {
                    write-host -ForegroundColor yellow -backgroundcolor $errormessagecolor "`n[010] - Online PowerShell module",$script.module,"not installed. Please install and re-run script`n"
                    Stop-Transcript                 ## Terminate transcription
                    exit 10                         ## Terminate script
                }
            }
            <# Test for script in current location #>
            if (-not (test-path -path $script.name)) {
                write-host -ForegroundColor yellow -backgroundcolor $errormessagecolor "`n[011] -",$script.name,"script not found in current directory - Please ensure exists first`n"
                Stop-Transcript | Out-Null      ## Terminate transcription
                exit 11                         ## Terminate script
            }
            else {
                write-host -ForegroundColor $processmessagecolor $script.name,"script found in current directory`n"
            }
            switch ($script.module) {
                "ExchangeOnlineManagement" {
                    try {
                        Get-Organizationconfig -ErrorAction continue | Out-Null
                    }
                    catch {
                        &($publicrepo+"o365-connect-exo.ps1") -wait            ## Connect to Exchange Online V2 and wait till complete 
                    };
                    break
                }
                "MicrosoftTeams" {
                    try {
                        Get-Team -ErrorAction continue | Out-Null
                    }
                    catch {
                        &($publicrepo+"o365-connect-tms.ps1") -wait            ## Connect to Microsoft Teams and wait till complete 
                    };
                    break
                }
                "microsoft.online.sharepoint.powershell" {
                    try {
                        Get-SPOTenant -ErrorAction continue | Out-Null
                    }
                    catch {
                        &($publicrepo+"o365-connect-spo.ps1") -wait            ## Connect to SharePoint and wait till complete 
                    };
                    break
                }
                "skypeonlineconnector" {
                    try {
                        get-csoauthconfiguration -ErrorAction continue | Out-Null
                    }
                    catch {
                        &($publicrepo+"o365-connect-s4b.ps1") -wait            ## Connect to Skype for Business and wait till complete 
                    };
                    break
                }
                "azuread" {
                    try {
                        Get-AzureADTenantDetail -ErrorAction continue | Out-Null
                    }
                    catch {
                        &($publicrepo+"o365-connect-aad.ps1") -wait            ## Connect to Skype for Business and wait till complete 
                    };
                    break
                }
            }
            if ($debug) {                             ## Is debug mode required?
                Write-Host -ForegroundColor $processmessagecolor (".\"+$script.name)
                &(".\"+$script.name) -debug           ## Run script
            }
            else {                                      ## If no debug mode is required
                Write-Host -ForegroundColor $processmessagecolor (".\"+$script.name)"-nodebug"
                &(".\"+$script.name)                    ## Run script
            }
        }
    }
}

write-host -foregroundcolor $systemmessagecolor "Script Complete`n"
if ($debug) {
    Stop-Transcript | Out-Null
}