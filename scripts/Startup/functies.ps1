#Requires -Version 7.0
#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Reports, Microsoft.Graph.Applications
<#
.NOTES
    EOO M365 management functies — herschreven naar Microsoft Graph (MSOnline + AzureAD verwijderd).
    Dot-source dit bestand vanuit je profiel of startup script.

    Vereist variabelen die door het startup script worden gezet:
        $upn      — UPN van de ingelogde beheerder
        $realname — weergavenaam (optioneel)

    CSP/partner-operaties: gebruik Connect-Tenant om $global:cid en $global:connectmsoldomain te vullen,
    waarna individuele functies verbinding maken met de klant-tenant via Connect-MgGraph -TenantId $cid.
#>

#region Startup

Connect-MgGraph -Scopes `
    'Domain.Read.All', 'Organization.Read.All', 'User.ReadWrite.All', `
    'Group.ReadWrite.All', 'RoleManagement.ReadWrite.Directory', `
    'AuditLog.Read.All', 'Application.Read.All', 'Domain.ReadWrite.All' `
    -NoWelcome

if ($realname) { Write-Host "Heey $realname. Succes vandaag!" }
else           { Write-Host "Heey $upn. Succes vandaag!" }

#endregion

#region Helpers

function Get-RandomCharacters {
    param (
        [int]    $Length,
        [string] $Characters
    )
    $random = 1..$Length | ForEach-Object { Get-Random -Maximum $Characters.Length }
    $private:ofs = ''
    return [string]$Characters[$random]
}

function Invoke-ScrambleString {
    param ([string]$InputString)
    $chars = $InputString.ToCharArray()
    return -join ($chars | Get-Random -Count $chars.Length)
}

function New-SecurePassword {
    param (
        [int]$Lowercase = 8,
        [int]$Uppercase = 2,
        [int]$Digits    = 2,
        [int]$Special   = 2
    )
    Invoke-ScrambleString (
        (Get-RandomCharacters -Length $Lowercase -Characters 'abcdefghiklmnoprstuvwxyz') +
        (Get-RandomCharacters -Length $Uppercase -Characters 'ABCDEFGHKLMNOPRSTUVWXYZ') +
        (Get-RandomCharacters -Length $Digits    -Characters '1234567890') +
        (Get-RandomCharacters -Length $Special   -Characters '!"$%&/()=?@#*+')
    )
}

function Get-DefaultDomain {
    (Get-MgDomain | Where-Object { $_.IsDefault }).Id
}

function Add-GlobalAdminRole {
    param ([Parameter(Mandatory)][string]$UserId)
    $role = Get-MgDirectoryRole | Where-Object { $_.DisplayName -eq 'Global Administrator' }
    New-MgDirectoryRoleMember -DirectoryRoleId $role.Id -BodyParameter @{
        '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId"
    }
}

function Set-ClipboardCrossPlatform {
    param ([string]$Text)
    if ($IsWindows) {
        Set-Clipboard $Text
    } elseif ($IsMacOS) {
        $Text | pbcopy
    } elseif ($IsLinux) {
        if (Get-Command wl-copy  -ErrorAction SilentlyContinue) { $Text | wl-copy }
        elseif (Get-Command xclip -ErrorAction SilentlyContinue) { $Text | xclip -selection clipboard }
        else { Write-Warning 'Klembord niet beschikbaar. Installeer xclip of wl-clipboard.' }
    }
}

#endregion

#region Menu

function Show-Menu {
    param ([string]$Title = 'Modules laden')
    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host '1: Exchange Online'
    Write-Host '2: Microsoft Entra ID (Graph)'
    Write-Host '3: Microsoft Teams'
    Write-Host '4: Intune / Graph'
    Write-Host 'Q: Afsluiten'
}

function Invoke-Menu {
    Show-Menu -Title 'Modules laden'
    $selection = Read-Host 'Welke modules wil je laden?'
    switch ($selection) {
        '1' {
            Write-Host 'Verbinding maken met Exchange Online...'
            Connect-ExchangeOnline -UserPrincipalName $upn -DelegatedOrganization (Get-DefaultDomain)
        }
        '2' {
            Write-Host 'Verbinding maken met Microsoft Entra ID...'
            Connect-MgGraph -TenantId $cid -Scopes `
                'User.ReadWrite.All', 'Group.ReadWrite.All', `
                'RoleManagement.ReadWrite.Directory', 'Domain.ReadWrite.All' `
                -NoWelcome
        }
        '3' {
            Write-Host 'Verbinding maken met Microsoft Teams...'
            Import-Module MicrosoftTeams
            Connect-MicrosoftTeams -TenantId $cid
        }
        '4' {
            Write-Host 'Verbinding maken met Intune / Graph...'
            Connect-MgGraph -TenantId $cid -Scopes `
                'DeviceManagementConfiguration.ReadWrite.All', `
                'DeviceManagementManagedDevices.ReadWrite.All' `
                -NoWelcome
        }
        'q' { return }
    }
}

#endregion

#region Verbinding / tenant selectie

function Test-ExoConnection {
    try {
        Get-AcceptedDomain -Identity $connectmsoldomain -ErrorAction Stop | Out-Null
    }
    catch {
        Connect-ExchangeOnline -UserPrincipalName $upn -DelegatedOrganization $connectmsoldomain
    }
}

function Connect-Tenant {
    param ([string]$Domain)
    if (-not $Domain) {
        $Domain = Read-Host 'Wat is het domein waarmee je wil verbinden?'
    }
    $global:connectmsoldomain = $Domain
    $contract = Get-MgContract -Filter "defaultDomainName eq '$Domain'" -ErrorAction Stop
    if (-not $contract) { throw "Geen CSP-contract gevonden voor domein '$Domain'." }
    $global:cid = $contract.CustomerId
    Write-Host "$($contract.DisplayName) geselecteerd. Gebruik `$cid voor Graph-operaties op deze klant."
}

#endregion

#region Exchange Online

function Enable-CopyOfSentItems {
    Test-ExoConnection
    Get-Mailbox -ResultSize Unlimited |
        Select-Object -ExpandProperty PrimarySmtpAddress |
        Set-Mailbox -MessageCopyForSentAsEnabled $true
}

function Add-SharedMailboxAccess {
    param (
        [Parameter(Mandatory)] [string]$User,
        [Parameter(Mandatory)] [string]$Mailbox,
        [switch]$AutoMapping
    )
    Test-ExoConnection
    $mbx = Get-Mailbox -Identity $Mailbox
    Remove-MailboxPermission -Identity $mbx -User $User -AccessRights FullAccess -ErrorAction SilentlyContinue -Confirm:$false
    Add-MailboxPermission    -Identity $mbx -User $User -AccessRights FullAccess -AutoMapping:$AutoMapping.IsPresent
    Add-RecipientPermission  -Identity $mbx -Trustee $User -AccessRights SendAs  -Confirm:$false
}

function Set-MailboxLocale {
    Test-ExoConnection
    Get-Mailbox -ResultSize Unlimited |
        Select-Object -ExpandProperty PrimarySmtpAddress |
        Set-MailboxRegionalConfiguration -Language 1043 -TimeZone 'W. Europe Standard Time' -LocalizeDefaultFolderName
}

function Add-MailboxAlias {
    Test-ExoConnection
    $user  = Read-Host 'Aan welke gebruiker wil je een alias toevoegen?'
    $alias = Read-Host 'Welke alias?'
    Set-Mailbox $user -EmailAddresses @{ add = $alias }
}

function Get-MailboxAliases {
    Get-Mailbox -ResultSize Unlimited |
        Select-Object DisplayName, @{
            Name       = 'EmailAddresses'
            Expression = { $_.EmailAddresses | Where-Object { $_ -like 'SMTP:*' } }
        }
}

function Export-DistributionGroups {
    $csvFile = Join-Path ([System.IO.Path]::GetTempPath()) 'ExportDGs.csv'

    Get-DistributionGroup -ResultSize Unlimited | ForEach-Object {
        $members = Get-DistributionGroupMember $_.Name
        [PSCustomObject]@{
            Name                               = $_.Name
            DisplayName                        = $_.DisplayName
            PrimarySmtpAddress                 = $_.PrimarySmtpAddress
            SecondarySmtpAddress               = ($_.EmailAddresses | Where-Object { $_ -clike 'smtp:*' } | ForEach-Object { $_ -replace 'smtp:', '' }) -join ','
            Alias                              = $_.Alias
            GroupType                          = $_.GroupType
            RecipientType                      = $_.RecipientType
            Members                            = ($members.Name -join ',')
            MembersPrimarySmtpAddress          = ($members.PrimarySmtpAddress -join ',')
            ManagedBy                          = $_.ManagedBy.Name
            HiddenFromAddressLists             = $_.HiddenFromAddressListsEnabled
            MemberJoinRestriction              = $_.MemberJoinRestriction
            MemberDepartRestriction            = $_.MemberDepartRestriction
            RequireSenderAuthenticationEnabled = $_.RequireSenderAuthenticationEnabled
            AcceptMessagesOnlyFrom             = ($_.AcceptMessagesOnlyFrom.Name -join ',')
            GrantSendOnBehalfTo                = $_.GrantSendOnBehalfTo.Name
        }
    } | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8

    Write-Output "Opgeslagen in: $csvFile"
}

function Set-AutoReply {
    Clear-Host
    $FormatEnumerationLimit = -1

    do {
        $mbname = Read-Host 'Voer het e-mailadres in van de mailbox'
    } until ($mbname -like '*@*' -and $mbname -like '*.*')

    $message = Read-Host 'Plak de OOO-tekst hier (leeg laten om uit te schakelen)'
    $oooHtml = '<pre>' + $message + '</pre>'
    $mode    = Read-Host '(e)nabled  (d)isabled  (s)cheduled'
    $mbx     = Get-Mailbox -Identity $mbname

    switch -Regex ($mode) {
        '^e' { $mbx | Set-MailboxAutoReplyConfiguration -AutoReplyState Enabled  -ExternalMessage $oooHtml }
        '^d' { $mbx | Set-MailboxAutoReplyConfiguration -AutoReplyState Disabled }
        '^s' {
            $startTime = Read-Host 'Starttijd (bijv. 2026-04-01 08:00:00)'
            $endTime   = Read-Host 'Eindtijd  (bijv. 2026-04-10 18:00:00)'
            $mbx | Set-MailboxAutoReplyConfiguration `
                -AutoReplyState Scheduled `
                -InternalMessage $oooHtml -ExternalMessage $oooHtml `
                -StartTime $startTime -EndTime $endTime
        }
    }

    $mbx | Get-MailboxAutoReplyConfiguration |
        Select-Object AutoReplyState, StartTime, EndTime, InternalMessage, ExternalMessage |
        Format-List
}

#endregion

#region Microsoft Entra ID / Graph

function Get-TenantAdmins {
    Connect-MgGraph -TenantId $cid -Scopes 'RoleManagement.Read.Directory' -NoWelcome
    $role = Get-MgDirectoryRole | Where-Object { $_.DisplayName -eq 'Global Administrator' }
    Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id |
        Select-Object Id, DisplayName, AdditionalProperties
}

function Add-TenantDomain {
    $addDomain = Read-Host 'Welke domeinnaam wil je toevoegen?'
    New-MgDomain -Id $addDomain
    Start-Sleep -Seconds 5

    $txtRecord = Get-MgDomainVerificationDnsRecord -DomainId $addDomain
    Write-Host ($txtRecord | Where-Object { $_.RecordType -eq 'Txt' } | Select-Object -ExpandProperty AdditionalProperties | Out-String)
    Read-Host 'Druk op Enter zodra je het TXT-record (TTL 1 minuut) hebt toegevoegd'

    Confirm-MgDomain -DomainId $addDomain
    Start-Sleep -Seconds 5
    Update-MgDomain -DomainId $addDomain -SupportedServices @('Email')
    Start-Sleep -Seconds 5

    Get-MgDomainServiceConfigurationRecord -DomainId $addDomain |
        Where-Object { $_.RecordType -eq 'Mx' } |
        Select-Object -ExpandProperty AdditionalProperties
}

function Get-TenantLicenses {
    Get-MgSubscribedSku | Select-Object SkuPartNumber, ConsumedUnits, @{
        Name       = 'Beschikbaar'
        Expression = { $_.PrepaidUnits.Enabled - $_.ConsumedUnits }
    }
}

function Get-TenantUsers {
    Connect-MgGraph -TenantId $cid -Scopes 'User.Read.All' -NoWelcome
    Get-MgUser -All | Select-Object UserPrincipalName, DisplayName, AssignedLicenses
}

function Add-TenantAdmin {
    $setAsAdmin = Read-Host 'Welke gebruiker wil je adminrechten geven? (UPN)'
    $user = Get-MgUser -UserId $setAsAdmin
    Add-GlobalAdminRole -UserId $user.Id
}

function Get-EntraApplication {
    $appName = Read-Host 'Naam van de Enterprise App?'
    Get-MgApplication -Filter "displayName eq '$appName'"
}

function Reset-UserPassword {
    $resetAddress = Read-Host 'Voer het e-mailadres in waarvan je het wachtwoord wilt resetten'
    $domain       = $resetAddress.Split('@')[1]

    $contract = Get-MgContract -Filter "defaultDomainName eq '$domain'" -ErrorAction Stop
    Connect-MgGraph -TenantId $contract.CustomerId -Scopes 'User.ReadWrite.All' -NoWelcome

    $newPassword = New-SecurePassword -Lowercase 8 -Uppercase 2 -Digits 2 -Special 2

    Update-MgUser -UserId $resetAddress -PasswordProfile @{
        Password                      = $newPassword
        ForceChangePasswordNextSignIn = $false
    }

    Write-Host ''
    Write-Host "Het tijdelijke wachtwoord van $resetAddress is: $newPassword"
    Write-Host 'Graag inloggen op https://portal.office.com om een nieuw wachtwoord in te stellen.'
    Write-Host ''
    Write-Host 'Tip: Open de browser in privémodus als er automatisch een ander account inlogt.'
}

function Export-SignInLogs {
    [CmdletBinding()]
    param (
        [ValidateRange(1, 30)]
        [int]$Days = 30
    )

    Connect-MgGraph -TenantId $cid -Scopes 'AuditLog.Read.All' -NoWelcome

    $startDate  = (Get-Date).AddDays(-$Days).ToString('yyyy-MM-dd')
    $endDate    = (Get-Date).ToString('yyyy-MM-dd')
    $baseFilter = "createdDateTime ge $startDate and createdDateTime le $endDate"

    # Filter server-side to reduce data transfer
    $failLogs = Get-MgAuditLogSignIn -Filter "$baseFilter and status/errorCode ne 0" -All
    $goodLogs = Get-MgAuditLogSignIn -Filter "$baseFilter and status/errorCode eq 0"  -All
    $allLogs  = @($failLogs) + @($goodLogs)

    $selectProps = @(
        'CreatedDateTime', 'UserPrincipalName', 'RiskState', 'AppId', 'ClientAppUsed', 'IpAddress',
        @{ N = 'City';            E = { $_.Location.City } },
        @{ N = 'CountryOrRegion'; E = { $_.Location.CountryOrRegion } },
        @{ N = 'FailureReason';   E = { $_.Status.FailureReason } },
        'ConditionalAccessStatus'
    )

    $exportPath = Join-Path ([System.IO.Path]::GetTempPath()) 'AzureADSignInAudit'
    New-Item -Path $exportPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null

    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $allLogs  | Select-Object $selectProps | Export-Csv -Path (Join-Path $exportPath "AllSignIn_${ts}_${connectmsoldomain}.csv")  -NoTypeInformation -Encoding UTF8
    $failLogs | Select-Object $selectProps | Export-Csv -Path (Join-Path $exportPath "FailSignIn_${ts}_${connectmsoldomain}.csv") -NoTypeInformation -Encoding UTF8
    $goodLogs | Select-Object $selectProps | Export-Csv -Path (Join-Path $exportPath "GoodSignIn_${ts}_${connectmsoldomain}.csv") -NoTypeInformation -Encoding UTF8

    Write-Output "Logs opgeslagen in: $exportPath"
}

#endregion

#region EOO Beheeraccount

function New-EooAdmin {
    $password = New-SecurePassword -Lowercase 13 -Uppercase 2 -Digits 1 -Special 2

    Connect-MgGraph -TenantId $cid -Scopes 'User.ReadWrite.All', 'RoleManagement.ReadWrite.Directory' -NoWelcome

    $upnAdmin = "eooadmin@$(Get-DefaultDomain)"
    $user = New-MgUser `
        -DisplayName      'Easy Office Online - Beheeraccount' `
        -UserPrincipalName $upnAdmin `
        -MailNickname     'eooadmin' `
        -AccountEnabled   `
        -PasswordProfile  @{ Password = $password; ForceChangePasswordNextSignIn = $false }

    Add-GlobalAdminRole -UserId $user.Id
    Write-Host "Aangemaakt: $upnAdmin"
}

function Set-EooAsGroupOwner {
    $name = Read-Host 'Wat is de naam van de groep?'

    Connect-MgGraph -TenantId $cid -Scopes 'Group.ReadWrite.All', 'User.Read.All' -NoWelcome

    $group    = Get-MgGroup -Filter "displayName eq '$name'" | Select-Object -First 1
    $eooAdmin = Get-MgUser  -Filter "displayName eq 'Easy Office Online - Beheeraccount'" | Select-Object -First 1

    New-MgGroupOwner -GroupId $group.Id -BodyParameter @{
        '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($eooAdmin.Id)"
    }
}

function Reset-EooPassword {
    $password = New-SecurePassword -Lowercase 4 -Uppercase 2 -Digits 1 -Special 1

    Connect-MgGraph -TenantId $cid -Scopes 'User.ReadWrite.All' -NoWelcome

    $eooAdmin = "eooadmin@$(Get-DefaultDomain)"
    Update-MgUser -UserId $eooAdmin -PasswordProfile @{
        Password                      = $password
        ForceChangePasswordNextSignIn = $false
    }

    Write-Host "$eooAdmin | $password"
    Set-ClipboardCrossPlatform $password
}

#endregion

#region Navigatie

function Set-ImportLocation  { Set-Location $env:import }
function Set-ScriptsLocation { Set-Location $env:ps }

function prompt {
    $p = Split-Path -Leaf -Path (Get-Location)
    "$p> "
}

#endregion
