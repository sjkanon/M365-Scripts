#Requires -Version 7.0
#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Reports, Microsoft.Graph.Applications
<#
.NOTES
    Generic MSP M365 management functions via Microsoft Graph and Exchange Online.
    Dot-source this file from your profile or startup script.

    Required variables (set by your startup script before dot-sourcing):
        $upn      — UPN of the logged-in administrator
        $realname — Display name for greeting (optional)

    CSP/partner operations: use Connect-Tenant to populate $global:cid and
    $global:connectmsoldomain. Individual functions then connect to the customer
    tenant via Connect-MgGraph -TenantId $cid.

    MSP-specific settings: see #region Configuration below.
#>

#region Configuration
# Customise these values for your organisation before dot-sourcing.

$script:MspAdminAlias       = 'msp-admin'
$script:MspAdminDisplayName = 'MSP - Admin Account'

#endregion

#region Startup

Connect-MgGraph -Scopes `
    'Domain.Read.All', 'Organization.Read.All', 'User.ReadWrite.All', `
    'Group.ReadWrite.All', 'RoleManagement.ReadWrite.Directory', `
    'AuditLog.Read.All', 'Application.Read.All', 'Domain.ReadWrite.All' `
    -NoWelcome

if ($realname) { Write-Host "Hey $realname. Good luck today!" }
else           { Write-Host "Hey $upn. Good luck today!" }

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
        else { Write-Warning 'Clipboard not available. Install xclip or wl-clipboard.' }
    }
}

#endregion

#region Menu

function Show-Menu {
    param ([string]$Title = 'Load modules')
    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host '1: Exchange Online'
    Write-Host '2: Microsoft Entra ID (Graph)'
    Write-Host '3: Microsoft Teams'
    Write-Host '4: Intune / Graph'
    Write-Host 'Q: Quit'
}

function Invoke-Menu {
    Show-Menu -Title 'Load modules'
    $selection = Read-Host 'Which modules do you want to load?'
    switch ($selection) {
        '1' {
            Write-Host 'Connecting to Exchange Online...'
            Connect-ExchangeOnline -UserPrincipalName $upn -DelegatedOrganization (Get-DefaultDomain)
        }
        '2' {
            Write-Host 'Connecting to Microsoft Entra ID...'
            Connect-MgGraph -TenantId $cid -Scopes `
                'User.ReadWrite.All', 'Group.ReadWrite.All', `
                'RoleManagement.ReadWrite.Directory', 'Domain.ReadWrite.All' `
                -NoWelcome
        }
        '3' {
            Write-Host 'Connecting to Microsoft Teams...'
            Import-Module MicrosoftTeams
            Connect-MicrosoftTeams -TenantId $cid
        }
        '4' {
            Write-Host 'Connecting to Intune / Graph...'
            Connect-MgGraph -TenantId $cid -Scopes `
                'DeviceManagementConfiguration.ReadWrite.All', `
                'DeviceManagementManagedDevices.ReadWrite.All' `
                -NoWelcome
        }
        'q' { return }
    }
}

#endregion

#region Connection / tenant selection

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
        $Domain = Read-Host 'Enter the domain you want to connect to'
    }
    $global:connectmsoldomain = $Domain
    $contract = Get-MgContract -Filter "defaultDomainName eq '$Domain'" -ErrorAction Stop
    if (-not $contract) { throw "No CSP contract found for domain '$Domain'." }
    $global:cid = $contract.CustomerId
    Write-Host "$($contract.DisplayName) selected. Use `$cid for Graph operations on this customer."
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
    param (
        [int]   $Language = 1043,
        [string]$TimeZone = 'W. Europe Standard Time'
    )
    Test-ExoConnection
    Get-Mailbox -ResultSize Unlimited |
        Select-Object -ExpandProperty PrimarySmtpAddress |
        Set-MailboxRegionalConfiguration -Language $Language -TimeZone $TimeZone -LocalizeDefaultFolderName
}

function Add-MailboxAlias {
    Test-ExoConnection
    $user  = Read-Host 'Which user do you want to add an alias to?'
    $alias = Read-Host 'Which alias?'
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

    Write-Output "Saved to: $csvFile"
}

function Set-AutoReply {
    Clear-Host
    $FormatEnumerationLimit = -1

    do {
        $mbname = Read-Host 'Enter the mailbox email address'
    } until ($mbname -like '*@*' -and $mbname -like '*.*')

    $message = Read-Host 'Paste the OOO message here (leave blank to disable)'
    $oooHtml = '<pre>' + $message + '</pre>'
    $mode    = Read-Host '(e)nabled  (d)isabled  (s)cheduled'
    $mbx     = Get-Mailbox -Identity $mbname

    switch -Regex ($mode) {
        '^e' { $mbx | Set-MailboxAutoReplyConfiguration -AutoReplyState Enabled  -ExternalMessage $oooHtml }
        '^d' { $mbx | Set-MailboxAutoReplyConfiguration -AutoReplyState Disabled }
        '^s' {
            $startTime = Read-Host 'Start time (e.g. 2026-04-01 08:00:00)'
            $endTime   = Read-Host 'End time   (e.g. 2026-04-10 18:00:00)'
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
    $addDomain = Read-Host 'Which domain name do you want to add?'
    New-MgDomain -Id $addDomain
    Start-Sleep -Seconds 5

    $txtRecord = Get-MgDomainVerificationDnsRecord -DomainId $addDomain
    Write-Host ($txtRecord | Where-Object { $_.RecordType -eq 'Txt' } | Select-Object -ExpandProperty AdditionalProperties | Out-String)
    Read-Host 'Press Enter once you have added the TXT record (TTL 1 minute)'

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
        Name       = 'Available'
        Expression = { $_.PrepaidUnits.Enabled - $_.ConsumedUnits }
    }
}

function Get-TenantUsers {
    Connect-MgGraph -TenantId $cid -Scopes 'User.Read.All' -NoWelcome
    Get-MgUser -All | Select-Object UserPrincipalName, DisplayName, AssignedLicenses
}

function Add-TenantAdmin {
    $setAsAdmin = Read-Host 'Which user do you want to grant admin rights? (UPN)'
    $user = Get-MgUser -UserId $setAsAdmin
    Add-GlobalAdminRole -UserId $user.Id
}

function Get-EntraApplication {
    $appName = Read-Host 'Name of the Enterprise App?'
    Get-MgApplication -Filter "displayName eq '$appName'"
}

function Reset-UserPassword {
    $resetAddress = Read-Host 'Enter the email address of the account to reset'
    $domain       = $resetAddress.Split('@')[1]

    $contract = Get-MgContract -Filter "defaultDomainName eq '$domain'" -ErrorAction Stop
    Connect-MgGraph -TenantId $contract.CustomerId -Scopes 'User.ReadWrite.All' -NoWelcome

    $newPassword = New-SecurePassword -Lowercase 8 -Uppercase 2 -Digits 2 -Special 2

    Update-MgUser -UserId $resetAddress -PasswordProfile @{
        Password                      = $newPassword
        ForceChangePasswordNextSignIn = $false
    }

    Write-Host ''
    Write-Host "The temporary password for $resetAddress is: $newPassword"
    Write-Host 'Please sign in at https://portal.office.com to set a new password.'
    Write-Host ''
    Write-Host 'Tip: Use a private browser window if another account signs in automatically.'
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

    $exportPath = Join-Path ([System.IO.Path]::GetTempPath()) 'SignInAudit'
    New-Item -Path $exportPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null

    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $allLogs  | Select-Object $selectProps | Export-Csv -Path (Join-Path $exportPath "AllSignIn_${ts}_${connectmsoldomain}.csv")  -NoTypeInformation -Encoding UTF8
    $failLogs | Select-Object $selectProps | Export-Csv -Path (Join-Path $exportPath "FailSignIn_${ts}_${connectmsoldomain}.csv") -NoTypeInformation -Encoding UTF8
    $goodLogs | Select-Object $selectProps | Export-Csv -Path (Join-Path $exportPath "GoodSignIn_${ts}_${connectmsoldomain}.csv") -NoTypeInformation -Encoding UTF8

    Write-Output "Logs saved to: $exportPath"
}

#endregion

#region MSP Admin Account

function New-MspAdmin {
    $password = New-SecurePassword -Lowercase 13 -Uppercase 2 -Digits 1 -Special 2

    Connect-MgGraph -TenantId $cid -Scopes 'User.ReadWrite.All', 'RoleManagement.ReadWrite.Directory' -NoWelcome

    $upnAdmin = "$($script:MspAdminAlias)@$(Get-DefaultDomain)"
    $user = New-MgUser `
        -DisplayName       $script:MspAdminDisplayName `
        -UserPrincipalName $upnAdmin `
        -MailNickname      $script:MspAdminAlias `
        -AccountEnabled    `
        -PasswordProfile   @{ Password = $password; ForceChangePasswordNextSignIn = $false }

    Add-GlobalAdminRole -UserId $user.Id
    Write-Host "Created: $upnAdmin"
}

function Set-MspAdminAsGroupOwner {
    $name = Read-Host 'What is the name of the group?'

    Connect-MgGraph -TenantId $cid -Scopes 'Group.ReadWrite.All', 'User.Read.All' -NoWelcome

    $group    = Get-MgGroup -Filter "displayName eq '$name'" | Select-Object -First 1
    $mspAdmin = Get-MgUser  -Filter "displayName eq '$($script:MspAdminDisplayName)'" | Select-Object -First 1

    New-MgGroupOwner -GroupId $group.Id -BodyParameter @{
        '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($mspAdmin.Id)"
    }
}

function Reset-MspAdminPassword {
    $password = New-SecurePassword -Lowercase 4 -Uppercase 2 -Digits 1 -Special 1

    Connect-MgGraph -TenantId $cid -Scopes 'User.ReadWrite.All' -NoWelcome

    $adminUpn = "$($script:MspAdminAlias)@$(Get-DefaultDomain)"
    Update-MgUser -UserId $adminUpn -PasswordProfile @{
        Password                      = $password
        ForceChangePasswordNextSignIn = $false
    }

    Write-Host "$adminUpn | $password"
    Set-ClipboardCrossPlatform $password
}

#endregion

#region Navigation

function Set-ImportLocation  { Set-Location $env:import }
function Set-ScriptsLocation { Set-Location $env:ps }

function prompt {
    $p = Split-Path -Leaf -Path (Get-Location)
    "$p> "
}

#endregion
