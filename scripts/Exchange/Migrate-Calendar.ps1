#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Applications, Microsoft.Graph.Authentication, Microsoft.Graph.Calendar, Microsoft.Graph.Groups, Microsoft.Graph.Users

<#
.SYNOPSIS
    Migreert de Onco3R "Holidays" M365 Group kalender naar een Room/Resource Mailbox.
    Maakt automatisch een App Registration aan als geen ClientId/ClientSecret opgegeven is.

.DESCRIPTION
    Script voor BraveHub - Ticket #0298048 - Onco3R Therapeutics
    Geschreven door: Sjoerd Kanon
    Datum: 19/03/2026

    Het script werkt in twee modi:

    MODUS A - Volledig automatisch (aanbevolen, eerste keer):
        Geen ClientId/ClientSecret opgeven. Het script maakt zelf een App Registration
        aan met de juiste application permissions en gebruikt die direct.

    MODUS B - Bestaande App Registration:
        ClientId en ClientSecret opgeven. Het script slaat de setup stap over.

    Stappen:
    1.  Platform detecteren (Windows/macOS/Linux)
    2.  Exchange Online verbinden
    3.  Graph verbinden (delegated) voor App Registration setup
    4.  App Registration aanmaken + admin consent verlenen (alleen Modus A)
    5.  Graph herverbinden met application permissions (client credentials)
    6.  Room Mailbox aanmaken
    7.  Kalender permissies + AutoAccept instellen
    8.  M365 Group "Holidays" opzoeken
    9.  Afspraken ophalen uit groepskalender
    10. Afspraken kopieren naar Room Mailbox
    11. (Optioneel) M365 Group verwijderen
    12. Samenvatting

    Waarom Room Mailbox i.p.v. M365 Group?
    - Geen e-mailnotificaties naar het hele bedrijf bij nieuwe events
    - Werkt als vergaderzaal: attendee toevoegen = verschijnt op de kalender
    - Iedereen kan de kalender lezen zonder lid te zijn
    - AutoAccept zodat verlof automatisch wordt goedgekeurd

    Installeer modules indien nodig:
        Install-Module ExchangeOnlineManagement      -Scope CurrentUser
        Install-Module Microsoft.Graph.Applications  -Scope CurrentUser
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
        Install-Module Microsoft.Graph.Calendar      -Scope CurrentUser
        Install-Module Microsoft.Graph.Groups        -Scope CurrentUser
        Install-Module Microsoft.Graph.Users         -Scope CurrentUser

.PARAMETER TenantId
    Azure AD Tenant ID (Entra ID > Overview > Tenant ID)

.PARAMETER AdminUPN
    UPN van de admin die het script uitvoert

.PARAMETER ClientId
    AppId van een bestaande App Registration (optioneel - wordt aangemaakt als leeg)

.PARAMETER ClientSecret
    Client Secret van de App Registration (optioneel - wordt aangemaakt als leeg)

.PARAMETER SourceGroupMail
    E-mailadres van de bestaande M365 Group

.PARAMETER SourceGroupDisplayName
    DisplayName van de M365 Group (fallback als mail niet gevonden wordt)

.PARAMETER RoomDisplayName
    Weergavenaam van de nieuwe Room Mailbox

.PARAMETER RoomAlias
    Alias voor de nieuwe Room Mailbox (moet uniek zijn in de tenant)

.PARAMETER RoomEmail
    Primair SMTP-adres van de nieuwe Room Mailbox

.PARAMETER DaysBack
    Hoeveel dagen terug afspraken ophalen (default: 365)

.PARAMETER DaysForward
    Hoeveel dagen vooruit afspraken ophalen (default: 730)

.PARAMETER DeleteSourceGroup
    Als $true wordt de M365 Group verwijderd na migratie (default: $false)

.EXAMPLE
    # Eerste keer - volledig automatisch
    .\Migrate-HolidaysCalendar.ps1 -TenantId "xxxx" -AdminUPN "admin@onco3r.com"

.EXAMPLE
    # Met bestaande App Registration
    .\Migrate-HolidaysCalendar.ps1 -TenantId "xxxx" -AdminUPN "admin@onco3r.com" -ClientId "yyyy" -ClientSecret "zzzz"

.EXAMPLE
    # Dry run
    .\Migrate-HolidaysCalendar.ps1 -TenantId "xxxx" -AdminUPN "admin@onco3r.com" -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId = "JOUW-TENANT-ID-HIER",

    [Parameter(Mandatory = $false)]
    [string]$AdminUPN = "admin@onco3r.com",

    # App Registration - leeg laten om automatisch aan te maken
    [Parameter(Mandatory = $false)]
    [string]$ClientId = "",

    [Parameter(Mandatory = $false)]
    [string]$ClientSecret = "",

    [Parameter(Mandatory = $false)]
    [string]$AppName = "BraveHub-HolidaysCalendarMigration",

    # Bestaande M365 Group
    [Parameter(Mandatory = $false)]
    [string]$SourceGroupMail = "holidays@onco3r.com",

    [Parameter(Mandatory = $false)]
    [string]$SourceGroupDisplayName = "Holidays",

    # Nieuwe Room Mailbox
    [Parameter(Mandatory = $false)]
    [string]$RoomDisplayName = "Holidays Calendar",

    [Parameter(Mandatory = $false)]
    [string]$RoomAlias = "holidays-calendar",

    [Parameter(Mandatory = $false)]
    [string]$RoomEmail = "holidays-calendar@onco3r.com",

    # Migratie tijdsvenster
    [Parameter(Mandatory = $false)]
    [int]$DaysBack = 365,

    [Parameter(Mandatory = $false)]
    [int]$DaysForward = 730,

    # Verwijder M365 Group na migratie
    [Parameter(Mandatory = $false)]
    [bool]$DeleteSourceGroup = $false
)

#region Helpers

function Write-Step { param([string]$m) Write-Host "`n==> $m" -ForegroundColor Cyan }
function Write-OK   { param([string]$m) Write-Host "    [OK]   $m" -ForegroundColor Green }
function Write-Warn { param([string]$m) Write-Host "    [WARN] $m" -ForegroundColor Yellow }
function Write-Fail { param([string]$m) Write-Host "    [FAIL] $m" -ForegroundColor Red }

function Connect-GraphDelegated {
    param([string]$TenantId, [string[]]$Scopes)
    if ($runOnWindows) {
        Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -NoWelcome
    } else {
        Write-Host "    Device code flow: kopieer de code en open de URL in je browser" -ForegroundColor DarkGray
        Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -UseDeviceAuthentication -NoWelcome
    }
}

function Connect-GraphAppAuth {
    param([string]$TenantId, [string]$ClientId, [string]$ClientSecret)
    # ClientSecretCredential via Microsoft.Graph.Authentication
    # Werkt met Graph SDK 2.x - gebruikt MSAL client credentials flow
    $secure     = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($ClientId, $secure)

    # Probeer eerst de nieuwere SDK methode
    try {
        Connect-MgGraph `
            -TenantId               $TenantId `
            -ClientSecretCredential $credential `
            -NoWelcome
        return
    } catch {
        Write-Warn "ClientSecretCredential methode mislukt, probeer via environment variables..."
    }

    # Fallback: via omgevingsvariabelen (werkt op alle SDK versies)
    $env:AZURE_CLIENT_ID     = $ClientId
    $env:AZURE_CLIENT_SECRET = $ClientSecret
    $env:AZURE_TENANT_ID     = $TenantId
    Connect-MgGraph -EnvironmentVariable -NoWelcome

    # Opruimen
    Remove-Item Env:AZURE_CLIENT_ID     -ErrorAction SilentlyContinue
    Remove-Item Env:AZURE_CLIENT_SECRET -ErrorAction SilentlyContinue
    Remove-Item Env:AZURE_TENANT_ID     -ErrorAction SilentlyContinue
}

#endregion

#region Stap 1: Platform detecteren

Write-Step "Platform detecteren"

$runOnMacOS   = ($IsMacOS -eq $true)
$runOnLinux   = ($IsLinux -eq $true)
$runOnWindows = ($IsWindows -eq $true) -or ($PSVersionTable.PSVersion.Major -le 5)

if (-not $runOnMacOS -and -not $runOnLinux -and -not $runOnWindows) {
    Write-Warn "Platform onbekend - device code flow wordt gebruikt als fallback"
    $runOnMacOS = $true
}

if ($runOnMacOS)   { Write-OK "Platform: macOS  - device code flow voor authenticatie" }
elseif ($runOnLinux)   { Write-OK "Platform: Linux  - device code flow voor authenticatie" }
elseif ($runOnWindows) { Write-OK "Platform: Windows - interactieve browser login" }

#endregion

#region Stap 2: Exchange Online verbinden

Write-Step "Verbinding maken met Exchange Online"
try {
    if ($runOnWindows) {
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false
    } else {
        Write-Host "    Device code flow: kopieer de code en open de URL in je browser" -ForegroundColor DarkGray
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false -Device
    }
    Write-OK "Exchange Online verbonden als $AdminUPN"
} catch {
    Write-Error "Exchange Online verbinding mislukt: $_"
    exit 1
}

#endregion

#region Stap 3 + 4: App Registration aanmaken of hergebruiken

$setupScopes = @("Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All", "Group.Read.All")

if (-not $ClientId -or -not $ClientSecret) {

    Write-Step "App Registration aanmaken (Modus A - geen ClientId/ClientSecret opgegeven)"

    # Delegated verbinding voor setup
    try {
        Connect-GraphDelegated -TenantId $TenantId -Scopes $setupScopes
        Write-OK "Graph verbonden (delegated) voor setup"
    } catch {
        Write-Error "Graph verbinding mislukt: $_"
        exit 1
    }

    # App aanmaken of hergebruiken
    $app = Get-MgApplication -Filter "displayName eq '$AppName'" -ErrorAction SilentlyContinue

    if ($app) {
        Write-Warn "App '$AppName' bestaat al (AppId: $($app.AppId)) - wordt hergebruikt"
    } else {
        if ($PSCmdlet.ShouldProcess($AppName, "App Registration aanmaken")) {
            $app = New-MgApplication -DisplayName $AppName
            Write-OK "App aangemaakt: $($app.DisplayName) | AppId: $($app.AppId)"
        }
    }

    # Graph service principal ophalen voor permission IDs
    $graphResourceId = "00000003-0000-0000-c000-000000000000"
    $graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphResourceId'"

    $requiredPerms = @("Calendars.ReadWrite", "Calendars.Read", "Group.Read.All", "Group.ReadWrite.All", "User.Read.All")
    $appRoles = [System.Collections.Generic.List[object]]::new()
    foreach ($perm in $requiredPerms) {
        $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $perm }
        if ($role) {
            Write-OK "Permission: $perm"
            $appRoles.Add([PSCustomObject]@{ Id = $role.Id; Type = "Role" })
        } else {
            Write-Warn "Permission '$perm' niet gevonden"
        }
    }

    # Permissions instellen op de app
    if ($PSCmdlet.ShouldProcess($app.AppId, "Permissions instellen")) {
        $resourceAccess = $appRoles | ForEach-Object {
            @{ id = $_.Id.ToString(); type = $_.Type }
        }
        Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess @(
            @{
                resourceAppId  = $graphResourceId
                resourceAccess = @($resourceAccess)
            }
        )
        Write-OK "Application permissions ingesteld"
    }

    # Service Principal aanmaken voor admin consent
    $appSp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue
    if (-not $appSp) {
        if ($PSCmdlet.ShouldProcess($app.AppId, "Service Principal aanmaken")) {
            $appSp = New-MgServicePrincipal -AppId $app.AppId
            Write-OK "Service Principal aangemaakt"
            Start-Sleep -Seconds 5
        }
    } else {
        Write-OK "Service Principal bestaat al"
    }

    # Admin consent verlenen
    foreach ($role in $appRoles) {
        try {
            if ($PSCmdlet.ShouldProcess($role.Id, "Admin consent verlenen")) {
                New-MgServicePrincipalAppRoleAssignment `
                    -ServicePrincipalId $appSp.Id `
                    -PrincipalId        $appSp.Id `
                    -ResourceId         $graphSp.Id `
                    -AppRoleId          $role.Id `
                    -ErrorAction Stop | Out-Null
                Write-OK "Admin consent verleend: $($requiredPerms[$appRoles.IndexOf($role)])"
            }
        } catch {
            if ($_ -match "already exists") {
                Write-Warn "Consent al aanwezig - overgeslagen"
            } else {
                Write-Warn "Consent mislukt: $_"
            }
        }
    }

    # Client Secret aanmaken
    if ($PSCmdlet.ShouldProcess($app.AppId, "Client Secret aanmaken")) {
        $secretEndDate = (Get-Date).AddMonths(3)
        $secret = Add-MgApplicationPassword `
            -ApplicationId $app.Id `
            -PasswordCredential @{
                DisplayName = "HolidaysMigration-$(Get-Date -Format 'yyyyMMdd')"
                EndDateTime = $secretEndDate
            }
        $ClientId     = $app.AppId
        $ClientSecret = $secret.SecretText
        Write-OK "Client Secret aangemaakt (vervalt: $($secretEndDate.ToString('dd/MM/yyyy')))"
        Write-Host ""
        Write-Host "    Bewaar deze waarden in Vaultwarden:" -ForegroundColor Yellow
        Write-Host "    ClientId    : $ClientId" -ForegroundColor Yellow
        Write-Host "    ClientSecret: $ClientSecret" -ForegroundColor Yellow
        Write-Host ""
    }

    # Wachten op consent propagatie in Azure AD (kan 30-60s duren)
    Write-Host "    Wachten 60s op admin consent propagatie in Azure AD..." -ForegroundColor DarkGray
    Start-Sleep -Seconds 60

    # Herverbinden met application permissions
    Write-Host "    Herverbinden met application permissions..." -ForegroundColor DarkGray
    Disconnect-MgGraph | Out-Null
    Start-Sleep -Seconds 5

} else {
    Write-Step "Bestaande App Registration gebruiken (Modus B)"
    Write-OK "ClientId: $ClientId"
}

#endregion

#region Stap 5: Graph verbinden met application permissions

Write-Step "Graph verbinden met application permissions"
try {
    Connect-GraphAppAuth -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    $ctx = Get-MgContext
    if (-not $ctx) { throw "Get-MgContext leeg na verbinding" }
    Write-OK "Graph verbonden | App: $($ctx.AppName) | AuthType: $($ctx.AuthType)"
} catch {
    Write-Error "Graph app auth mislukt: $_"
    exit 1
}

#endregion

#region Stap 6: Room Mailbox aanmaken

Write-Step "Room Mailbox controleren / aanmaken: $RoomEmail"

$existingRoom = Get-Mailbox -Identity $RoomEmail -ErrorAction SilentlyContinue

if ($existingRoom) {
    Write-Warn "Mailbox '$RoomEmail' bestaat al - stap overgeslagen"
} else {
    if ($PSCmdlet.ShouldProcess($RoomEmail, "Room Mailbox aanmaken")) {
        New-Mailbox `
            -Name               $RoomDisplayName `
            -Alias              $RoomAlias `
            -PrimarySmtpAddress $RoomEmail `
            -Room | Out-Null
        Write-OK "Room Mailbox aangemaakt: $RoomDisplayName ($RoomEmail)"
        Write-Host "    Wachten 15s op Exchange initialisatie..." -ForegroundColor DarkGray
        Start-Sleep -Seconds 15
    }
}

#endregion

#region Stap 7: Permissies en AutoAccept

Write-Step "Kalender permissies instellen op $RoomEmail"

if ($PSCmdlet.ShouldProcess($RoomEmail, "Permissies instellen")) {
    Set-MailboxFolderPermission `
        -Identity     "$($RoomEmail):\Calendar" `
        -User         Default `
        -AccessRights Reviewer `
        -ErrorAction  SilentlyContinue
    Write-OK "Default gebruikers: Reviewer (lezen met details)"

    Set-CalendarProcessing `
        -Identity              $RoomEmail `
        -AutomateProcessing    AutoAccept `
        -AllowConflicts        $true `
        -AddOrganizerToSubject $false `
        -DeleteComments        $false `
        -DeleteSubject         $false `
        -BookingWindowInDays   730
    Write-OK "AutoAccept ingesteld (overlappende verloven toegestaan)"
}

#endregion

#region Stap 8: Herverbinden delegated voor groepskalender lezen

# Microsoft Graph blokkeert Get-MgGroupCalendarEvent voor application permissions (AppOnly).
# Dit is een bekende beperking: https://learn.microsoft.com/en-us/graph/known-issues#group-calendar
# Oplossing: tijdelijk herverbinden als delegated gebruiker om de groepskalender te lezen,
# daarna terugschakelen naar app auth om te schrijven naar de Room Mailbox.

Write-Step "Herverbinden als delegated gebruiker (vereist voor groepskalender)"
Write-Host "    Microsoft Graph blokkeert groepskalender lezen via app auth." -ForegroundColor DarkGray
Write-Host "    Delegated verbinding nodig voor stap 8 en 9." -ForegroundColor DarkGray

# Volledig disconnecten en MSAL token cache wissen voor een schone sessie
try { Disconnect-MgGraph | Out-Null } catch {}
Start-Sleep -Seconds 3

# Schone delegated verbinding - module opnieuw importeren om token cache te wissen
Remove-Module Microsoft.Graph.Authentication -Force -ErrorAction SilentlyContinue
Import-Module Microsoft.Graph.Authentication -Force

try {
    if ($runOnWindows) {
        Connect-MgGraph `
            -TenantId $TenantId `
            -Scopes @("Group.Read.All", "Calendars.Read", "Calendars.ReadWrite", "User.Read.All") `
            -NoWelcome
    } else {
        Write-Host "    Device code flow: kopieer de code en open de URL in je browser" -ForegroundColor DarkGray
        Connect-MgGraph `
            -TenantId $TenantId `
            -Scopes @("Group.Read.All", "Calendars.Read", "Calendars.ReadWrite", "User.Read.All") `
            -UseDeviceAuthentication `
            -NoWelcome
    }

    $ctx = Get-MgContext
    if (-not $ctx) { throw "Geen context na verbinding" }
    Write-OK "Graph herverbonden (delegated) als $($ctx.Account)"
} catch {
    Write-Error "Delegated herverbinding mislukt: $_"
    exit 1
}

#endregion

#region Stap 9: M365 Group opzoeken + afspraken ophalen

Write-Step "M365 Group opzoeken"

$group = $null

$mailLower = $SourceGroupMail.ToLower()
$group = Get-MgGroup -Filter "mail eq '$mailLower'" -ErrorAction SilentlyContinue

if (-not $group) {
    Write-Warn "Niet gevonden op '$mailLower', poging met '$SourceGroupMail'..."
    $group = Get-MgGroup -Filter "mail eq '$SourceGroupMail'" -ErrorAction SilentlyContinue
}

if (-not $group) {
    Write-Warn "Niet gevonden op mail, fallback naar displayName '$SourceGroupDisplayName'..."
    $group = Get-MgGroup -Filter "displayName eq '$SourceGroupDisplayName'" -ErrorAction SilentlyContinue
}

if (-not $group) {
    Write-Warn "Niet gevonden via filter, laatste poging via Search..."
    $group = Get-MgGroup -Search "`"displayName:$SourceGroupDisplayName`"" `
        -ConsistencyLevel eventual -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -eq $SourceGroupDisplayName -or $_.Mail -like "*holiday*" } |
        Select-Object -First 1
}

if (-not $group) {
    Write-Fail "Groep niet gevonden na 4 pogingen."
    Write-Host "    Get-MgGroup -Search `"`"displayName:Holidays`"`" -ConsistencyLevel eventual | Select DisplayName,Mail,Id" -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph | Out-Null
    exit 1
}

Write-OK "Groep gevonden: '$($group.DisplayName)' | Mail: $($group.Mail) | ID: $($group.Id)"

Write-Step "Afspraken ophalen ($DaysBack dagen terug t/m $DaysForward dagen vooruit)"

$startDate = (Get-Date).AddDays(-$DaysBack).ToString("yyyy-MM-ddT00:00:00")
$endDate   = (Get-Date).AddDays($DaysForward).ToString("yyyy-MM-ddT23:59:59")

try {
    $events = Get-MgGroupCalendarEvent `
        -GroupId     $group.Id `
        -Filter      "start/dateTime ge '$startDate' and start/dateTime le '$endDate'" `
        -All `
        -ErrorAction Stop
    Write-OK "$($events.Count) afspraken opgehaald"
} catch {
    Write-Error "Kon afspraken niet ophalen: $_"
    exit 1
}

# Terugschakelen naar app auth voor schrijven naar Room Mailbox
Write-Step "Terugschakelen naar application permissions voor schrijven"
Disconnect-MgGraph | Out-Null

try {
    Connect-GraphAppAuth -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    Write-OK "Graph herverbonden (app auth) | AuthType: $((Get-MgContext).AuthType)"
} catch {
    Write-Error "App auth herverbinding mislukt: $_"
    exit 1
}

#endregion

#region Stap 10: Afspraken kopieren

Write-Step "Afspraken kopieren naar Room Mailbox kalender"

$successCount = 0
$failCount    = 0
$skippedCount = 0

foreach ($event in $events) {

    if ($event.IsCancelled) {
        $skippedCount++
        continue
    }

    if ($PSCmdlet.ShouldProcess($event.Subject, "Event kopieren naar $RoomEmail")) {
        try {
            $tz = if ($event.Start.TimeZone) { $event.Start.TimeZone } else { "Europe/Brussels" }

            $params = @{
                Subject  = $event.Subject
                IsAllDay = $event.IsAllDay
                ShowAs   = "oof"
                Start    = @{ DateTime = $event.Start.DateTime; TimeZone = $tz }
                End      = @{ DateTime = $event.End.DateTime;   TimeZone = $tz }
                Body     = @{
                    ContentType = "text"
                    Content     = "Gemigreerd vanuit Holidays M365 groep. Originele organisator: $($event.Organizer.EmailAddress.Address)"
                }
            }

            New-MgUserEvent -UserId $RoomEmail -BodyParameter $params | Out-Null
            Write-Host "    [+] $($event.Subject) | $($event.Start.DateTime)" -ForegroundColor DarkGreen
            $successCount++
        } catch {
            Write-Warn "Niet gekopieerd: '$($event.Subject)' - $_"
            $failCount++
        }
    }
}

Write-OK "Kopieren klaar: $successCount OK | $skippedCount overgeslagen | $failCount mislukt"

#endregion

#region Stap 11: M365 Group verwijderen (optioneel)

if ($DeleteSourceGroup) {
    Write-Step "M365 Group verwijderen: $SourceGroupDisplayName"
    if ($PSCmdlet.ShouldProcess($group.Id, "M365 Group verwijderen")) {
        try {
            Remove-MgGroup -GroupId $group.Id -Confirm:$false
            Write-OK "M365 Group verwijderd"
        } catch {
            Write-Warn "Kon groep niet verwijderen: $_"
        }
    }
} else {
    Write-Warn "M365 Group is NIET verwijderd (DeleteSourceGroup = false)"
    Write-Host "    Aanbevolen: verwijder de Dynamic Membership rule zodat niemand meer notificaties krijgt" -ForegroundColor Yellow
    Write-Host "    Of verwijder de groep volledig: Remove-MgGroup -GroupId $($group.Id)" -ForegroundColor DarkYellow
}

#endregion

#region Stap 12: Samenvatting

$line = "=" * 65
Write-Host "`n$line" -ForegroundColor Cyan
Write-Host " SAMENVATTING  -  Ticket 0298048  -  Onco3R Holidays migratie" -ForegroundColor Cyan
Write-Host $line -ForegroundColor Cyan
Write-Host "  Platform            : $(if ($runOnWindows) { 'Windows' } elseif ($runOnMacOS) { 'macOS' } else { 'Linux' })"
Write-Host "  App Registration    : $ClientId"
Write-Host "  Nieuwe Room Mailbox : $RoomEmail"
Write-Host "  Afspraken gekopieerd: $successCount"
Write-Host "  Overgeslagen        : $skippedCount (geannuleerd)"
Write-Host "  Mislukt             : $failCount"
Write-Host ""
Write-Host " HOE VERLOF BOEKEN (uitleg voor eindgebruikers):" -ForegroundColor Yellow
Write-Host "  1. Maak een afspraak in Outlook (All day, status = Out of office)"
Write-Host "  2. Voeg '$RoomEmail' toe als attendee (zoals een vergaderzaal)"
Write-Host "  3. Opslaan - boeking wordt automatisch goedgekeurd"
Write-Host "  4. Afspraak verschijnt op de gedeelde Holidays Calendar"
Write-Host ""
Write-Host " KALENDER TOEVOEGEN IN OUTLOOK (eenmalig per gebruiker):" -ForegroundColor Yellow
Write-Host "  1. Calendar > Add calendar > Add from directory"
Write-Host "  2. Zoek: '$RoomDisplayName' of '$RoomEmail'"
Write-Host "  3. Toevoegen - daarna altijd zichtbaar onder People's calendars"
Write-Host $line -ForegroundColor Cyan

Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph | Out-Null
Write-OK "Script klaar. Verbindingen verbroken."

#endregion