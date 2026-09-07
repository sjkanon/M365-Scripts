<#
.SYNOPSIS
    Houdt twee elkaar uitsluitende statische groepen bij op basis van of een
    gebruiker een geaccepteerde MFA-methode heeft geregistreerd:
      - Rollout-groep    : moet nog registreren (in scope van de CA-eis)
      - Registered-groep : heeft al geregistreerd (compliant)

.DESCRIPTION
    Voor elke gebruiker die momenteel in Rollout OF Registered zit:
      - Heeft geaccepteerde methode EN zit in Rollout
            -> verwijderen uit Rollout, toevoegen aan Registered (graduatie)
      - Heeft GEEN geaccepteerde methode meer EN zit in Registered
            -> verwijderen uit Registered, toevoegen aan Rollout (terugval,
               methode verwijderd/verlopen -> moet opnieuw registreren)
      - Andere combinaties: geen actie, status is al correct.

    Een gebruiker zit dus nooit in beide groepen tegelijk.

    Wat als "geaccepteerd" telt bepaal je met -AcceptedMethod:
      - AuthenticatorPasskey (default): alleen een passkey in Microsoft
        Authenticator. Dat is een fido2AuthenticationMethod waarvan de AAGUID
        in -AllowedAaGuids staat. Een YubiKey, Windows Hello for Business of
        CBA telt dus NIET; die laten de gebruiker in Rollout staan.
      - AnyPhishingResistant: elke phishing-resistant methode telt.

    Let op: in de Entra-portal heet zo'n methode gewoon "Passkey" met als
    detail bv. "MS Authenticator iOS". In Graph is dat geen aparte
    methodesoort maar een fido2-methode; alleen de AAGUID verraadt dat het
    om Authenticator gaat.

.NOTES
    Benodigde Graph applicatiepermissies (Managed Identity / App Registration):
        User.Read.All
        UserAuthenticationMethod.Read.All
        GroupMember.ReadWrite.All  (of breder: Group.ReadWrite.All)

    Bedoeld als Azure Automation Runbook (system-assigned managed identity),
    periodiek te draaien (bv. elke 1-4 uur).

.PARAMETER RolloutGroupId
    Object Id van de statische Rollout-groep (moet nog registreren).

.PARAMETER RegisteredGroupId
    Object Id van de statische Registered-groep (al gecompliant).

.PARAMETER AcceptedMethod
    Bepaalt wat als "geregistreerd" telt:
      AuthenticatorPasskey  (default) - alleen een passkey in Microsoft
                                        Authenticator (AAGUID-filter).
      AnyPhishingResistant            - elke phishing-resistant methode
                                        (FIDO2-sleutels, WHfB, Platform SSO, CBA).

.PARAMETER AllowedAaGuids
    AAGUIDs die als "passkey in Authenticator" tellen bij
    -AcceptedMethod AuthenticatorPasskey. Default: de iOS- en Android-AAGUID
    van Microsoft Authenticator. Vul aan als je bv. ook een specifiek
    security-key-model wilt accepteren. Wordt genegeerd bij AnyPhishingResistant.

.PARAMETER Interactive
    Log interactief in (browser-prompt) i.p.v. managed identity. Gebruik dit
    als je het script lokaal op je eigen machine draait. Zonder deze switch
    en zonder -UseAppRegistration wordt managed identity geprobeerd, wat
    buiten Azure altijd faalt.

    Het inloggen gebeurt met -ContextScope CurrentUser, zodat de tokencache op
    schijf staat en de SDK het token stil kan vernieuwen. Zonder dat leeft de
    cache alleen in het huidige proces en opent de SDK bij een verlopen token
    alsnog een browser - midden in de run, per gebruiker.

    Voor een run over veel gebruikers of onbewaakt draaien is
    -UseAppRegistration met certificaat betrouwbaarder: app-only tokens kunnen
    nooit om interactie vragen.

.PARAMETER UseTemporaryApp
    Maakt zelf een wegwerp-app-registratie aan, draait de hele run app-only, en
    verwijdert die app daarna weer. Zelfde patroon als de SharePoint- en
    Intune-scripts in deze repo.

    Dit is de betrouwbaarste modus voor een run over veel gebruikers: een
    app-only token wordt bij elke call opnieuw gemunt met het client secret en
    kan dus nooit een browser-prompt opleveren. Je logt één keer interactief in
    om de app te mogen aanmaken; daarna is de delegated sessie niet meer nodig.

    De app krijgt als applicatierechten User.Read.All,
    UserAuthenticationMethod.Read.All en GroupMember.ReadWrite.All, plus
    Application.ReadWrite.OwnedBy zodat hij zichzelf kan opruimen. Hij wordt
    eigenaar van zichzelf en is dus verder nergens eigenaar van.

    Vereist Application Administrator of Global Administrator: het script maakt
    een app aan én kent er app-rollen aan toe.

    LET OP: de app wordt ook bij -WhatIf echt aangemaakt en weer verwijderd -
    zonder token valt er niets te simuleren. De groepswijzigingen zelf blijven
    met -WhatIf gewoon achterwege.

.PARAMETER UseExistingSession
    Draai op de Connect-MgGraph-sessie die je zelf al hebt opgezet. Het script
    roept dan zelf nooit Connect-MgGraph of Disconnect-MgGraph aan, ook niet als
    het token halverwege sneuvelt. In dat geval stopt het met de lijst van wat
    er nog openstond, zodat je zelf opnieuw kunt inloggen en herstarten.

    LET OP: is het token dood, dan kan de Graph SDK zelf nog één browser-prompt
    openen bij de eerstvolgende call - dat zit in de credential van jouw sessie
    en daar kan het script niet tussen komen. Klik je die weg, dan stopt het
    script met een duidelijke melding i.p.v. de prompt per gebruiker te herhalen.

    Zonder deze switch gebeurt hetzelfde impliciet als je geen modus kiest en
    er toevallig een delegated sessie is; met de switch is het een bewuste
    keuze en klaagt het script meteen als die sessie er niet is.

    Log zelf in met de scopes die het script nodig heeft, en met
    -ContextScope CurrentUser zodat de tokencache op schijf staat en stil
    vernieuwd kan worden:

      Connect-MgGraph -Scopes User.Read.All,UserAuthenticationMethod.Read.All,GroupMember.ReadWrite.All -ContextScope CurrentUser

.PARAMETER ForceLogin
    Negeer een bestaande Graph-sessie en log opnieuw in. Zonder deze switch
    hergebruikt -Interactive een sessie die de benodigde scopes al heeft, zodat
    je niet bij elke run een browser-prompt krijgt. Gebruik dit als je van
    account wilt wisselen of opnieuw wilt consenten.

.PARAMETER PendingCsvPath
    Optioneel pad naar een CSV met de gebruikers die nog actie vragen: wie in
    de Rollout-groep zit zonder geaccepteerde methode (NogTeRegistreren) en
    wie het script niet heeft kunnen checken (NietGecheckt). Handig om
    achteraan te bellen zonder de console-output door te spitten. Wordt bij
    -RunContinuously elke pass overschreven.

.PARAMETER RunContinuously
    Als aanwezig: het script blijft draaien en voert elke $IntervalMinutes
    minuten een nieuwe sync-pass uit (interne Start-Sleep), i.p.v. eenmalig te
    draaien en te stoppen. Handig op een Hybrid Runbook Worker, VM of server.

    LET OP: in een Azure Automation CLOUD job (geen Hybrid Worker) geldt een
    fair-share limiet van 3 uur per job -> een langere interne lus wordt dan
    halverwege afgebroken. Gebruik in dat geval liever GEEN -RunContinuously,
    en laat de scheduler zelf elke run starten.

.PARAMETER IntervalMinutes
    Aantal minuten pauze tussen elke sync-pass bij -RunContinuously.
    Default: 240 (= 4 uur). Voor elk half uur: -IntervalMinutes 30.

.PARAMETER IntervalHours
    Verouderd; alleen voor compatibiliteit. Wordt omgerekend naar
    -IntervalMinutes als die niet expliciet is opgegeven.

.EXAMPLE
    # Eenmalig draaien (aanbevolen i.c.m. een externe scheduler / Automation Schedule)
    ./Phising-rollout.ps1 `
        -RolloutGroupId "11111111-1111-1111-1111-111111111111" `
        -RegisteredGroupId "22222222-2222-2222-2222-222222222222"

.EXAMPLE
    # Tijdelijke app-registratie: één keer inloggen, daarna app-only, app wordt
    # na afloop weer verwijderd. Aanbevolen voor grotere groepen.
    ./Phising-rollout.ps1 `
        -RolloutGroupId "11111111-1111-1111-1111-111111111111" `
        -RegisteredGroupId "22222222-2222-2222-2222-222222222222" `
        -UseTemporaryApp

.EXAMPLE
    # Draaien op je eigen Graph-sessie; het script logt zelf nooit in
    Connect-MgGraph -Scopes User.Read.All,UserAuthenticationMethod.Read.All,GroupMember.ReadWrite.All -ContextScope CurrentUser

    ./Phising-rollout.ps1 `
        -RolloutGroupId "11111111-1111-1111-1111-111111111111" `
        -RegisteredGroupId "22222222-2222-2222-2222-222222222222" `
        -UseExistingSession

.EXAMPLE
    # Lokaal testen met interactieve login, zonder wijzigingen door te voeren
    ./Phising-rollout.ps1 `
        -RolloutGroupId "11111111-1111-1111-1111-111111111111" `
        -RegisteredGroupId "22222222-2222-2222-2222-222222222222" `
        -Interactive -WhatIf

.EXAMPLE
    # Elke phishing-resistant methode laten meetellen (YubiKey, WHfB, CBA, ...)
    ./Phising-rollout.ps1 `
        -RolloutGroupId "11111111-1111-1111-1111-111111111111" `
        -RegisteredGroupId "22222222-2222-2222-2222-222222222222" `
        -AcceptedMethod AnyPhishingResistant

.EXAMPLE
    # Continu draaien, elke 30 minuten (Hybrid Worker / VM / server)
    ./Phising-rollout.ps1 `
        -RolloutGroupId "11111111-1111-1111-1111-111111111111" `
        -RegisteredGroupId "22222222-2222-2222-2222-222222222222" `
        -RunContinuously -IntervalMinutes 30
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$RolloutGroupId,

    [Parameter(Mandatory = $true)]
    [string]$RegisteredGroupId,

    [ValidateSet('AuthenticatorPasskey', 'AnyPhishingResistant')]
    [string]$AcceptedMethod = 'AuthenticatorPasskey',

    # AAGUIDs van Microsoft Authenticator; komt overeen met het passkey-profiel
    # in Entra (Authentication methods > Passkey (FIDO2) > Target specific AAGUIDs).
    [string[]]$AllowedAaGuids = @(
        '90a3ccdf-635c-4729-a248-9b709135078f',   # Microsoft Authenticator for iOS
        'de1e552d-db1d-4423-a619-566b625cdc84'    # Microsoft Authenticator for Android
    ),

    [switch]$UseAppRegistration,
    [switch]$UseTemporaryApp,
    [switch]$Interactive,
    [switch]$UseExistingSession,
    [switch]$ForceLogin,
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientCertificateThumbprint,

    [string]$PendingCsvPath,

    [switch]$RunContinuously,
    [int]$IntervalMinutes = 240,
    [double]$IntervalHours
)

# -IntervalHours is verouderd; alleen gebruiken als -IntervalMinutes niet is opgegeven
if ($PSBoundParameters.ContainsKey('IntervalHours') -and -not $PSBoundParameters.ContainsKey('IntervalMinutes')) {
    $IntervalMinutes = [int]($IntervalHours * 60)
}
if ($IntervalMinutes -lt 1) { throw "IntervalMinutes moet minimaal 1 zijn." }

# ---------------------------------------------------------------------------
# 0. Modules & connectie
# ---------------------------------------------------------------------------
if ($UseTemporaryApp -and ($UseAppRegistration -or $UseExistingSession)) {
    throw "-UseTemporaryApp gaat niet samen met -UseAppRegistration of -UseExistingSession: kies één authenticatiemodus."
}

$RequiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Groups", "Microsoft.Graph.Identity.SignIns")
if ($UseTemporaryApp) { $RequiredModules += "Microsoft.Graph.Applications" }
foreach ($m in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Verbose "Installeren van module $m..."
        Install-Module -Name $m -Force -Scope CurrentUser -AllowClobber
    }
    Import-Module $m -ErrorAction Stop
}

$InteractiveScopes = @(
    "User.Read.All",
    "UserAuthenticationMethod.Read.All",
    "GroupMember.ReadWrite.All"
)

function Test-GraphSessionUsable {
    <#
        Ruwe check of een bestaande sessie genoeg rechten heeft. Bewust ruim
        matchen: Entra geeft de daadwerkelijk geconsente scopes terug en die
        kunnen breder of anders benoemd zijn dan wat we vroegen (bv.
        Group.ReadWrite.All i.p.v. GroupMember.ReadWrite.All). Exact vergelijken
        zou dus een prima sessie afkeuren.

        Doel is alleen het onderscheid met een kale 'Connect-MgGraph' (die enkel
        User.Read krijgt); die moet wél opnieuw inloggen, anders faalt de eerste
        groepsactie pas halverwege.
    #>
    param($Context)

    if (-not $Context) { return $false }

    $scopes = @($Context.Scopes)
    $canWriteGroups = @($scopes | Where-Object { $_ -match '^(GroupMember|Group|Directory)\.ReadWrite' }).Count -gt 0
    $canReadMethods = @($scopes | Where-Object { $_ -match '^UserAuthenticationMethod\.Read' }).Count -gt 0

    return ($canWriteGroups -and $canReadMethods)
}

# Meeliften op de Connect-MgGraph die de gebruiker zelf heeft gedaan. Met
# -UseExistingSession is dat een expliciete keuze; zonder enige modus-switch is
# managed identity de bedoeling, maar die faalt buiten Azure altijd - dan is een
# bestaande delegated sessie duidelijk wat je bedoelde, en is hergebruiken beter
# dan afbreken op een IMDS-timeout van 169.254.169.254.
#
# In deze modus roept het script zelf geen Connect-MgGraph of Disconnect-MgGraph
# aan. Is het token op, dan stopt het en log je zelf opnieuw in. (De SDK kan bij
# een dood token nog wel zelf een prompt openen; dat zit in de credential van de
# sessie en daar komt het script niet tussen.)
$script:ReuseCallerSession = $false
if ($UseExistingSession -or -not ($UseAppRegistration -or $UseTemporaryApp -or $Interactive)) {
    $existingContext = Get-MgContext

    if ($existingContext -and $existingContext.AuthType -eq 'Delegated') {
        $script:ReuseCallerSession = $true

        if ($UseExistingSession) {
            Write-Output "Hergebruikt je eigen Graph-sessie: $($existingContext.Account)"
        }
        else {
            Write-Warning "Bestaande Graph-sessie gevonden ($($existingContext.Account)); die wordt hergebruikt i.p.v. managed identity."
        }

        if (-not (Test-GraphSessionUsable -Context $existingContext)) {
            Write-Warning "Die sessie mist waarschijnlijk scopes: $($existingContext.Scopes -join ', '). Verwacht: $($InteractiveScopes -join ', '). Een kale 'Connect-MgGraph' krijgt alleen User.Read en loopt dus vast op de groepen."
        }
    }
    elseif ($UseExistingSession) {
        # Expliciet gevraagd om hergebruik, maar er is niets om te hergebruiken:
        # doorgaan zou stilletjes managed identity proberen.
        $found = if ($existingContext) { "authenticatietype $($existingContext.AuthType)" } else { "geen Graph-sessie" }
        throw "-UseExistingSession vraagt om een bestaande delegated sessie, maar ik vond $found. Log eerst in met: Connect-MgGraph -Scopes $($InteractiveScopes -join ',') -ContextScope CurrentUser"
    }
}

function Test-GraphTokenAlive {
    <#
        Get-MgContext bewijst NIETS over het token: dat object blijft ook staan
        als het access token verlopen is of Conditional Access de refresh
        weigert. Daarom een echte, goedkope call als bewijs.

        Zonder deze check denkt het script "sessie is prima", slaat het inloggen
        over, en ontdekt de SDK het pas bij de eerste gebruiker - waarna hij per
        gebruiker een browser-prompt opent.
    #>
    [CmdletBinding()]
    param()

    $context = Get-MgContext
    if (-not $context) { return $false }

    # App-only en managed identity halen hun token per call op en kunnen nooit
    # om interactie vragen; daar valt niets te proben.
    if ($context.AuthType -ne 'Delegated') { return $true }

    try {
        $null = Invoke-MgGraphRequest -Method GET -Uri 'v1.0/me?$select=id' -ErrorAction Stop
        return $true
    }
    catch {
        Write-Verbose "Sessie-probe mislukt: $($_.Exception.Message)"
        return $false
    }
}

function Test-IsGraphAuthError {
    <#
        Onderscheid tussen "deze gebruiker gaat mis" (overslaan is prima) en
        "de sessie is stuk" (doorgaan levert alleen maar meer browser-prompts
        op). Op de tekst matchen is lelijk, maar de SDK verpakt MSAL-fouten in
        een generieke AuthenticationFailedException zonder bruikbare code.
    #>
    param($ErrorRecord)

    $message = "$($ErrorRecord.Exception.Message)"
    return $message -match 'InteractiveBrowserCredential|DeviceCodeCredential|ClientCertificateCredential|ManagedIdentityCredential|AuthenticationFailed|MsalUiRequired|MsalServiceException|AADSTS|InvalidAuthenticationToken|Access token has expired|CredentialUnavailable'
}

$GraphResourceAppId = '00000003-0000-0000-c000-000000000000'

# App-only rollen voor de tijdelijke app. Application.ReadWrite.OwnedBy staat er
# bewust bij: daarmee kan de app zichzelf aan het eind verwijderen, zonder de
# beheerder nog een keer door een login te sturen. Hij is alleen eigenaar van
# zichzelf, dus verder reikt die rol niet.
$TemporaryAppRoles = @(
    'User.Read.All',
    'UserAuthenticationMethod.Read.All',
    'GroupMember.ReadWrite.All',
    'Application.ReadWrite.OwnedBy'
)

$script:TempApp        = $null
$script:TenantIdForApp = $null

function New-TemporaryGraphApp {
    <#
        Maakt een wegwerp-app-registratie met app-only rechten, zoals de
        SharePoint- en Intune-scripts in deze repo dat ook doen.

        Waarom: een app-only token wordt bij elke call opnieuw opgehaald met het
        client secret en kan dus nooit om een browser vragen. Een delegated
        sessie sneuvelt halverwege een lange run (verlopen refresh token,
        Conditional Access) en opent dan per gebruiker een prompt.

        Vereist Application Administrator of Global Administrator: het script
        maakt een app aan én kent er app-rollen aan toe.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$AppName)

    Write-Output "Tijdelijke app-registratie aanmaken: $AppName"
    $app = New-MgApplication -DisplayName $AppName -SignInAudience 'AzureADMyOrg' -ErrorAction Stop
    $script:TempApp = [pscustomobject]@{
        ObjectId    = $app.Id
        AppId       = $app.AppId
        DisplayName = $AppName
        Credential  = $null
    }

    $sp = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop

    # Eigenaar van zichzelf maken, anders doet Application.ReadWrite.OwnedBy
    # niets en kan de app zichzelf straks niet opruimen. Een net aangemaakte
    # service principal is soms nog niet vindbaar: even doorproberen.
    $ownerSet = $false
    for ($attempt = 1; $attempt -le 4; $attempt++) {
        try {
            New-MgApplicationOwnerByRef -ApplicationId $app.Id `
                -OdataId "https://graph.microsoft.com/v1.0/directoryObjects/$($sp.Id)" -ErrorAction Stop
            $ownerSet = $true
            break
        }
        catch {
            Write-Verbose "Eigenaar zetten nog niet gelukt (poging $attempt/4): $($_.Exception.Message)"
            Start-Sleep -Seconds 3
        }
    }
    if (-not $ownerSet) {
        throw "Kon de tijdelijke app geen eigenaar van zichzelf maken; zonder dat kan hij zichzelf niet opruimen."
    }

    $graphSp = Get-MgServicePrincipal -Filter "appId eq '$GraphResourceAppId'" -ErrorAction Stop
    foreach ($roleValue in $TemporaryAppRoles) {
        $appRole = $graphSp.AppRoles | Where-Object { $_.Value -eq $roleValue -and $_.AllowedMemberTypes -contains 'Application' } | Select-Object -First 1
        if (-not $appRole) {
            throw "App-rol '$roleValue' niet gevonden op de Microsoft Graph service principal."
        }

        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $sp.Id `
            -PrincipalId        $sp.Id `
            -ResourceId         $graphSp.Id `
            -AppRoleId          $appRole.Id `
            -ErrorAction Stop | Out-Null
        Write-Output "  [OK] $roleValue toegekend"
    }

    # Secret-levensduur volgt de run: een eenmalige pass is klaar binnen een dag,
    # maar bij -RunContinuously moet er ook over weken nog een token gemunt
    # kunnen worden. Een verlopen secret is niet te vernieuwen door de app zelf.
    $secretLifetime = if ($RunContinuously) { (Get-Date).AddDays(30) } else { (Get-Date).AddDays(1) }
    $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
        displayName = 'temp'
        endDateTime = $secretLifetime
    } -ErrorAction Stop

    $secureSecret = ConvertTo-SecureString $secret.SecretText -AsPlainText -Force
    $script:TempApp.Credential = [System.Management.Automation.PSCredential]::new($app.AppId, $secureSecret)

    return $script:TempApp
}

function Remove-TemporaryGraphApp {
    <#
        Ruimt de wegwerp-app op. Lukt dat niet, dan is de exacte opruimopdracht
        belangrijker dan de foutmelding: blijft hij staan, dan heb je een app met
        UserAuthenticationMethod.Read.All in je tenant zonder dat iemand weet
        waar hij vandaan komt.
    #>
    [CmdletBinding()]
    param()

    if (-not $script:TempApp) { return }

    $app = $script:TempApp
    $script:TempApp = $null

    try {
        Remove-MgApplication -ApplicationId $app.ObjectId -ErrorAction Stop
        Write-Output "Tijdelijke app verwijderd: $($app.DisplayName)"
    }
    catch {
        Write-Warning "Kon de tijdelijke app niet verwijderen: $($_.Exception.Message)"
        Write-Warning "Ruim hem handmatig op - Entra ID > App registrations > '$($app.DisplayName)' (AppId $($app.AppId)), of:"
        Write-Warning "  Connect-MgGraph -Scopes Application.ReadWrite.All; Remove-MgApplication -ApplicationId $($app.ObjectId)"
    }
}

function Connect-TemporaryAppSession {
    <#
        Maakt (indien nodig) de tijdelijke app en verbindt daarna app-only.
        Bestaat de app al, dan alleen opnieuw verbinden met hetzelfde secret.
    #>
    [CmdletBinding()]
    param()

    if (-not $script:TempApp) {
        # De app aanmaken vraagt om een beheerderssessie; daarna is de delegated
        # sessie niet meer nodig en nemen we het app-only token over.
        Write-Output "Interactief inloggen om de tijdelijke app aan te maken (Application Administrator of Global Administrator vereist)..."
        $adminParams = @{
            Scopes       = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
            ContextScope = 'CurrentUser'
            NoWelcome    = $true
            ErrorAction  = 'Stop'
        }
        if ($TenantId) { $adminParams['TenantId'] = $TenantId }
        Connect-MgGraph @adminParams

        $adminContext  = Get-MgContext
        $script:TenantIdForApp = if ($TenantId) { $TenantId } else { $adminContext.TenantId }

        New-TemporaryGraphApp -AppName "PhishingRollout-Temp-$(Get-Date -Format 'yyyyMMddHHmmss')" | Out-Null
    }

    # Een vers aangemaakte app en secret zijn niet meteen overal bekend; de
    # eerste token-aanvraag faalt dan met "unauthorized_client" of "invalid
    # client". Even doorproberen i.p.v. de run afbreken.
    $connected = $false
    $lastError = $null
    for ($attempt = 1; $attempt -le 6; $attempt++) {
        try {
            Connect-MgGraph -TenantId $script:TenantIdForApp -ClientSecretCredential $script:TempApp.Credential -NoWelcome -ErrorAction Stop
            $connected = $true
            break
        }
        catch {
            $lastError = $_
            Write-Verbose "App-only verbinden nog niet gelukt (poging $attempt/6): $($_.Exception.Message)"
            Start-Sleep -Seconds 5
        }
    }
    if (-not $connected) {
        throw "App-only verbinden met de tijdelijke app mislukt: $($lastError.Exception.Message)"
    }

    # App-rollen zijn niet meteen actief; de eerste calls geven dan 403. Wachten
    # tot een echte lees-actie lukt, anders lijkt straks elke gebruiker te falen.
    Write-Output "Wachten tot de rechten actief zijn..."
    $lastError = $null
    for ($attempt = 1; $attempt -le 12; $attempt++) {
        try {
            $null = Get-MgGroupMember -GroupId $RolloutGroupId -Top 1 -ErrorAction Stop
            Write-Output "Verbonden als tijdelijke app: $($script:TempApp.DisplayName)"
            return
        }
        catch {
            $lastError = $_
            Write-Verbose "Rechten nog niet actief (poging $attempt/12): $($_.Exception.Message)"
            Start-Sleep -Seconds 5
        }
    }

    throw "De tijdelijke app kreeg binnen een minuut geen toegang tot de Rollout-groep. Laatste fout: $($lastError.Exception.Message)"
}

function Connect-GraphSession {
    <#
        Verbindt volgens de gekozen modus en gooit een terminating error als dat
        niet lukt. Zonder deze check draait het script vrolijk door en rapporteert
        het 0 leden, wat lijkt op "niks te doen" terwijl er niets is gecontroleerd.

        -Force slaat het hergebruik over en logt hoe dan ook opnieuw in; dat is
        de uitweg als het token halverwege een pass sneuvelt.
    #>
    [CmdletBinding()]
    param(
        [switch]$Silent,
        [switch]$Force
    )

    # Ook bij -Force: in deze modus is de sessie van de gebruiker, daar logt het
    # script niet overheen. Wel een echte probe, want alleen "er is een context"
    # is geen bewijs dat het token nog werkt.
    if ($script:ReuseCallerSession) {
        if (Test-GraphTokenAlive) { return }
        throw "Je eigen Graph-sessie is verlopen of geweigerd. Het script logt bewust niet zelf in. Herstel met: Connect-MgGraph -Scopes $($InteractiveScopes -join ',') -ContextScope CurrentUser"
    }

    try {
        if ($UseTemporaryApp) {
            # App-only tokens verlopen wel, maar worden stil vernieuwd met het
            # secret; opnieuw verbinden is alleen nodig bij -Force.
            if ($script:TempApp -and -not $Force -and (Get-MgContext)) { return }
            Connect-TemporaryAppSession
        }
        elseif ($UseAppRegistration) {
            if (-not ($TenantId -and $ClientId -and $ClientCertificateThumbprint)) {
                throw "-UseAppRegistration vereist -TenantId, -ClientId en -ClientCertificateThumbprint."
            }
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $ClientCertificateThumbprint -NoWelcome -ErrorAction Stop
        }
        elseif ($Interactive) {
            # Een bruikbare bestaande sessie hergebruiken, ook bij de eerste
            # aanroep. De SDK vernieuwt het access token zelf via de refresh
            # token in de cache; opnieuw Connect-MgGraph draaien levert alleen
            # een extra browser-prompt op.
            $context = Get-MgContext
            if ($context -and -not $ForceLogin -and -not $Force) {
                if ((Test-GraphSessionUsable -Context $context) -and (Test-GraphTokenAlive)) { return }

                # In de continue lus nooit alsnog een browser openen: dat blokkeert
                # een onbewaakte run tot iemand toevallig langsloopt.
                if ($Silent) { return }

                Write-Warning "Bestaande sessie ($($context.Account)) is niet bruikbaar (scopes: $($context.Scopes -join ', ')). Opnieuw inloggen."
            }

            # -ContextScope CurrentUser: zet de tokencache op schijf i.p.v.
            # alleen in dit proces. Daardoor kan de SDK het token stil
            # vernieuwen en hoef je niet bij elke run (of elk verlopen token)
            # opnieuw door de browser.
            $connectParams = @{
                Scopes       = $InteractiveScopes
                ContextScope = 'CurrentUser'
                NoWelcome    = $true
                ErrorAction  = 'Stop'
            }
            if ($TenantId) { $connectParams['TenantId'] = $TenantId }
            Connect-MgGraph @connectParams
        }
        else {
            # Azure Automation Runbook / Azure VM met (system-assigned) managed identity
            if ($ClientId) {
                Connect-MgGraph -Identity -ClientId $ClientId -NoWelcome -ErrorAction Stop
            }
            else {
                Connect-MgGraph -Identity -NoWelcome -ErrorAction Stop
            }
        }
    }
    catch {
        if (-not ($UseAppRegistration -or $Interactive)) {
            Write-Warning "Managed identity is niet beschikbaar. Draai je dit lokaal? Gebruik dan -Interactive, of -UseAppRegistration met certificaat."
        }
        throw "Verbinden met Microsoft Graph mislukt: $($_.Exception.Message)"
    }

    if (-not (Get-MgContext)) {
        throw "Verbinden met Microsoft Graph mislukt: geen actieve Graph-context na Connect-MgGraph."
    }
}

# Eén keer per pass mag het script zichzelf herstellen; daarna is doorgaan
# zinloos en zou elke gebruiker een nieuwe prompt opleveren.
$script:ReconnectAttempted = $false

function Get-AuthMethodsResilient {
    <#
        Haalt de authenticatiemethoden op. Sneuvelt het token halverwege een
        pass (Conditional Access, verlopen refresh token), dan verbindt dit
        eenmalig opnieuw en probeert het die gebruiker nog een keer, i.p.v. de
        SDK per gebruiker een browser te laten openen.
    #>
    param([string]$UserId)

    try {
        return @(Get-MgUserAuthenticationMethod -UserId $UserId -ErrorAction Stop)
    }
    catch {
        $original = $_

        if (-not (Test-IsGraphAuthError -ErrorRecord $original)) { throw }
        if ($script:ReconnectAttempted) { throw }

        # Draaien we op de sessie van de gebruiker, dan is opnieuw inloggen niet
        # aan ons. De aanroeper breekt af met de lijst van wat nog openstond.
        if ($script:ReuseCallerSession) { throw }

        $script:ReconnectAttempted = $true
        Write-Warning "Graph-sessie sneuvelde tijdens de run; eenmalig opnieuw verbinden..."

        try {
            Connect-GraphSession -Force
        }
        catch {
            # De oorspronkelijke auth-fout doorgooien, niet deze: de aanroeper
            # herkent daaraan dat de sessie stuk is en stopt meteen, i.p.v. de
            # volgende gebruiker te proberen (en dus een nieuwe prompt te openen).
            Write-Warning "Opnieuw verbinden mislukt: $($_.Exception.Message)"
            throw $original
        }

        Write-Output "Opnieuw verbonden als $((Get-MgContext).Account); hervatten bij $(Get-UserLabel -UserId $UserId)."
        return @(Get-MgUserAuthenticationMethod -UserId $UserId -ErrorAction Stop)
    }
}

Connect-GraphSession

Write-Output "Verbonden met tenant: $((Get-MgContext).TenantId)"

# ---------------------------------------------------------------------------
# 1. Classificatie
# ---------------------------------------------------------------------------
$Fido2Type = "#microsoft.graph.fido2AuthenticationMethod"

$PhishingResistantTypes = @(
    $Fido2Type,                                                   # FIDO2-sleutels en passkeys (incl. Authenticator passkey)
    "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod",
    "#microsoft.graph.platformCredentialAuthenticationMethod",    # macOS Platform SSO
    "#microsoft.graph.x509CertificateAuthenticationMethod"        # CBA
)

# Normaliseren zodat een AAGUID met accolades of hoofdletters ook matcht.
$NormalizedAllowedAaGuids = @(
    $AllowedAaGuids | Where-Object { $_ } | ForEach-Object { $_.Trim().Trim('{', '}').ToLowerInvariant() }
)
if ($AcceptedMethod -eq 'AuthenticatorPasskey' -and $NormalizedAllowedAaGuids.Count -eq 0) {
    throw "-AcceptedMethod AuthenticatorPasskey vereist minimaal één AAGUID in -AllowedAaGuids."
}

# Id -> UPN/weergavenaam, gevuld uit de groepsleden. Graph levert bij
# /groups/{id}/members het volledige user-object mee, dus dit kost geen extra
# call; zonder deze map staan er alleen GUID's in de logging en moet je zelf
# gaan opzoeken over wie het gaat.
$script:UserLabels = @{}

function Get-UserLabel {
    param([string]$UserId)

    if ($script:UserLabels.ContainsKey($UserId)) {
        return $script:UserLabels[$UserId]
    }
    return $UserId
}

function Register-UserLabel {
    <#
        Haalt UPN (of anders weergavenaam/mail) uit het member-object. Valt
        terug op de GUID zodat de logging nooit leeg is.
    #>
    param($Member)

    if (-not $Member -or -not $Member.Id) { return }
    if ($script:UserLabels.ContainsKey($Member.Id)) { return }

    $label = $null
    $props = $Member.AdditionalProperties
    if ($props) {
        foreach ($key in 'userPrincipalName', 'mail', 'displayName') {
            if ($props.ContainsKey($key) -and $props[$key]) {
                $label = [string]$props[$key]
                break
            }
        }
    }

    if (-not $label) { $label = $Member.Id }
    $script:UserLabels[$Member.Id] = $label
}

function Get-MethodType {
    <#
        De Graph SDK hangt het @odata.type in AdditionalProperties, niet als
        directe property. Beide varianten afvangen zodat de detectie niet stil
        leeg blijft (en dus iedereen als "niet geregistreerd" telt).
    #>
    param($Method)

    if ($Method.AdditionalProperties -and $Method.AdditionalProperties['@odata.type']) {
        return $Method.AdditionalProperties['@odata.type']
    }
    return $Method.'@odata.type'
}

function Get-MethodAaGuid {
    <#
        Idem voor de AAGUID van een fido2-methode: die zit in
        AdditionalProperties zolang de SDK het generieke authenticationMethod-
        type teruggeeft. Leeg -> onbekend, dan valt de aanroeper terug op
        /authentication/fido2Methods.
    #>
    param($Method)

    $aaGuid = $null
    if ($Method.AdditionalProperties -and $Method.AdditionalProperties['aaGuid']) {
        $aaGuid = $Method.AdditionalProperties['aaGuid']
    }
    elseif ($Method.aaGuid) {
        $aaGuid = $Method.aaGuid
    }
    elseif ($Method.AaGuid) {
        $aaGuid = $Method.AaGuid
    }

    if (-not $aaGuid) { return $null }
    return ([string]$aaGuid).Trim().Trim('{', '}').ToLowerInvariant()
}

function Get-Fido2AaGuidList {
    <#
        Fallback: haal de AAGUIDs op via het dedicated fido2Methods-endpoint.
        Alleen nodig als /authentication/methods de aaGuid niet meelevert;
        zonder deze fallback zou elke passkey stil als "niet Authenticator"
        gelden en zou niemand meer graduaten.
    #>
    param([string]$UserId)

    try {
        return @(Get-MgUserAuthenticationFido2Method -UserId $UserId -ErrorAction Stop |
            ForEach-Object { Get-MethodAaGuid -Method $_ } |
            Where-Object { $_ })
    }
    catch {
        Write-Warning "Kon fido2Methods niet ophalen voor $(Get-UserLabel -UserId $UserId) : $($_.Exception.Message)"
        return @()
    }
}

function Test-HasAcceptedMethod {
    <#
        Bepaalt of de gebruiker voldoet aan -AcceptedMethod. Bij
        AuthenticatorPasskey telt uitsluitend een fido2-methode waarvan de
        AAGUID in $NormalizedAllowedAaGuids staat; een YubiKey of WHfB is dan
        wel phishing-resistant maar niet wat we uitrollen.
    #>
    param(
        [array]$Methods,
        [string]$UserId
    )

    if ($AcceptedMethod -eq 'AnyPhishingResistant') {
        foreach ($m in $Methods) {
            if ((Get-MethodType -Method $m) -in $PhishingResistantTypes) { return $true }
        }
        return $false
    }

    $fido2Methods = @($Methods | Where-Object { (Get-MethodType -Method $_) -eq $Fido2Type })
    if ($fido2Methods.Count -eq 0) { return $false }

    $aaGuids = @($fido2Methods | ForEach-Object { Get-MethodAaGuid -Method $_ } | Where-Object { $_ })
    if ($aaGuids.Count -lt $fido2Methods.Count) {
        # Minstens één passkey zonder AAGUID in de payload -> los opvragen.
        $aaGuids = Get-Fido2AaGuidList -UserId $UserId
    }

    foreach ($aaGuid in $aaGuids) {
        if ($NormalizedAllowedAaGuids -contains $aaGuid) {
            Write-Verbose "$(Get-UserLabel -UserId $UserId) : passkey met toegestane AAGUID $aaGuid"
            return $true
        }
    }

    if ($aaGuids.Count -gt 0) {
        Write-Verbose "$(Get-UserLabel -UserId $UserId) : wel passkey(s), maar geen toegestane AAGUID ($($aaGuids -join ', '))"
    }
    return $false
}

# ---------------------------------------------------------------------------
# 2. Eén sync-pass: leden checken en graduatie/terugval uitvoeren
# ---------------------------------------------------------------------------
function Invoke-MfaSyncPass {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [string]$RolloutGroupId,
        [string]$RegisteredGroupId
    )

    Write-Output "=== Sync-pass gestart: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="

    # Elke pass mag zichzelf één keer herstellen van een verlopen token.
    $script:ReconnectAttempted = $false

    if ($AcceptedMethod -eq 'AuthenticatorPasskey') {
        Write-Output "Telt als geregistreerd : passkey met AAGUID $($NormalizedAllowedAaGuids -join ', ')"
    }
    else {
        Write-Output "Telt als geregistreerd : elke phishing-resistant methode"
    }

    # -ErrorAction Stop: zonder dit loopt een mislukte lees-actie door met 0 leden
    # en meldt de pass "niks te doen", terwijl er niets is gecontroleerd.
    $rolloutMembers    = @(Get-MgGroupMember -GroupId $RolloutGroupId -All -ErrorAction Stop)
    $registeredMembers = @(Get-MgGroupMember -GroupId $RegisteredGroupId -All -ErrorAction Stop)

    # Per pass opnieuw vullen: leden kunnen tussen twee passes wisselen.
    $script:UserLabels = @{}
    foreach ($member in @($rolloutMembers) + @($registeredMembers)) {
        Register-UserLabel -Member $member
    }

    $rolloutMemberIds    = @($rolloutMembers    | ForEach-Object { $_.Id })
    $registeredMemberIds = @($registeredMembers | ForEach-Object { $_.Id })

    Write-Output "Rollout-groep leden    : $($rolloutMemberIds.Count)"
    Write-Output "Registered-groep leden : $($registeredMemberIds.Count)"

    # Iedereen die in een van de twee groepen zit moet gecheckt worden
    $allUserIds = @(@($rolloutMemberIds) + @($registeredMemberIds) | Select-Object -Unique)
    $total      = $allUserIds.Count

    Write-Output "Te checken (uniek)     : $total"

    $toGraduate    = [System.Collections.Generic.List[string]]::new()   # Rollout -> Registered
    $toRevert      = [System.Collections.Generic.List[string]]::new()   # Registered -> Rollout
    $stillMissing  = [System.Collections.Generic.List[string]]::new()   # in Rollout, nog geen geaccepteerde methode
    $failedUserIds = [System.Collections.Generic.List[string]]::new()   # niet kunnen checken

    $i = 0
    $consecutiveFailures = 0
    foreach ($userId in $allUserIds) {
        $i++
        $label     = Get-UserLabel -UserId $userId
        $remaining = $total - $i

        Write-Progress -Activity "Authenticatiemethoden checken" `
                       -Status   "$i van $total - nu: $label (nog $remaining te gaan)" `
                       -PercentComplete ([int](100 * $i / [Math]::Max($total, 1)))
        Write-Verbose "[$i/$total] Checken: $label"

        try {
            $methods = Get-AuthMethodsResilient -UserId $userId
            $consecutiveFailures = 0
        }
        catch {
            $failedUserIds.Add($userId)
            $consecutiveFailures++

            # Een kapotte sessie raakt élke gebruiker. Doorgaan levert alleen een
            # browser-prompt per gebruiker op (die je dan 30x wegklikt), dus
            # meteen stoppen. Andere fouten zijn gebruiker-specifiek: overslaan,
            # tenzij er vijf op rij misgaan.
            $isAuthError = Test-IsGraphAuthError -ErrorRecord $_
            if ($isAuthError -or $consecutiveFailures -ge 5) {
                Write-Progress -Activity "Authenticatiemethoden checken" -Completed

                # Expliciet benoemen waar de pass bleef steken, anders weet je na
                # het afbreken niet wie er nog gecheckt moet worden.
                $skipBack   = if ($isAuthError) { 1 } else { $consecutiveFailures }
                $notChecked = @($allUserIds | Select-Object -Skip ($i - $skipBack))
                Write-Warning "Afgebroken bij $label ($i van $total). Nog niet gecheckt: $($notChecked.Count) gebruiker(s)."
                foreach ($pendingId in $notChecked) {
                    Write-Output "  ? Nog te checken: $(Get-UserLabel -UserId $pendingId)"
                }

                if ($isAuthError) {
                    if ($script:ReuseCallerSession) {
                        throw "Je eigen Graph-sessie is tijdens de run verlopen. Log opnieuw in en start het script opnieuw: Connect-MgGraph -Scopes $($InteractiveScopes -join ',') -ContextScope CurrentUser. Loopt dit vaker mis, draai dan met -UseTemporaryApp: dan werkt de run app-only en kan er geen prompt meer komen. Laatste fout: $($_.Exception.Message)"
                    }
                    throw "Graph-sessie is niet meer bruikbaar en opnieuw inloggen hielp niet. Draai met -UseTemporaryApp (wegwerp-app, app-only) of -UseAppRegistration met certificaat. Laatste fout: $($_.Exception.Message)"
                }
                throw "Vijf gebruikers op rij mislukt. Laatste fout: $($_.Exception.Message)"
            }

            Write-Warning "Kon authenticatiemethoden niet ophalen voor $label : $($_.Exception.Message)"
            continue
        }

        $hasAcceptedMethod = Test-HasAcceptedMethod -Methods $methods -UserId $userId
        $inRollout         = $userId -in $rolloutMemberIds
        $inRegistered      = $userId -in $registeredMemberIds

        if ($hasAcceptedMethod -and $inRollout) {
            $toGraduate.Add($userId)
        }
        elseif (-not $hasAcceptedMethod -and $inRegistered) {
            $toRevert.Add($userId)
        }
        elseif (-not $hasAcceptedMethod -and $inRollout) {
            # Zit al in Rollout en heeft nog niets geregistreerd: geen actie voor
            # het script, wél de groep waar jij achteraan moet.
            $stillMissing.Add($userId)
        }
        # else: status is al correct, geen actie
    }

    Write-Progress -Activity "Authenticatiemethoden checken" -Completed

    # -- Graduatie: Rollout -> Registered --
    foreach ($userId in $toGraduate) {
        if ($PSCmdlet.ShouldProcess($userId, "Graduatie: Rollout -> Registered")) {
            try {
                New-MgGroupMember -GroupId $RegisteredGroupId -DirectoryObjectId $userId -ErrorAction Stop
                Remove-MgGroupMemberByRef -GroupId $RolloutGroupId -DirectoryObjectId $userId -ErrorAction Stop
                Write-Output "  -> Gegradueerd: $(Get-UserLabel -UserId $userId)"
            }
            catch {
                Write-Warning "Graduatie mislukt voor $(Get-UserLabel -UserId $userId) : $($_.Exception.Message)"
            }
        }
    }

    # -- Terugval: Registered -> Rollout --
    foreach ($userId in $toRevert) {
        if ($PSCmdlet.ShouldProcess($userId, "Terugval: Registered -> Rollout")) {
            try {
                New-MgGroupMember -GroupId $RolloutGroupId -DirectoryObjectId $userId -ErrorAction Stop
                Remove-MgGroupMemberByRef -GroupId $RegisteredGroupId -DirectoryObjectId $userId -ErrorAction Stop
                Write-Output "  <- Teruggevallen (methode verloren): $(Get-UserLabel -UserId $userId)"
            }
            catch {
                Write-Warning "Terugval mislukt voor $(Get-UserLabel -UserId $userId) : $($_.Exception.Message)"
            }
        }
    }

    # -- Wie moet er nog wat doen --
    if ($stillMissing.Count -gt 0) {
        Write-Output "----------------------------------------"
        Write-Output "Nog te registreren (in Rollout, nog geen geaccepteerde methode):"
        foreach ($userId in $stillMissing) {
            Write-Output "  ! $(Get-UserLabel -UserId $userId)"
        }
    }

    if ($failedUserIds.Count -gt 0) {
        Write-Output "----------------------------------------"
        Write-Output "Niet kunnen checken (status onbekend, volgende pass opnieuw):"
        foreach ($userId in $failedUserIds) {
            Write-Output "  ? $(Get-UserLabel -UserId $userId)"
        }
    }

    if ($PendingCsvPath) {
        # Los bestand zodat je de openstaande gebruikers kunt mailen/opvolgen
        # zonder de console-output uit te moeten pluizen.
        try {
            $pendingRows = @(
                foreach ($userId in $stillMissing) {
                    [pscustomobject]@{
                        UserId = $userId
                        User   = Get-UserLabel -UserId $userId
                        Status = 'NogTeRegistreren'
                    }
                }
                foreach ($userId in $failedUserIds) {
                    [pscustomobject]@{
                        UserId = $userId
                        User   = Get-UserLabel -UserId $userId
                        Status = 'NietGecheckt'
                    }
                }
            )
            $pendingRows | Export-Csv -Path $PendingCsvPath -NoTypeInformation -Encoding UTF8 -Force
            Write-Output "Openstaande gebruikers weggeschreven naar: $PendingCsvPath ($($pendingRows.Count) regel(s))"
        }
        catch {
            Write-Warning "Kon $PendingCsvPath niet schrijven: $($_.Exception.Message)"
        }
    }

    Write-Output "----------------------------------------"
    Write-Output "Gecheckte gebruikers (Rollout + Registered) : $($total - $failedUserIds.Count) van $total"
    Write-Output "Gegradueerd (Rollout -> Registered)          : $($toGraduate.Count)"
    Write-Output "Teruggevallen (Registered -> Rollout)        : $($toRevert.Count)"
    Write-Output "Nog te registreren (in Rollout)              : $($stillMissing.Count)"
    Write-Output "Niet kunnen checken                          : $($failedUserIds.Count)"
    Write-Output "=== Sync-pass klaar: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
}

# ---------------------------------------------------------------------------
# 3. Uitvoering: eenmalig, of continu met interne pauze
# ---------------------------------------------------------------------------
# finally: ook bij een fout of Ctrl+C moet de tijdelijke app weg. Blijft hij
# staan, dan heb je een app met UserAuthenticationMethod.Read.All in je tenant.
try {
    if ($RunContinuously) {
        Write-Output "Continue modus: sync-pass elke $IntervalMinutes minuten. Stop met Ctrl+C of het stoppen van de job."
        while ($true) {
            try {
                Invoke-MfaSyncPass -RolloutGroupId $RolloutGroupId -RegisteredGroupId $RegisteredGroupId
            }
            catch {
                Write-Warning "Sync-pass mislukt: $($_.Exception.Message)"
            }

            Write-Output "Volgende sync-pass over $IntervalMinutes minuten ($((Get-Date).AddMinutes($IntervalMinutes)))..."
            Start-Sleep -Seconds ($IntervalMinutes * 60)

            # Graph-tokens verlopen doorgaans na ~1 uur; her-verbinden voor de zekerheid.
            # -Silent hergebruikt een geldige interactieve sessie i.p.v. elke keer een browser te openen.
            try {
                Connect-GraphSession -Silent
            }
            catch {
                Write-Warning "Her-verbinden mislukt: $($_.Exception.Message)"
            }
        }
    }
    else {
        Invoke-MfaSyncPass -RolloutGroupId $RolloutGroupId -RegisteredGroupId $RegisteredGroupId
    }
}
finally {
    Remove-TemporaryGraphApp
}