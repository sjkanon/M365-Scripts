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

.PARAMETER ForceLogin
    Negeer een bestaande Graph-sessie en log opnieuw in. Zonder deze switch
    hergebruikt -Interactive een sessie die de benodigde scopes al heeft, zodat
    je niet bij elke run een browser-prompt krijgt. Gebruik dit als je van
    account wilt wisselen of opnieuw wilt consenten.

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
    [switch]$Interactive,
    [switch]$ForceLogin,
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientCertificateThumbprint,

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
$RequiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Groups", "Microsoft.Graph.Identity.SignIns")
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

# Zonder -Interactive/-UseAppRegistration is managed identity de bedoelde modus,
# maar buiten Azure faalt die altijd. Draai je lokaal en heb je zelf al een
# Connect-MgGraph gedaan, dan is die sessie duidelijk de bedoeling: hergebruiken
# i.p.v. afbreken op een IMDS-timeout van 169.254.169.254.
$script:UseExistingSession = $false
if (-not ($UseAppRegistration -or $Interactive)) {
    $existingContext = Get-MgContext
    if ($existingContext -and $existingContext.AuthType -eq 'Delegated') {
        $script:UseExistingSession = $true
        Write-Warning "Bestaande Graph-sessie gevonden ($($existingContext.Account)); die wordt hergebruikt i.p.v. managed identity."
        if (-not (Test-GraphSessionUsable -Context $existingContext)) {
            Write-Warning "Die sessie mist waarschijnlijk scopes: $($existingContext.Scopes -join ', '). Verwacht: $($InteractiveScopes -join ', '). Een kale 'Connect-MgGraph' krijgt alleen User.Read en loopt dus vast op de groepen - draai met -Interactive."
        }
    }
}

function Connect-GraphSession {
    <#
        Verbindt volgens de gekozen modus en gooit een terminating error als dat
        niet lukt. Zonder deze check draait het script vrolijk door en rapporteert
        het 0 leden, wat lijkt op "niks te doen" terwijl er niets is gecontroleerd.
    #>
    [CmdletBinding()]
    param([switch]$Silent)

    if ($script:UseExistingSession) {
        # De SDK vernieuwt het token zelf via de refresh token in de cache.
        if (Get-MgContext) { return }
        throw "De hergebruikte Graph-sessie is verlopen of verbroken. Draai opnieuw met -Interactive."
    }

    try {
        if ($UseAppRegistration) {
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
            if ($context -and -not $ForceLogin) {
                if (Test-GraphSessionUsable -Context $context) { return }

                # In de continue lus nooit alsnog een browser openen: dat blokkeert
                # een onbewaakte run tot iemand toevallig langsloopt.
                if ($Silent) { return }

                Write-Warning "Bestaande sessie ($($context.Account)) mist waarschijnlijk scopes: $($context.Scopes -join ', '). Opnieuw inloggen."
            }

            $connectParams = @{ Scopes = $InteractiveScopes; NoWelcome = $true; ErrorAction = 'Stop' }
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
        Write-Warning "Kon fido2Methods niet ophalen voor $UserId : $($_.Exception.Message)"
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
            Write-Verbose "$UserId : passkey met toegestane AAGUID $aaGuid"
            return $true
        }
    }

    if ($aaGuids.Count -gt 0) {
        Write-Verbose "$UserId : wel passkey(s), maar geen toegestane AAGUID ($($aaGuids -join ', '))"
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
    if ($AcceptedMethod -eq 'AuthenticatorPasskey') {
        Write-Output "Telt als geregistreerd : passkey met AAGUID $($NormalizedAllowedAaGuids -join ', ')"
    }
    else {
        Write-Output "Telt als geregistreerd : elke phishing-resistant methode"
    }

    # -ErrorAction Stop: zonder dit loopt een mislukte lees-actie door met 0 leden
    # en meldt de pass "niks te doen", terwijl er niets is gecontroleerd.
    $rolloutMemberIds    = (Get-MgGroupMember -GroupId $RolloutGroupId -All -ErrorAction Stop).Id
    $registeredMemberIds = (Get-MgGroupMember -GroupId $RegisteredGroupId -All -ErrorAction Stop).Id

    Write-Output "Rollout-groep leden    : $($rolloutMemberIds.Count)"
    Write-Output "Registered-groep leden : $($registeredMemberIds.Count)"

    # Iedereen die in een van de twee groepen zit moet gecheckt worden
    $allUserIds = @($rolloutMemberIds) + @($registeredMemberIds) | Select-Object -Unique

    $toGraduate = [System.Collections.Generic.List[string]]::new()   # Rollout -> Registered
    $toRevert   = [System.Collections.Generic.List[string]]::new()   # Registered -> Rollout

    $i = 0
    $consecutiveFailures = 0
    foreach ($userId in $allUserIds) {
        $i++
        if ($i % 50 -eq 0) { Write-Output "Verwerkt: $i / $($allUserIds.Count)" }

        try {
            $methods = Get-MgUserAuthenticationMethod -UserId $userId -ErrorAction Stop
            $consecutiveFailures = 0
        }
        catch {
            # Bij een kapotte sessie faalt élke gebruiker; dan de pass afbreken i.p.v.
            # honderden identieke waarschuwingen produceren en daarna "0 acties" melden.
            $consecutiveFailures++
            if ($consecutiveFailures -ge 5) {
                throw "Vijf gebruikers op rij mislukt - waarschijnlijk een verlopen of afgebroken Graph-sessie. Laatste fout: $($_.Exception.Message)"
            }
            Write-Warning "Kon authenticatiemethoden niet ophalen voor $userId : $($_.Exception.Message)"
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
        # else: status is al correct, geen actie
    }

    # -- Graduatie: Rollout -> Registered --
    foreach ($userId in $toGraduate) {
        if ($PSCmdlet.ShouldProcess($userId, "Graduatie: Rollout -> Registered")) {
            try {
                New-MgGroupMember -GroupId $RegisteredGroupId -DirectoryObjectId $userId -ErrorAction Stop
                Remove-MgGroupMemberByRef -GroupId $RolloutGroupId -DirectoryObjectId $userId -ErrorAction Stop
                Write-Output "  -> Gegradueerd: $userId"
            }
            catch {
                Write-Warning "Graduatie mislukt voor $userId : $($_.Exception.Message)"
            }
        }
    }

    # -- Terugval: Registered -> Rollout --
    foreach ($userId in $toRevert) {
        if ($PSCmdlet.ShouldProcess($userId, "Terugval: Registered -> Rollout")) {
            try {
                New-MgGroupMember -GroupId $RolloutGroupId -DirectoryObjectId $userId -ErrorAction Stop
                Remove-MgGroupMemberByRef -GroupId $RegisteredGroupId -DirectoryObjectId $userId -ErrorAction Stop
                Write-Output "  <- Teruggevallen (methode verloren): $userId"
            }
            catch {
                Write-Warning "Terugval mislukt voor $userId : $($_.Exception.Message)"
            }
        }
    }

    Write-Output "----------------------------------------"
    Write-Output "Gecheckte gebruikers (Rollout + Registered) : $($allUserIds.Count)"
    Write-Output "Gegradueerd (Rollout -> Registered)          : $($toGraduate.Count)"
    Write-Output "Teruggevallen (Registered -> Rollout)        : $($toRevert.Count)"
    Write-Output "=== Sync-pass klaar: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
}

# ---------------------------------------------------------------------------
# 3. Uitvoering: eenmalig, of continu met interne pauze
# ---------------------------------------------------------------------------
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