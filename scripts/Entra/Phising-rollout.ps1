<#
.SYNOPSIS
    Houdt twee elkaar uitsluitende statische groepen bij op basis van of een
    gebruiker een phishing-resistant MFA-methode heeft geregistreerd:
      - Rollout-groep    : moet nog registreren (in scope van de CA-eis)
      - Registered-groep : heeft al geregistreerd (compliant)

.DESCRIPTION
    Voor elke gebruiker die momenteel in Rollout OF Registered zit:
      - Heeft phishing-resistant methode EN zit in Rollout
            -> verwijderen uit Rollout, toevoegen aan Registered (graduatie)
      - Heeft GEEN phishing-resistant methode meer EN zit in Registered
            -> verwijderen uit Registered, toevoegen aan Rollout (terugval,
               methode verwijderd/verlopen -> moet opnieuw registreren)
      - Andere combinaties: geen actie, status is al correct.

    Een gebruiker zit dus nooit in beide groepen tegelijk.

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

.PARAMETER Interactive
    Log interactief in (browser-prompt) i.p.v. managed identity. Gebruik dit
    als je het script lokaal op je eigen machine draait. Zonder deze switch
    en zonder -UseAppRegistration wordt managed identity geprobeerd, wat
    buiten Azure altijd faalt.

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

    [switch]$UseAppRegistration,
    [switch]$Interactive,
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

function Connect-GraphSession {
    <#
        Verbindt volgens de gekozen modus en gooit een terminating error als dat
        niet lukt. Zonder deze check draait het script vrolijk door en rapporteert
        het 0 leden, wat lijkt op "niks te doen" terwijl er niets is gecontroleerd.
    #>
    [CmdletBinding()]
    param([switch]$Silent)

    try {
        if ($UseAppRegistration) {
            if (-not ($TenantId -and $ClientId -and $ClientCertificateThumbprint)) {
                throw "-UseAppRegistration vereist -TenantId, -ClientId en -ClientCertificateThumbprint."
            }
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $ClientCertificateThumbprint -NoWelcome -ErrorAction Stop
        }
        elseif ($Interactive) {
            # Bestaande sessie ALTIJD hergebruiken. De SDK vernieuwt het access
            # token zelf via de refresh token in de cache; opnieuw Connect-MgGraph
            # draaien levert alleen een browser-prompt op. Niet vergelijken op
            # scope-namen: Entra geeft consented scopes terug, die kunnen breder
            # of anders benoemd zijn dan wat we vroegen (bv. Group.ReadWrite.All
            # i.p.v. GroupMember.ReadWrite.All) -> zo'n check faalt dan elke pass.
            if ($Silent -and (Get-MgContext)) { return }

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
$PhishingResistantTypes = @(
    "#microsoft.graph.fido2AuthenticationMethod",                 # FIDO2-sleutels en passkeys (incl. Authenticator passkey)
    "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod",
    "#microsoft.graph.platformCredentialAuthenticationMethod",    # macOS Platform SSO
    "#microsoft.graph.x509CertificateAuthenticationMethod"        # CBA
)

function Get-MethodType {
    <#
        De Graph SDK hangt het @odata.type in AdditionalProperties, niet als
        directe property. Beide varianten afvangen zodat de detectie niet stil
        leeg blijft (en dus iedereen als "niet phishing-resistant" telt).
    #>
    param($Method)

    if ($Method.AdditionalProperties -and $Method.AdditionalProperties['@odata.type']) {
        return $Method.AdditionalProperties['@odata.type']
    }
    return $Method.'@odata.type'
}

function Test-IsPhishingResistant {
    param([array]$Methods)
    foreach ($m in $Methods) {
        if ((Get-MethodType -Method $m) -in $PhishingResistantTypes) { return $true }
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

        $isPhishingResistant = Test-IsPhishingResistant -Methods $methods
        $inRollout           = $userId -in $rolloutMemberIds
        $inRegistered        = $userId -in $registeredMemberIds

        if ($isPhishingResistant -and $inRollout) {
            $toGraduate.Add($userId)
        }
        elseif (-not $isPhishingResistant -and $inRegistered) {
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