# ============================================================
# Teams Archivering - Volledig Automatisch Script v9.0
# PowerShell 7+ vereist | Uitvoeren als Global Admin
# ============================================================
#
# Aanmelden (via scripts\Startup\Connect-M365.ps1):
#   - Standaard gedelegeerd: je meldt je aan als de admin van de klant, met een
#     apparaatcode zoals load.config.ps1 zegt (zonder load.config: apparaatcode,
#     zoals vroeger). Geen tijdelijke app-registratie meer.
#   - App-only op aanvraag: -ClientId + -CertificateThumbprint, of -AppOnly
#     (graph.appid.json). De app heeft dan de application-varianten nodig van
#     Group.Read.All, Sites.Read.All, TeamMember.Read.All, ChannelMessage.Read.All
#     (een protected API) en TeamSettings/ChannelSettings.ReadWrite.All.
#   - Alles loopt via Microsoft Graph (teams, leden, kanalen, bestanden, chat,
#     archiveren). Alleen het toekennen van site-admin rechten bij een "access
#     denied" (Set-PnPSite -Owners) heeft geen Graph-API en blijft PnP; dat
#     gebruikt de PnP-app uit pnp.appid.json of -PnPClientId.

param(
    [ValidateSet("interactive", "archive", "undo", "skip")]
    [string]$Step10Action = "interactive",
    [switch]$Step10Only,
    [ValidateSet("none", "archive", "undo")]
    [string]$ChannelAction = "none",
    [string]$ChannelArchiveTag = "[ARCHIEF]",
    [switch]$ChannelFallbackToRename,
    [switch]$DryRun,
    # Werkblad met de Teams-lijst. Leeg = het eerste werkblad van het Excel-bestand.
    [string]$WorksheetName,
    # Tenant ID van de klant. Leeg = de GDAP-klant (Connect-Tenant) als standaard in de wizard.
    [string]$TenantId,
    # App-only in plaats van gedelegeerd: app-registratie + certificaat.
    [string]$ClientId,
    [string]$CertificateThumbprint,
    # App-only met ClientId/CertificateThumbprint uit graph.appid.json.
    [switch]$AppOnly,
    # PnP-app voor de site-admin fallback. Leeg = pnp.appid.json.
    [string]$PnPClientId
)

#region ZELFHERSTART - Modules opkuisen en sessie hernieuwen
# Dit blok zorgt ervoor dat het script zichzelf herstart in een schone sessie
# nadat conflicterende modules zijn verwijderd. Zonder herstart blijven
# oude module-versies actief in het geheugen en crashen alle Graph-calls.

$herstart = $env:TEAMS_ARCHIVER_HERSTART

if ($herstart -ne "1") {

    Write-Host "`n[PRE] Conflicterende Graph-modules opkuisen..." -ForegroundColor Cyan

    # Verwijder alle Graph-module bestanden van schijf
    $modulePaden = if ($IsWindows) {
        @(
            "$env:USERPROFILE\Documents\PowerShell\Modules",
            "$env:USERPROFILE\Documents\WindowsPowerShell\Modules",
            "C:\Program Files\PowerShell\Modules",
            "C:\Program Files\WindowsPowerShell\Modules"
        )
    } else {
        @(
            (Join-Path $HOME ".local" "share" "powershell" "Modules"),
            (Join-Path $HOME "Documents" "PowerShell" "Modules"),
            "/usr/local/share/powershell/Modules",
            "/usr/share/powershell/Modules"
        )
    }
    foreach ($pad in $modulePaden) {
        $gevonden = Get-ChildItem (Join-Path $pad "Microsoft.Graph*") -ErrorAction SilentlyContinue
        foreach ($map in $gevonden) {
            Remove-Item $map.FullName -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "  Verwijderd: $($map.FullName)" -ForegroundColor Gray
        }
    }

    # Installeer correcte versies
    Write-Host "  Installeren: Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop

    # De Teams-module is niet meer nodig: teams, leden en kanalen komen uit Graph.
    foreach ($mod in @("PnP.PowerShell","ImportExcel")) {
        if (-not (Get-Module -ListAvailable $mod -ErrorAction SilentlyContinue)) {
            Write-Host "  Installeren: $mod..." -ForegroundColor Yellow
            Install-Module $mod -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        Write-Host "  OK: $mod" -ForegroundColor Green
    }

    Write-Host "`n  Modules geinstalleerd. Script herstart in schone sessie...`n" -ForegroundColor Cyan

    # Herstart het script in een nieuwe pwsh-sessie met de herstart-vlag.
    # De parameters gaan mee: zonder dat viel o.a. -DryRun bij elke eerste run weg
    # en archiveerde de herstarte sessie echt.
    $env:TEAMS_ARCHIVER_HERSTART = "1"
    # De nieuwe sessie start met -NoProfile en kent load.config.ps1 niet: geef de
    # aanmeldinstellingen mee. Zonder instelling blijft het de apparaatcode van vroeger.
    $env:TEAMS_ARCHIVER_DEVICECODE = if ($null -eq $global:useDeviceCodeAuth) { "1" } elseif ($global:useDeviceCodeAuth) { "1" } else { "0" }
    $env:TEAMS_ARCHIVER_AUTHMODE   = [string]$global:authMode
    $env:TEAMS_ARCHIVER_CID        = [string]$global:cid
    $env:TEAMS_ARCHIVER_UPN        = [string]$global:upn
    $pwshPath = (Get-Command pwsh).Source
    $restartArgs = if ($IsWindows) {
        @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $MyInvocation.MyCommand.Path)
    } else {
        @("-NoProfile", "-File", $MyInvocation.MyCommand.Path)
    }
    foreach ($param in $PSBoundParameters.GetEnumerator()) {
        if ($param.Value -is [switch]) {
            if ($param.Value.IsPresent) { $restartArgs += "-$($param.Key)" }
        } else {
            $restartArgs += "-$($param.Key)"
            $restartArgs += [string]$param.Value
        }
    }
    & $pwshPath @restartArgs
    $restartExit = $LASTEXITCODE
    # Niet laten hangen in de aanroepende sessie: anders slaat een volgende run de opkuis over.
    foreach ($v in 'TEAMS_ARCHIVER_HERSTART','TEAMS_ARCHIVER_DEVICECODE','TEAMS_ARCHIVER_AUTHMODE','TEAMS_ARCHIVER_CID','TEAMS_ARCHIVER_UPN') {
        Remove-Item "Env:$v" -ErrorAction SilentlyContinue
    }
    exit $restartExit
}

# Vanaf hier: we zitten in de hergestarte schone sessie
Write-Host "`n  Schone sessie actief. Modules worden geladen..." -ForegroundColor Green

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# Aanmeldinstellingen van de aanroepende sessie terugzetten (zie de herstart hierboven).
if ($env:TEAMS_ARCHIVER_DEVICECODE) { $global:useDeviceCodeAuth = $env:TEAMS_ARCHIVER_DEVICECODE -eq "1" }
if ($env:TEAMS_ARCHIVER_AUTHMODE)   { $global:authMode = $env:TEAMS_ARCHIVER_AUTHMODE }
if ($env:TEAMS_ARCHIVER_CID)        { $global:cid = $env:TEAMS_ARCHIVER_CID }
if ($env:TEAMS_ARCHIVER_UPN)        { $global:upn = $env:TEAMS_ARCHIVER_UPN }

. (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')

Write-Host "  Graph module geladen." -ForegroundColor Green

$wantAppOnly = [bool]$ClientId -or $AppOnly

# Cross-platform tijdelijke map
$tempDir = [System.IO.Path]::GetTempPath()
#endregion

#region CONFIGURATIE - Interactief opvragen
Clear-Host
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Teams Archivering - Setup Wizard v9.0" -ForegroundColor Cyan
Write-Host "============================================`n" -ForegroundColor Cyan

# Excel-bestand
Write-Host "Stap 1/5 - Excel-bestand met de Teams-lijst" -ForegroundColor Yellow
$xlStandaard = if ($IsWindows) { "C:\Temp\Teams_Channels.xlsx" } else { Join-Path $HOME "Downloads" "Teams_Channels.xlsx" }
Write-Host "  Standaard: $xlStandaard"
Write-Host "  Druk Enter voor standaard, of typ een ander pad.`n"
$xlInput = Read-Host "  Pad naar Excel-bestand"
$xlPath  = if ($xlInput.Trim()) { $xlInput.Trim() } else { $xlStandaard }
if (-not (Test-Path $xlPath)) {
    Write-Host "  FOUT: Bestand niet gevonden: $xlPath" -ForegroundColor Red; exit
}
Write-Host "  OK: $xlPath`n" -ForegroundColor Green

# Archief-map
Write-Host "Stap 2/5 - Archief-map" -ForegroundColor Yellow
$archiveStandaard = if ($IsWindows) { "C:\Temp\Teams_Archive" } else { Join-Path $HOME "Documents" "Teams_Archive" }
Write-Host "  Standaard: $archiveStandaard"
Write-Host "  Druk Enter voor standaard, of typ een ander pad.`n"
$archiveInput = Read-Host "  Archief-map"
$archiveRoot  = if ($archiveInput.Trim()) { $archiveInput.Trim() } else { $archiveStandaard }
$archiveParent = Split-Path $archiveRoot -Parent
if ($archiveParent -and -not (Test-Path $archiveParent)) {
    Write-Host "  FOUT: Map niet bereikbaar: $archiveParent" -ForegroundColor Red; exit
}
New-Item -ItemType Directory -Path $archiveRoot -Force | Out-Null
Write-Host "  OK: $archiveRoot`n" -ForegroundColor Green

# Tenant ID
Write-Host "Stap 3/5 - Tenant ID van de Microsoft 365-omgeving van de klant" -ForegroundColor Yellow
Write-Host "  Vind je via: https://entra.microsoft.com > Microsoft Entra ID > Overview"
Write-Host "  Formaat: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`n"
$tenantStandaard = Resolve-M365TenantId -TenantId $TenantId
if ($tenantStandaard -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') { $tenantStandaard = $null }
if ($tenantStandaard) { Write-Host "  Standaard (Enter): $tenantStandaard" }
do {
    $tenantId = (Read-Host "  Tenant ID").Trim()
    if (-not $tenantId -and $tenantStandaard) { $tenantId = $tenantStandaard.ToLowerInvariant() }
    if ($tenantId -notmatch '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$') {
        Write-Host "  Ongeldig formaat. Probeer opnieuw.`n" -ForegroundColor Red
    }
} while ($tenantId -notmatch '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$')
Write-Host "  OK: $tenantId`n" -ForegroundColor Green

# SharePoint URL
Write-Host "Stap 4/5 - SharePoint URL van de tenant van de klant" -ForegroundColor Yellow
Write-Host "  Formaat: https://naam.sharepoint.com  (geen slash aan het einde)`n"
do {
    $tenantUrl = (Read-Host "  SharePoint URL").Trim().TrimEnd('/')
    if ($tenantUrl -notmatch '^https://[a-zA-Z0-9-]+\.sharepoint\.com$') {
        Write-Host "  Ongeldig formaat. Voorbeeld: https://contoso.sharepoint.com`n" -ForegroundColor Red
    }
} while ($tenantUrl -notmatch '^https://[a-zA-Z0-9-]+\.sharepoint\.com$')
Write-Host "  OK: $tenantUrl`n" -ForegroundColor Green

# Chat-methode
Write-Host "Stap 5/5 - Methode voor chat-export" -ForegroundColor Yellow
Write-Host "  A) Graph API  - Automatisch, JSON + CSV per kanaal"
Write-Host "  B) Purview    - Manueel via compliance.microsoft.com (E3/E5 vereist)`n"
do {
    $chatMethode = (Read-Host "  Kies methode (a/b)").Trim().ToLower()
} while ($chatMethode -notin @("a","b"))
Write-Host "  OK: Methode $($chatMethode.ToUpper())`n" -ForegroundColor Green

# Bevestiging
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Configuratie:"
Write-Host "  Excel          : $xlPath"
Write-Host "  Archief        : $archiveRoot"
Write-Host "  Tenant ID      : $tenantId"
Write-Host "  SharePoint URL : $tenantUrl"
Write-Host "  Chat methode   : $(if ($chatMethode -eq 'a') {'Graph API (auto)'} else {'Purview (manueel)'})"
Write-Host "============================================`n" -ForegroundColor Cyan
$bevestig = Read-Host "Klopt alles? Start? (j/n)"
if ($bevestig -ne "j") { Write-Host "Geannuleerd." -ForegroundColor Yellow; exit }
#endregion

#region STAP 1-4 - Inloggen (Graph, gedelegeerd standaard)
Write-Host "`n[1/12] Inloggen op Microsoft Graph..." -ForegroundColor Cyan

# Gedelegeerd: de Microsoft Graph Command Line Tools-app met deze scopes; de admin geeft
# bij de eerste aanmelding toestemming. Een tijdelijke app-registratie is niet meer nodig.
$graphScopes = @(
    "Group.ReadWrite.All","Sites.Read.All","Files.ReadWrite.All",
    "TeamSettings.ReadWrite.All","TeamMember.Read.All","ChannelMessage.Read.All","ChannelSettings.ReadWrite.All"
)

if ($wantAppOnly) {
    Write-Host "  App-only: de app moet in de tenant van de klant toestemming hebben.`n" -ForegroundColor Yellow
} else {
    if ($global:useDeviceCodeAuth) { Write-Host "  Je krijgt een code + link." -ForegroundColor Yellow }
    Write-Host "  Log in met het GLOBAL ADMIN account van de klant (of je GDAP-partneraccount).`n" -ForegroundColor Yellow
}
if ($DryRun) {
    Write-Host "  DRY RUN: login wordt normaal uitgevoerd; alleen mutaties worden gesimuleerd." -ForegroundColor DarkYellow
}

$graphParams = @{ Scopes = $graphScopes; TenantId = $tenantId }
if ($ClientId)              { $graphParams['ClientId'] = $ClientId }
if ($CertificateThumbprint) { $graphParams['CertificateThumbprint'] = $CertificateThumbprint }
if ($AppOnly)               { $graphParams['AppOnly'] = $true }

try {
    $graphConn = Connect-M365Graph @graphParams
} catch {
    Write-Host "`n  FOUT bij inloggen: $_" -ForegroundColor Red
    Write-Host "  Controleer:" -ForegroundColor Yellow
    Write-Host "   - Tenant ID correct? ($tenantId)"
    Write-Host "   - Log je in met het admin-account van de klant, of heb je GDAP-rechten op deze klant?"
    Write-Host "   - Is MFA ingesteld op dit account?"
    exit 1
}

Write-Host "`n[2/12] Sessie controleren..." -ForegroundColor Cyan
$ctx = Get-MgContext
if ($ctx.TenantId -and $ctx.TenantId -ne $tenantId) {
    Write-Host "  FOUT: aangemeld in tenant $($ctx.TenantId), niet in $tenantId." -ForegroundColor Red
    Disconnect-M365Graph $graphConn
    exit 1
}
$account = $ctx.Account
if ($graphConn.AuthType -eq 'Delegated') {
    if (-not $account) {
        Write-Host "  FOUT: Inloggen mislukt, geen account gevonden." -ForegroundColor Red; exit 1
    }
    Write-Host "  Ingelogd als: $account" -ForegroundColor Green
    $ontbrekend = @($graphScopes | Where-Object { $_ -notin $ctx.Scopes })
    if ($ontbrekend) { Write-Warning "  Scopes niet toegekend: $($ontbrekend -join ', ')" }
} else {
    Write-Host "  App-only verbonden ($($ctx.ClientId))." -ForegroundColor Green
}

Write-Host "`n[3-4/12] SharePoint site-admin fallback voorbereiden..." -ForegroundColor Cyan
# Alleen gebruikt als een kanaalmap "access denied" geeft: Graph heeft geen API om een
# site-collectiebeheerder toe te voegen, dus dat ene gebeurt met PnP (Set-PnPSite -Owners).
if ($graphConn.AuthType -ne 'Delegated') {
    Write-Host "  App-only: niet nodig, de app leest alle sites." -ForegroundColor Gray
} elseif ($PnPClientId) {
    Write-Host "  PnP-app: $PnPClientId" -ForegroundColor Green
} elseif (Test-Path (Join-Path $PSScriptRoot '..\..\pnp.appid.json')) {
    Write-Host "  PnP-app: uit pnp.appid.json (als die een regel voor deze tenant heeft)." -ForegroundColor Green
} else {
    Write-Host "  Geen PnP-app (pnp.appid.json / -PnPClientId): bij 'access denied' moet je jezelf handmatig site-beheerder maken." -ForegroundColor Yellow
}
#endregion

#region Graph-hulpfuncties
function Invoke-GraphRequestWithRetry {
    param(
        [ValidateSet("GET","POST","PATCH","DELETE")]
        [string]$Method,
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [string]$OutputFilePath,
        [int]$MaxRetries = 5
    )

    $attempt = 0
    while ($true) {
        try {
            if ($OutputFilePath) {
                return Invoke-MgGraphRequest -Method $Method -Uri $Uri -OutputFilePath $OutputFilePath -ErrorAction Stop
            }
            return Invoke-MgGraphRequest -Method $Method -Uri $Uri -ErrorAction Stop
        } catch {
            $attempt++
            $message = $_.Exception.Message
            $isThrottle = $message -match "Status:\s*429|Too Many Requests|TooManyRequests|throttl"
            $isTransient = $message -match "Status:\s*5\d\d|ServiceUnavailable|GatewayTimeout|BadGateway|timeout|temporar"

            if (($isThrottle -or $isTransient) -and $attempt -lt $MaxRetries) {
                $waitSeconds = [Math]::Min(30, [Math]::Pow(2, $attempt))
                Write-Host "    Graph retry na ${waitSeconds}s (poging $attempt/$MaxRetries)..." -ForegroundColor DarkYellow
                Start-Sleep -Seconds $waitSeconds
                continue
            }

            throw
        }
    }
}

function Get-GraphPagedCollection {
    param([Parameter(Mandatory = $true)][string]$StartUri)

    $all = [System.Collections.Generic.List[object]]::new()
    $uri = $StartUri
    do {
        $result = Invoke-GraphRequestWithRetry -Method GET -Uri $uri
        if ($result.value) {
            $all.AddRange(@($result.value))
        }
        $uri = $result.'@odata.nextLink'
    } while ($uri)

    return $all
}

function Find-TeamByName {
    # Vervangt Get-Team -DisplayName: de groep met die naam die een Team is.
    param([Parameter(Mandatory = $true)][string]$DisplayName)
    $safe   = $DisplayName.Replace("'", "''")
    $filter = [System.Uri]::EscapeDataString("displayName eq '$safe'")
    $groups = Get-GraphPagedCollection -StartUri "https://graph.microsoft.com/v1.0/groups?`$filter=$filter&`$select=id,displayName,resourceProvisioningOptions"
    return $groups | Where-Object { @($_.resourceProvisioningOptions) -contains 'Team' } | Select-Object -First 1
}
#endregion

#region STAP 5 - Excel inlezen en Teams-IDs ophalen
Write-Host "`n[5/12] Excel inlezen en Teams-IDs ophalen..." -ForegroundColor Cyan

Import-Module ImportExcel -ErrorAction Stop

$excelArgs = @{ Path = $xlPath }
if ($WorksheetName) { $excelArgs['WorksheetName'] = $WorksheetName }
$data      = Import-Excel @excelArgs
$toArchive = $data | Where-Object { $_.Archive -eq "Archive" }

# Normaliseer Excel-waarden om lookup-missers (bv. trailing spaces) te vermijden.
foreach ($row in $toArchive) {
    $row.TeamName = ([string]$row.TeamName).Trim()
    $row.ChannelName = ([string]$row.ChannelName).Trim()
}

Write-Host "  $($toArchive.Count) kanalen gevonden." -ForegroundColor Green

$archiveTeams = $toArchive | Select-Object -ExpandProperty TeamName -Unique
$teamMapping  = @{}

foreach ($teamName in $archiveTeams) {
    $team = $null
    try { $team = Find-TeamByName -DisplayName $teamName } catch { Write-Warning "  Opzoeken mislukt voor ${teamName}: $_" }
    if ($team) {
        $teamMapping[$teamName] = [string]$team.id
        Write-Host "  OK: $teamName" -ForegroundColor Green
    } else {
        Write-Warning "  Niet gevonden: $teamName"
    }
}

if ($DryRun) {
    Write-Host "  [DRYRUN] Team mapping file wordt niet weggeschreven." -ForegroundColor Cyan
} else {
    $teamMapping | ConvertTo-Json | Out-File (Join-Path $tempDir "teams_archiver_team_mapping.json") -Force
}
Write-Host "  $($teamMapping.Count) van de $($archiveTeams.Count) Teams gevonden." -ForegroundColor Green
#endregion

function New-ArchiveRowKey {
    param(
        [Parameter(Mandatory = $true)][string]$TeamName,
        [Parameter(Mandatory = $true)][string]$ChannelName
    )
    return "$TeamName||$ChannelName"
}

function Normalize-LookupValue {
    param([Parameter(Mandatory = $false)][string]$Value)

    if ($null -eq $Value) { return "" }
    $s = [string]$Value
    $s = $s -replace "\s+", " "
    return $s.Trim().ToLowerInvariant()
}

$dryRunFileCounts = @{}
$dryRunChatCounts = @{}

if (-not $Step10Only) {

#region STAP 6 - Mapstructuur
Write-Host "`n[6/12] Mapstructuur aanmaken..." -ForegroundColor Cyan
foreach ($row in $toArchive) {
    $safeTeam = $row.TeamName -replace '[\\/:*?"<>|]', '_'
    $safeChannel = $row.ChannelName -replace '[\\/:*?"<>|]', '_'
    $channelRoot = Join-Path $archiveRoot $safeTeam $safeChannel
    if ($DryRun) {
        Write-Host "  [DRYRUN] Structuur gevalideerd: $($row.TeamName) / $($row.ChannelName)" -ForegroundColor Cyan
    } else {
        New-Item -ItemType Directory -Path (Join-Path $channelRoot "Files")   -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $channelRoot "Chat")    -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $channelRoot "Members") -Force | Out-Null
    }
}
Write-Host "  Mappen aangemaakt." -ForegroundColor Green
#endregion

#region STAP 7 - Ledenlijsten
Write-Host "`n[7/12] Ledenlijsten exporteren..." -ForegroundColor Cyan
foreach ($row in $toArchive) {
    $teamName = $row.TeamName
    $groupId = $teamMapping[$teamName]
    if (-not $groupId) { continue }
    $safeTeam = $teamName -replace '[\\/:*?"<>|]', '_'
    $safeChannel = $row.ChannelName -replace '[\\/:*?"<>|]', '_'
    try {
        # Vervangt Get-TeamUser: dezelfde kolommen Name, User en Role.
        $members = @(Get-GraphPagedCollection -StartUri "https://graph.microsoft.com/v1.0/teams/$groupId/members" | ForEach-Object {
            $roles = @($_.roles)
            [PSCustomObject]@{
                Name = $_.displayName
                User = $_.email
                Role = if ($roles -contains 'owner') { 'Owner' } elseif ($roles -contains 'guest') { 'Guest' } else { 'Member' }
            }
        })
        if ($DryRun) {
            Write-Host "  [DRYRUN] Leden gevonden: $teamName / $($row.ChannelName) ($($members.Count))" -ForegroundColor Cyan
        } else {
            $members |
                Select-Object Name, User, Role |
                Export-Csv (Join-Path $archiveRoot $safeTeam $safeChannel "Members" "members.csv") -NoTypeInformation -Encoding UTF8
            Write-Host "  OK: $teamName / $($row.ChannelName)" -ForegroundColor Green
        }
    } catch {
        Write-Warning "  Fout ledenlijst $teamName / $($row.ChannelName) : $_"
    }
}
#endregion

#region STAP 8 - Bestanden exporteren
Write-Host "`n[8/12] Bestanden exporteren (SharePoint -> $archiveRoot)..." -ForegroundColor Cyan

# Lezen en downloaden gaat via Graph (/drives/{id}/items/{id}/children en /content).
# PnP is alleen nog nodig voor de site-admin fallback hieronder.

function Test-IsAccessDeniedError {
    param([string]$Message)
    return $Message -match "Access denied|accessDenied|Unauthorized|Forbidden|Status:\s*401|Status:\s*403|Insufficient privileges"
}

function Test-IsNotFoundError {
    param([string]$Message)
    return $Message -match "Status:\s*404|NotFound|itemNotFound|Item does not exist|bestaat niet"
}

function Get-SiteUrlFromWebUrl {
    # https://contoso.sharepoint.com/sites/Team-Kanaal/Gedeelde documenten/Kanaal -> .../sites/Team-Kanaal
    param([string]$WebUrl)
    if ($WebUrl -match '^(https://[^/]+/(?:sites|teams)/[^/]+)') { return $Matches[1] }
    return $null
}

function Get-TeamChannelCached {
    param(
        [Parameter(Mandatory = $true)][string]$GroupId,
        [Parameter(Mandatory = $true)][string]$ChannelName,
        [Parameter(Mandatory = $true)][hashtable]$Cache
    )

    # Vervangt Get-TeamChannel: de kanalen komen uit Graph, eenmaal per team.
    if (-not $Cache.ContainsKey($GroupId)) {
        $Cache[$GroupId] = @(Get-GraphPagedCollection -StartUri "https://graph.microsoft.com/v1.0/teams/$GroupId/channels")
    }

    $targetNormalized = Normalize-LookupValue -Value $ChannelName

    $channel = $Cache[$GroupId] | Where-Object { $_.displayName -eq $ChannelName } | Select-Object -First 1
    if ($channel) { return $channel }

    return $Cache[$GroupId] |
        Where-Object { (Normalize-LookupValue -Value $_.displayName) -eq $targetNormalized } |
        Select-Object -First 1
}

function Grant-SiteAdminAccess {
    param(
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$AdminAccount,
        [Parameter(Mandatory = $true)][string]$TenantUrl
    )
    # Geen Graph-API om een site-collectiebeheerder toe te voegen: dit ene blijft PnP.
    $spoAdminUrl = $TenantUrl -replace '(https://[^.]+)(\.sharepoint\.com)', '$1-admin$2'
    try {
        Write-Host "    Site-admin rechten verlenen via SPO Admin voor: $SiteUrl" -ForegroundColor Yellow
        Import-Module PnP.PowerShell -ErrorAction Stop
        # pnp.appid.json is per tenantdomein (contoso.onmicrosoft.com), niet per GUID.
        $pnpTenant = if ($TenantUrl -match '^https://([^.]+)\.sharepoint\.com') { "$($Matches[1]).onmicrosoft.com" } else { $tenantId }
        $pnpParams = @{ Url = $spoAdminUrl; TenantId = $pnpTenant }
        if ($PnPClientId) { $pnpParams['ClientId'] = $PnPClientId }
        $pnp = Connect-M365PnP @pnpParams
        Set-PnPSite -Identity $SiteUrl -Owners @($AdminAccount) -Connection $pnp -ErrorAction Stop
        Write-Host "    Site-admin rechten verleend. Wachten op propagatie..." -ForegroundColor Green
        Start-Sleep -Seconds 10
        return $true
    } catch {
        Write-Warning "    Kon site-admin rechten niet verlenen voor ${SiteUrl}: $_"
        Write-Host "    TIP: Voeg '$AdminAccount' handmatig toe als site-beheerder in het SharePoint Admin Center." -ForegroundColor Yellow
        return $false
    }
}

function Get-DriveFilesRecursive {
    # Alle bestanden onder een map, met hun pad relatief tegenover die map.
    param(
        [Parameter(Mandatory = $true)][string]$DriveId,
        [Parameter(Mandatory = $true)][string]$ItemId,
        [string]$RelativePath = ""
    )
    $children = Get-GraphPagedCollection -StartUri "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ItemId/children?`$select=id,name,file,folder,size&`$top=999"
    foreach ($child in $children) {
        $rel = if ($RelativePath) { Join-Path $RelativePath $child.name } else { [string]$child.name }
        if ($child.folder) {
            Get-DriveFilesRecursive -DriveId $DriveId -ItemId $child.id -RelativePath $rel
        } elseif ($child.file) {
            [PSCustomObject]@{ DriveId = $DriveId; Id = $child.id; Name = $child.name; RelativePath = $rel; Size = $child.size }
        }
    }
}

function Save-DriveFilesReliable {
    param(
        [Parameter(Mandatory = $true)][array]$Items,
        [Parameter(Mandatory = $true)][string]$DestPath,
        [int]$MaxRetriesPerFile = 3
    )

    $failed = [System.Collections.Generic.List[object]]::new()
    foreach ($item in $Items) {
        # De mappenstructuur blijft behouden: twee bestanden met dezelfde naam in
        # verschillende submappen overschrijven elkaar niet meer.
        $target = Join-Path $DestPath $item.RelativePath
        $dir = Split-Path $target -Parent
        if ($dir) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $ok = $false
        for ($attempt = 1; $attempt -le $MaxRetriesPerFile; $attempt++) {
            try {
                Invoke-GraphRequestWithRetry -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/drives/$($item.DriveId)/items/$($item.Id)/content" `
                    -OutputFilePath $target | Out-Null
                $ok = $true
                break
            } catch {
                if ($attempt -lt $MaxRetriesPerFile) {
                    Start-Sleep -Seconds ([Math]::Min(5, $attempt))
                }
            }
        }
        if (-not $ok) { $failed.Add($item) }
    }

    return [PSCustomObject]@{
        FailedCount = $failed.Count
        FailedItems = $failed
    }
}

$adminGrantedSites   = @{}
$teamChannelCache    = @{}

foreach ($row in $toArchive) {
    $rowKey = New-ArchiveRowKey -TeamName $row.TeamName -ChannelName $row.ChannelName
    $groupId = $teamMapping[$row.TeamName]
    if (-not $groupId) {
        Write-Warning "  Geen GroupId voor: $($row.TeamName) — overgeslagen"
        continue
    }

    $safeteam    = $row.TeamName    -replace '[\\/:*?"<>|]', '_'
    $safechannel = $row.ChannelName -replace '[\\/:*?"<>|]', '_'
    $destPath    = Join-Path $archiveRoot $safeteam $safechannel "Files"
    if (-not $DryRun) {
        New-Item -ItemType Directory -Path $destPath -Force | Out-Null
    }

    try {
        $channel = Get-TeamChannelCached -GroupId $groupId -ChannelName $row.ChannelName -Cache $teamChannelCache
        if (-not $channel) {
            Write-Warning "  Kanaal niet gevonden: $($row.ChannelName)"
            continue
        }

        # filesFolder werkt voor standaard, privé en gedeelde kanalen (die laatste twee
        # hebben hun eigen site) en geeft de drive en de map rechtstreeks.
        $ff = Invoke-GraphRequestWithRetry -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/teams/$groupId/channels/$($channel.id)/filesFolder"
        $driveId  = [string]$ff.parentReference.driveId
        $folderId = [string]$ff.id
        $siteUrl  = Get-SiteUrlFromWebUrl -WebUrl $ff.webUrl
        if (-not $driveId -or -not $folderId) {
            Write-Warning "  FilesFolder zonder drive/map: $($row.TeamName) / $($row.ChannelName)"
            continue
        }

        try {
            $items = @(Get-DriveFilesRecursive -DriveId $driveId -ItemId $folderId)
        } catch {
            $errMsg = $_.Exception.Message
            if (Test-IsAccessDeniedError -Message $errMsg) {
                # Per site eenmalig SPO site-admin rechten toekennen. Alleen gedelegeerd:
                # app-only leest alle sites al.
                if ($graphConn.AuthType -ne 'Delegated' -or -not $siteUrl -or $adminGrantedSites.ContainsKey($siteUrl)) { throw }
                $adminGrantedSites[$siteUrl] = $true
                if ($DryRun) {
                    Write-Host "  [DRYRUN] Geen toegang; zou site-admin rechten verlenen op $siteUrl." -ForegroundColor Cyan
                    continue
                }
                $granted = Grant-SiteAdminAccess -SiteUrl $siteUrl -AdminAccount $account -TenantUrl $tenantUrl
                if (-not $granted) { throw }
                $items = @(Get-DriveFilesRecursive -DriveId $driveId -ItemId $folderId)
            } elseif (Test-IsNotFoundError -Message $errMsg) {
                Write-Warning "  Niet gevonden in SharePoint (overgeslagen): $($row.TeamName) / $($row.ChannelName)"
                continue
            } else {
                throw
            }
        }

        if (-not $items -or $items.Count -eq 0) {
            if ($DryRun) { $dryRunFileCounts[$rowKey] = 0 }
            Write-Host "  Leeg: $($row.TeamName) / $($row.ChannelName)" -ForegroundColor Gray
            continue
        }

        $itemList = @($items)
        if ($DryRun) {
            $dryRunFileCounts[$rowKey] = $itemList.Count
            Write-Host "  [DRYRUN] Bestanden gevonden: $($row.TeamName) / $($row.ChannelName) ($($itemList.Count))" -ForegroundColor Cyan
            continue
        }
        $dl = Save-DriveFilesReliable -Items $itemList -DestPath $destPath
        if ($dl.FailedCount -gt 0) {
            throw "Niet alle bestanden konden gedownload worden ($($itemList.Count - $dl.FailedCount)/$($itemList.Count))."
        }

        $localCount = (Get-ChildItem -Path $destPath -Recurse -File -ErrorAction SilentlyContinue).Count
        if ($localCount -lt $itemList.Count) {
            throw "Bestandscontrole mislukt: lokaal $localCount van $($itemList.Count) bestanden aanwezig."
        }

        Write-Host "  OK [$($row.ChannelType)] $($row.TeamName) / $($row.ChannelName) ($($itemList.Count) bestanden)" `
            -ForegroundColor Green

    } catch {
        Write-Warning "  FOUT $($row.TeamName) / $($row.ChannelName) : $_"
    }
}
#endregion

#region STAP 9 - Chat exporteren (VOOR archivering)
Write-Host "`n[9/12] Chat history exporteren..." -ForegroundColor Cyan

if ($chatMethode -eq "a") {
    Write-Host "  Methode: Graph API (automatisch)`n" -ForegroundColor White

    foreach ($row in $toArchive) {
        $rowKey = New-ArchiveRowKey -TeamName $row.TeamName -ChannelName $row.ChannelName
        $groupId = $teamMapping[$row.TeamName]
        if (-not $groupId) { continue }

        $safeteam    = $row.TeamName    -replace '[\\/:*?"<>|]', '_'
        $safechannel = $row.ChannelName -replace '[\\/:*?"<>|]', '_'
        $chatPath    = Join-Path $archiveRoot $safeteam $safechannel "Chat"

        try {
            $channel = Get-TeamChannelCached -GroupId $groupId -ChannelName $row.ChannelName -Cache $teamChannelCache
            if (-not $channel) {
                Write-Warning "  Kanaal niet gevonden voor chat: $($row.ChannelName)"
                continue
            }

            $allMessages = Get-GraphPagedCollection -StartUri "https://graph.microsoft.com/v1.0/teams/$groupId/channels/$($channel.Id)/messages"

            $allData = [System.Collections.Generic.List[object]]::new()
            foreach ($msg in $allMessages) {
                $msgObj = [PSCustomObject]@{
                    Id       = $msg.id
                    Datum    = $msg.createdDateTime
                    Afzender = if ($msg.from.user.displayName) { $msg.from.user.displayName } else { "Onbekend" }
                    Email    = if ($msg.from.user.email) { $msg.from.user.email } else { "" }
                    Bericht  = if ($msg.body.content) { ($msg.body.content -replace '<[^>]+>', '') } else { "" }
                    Bijlagen = ($msg.attachments | ForEach-Object { $_.name }) -join ", "
                    Replies  = @()
                }
                try {
                    $replyItems = Get-GraphPagedCollection -StartUri "https://graph.microsoft.com/v1.0/teams/$groupId/channels/$($channel.Id)/messages/$($msg.id)/replies"
                    if ($replyItems.Count -gt 0) {
                        $msgObj.Replies = $replyItems | ForEach-Object {
                            [PSCustomObject]@{
                                Datum    = $_.createdDateTime
                                Afzender = if ($_.from.user.displayName) { $_.from.user.displayName } else { "Onbekend" }
                                Bericht  = if ($_.body.content) { ($_.body.content -replace '<[^>]+>', '') } else { "" }
                            }
                        }
                    }
                } catch { }
                $allData.Add($msgObj)
            }

            if ($DryRun) {
                $dryRunChatCounts[$rowKey] = $allMessages.Count
                Write-Host "  [DRYRUN] Chat-berichten gevonden: $($row.TeamName) / $($row.ChannelName) ($($allMessages.Count))" -ForegroundColor Cyan
                continue
            }

            $allData | ConvertTo-Json -Depth 10 |
                Out-File (Join-Path $chatPath "${safechannel}_chat.json") -Encoding UTF8

            $allData | Select-Object Datum, Afzender, Email, Bericht, Bijlagen |
                Export-Csv (Join-Path $chatPath "${safechannel}_chat.csv") -NoTypeInformation -Encoding UTF8

            Write-Host "  OK: $($row.TeamName) / $($row.ChannelName) ($($allMessages.Count) berichten)" `
                -ForegroundColor Green

        } catch {
            Write-Warning "  FOUT chat $($row.TeamName) / $($row.ChannelName) : $_"
        }
    }

} else {
    Write-Host "  Methode: Microsoft Purview eDiscovery (manueel)`n" -ForegroundColor White
    if ($DryRun) {
        Write-Host "  [DRYRUN] Purview-export wordt niet uitgevoerd; automatische chat-validatie is in deze modus niet mogelijk." -ForegroundColor Yellow
    }
    Write-Host "  1. Ga naar https://compliance.microsoft.com"
    Write-Host "  2. eDiscovery > Standard > + Create a case"
    Write-Host "     Naam: Teams Archivering $(Get-Date -Format 'yyyy')"
    Write-Host "  3. Searches > + New search > Teams chats & channels"
    Write-Host "     Selecteer deze Teams:"
    $archiveTeams | ForEach-Object { Write-Host "       - $_" -ForegroundColor Gray }
    Write-Host "  4. Save & run > wacht tot klaar (5-30 min)"
    Write-Host "  5. Actions > Export results > HTML reports"
    Write-Host "  6. Download en sla op in: $archiveRoot\[Teamnaam]\Chat\`n"
    if (-not $DryRun) {
        Write-Host "  Druk Enter zodra de export gedownload en opgeslagen is." -ForegroundColor Red
        Read-Host "  Klaar? Druk Enter"
    }
}

# Verificatie chat-mappen
Write-Host "`n  Chat-verificatie:" -ForegroundColor White
$chatOntbreekt = 0
foreach ($row in $toArchive) {
    $safeTeam  = $row.TeamName -replace '[\\/:*?"<>|]', '_'
    $safeChannel = $row.ChannelName -replace '[\\/:*?"<>|]', '_'
    $rowKey = New-ArchiveRowKey -TeamName $row.TeamName -ChannelName $row.ChannelName
    $count = if ($DryRun -and $chatMethode -eq "a") {
        if ($dryRunChatCounts.ContainsKey($rowKey)) { [int]$dryRunChatCounts[$rowKey] } else { 0 }
    } elseif ($DryRun -and $chatMethode -eq "b") {
        -1
    } else {
        (Get-ChildItem (Join-Path $archiveRoot $safeTeam $safeChannel "Chat") -File -ErrorAction SilentlyContinue).Count
    }

    if ($count -eq -1) {
        Write-Host "  [DRYRUN] Purview chat-validatie niet automatisch beschikbaar: $($row.TeamName) / $($row.ChannelName)" -ForegroundColor DarkYellow
        continue
    }

    if ($count -eq 0) {
        Write-Warning "  Geen chat-export: $($row.TeamName) / $($row.ChannelName)"
        $chatOntbreekt++
    } else {
        Write-Host "  OK: $($row.TeamName) / $($row.ChannelName) ($count bestanden)" -ForegroundColor Green
    }
}

if ($chatOntbreekt -gt 0) {
    Write-Host "`n  $chatOntbreekt Teams zonder chat-export." -ForegroundColor Yellow
    if ($DryRun) {
        Write-Host "  [DRYRUN] Doorgaan zonder archiveringswijzigingen." -ForegroundColor Yellow
    } else {
        $doorgaan = Read-Host "  Toch archiveren in M365? (j/n)"
        if ($doorgaan -ne "j") {
            Write-Host "  Gepauzeerd. Exporteer de ontbrekende chats en herstart." -ForegroundColor Yellow
            exit
        }
    }
}
#endregion

} else {
    Write-Host "`n[QUICK MODE] Step10Only actief: stappen 6 t/m 9 worden overgeslagen." -ForegroundColor Yellow
}

#region STAP 10 - Teams archiveren (na chat-export)
Write-Host "`n[10/12] Teams archiveren in Microsoft 365..." -ForegroundColor Cyan
Write-Host "  Standaard is archiveren UITGESCHAKELD in v9.0." -ForegroundColor Yellow
Write-Host "  Let op: echte archiveren/unarchiven gebeurt op TEAM-niveau." -ForegroundColor Yellow
Write-Host "  C/D gebruiken nu echte kanaal archiveren/unarchiven via Graph." -ForegroundColor Yellow
if ($ChannelFallbackToRename) {
    Write-Host "  Fallback actief: bij API-fout wordt kanaalnaam-marker gebruikt." -ForegroundColor Yellow
}
if ($DryRun) {
    Write-Host "  DRY RUN actief: er worden geen wijzigingen uitgevoerd." -ForegroundColor Yellow
}
Write-Host "  A) Team archiveren (read-only)"
Write-Host "  U) Team undo archivering (unarchive)"
Write-Host "  C) Kanaal archiveren"
Write-Host "  D) Kanaal undo archivering"
Write-Host "  N) Overslaan (standaard)"

function Get-ChannelByNameForStep10 {
    param(
        [Parameter(Mandatory = $true)][string]$GroupId,
        [Parameter(Mandatory = $true)][string]$ChannelName
    )

    $uri = "https://graph.microsoft.com/v1.0/teams/$GroupId/channels"
    do {
        $result = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        if ($result.value) {
            $found = $result.value | Where-Object { $_.displayName -eq $ChannelName } | Select-Object -First 1
            if ($found) { return $found }
        }
        $uri = $result.'@odata.nextLink'
    } while ($uri)

    return $null
}

function Get-ChannelTargetName {
    param(
        [Parameter(Mandatory = $true)][string]$CurrentName,
        [Parameter(Mandatory = $true)][string]$Mode,
        [Parameter(Mandatory = $true)][string]$Tag
    )

    $prefix = "$Tag "
    if ($Mode -eq "archive") {
        if ($CurrentName -like "$prefix*") { return $CurrentName }
        return "$prefix$CurrentName"
    }

    if ($CurrentName -like "$prefix*") {
        return $CurrentName.Substring($prefix.Length)
    }
    return $CurrentName
}

function Set-ChannelArchiveMarker {
    param(
        [Parameter(Mandatory = $true)][string]$GroupId,
        [Parameter(Mandatory = $true)][string]$ChannelName,
        [Parameter(Mandatory = $true)][string]$Mode,
        [Parameter(Mandatory = $true)][string]$Tag
    )

    if ($ChannelName -eq "General") {
        Write-Warning "  General kanaal wordt overgeslagen: $GroupId / $ChannelName"
        return
    }

    $channel = Get-ChannelByNameForStep10 -GroupId $GroupId -ChannelName $ChannelName
    if (-not $channel) {
        Write-Warning "  Kanaal niet gevonden voor soft-archive: $GroupId / $ChannelName"
        return
    }

    $currentName = [string]$channel.displayName
    $targetName = Get-ChannelTargetName -CurrentName $currentName -Mode $Mode -Tag $Tag
    if ($targetName -eq $currentName) {
        Write-Host "  Geen wijziging nodig: $currentName" -ForegroundColor Gray
        return
    }

    $body = @{ displayName = $targetName } | ConvertTo-Json
    for ($poging = 1; $poging -le 3; $poging++) {
        try {
            if ($DryRun) {
                Write-Host "  [DRYRUN] Kanaal marker: $currentName -> $targetName" -ForegroundColor Cyan
                return
            }
            Invoke-MgGraphRequest -Method PATCH `
                -Uri "https://graph.microsoft.com/v1.0/teams/$GroupId/channels/$($channel.id)" `
                -Body $body -ContentType "application/json" -ErrorAction Stop

            if ($Mode -eq "archive") {
                Write-Host "  Kanaal gemarkeerd: $currentName -> $targetName" -ForegroundColor Green
            } else {
                Write-Host "  Kanaal hersteld: $currentName -> $targetName" -ForegroundColor Green
            }
            return
        } catch {
            if ($poging -lt 3) {
                $wacht = 10 * $poging
                Start-Sleep -Seconds $wacht
            } else {
                Write-Warning "  Fout kanaalwijziging $currentName : $_"
            }
        }
    }
}

function Wait-ChannelArchiveState {
    param(
        [Parameter(Mandatory = $true)][string]$GroupId,
        [Parameter(Mandatory = $true)][string]$ChannelId,
        [Parameter(Mandatory = $true)][bool]$DesiredArchived,
        [int]$MaxChecks = 12,
        [int]$IntervalSeconds = 5
    )

    for ($i = 1; $i -le $MaxChecks; $i++) {
        try {
            $state = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$GroupId/channels/$ChannelId" -ErrorAction Stop
            if ([bool]$state.isArchived -eq $DesiredArchived) {
                return $true
            }
        } catch { }
        Start-Sleep -Seconds $IntervalSeconds
    }

    return $false
}

function Set-ChannelArchiveStateGraph {
    param(
        [Parameter(Mandatory = $true)][string]$GroupId,
        [Parameter(Mandatory = $true)][string]$ChannelName,
        [Parameter(Mandatory = $true)][string]$Mode,
        [Parameter(Mandatory = $true)][string]$Tag,
        [switch]$FallbackToRename
    )

    $channel = Get-ChannelByNameForStep10 -GroupId $GroupId -ChannelName $ChannelName
    if (-not $channel) {
        Write-Warning "  Kanaal niet gevonden: $GroupId / $ChannelName"
        return
    }

    $isArchivedNow = [bool]$channel.isArchived
    if ($Mode -eq "archive" -and $isArchivedNow) {
        Write-Host "  Kanaal al gearchiveerd: $($channel.displayName)" -ForegroundColor Gray
        return
    }
    if ($Mode -eq "undo" -and -not $isArchivedNow) {
        Write-Host "  Kanaal al actief: $($channel.displayName)" -ForegroundColor Gray
        return
    }

    $action = if ($Mode -eq "archive") { "archive" } else { "unarchive" }
    $uri = "https://graph.microsoft.com/v1.0/teams/$GroupId/channels/$($channel.id)/$action"
    $desiredArchived = ($Mode -eq "archive")

    for ($poging = 1; $poging -le 3; $poging++) {
        try {
            if ($DryRun) {
                Write-Host "  [DRYRUN] Kanaal ${action}: $($channel.displayName)" -ForegroundColor Cyan
                return
            }
            if ($Mode -eq "archive") {
                $body = @{ shouldSetSpoSiteReadOnlyForMembers = $true } | ConvertTo-Json
                Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType "application/json" -ErrorAction Stop
            } else {
                Invoke-MgGraphRequest -Method POST -Uri $uri -ContentType "application/json" -ErrorAction Stop
            }

            if (Wait-ChannelArchiveState -GroupId $GroupId -ChannelId $channel.id -DesiredArchived $desiredArchived) {
                if ($Mode -eq "archive") {
                    Write-Host "  Kanaal gearchiveerd: $($channel.displayName)" -ForegroundColor Green
                } else {
                    Write-Host "  Kanaal geunarchived: $($channel.displayName)" -ForegroundColor Green
                }
                return
            }

            throw "Kanaal status bevestiging timeout"
        } catch {
            if ($poging -lt 3) {
                Start-Sleep -Seconds (10 * $poging)
                continue
            }

            if ($FallbackToRename) {
                Write-Warning "  API kanaal $action mislukt, fallback naar naammarker: $($channel.displayName)"
                Set-ChannelArchiveMarker -GroupId $GroupId -ChannelName $ChannelName -Mode $Mode -Tag $Tag
            } else {
                Write-Warning "  Fout kanaal $action $($channel.displayName) : $_"
            }
        }
    }
}

$archiveKeuze = "interactive"
switch ($Step10Action) {
    "archive" { $archiveKeuze = "a" }
    "undo"    { $archiveKeuze = "u" }
    "skip"    { $archiveKeuze = "n" }
}

if ($ChannelAction -eq "archive") { $archiveKeuze = "c" }
if ($ChannelAction -eq "undo")    { $archiveKeuze = "d" }

if ($archiveKeuze -eq "interactive") {
    do {
        $archiveKeuze = (Read-Host "  Kies actie (a/u/c/d/n, standaard n)").Trim().ToLower()
        if ([string]::IsNullOrWhiteSpace($archiveKeuze)) { $archiveKeuze = "n" }
    } while ($archiveKeuze -notin @("a", "u", "c", "d", "n"))
} else {
    if ($ChannelAction -ne "none") {
        Write-Host "  ChannelAction toegepast: $ChannelAction" -ForegroundColor Cyan
    } else {
        Write-Host "  Step10Action toegepast: $Step10Action" -ForegroundColor Cyan
    }
}

if ($archiveKeuze -eq "a") {
    Write-Host "  Team archivering gestart (team-niveau).`n" -ForegroundColor White
    foreach ($teamName in $archiveTeams) {
        $groupId = $teamMapping[$teamName]
        if (-not $groupId) { continue }
        if ($DryRun) {
            Write-Host "  [DRYRUN] Team archiveren: $teamName" -ForegroundColor Cyan
            continue
        }
        $gelukt = $false
        for ($poging = 1; $poging -le 3 -and -not $gelukt; $poging++) {
            try {
                Invoke-MgGraphRequest -Method POST `
                    -Uri "https://graph.microsoft.com/v1.0/teams/$groupId/archive" `
                    -Body (@{ shouldSetSpoSiteReadOnlyForMembers = $true } | ConvertTo-Json) `
                    -ContentType "application/json"
                Write-Host "  Gearchiveerd: $teamName" -ForegroundColor Green
                $gelukt = $true
            } catch {
                if ($poging -lt 3) {
                    $wacht = 30 * $poging
                    Write-Host "  Wachten ${wacht}s voor retry ($poging/3): $teamName..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $wacht
                } else {
                    Write-Warning "  Fout archivering $teamName : $_"
                }
            }
        }
        if ($gelukt) { Start-Sleep -Seconds 2 }
    }
} elseif ($archiveKeuze -eq "u") {
    Write-Host "  Team undo archivering gestart (team-niveau).`n" -ForegroundColor White
    foreach ($teamName in $archiveTeams) {
        $groupId = $teamMapping[$teamName]
        if (-not $groupId) { continue }
        if ($DryRun) {
            Write-Host "  [DRYRUN] Team unarchive: $teamName" -ForegroundColor Cyan
            continue
        }
        $gelukt = $false
        for ($poging = 1; $poging -le 3 -and -not $gelukt; $poging++) {
            try {
                Invoke-MgGraphRequest -Method POST `
                    -Uri "https://graph.microsoft.com/v1.0/teams/$groupId/unarchive" `
                    -ContentType "application/json"
                Write-Host "  Unarchived: $teamName" -ForegroundColor Green
                $gelukt = $true
            } catch {
                if ($poging -lt 3) {
                    $wacht = 30 * $poging
                    Write-Host "  Wachten ${wacht}s voor retry ($poging/3): $teamName..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $wacht
                } else {
                    Write-Warning "  Fout undo archivering $teamName : $_"
                }
            }
        }
        if ($gelukt) { Start-Sleep -Seconds 2 }
    }
} elseif ($archiveKeuze -in @("c", "d")) {
    $mode = if ($archiveKeuze -eq "c") { "archive" } else { "undo" }
    Write-Host "  Kanaal API mode: $mode.`n" -ForegroundColor White
    foreach ($row in $toArchive) {
        $groupId = $teamMapping[$row.TeamName]
        if (-not $groupId) { continue }
        Set-ChannelArchiveStateGraph -GroupId $groupId -ChannelName $row.ChannelName -Mode $mode `
            -Tag $ChannelArchiveTag -FallbackToRename:$ChannelFallbackToRename
    }
} else {
    Write-Host "  Archiveren/undo overgeslagen. Teams-status blijft ongewijzigd." -ForegroundColor Yellow
}
#endregion

#region STAP 11 - Rapport
Write-Host "`n[11/12] Verificatierapport genereren..." -ForegroundColor Cyan

$report = [System.Collections.Generic.List[object]]::new()

foreach ($row in $toArchive) {
    $safeteam    = $row.TeamName    -replace '[\\/:*?"<>|]', '_'
    $safechannel = $row.ChannelName -replace '[\\/:*?"<>|]', '_'
    $groupId     = $teamMapping[$row.TeamName]
    $rowKey = New-ArchiveRowKey -TeamName $row.TeamName -ChannelName $row.ChannelName

    $fileCount = if ($DryRun) {
        if ($dryRunFileCounts.ContainsKey($rowKey)) { [int]$dryRunFileCounts[$rowKey] } else { 0 }
    } else {
        try {
            (Get-ChildItem ([System.IO.Path]::Combine($archiveRoot, $safeteam, $safechannel, "Files")) `
                -Recurse -File -ErrorAction SilentlyContinue).Count
        } catch { 0 }
    }
    $chatCount = if ($DryRun -and $chatMethode -eq "a") {
        if ($dryRunChatCounts.ContainsKey($rowKey)) { [int]$dryRunChatCounts[$rowKey] } else { 0 }
    } elseif ($DryRun -and $chatMethode -eq "b") {
        0
    } else {
        try {
            (Get-ChildItem ([System.IO.Path]::Combine($archiveRoot, $safeteam, $safechannel, "Chat")) `
                -File -ErrorAction SilentlyContinue).Count
        } catch { 0 }
    }

    $m365Status = if ($groupId) {
        try {
            $r = Invoke-MgGraphRequest -Method GET `
                     -Uri "https://graph.microsoft.com/v1.0/teams/$groupId"
            if ($r.isArchived) { "Gearchiveerd" } else { "Actief" }
        } catch { "Onbekend" }
    } else { "Niet gevonden" }

    $report.Add([PSCustomObject]@{
        Team          = $row.TeamName
        Kanaal        = $row.ChannelName
        Type          = $row.ChannelType
        AantalFiles   = $fileCount
        ChatBestanden = $chatCount
        ChatOK        = if ($chatCount -gt 0) { "Ja" } else { "NEE" }
        M365Status    = $m365Status
        Datum         = (Get-Date -Format "dd/MM/yyyy HH:mm")
    })
}

$rapportPad = [System.IO.Path]::Combine($archiveRoot, "Teams_Archivering_Rapport_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx")
if (-not (Test-Path $archiveRoot -ErrorAction SilentlyContinue)) {
    $rapportPad = [System.IO.Path]::Combine(
        [System.IO.Path]::GetTempPath(),
        "Teams_Archivering_Rapport_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"
    )
    Write-Warning "  Archief-locatie niet bereikbaar. Rapport wordt opgeslagen in: $rapportPad"
}
$report | Export-Excel -Path $rapportPad `
    -AutoSize -BoldTopRow -FreezeTopRow `
    -TableName "ArchivRapport" -WorksheetName "Archivering" `
    -ConditionalText $(
        New-ConditionalText "NEE"          -Range "G:G" -BackgroundColor "#FFD7D7" -ConditionalTextColor "#CC0000"
        New-ConditionalText "Gearchiveerd" -Range "H:H" -BackgroundColor "#D4EDDA" -ConditionalTextColor "#155724"
        New-ConditionalText "Actief"       -Range "H:H" -BackgroundColor "#FFF3CD" -ConditionalTextColor "#856404"
    )

Write-Host "  Rapport: $rapportPad" -ForegroundColor Green
#endregion

#region STAP 12 - Opruimen
Write-Host "`n[12/12] Opruimen..." -ForegroundColor Cyan
# Geen tijdelijke app meer om op te ruimen. Alleen de Graph-sessie sluiten als dit
# script haar zelf opende (de herstarte sessie is sowieso eigen).
Disconnect-M365Graph $graphConn
Write-Host "  Graph-sessie afgesloten." -ForegroundColor Yellow

# Omgevingsvariabele opruimen
$env:TEAMS_ARCHIVER_HERSTART = $null

Write-Host "`n===== ARCHIVERING VOLTOOID =====" -ForegroundColor Cyan
Write-Host "  Rapport : $rapportPad"            -ForegroundColor Green
Write-Host "  Archief : $archiveRoot"            -ForegroundColor Green
Write-Host "================================`n"  -ForegroundColor Cyan
#endregion