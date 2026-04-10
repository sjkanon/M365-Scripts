# ============================================================
# Vias Teams Archivering - Volledig Automatisch Script v8.6
# PowerShell 7+ vereist | Uitvoeren als Global Admin
# ============================================================

#region ZELFHERSTART - Modules opkuisen en sessie hernieuwen
# Dit blok zorgt ervoor dat het script zichzelf herstart in een schone sessie
# nadat conflicterende modules zijn verwijderd. Zonder herstart blijven
# oude module-versies actief in het geheugen en crashen alle Graph-calls.

$herstart = $env:VIAS_ARCHIVER_HERSTART

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

    foreach ($mod in @("MicrosoftTeams","PnP.PowerShell","ImportExcel")) {
        if (-not (Get-Module -ListAvailable $mod -ErrorAction SilentlyContinue)) {
            Write-Host "  Installeren: $mod..." -ForegroundColor Yellow
            Install-Module $mod -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        Write-Host "  OK: $mod" -ForegroundColor Green
    }

    Write-Host "`n  Modules geinstalleerd. Script herstart in schone sessie...`n" -ForegroundColor Cyan

    # Herstart het script in een nieuwe pwsh-sessie met de herstart-vlag
    # Alle configuratievariabelen worden via omgevingsvariabelen doorgegeven
    $env:VIAS_ARCHIVER_HERSTART = "1"
    $pwshPath = (Get-Command pwsh).Source
    $restartArgs = if ($IsWindows) {
        @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $MyInvocation.MyCommand.Path)
    } else {
        @("-NoProfile", "-File", $MyInvocation.MyCommand.Path)
    }
    & $pwshPath @restartArgs
    exit
}

# Vanaf hier: we zitten in de hergestarte schone sessie
Write-Host "`n  Schone sessie actief. Modules worden geladen..." -ForegroundColor Green

Import-Module Microsoft.Graph.Authentication,
              Microsoft.Graph.Applications,
              Microsoft.Graph.Groups -ErrorAction Stop

Write-Host "  Graph modules geladen." -ForegroundColor Green

# Cross-platform tijdelijke map
$tempDir = [System.IO.Path]::GetTempPath()
$cleanupEventName = "ViasArchiverCleanup"

# Tijdelijke app-tracking
$isTempApp = $false
$app = $null
$sp = $null

function Remove-TempArchiverApp {
    param(
        [Parameter(Mandatory = $false)]$App,
        [Parameter(Mandatory = $false)]$Sp,
        [switch]$Silent
    )

    if (-not $App) { return }

    try {
        if ($Sp -and $Sp.Id) {
            Remove-MgServicePrincipal -ServicePrincipalId $Sp.Id -ErrorAction SilentlyContinue
        }
        if ($App.Id) {
            Remove-MgApplication -ApplicationId $App.Id -ErrorAction SilentlyContinue
        }
        if (-not $Silent) {
            Write-Host "  Tijdelijke app verwijderd." -ForegroundColor Yellow
        }
    } catch {
        if (-not $Silent) {
            Write-Warning "  Kon tijdelijke app niet volledig verwijderen: $_"
        }
    }
}

function Register-TempAppCleanupEvent {
    param(
        [Parameter(Mandatory = $true)][string]$EventName,
        [Parameter(Mandatory = $true)]$App,
        [Parameter(Mandatory = $true)]$Sp
    )

    Unregister-Event -SourceIdentifier $EventName -ErrorAction SilentlyContinue
    Register-EngineEvent -SourceIdentifier PowerShell.Exiting -SupportEvent -Action {
        if ($using:App -and $using:App.Id) {
            try {
                if ($using:Sp -and $using:Sp.Id) {
                    Remove-MgServicePrincipal -ServicePrincipalId $using:Sp.Id -ErrorAction SilentlyContinue
                }
                Remove-MgApplication -ApplicationId $using:App.Id -ErrorAction SilentlyContinue
            } catch { }
        }
    } | Out-Null
}
#endregion

#region CONFIGURATIE - Interactief opvragen
Clear-Host
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Vias Teams Archivering - Setup Wizard v8.6" -ForegroundColor Cyan
Write-Host "============================================`n" -ForegroundColor Cyan

# Excel-bestand
Write-Host "Stap 1/5 - Excel-bestand met de Teams-lijst" -ForegroundColor Yellow
$xlStandaard = if ($IsWindows) { "C:\Temp\Vias_Teams_Channels_JFG.xlsx" } else { Join-Path $HOME "Downloads" "Vias_Teams_Channels_JFG.xlsx" }
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
$archiveStandaard = if ($IsWindows) { "N:\Archives\Teams" } else { Join-Path $HOME "Documents" "Vias_Teams_Archive" }
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
Write-Host "Stap 3/5 - Tenant ID van de Vias Microsoft 365 omgeving" -ForegroundColor Yellow
Write-Host "  Vind je via: https://entra.microsoft.com > Microsoft Entra ID > Overview"
Write-Host "  Formaat: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`n"
do {
    $tenantId = (Read-Host "  Tenant ID").Trim()
    if ($tenantId -notmatch '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$') {
        Write-Host "  Ongeldig formaat. Probeer opnieuw.`n" -ForegroundColor Red
    }
} while ($tenantId -notmatch '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$')
Write-Host "  OK: $tenantId`n" -ForegroundColor Green

# SharePoint URL
Write-Host "Stap 4/5 - SharePoint URL van de Vias tenant" -ForegroundColor Yellow
Write-Host "  Formaat: https://naam.sharepoint.com  (geen slash aan het einde)`n"
do {
    $tenantUrl = (Read-Host "  SharePoint URL").Trim().TrimEnd('/')
    if ($tenantUrl -notmatch '^https://[a-zA-Z0-9-]+\.sharepoint\.com$') {
        Write-Host "  Ongeldig formaat. Voorbeeld: https://bivv.sharepoint.com`n" -ForegroundColor Red
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

#region STAP 1 - Login Global Admin
Write-Host "`n[1/12] Inloggen als Global Admin..." -ForegroundColor Cyan
Write-Host "  Je krijgt een code + link." -ForegroundColor Yellow
Write-Host "  Log in met het VIAS GLOBAL ADMIN account.`n" -ForegroundColor Yellow

Disconnect-MgGraph -ErrorAction SilentlyContinue

try {
    Connect-MgGraph `
        -TenantId $tenantId `
        -Scopes "Application.ReadWrite.All","AppRoleAssignment.ReadWrite.All" `
    -ContextScope Process `
        -UseDeviceAuthentication -NoWelcome -ErrorAction Stop
} catch {
    Write-Host "`n  FOUT bij inloggen: $_" -ForegroundColor Red
    Write-Host "  Controleer:" -ForegroundColor Yellow
    Write-Host "   - Tenant ID correct? ($tenantId)"
    Write-Host "   - Log je in met het VIAS admin account (niet BraveHub)?"
    Write-Host "   - Is MFA ingesteld op dit account?"
    exit
}

$account = (Get-MgContext).Account
if (-not $account) {
    Write-Host "  FOUT: Inloggen mislukt, geen account gevonden." -ForegroundColor Red; exit
}
Write-Host "  Ingelogd als: $account" -ForegroundColor Green
#endregion

#region STAP 2 - Entra App aanmaken of hergebruiken
Write-Host "`n[2/12] Tijdelijke Entra app registreren..." -ForegroundColor Cyan

$appName = "Temp-Vias-Teams-Archiver-$(Get-Date -Format 'yyyyMMddHHmmss')"
$isTempApp = $true

try {
    $app = New-MgApplication -DisplayName $appName `
        -PublicClient @{ RedirectUris = @("http://localhost") } `
        -IsFallbackPublicClient -ErrorAction Stop
    Write-Host "  Tijdelijke app aangemaakt: $($app.AppId)" -ForegroundColor Green
} catch {
    Write-Host "  FOUT bij aanmaken tijdelijke app: $_" -ForegroundColor Red; exit
}

try {
    $sp = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop
    Write-Host "  Service Principal aangemaakt." -ForegroundColor Green
    Register-TempAppCleanupEvent -EventName $cleanupEventName -App $app -Sp $sp
} catch {
    Write-Host "  FOUT bij aanmaken Service Principal: $_" -ForegroundColor Red
    Remove-TempArchiverApp -App $app -Sp $sp
    exit
}

# ClientId ophalen en valideren
$clientId = $app.AppId
if ([string]::IsNullOrWhiteSpace($clientId)) {
    Write-Host "  FOUT: ClientId is leeg na aanmaken app." -ForegroundColor Red; exit
}

# Opslaan als fallback
$clientId | Out-File (Join-Path $tempDir "vias_archiver_clientid.txt") -Force
Write-Host "  Client ID: $clientId" -ForegroundColor Green
#endregion

#region STAP 3 - Permissies instellen en admin consent geven
Write-Host "`n[3/12] API Permissies instellen en admin consent geven..." -ForegroundColor Cyan

$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" |
           Select-Object -First 1
$spSp    = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0ff1-ce00-000000000000'" |
           Select-Object -First 1

function Grant-DelegatedPermission {
    param($ResourceSp, [string[]]$Scopes, $ResourceName)
    $validScopes = @()
    foreach ($scopeName in $Scopes) {
        $permDef = $ResourceSp.Oauth2PermissionScopes |
                   Where-Object { $_.Value -eq $scopeName } | Select-Object -First 1
        if (-not $permDef) {
            Write-Warning "    Scope '$scopeName' niet gevonden op $ResourceName"
            continue
        }
        $validScopes += $scopeName
    }

    if (-not $validScopes -or $validScopes.Count -eq 0) {
        return
    }

    $grant = Get-MgOauth2PermissionGrant `
        -Filter "clientId eq '$($sp.Id)' and resourceId eq '$($ResourceSp.Id)' and consentType eq 'AllPrincipals'" `
        -ErrorAction SilentlyContinue |
        Select-Object -First 1

    $existingScopes = @()
    if ($grant -and $grant.Scope) {
        $existingScopes = $grant.Scope -split ' '
    }

    $missingScopes = $validScopes | Where-Object { $_ -notin $existingScopes }
    if (-not $missingScopes -or $missingScopes.Count -eq 0) {
        Write-Host "    Alle scopes al aanwezig op $ResourceName" -ForegroundColor Gray
        return
    }

    $mergedScopes = ($existingScopes + $validScopes | Select-Object -Unique) -join ' '

    if ($grant) {
        Update-MgOauth2PermissionGrant `
            -OAuth2PermissionGrantId $grant.Id `
            -Scope $mergedScopes `
            -ErrorAction Stop | Out-Null
        Write-Host "    Bijgewerkt op ${ResourceName}: $($missingScopes -join ', ')" -ForegroundColor Green
    } else {
        New-MgOauth2PermissionGrant `
            -ClientId    $sp.Id `
            -ResourceId  $ResourceSp.Id `
            -Scope       $mergedScopes `
            -ConsentType "AllPrincipals" `
            -ErrorAction Stop | Out-Null
        Write-Host "    Toegekend op ${ResourceName}: $($missingScopes -join ', ')" -ForegroundColor Green
    }
}

try {
    Write-Host "  Microsoft Graph permissies..." -ForegroundColor White
    Grant-DelegatedPermission -ResourceSp $graphSp -ResourceName "Graph" -Scopes @(
        "Group.ReadWrite.All","Sites.Read.All","Files.ReadWrite.All",
        "ChannelMessage.Read.All","TeamSettings.ReadWrite.All","TeamMember.Read.All"
    )
    Write-Host "  SharePoint permissies..." -ForegroundColor White
    Grant-DelegatedPermission -ResourceSp $spSp -ResourceName "SharePoint" -Scopes @(
        "AllSites.FullControl"
    )
    Write-Host "  Alle permissies ingesteld." -ForegroundColor Green
} catch {
    Write-Host "  FOUT bij permissies: $_" -ForegroundColor Red
    if ($isTempApp) { Remove-TempArchiverApp -App $app -Sp $sp }
    exit
}
#endregion

#region STAP 4 - Opnieuw inloggen met volledige permissies
Write-Host "`n[4/12] Opnieuw inloggen met volledige permissies..." -ForegroundColor Cyan
Write-Host "  Gebruik opnieuw het VIAS ADMIN account.`n" -ForegroundColor Yellow

Disconnect-MgGraph -ErrorAction SilentlyContinue

try {
    Connect-MgGraph `
        -ClientId $clientId `
        -TenantId $tenantId `
        -Scopes "Group.ReadWrite.All","Sites.Read.All","Files.ReadWrite.All",
                "TeamSettings.ReadWrite.All","TeamMember.Read.All","ChannelMessage.Read.All" `
        -ContextScope Process `
        -UseDeviceAuthentication -NoWelcome -ErrorAction Stop
} catch {
    Write-Host "  FOUT bij tweede login: $_" -ForegroundColor Red
    if ($isTempApp) { Remove-TempArchiverApp -App $app -Sp $sp }
    exit
}

Import-Module MicrosoftTeams -ErrorAction Stop

try {
    Connect-MicrosoftTeams -TenantId $tenantId -ErrorAction Stop
} catch {
    Write-Host "  FOUT bij Teams-verbinding: $_" -ForegroundColor Red
    if ($isTempApp) { Remove-TempArchiverApp -App $app -Sp $sp }
    exit
}

Write-Host "  Verbonden als: $((Get-MgContext).Account)" -ForegroundColor Green
#endregion

#region STAP 5 - Excel inlezen en Teams-IDs ophalen
Write-Host "`n[5/12] Excel inlezen en Teams-IDs ophalen..." -ForegroundColor Cyan

Import-Module ImportExcel -ErrorAction Stop

$data      = Import-Excel -Path $xlPath -WorksheetName "Teams channels Vias"
$toArchive = $data | Where-Object { $_.Archive -eq "Archive" }
Write-Host "  $($toArchive.Count) kanalen gevonden." -ForegroundColor Green

$archiveTeams = $toArchive | Select-Object -ExpandProperty TeamName -Unique
$teamMapping  = @{}

foreach ($teamName in $archiveTeams) {
    $team = Get-Team -DisplayName $teamName -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($team) {
        $teamMapping[$teamName] = [string]$team.GroupId
        Write-Host "  OK: $teamName" -ForegroundColor Green
    } else {
        Write-Warning "  Niet gevonden: $teamName"
    }
}

$teamMapping | ConvertTo-Json | Out-File (Join-Path $tempDir "vias_team_mapping.json") -Force
Write-Host "  $($teamMapping.Count) van de $($archiveTeams.Count) Teams gevonden." -ForegroundColor Green
#endregion

#region STAP 6 - Mapstructuur
Write-Host "`n[6/12] Mapstructuur aanmaken..." -ForegroundColor Cyan
foreach ($teamName in $archiveTeams) {
    $safe = $teamName -replace '[\\/:*?"<>|]', '_'
    New-Item -ItemType Directory -Path (Join-Path $archiveRoot $safe "Files")   -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $archiveRoot $safe "Chat")    -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $archiveRoot $safe "Members") -Force | Out-Null
}
Write-Host "  Mappen aangemaakt." -ForegroundColor Green
#endregion

#region STAP 7 - Ledenlijsten
Write-Host "`n[7/12] Ledenlijsten exporteren..." -ForegroundColor Cyan
foreach ($teamName in $archiveTeams) {
    $groupId = $teamMapping[$teamName]
    if (-not $groupId) { continue }
    $safe = $teamName -replace '[\\/:*?"<>|]', '_'
    try {
        Get-TeamUser -GroupId $groupId |
            Select-Object Name, User, Role |
            Export-Csv (Join-Path $archiveRoot $safe "Members" "members.csv") -NoTypeInformation -Encoding UTF8
        Write-Host "  OK: $teamName" -ForegroundColor Green
    } catch {
        Write-Warning "  Fout ledenlijst $teamName : $_"
    }
}
#endregion

#region STAP 8 - Bestanden exporteren
Write-Host "`n[8/12] Bestanden exporteren (SharePoint -> $archiveRoot)..." -ForegroundColor Cyan

Import-Module PnP.PowerShell -ErrorAction Stop

function Test-IsAccessDeniedError {
    param([string]$Message)
    return $Message -match "Access denied|Unauthorized|Forbidden|Status:\s*401|Status:\s*403|Insufficient privileges"
}

function Test-IsNotFoundError {
    param([string]$Message)
    return $Message -match "Status:\s*404|NotFound|Item does not exist|bestaat niet"
}

function Resolve-PnPFolderLocationFromFilesFolder {
    param([Parameter(Mandatory = $true)][string]$WebUrl)

    if ([string]::IsNullOrWhiteSpace($WebUrl)) { return $null }

    $u = [System.Uri]$WebUrl
    $decodedPath = [System.Uri]::UnescapeDataString($u.AbsolutePath)

    if ($decodedPath -notmatch '^(?<site>.+?)/Shared Documents(?:/(?<tail>.*))?$') {
        return $null
    }

    $sitePath = $Matches.site
    $tail = $Matches.tail
    $siteUrl = "$($u.Scheme)://$($u.Host)$sitePath"
    $folder = if ([string]::IsNullOrWhiteSpace($tail)) { "Shared Documents" } else { "Shared Documents/$tail" }

    return [PSCustomObject]@{
        SiteUrl = $siteUrl
        Folder  = $folder
    }
}

function Get-TeamChannelCached {
    param(
        [Parameter(Mandatory = $true)][string]$GroupId,
        [Parameter(Mandatory = $true)][string]$ChannelName,
        [Parameter(Mandatory = $true)][hashtable]$Cache
    )

    if (-not $Cache.ContainsKey($GroupId)) {
        $Cache[$GroupId] = @(Get-TeamChannel -GroupId $GroupId -ErrorAction SilentlyContinue)
    }

    $channel = $Cache[$GroupId] | Where-Object { $_.DisplayName -eq $ChannelName } | Select-Object -First 1
    if ($channel) { return $channel }

    # Fallback op Graph-lijst als Teams-module niets teruggeeft
    $uri = "https://graph.microsoft.com/v1.0/teams/$GroupId/channels"
    $allChannels = [System.Collections.Generic.List[object]]::new()
    do {
        $result = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue
        if ($result.value) { $allChannels.AddRange($result.value) }
        $uri = $result.'@odata.nextLink'
    } while ($uri)

    return $allChannels | Where-Object { $_.displayName -eq $ChannelName } | Select-Object -First 1
}

function Ensure-HigherRights {
    param(
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$ClientId,
        [Parameter(Mandatory = $true)]$GraphSp
    )

    Write-Host "  Hogere rechten nodig gedetecteerd. Extra Graph-rechten worden toegekend..." -ForegroundColor Yellow

    Grant-DelegatedPermission -ResourceSp $GraphSp -ResourceName "Graph" -Scopes @(
        "Sites.ReadWrite.All",
        "Sites.FullControl.All"
    )

    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Connect-MgGraph `
        -ClientId $ClientId `
        -TenantId $TenantId `
        -Scopes "Group.ReadWrite.All","Sites.Read.All","Sites.ReadWrite.All","Sites.FullControl.All","Files.ReadWrite.All",
                "TeamSettings.ReadWrite.All","TeamMember.Read.All","ChannelMessage.Read.All" `
        -ContextScope Process `
        -UseDeviceAuthentication -NoWelcome -ErrorAction Stop

    Write-Host "  Hogere rechten toegekend en nieuwe Graph-sessie actief." -ForegroundColor Green
}

function Download-PnPFilesReliable {
    param(
        [Parameter(Mandatory = $true)][array]$Items,
        [Parameter(Mandatory = $true)][string]$DestPath,
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$ClientId,
        [int]$MaxRetriesPerFile = 3
    )

    $remaining = [System.Collections.Generic.List[object]]::new()
    if ($Items) { $remaining.AddRange(@($Items)) }

    for ($pass = 1; $pass -le 2 -and $remaining.Count -gt 0; $pass++) {
        if ($pass -gt 1) {
            # Nieuwe PnP-verbinding voor hardnekkige gevallen
            Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
        }

        $nextRound = [System.Collections.Generic.List[object]]::new()
        foreach ($item in $remaining) {
            $ok = $false
            for ($attempt = 1; $attempt -le $MaxRetriesPerFile; $attempt++) {
                try {
                    Get-PnPFile -Url $item.ServerRelativeUrl `
                        -Path $DestPath -Filename $item.Name -AsFile -Force -ErrorAction Stop
                    $ok = $true
                    break
                } catch {
                    if ($attempt -lt $MaxRetriesPerFile) {
                        Start-Sleep -Seconds ([Math]::Min(5, $attempt))
                    }
                }
            }

            if (-not $ok) {
                $nextRound.Add($item)
            }
        }

        $remaining = $nextRound
    }

    return [PSCustomObject]@{
        FailedCount = $remaining.Count
        FailedItems = $remaining
    }
}

$higherRightsGranted = $false
$teamChannelCache = @{}

# Laatste check op ClientId voor gebruik in PnP
if ([string]::IsNullOrWhiteSpace($clientId)) {
    $clientId = (Get-Content (Join-Path $tempDir "vias_archiver_clientid.txt") -Raw -ErrorAction SilentlyContinue).Trim()
    if ([string]::IsNullOrWhiteSpace($clientId)) {
        Write-Host "  FOUT: ClientId is null. Herstart het script volledig." -ForegroundColor Red; exit
    }
    Write-Host "  ClientId hersteld: $clientId" -ForegroundColor Yellow
}

$currentSiteUrl = $null

foreach ($row in $toArchive) {
    $groupId = $teamMapping[$row.TeamName]
    if (-not $groupId) {
        Write-Warning "  Geen GroupId voor: $($row.TeamName) — overgeslagen"
        continue
    }

    $safeteam    = $row.TeamName    -replace '[\\/:*?"<>|]', '_'
    $safechannel = $row.ChannelName -replace '[\\/:*?"<>|]', '_'
    $destPath    = Join-Path $archiveRoot $safeteam "Files" $safechannel
    New-Item -ItemType Directory -Path $destPath -Force | Out-Null

    try {
        $channel = Get-TeamChannelCached -GroupId $groupId -ChannelName $row.ChannelName -Cache $teamChannelCache
        if (-not $channel) {
            Write-Warning "  Kanaal niet gevonden: $($row.ChannelName)"
            continue
        }

        $ff = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/teams/$groupId/channels/$($channel.Id)/filesFolder" `
            -ErrorAction Stop
        $loc = Resolve-PnPFolderLocationFromFilesFolder -WebUrl $ff.webUrl
        if (-not $loc) {
            Write-Warning "  FilesFolder kon niet vertaald worden: $($row.TeamName) / $($row.ChannelName)"
            continue
        }

        $siteUrl = $loc.SiteUrl
        $folder = $loc.Folder

        if ($siteUrl -ne $currentSiteUrl) {
            Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $clientId
            $currentSiteUrl = $siteUrl
        }

        try {
            $items = Get-PnPFolderItem -FolderSiteRelativeUrl $folder -ItemType File -Recursive -ErrorAction Stop
        } catch {
            $errMsg = $_.Exception.Message
            if ((Test-IsAccessDeniedError -Message $errMsg) -and -not $higherRightsGranted) {
                Ensure-HigherRights -TenantId $tenantId -ClientId $clientId -GraphSp $graphSp
                $higherRightsGranted = $true
                Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $clientId
                $items = Get-PnPFolderItem -FolderSiteRelativeUrl $folder -ItemType File -Recursive -ErrorAction Stop
            } elseif (Test-IsNotFoundError -Message $errMsg) {
                Write-Warning "  Niet gevonden in SharePoint (overgeslagen): $($row.TeamName) / $($row.ChannelName)"
                continue
            } else {
                throw
            }
        }

        if (-not $items -or $items.Count -eq 0) {
            Write-Host "  Leeg: $($row.TeamName) / $($row.ChannelName)" -ForegroundColor Gray
            continue
        }

        $itemList = @($items)
        $dl = Download-PnPFilesReliable -Items $itemList -DestPath $destPath -SiteUrl $siteUrl -ClientId $clientId
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

function Invoke-GraphRequestWithRetry {
    param(
        [ValidateSet("GET","POST","PATCH","DELETE")]
        [string]$Method,
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [int]$MaxRetries = 5
    )

    $attempt = 0
    while ($true) {
        try {
            return Invoke-MgGraphRequest -Method $Method -Uri $Uri -ErrorAction Stop
        } catch {
            $attempt++
            $message = $_.Exception.Message
            $isThrottle = $message -match "Status:\s*429|Too Many Requests|throttl"
            $isTransient = $message -match "Status:\s*5\d\d|timeout|temporar"

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
            $all.AddRange($result.value)
        }
        $uri = $result.'@odata.nextLink'
    } while ($uri)

    return $all
}

if ($chatMethode -eq "a") {
    Write-Host "  Methode: Graph API (automatisch)`n" -ForegroundColor White

    foreach ($row in $toArchive) {
        $groupId = $teamMapping[$row.TeamName]
        if (-not $groupId) { continue }

        $safeteam    = $row.TeamName    -replace '[\\/:*?"<>|]', '_'
        $safechannel = $row.ChannelName -replace '[\\/:*?"<>|]', '_'
        $chatPath    = Join-Path $archiveRoot $safeteam "Chat"

        try {
            $channel = Get-MgTeamChannel -TeamId $groupId |
                       Where-Object { $_.DisplayName -eq $row.ChannelName } |
                       Select-Object -First 1
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
    Write-Host "  1. Ga naar https://compliance.microsoft.com"
    Write-Host "  2. eDiscovery > Standard > + Create a case"
    Write-Host "     Naam: Vias Teams Archivering $(Get-Date -Format 'yyyy')"
    Write-Host "  3. Searches > + New search > Teams chats & channels"
    Write-Host "     Selecteer deze Teams:"
    $archiveTeams | ForEach-Object { Write-Host "       - $_" -ForegroundColor Gray }
    Write-Host "  4. Save & run > wacht tot klaar (5-30 min)"
    Write-Host "  5. Actions > Export results > HTML reports"
    Write-Host "  6. Download en sla op in: $archiveRoot\[Teamnaam]\Chat\`n"
    Write-Host "  Druk Enter zodra de export gedownload en opgeslagen is." -ForegroundColor Red
    Read-Host "  Klaar? Druk Enter"
}

# Verificatie chat-mappen
Write-Host "`n  Chat-verificatie:" -ForegroundColor White
$chatOntbreekt = 0
foreach ($teamName in $archiveTeams) {
    $safe  = $teamName -replace '[\\/:*?"<>|]', '_'
    $count = (Get-ChildItem (Join-Path $archiveRoot $safe "Chat") -File -ErrorAction SilentlyContinue).Count
    if ($count -eq 0) {
        Write-Warning "  Geen chat-export: $teamName"
        $chatOntbreekt++
    } else {
        Write-Host "  OK: $teamName ($count bestanden)" -ForegroundColor Green
    }
}

if ($chatOntbreekt -gt 0) {
    Write-Host "`n  $chatOntbreekt Teams zonder chat-export." -ForegroundColor Yellow
    $doorgaan = Read-Host "  Toch archiveren in M365? (j/n)"
    if ($doorgaan -ne "j") {
        Write-Host "  Gepauzeerd. Exporteer de ontbrekende chats en herstart." -ForegroundColor Yellow
        exit
    }
}
#endregion

#region STAP 10 - Teams archiveren (na chat-export)
Write-Host "`n[10/12] Teams archiveren in Microsoft 365..." -ForegroundColor Cyan
Write-Host "  Standaard is archiveren UITGESCHAKELD in v8.6." -ForegroundColor Yellow
$archiveNu = Read-Host "  Wil je NU toch archiveren? (j/n, standaard n)"
if ($archiveNu -eq "j") {
    Write-Host "  Chat-export voltooid. Teams worden nu read-only gemaakt.`n" -ForegroundColor White
    foreach ($teamName in $archiveTeams) {
        $groupId = $teamMapping[$teamName]
        if (-not $groupId) { continue }
        try {
            Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/teams/$groupId/archive" `
                -Body (@{ shouldSetSpoSiteReadOnlyForMembers = $true } | ConvertTo-Json) `
                -ContentType "application/json"
            Write-Host "  Gearchiveerd: $teamName" -ForegroundColor Green
            Start-Sleep -Seconds 2
        } catch {
            Write-Warning "  Fout archivering $teamName : $_"
        }
    }
} else {
    Write-Host "  Archiveren overgeslagen. Teams blijven actief." -ForegroundColor Yellow
}
#endregion

#region STAP 11 - Rapport
Write-Host "`n[11/12] Verificatierapport genereren..." -ForegroundColor Cyan

$report = [System.Collections.Generic.List[object]]::new()

foreach ($row in $toArchive) {
    $safeteam    = $row.TeamName    -replace '[\\/:*?"<>|]', '_'
    $safechannel = $row.ChannelName -replace '[\\/:*?"<>|]', '_'
    $groupId     = $teamMapping[$row.TeamName]

    $fileCount = (Get-ChildItem (Join-Path $archiveRoot $safeteam "Files" $safechannel) `
                     -Recurse -File -ErrorAction SilentlyContinue).Count
    $chatCount = (Get-ChildItem (Join-Path $archiveRoot $safeteam "Chat") `
                     -File -ErrorAction SilentlyContinue).Count

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

$rapportPad = Join-Path $archiveRoot "Vias_Archivering_Rapport_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"
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
if ($isTempApp) {
    Remove-TempArchiverApp -App $app -Sp $sp
} else {
    Write-Host "  Niet-tijdelijke app behouden. Client ID: $clientId" -ForegroundColor Gray
}
Unregister-Event -SourceIdentifier $cleanupEventName -ErrorAction SilentlyContinue
Remove-Item (Join-Path $tempDir "vias_archiver_clientid.txt") -ErrorAction SilentlyContinue
Write-Host "  Tijdelijke bestanden verwijderd." -ForegroundColor Yellow

# Omgevingsvariabele opruimen
$env:VIAS_ARCHIVER_HERSTART = $null

Write-Host "`n===== ARCHIVERING VOLTOOID =====" -ForegroundColor Cyan
Write-Host "  Rapport : $rapportPad"            -ForegroundColor Green
Write-Host "  Archief : $archiveRoot"            -ForegroundColor Green
Write-Host "================================`n"  -ForegroundColor Cyan
#endregion