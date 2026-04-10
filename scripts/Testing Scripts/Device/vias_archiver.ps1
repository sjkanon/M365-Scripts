# ============================================================
# Vias Teams Archivering - Volledig Automatisch Script v8
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
    & $pwshPath -NoProfile -ExecutionPolicy Bypass -File $MyInvocation.MyCommand.Path
    exit
}

# Vanaf hier: we zitten in de hergestarte schone sessie
Write-Host "`n  Schone sessie actief. Modules worden geladen..." -ForegroundColor Green

Import-Module Microsoft.Graph.Authentication,
              Microsoft.Graph.Applications,
              Microsoft.Graph.Groups,
              MicrosoftTeams, PnP.PowerShell, ImportExcel -ErrorAction Stop

Write-Host "  Alle modules geladen." -ForegroundColor Green

# Cross-platform tijdelijke map
$tempDir = [System.IO.Path]::GetTempPath()
#endregion

#region CONFIGURATIE - Interactief opvragen
Clear-Host
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Vias Teams Archivering - Setup Wizard v8" -ForegroundColor Cyan
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
$archiveStandaard = if ($IsWindows) { "N:\Archive\Vias_Teams" } else { Join-Path $HOME "Documents" "Vias_Teams_Archive" }
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
        -Scopes "Application.ReadWrite.All","AppRoleAssignment.ReadWrite.All","Directory.ReadWrite.All" `
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
Write-Host "`n[2/12] Entra ID App registreren..." -ForegroundColor Cyan

$appName = "BraveHub-Vias-Teams-Archiver"
$app     = Get-MgApplication -Filter "displayName eq '$appName'" -ErrorAction SilentlyContinue |
           Select-Object -First 1

if ($app) {
    Write-Host "  Bestaande app gevonden: $($app.AppId)" -ForegroundColor Yellow
} else {
    try {
        $app = New-MgApplication -DisplayName $appName `
            -PublicClient @{ RedirectUris = @("http://localhost") } `
            -IsFallbackPublicClient -ErrorAction Stop
        Write-Host "  Nieuwe app aangemaakt: $($app.AppId)" -ForegroundColor Green
    } catch {
        Write-Host "  FOUT bij aanmaken app: $_" -ForegroundColor Red; exit
    }
}

$sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue |
     Select-Object -First 1
if (-not $sp) {
    try {
        $sp = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop
        Write-Host "  Service Principal aangemaakt." -ForegroundColor Green
    } catch {
        Write-Host "  FOUT bij aanmaken Service Principal: $_" -ForegroundColor Red; exit
    }
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
    foreach ($scopeName in $Scopes) {
        $permDef = $ResourceSp.Oauth2PermissionScopes |
                   Where-Object { $_.Value -eq $scopeName } | Select-Object -First 1
        if (-not $permDef) {
            Write-Warning "    Scope '$scopeName' niet gevonden op $ResourceName"
            continue
        }
        $existing = Get-MgOauth2PermissionGrant `
            -Filter "clientId eq '$($sp.Id)' and resourceId eq '$($ResourceSp.Id)'" `
            -ErrorAction SilentlyContinue |
            Where-Object { $_.Scope -split " " -contains $scopeName }

        if ($existing) {
            Write-Host "    Al aanwezig: $scopeName" -ForegroundColor Gray
        } else {
            New-MgOauth2PermissionGrant `
                -ClientId    $sp.Id `
                -ResourceId  $ResourceSp.Id `
                -Scope       $scopeName `
                -ConsentType "AllPrincipals" | Out-Null
            Write-Host "    Toegekend: $scopeName" -ForegroundColor Green
        }
    }
}

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
        -UseDeviceAuthentication -NoWelcome -ErrorAction Stop
} catch {
    Write-Host "  FOUT bij tweede login: $_" -ForegroundColor Red; exit
}

try {
    Connect-MicrosoftTeams -TenantId $tenantId -ErrorAction Stop
} catch {
    Write-Host "  FOUT bij Teams-verbinding: $_" -ForegroundColor Red; exit
}

Write-Host "  Verbonden als: $((Get-MgContext).Account)" -ForegroundColor Green
#endregion

#region STAP 5 - Excel inlezen en Teams-IDs ophalen
Write-Host "`n[5/12] Excel inlezen en Teams-IDs ophalen..." -ForegroundColor Cyan

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
        if ($row.ChannelType -eq "Standard") {
            $siteUrl = "$tenantUrl/teams/$($row.TeamName -replace '\s','')"
            $folder  = "Shared Documents/$($row.ChannelName)"
        } else {
            $channel = Get-TeamChannel -GroupId $groupId |
                       Where-Object { $_.DisplayName -eq $row.ChannelName } |
                       Select-Object -First 1
            if (-not $channel) {
                Write-Warning "  Kanaal niet gevonden: $($row.ChannelName)"
                continue
            }
            $ff = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/v1.0/teams/$groupId/channels/$($channel.Id)/filesFolder"
            $siteUrl = $ff.webUrl -replace "(/[^/]+){2}$", ""
            $folder  = "Shared Documents"
        }

        if ($siteUrl -ne $currentSiteUrl) {
            Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $clientId
            $currentSiteUrl = $siteUrl
        }

        $items = Get-PnPFolderItem -FolderSiteRelativeUrl $folder -ItemType File -Recursive `
                     -ErrorAction SilentlyContinue

        if (-not $items -or $items.Count -eq 0) {
            Write-Host "  Leeg: $($row.TeamName) / $($row.ChannelName)" -ForegroundColor Gray
            continue
        }

        foreach ($item in $items) {
            Get-PnPFile -Url $item.ServerRelativeUrl `
                -Path $destPath -Filename $item.Name -AsFile -Force
        }

        Write-Host "  OK [$($row.ChannelType)] $($row.TeamName) / $($row.ChannelName) ($($items.Count) bestanden)" `
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

            $allMessages = [System.Collections.Generic.List[object]]::new()
            $uri = "https://graph.microsoft.com/v1.0/teams/$groupId/channels/$($channel.Id)/messages"
            do {
                $result = Invoke-MgGraphRequest -Method GET -Uri $uri
                if ($result.value) { $allMessages.AddRange($result.value) }
                $uri = $result.'@odata.nextLink'
            } while ($uri)

            $allData = [System.Collections.Generic.List[object]]::new()
            foreach ($msg in $allMessages) {
                $msgObj = [PSCustomObject]@{
                    Id       = $msg.id
                    Datum    = $msg.createdDateTime
                    Afzender = $msg.from.user.displayName
                    Email    = $msg.from.user.email
                    Bericht  = ($msg.body.content -replace '<[^>]+>', '')
                    Bijlagen = ($msg.attachments | ForEach-Object { $_.name }) -join ", "
                    Replies  = @()
                }
                try {
                    $rep = Invoke-MgGraphRequest -Method GET `
                        -Uri "https://graph.microsoft.com/v1.0/teams/$groupId/channels/$($channel.Id)/messages/$($msg.id)/replies"
                    if ($rep.value) {
                        $msgObj.Replies = $rep.value | ForEach-Object {
                            [PSCustomObject]@{
                                Datum    = $_.createdDateTime
                                Afzender = $_.from.user.displayName
                                Bericht  = ($_.body.content -replace '<[^>]+>', '')
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
$verwijder = Read-Host "Entra app '$appName' verwijderen? (j/n)"
if ($verwijder -eq "j") {
    Remove-MgServicePrincipal -ServicePrincipalId $sp.Id -ErrorAction SilentlyContinue
    Remove-MgApplication -ApplicationId $app.Id
    Remove-Item (Join-Path $tempDir "vias_archiver_clientid.txt") -ErrorAction SilentlyContinue
    Write-Host "  App en tijdelijke bestanden verwijderd." -ForegroundColor Yellow
} else {
    Write-Host "  App behouden. Client ID: $clientId" -ForegroundColor Gray
}

# Omgevingsvariabele opruimen
$env:VIAS_ARCHIVER_HERSTART = $null

Write-Host "`n===== ARCHIVERING VOLTOOID =====" -ForegroundColor Cyan
Write-Host "  Rapport : $rapportPad"            -ForegroundColor Green
Write-Host "  Archief : $archiveRoot"            -ForegroundColor Green
Write-Host "================================`n"  -ForegroundColor Cyan
#endregion