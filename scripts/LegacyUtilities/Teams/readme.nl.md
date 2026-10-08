[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [LegacyUtilities](../readme.nl.md) › **Teams**

# Legacy Utilities — Teams

Hulpscripts voor het inrichten van Teams en Planner, volledig via Microsoft Graph — de MicrosoftTeams-module is niet meer nodig. Ze melden aan via [`Connect-M365.ps1`](../../Startup/readme.nl.md): standaard delegated als beheerder (browser, of apparaatcode / GDAP-klant volgens `load.config.ps1`), app-only met `-ClientId` + `-CertificateThumbprint` of `-AppOnly` (app uit `graph.appid.json`). Een passende sessie voor de juiste tenant wordt hergebruikt en blijft verbonden; alleen een sessie die het script zelf opende, wordt verbroken. Elk script accepteert `-TenantId`, `-ClientId`, `-CertificateThumbprint` en `-AppOnly`.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Copy-Team.ps1`](Copy-Team.ps1) ([docs](#copy-teamps1)) | Kloon een bestaand Team (apps/tabbladen/instellingen/kanalen/leden) |
| [`Copy-PlannerPlan.ps1`](Copy-PlannerPlan.ps1) ([docs](#copy-plannerplanps1)) | Kopieer de buckets/taken/checklists van een Planner-plan naar een nieuw plan |
| [`New-ProjectTeam.ps1`](New-ProjectTeam.ps1) ([docs](#new-projectteamps1)) | Maak Teams in bulk aan met een standaard kanaalsjabloon uit een CSV |

---

### Copy-Team.ps1

Dient via Graph een kloonopdracht voor een Team in (`POST /teams/{id}/clone`, asynchroon) en
blijft pollen tot het nieuwe Team verschijnt. Standaard een proefdraai.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-SourceTeamId` | Ja | Object-ID of weergavenaam van het Team dat gekloond wordt |
| `-NewTeamName` | Ja | Weergavenaam van de kloon |
| `-NewTeamDescription` / `-NewMailNickname` | Nee | Vallen terug op `-NewTeamName` als ze worden weggelaten |
| `-Visibility` | Nee | `Private` (standaard) of `Public` |
| `-PartsToClone` | Nee | Een of meer van Apps, Tabs, Settings, Channels, Members (standaard: alles) |
| `-Apply` | Nee | Dien de kloonopdracht echt in (standaard: voorbeeldweergave) |

```powershell
.\Copy-Team.ps1 -SourceTeamId "Project Template" -NewTeamName "Project 1234" -Apply
```

**Opmerkingen**
- Pollt de kloonbewerking uit de `Location`-header en meldt de ID van het nieuwe Team, of de fout als het klonen mislukte; voorheen wachtte het op een willekeurige groep met de nieuwe naam, waaraan een bestaande groep met die naam ook voldeed
- Aanhalingstekens in een weergavenaam voor `-SourceTeamId` worden ge-escaped voor het filter

---

### Copy-PlannerPlan.ps1

Kopieert elke bucket en taak (met beschrijvingen en checklists) van een bron-Planner-plan naar
een nieuw aangemaakt plan in een andere groep. Voegt twee oude varianten (kale Graph REST en
PnP.PowerShell) samen tot één script dat Microsoft.Graph.Planner gebruikt; PnP is niet nodig.
Standaard een proefdraai.

```powershell
.\Copy-PlannerPlan.ps1 -SourcePlanId "xqQg5FS2LkCp935s-FIFm2QAFkHM" -DestinationGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Apply
```

**Opmerkingen**
- Leest alle buckets en taken pagina voor pagina (bij een plan groter dan één pagina ging de rest verloren)
- Checklists worden weer gekopieerd — de items werden uit de verkeerde eigenschappen gelezen — ook bij taken met een checklist maar zonder beschrijving
- Delegated moet de beheerder lid zijn van beide groepen; app-only werkt met de applicatiemachtiging `Tasks.ReadWrite.All`

---

### New-ProjectTeam.ps1

Maakt Teams in bulk aan vanuit een CSV (één rij per Team) en past op elk Team hetzelfde
kanaalsjabloon toe (een tweede CSV). Gegeneraliseerde vervanging van een oud script waarin de
vaste indeling van projectkanalen van één bedrijf hardcoded stond. Standaard een proefdraai.

**CSV-invoer**

```csv
# teams.csv
TeamName,MailNickname,Owner,Visibility
"Project 1001","project-1001","pm@contoso.com","Private"

# channels.csv
ChannelName,Description
"Documents","Signed customer documents"
"Internal",""
```

```powershell
.\New-ProjectTeam.ps1 -TeamsCsvPath .\teams.csv -ChannelsCsvPath .\channels.csv -Apply
```

---

**Vereiste modules**

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

**Hoe het werkt**

1. `POST /groups` maakt de Microsoft 365-groep aan met de `MailNickname` uit de CSV, met de eigenaar als eigenaar en lid
2. `POST /teams` met `group@odata.bind` en de sjabloon `standard` maakt er een Team van (404's worden opnieuw geprobeerd zolang de nieuwe groep repliceert)
3. De asynchrone bewerking uit de `Location`-header wordt gepolld tot het Team is ingericht
4. `POST /teams/{id}/channels` voegt elk kanaal toe

Gedelegeerde scopes: `Group.ReadWrite.All`, `User.Read.All`, `Team.Create`, `Channel.Create`. App-only: `Group.ReadWrite.All`, `User.Read.All` (applicatiemachtigingen). Een CSV zonder `TeamName` of `ChannelName` wordt nu geweigerd (die controle werkte voorheen nooit).
