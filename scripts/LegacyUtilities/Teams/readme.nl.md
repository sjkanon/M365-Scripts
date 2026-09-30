[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [LegacyUtilities](../readme.nl.md) › **Teams**

# Legacy Utilities — Teams

Hulpscripts voor het inrichten van Teams en Planner via Microsoft Graph (en Microsoft Teams
PowerShell voor het in bulk aanmaken van Teams/kanalen). Ze maken automatisch verbinding als er
geen sessie actief is, en hergebruiken een bestaande sessie als je al verbonden bent.

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

---

### Copy-PlannerPlan.ps1

Kopieert elke bucket en taak (met beschrijvingen en checklists) van een bron-Planner-plan naar
een nieuw aangemaakt plan in een andere groep. Voegt twee oude varianten (kale Graph REST en
PnP.PowerShell) samen tot één script dat Microsoft.Graph.Planner gebruikt; PnP is niet nodig.
Standaard een proefdraai.

```powershell
.\Copy-PlannerPlan.ps1 -SourcePlanId "xqQg5FS2LkCp935s-FIFm2QAFkHM" -DestinationGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Apply
```

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
Install-Module MicrosoftTeams -Scope CurrentUser
```
