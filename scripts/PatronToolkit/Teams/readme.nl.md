[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [PatronToolkit](../readme.nl.md) › **Teams**

# Patron Toolkit — Teams

Governance van de Microsoft Teams-tenant en inventarisrapportage.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Get-TeamsConfigReport.ps1`](Get-TeamsConfigReport.ps1) ([docs](#get-teamsconfigreportps1)) | Rapport van de governance-instellingen van de Teams-tenant en een inventaris van teams |

---

### Get-TeamsConfigReport.ps1

Rapporteert het tenantbrede Teams-beleid dat er bij een governance-/beveiligings-
review het meest toe doet: externe toegang (federatie), gasttoegang, globaal vergaderbeleid (anoniem
deelnemen, opnemen, presentatorrol), globaal berichtenbeleid en app-instellingsbeleid
(sideloading). Toont ook elk team met zichtbaarheid, archiefstatus en aantallen eigenaren/leden
— en markeert elk team zonder eigenaren.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-IncludeTeamsInventory` | Nee | Ook elk team tonen met het aantal leden/eigenaren, via Graph (standaard: aan) |
| `-OutputPath` | Nee | Map voor de CSV-rapport(en) (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard de GDAP-klant (`load.config.ps1`) of je eigen tenant; verplicht voor app-only |
| `-ClientId` | Nee | App-registratie voor app-only aanmelden (met `-CertificateThumbprint`). Zonder meldt het script gedelegeerd aan, als jezelf |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor app-only aanmelden met `-ClientId` |
| `-AppOnly` | Nee | App-only aanmelden met de ClientId en vingerafdruk voor de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
.\Get-TeamsConfigReport.ps1

.\Get-TeamsConfigReport.ps1 -IncludeTeamsInventory:$false

# App-only, Teams en Graph met de app uit graph.appid.json
.\Get-TeamsConfigReport.ps1 -TenantId contoso.onmicrosoft.com -AppOnly
```

**Opmerkingen**
- Alleen-lezen
- Het `Cs*`-tenantbeleid blijft op Teams PowerShell (`Connect-M365Teams`) — Microsoft Graph
  heeft er geen API voor. De teaminventaris is naar Graph verhuisd: `/groups` gefilterd op
  Team-provisioning, `/teams/{id}` (`isArchived`) en `/teams/{id}/members` (eigenaarsrol),
  in plaats van `Get-Team` / `Get-TeamUser`
- Gasttoegang leest nu `AllowGuestUser` (`Get-CsTeamsClientConfiguration`) en
  `DisableAnonymousJoin` (`Get-CsTeamsMeetingConfiguration`); de vorige versie las
  `AllowAnonymousUsersToJoinMeeting` uit de gastvergaderconfiguratie, die die eigenschap niet
  heeft, en meldde altijd een lege waarde
- Aanmelden via [`Connect-M365.ps1`](../../Startup/readme.nl.md#connect-m365ps1), standaard
  gedelegeerd (rol Teams Administrator of Global Reader; Graph-scopes `Group.Read.All`,
  `TeamMember.Read.All`, `TeamSettings.Read.All`); app-only met `-ClientId` +
  `-CertificateThumbprint` of `-AppOnly` voor beide. Met `-IncludeTeamsInventory:$false`
  wordt er niet bij Graph aangemeld

**Vereiste modules**
```powershell
Install-Module MicrosoftTeams -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```
