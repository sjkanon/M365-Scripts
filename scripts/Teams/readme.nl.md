[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **Teams**

# Teams

Tooling voor export en archivering van Microsoft Teams / SharePoint.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Invoke-TeamsArchive.ps1`](Invoke-TeamsArchive.ps1) ([docs](#invoke-teamsarchiveps1)) | Voert een export- en archiveringsflow voor Teams/SharePoint uit voor een lijst teams en kanalen uit Excel |

---

### Invoke-TeamsArchive.ps1

Teams-archiver: exporteert leden, bestanden en chat van de opgegeven kanalen en archiveert daarna (optioneel) teams of kanalen. Alles loopt via Microsoft Graph. Vereist PowerShell 7+, uitvoeren als Global Admin van de klant (of met GDAP-rechten op die klant).

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-Step10Action` | `interactive` (standaard), `archive`, `undo` of `skip` |
| `-Step10Only` | Alleen Stap 10 (archiveren/dearchiveren) uitvoeren, niet-interactief |
| `-ChannelAction` | `none` (standaard), `archive` of `undo` — snelle modus per kanaal |
| `-ChannelArchiveTag` | Markeringstekst voor de terugval via hernoemen (standaard: `[ARCHIEF]`) |
| `-ChannelFallbackToRename` | Terugvallen op een hernoemmarkering als de Graph-API-aanroep voor archiveren/dearchiveren mislukt |
| `-DryRun` | Simuleren — behoudt de volledige aanmelding en valideert Stap 6-9 door aantallen op te vragen, zonder exports te schrijven of de archiefstatus te wijzigen |
| `-WorksheetName` | Werkblad met de teamlijst. Standaard: het eerste werkblad van het Excel-bestand |
| `-TenantId` | Tenant-ID van de klant. Standaard voorgesteld in de wizard: de GDAP-klant uit `Connect-Tenant` |
| `-ClientId` / `-CertificateThumbprint` | App-only in plaats van gedelegeerd, met deze app-registratie en dit certificaat |
| `-AppOnly` | App-only met de app-registratie voor de tenant uit `graph.appid.json` |
| `-PnPClientId` | PnP-app voor de site-admin-terugval. Standaard: de regel voor de tenant in `pnp.appid.json` |

**Voorbeelden**

```powershell
# Gedelegeerd (standaard): apparaatcode volgens load.config.ps1, eerst een dry run
.\Invoke-TeamsArchive.ps1 -DryRun

# Alleen de opgegeven kanalen archiveren, app-only
.\Invoke-TeamsArchive.ps1 -Step10Only -ChannelAction archive -AppOnly -TenantId <tenant-guid>
```

**Opmerkingen**

Het Excel-bestand heeft de kolommen `TeamName`, `ChannelName` en `Archive` nodig (rijen met `Archive` = `Archive` worden verwerkt). Niets in het script is aan één klant gebonden: tenant-ID, SharePoint-URL, Excel-bestand en archiefmap vraagt de setupwizard (standaard `C:\Temp\Teams_Channels.xlsx` en `C:\Temp\Teams_Archive`).

Aanmelden (v9.0):
- Loopt via [`Connect-M365.ps1`](../Startup/readme.nl.md#connect-m365ps1). **Standaard gedelegeerd**: je meldt je aan als de admin, met een apparaatcode als `useDeviceCodeAuth` in `load.config.ps1` aan staat (en zonder `load.config.ps1`, zoals vroeger, met een apparaatcode). De herstart in een schone sessie neemt die instellingen en de GDAP-klant mee naar de nieuwe sessie.
- **Geen tijdelijke app-registratie meer.** Eerdere versies maakten er bij elke run een aan, gaven toestemming en verwijderden ze weer; de gedelegeerde scopes worden nu bij het aanmelden gevraagd op de app Microsoft Graph Command Line Tools.
- **App-only** met `-ClientId` + `-CertificateThumbprint`, of `-AppOnly`. De app heeft dan de application-rechten `Group.Read.All`, `Sites.Read.All`, `TeamMember.Read.All`, `ChannelMessage.Read.All` (een protected API die Microsoft moet goedkeuren) en `TeamSettings.ReadWrite.All` / `ChannelSettings.ReadWrite.All` voor het archiveren nodig.
- De Graph-sessie wordt op het einde gesloten; ze is van de herstarte sessie zelf.

Wat waar loopt:
- **Graph** voor alles wat kan: de teams vinden (`/groups`, gefilterd op `resourceProvisioningOptions` = `Team`), leden (`/teams/{id}/members`), kanalen (`/teams/{id}/channels`), de bestandslocatie van het kanaal (`filesFolder`), het oplijsten en downloaden van bestanden (`/drives/{id}/items/{id}/children` en `/content`), chat (`/messages`, `/replies`) en archiveren (`/teams/{id}/archive`, `/channels/{id}/archive`). De MicrosoftTeams-module wordt niet meer gebruikt of geïnstalleerd.
- **PnP** alleen voor de site-admin-terugval: als de bestanden van een kanaal bij een gedelegeerde aanmelding "access denied" geven, wordt de aangemelde admin eenmaal per site sitecollectiebeheerder gemaakt (`Set-PnPSite -Owners`) — daar heeft Graph geen API voor. Dat gebruikt de PnP-app uit `pnp.appid.json` (of `-PnPClientId`); zonder app zegt het script hoe je het met de hand doet. App-only heeft het nooit nodig.

Gedrag:
- Bepaalt de bestandslocaties van kanalen via Graph `filesFolder` voor alle kanaaltypen (standaard/privé/gedeeld), met de kanaallijst gecachet per team.
- Normaliseert TeamName-/ChannelName-waarden uit Excel (trim) en vergelijkt kanaalnamen genormaliseerd (trim + witruimte samenvoegen + kleine letters) om onterechte "Kanaal niet gevonden"-gevallen te vermijden.
- Behandelt SharePoint NotFound tijdens de export als een gecontroleerde overslag.
- Downloadt bestanden met nieuwe pogingen per bestand en een controle van het aantal na het downloaden. De mappenstructuur binnen het kanaal blijft nu behouden onder `Files` (vroeger werd ze platgeslagen, waardoor twee bestanden met dezelfde naam in verschillende submappen elkaar overschreven en de telling mislukte).
- Slaat de uitvoer op per kanaal: `Teams > Team > Channel > Files, Chat, Members`. `members.csv` houdt de kolommen `Name`, `User`, `Role`.
- Archiveert Teams niet standaard; archiveren vereist expliciete bevestiging in Stap 10, die ook `unarchive` ondersteunt, een niet-interactieve snelle modus (`-Step10Only -Step10Action undo|archive|skip`) en echt archiveren/dearchiveren per kanaal (`-Step10Only -ChannelAction archive|undo`), met een optionele terugval op een hernoemmarkering (`-ChannelFallbackToRename`, `-ChannelArchiveTag`).
- Een team archiveren/dearchiveren is in Microsoft Teams een actie op teamniveau.
- Dry-run behoudt de volledige aanmelding, valideert Stap 6-9 door aantallen op te vragen zonder naar schijf te schrijven, gebruikt die aantallen in het rapport van Stap 11 en simuleert Stap 10 enkel (`[DRYRUN]`).
- De herstart in een schone sessie na het opkuisen van de modules geeft alle parameters door (ook `-DryRun`), geeft de exitcode van de herstarte run terug en laat haar markeringsvariabele niet meer achter in de aanroepende sessie (een tweede run in dezelfde sessie sloeg vroeger de opkuis over).
