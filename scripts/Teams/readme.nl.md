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

Teams-archiver met een exportflow voor Graph, Teams en SharePoint. Vereist PowerShell 7+, uitvoeren als Global Admin.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-Step10Action` | `interactive` (standaard), `archive`, `undo` of `skip` |
| `-Step10Only` | Alleen Stap 10 (archiveren/dearchiveren) uitvoeren, niet-interactief |
| `-ChannelAction` | `none` (standaard), `archive` of `undo` — snelle modus per kanaal |
| `-ChannelArchiveTag` | Markeringstekst voor de terugval via hernoemen (standaard: `[ARCHIEF]`) |
| `-ChannelFallbackToRename` | Terugvallen op een hernoemmarkering als de Graph-API-aanroep voor archiveren/dearchiveren mislukt |
| `-DryRun` | Simuleren — behoudt de volledige authenticatie/bootstrap en valideert Stap 6-9 door aantallen op te vragen, zonder exports te schrijven of de archiefstatus te wijzigen |
| `-WorksheetName` | Werkblad met de teamlijst. Standaard: het eerste werkblad van het Excel-bestand |

Het Excel-bestand heeft de kolommen `TeamName`, `ChannelName` en `Archive` nodig (rijen met `Archive` = `Archive` worden verwerkt). Niets in het script is aan één klant gebonden: tenant-ID, SharePoint-URL, Excel-bestand en archiefmap vraagt de setupwizard (standaard `C:\Temp\Teams_Channels.xlsx` en `C:\Temp\Teams_Archive`).

Huidig gedrag (v8.19):
- Maakt voor de run een unieke tijdelijke Entra-app-registratie aan.
- Kent tijdens de bootstrap alleen de vereiste gedelegeerde setuprechten toe.
- Past gedelegeerde toestemming voor Graph/SharePoint toe op die tijdelijke app.
- Verwijdert de tijdelijke app en service principal bij het opruimen (en bij belangrijke setupfouten).
- Registreert een opruimhook bij afsluiten, zodat de tijdelijke app ook wordt verwijderd bij het afsluiten van PowerShell/Ctrl+C.
- Controleert tijdens de bestandsexport eerst de toegang tot Teams-/SharePoint-mappen en kent pas hogere Graph-rechten toe als de toegang geweigerd wordt.
- Bepaalt de bestandslocaties van kanalen via Graph filesFolder voor alle kanaaltypen (standaard/privé/gedeeld), met caching van kanalen en een terugvalzoekactie.
- Normaliseert TeamName-/ChannelName-waarden uit Excel (trim) om missers bij het opzoeken door afsluitende spaties te voorkomen.
- Hergebruikt dezelfde gecachte kanaalresolver met Graph-terugval in de chatexport, wat de consistentie van kanaaldetectie in dry-run en normale runs verbetert.
- Past genormaliseerde matching van kanaalnamen toe (trim + witruimte samenvoegen + kleine letters) in de gecachte en de Graph-terugvalzoekactie, om onterechte "Kanaal niet gevonden"-gevallen te verminderen.
- Behandelt SharePoint NotFound tijdens de export als een gecontroleerde overslag in plaats van rumoerige harde fouten.
- Downloadt bestanden met nieuwe pogingen per bestand, terugval via opnieuw verbinden en controle van het aantal na het downloaden om volledigheid te garanderen.
- Slaat de uitvoer op in een structuur per kanaal: `Teams > Team > Channel > Files, Chat, Members`.
- Archiveert Teams niet standaard; archiveren vereist nu expliciete bevestiging tijdens Stap 10.
- Stap 10 ondersteunt ook het ongedaan maken van archiveren (`unarchive`) met retrylogica.
- Stap 10 ondersteunt een niet-interactieve snelle modus: `-Step10Only -Step10Action undo|archive|skip`.
- Belangrijk: archiveren/dearchiveren is in Microsoft Teams een actie op teamniveau, niet op kanaalniveau.
- Stap 10 ondersteunt echt archiveren/dearchiveren per kanaal via Microsoft Graph (`/channels/{id}/archive|unarchive`).
- Snelle modus per kanaal: `-Step10Only -ChannelAction archive|undo`.
- Optionele terugval op een hernoemmarkering bij een API-fout: `-ChannelFallbackToRename` (markering via `-ChannelArchiveTag`).
- De dry-runmodus behoudt de volledige authenticatie/bootstrap en valideert Stap 6-9 door het bestaan en de aantallen in Teams/SharePoint/Graph op te vragen, zonder export van leden/chats/bestanden naar schijf te schrijven.
- Het rapport van Stap 11 gebruikt in dry-run de opgevraagde aantallen (gedetecteerde bestanden/berichten) in plaats van lokaal geëxporteerde bestanden.
- Mutaties voor archiveren/dearchiveren in Stap 10 blijven gesimuleerd, met `[DRYRUN]`-uitvoer.
- De herstart in een schone sessie na het opkuisen van de modules geeft alle parameters door (ook `-DryRun`) en geeft de exitcode van de herstarte run terug.
