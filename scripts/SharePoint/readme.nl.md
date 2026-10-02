[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **SharePoint**

# SharePoint Scripts

Contentbewerkingen op SharePoint Online en OneDrive via PnP PowerShell, met interactieve
admin-aanmelding op elke (klant)tenant.

---

## Mappen

| Map | Omschrijving |
|--------|-------------|
| [`Provisioning/`](Provisioning/readme.nl.md) | Een complete structuur inrichten en onderhouden — metadatamodel, contenttypes, bibliotheken en groepsrechten — vanuit één configbestand, plus een deelaudit en een drift-check |

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Find-SiteContent.ps1`](Find-SiteContent.ps1) ([docs](#find-sitecontentps1)) | Doorzoek een hele site (naam, pad, type, grootte, datum of volledige tekst) en rapporteer de rechten op elke treffer — PnP/CSOM, meldt zich aan als jou |
| [`Search-SharePointContent.ps1`](Search-SharePointContent.ps1) ([docs](#search-sharepointcontentps1)) | Dezelfde vraag tenantbreed via Microsoft Graph, app-only, zonder interactieve login — bestanden en mappen |
| [`Restore-RecycleBinItems.ps1`](Restore-RecycleBinItems.ps1) ([docs](#restore-recyclebinitemsps1)) | Verwijderde bestanden/mappen terugzetten uit de prullenbak van een site of OneDrive (standaard een proefdraai) |
| [`Trace-SharePointFile.ps1`](Trace-SharePointFile.ps1) ([docs](#trace-sharepointfileps1)) | Waar is een bestand gebleven? Hernoemingen, verplaatsingen, kopieën en verwijderingen uit het auditlog, ook via een map — in Brusselse tijd, over een periode naar keuze |
| [`Revoke-SharePointUserAccess.ps1`](Revoke-SharePointUserAccess.ps1) ([docs](#revoke-sharepointuseraccessps1)) | Neem één gebruiker overal de toegang af: site collection-beheerder, directe toekenningen op elk niveau, SharePoint-groepen en deellinks. Rapporteert standaard, verwijdert met `-Apply` |
| [`Test-SharePointAccessScripts.ps1`](Test-SharePointAccessScripts.ps1) ([docs](#test-sharepointaccessscriptsps1)) | Controleer de twee toegangsscripts zonder een tenant aan te raken — gedeeld authenticatieblok identiek, en de revocatie-trechter gedraagt zich |

---

## Welke van de twee zoekscripts?

Beide beantwoorden "waar staat dit en wie kan erbij", en beide schrijven dezelfde soort
CSV. Ze verschillen in wat ze kunnen zien en in wat ze van jou nodig hebben.

| | `Find-SiteContent.ps1` (PnP) | `Search-SharePointContent.ps1` (Graph) |
|---|---|---|
| Aanmelden | Interactief, als jij | App-only, onbeheerd — prima in een geplande taak |
| Benodigde rechten | Toegang tot de site, of `-GrantSiteAdmin` per site | Eén admin consent, één keer, voor de hele tenant |
| Bereik | Eén site collection (+ subsites) | Eén site, of **elke site in de tenant**, OneDrive inbegrepen |
| Bestanden en mappen | Ja | Ja, en sneller — `delta` leest een bibliotheek in pagina's van duizend en rechten komen per 20 in een `$batch` |
| Items in gewone lijsten | Ja, met hun rechten | **Nee** — Graph geeft alleen rechten voor driveItems |
| Rechten op site- en lijstniveau | Ja: "dit erft van de bibliotheek, die Bewerken geeft aan Site Members" | **Nee** — Graph heeft geen API voor SharePoint-roltoewijzingen; het zegt *of* een item erft en waarvan, niet wat de site toekent |
| Deellinks | Ja, uit de `SharingLinks.*`-groepen | Ja, rijker: linkbereik, bewerken/bekijken, vervaldatum en de link-URL zelf |

Vuistregel: **Graph** voor "vind het ergens in de tenant en laat me de links en gasten erop
zien", **PnP** als je het volledige rechtenverhaal van één site nodig hebt, inclusief de
lijsten en groepen.

---

### Find-SiteContent.ps1

Beantwoordt de twee vragen die je meestal tegelijk hebt: **waar staat dit** en **wie kan
erbij**. Alleen-lezen — het script wijzigt nooit iets.

**Twee engines**

| Engine | Wanneer | Wat die ziet |
|--------|------|--------------|
| Crawl (standaard) | Geen `-Content` opgegeven | Loopt elke lijst en bibliotheek af. `-Name` en `-ItemType` worden waar mogelijk in een CAML-query gezet, zodat SharePoint alleen de treffers teruggeeft; wat CAML niet kan uitdrukken valt terug op het volledig lezen van die lijst. Ziet alles, ook wat de zoekindex nog niet heeft opgepikt |
| Search (`-Content`) | Zoeken op volledige tekst | Een KQL-query tegen de zoekindex, beperkt tot het sitepad — dit is degene die tekst *binnen* documenten vindt. Snel, maar beperkt tot wat geïndexeerd is en wat het aangemelde account mag zien. Neemt subsites automatisch mee |

Beide engines voeden dezelfde filters: `-Name` (wildcards), `-Path`, `-Extension`,
`-ItemType`, `-ListName`, `-ModifiedBy`, `-ModifiedAfter` / `-ModifiedBefore`,
`-MinSizeMB`. `-Name` wordt vergeleken met de bestandsnaam, de itemtitel *én* het laatste
segment van de URL, zodat een item waarvan de titel afwijkt van de bestandsnaam toch
gevonden wordt.

Verborgen en systeembibliotheken worden overgeslagen tenzij je `-IncludeHidden` meegeeft,
en bij het crawlen wordt alleen de bovenste web doorzocht tenzij je `-IncludeSubsites`
toevoegt — het script meldt per web hoeveel lijsten het oversloeg en waarom.
**`-Everything` zet dat allemaal in één keer uit**: elke subsite, elke verborgen en
systeemlijst, geen limiet op de treffers en geen op de rechten-lookups. Een crawl vergelijkt
nog steeds alleen namen en metadata; gebruik `-Content` om in de documenten zelf te zoeken.

**Rechten per treffer**

Voor elke treffer bepaalt het script waar de rechten werkelijk vandaan komen:

| Bron | Betekenis |
|--------|---------|
| `Item` | Het item heeft de overerving doorbroken en heeft eigen roltoewijzingen |
| `List` | Het erft van een bibliotheek/lijst met unieke rechten |
| `Site` | Het erft helemaal door tot de (sub)site |

Roltoewijzingen worden platgeslagen tot één CSV-regel per principal — principaltype, login,
e-mail en de rolnamen (`Full Control`, `Edit`, …). `Limited Access` wordt verborgen tenzij je
`-IncludeLimitedAccess` meegeeft; die toewijzingen bestaan alleen zodat iemand een dieper
item kan bereiken en geven op zichzelf niets.

Drie dingen worden apart uitgelicht, omdat juist die mensen verrassen:

- **Deellinks** — de `SharingLinks.*`-groepen achter elke "Koppeling kopiëren". Altijd
  uitgeklapt naar de mensen erin en gelabeld als *Anyone* / *Organization* / *Specific
  people*, zodat een anonieme link niet tussen de ruis kan verdwijnen
- **Externe gebruikers** — gastaccounts (`#ext#`) in welke toewijzing dan ook
- **Everyone** — "Everyone" en "Everyone except external users"

Site- en lijstrechten worden één keer gelezen en gecachet, itemrechten alleen voor items die
de overerving echt hebben doorbroken, dus een zoekopdracht met een handvol treffers kost een
handvol extra calls. `-Permissions Unique` is de snelle manier om te beantwoorden "wat in deze
site is anders gedeeld dan de rest"; `-Permissions None` slaat rechten helemaal over.
`-MaxPermissionLookups` (standaard 1000) voorkomt dat een te brede zoekopdracht urenlang
doorloopt.

**Aanmelden**

Hetzelfde als `Restore-RecycleBinItems.ps1`: de eerste run tegen een tenant registreert een
public-client Entra-app (delegated `AllSites.FullControl`, met admin consent) en cachet de
client ID per tenant in `pnp.appid.json` in de root van de repo (gitignored). Een client ID
die het andere script al heeft gecachet wordt hergebruikt, dus meestal kost dit niets. Geef
`-ClientId` mee om de app-registratie helemaal over te slaan.

Rechten uitlezen vereist toegang tot de site. `-GrantSiteAdmin` maakt de aangemelde admin
site collection-beheerder voor de duur van de run en haalt die rechten daarna weer weg
(behoud ze met `-KeepSiteAdmin`) — zo kun je de OneDrive van iemand anders doorzoeken, of
een site waar je geen lid van bent.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-SiteUrl` | Ja | Te doorzoeken site collection (teamsite, communicatiesite of een OneDrive) |
| `-IncludeSubsites` | Nee | Crawl ook elke subsite eronder |
| `-Content` | Nee | Volledige-tekst-/KQL-query — schakelt over naar de zoekindex |
| `-Name` | Nee | Filter op de naam, wildcards toegestaan (`*offerte*`) — vergeleken met de bestandsnaam, de itemtitel *én* het laatste URL-segment |
| `-Path` | Nee | Filter op de map, deelstring-match op de server relative URL |
| `-Extension` | Nee | Eén of meer extensies, met of zonder punt (`xlsx`,`pdf`) |
| `-ItemType` | Nee | `All` (standaard), `File`, `Folder` of `ListItem` |
| `-ListName` | Nee | Alleen deze lijsten/bibliotheken op titel, wildcards toegestaan |
| `-ModifiedBy` | Nee | Wie het laatst wijzigde — weergavenaam of e-mail, wildcards toegestaan |
| `-ModifiedAfter` / `-ModifiedBefore` | Nee | Beperk tot een wijzigingsperiode |
| `-MinSizeMB` | Nee | Alleen bestanden van minstens deze grootte |
| `-IncludeHidden` | Nee | Doorzoek ook verborgen lijsten, catalogi en systeembibliotheken |
| `-Everything` | Nee | Laat niets weg: `-IncludeSubsites -IncludeHidden` plus helemaal geen limieten |
| `-Permissions` | Nee | `Effective` (standaard), `Unique` (alleen doorbroken overerving) of `None` |
| `-ExpandGroups` | Nee | Toon ook de leden van gewone SharePoint-groepen (linkgroepen worden altijd uitgeklapt) |
| `-IncludeLimitedAccess` | Nee | Houd `Limited Access`-toewijzingen in het rapport |
| `-MaxItems` | Nee | Stop na zoveel treffers (standaard 5000, `0` = geen limiet) |
| `-MaxPermissionLookups` | Nee | Limiet op het aantal treffers waarvan de rechten worden opgehaald (standaard 1000, `0` = geen limiet) |
| `-PageSize` | Nee | Items per servercall tijdens het crawlen (standaard 500) |
| `-GrantSiteAdmin` | Nee | Maak jezelf tijdelijk site collection-beheerder (SharePoint Administrator vereist) |
| `-KeepSiteAdmin` | Nee | Behoud die rechten in plaats van ze na afloop weg te halen |
| `-AdminUpn` | Nee | UPN die site-admin wordt (standaard: het aangemelde account) |
| `-TenantId` / `-ClientId` / `-AppName` | Nee | Overschrijvingen voor het aanmelden, zoals in `Restore-RecycleBinItems.ps1` |
| `-OutputPath` | Nee | Pad van het CSV-rapport (standaard: `C:\Temp\SharePointFind_<timestamp>.csv`) |
| `-Disconnect` | Nee | Meld af bij PnP als het klaar is |

**Voorbeelden**

```powershell
# Waar staat alles met "offerte" in de naam, en wie kan het zien?
.\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Sales -Name "*offerte*"

# Volledige tekst: welke documenten noemen "salarisschaal", ergens in de sitestructuur?
.\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/HR `
    -Content "salarisschaal"

# Niets weglaten: alle subsites, alle verborgen/systeemlijsten, geen limieten
.\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -Name "*veiligheid*" -Everything

# Alles in de site dat anders gedeeld is dan de rest
.\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -Permissions Unique -IncludeSubsites

# Grote PDF's in één bibliotheek, met de groepen achter de rechten uitgeklapt
.\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -ListName "Gedeelde documenten" -Extension pdf -MinSizeMB 10 -ExpandGroups

# Een OneDrive doorzoeken waar je geen rechten op hebt
.\Find-SiteContent.ps1 `
    -SiteUrl https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com `
    -Name "*.xlsx" -GrantSiteAdmin
```

**Opmerkingen**
- De CSV heeft één regel per treffer *per principal*, dus hij filtert en pivot goed: sorteer
  op `SharingLink`, `External` of `UniqueRights` om meteen bij de interessante regels te komen.
- `-Content` vindt alleen wat de zoekindex kent. Net geüploade of recent gewijzigde documenten
  kunnen minuten tot uren nodig hebben om te verschijnen — crawl-modus ziet ze altijd.
- Een crawl leest elk item in elke bibliotheek. Op een site met honderdduizenden items duurt
  dat even; beperk het met `-ListName` of gebruik in plaats daarvan `-Content`.

---

### Search-SharePointContent.ps1

De Graph-tegenhanger van `Find-SiteContent.ps1`: dezelfde vraag, app-only, en het bereikt
elke site en elke OneDrive in de tenant zonder dat jij op één daarvan rechten hebt.
Alleen-lezen.

**Twee engines**

| Engine | Wanneer | Wat die doet |
|--------|------|--------------|
| Delta (standaard) | Geen `-Content` | `/drives/{id}/root/delta` loopt een hele bibliotheekboom af in pagina's van duizend items. Filteren gebeurt client-side, dus `*bevat*`-wildcards werken |
| Search (`-Content`) | Zoeken op volledige tekst | `/search/query` tegen de zoekindex — vindt tekst *binnen* documenten. KQL doet alleen wildcards achteraan (`veiligheid*`), niet vooraan. Een app-only zoekopdracht moet een geografie noemen; het script leest die uit de datalocatie van de site, of vindt hem door te proberen, en `-Region` overschrijft dat |

**Rechten**

`/drives/{id}/items/{id}/permissions`, per 20 tegelijk opgehaald via `/$batch`. Die ene call
bevat per item het hele beeld:

| Veld | Gerapporteerd als |
|-------|-------------|
| `roles` | `Read` / `Edit` / `Full Control` |
| `grantedToV2` / `grantedToIdentitiesV2` | De gebruiker, Entra-groep, SharePoint-groep of sitegebruiker — elk een eigen CSV-regel |
| `link` | `SharingLink` (Anyone / Organization / Specific people, bekijken of bewerken), `LinkUrl`, `LinkExpires` |
| `inheritedFrom` | Afwezig → `PermissionSource = Item` (unieke rechten). Aanwezig → `Inherited`, met de map waar het vandaan komt |

Rechten via "iedereen met de link" worden apart geteld in de samenvatting en in het rood
getoond, omdat die bereikbaar zijn zonder zelfs maar aan te melden. Externe gasten (`#EXT#`)
en "Everyone except external users" krijgen ook hun eigen tellers.

**Wat Graph niet kan** — goed om te weten voordat je naar deze grijpt:

- **Geen rechten op site- of lijstniveau.** Er is geen Graph-API voor SharePoint-
  roltoewijzingen; `/sites/{id}/permissions` geeft alleen app-grants terug (Sites.Selected).
  Dit script vertelt je of een item erft en van welke map, niet wat de site zelf aan welke
  groep toekent. Gebruik daarvoor `Find-SiteContent.ps1`.
- **Alleen bibliotheken.** Rechten bestaan voor driveItems, niet voor items in gewone lijsten.
- Geen roldefinities, en geen "Limited Access"-nuance.

**Aanmelden — app-only, voor je aangemaakt**

De eerste run tegen een tenant zet de app op:

1. Meld je (één keer) aan bij Graph als Global Administrator
2. Maak een app aan, of hergebruik die, met de naam uit `-AppName`
3. Ken de **applicatie**rol `Sites.Read.All` toe met admin consent (plus `Group.Read.All`
   als je `-ExpandGroups` gebruikt)
4. Maak een self-signed certificaat aan in `Cert:\CurrentUser\My` en upload de publieke sleutel —
   **er wordt geen secret naar schijf geschreven**
5. Cache de client ID en thumbprint per tenant in `graph.appid.json` in de root van de repo
   (gitignored)

Latere runs verbinden app-only zonder enige prompt, en dat maakt deze bruikbaar vanuit een
geplande taak. Gebruik je eigen app met `-ClientId` plus `-CertificateThumbprint` of
`-ClientSecret`. Het certificaat is niet exporteerbaar en staat in de store van de gebruiker
die het aanmaakte, dus een geplande taak moet onder datzelfde account draaien.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-SiteUrl` | * | Eén site collection (of een OneDrive) om te doorzoeken |
| `-AllSites` | * | Doorzoek in plaats daarvan elke site in de tenant |
| `-SiteFilter` | Nee | Met `-AllSites`: alleen sites waarvan de URL overeenkomt met deze wildcard |
| `-MaxSites` | Nee | Met `-AllSites`: stop na zoveel sites |
| `-IncludePersonalSites` | Nee | Met `-AllSites`: neem ieders OneDrive mee |
| `-IncludeSubsites` | Nee | Met `-SiteUrl`: doorzoek ook de subsites |
| `-Content` | Nee | Volledige-tekst-/KQL-query — schakelt over naar de zoekindex |
| `-Region` | Nee | Geografie voor `-Content` (`EUR`, `NAM`, `DEU`, …). App-only zoeken vereist er een; wordt waar mogelijk automatisch gedetecteerd |
| `-Name` | Nee | Filter op de bestands-/mapnaam, wildcards toegestaan |
| `-Path` | Nee | Filter op het mappad, deelstring-match |
| `-Extension` | Nee | Eén of meer extensies, met of zonder punt |
| `-ItemType` | Nee | `All` (standaard), `File` of `Folder` |
| `-LibraryName` | Nee | Alleen deze documentbibliotheken, wildcards toegestaan |
| `-ModifiedBy` | Nee | Wie het laatst wijzigde — weergavenaam of e-mail |
| `-ModifiedAfter` / `-ModifiedBefore` | Nee | Beperk tot een wijzigingsperiode |
| `-MinSizeMB` | Nee | Alleen bestanden van minstens deze grootte |
| `-Permissions` | Nee | `Effective` (standaard), `Unique` of `None` |
| `-ExpandGroups` | Nee | Toon ook de leden van Entra-groepen (vereist `Group.Read.All`) |
| `-Everything` | Nee | Subsites, persoonlijke sites en geen limieten |
| `-MaxItems` | Nee | Stop na zoveel treffers (standaard 5000, `0` = geen limiet) |
| `-MaxPermissionLookups` | Nee | Limiet op het aantal rechten-lookups (standaard 2000, `0` = geen limiet) |
| `-TenantId` | * | Verplicht met `-AllSites`; anders afgeleid van `-SiteUrl` |
| `-ClientId` / `-CertificateThumbprint` / `-ClientSecret` / `-AppName` | Nee | Gebruik je eigen app-registratie |
| `-OutputPath` | Nee | CSV-pad (standaard: `C:\Temp\GraphSharePointFind_<timestamp>.csv`) |
| `-MaxRetries` | Nee | Retries bij throttling (429), standaard 5 |

**Voorbeelden**

```powershell
# Eén site plus de subsites, alles met "veiligheid" in de naam
.\Search-SharePointContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -Name "*veiligheid*" -IncludeSubsites

# Tenantbreed op volledige tekst: welke documenten noemen "salarisschaal"?
.\Search-SharePointContent.ps1 -AllSites -TenantId contoso.onmicrosoft.com `
    -Content "salarisschaal"

# Alles in de tenant met eigen rechten — begin met 25 sites
.\Search-SharePointContent.ps1 -AllSites -TenantId contoso.onmicrosoft.com `
    -Permissions Unique -MaxSites 25

# Onbeheerd, met een app die je al hebt
.\Search-SharePointContent.ps1 -AllSites -TenantId contoso.onmicrosoft.com `
    -ClientId 0000...-4444 -CertificateThumbprint A1B2C3 `
    -Name "*.pfx" -OutputPath C:\Reports\keys.csv
```

**Opmerkingen**
- Een tenantbrede delta-crawl leest elk item van elke bibliotheek die hij raakt. Begin met
  `-MaxSites`, en voeg `-IncludePersonalSites` alleen toe als je dat echt wilt — dat
  vermenigvuldigt het werk met het aantal gebruikers.
- Throttling (HTTP 429) wordt afgehandeld: losse calls en batch-subrequests worden opnieuw
  geprobeerd, met respect voor `Retry-After`.
- `-Content` doorzoekt de hele tenantindex en de resultaten worden daarna teruggefilterd naar
  de sites binnen het bereik, dus een query met heel veel treffers besteedt wat tijd aan
  resultaten die hij weggooit.

**Benodigde modules**
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser  # om te draaien
Install-Module Microsoft.Graph.Applications  -Scope CurrentUser   # alleen voor de eenmalige app-registratie
```

---

### Restore-RecycleBinItems.ps1

Toont de prullenbak van de eerste en/of tweede fase, filtert de items (naam, pad, wie ze
verwijderde, wanneer) en zet de treffers terug op hun oorspronkelijke locatie. Standaard een
proefdraai — er wordt niets teruggezet tot je `-Apply` meegeeft.

**Twee modi**

| Modus | Wat het doet |
|------|--------------|
| `-SiteUrl <url>` | Eén site collection — teamsite, communicatiesite of de OneDrive van een gebruiker (`https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com`) |
| `-AllSites -TenantUrl <url>` | Elke SharePoint-site in de tenant. **Persoonlijke OneDrive-sites worden uitgesloten** — zet die één voor één terug met `-SiteUrl` |

De tenantbrede ronde slaat ook de My Site-host over, redirect-sites, en sites die
vergrendeld zijn als alleen-lezen of zonder toegang (terugzetten mislukt daar toch). Het meldt
hoeveel van elk het oversloeg. Beperk het met `-SiteFilter "*/sites/Finance*"` en probeer het
uit met `-MaxSites 5` voordat je het op de hele tenant loslaat.

Een site die een fout geeft (geen toegang, gethrottled, verdwenen) wordt gemeld en de ronde
gaat door — één slechte site breekt de run niet af. Resultaten per site komen in de
samenvattingstabel en in de CSV, die in beide modi een kolom `Site` heeft.

Tenantbreed is waar een proefdraai zijn waarde bewijst: hij beantwoordt "op welke sites staan
bestanden die dit account dinsdag heeft verwijderd" zonder iets aan te raken.

**Snelheid — batches, geen threads**

Terugzetten gaat via `Restore-PnPRecycleBinItem -IdList`, dat SharePoint een hele set items in
één servercall geeft. 200 bestanden kosten dan één round trip in plaats van 200, en daar zit de
snelheid — het terugzetten van een volledige bibliotheek gaat van tientallen minuten naar een
paar minuten.

Terugzetten in parallelle runspaces is hier *niet* de snellere route, en het script doet dat
bewust niet:

- PnP PowerShell is niet thread-safe; elke runspace heeft zijn eigen module-import (seconden)
  en zijn eigen verbinding nodig, en verbindingsobjecten gaan niet netjes over runspace-grenzen
- Alle items belanden in dezelfde site collection, dus parallelle calls concurreren om dezelfde
  lijsten en triggeren SharePoint-throttling (HTTP 429) — je krijgt retries en gedeeltelijke
  mislukkingen, geen snelheid
- Eén gebatchte call doet server-side al wat de threads probeerden te paralleliseren

Batchgedrag:

- `-BatchSize` (1-200, standaard 200) bepaalt de batchgrootte; `-BatchSize 1` zet strikt item
  voor item terug
- Mappen en bestanden delen nooit een batch, zodat de volgorde mappen-eerst behouden blijft
- Een batch is alles-of-niets en de fout zegt niet welk item hem liet mislukken, dus een
  mislukte batch wordt automatisch opnieuw geprobeerd, één item tegelijk — de goede items komen
  alsnog terug en de CSV noemt de items die dat niet deden

**Hoe lang duurt het?**

Het script vertelt het je, zodat je weet of je moet wachten of koffie kunt halen:

- Het lezen van de prullenbak wordt per site getimed en gemeld — op een site met tienduizenden
  verwijderde items kan dat alleen al minuten duren (gebruik `-RowLimit` om het te begrenzen).
  Tenantbreed is dit lezen meestal het grootste deel van de looptijd, niet het terugzetten
- Tijdens `-Apply` toont de voortgangsbalk de verstreken tijd en een live ETA, herberekend uit
  het tempo dat tegen die tenant werkelijk gemeten wordt in plaats van een schatting vooraf.
  Met `-AllSites` zijn er twee balken: sites, en batches binnen de huidige site
- De samenvatting meldt de werkelijke duur en, tenantbreed, een tabel per site met gevonden,
  gematchte, teruggezette en mislukte items en de lees-/terugzetseconden per site
- De CSV heeft de kolommen `Site`, `Batch` en `DurationSeconds`, en daarmee vind je de trage
  exemplaren

**Aanmelden — de app-registratie wordt voor je aangemaakt**

PnP PowerShell levert geen gedeelde multi-tenant app meer mee, dus een Entra-app-registratie is
vereist. De eerste run tegen een tenant maakt er automatisch een aan:

1. Meld je aan bij Microsoft Graph als Global Administrator van de doeltenant
2. Het script maakt een public-client app aan (of hergebruikt die) met de naam uit `-AppName`
3. Het kent de delegated SharePoint-scope `AllSites.FullControl` toe, met admin consent
4. De client ID wordt per tenant gecachet in `pnp.appid.json` in de root van de repo (gitignored)

Latere runs lezen de gecachete client ID en gaan meteen naar de interactieve SharePoint-login —
de Graph-admin-aanmelding gebeurt maar één keer per tenant. Geef een bestaande `-ClientId` mee
om het aanmaken van de app helemaal over te slaan.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-SiteUrl` | * | URL van de site collection om in terug te zetten (teamsite, communicatiesite of OneDrive) |
| `-AllSites` | * | Loop in plaats daarvan elke SharePoint-site in de tenant af; OneDrive wordt uitgesloten |
| `-TenantUrl` | * | Tenant-URL voor `-AllSites`, bijv. `https://contoso.sharepoint.com` |
| `-SiteFilter` | Nee | Met `-AllSites`: alleen sites waarvan de URL overeenkomt met deze wildcard |
| `-MaxSites` | Nee | Met `-AllSites`: stop na zoveel sites |
| `-TenantId` | Nee | Tenant-ID/domein voor de eenmalige Graph-aanmelding (standaard: afgeleid van `-SiteUrl`) |
| `-ClientId` | Nee | Bestaande Entra-app om mee aan te melden — slaat de app-registratie over |
| `-AppName` | Nee | Weergavenaam van de app die aangemaakt/hergebruikt wordt (standaard: `M365-Scripts SharePoint Restore`) |
| `-Name` | Nee | Filter op itemnaam, wildcards toegestaan (`*.xlsx`, `Budget*`) |
| `-Path` | Nee | Filter op de oorspronkelijke locatie, deelstring-match (`Shared Documents/Finance`) |
| `-DeletedBy` | Nee | Filter op wie het verwijderde — weergavenaam of e-mail, wildcards toegestaan |
| `-DeletedAfter` / `-DeletedBefore` | Nee | Beperk tot een verwijderperiode |
| `-ItemType` | Nee | `All` (standaard), `File`, `Folder` of `ListItem` |
| `-Stage` | Nee | `All` (standaard), `FirstStage` (prullenbak van de gebruiker) of `SecondStage` (prullenbak van de site collection-beheerder) |
| `-RowLimit` | Nee | Limiet op het aantal prullenbakitems dat wordt opgehaald (gebruik bij enorme prullenbakken) |
| `-BatchSize` | Nee | Items per servercall, 1-200 (standaard 200). `1` = strikt één voor één |
| `-GrantSiteAdmin` | Nee | Voeg jezelf per site tijdelijk toe als site collection-beheerder — nodig voor de OneDrive van een andere gebruiker en in de praktijk vereist voor `-AllSites` |
| `-KeepSiteAdmin` | Nee | Behoud die rechten in plaats van ze na afloop weg te halen |
| `-AdminUpn` | Nee | UPN die site-admin wordt (standaard: het aangemelde account) |
| `-Apply` | Nee | Echt terugzetten. Zonder dit rapporteert het script alleen |
| `-OutputPath` | Nee | Pad van het CSV-rapport (standaard: `C:\Temp\RecycleBinRestore_<timestamp>.csv`) |
| `-Disconnect` | Nee | Meld af bij PnP als het klaar is |

**Voorbeelden**

```powershell
# Proefdraai — toon alles in beide prullenbakken van een site
.\Restore-RecycleBinItems.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance

# Alles terugzetten wat één gebruiker gisteravond heeft verwijderd
.\Restore-RecycleBinItems.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -DeletedBy "jane.doe@contoso.com" -DeletedAfter "2026-08-27 18:00" -Apply

# Excel-bestanden uit één bibliotheekmap terugzetten, alleen uit de tweede-fase-prullenbak
.\Restore-RecycleBinItems.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -Name "*.xlsx" -Path "Shared Documents/Budget" -Stage SecondStage -Apply

# De OneDrive van een gebruiker terugzetten, en jezelf tijdelijk site-admin maken
.\Restore-RecycleBinItems.ps1 `
    -SiteUrl https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com `
    -GrantSiteAdmin -Apply

# Tenantbrede proefdraai: op welke sites staan bestanden die dit account vandaag verwijderde?
.\Restore-RecycleBinItems.ps1 -AllSites -TenantUrl https://contoso.sharepoint.com `
    -DeletedBy "jane.doe@contoso.com" -DeletedAfter "2026-08-28" -GrantSiteAdmin

# Tenantbreed terugzetten na een bulkverwijdering — probeer eerst 5 sites
.\Restore-RecycleBinItems.ps1 -AllSites -TenantUrl https://contoso.sharepoint.com `
    -DeletedBy "jane.doe@contoso.com" -DeletedAfter "2026-08-28" `
    -GrantSiteAdmin -MaxSites 5 -Apply
```

**Opmerkingen**
- Items worden mappen-eerst en ondiepe-paden-eerst teruggezet: een bestand kan niet worden
  teruggezet zolang de map waarin het stond zelf nog verwijderd is.
- Terugzetten mislukt als er op de oorspronkelijke locatie al een item met dezelfde naam
  bestaat — die items worden in de CSV gemeld met de SharePoint-fout, verder wordt niets
  afgebroken.
- Items uit de tweede fase (site collection) worden direct teruggezet op hun oorspronkelijke
  locatie, niet in de prullenbak van de eerste fase.
- De standaard bewaartermijn is 93 dagen over beide fasen samen. Oudere items zijn weg en
  geen enkel script kan ze terughalen — dat is een limiet van het Microsoft-platform, niet van
  het script.
- Vereist SharePoint Administrator (of Global Administrator) voor `-GrantSiteAdmin` en voor de
  eenmalige app-registratie; het terugzetten zelf heeft alleen toegang tot de site nodig.

**Benodigde modules**
```powershell
Install-Module PnP.PowerShell -Scope CurrentUser              # PowerShell 7.4+
Install-Module Microsoft.Graph.Applications -Scope CurrentUser # alleen voor de eenmalige app-registratie
```

---

### Trace-SharePointFile.ps1

Beantwoordt "waar is mijn bestand gebleven?" voor OneDrive en SharePoint: hernoemd,
verplaatst, gekopieerd, verwijderd of teruggezet, door wie en wanneer. Elke tijd staat in
**Brusselse tijd** (zomer- en wintertijd inbegrepen, met de UTC-offset erbij), en je kunt
een periode meegeven. Alleen-lezen: er verandert niets in de tenant.

SharePoint onthoudt de oude naam of locatie van een bestand niet; het **Unified Audit Log**
wel, dus dat is de bron. Het script leest het en bouwt het spoor van het bestand opnieuw op:

| Stap | Wat het doet |
|------|--------------|
| Lezen | Elke hernoeming, verplaatsing, kopie, verwijdering, prullenbak-actie, terugzetting en upload van bestanden **en mappen** in de periode. Per dag gelezen; een stuk met meer dan de 50.000 records die één zoekopdracht kan teruggeven wordt gesplitst tot het past (tot 15 minuten), en een mislukte of inconsistente zoekopdracht wordt opnieuw geprobeerd |
| Startpunt | De records die het bestand noemen via `-Name`, `-Url` of `-ItemId` |
| Volgen | Elk record van hetzelfde item (`ListItemUniqueId`, dat hernoemen en verplaatsen overleeft) en elk record dat begint op een pad waarnaar het bestand hernoemd of verplaatst werd. Een keten `A → B → C` eindigt bij C, ook al lijkt C in niets op de naam waarop je zocht. Een verplaatsing naar een andere site wordt gevolgd via het doelpad |
| Mappen | Een map hernoemen, verplaatsen of verwijderen verplaatst elk bestand erin **zonder record per bestand**. Maprecords worden daarom afgespeeld op het pad van het bestand op dat moment, zodat "meeverhuisd met map X" en "mee verwijderd met map X" ook zichtbaar zijn |

Per item krijg je de tijdlijn, de **laatst bekende locatie** en een status: `Present`,
`In recycle bin` (terugzetten met [`Restore-RecycleBinItems.ps1`](#restore-recyclebinitemsps1)),
`In second-stage recycle bin` of `Permanently deleted`.

**Parameters**

| Parameter | Type | Standaard | Beschrijving |
|-----------|------|-----------|--------------|
| `-Name` | string | — | Bestandsnaam zoals die ooit was. Zonder extensie past elke extensie (`Offerte` vindt `Offerte.docx`); `*` en `?` zijn jokertekens |
| `-Url` | string | — | Volledige URL van het bestand zoals die was. `?web=1` wordt genegeerd en een deellink `/:w:/r/` wordt terug omgezet naar het pad. Ondoorzichtige deellinks (`/:w:/s/`, `/:w:/g/`) zijn niet te volgen — open ze en kopieer het adres waarop ze uitkomen |
| `-ItemId` | guid | — | De `ListItemUniqueId`, bv. uit de CSV van een eerdere run |
| `-SiteUrl` | string | — | Alleen records van deze site of OneDrive. Veel sneller in een grote tenant |
| `-StartDate` | datum of string | `-Days` voor `-EndDate` | Kloktijd in `-TimeZone`: `15-09-2026`, `15/09/2026 08:30`, `2026-09-15 08:30` (dag eerst, Belgisch) |
| `-EndDate` | datum of string | nu | Zelfde notatie. Een datum zonder tijd neemt die hele dag mee |
| `-Days` | int | `30` | Lengte van de periode als `-StartDate` niet is opgegeven |
| `-TimeZone` | string | `Europe/Brussels` | IANA- of Windows-ID; werkt in Windows PowerShell 5.1 en PowerShell 7 |
| `-FollowCopies` | switch | uit | Ook kopieën volgen. Standaard wordt een kopie gemeld maar niet gevolgd — het origineel blijft waar het was |
| `-IncludeActivity` | switch | uit | Ook openen, bewerken, downloaden, synchronisatie en in-/uitchecken: wie er het laatst in werkte. Veel meer records, dus trager |
| `-OutputPath` | string | `C:\Temp\FileTrail_<naam>_<ts>.csv` | CSV-pad; de ruwe auditrecords van het spoor gaan naar dezelfde naam met `.json` |
| `-TenantId` | string | — | Tenantdomein voor `Connect-ExchangeOnline`; niet nodig als je al verbonden bent |
| `-PassThru` | switch | uit | Geeft de tijdlijnrijen ook als objecten terug |

**Voorbeelden**

```powershell
# Waar is "Offerte Janssens.docx" de laatste 30 dagen gebleven?
.\Trace-SharePointFile.ps1 -Name "Offerte Janssens.docx" -TenantId contoso.onmicrosoft.com

# Een periode in Brusselse tijd, alleen één OneDrive
.\Trace-SharePointFile.ps1 -Name "Budget*" -StartDate '01-09-2026' -EndDate '15-09-2026' `
    -SiteUrl https://contoso-my.sharepoint.com/personal/jan_contoso_com

# Vanaf de link die iemand ooit stuurde, één namiddag
.\Trace-SharePointFile.ps1 -Url "https://contoso.sharepoint.com/sites/Sales/Shared Documents/2026/Prijslijst.xlsx" `
    -StartDate '2026-09-12 13:00' -EndDate '2026-09-12 18:00'
```

**Opmerkingen**

- Vereist de rol **View-Only Audit Logs** of **Audit Logs** in Exchange Online, en de module `ExchangeOnlineManagement`. Draait in Windows PowerShell 5.1 en PowerShell 7
- Het auditlog loopt 30–90 minuten (soms 24 uur) achter. Audit Standard bewaart **180 dagen**; het script waarschuwt als de periode vroeger begint
- Alleen wat **binnen** de periode gebeurde, kan gevolgd worden. Een map die vóór `-StartDate` hernoemd werd, is onzichtbaar; lijkt het spoor halverwege te beginnen, maak de periode dan groter
- Hernoemingen door de OneDrive-synchronisatieclient (in Verkenner) worden net als die in de browser geaudit; de kolom `UserAgent` onderscheidt ze
- De laatst bekende locatie is wat het auditlog zegt — het script controleert niet of het bestand er nog staat
- CSV-kolommen: `Item, Time, TimeUtc, Action, Operation, User, From, To, ViaFolder, ItemId, ClientIP, UserAgent, RecordId`. `ViaFolder` is ingevuld als de stap uit een mapactie kwam

---

### Revoke-SharePointUserAccess.ps1

De tegenhanger van [`Get-SharePointPermissionsReport.ps1`](../Reporting/readme.nl.md#get-sharepointpermissionsreportps1): dat script vertelt wie waar bij kan, dit script haalt het weg. Het zoekt elke plek waar één genoemde gebruiker toegang heeft en verwijdert die:

- **Site collection-beheerder** — als eerste, want die overschrijft elke roltoewijzing eronder; laten staan zou de rest cosmetisch maken
- **Directe roltoewijzingen** op site, sub-site, lijst/bibliotheek, map of los bestand
- **SharePoint-groepen** (Owners, Members, Visitors en eigen groepen)
- **Deellinks** — de `SharingLinks.*`-groepen waar een gedeelde link zijn ontvangers in zet. Dat is hoe "iedereen met de link" en "specifieke personen" een persoon daadwerkelijk toegang geven

Rapporteren is de standaard. Er verandert niets zonder `-Apply`, en elke run schrijft een CSV met precies wat er gevonden is en wat ermee gebeurd is.

#### Wat het bewust níét doet

| | Waarom |
|---|---|
| Entra ID-groepslidmaatschap wijzigen | **Tenzij je `-RemoveFromEntraGroups` meegeeft.** Standaard houdt wie via een security- of M365-groep binnenkomt die toegang — de groep *is* de toekenning — en worden die routes nadrukkelijk gemeld, mét groepsnaam, zodat je niet denkt dat het dicht is terwijl het openstaat |
| `Everyone` / `Everyone except external users` verwijderen | Dat ontneemt de hele tenant toegang, niet deze persoon. Wordt gemeld, niet aangeraakt |
| Eigenaarschap en metadata opschonen | Een ingetrokken gebruiker blijft de auteur van wat die gemaakt heeft |

> **Offboarding is dus twee stappen.** Draai dit script, en werk daarna de Entra-groepen af die in de CSV onder `Action = CannotRevoke` staan. Zonder die tweede stap is de toegang niet weg.

#### Robuustheid

Dit script verwijdert rechten, dus de faalmodi zijn andere dan bij een rapport: stil de verkeerde persoon raken, of niet kunnen navertellen wat je hebt weggehaald.

| Situatie | Gedrag |
|---|---|
| Gebruiker identificeren | **Alleen exacte vergelijkingen.** Op UPN, e-mail, het claim-achtervoegsel, en de gedecodeerde gastnaam (`jan_partner.com#ext#@tenant` wordt `jan@partner.com`). Nooit op deelstring: `an@contoso.com` zit in `jan@contoso.com`, en dat is precies hoe je de verkeerde persoon intrekt |
| Twee accounts met hetzelfde adres | De site wordt **niet** aangeraakt; het script stopt met beide loginnamen in de fout. Kiezen is aan jou, niet aan het script |
| Audit-CSV | Regel voor regel weggeschreven tijdens de run, niet aan het eind. Een run die tweehonderd dingen intrekt en dan crasht, moet nog steeds kunnen navertellen wát er weg is |
| CSV staat open in Excel | Vijf pogingen met oplopende wachttijd, daarna stopt de run — liever een afgebroken run dan rechten verwijderen zonder spoor |
| Verwijdering geeft `404` | `AlreadyGone`, geen fout. Bij een tweede run is dat de normale uitkomst; als fout geteld zou een schone run kapot lijken |
| `-WhatIf` | Zelfde tak als een proefdraai, dus `WouldRevoke` in de CSV — niet `Skipped`, wat zou suggereren dat iemand een prompt heeft geweigerd |
| Site collection-beheerder niet te verwijderen | **Luid gemeld, en de run telt het als mislukt.** Die rol bereikt elke scope in de site, dus alle andere verwijderingen daar zijn dan cosmetisch |
| Throttling (`429`/`503`) | Opnieuw proberen met `Retry-After`; een geweigerd token wordt één keer vers opgehaald voordat de run stopt |
| Onverwachte fout | Een `trap` ruimt de tijdelijke Full Control-app op voordat het script stopt |

#### Parameters

| Parameter | Type | Standaard | Omschrijving |
|---|---|---|---|
| `-UserPrincipalName` | string | — | **Verplicht.** De gebruiker, bijv. `jan@contoso.com`. Voor een gast mag ook het echte adres (`jan@partner.com`) — het script vindt de `#ext#`-variant zelf |
| `-TenantUrl` | string | — | Tenant-root. Verplicht voor een tenantbrede run |
| `-SiteUrl` | string | — | Eén site collection in plaats van de hele tenant |
| `-Apply` | switch | uit | Daadwerkelijk intrekken. Zonder dit alleen rapporteren |
| `-Scope` | `Site`/`List`/`Item` | `Item` | Hoe diep naar directe toekenningen wordt gezocht |
| `-IncludeGroupAccess` | switch | uit | Meldt ook de sites die de gebruiker via Entra-groepen bereikt, óók waar die verder niets heeft. Alleen rapporteren |
| `-KeepSharingLinks` | switch | uit | Deellinks met rust laten; alle andere routes worden wel ingetrokken |
| `-RemoveFromEntraGroups` | switch | uit | **Verwijdert de gebruiker ook uit de Entra ID-groepen die toegang bleken te geven** — alleen die, nooit elke groep waar iemand in zit. Vereist Graph `GroupMember.ReadWrite.All`, die de tijdelijke app alleen met deze schakelaar vraagt. Zie hieronder |
| `-RemoveFromSite` | switch | uit | Verwijdert de gebruiker daarna ook uit de gebruikerslijst van elke site collection. Vangt wat de scope-voor-scope-ronde niet zag, maar de naam rendert daarna als verwijderd account in oudere metadata |
| `-IncludeOneDriveSites` | switch | uit | Ook persoonlijke OneDrive-sites doorzoeken |
| `-IncludeHiddenLists` | switch | uit | Ook verborgen en systeemlijsten |
| `-TenantId` / `-ClientId` / `-CertificateThumbprint` | string | — | Eigen app-registratie in plaats van de tijdelijke. Die heeft SharePoint `Sites.FullControl.All` nodig plus Graph `Sites.Read.All`, `User.Read.All` en `GroupMember.Read.All`. Zonder `User.Read.All` geeft de gebruikerslookup `403` en stopt de run, in plaats van dat aan te zien voor een niet-bestaand account |
| `-ClientSecret` | string | — | Werkt voor Graph maar **niet** voor SharePoint (zie authenticatie bij het rapport) |
| `-OutputPath` | string | `C:\Temp` | Outputmap |
| `-GraphTimeoutSec` / `-MaxGraphRetry` | int | `120` / `6` | Timeout en retries |

Authenticatie is identiek aan het rapport: een kortlevende, certificaat-gebaseerde app-registratie met SharePoint `Sites.FullControl.All`, die na afloop weer wordt verwijderd.

#### Ook de Entra ID-groepen verwijderen

`-RemoveFromEntraGroups` maakt de tweede helft van een offboarding af in plaats van hem alleen te melden. **Alleen de groepen die deze run daadwerkelijk een roltoewijzing zag houden op een scope binnen bereik worden aangeraakt** — nooit elke groep waar de gebruiker in zit. Iemand die vertrekt zit vaak in vijftig groepen; alleen die SharePoint-toegang geven zijn hier in scope.

> **Dit reikt verder dan SharePoint.** Een Entra-groep is geen SharePoint-object. Datzelfde lidmaatschap draagt vaak een Teams-team, een mailbox, licenties en app-toewijzingen — niets daarvan ziet dit rapport. Lees eerst het rapport van een run zónder `-Apply`, draai daarna pas met de schakelaar.

> **De lijst is net zo volledig als de scan.** Een run op één site, een versmalde `-Scope`, uitgesloten OneDrive- of verborgen lijsten en onleesbare scopes maken hem kleiner — een groep die toegang geeft op een plek die nooit doorzocht is, staat er niet in. De run noemt elke beperking vóór er iets verwijderd wordt en opnieuw in de samenvatting. Het revoke-script draait zijn eigen scan, dus het rechtenrapport is geen voorwaarde.

Vier gevallen worden gemeld in plaats van afgedwongen, omdat afdwingen óf zou falen óf het verkeerde zou doen:

| Geval | Waarom het blijft staan |
|---|---|
| Dynamische groep | Lidmaatschap volgt een regel en wordt niet opgeslagen, dus er valt niets te verwijderen. Pas de regel aan, of de gebruikerskenmerken waarop die matcht |
| Gesynchroniseerd uit on-premises AD | Alleen-lezen in de cloud. Het lidmaatschap moet in Active Directory weg |
| Lid via een geneste groep | De gebruiker is geen direct lid, dus verwijderen zou hier mislukken. De toegang moet worden afgesneden bij de groep die hem écht bevat — het rapport noemt die |
| Gebruiker niet gevonden in Entra | Er is niets om hem uit te verwijderen; de SharePoint-kant draait gewoon door |

De verwijdering loopt door dezelfde trechter als elke andere wijziging, dus `-Apply`, `-WhatIf`, de bevestigingsvraag en de audit-CSV gedragen zich identiek. Vereist Graph `GroupMember.ReadWrite.All`, die de tijdelijke app **alleen** vraagt als de schakelaar is meegegeven — een rapportage-run houdt geen permissie die groepslidmaatschap kan wijzigen.

#### Uitvoer

`SharePoint_Revoke_<user>_<ts>.csv`, één regel per gevonden toekenning, met kolom `Action`:

| Action | Betekenis |
|---|---|
| `WouldRevoke` | Gevonden, en zou verwijderd worden — dit is wat je zonder `-Apply` krijgt |
| `Revoked` | Verwijderd |
| `Failed` | Poging mislukt; de reden staat in `Detail` |
| `AlreadyGone` | Er viel niets meer te verwijderen — bij een tweede run de normale uitkomst |
| `CannotRevoke` | Via een Entra-groep of `Everyone` — moet elders opgelost worden |
| `Kept` | Bewust laten staan door `-KeepSharingLinks` |
| `Skipped` | Bij de bevestigingsvraag geweigerd |

#### Voorbeelden

```powershell
# Wat kan Jan allemaal bereiken? Verandert niets
.\Revoke-SharePointUserAccess.ps1 -UserPrincipalName jan@contoso.com -TenantUrl "https://contoso.sharepoint.com"

# Hetzelfde, en nu daadwerkelijk intrekken
.\Revoke-SharePointUserAccess.ps1 -UserPrincipalName jan@contoso.com -TenantUrl "https://contoso.sharepoint.com" -Apply

# Een gast uit één site collection halen, deellinks inbegrepen
.\Revoke-SharePointUserAccess.ps1 -UserPrincipalName gast@partner.com -SiteUrl "https://contoso.sharepoint.com/sites/Finance" -Apply

# Offboarding-checklist: ook de Entra-groepen die toegang geven
.\Revoke-SharePointUserAccess.ps1 -UserPrincipalName jan@contoso.com -TenantUrl "https://contoso.sharepoint.com" -IncludeGroupAccess
```

> Onbeheerd draaien? Geef `-Confirm:$false` mee, anders vraagt het script per verwijdering om bevestiging (`ConfirmImpact = 'High'`).

---

### Test-SharePointAccessScripts.ps1

Controleert `Revoke-SharePointUserAccess.ps1` en `Get-SharePointPermissionsReport.ps1` zonder een tenant aan te raken. Draai het na elke wijziging aan één van beide; exitcode 0 betekent dat beide in orde zijn.

Twee dingen worden gecontroleerd, allebei fouten die in productie geruisloos misgaan:

1. **Het gedeelde authenticatieblok is byte-identiek.** Beide scripts bevatten dezelfde app-only auth- en SharePoint REST-laag, afgebakend met `SHARED BLOCK START/END`. Die laag kostte vier live-runs tegen een tenant om goed te krijgen — certificaat in plaats van secret, tokens die hun app-rollen moeten aantonen vóór ze gecachet worden, 401 als fataal in plaats van per site, paging die niet kan blijven hangen. Een tweede kopie die stilletjes afdrijft is een correctheidsrisico in júist het script dat rechten verwijdert. Bij verschil wordt de eerste afwijkende regel getoond.
2. **De revocatie-trechter gedraagt zich.** Een proefdraai moet zijn voornemen vastleggen en niets uitvoeren, `-Apply` moet uitvoeren én vastleggen, een mislukking moet in het audit-spoor belanden in plaats van te verdwijnen, en de toekenningen die het script moet weigeren te verwijderen (via een Entra-groep, of aan iedereen) moeten geweigerd blijven.

```powershell
.\Test-SharePointAccessScripts.ps1
```
