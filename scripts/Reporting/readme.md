# Reporting Scripts

Scripts voor het genereren van rapporten over Active Directory, devices en licenties.

---

## Get-ComputerLastLogon.ps1

Rapporteert de **laatste inlogdatum** van computerobjecten in één of meerdere OUs, met export naar CSV.

### Werking

Twee nauwkeurigheidsmodi:

| Modus | Attribuut | Vertraging | Snelheid |
|---|---|---|---|
| Standaard | `LastLogonTimestamp` (gerepliceerd) | max. 14 dagen | Snel |
| `-AllDCs` | `LastLogon` per DC, beste waarde | Geen | Trager |

Gebruik de standaardmodus voor stale-device rapportages. Gebruik `-AllDCs` als absolute nauwkeurigheid vereist is.

### Parameters

| Parameter | Type | Standaard | Omschrijving |
|---|---|---|---|
| `-SearchBase` | `string[]` | _(heel domein)_ | Een of meerdere OU distinguished names |
| `-AllDCs` | switch | uit | Bevraagt alle DC's voor nauwkeurigste `LastLogon` |
| `-InactiveDays` | int | `90` | Drempel in dagen waarna een computer als _Stale_ geldt |
| `-IncludeDisabled` | switch | uit | Neemt uitgeschakelde computerobjecten ook mee |
| `-ExportPath` | string | `C:\Temp\` | Map voor het CSV-bestand |

### Vereisten

- ActiveDirectory PowerShell module (RSAT)
- Leesrechten op de opgegeven OUs

### Voorbeelden

```powershell
# Laptops en Computers OU
.\Get-ComputerLastLogon.ps1 `
    -SearchBase "OU=Laptops,OU=Computers,DC=bedrijf,DC=local",
               "OU=Computers,DC=bedrijf,DC=local"

# Nauwkeurigste modus — bevraagt alle DC's
.\Get-ComputerLastLogon.ps1 -SearchBase "OU=Computers,DC=bedrijf,DC=local" -AllDCs

# Inclusief uitgeschakelde computers, drempel op 60 dagen
.\Get-ComputerLastLogon.ps1 -SearchBase "OU=Computers,DC=bedrijf,DC=local" `
    -IncludeDisabled -InactiveDays 60
```

### Statuswaarden

| Status | Betekenis |
|---|---|
| `Active` | LastLogon binnen de `-InactiveDays` drempel |
| `Active (pwd recent)` | LastLogon ziet er stale uit door replicatievertraging, maar het computeraccount-wachtwoord werd < 35 dagen geleden vernieuwd — device is online |
| `Stale` | Zowel LastLogon als PasswordLastSet overschrijden de drempel — vermoedelijk echt inactief |
| `Never` | Nooit ingelogd én geen recent wachtwoord |
| `Disabled` | Account uitgeschakeld in AD |

> **Tip:** `Active (pwd recent)` zijn PC's die wél actief zijn maar door de 9-14 daagse replicatievertraging van `LastLogonTimestamp` ten onrechte als stale verschijnen. Gebruik `-AllDCs` voor exacte gegevens als dit onderscheid kritisch is.

### CSV-kolommen

| Kolom | Omschrijving |
|---|---|
| `Name` | Computernaam |
| `Status` | Zie statuswaarden hierboven |
| `Enabled` | True/False |
| `LastLogon` | Laatste inlogdatum (dd/MM/yyyy HH:mm) |
| `DaysSinceLogon` | Aantal dagen geleden |
| `PasswordLastSet` | Datum laatste wachtwoordwijziging computeraccount (dd/MM/yyyy) |
| `DaysSincePasswordSet` | Dagen geleden dat het wachtwoord vernieuwd werd |
| `OperatingSystem` | OS-naam |
| `OperatingSystemVersion` | OS-versie |
| `IPv4Address` | IP-adres (indien beschikbaar) |
| `OU` | OU-pad (leesbaar formaat) |
| `Created` | Aanmaakdatum in AD |
| `Description` | Omschrijving uit AD |
| `DistinguishedName` | Volledig AD-pad |

---

## Licensing/

Zie [Licensing/](Licensing/) voor het maandelijkse licentie-rapport.

---

## Get-SharePointStorageReport.ps1

Rapporteert opslaggebruik over SharePoint Online met een tenantbrede scan. Standaard connecteert het script delegated en maakt het tijdelijk een App Registration (`Sites.Read.All`) aan voor site-enumeratie; die app wordt na afloop weer verwijderd.


### Dekking

- Alle SharePoint site collections (OneDrive personal sites worden uitgesloten)
- Onderliggende sub-sites op alle niveaus
- Teams-gerelateerde SharePoint locaties:
    - Standaard channels als libraries/folders in de parent Teams-site
    - Private/shared channels als aparte site collections
- Onderliggende mappen en bestanden in document libraries (alleen met `-Apply`)
- Detailoutput bevat zowel folders als files (`ItemType`) zodat je volledige structuur ziet
- In `-Apply` mode wordt alles in 1 ranked CSV gezet (grootste folders + files, inclusief version history)
- CSV bevat ook `Level` (diepte): root = `0`, topfolder = `1`, etc.
- CSV bevat ook `ParentPath` voor hiërarchische analyses (Excel/Power BI tree-opbouw)

### Prullenbak (recycle bin)

De prullenbak (stage 1 + stage 2) telt mee voor de tenant-opslagquota en wordt daarom **apart** van de library-scan opgehaald, alleen voor echte SharePoint site collections (geen OneDrive):

- Standaard (`-Apply`, als Phase 2b) of los via **`-RecycleBinOnly`** (slaat de library-scan helemaal over, alleen prullenbak)
- Alleen root site collections hebben een eigen prullenbak (sub-webs delen die van de root)
- Output: extra rij per site in de summary-CSV (`Library = "Recycle Bin (stage 1 + 2)"`) plus een rij per verwijderd item in de detail-CSV

```powershell
# Alleen prullenbak
.\Get-SharePointStorageReport.ps1 -RecycleBinOnly

---

## Remove-SharePointFileVersionsByDate.ps1

Rapporteert of verwijdert **oude bestandsversies** in SharePoint Online document libraries op basis van een cutoff-datum, terwijl de **huidige versie behouden blijft**.

### Gedrag

- Standaard: alleen preview/reporting
- Met `-Apply`: verwijdert matching vorige versies echt
- Werkt op één site of tenantbreed over alle sites
- Standaard geen OneDrive-sites en geen hidden libraries
- Gebaseerd op Microsoft Graph (`Invoke-MgGraphRequest`) — **geen** `PnP.PowerShell` en **geen** eigen Entra app-registratie nodig voor het standaardgeval

### Authenticatie

Standaard verbindt het script interactief (delegated) met `Sites.ReadWrite.All` + `Files.ReadWrite.All` via `Connect-MgGraph` — dat gebruikt Microsoft's eigen voorgeconsente app, dus zonder eigen App Registration of `-ClientId`. Alleen een **tenantbrede scan** (geen `-SiteUrl`) heeft daarnaast een kortstondige, read-only tijdelijke App Registration nodig (`Sites.Read.All`) om alle sites op te sommen — Microsoft ondersteunt tenantbrede site-enumeratie niet delegated. Die tijdelijke app wordt na afloop weer verwijderd; alle daadwerkelijke file-reads en version-deletes lopen altijd via je eigen delegated permissies, nooit via die tijdelijke app.

Wil je de tijdelijke app overslaan en je eigen bestaande app-registratie gebruiken? Geef dan `-ClientId` + `-TenantId` + `-ClientSecret` (of `-CertificateThumbprint`) mee; die app moet dan al `Sites.ReadWrite.All` application permission hebben.

> **Let op:** het verwijderen van een specifieke versie (`DELETE .../versions/{id}`) staat niet in Microsoft's officiële Graph API-referentie, maar is een breed gebruikte en bevestigd werkende operatie (zowel voor OneDrive als SharePoint document libraries). De huidige/laatste versie kan hiermee niet verwijderd worden — Graph weigert dat, wat precies de behouden-huidige-versie garantie is.

### Parameters

| Parameter | Type | Omschrijving |
|---|---|---|
| `-BeforeDate` | `datetime` | Verwijder versies ouder dan deze datum |
| `-SiteUrl` | `string` | Optioneel: scan één site |
| `-TenantUrl` | `string` | Vereist voor all-sites scan, bv. `https://contoso.sharepoint.com` |
| `-TenantId` | `string` | Entra ID tenant ID — automatisch gedetecteerd indien niet opgegeven; verplicht in combinatie met `-ClientId` |
| `-ClientId` | `string` | Bestaande App Registration client ID — slaat de tijdelijke app over; gebruik samen met `-TenantId` en `-ClientSecret` of `-CertificateThumbprint` |
| `-ClientSecret` | `string` | Client secret voor een bestaande app registration |
| `-CertificateThumbprint` | `string` | Certificate thumbprint voor een bestaande app registration |
| `-Apply` | `switch` | Voert de verwijdering echt uit |
| `-IncludeOneDriveSites` | `switch` | Neemt OneDrive-sites mee in tenantscan |
| `-IncludeHiddenLibraries` | `switch` | Neemt hidden document libraries mee |
| `-LibraryTitle` | `string[]` | Optionele filter op librarytitel |
| `-GraphTimeoutSec` | `int` | Timeout in seconden per Graph-call (standaard: `120`) |
| `-MaxGraphRetry` | `int` | Max. aantal retries bij Graph throttling/timeouts (standaard: `6`) |

### Voorbeelden

```powershell
# Preview tenantbreed: alles ouder dan 1 januari 2025
.\Remove-SharePointFileVersionsByDate.ps1 `
    -TenantUrl "https://contoso.sharepoint.com" `
    -BeforeDate "2025-01-01"

# Echt verwijderen op één site
.\Remove-SharePointFileVersionsByDate.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/Finance" `
    -BeforeDate "2025-01-01" `
    -Apply
```

# Volledige scan + prullenbak als extra fase
.\Get-SharePointStorageReport.ps1 -Apply
```

### Site collection totalen (vergelijken met het adminportaal)

Sub-sites en Teams-kanalen delen de opslagquota van hun root site collection, maar worden in de scan als **losse site-rijen** gerapporteerd; de prullenbak wordt weer in een **eigen rij** bijgehouden. Los van elkaar zijn die cijfers dus niet 1-op-1 te vergelijken met het ene "storage used"-getal dat het SharePoint-adminportaal per site collection toont.

Bij `-Apply` (Phase 2c) telt het script daarom alles automatisch weer bij elkaar op per root site collection: de library-totalen van alle onderliggende sub-sites/kanalen + de prullenbak van die site collection. Output: `SharePoint_SiteCollectionTotals_<timestamp>.csv`, met per site collection `LibrariesMB`, `RecycleBinMB`, `GrandTotalMB`/`GrandTotalGB` en het aantal sub-sites/kanalen dat is meegeteld. De console toont ook de top 10.

Wijkt `GrandTotalGB` voor een site nog steeds af van het adminportaal-cijfer, dan is de meest waarschijnlijke oorzaak een van:
- **Timing** — het adminportaal-cijfer kan tot 24u vertraagd zijn t.o.v. een live scan
- **Stil overgeslagen mappen** — een `[ERROR] Cannot read folder`-melding in de console betekent dat die submap-boom (permissieprobleem) niet is meegeteld
- **Mislukte version-lookups** — vallen terug op "0 versies" bij herhaalde Graph-fouten (zeldzaam, alleen na 3 mislukte retries)
- Vergelijk eerst zonder `-Apply` (quick mode) — dat gebruikt hetzelfde officiële `quota.used`-cijfer als het adminportaal, dus wijkt dat ook al af, dan zit het verschil niet in de `-Apply`-telling zelf

### Performance (version history lookups)

Version history is de duurste stap: van nature 1 Graph-call per bestand. Drie optimalisaties beperken dat:

- **Overgeslagen wanneer versiebeheer uit staat** — is voor een library met zekerheid bekend dat versiebeheer uitstaat, dan wordt er geen version-call per bestand gedaan (0 versies is dan toch het antwoord). Bij onbekende status (fallback via `Get-MgSiteDrive`) wordt uit voorzichtigheid altijd nog opgehaald.
- **Batched via Graph's `$batch`-endpoint** — version-lookups voor bestanden in een library worden nu in groepen van 20 in 1 HTTP-call opgehaald, in plaats van 1 losse call per bestand.
- **Kortere retry voor deze specifieke calls** — max. 3 pogingen met een korte backoff (in plaats van de standaard `-MaxGraphRetry`/backoff die voor kritieke calls tot ~2 minuten per poging kan oplopen). Een mislukte version-lookup valt terug op "0 versies" in plaats van de hele scan op te houden.

`-SkipVersions` blijft de snelste optie als versiehistorie niet nodig is — dan wordt er helemaal geen version-call gedaan.

Daarnaast is `-SiteUrl` (1 specifieke site) geoptimaliseerd: in normale mode gebruikt het script direct delegated Graph-calls voor alleen die site, wat de opstarttijd gelijk trekt met andere commando's.

Voor GDAP-betrouwbaarheid schakelt het script bij single-site scans automatisch naar app-only bootstrap wanneer `authMode=GDAP` is gedetecteerd (uit `load.config.ps1`/launcher context). Wil je dat altijd forceren, gebruik dan `-ForceAppOnlySingleSite`.

Voor full-site scans in GDAP gebruikt het script dezelfde customer-tenant context (`$global:cid`/`-TenantId`) voor zowel `Connect-MgGraph` als de tijdelijke app-bootstrap, zodat consent en site-enumeratie altijd in de juiste tenant plaatsvinden.

### Parameters

| Parameter | Omschrijving |
|---|---|
| `-SiteUrl` | Scan 1 specifieke site. Tenant-root URL (bijv. `https://contoso.sharepoint.com`) triggert automatisch een tenantbrede scan |
| `-SkipVersions` | Neemt versiehistorie niet mee (sneller) |
| `-OutputPath` | Overschrijft de standaard outputmap (`C:\Temp\` / `~/Downloads/`) |
| `-TenantId` | Entra ID tenant ID — automatisch gedetecteerd indien niet opgegeven; verplicht in combinatie met `-ClientId` |
| `-ClientId` | Bestaande App Registration client ID — slaat auto-create over; gebruik samen met `-TenantId` en `-ClientSecret` of `-CertificateThumbprint` |
| `-ClientSecret` | Client secret voor een bestaande app registration |
| `-CertificateThumbprint` | Certificate thumbprint voor een bestaande app registration |
| `-Apply` | Volledige recursieve scan van libraries, mappen en bestanden. Zonder deze switch alleen quota-samenvatting |
| `-UseHighPrivilege` | Auto mode: kent tijdelijk `Sites.FullControl.All` toe i.p.v. `Sites.Read.All` wanneer read-only rechten niet voldoende blijken |
| `-RecycleBinOnly` | Slaat storage/library scanning over — leest alleen recycle bin items (stage 1 + stage 2) per site collection |
| `-ForceAppOnlySingleSite` | Forceert tijdelijke app-bootstrap voor `-SiteUrl` scans (handig voor GDAP/delegated beperkingen) |
| `-GraphTimeoutSec` | Timeout in seconden per Graph-call (standaard: `120`) |
| `-MaxGraphRetry` | Max. aantal retries bij Graph throttling/timeouts (standaard: `6`) |

### Voorbeelden

```powershell
# Snelle samenvatting — alleen site quota, geen file scan
.\Get-SharePointStorageReport.ps1

# Volledige tenantscan inclusief sub-sites en bestanden (auto app registration)
.\Get-SharePointStorageReport.ps1 -Apply

# Idem, maar met hogere tijdelijke app-rechten indien nodig
.\Get-SharePointStorageReport.ps1 -Apply -UseHighPrivilege

# Enkel een specifieke Teams-site
.\Get-SharePointStorageReport.ps1 -SiteUrl "https://contoso.sharepoint.com/teams/Operations" -Apply

# Volledige scan met een bestaande app registration
.\Get-SharePointStorageReport.ps1 -Apply -ClientId "..." -TenantId "..." -ClientSecret "..."
```
