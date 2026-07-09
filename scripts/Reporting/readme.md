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

### Performance (version history lookups)

Version history is de duurste stap: van nature 1 Graph-call per bestand. Drie optimalisaties beperken dat:

- **Overgeslagen wanneer versiebeheer uit staat** — is voor een library met zekerheid bekend dat versiebeheer uitstaat, dan wordt er geen version-call per bestand gedaan (0 versies is dan toch het antwoord). Bij onbekende status (fallback via `Get-MgSiteDrive`) wordt uit voorzichtigheid altijd nog opgehaald.
- **Batched via Graph's `$batch`-endpoint** — version-lookups voor bestanden in een library worden nu in groepen van 20 in 1 HTTP-call opgehaald, in plaats van 1 losse call per bestand.
- **Kortere retry voor deze specifieke calls** — max. 3 pogingen met een korte backoff (in plaats van de standaard `-MaxGraphRetry`/backoff die voor kritieke calls tot ~2 minuten per poging kan oplopen). Een mislukte version-lookup valt terug op "0 versies" in plaats van de hele scan op te houden.

`-SkipVersions` blijft de snelste optie als versiehistorie niet nodig is — dan wordt er helemaal geen version-call gedaan.

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
