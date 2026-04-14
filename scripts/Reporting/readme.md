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

Rapporteert opslaggebruik over SharePoint Online met een tenantbrede scan.

### Dekking

- Alle SharePoint site collections (OneDrive personal sites worden uitgesloten)
- Onderliggende sub-sites op alle niveaus
- Teams-gerelateerde SharePoint locaties:
    - Standaard channels als libraries/folders in de parent Teams-site
    - Private/shared channels als aparte site collections
- Onderliggende mappen en bestanden in document libraries (alleen met `-Apply`)
- Detailoutput bevat zowel folders als files (`ItemType`) zodat je volledige structuur ziet

### Belangrijkste parameters

| Parameter | Omschrijving |
|---|---|
| `-Apply` | Volledige recursieve scan van libraries, mappen en bestanden. Zonder deze switch alleen quota-samenvatting. |
| `-SkipVersions` | Neemt versiehistorie niet mee (sneller). |
| `-SiteUrl` | Scan 1 specifieke site (`/sites/...` of `/teams/...`). |
| `-UseHighPrivilege` | Auto mode: kent tijdelijk `Sites.FullControl.All` toe i.p.v. `Sites.Read.All` wanneer read-only rechten niet voldoende blijken. |
| `-ClientId/-TenantId` | Gebruik eigen app-registratie (met passende Graph application permissions). |

### Voorbeelden

```powershell
# Volledige tenantscan inclusief sub-sites en bestanden
.\Get-SharePointStorageReport.ps1 -Apply

# Idem, maar met hogere tijdelijke app-rechten indien nodig
.\Get-SharePointStorageReport.ps1 -Apply -UseHighPrivilege

# Enkel een specifieke Teams-site
.\Get-SharePointStorageReport.ps1 -SiteUrl "https://contoso.sharepoint.com/teams/Operations" -Apply
```
