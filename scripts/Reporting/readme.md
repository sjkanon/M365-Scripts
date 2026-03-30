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
