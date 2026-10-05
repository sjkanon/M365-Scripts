[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **DNS**

# DNS-scripts

Scripts om DNS-records te resolven en te importeren in Active Directory-geïntegreerde DNS-zones.

---

## Scripts

| Script | Omschrijving |
|--------|--------------|
| [`Import-DnsRecords.ps1`](Import-DnsRecords.ps1) ([docs](#import-dnsrecordsps1)) | FQDN's uit een CSV resolven via Google DNS en de records optioneel importeren in een AD-geïntegreerde DNS-zone (voorbeeldinvoer: [`example-records.csv`](example-records.csv)) |

---

### Import-DnsRecords.ps1

Leest een lijst FQDN's uit een CSV, resolvet elk ervan via Google DNS (8.8.8.8) met `dig`, en importeert de resultaten optioneel in een Active Directory-DNS-zone.

**Resolutielogica per FQDN**

1. Controleer op CNAME → indien gevonden is het recordtype CNAME
2. Controleer op A → indien gevonden is het recordtype A
3. Geen antwoord → gerapporteerd als niet te resolven, overgeslagen

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-CsvPath` | Ja | Pad naar het CSV-bestand met een kolom `FQDN` |
| `-ExportCsv` | Nee | Geresolvede records opslaan in een CSV om handmatig te controleren |
| `-ExportPath` | Nee | Eigen pad voor de export-CSV (impliceert `-ExportCsv`). Standaard: `C:\Temp\` / `~/Downloads\` |
| `-Apply` | Nee | Geresolvede records naar AD DNS schrijven (vereist Windows + de module DnsServer) |
| `-ZoneName` | Alleen met `-Apply` | AD-DNS-zone waaraan de records worden toegevoegd (bijv. `contoso.com`) |
| `-DnsServer` | Nee | DNS-server waarnaar geschreven wordt (standaard: `localhost`) |
| `-Ttl` | Nee | TTL in seconden (standaard: `3600`) |

**CSV-formaat**

```csv
FQDN
mail.contoso.com
webmail.contoso.com
portal.contoso.com
www.contoso.com
```

**Voorbeelden**

```powershell
# Resolven en op het scherm tonen — werkt op macOS/Linux
.\Import-DnsRecords.ps1 -CsvPath .\records.csv

# Resolven en exporteren naar CSV voor handmatige import
.\Import-DnsRecords.ps1 -CsvPath .\records.csv -ExportCsv

# Resolven en direct importeren in AD DNS
.\Import-DnsRecords.ps1 -CsvPath .\records.csv -ZoneName contoso.com -Apply

# DNS-server op afstand
.\Import-DnsRecords.ps1 -CsvPath .\records.csv -ZoneName contoso.com -DnsServer dc01.contoso.com -Apply
```

**Vereisten**

- `dig` — installeer via `choco install bind-toolsonly` of [isc.org/download](https://www.isc.org/download/)
- Module `DnsServer` — alleen nodig met `-Apply` (RSAT of de Windows-rol DNS Server)

**Opmerkingen**

- Slaat records over die al bestaan — veilig om opnieuw uit te voeren
- FQDN's die niet bij `-ZoneName` horen, worden bij gebruik van `-Apply` met een waarschuwing overgeslagen
- CNAME wordt eerst gedetecteerd; het A-record dient als terugval
