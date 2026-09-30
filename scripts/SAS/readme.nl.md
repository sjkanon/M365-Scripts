[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **SAS**

# Foutmonitoring van SAS-batchjobs

Controleert de logs van SAS-batchjobs en de Windows Event Viewer op fouten, met optionele Zabbix-integratie en e-mailmeldingen.

---

## Bestanden

| Bestand | Omschrijving |
|------|-------------|
| [`Monitor-SASBatchErrors.ps1`](Monitor-SASBatchErrors.ps1) | Hoofdscript — doorzoekt logbestanden en Event Viewer |
| [`Setup-SASMonitoring.ps1`](Setup-SASMonitoring.ps1) | Eenmalige setup — installeert het script, de geplande taak en de Zabbix-configuratie |
| [`Test-SASWorkDirectory.ps1`](Test-SASWorkDirectory.ps1) | Controleert de gezondheid en de rechten van de SAS WORK-map |
| `rca.md` | Oorzaakanalyse van periodieke access-denied-fouten bij het verwijderen in SAS WORK |
| `zabbix_sas_monitor.conf` | Voorbeeldconfiguratie voor Zabbix UserParameter |

---

## Gedetecteerde fouten

| Type | Ernst | Patroon |
|------|----------|---------|
| `SpawnError` | Critical | `Can't spawn "sas.bat"` |
| `WorkLibAuth` | Critical | Autorisatiefout op de WORK-library |
| `SASAbort` | Critical | `SAS has ABORTED processing` |
| `DiskError` | Critical | Schijf-/bestandssysteemfouten (Event Viewer) |
| `SQLViewError` | High | Fout in de definitie van een SQL-view |
| `GeneralError` | Medium | Elke regel met `^ERROR:` |

---

## Setup

```powershell
# Uitvoeren als Administrator
.\Setup-SASMonitoring.ps1

# Met Zabbix-integratie
.\Setup-SASMonitoring.ps1 -InstallZabbix

# Met e-mailmeldingen
.\Setup-SASMonitoring.ps1 -EmailAlerts -SmtpServer "smtp.contoso.com" -EmailTo "admin@contoso.com" -EmailFrom "sas-monitor@contoso.com"
```

De setup installeert `Monitor-SASBatchErrors.ps1` in `C:\Scripts\` en maakt een geplande taak aan die dagelijks om 08:00 draait.

---

## Gebruik

```powershell
# Laatste 7 dagen doorzoeken, tekstuitvoer
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs"

# Laatste 24 uur doorzoeken
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -DaysToCheck 1

# JSON-uitvoer (voor automatisering)
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -OutputFormat JSON

# Zabbix-uitvoer (geeft het aantal fouten terug)
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -OutputFormat Zabbix

# Event Viewer meenemen
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -IncludeEventLog

# Opslaan in een bestand
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -OutputFile "C:\Temp\report.txt"

# Gezondheidscontrole van WORK/USERWORK met AV-/filterdiagnose
.\Test-SASWorkDirectory.ps1 -Iterations 1000 -EventLogHours 2 -IncludeAVDiagnostics $true -AVLogHours 2
```

**Exitcodes:** `0` = geen critical/high fouten · `1` = ernst high · `2` = critical · `-1` = scriptfout

De AV-diagnose van `Test-SASWorkDirectory.ps1` omvat:

- Momentopname van de Defender-status (realtimebeveiliging, gedragsmonitoring, AV ingeschakeld)
- Controle van Defender-uitsluitingen tegen de gemonitorde WORK-/USERWORK-paden
- Defender Operational-events (blocked/quarantine/denied/CFA/tamper en padgerelateerde meldingen)
- Filterdriver-events uit het System-log (FilterManager/WdFilter)
- Momentopname van actieve minifilters (`fltmc filters`)

---

## Opmerking over tijdelijke schijven (G: / U:)

Als `G:` en `U:` in je omgeving tijdelijke/lokale scratchschijven zijn:

- Ze zijn geschikt voor tijdelijke SAS-data in `WORK`/`USERWORK`, maar niet als persistente opslag.
- Werklast van `G:` naar `U:` verplaatsen is geen structurele failover als beide tijdelijk zijn.
- Periodieke `Access is denied`-fouten tijdens zware I/O-lussen worden vaak veroorzaakt door kortstondige vergrendelingen/filteractiviteit (AV, back-up, indexering, EDR) in plaats van door het verlopen van Kerberos.

`Test-SASWorkDirectory.ps1` logt nu de context van het schijfprofiel en markeert de geconfigureerde tijdelijke stationsletters expliciet, zodat de foutanalyse in omgeleide logs duidelijker is.

Het classificeert ook de SAS-events in het Application-log nauwkeuriger:

- `hc_disk_delete*` + `Access is denied` / returncode `5` worden gemarkeerd als WORK-verwijderfouten waar actie op nodig is.
- `ARM Application data not available` wordt gelogd als informatieve telemetrieruis, tenzij er andere SAS-fouten zijn.
- Dubbele SAS-events met dezelfde tijdstempel/melding worden in de uitvoer ontdubbeld.

---

## Zabbix-integratie

Kopieer `zabbix_sas_monitor.conf` naar `C:\Program Files\Zabbix Agent 2\zabbix_agent2.d\` en herstart daarna de service Zabbix Agent. Of gebruik `Setup-SASMonitoring.ps1 -InstallZabbix` om dit automatisch te doen.

Zabbix-items:

| Key | Omschrijving |
|-----|-------------|
| `sas.batch.errors.critical` | Critical fouten in de laatste 7 dagen |
| `sas.batch.errors.critical.24h` | Critical fouten in de laatste 24 uur |
| `sas.batch.errors.json` | Volledig JSON-rapport |

---

## Foutpatronen toevoegen

Bewerk `Monitor-SASBatchErrors.ps1` en voeg toe aan `$ErrorPatterns`:

```powershell
'MyNewError' = @{
    Pattern     = 'your regex here'
    Severity    = 'Critical'   # Critical / High / Medium
    Description = 'What this error means'
}
```
