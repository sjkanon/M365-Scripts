[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **RDS**

# RDS

Diagnose- en monitoringscripts voor RDP- / RD Web Access-infrastructuur. Voer ze rechtstreeks op de RDS-/RDWeb-server uit voor volledige resultaten — doelen op afstand krijgen alleen controles op connectiviteitsniveau.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Test-RDSDiagnostics.ps1`](Test-RDSDiagnostics.ps1) ([docs](#test-rdsdiagnosticsps1)) | Eenmalige gezondheidscontrole — services, configuratie, certificaten, gebruikersaccount, eventlogs |
| [`Watch-RDSLive.ps1`](Watch-RDSLive.ps1) ([docs](#watch-rdsliveps1)) | Realtime monitor van sessie- en licentie-events |
| [`Get-FSlogix-errors.ps1`](Get-FSlogix-errors.ps1) ([docs](#get-fslogix-errorsps1)) | Diagnose van FSLogix- / Azure Files-profielen op een AVD-sessiehost |
| [`Invoke-FSLogixShrink.ps1`](Invoke-FSLogixShrink.ps1) ([docs](#invoke-fslogixshrinkps1)) | FSLogix-profielschijven op een share verkleinen (Invoke-FslShrinkDisk), of controleren of FSLogix ze zelf comprimeert bij afmelden |

---

### Test-RDSDiagnostics.ps1

Achterhaalt waarom gebruikers zich niet kunnen aanmelden op een RDP- of RD Web Access-server.

**Uitgevoerde controles**

| Onderdeel | Details |
|------|---------|
| RDP-server | Status van `TermService`/`SessionEnv`/`UmRdpService`, RDP in-/uitgeschakeld, NLA, sessielimieten, RD Licensing-modus, groep Remote Desktop Users, firewallregels, actieve sessies (`quser`) |
| RDWeb-server | Bereikbaarheid van poort 443, geldigheid/verloopdatum van het HTTPS-certificaat, IIS + RDWeb-app-pool (alleen lokaal), RD Gateway-service (alleen lokaal) |
| Gebruikersaccount (optioneel, `-Username`) | Ingeschakeld/vergrendeld/verlopen, groepslidmaatschap, beperkingen op aanmeldwerkstations, laatste aanmelding, leeftijd van het wachtwoord |
| Eventlogs (optioneel, `-IncludeEventLogs`) | Security 4625 (mislukte RDP-aanmelding), 4740 (vergrendeling), `TerminalServices-LocalSessionManager` 20/40 (sessiefout/reden van verbreken) |

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-RdpServer` | Hostnaam/IP van de RDP-/Terminal Server. Laat weg om lokale controles uit te voeren |
| `-RdWebServer` | Hostnaam/IP van de RD Web Access-server |
| `-Username` | Een specifiek account controleren op zaken die het aanmelden blokkeren (vereist de module ActiveDirectory of de ADSI-terugval) |
| `-LogPath` | Map voor het logbestand (standaard: `C:\Temp\`) |
| `-IncludeEventLogs` | Analyse van de eventlogs over de laatste `-Hours` uur meenemen |
| `-Hours` | Aantal uren eventloggeschiedenis om te analyseren (standaard: `24`) |

**Voorbeelden**

```powershell
# Volledige controle — RDP + RDWeb + gebruikersaccount
.\Test-RDSDiagnostics.ps1 -RdpServer rdp01.company.local -RdWebServer rdweb.company.local -Username jdoe -IncludeEventLogs

# Lokaal op de RDS-host uitvoeren, eventlogs controleren
.\Test-RDSDiagnostics.ps1 -IncludeEventLogs -Hours 48
```

De resultaten worden naar de console geschreven en naar een logbestand met tijdstempel in `C:\Temp\`.

---

### Watch-RDSLive.ps1

Bevraagt de Windows-eventlogs elke N seconden en streamt nieuwe events naar de console + een logbestand. Voer het rechtstreeks uit op elke RDS-/RDWeb-server.

**Gemonitorde events**

| Bron | Events |
|--------|--------|
| `TerminalServices-LocalSessionManager` | Aanmelding (21), opnieuw verbinden (22/25), afmelding (23), verbinding verbroken (24), aanmelding mislukt (20), reden van verbreken (40) — leesbare redencodes |
| Security | Mislukte RDP-aanmelding (4625, type 10), accountvergrendeling (4740) |
| `TerminalServices-Licensing` | Events voor licentie toegekend/geweigerd/waarschuwing |
| System | Provider `TermServLicensing` (respijtperiode, fouten van de licentieserver) |

Toont per poll een heartbeatregel met het aantal actieve sessies.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-IntervalSeconds` | Pollinterval in seconden (standaard: `20`) |
| `-LogPath` | Map voor het logbestand (standaard: `C:\Temp\`) |
| `-NoLogFile` | Alleen uitvoer naar de console, geen logbestand |

**Voorbeelden**

```powershell
# Uitvoeren op de RDS-server
.\Watch-RDSLive.ps1

# Sneller pollen, geen logbestand
.\Watch-RDSLive.ps1 -IntervalSeconds 10 -NoLogFile
```

> Druk op `Ctrl+C` om te stoppen. Voer uit als Administrator voor toegang tot het Security-log.

---

### Get-FSlogix-errors.ps1

Verzamelt in één run alles wat nodig is om te achterhalen waarom een FSLogix-profiel
niet mount op een AVD-sessiehost — mountfouten, een vergrendelde VHDX, problemen met
SMB/Azure Files of schijffouten. Alleen-lezen: het verzamelt en rapporteert, het repareert niets.

**Wat het verzamelt**

| Onderdeel | Details |
|------|---------|
| Systeem | Hostnaam, OS-build, uptime |
| FSLogix | Geïnstalleerde versie, de volledige configuratie van `Profiles`/`Containers` en de status van de service |
| Containers | Gekoppelde VHD(X)-bestanden, het sessieregister van FSLogix en de profielpaden uit `ProfileList` |
| Opslag | SMB-verbindingen met Azure Files, en of de VHD-share überhaupt bereikbaar is |
| Events | FSLogix-events over de laatste `-Days` dagen, vergeleken met de bekende kritieke foutpatronen, plus schijf-/NTFS-fouten en events van de User Profile Service |
| Restanten | Lokale profielen onder `C:\Users` en de logbestanden van FSLogix |

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-User` | Ook een sectie maken die gefilterd is op één gebruiker — het account waarvan het profiel faalt |
| `-Days` | Aantal dagen eventgeschiedenis om te analyseren (standaard: `7`) |
| `-OutputPath` | Map voor het transcript (standaard: `%SystemDrive%\Temp\FSLogixDiag`) |

**Voorbeelden**

```powershell
# Alles van de afgelopen week
.\Get-FSlogix-errors.ps1

# Eén gebruiker, twee weken terug, rapport op een andere plek
.\Get-FSlogix-errors.ps1 -User jdoe -Days 14 -OutputPath C:\Temp
```

> Voer het uit in een verhoogde sessie **op de sessiehost zelf** — de gegevens over
> containers, SMB en events bestaan alleen daar. De hele run wordt geschreven naar
> `FSLogixDiag_<host>_<timestamp>.log` in de uitvoermap; dat is het bestand dat je
> aan een ticket toevoegt.

---

### Invoke-FSLogixShrink.ps1

Geeft de ruimte terug die FSLogix-profiel- en ODFC-containers vasthouden nadat er data in
is verwijderd: een dynamische VHDX groeit, maar krimpt nooit vanzelf. Het is een schil om
[Invoke-FslShrinkDisk](https://github.com/FSLogix/Invoke-FslShrinkDisk), het eigen
verkleinscript van het FSLogix-team — nog steeds het beste gereedschap hiervoor; de forks en
alternatieven op GitHub doen hetzelfde met minder erachter.

**Wat het doet**

| Stap | Details |
|------|---------|
| Downloaden | Haalt Invoke-FslShrinkDisk op een **vastgezette commit** op naar `C:\Scripts\Invoke-FslShrinkDisk`, deblokkeert het en controleert de SHA-256 van het script. Een gewijzigde upstream-versie draait nooit ongezien; een aangepaste lokale kopie wordt geweigerd |
| Rapport | Elke `.vhd`/`.vhdx` op de share, grootste eerst, met map, grootte en laatste schrijfdatum, en het totaal |
| Verkleinen | Draait Invoke-FslShrinkDisk recursief en vat daarna het CSV-log samen: verkleinde schijven, teruggewonnen GB, en de schijven die niet verwerkt konden worden |
| `-CheckHost` | Op een sessiehost: kan de **ingebouwde compressie van FSLogix bij afmelden** draaien? FSLogix-versie (2210 / 2.9.8361 of later), `VHDCompactDisk`, de service Optimize Drives (`defragsvc` niet Disabled) en dynamische schijven |

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-Path` | De share met de containers, bijv. `\\<storageaccount>.file.core.windows.net\<share>\Profiles`. Wordt recursief doorzocht |
| `-ReportOnly` | Alleen de schijven en hun grootte tonen; niets verkleinen |
| `-IgnoreLessThanGB` | Schijven kleiner dan dit overslaan (standaard: `5`) |
| `-RatioFreeSpace` | Alleen een schijf verkleinen met minstens dit deel vrij erbinnen (standaard: `0.1` = 10%) |
| `-ThrottleLimit` | Aantal schijven tegelijk (standaard: `4`; hooguit twee keer het aantal CPU-cores) |
| `-LogFilePath` | CSV-log (standaard: `C:\Temp\FslShrink_<timestamp>.csv`); de map wordt aangemaakt als die ontbreekt |
| `-ToolPath` | Waar Invoke-FslShrinkDisk staat (standaard: `C:\Scripts\Invoke-FslShrinkDisk`) |
| `-Force` | Invoke-FslShrinkDisk opnieuw downloaden |
| `-CheckHost` | In plaats daarvan de eigen compressie van FSLogix op deze host controleren; geen `-Path` nodig |

**Voorbeelden**

```powershell
# Eerst kijken: elke container op de share, grootste eerst
.\Invoke-FSLogixShrink.ps1 -Path \\sa.file.core.windows.net\profiles\Profiles -ReportOnly

# Alles van 5 GB of meer verkleinen met minstens 10% vrij erbinnen
.\Invoke-FSLogixShrink.ps1 -Path \\sa.file.core.windows.net\profiles\Profiles

# Comprimeert FSLogix de schijven zelf op deze host?
.\Invoke-FSLogixShrink.ps1 -CheckHost
```

**Opmerkingen**

- FSLogix 2210 en later comprimeert een container zelf bij elke afmelding, als de schijf
  groter is dan 1 GB en er minstens 20% te winnen valt ([Microsoft Learn](https://learn.microsoft.com/en-us/fslogix/concepts-vhd-disk-compaction)).
  Draai eerst `-CheckHost`: slaagt die, dan haalt een handmatige verkleining alleen de
  schijven in van gebruikers die zelden afmelden, of die onder de drempel van 20% blijven.
- Een schijf die gekoppeld is — de gebruiker is aangemeld — kan niet verkleind worden; die
  staat in de samenvatting als niet verwerkt en de run eindigt met exitcode 1. Draai het
  buiten kantooruren of met de hosts leeggemaakt.
- Verhoogd uitvoeren (elke schijf wordt gekoppeld), met toegang tot de share: op Azure Files
  via Kerberos of de sleutel van het opslagaccount. Hyper-V is niet nodig.
- Naar een nieuwere Invoke-FslShrinkDisk: zet de nieuwe commit en de SHA-256 van zijn
  `Invoke-FslShrinkDisk.ps1` in `$ToolCommit` / `$ToolHash` bovenin het script, nadat je de
  diff hebt gelezen.
