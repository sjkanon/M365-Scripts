[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **Network**

# Network

Diagnosescripts voor netwerk en connectiviteit. Cross-platform waar dat vermeld staat; de rest vereist Windows + Administrator.

---

## Mappen

| Map | Omschrijving |
|--------|-------------|
| [`UniFi/`](UniFi/readme.nl.md) | Netwerkdocumentatierapport voor de UniFi Controller + tooling voor firmware-upgrades |

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Test-Ports.ps1`](Test-Ports.ps1) ([docs](#test-portsps1)) | Controle van TCP-poortconnectiviteit (cross-platform) |
| [`Test-AuthNetworkDiagnostics.ps1`](Test-AuthNetworkDiagnostics.ps1) ([docs](#test-authnetworkdiagnosticsps1)) | Diagnose van authenticatie-/netwerkproblemen (Event Viewer, Kerberos, DNS, shares) |
| [`Test-FileIODiagnostics.ps1`](Test-FileIODiagnostics.ps1) ([docs](#test-fileiodiagnosticsps1)) | Stresstest voor bestands-I/O met live diagnose van fouten |

---

### Test-Ports.ps1

Test TCP-connectiviteit op één of meer poorten naar één of meer hosts. Ondersteunt losse poorten, bereiken (`1294:1494`) en door komma's gescheiden combinaties. Werkt op Windows (PS 5.1+), macOS en Linux (PS 7+).

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Target` | Ja | IP-adres of hostnaam/-namen om te testen |
| `-Ports` | Ja | Poortspecificatie: los (`80`), bereik (`1294:1494`), lijst (`80,443,3389`) of gemengd (`80,443,1294:1494`) |
| `-TimeoutMs` | Nee | Time-out voor de TCP-verbinding in ms (standaard: `500`) |
| `-ShowClosed` | Nee | Gesloten poorten in de uitvoer opnemen (standaard worden alleen open poorten getoond) |

**Voorbeelden**

```powershell
.\Test-Ports.ps1 -Target 192.168.1.1 -Ports 1294:1494
.\Test-Ports.ps1 -Target 10.0.0.1 -Ports 80,443,3389,8080:8090
.\Test-Ports.ps1 -Target server01.contoso.local -Ports 22,3389 -TimeoutMs 1000 -ShowClosed
.\Test-Ports.ps1 -Target 10.0.0.1,10.0.0.2 -Ports 80,443
```

---

### Test-AuthNetworkDiagnostics.ps1

Diagnosticeert authenticatie- en netwerkproblemen op een Windows-server/-endpoint: doorzoekt Event Viewer op aanmeldfouten (4625), Kerberos-fouten (4771/4768), NTLM-fouten (4776) en fouten van SMB/Netlogon/DNS-client; controleert tijdsynchronisatie, de Kerberos-ticketcache, DNS-resolutie, TCP-connectiviteit, toegang tot UNC-shares en de status van de netwerkadapters; doorzoekt optioneel logbestanden op foutpatronen. Schrijft een txt-rapport naar `C:\Temp\`.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-TestHosts` | Hostnamen/IP's om DNS-resolutie en TCP-connectiviteit tegen te testen |
| `-TestPorts` | TCP-poorten om op elke `-TestHosts`-vermelding te testen (standaard: `445, 88, 389, 636`) |
| `-TestShares` | UNC-paden om leestoegang te testen (bijv. `\\server\share`) |
| `-LogDirectory` | Optionele map om te doorzoeken op foutpatronen voor authenticatie/netwerk |
| `-LogDaysBack` | Aantal dagen terug waarvan logbestanden worden meegenomen (standaard: `1`) |
| `-EventLogHours` | Aantal uren terug om in Event Viewer te controleren (standaard: `24`) |
| `-OutputPath` | Andere uitvoermap dan de standaard (`C:\Temp\`) |

**Voorbeelden**

```powershell
# Basisrun — alleen Event Viewer
.\Test-AuthNetworkDiagnostics.ps1

# Connectiviteit + UNC-shares testen
.\Test-AuthNetworkDiagnostics.ps1 -TestHosts "dc01","fileserver" -TestShares "\\fileserver\data"

# Volledige scan inclusief SAS-logbestanden
.\Test-AuthNetworkDiagnostics.ps1 -TestHosts "sasserver" -LogDirectory "E:\SAS\Logs"

# De events van de afgelopen 48 uur controleren
.\Test-AuthNetworkDiagnostics.ps1 -EventLogHours 48
```

Vereist Administrator.

---

### Test-FileIODiagnostics.ps1

Voert een lus van schrijven/toevoegen/lezen/verwijderen uit op een doelpad en classificeert bij elke fout de foutsoort (AUTH/NETWORK/TIMEOUT/DISK/PATH/IO) en legt diagnostische context vast: FileSystemWatcher-events, een diff van de NTFS-rechten (`icacls`) ten opzichte van de baseline bij het opstarten, open bestandshandles via Sysinternals `Handle.exe` (automatisch gedownload naar `C:\Temp\handle\`), een momentopname van nieuwe processen, Kerberos-tickets en vermeldingen in het Security-eventlog. Stopt na 3 fouten.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-TestPath` | Te testen map — lokaal of UNC (bijv. `G:\sas\work`, `\\server\share\folder`) |
| `-Iterations` | Aantal schrijf-/verwijdercycli (standaard: `5000`) |
| `-DelayMs` | Milliseconden tussen iteraties (standaard: `50`) |
| `-StopOnFirstError` | Stoppen na de eerste fout in plaats van door te gaan tot 3 |
| `-HandleExe` | Pad naar een bestaande `Handle.exe` — slaat de automatische download over |
| `-SkipHandleDownload` | Niet proberen `Handle.exe` te downloaden (bijv. een air-gapped server) |
| `-LogPath` | Uitvoermap voor het logbestand (standaard: `C:\Temp\`) |

**Voorbeelden**

```powershell
# Basisrun — downloadt Handle.exe automatisch
.\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work

# Snelle stresstest, stoppen bij de eerste fout
.\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work -Iterations 10000 -DelayMs 0 -StopOnFirstError

# Een bestaande Handle.exe gebruiken, geen download
.\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work -HandleExe C:\Tools\handle.exe

# Air-gapped server — de download overslaan
.\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work -SkipHandleDownload
```
