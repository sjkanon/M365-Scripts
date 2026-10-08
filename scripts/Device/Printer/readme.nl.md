[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Device](../readme.nl.md) › **Printer**

# Printer

Installeert printerdrivers en TCP/IP-printers aan de hand van één JSON-bestand. De drivers worden gedownload uit een GitHub-repository (release-asset of map), van een willekeurige https-URL of van een share.

Gemaakt om zonder toezicht te draaien nadat een golden image een nieuwe server of sessiehost heeft geprovisioned. Bij de eerste opstart wacht het op de Print Spooler en het netwerk in plaats van te falen, en omdat elke run idempotent is, kan hetzelfde commando ook bij elke opstart draaien.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Install-Printer.ps1`](Install-Printer.ps1) ([docs](#install-printerps1)) | Printerdrivers (gedownload van GitHub) en printers installeren zoals beschreven in een JSON-bestand |

## Bestanden

| Bestand | Omschrijving |
|---------|-------------|
| [`printers.example.json`](printers.example.json) | Voorbeeldconfiguratie met elk veld en elke driverbron |

---

### Install-Printer.ps1

Per run:

1. **Config** — de JSON inlezen (lokaal, UNC of https-URL) en de printers kiezen die met `-Printer` zijn gevraagd. Het hele bestand wordt gevalideerd voordat er iets wordt gedownload
2. **Preflight** — de Print Spooler zo nodig starten en erop wachten, en daarna elke driver (geïnstalleerd? welke versie?) en elke printer (poort, driver, instellingen) vergelijken met de JSON
3. **Download** — alleen voor een driver die ontbreekt of ouder is dan de `version` in de JSON. Een optionele `sha256` wordt gecontroleerd, en `.zip`/`.cab`-bestanden worden uitgepakt
4. **Driver** — de INF zoeken (genoemd in de JSON, of de INF die de drivernaam declareert), de handtekening van de catalogus controleren, en daarna `pnputil /add-driver /install` + `Add-PrinterDriver` uitvoeren
5. **Printer** — de TCP/IP-poort aanmaken, de printer toevoegen of de driver/poort ervan corrigeren, en locatie, opmerking, delen en afdrukstandaarden (duplex, kleur, papierformaat) instellen
6. **Verificatie** — alles teruglezen

Een apparaat dat al in orde is, kost één keer de JSON lezen. Er wordt niets gedownload en er verandert niets.

**De JSON**

```json
{
  "drivers": [
    {
      "name": "HP Universal Printing PCL 6",
      "version": "7.2.0.25780",
      "inf": "hpcu270u.inf",
      "source": { "type": "githubRelease", "repository": "contoso/printer-drivers",
                  "tag": "latest", "asset": "hp-upd-pcl6-x64-*.zip" }
    }
  ],
  "printers": [
    { "name": "Office 1st floor", "driver": "HP Universal Printing PCL 6",
      "address": "10.0.5.20", "location": "1st floor", "duplex": "TwoSidedLongEdge" }
  ]
}
```

| Driverveld | Omschrijving |
|------------|-------------|
| `name` | Exacte drivernaam uit de INF (zoals `Get-PrinterDriver` hem toont). Printers verwijzen ernaar |
| `version` | Optioneel: een geïnstalleerde driver die ouder is, wordt bijgewerkt. Zonder dit veld blijft een geïnstalleerde driver ongemoeid (behalve met `-Force`) |
| `inf` | Optioneel: INF-naam of pad binnen het pakket (wildcards toegestaan). Zonder dit veld neemt het script de INF die `name` noemt, met voorkeur voor de x64/arm64-map |
| `install` | Optioneel `true`: ook installeren als geen van de gekozen printers hem gebruikt |
| `source` | Waar de bestanden vandaan komen — zie hieronder |

| Bron-`type` | Velden |
|-------------|--------|
| `githubRelease` | `repository` (`eigenaar/naam`), `tag` (standaard `latest`), `asset` (naam, wildcards toegestaan — moet precies één asset opleveren) |
| `github` | `repository`, `path` (een `.zip` of een map met de INF), `ref` (branch/tag/commit, standaard: de standaardbranch) |
| `url` | `url` (https) — een `.zip`, `.cab` of los bestand |
| `path` | `path` naar een map of `.zip`; relatieve paden gelden ten opzichte van de map van de JSON |
| *(alle)* | `sha256` — optioneel, gecontroleerd vóór het uitpakken (alleen losse bestanden) |

| Printerveld | Omschrijving |
|-------------|-------------|
| `name`, `driver`, `address` | Verplicht (`address` = IP of hostnaam) |
| `portName` | Standaard `IP_<address>` |
| `portNumber` | RAW-poort, standaard `9100` |
| `lprQueue` | LPR met deze wachtrij gebruiken in plaats van RAW |
| `snmp` | `true` zet de SNMP-status aan. Standaard uit, zodat een printer die niet op SNMP antwoordt niet als Offline wordt getoond |
| `location`, `comment` | Zichtbaar voor gebruikers |
| `shared`, `shareName` | De printer delen (sharenaam standaard gelijk aan `name`) |
| `duplex` | `OneSided`, `TwoSidedLongEdge` of `TwoSidedShortEdge` |
| `color` | `true` / `false` |
| `paperSize` | bv. `A4`, `Letter` |
| `ensure` | `absent` verwijdert de printer (en de poort als geen andere printer die gebruikt). De driver blijft altijd staan |

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-ConfigPath` | De JSON: lokaal/UNC-pad of https-URL (bv. een raw-link in dezelfde GitHub-repository) |
| `-Printer` | Alleen deze printers uit de JSON (namen, wildcards toegestaan). Standaard: alle |
| `-GitHubToken` | Token voor een privé-repository (valt terug op `$env:GITHUB_TOKEN`). Wordt alleen naar GitHubs eigen hosts gestuurd, nooit naar een `url`-bron |
| `-WorkingDir` | Map voor downloaden/uitpakken (standaard: `C:\IT\Printers`) |
| `-LogPath` | Map voor het transcript van een run die iets wijzigt (standaard: `C:\Temp`) |
| `-WaitSeconds` | Hoe lang een vers geprovisionde machine krijgt voor de spooler en het netwerk (standaard: `300`, `0` = niet wachten) |
| `-CheckOnly` | Alleen rapporteren, niets wijzigen. Exitcode `2` betekent dat er werk te doen is |
| `-Quiet` | Niets tonen tenzij er werk is of iets faalt |
| `-Force` | Drivers opnieuw installeren, ook als de versie klopt |
| `-SkipSignatureCheck` | De eigen handtekeningcontrole van de catalogus overslaan (Windows weigert niet-ondertekende pakketdrivers nog steeds) |

**Voorbeelden**

```powershell
# Wat er zou gebeuren - er wordt niets gewijzigd
.\Install-Printer.ps1 -ConfigPath .\printers.json -CheckOnly

# Alleen de Office-printers, als proefrun, JSON rechtstreeks van GitHub
.\Install-Printer.ps1 -ConfigPath https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json -Printer 'Office*' -WhatIf

# Eerste opstart van een server uit een golden image (Custom Script Extension / opstarttaak als SYSTEM)
powershell.exe -ExecutionPolicy Bypass -File C:\IT\Install-Printer.ps1 -ConfigPath https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json -Quiet -Confirm:$false

# Privé-repository
$env:GITHUB_TOKEN = '<fine-grained token, Contents: read>'
.\Install-Printer.ps1 -ConfigPath \\fs01\it$\printers.json -Confirm:$false
```

**Exitcodes**

| Code | Betekenis |
|------|---------|
| `0` | Alles zoals geconfigureerd, of succesvol geïnstalleerd |
| `1` | Fout (een driver of printer die niet geïnstalleerd kon worden, of een ongeldige JSON) |
| `2` | Alleen bij `-CheckOnly`: er is werk te doen |

**Opmerkingen**
- **Na een golden image:** draai het als SYSTEM vanuit de Azure Custom Script Extension, een opstarttaak die in het image zit, `SetupComplete.cmd`, Intune of een RMM. Als de Print Spooler nog niet gestart is, start het script hem en wacht het. Downloads die falen op DNS of een time-out worden opnieuw geprobeerd met oplopende wachttijd, beide binnen `-WaitSeconds`. Een 401/403/404 faalt meteen, want wachten verandert daar niets aan. Een spooler die het image heeft *uitgeschakeld* (PrintNightmare-hardening) wordt gemeld, niet stilletjes ingeschakeld
- Printers worden machinebreed aangemaakt, dus op een RDS/AVD-sessiehost ziet elke gebruiker ze. Er komt geen Point and Print per gebruiker aan te pas, dus de beperking `RestrictDriverInstallationToAdministrators` (KB5005652) is niet van toepassing — de driver wordt door een beheerder/SYSTEM geïnstalleerd
- **GitHub:** een release-asset wordt via de asset-URL van de API gedownload, zodat dezelfde code met en zonder token werkt. Een map wordt één keer via de git tree-API opgevraagd en daarna wordt elk bestand eronder opgehaald. Zonder authenticatie staat GitHub 60 API-verzoeken per uur per IP toe, en een map kost één verzoek per bestand, dus gebruik voor veel hosts achter één NAT een release-`.zip` of een token. Een privé-repository antwoordt zonder token met 404, en de foutmelding zegt dat
- Alleen INF-gebaseerde driverpakketten worden ondersteund. Een `.exe`-setup van een leverancier niet — pak het pakket uit (de meeste leveranciers bieden een "driver only"-zip) en zet dat in de repository. De catalogus (`.cat`) moet een geldige handtekening hebben
- `pnputil`-exitcodes `0`, `259` (geen apparaat dat wacht — normaal voor printers), `3010` en `1641` (herstart) gelden als succes. Bij alle andere verwijst de foutmelding naar `C:\Windows\INF\setupapi.dev.log`
- Er bestaat geen `Set-PrinterPort`: een poort die al bestaat voor een ander adres wordt gemeld, niet opnieuw aangemaakt, omdat andere printers hem kunnen gebruiken. Geef de printer een eigen `portName`
- `Set-PrintConfiguration` voert de eigen code van de driver van de leverancier uit en blijft bij sommige universele drivers hangen, dus het draait in een job met een time-out van 2 minuten. Een fout daar is een waarschuwing — de printer is evengoed geïnstalleerd
- De standaardprinter is een instelling per gebruiker en wordt bewust niet ingesteld: als SYSTEM is er geen gebruiker om hem voor in te stellen
- Gestart als 32-bit (Intune Management Extension, sommige RMM-agents) start het script zichzelf opnieuw als 64-bit, omdat `pnputil` niet bestaat onder SysWOW64. Handmatig gestart zonder verhoogde rechten vraagt het daarom
- NinjaOne-scriptvariabelen: `configPath`, `printer`, `githubToken`, `workingDir`, `logPath`, `waitSeconds`, en de selectievakjes `checkOnly`, `whatIf`, `quiet`, `force`, `skipSignatureCheck`
- `-CheckOnly` is de gezondheidscontrole voor een RMM-conditie of een geplande detectietaak: exitcode `2` betekent dat het apparaat niet overeenkomt met de JSON
