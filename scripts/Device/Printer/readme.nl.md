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
| `inputBin` | Standaardlade, met de naam die de driver toont (`Tray 2`, `Manual Feed`) of de Print Schema-naam (`ns0000:Tray2`); spaties en hoofdletters maken niet uit. Een naam die de driver niet kent geeft een waarschuwing met de lades die hij wel aanbiedt |
| `ensure` | `absent` verwijdert de printer (en de poort als geen andere printer die gebruikt). De driver blijft altijd staan |

**Eén apparaat, meerdere wachtrijen (lades)**

Gebruikers kiezen een wachtrij, geen lade. Geef elke wachtrij een eigen regel met **hetzelfde adres**: ze delen de poort `IP_<address>`, en elk krijgt een eigen standaardlade, duplex of kleur:

```json
"printers": [
  { "name": "Office",            "driver": "HP Universal Printing PCL 6", "address": "10.0.5.20", "inputBin": "Tray 1" },
  { "name": "Office - letterhead", "driver": "HP Universal Printing PCL 6", "address": "10.0.5.20", "inputBin": "Tray 2", "duplex": "OneSided" },
  { "name": "Office - envelopes", "driver": "HP Universal Printing PCL 6", "address": "10.0.5.20", "inputBin": "Manual Feed" }
]
```

De lade wordt ingesteld in het standaard printticket van de wachtrij, dus elke gebruiker van die wachtrij begint daarmee; een gebruiker kan in één afdrukvenster nog steeds een andere lade kiezen. Ladenamen verschillen per driver - draai het script één keer met de naam die je verwacht, en een verkeerde naam wordt beantwoord met de lijst die de driver aanbiedt. Universele drivers tonen pas de lades die het apparaat echt heeft als de opties van het apparaat bekend zijn; komt alleen `Auto Select` terug, stel dan eerst de installeerbare opties in bij de printereigenschappen, of gebruik de modelspecifieke driver van de fabrikant.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-ConfigPath` | De JSON: lokaal/UNC-pad of https-URL (bv. een raw-link in dezelfde GitHub-repository) |
| `-Printer` | Alleen deze printers uit de JSON (namen, wildcards toegestaan). Standaard: alle |
| `-GitHubToken` | Token voor een privé-repository (valt terug op `$env:GITHUB_TOKEN`). Wordt alleen naar GitHubs eigen hosts gestuurd, nooit naar een `url`-bron |
| `-Proxy` | Proxy voor elke download, bv. `http://proxy.contoso.local:8080`. Als SYSTEM is er geen gebruikersproxy om over te nemen; gebruikt de referenties van het computeraccount |
| `-WorkingDir` | Map voor downloaden/uitpakken (standaard: `C:\IT\Printers`) |
| `-LogPath` | Map voor `Install-Printer.log` (standaard: `C:\Temp`), aangevuld door elke run die iets kan wijzigen en geroteerd boven 1 MB |
| `-WaitSeconds` | Hoe lang een vers geprovisionde machine krijgt voor de spooler, het netwerk en een andere run van dit script die nog bezig is (standaard: `300`, `0` = niet wachten) |
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

# Eerste opstart of deployment (Custom Script Extension, Run Command, opstarttaak als SYSTEM)
powershell.exe -ExecutionPolicy Bypass -File C:\IT\Install-Printer.ps1 -ConfigPath https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json -Quiet

# Privé-repository
$env:GITHUB_TOKEN = '<fine-grained token, Contents: read>'
.\Install-Printer.ps1 -ConfigPath \\fs01\it$\printers.json -Confirm:$false
```

**Installeren bij de deployment in plaats van in de image**

De image blijft vrij van drivers en printers; elke nieuwe sessiehost krijgt ze bij het uitrollen, zodat een gewijzigde printer een wijziging in de JSON is en geen nieuwe image. Draai het script eenmaal als SYSTEM via de Azure Custom Script Extension in de Bicep van de VM (elke host die de pool erbij krijgt, draait het dan vanzelf), of via Run Command op een host die al bestaat:

```bicep
// Custom Script Extension: runs once, as SYSTEM, when the VM is deployed
resource installPrinters 'Microsoft.Compute/virtualMachines/extensions@2024-07-01' = {
  parent: vm
  name: 'InstallPrinters'
  location: location
  properties: {
    publisher: 'Microsoft.Compute'
    type: 'CustomScriptExtension'
    typeHandlerVersion: '1.10'
    autoUpgradeMinorVersion: true
    settings: {
      fileUris: [
        'https://raw.githubusercontent.com/sjkanon/M365-Scripts/<commit>/scripts/Device/Printer/Install-Printer.ps1'
      ]
    }
    protectedSettings: {
      commandToExecute: 'powershell.exe -ExecutionPolicy Bypass -File Install-Printer.ps1 -ConfigPath https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json -Quiet'
    }
  }
}
```

```powershell
# Run Command: afterwards, on a VM that is already running
az vm run-command invoke -g <resource-group> -n <vm> --command-id RunPowerShellScript `
  --scripts '@Install-Printer.ps1' `
  --parameters 'ConfigPath=https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json'
```

- Laat `fileUris` naar een **commit** wijzen in plaats van `main`, zodat een host die volgende maand wordt uitgerold het script draait dat je getest hebt. De JSON mag op een branch blijven: dat is data, en die wordt volledig gecontroleerd voordat er iets verandert.
- **Geen `-Confirm:$false`** op deze opdrachtregels: een run zonder console vraagt nooit iets, en via `-File` zou de switch aankomen als de tekst `'$false'` en het script stoppen voordat het begint.
- Een VM heeft **één** Custom Script Extension. Gebruikt de deployment die al voor iets anders, gebruik dan hiervoor een managed Run Command (`Microsoft.Compute/virtualMachines/runCommands`).
- Een GitHub-token voor een privé-driverrepository hoort in `protectedSettings` (`-GitHubToken`), nooit in `settings`, dat op de VM leesbaar is.
- Exitcode `1` laat de extensie falen, zodat een deployment met een printer die niet geïnstalleerd kon worden zichtbaar is in de portal; `Install-Printer.log` in `-LogPath` zegt welke.

**Exitcodes**

| Code | Betekenis |
|------|---------|
| `0` | Alles zoals geconfigureerd, of succesvol geïnstalleerd |
| `1` | Fout (een driver of printer die niet geïnstalleerd kon worden, of een ongeldige JSON) |
| `2` | Alleen bij `-CheckOnly`: er is werk te doen |

**Wat een mislukte run voorkomt**
- **De hele JSON wordt gecontroleerd voordat er iets gebeurt**: verplichte velden, onbekende velden (een typfout als `adress` of `loaction` is een fout en wordt niet stilletjes genegeerd), dubbele namen, tekens die Windows in een printernaam weigert, IP-adressen/hostnamen, poortnummers, waarden voor `duplex`/`ensure`, `true`/`false` als tekst geschreven, en twee printers op één poort met verschillende adressen. Alle problemen worden in één keer gemeld en er wordt niets gewijzigd
- **De INF wordt gecontroleerd voordat pnputil hem ziet**: printerklasse, een sectie voor deze architectuur (een pakket dat alleen x86 is, wordt op een x64-server ook zo benoemd), de exacte drivernaam tussen de modellen die hij declareert — met in de foutmelding de namen die hij *wel* declareert, de best passende eerst — en een ondertekende catalogus voor deze architectuur
- **Downloads worden gecontroleerd op wat ze zijn**: een blokkeerpagina van een proxy, een loginportaal of een GitHub-foutmelding die als `.zip` is opgeslagen wordt geweigerd, net als een bestand dat geen zip of cab is; er moet 1 GB vrij zijn voordat er iets wordt opgehaald
- **Eén run tegelijk**: een opstarttaak en een RMM-job die overlappen wachten op elkaar via een machinebrede vergrendeling (tot `-WaitSeconds`) in plaats van dezelfde driver twee keer te installeren. Een run die halverwege is afgebroken wordt herkend en de volgende gaat verder vanaf wat hij achterliet
- **De spooler**: bij de eerste opstart wordt hij gestart en afgewacht tot hij echt antwoordt, niet alleen tot hij Running meldt. Als hij tijdens een installatie stopt of blijft hangen — gebruikelijk vlak nadat een leveranciersdriver is geplaatst — wordt hij herstart en wordt die ene stap nog één keer geprobeerd. Hij wordt bewust *niet* na elke driver herstart, zoals sommige gepubliceerde scripts doen: op een sessiehost in gebruik onderbreekt dat ieders afdrukken
- **De JSON van een URL wordt in de cache bewaard** in `-WorkingDir`. Als de URL niet bereikbaar is, wordt de laatste goede kopie gebruikt met een waarschuwing, zodat een host die opstart terwijl GitHub plat ligt zijn printers houdt
- **Code van de driverleverancier wordt afgeschermd**: het lezen en schrijven van de afdrukstandaarden draait in een job met een time-out, omdat sommige universele drivers daar blijven hangen
- **Eén falende driver houdt de rest niet tegen**: de printers ervan worden overgeslagen en gemeld, alle andere worden geïnstalleerd, en de exitcode is `1`. De gedownloade bestanden van de mislukte driver blijven bewaard voor onderzoek; na een geslaagde installatie worden ze verwijderd (de driver store houdt een eigen kopie)
- **Een gewijzigd adres** wordt doorgevoerd: met standaardpoortnamen verhuist de printer naar de nieuwe poort `IP_<address>` en wordt de oude verwijderd zodra die ongebruikt is; een benoemde poort die alleen deze printer gebruikt wordt ter plekke opnieuw opgebouwd. Een poort die met andere printers wordt gedeeld, wordt niet achter hun rug om gewijzigd
- **Detectie**: een schone run schrijft `ConfigSha256`, `ConfigPath`, `LastSuccess` en `Printers` onder `HKLM:\SOFTWARE\M365-Scripts\InstallPrinter`. Een Intune-detectieregel of RMM-conditie kan de hash vergelijken met die van de JSON, en ziet zo het verschil tussen "geïnstalleerd met de huidige configuratie" en "geïnstalleerd met die van vorige maand"
- **Alles wordt gelogd** in `Install-Printer.log` in `-LogPath`, ook als de run vroeg faalt, zodat een onbewaakte eerste opstart achteraf terug te lezen is

**Opmerkingen**
- **Na een golden image:** draai het als SYSTEM vanuit de Azure Custom Script Extension, een opstarttaak die in het image zit, `SetupComplete.cmd`, Intune of een RMM. Als de Print Spooler nog niet gestart is, start het script hem en wacht het. Downloads die falen op DNS of een time-out worden opnieuw geprobeerd met oplopende wachttijd, beide binnen `-WaitSeconds`. Een 401/403/404 faalt meteen, want wachten verandert daar niets aan. Een spooler die het image heeft *uitgeschakeld* (PrintNightmare-hardening) wordt gemeld, niet stilletjes ingeschakeld
- Printers worden machinebreed aangemaakt, dus op een RDS/AVD-sessiehost ziet elke gebruiker ze. Er komt geen Point and Print per gebruiker aan te pas, dus de beperking `RestrictDriverInstallationToAdministrators` (KB5005652) is niet van toepassing — de driver wordt door een beheerder/SYSTEM geïnstalleerd
- **GitHub:** de releasegegevens en de maplijst gaan via de API (één of twee verzoeken per driver). De bestanden zelf komen zonder token via de downloadlink van de release en `raw.githubusercontent.com` — die tellen geen van beide mee voor de 60 API-verzoeken per uur die GitHub per publiek IP toestaat, dus een pool hosts die achter één NAT opstart raakt niet door de limiet heen. Met een token gaat alles via de API, en dat heeft een privé-repository nodig. Een privé-repository antwoordt zonder token met 404, en de foutmelding zegt dat. Een repository die te groot is voor één tree-lijst, of een bestand dat via Git LFS is opgeslagen, wordt geweigerd in plaats van half gedownload — publiceer de driver dan als release-asset
- Alleen INF-gebaseerde driverpakketten worden ondersteund. Een `.exe`-setup van een leverancier niet — pak het pakket uit (de meeste leveranciers bieden een "driver only"-zip) en zet dat in de repository. De catalogus (`.cat`) moet een geldige handtekening hebben
- `pnputil`-exitcodes `0`, `259` (geen apparaat dat wacht — normaal voor printers), `3010` en `1641` (herstart) gelden als succes. Bij alle andere verwijst de foutmelding naar `C:\Windows\INF\setupapi.dev.log`
- Er bestaat geen `Set-PrinterPort`: een poort die al bestaat voor een ander adres wordt gemeld, niet opnieuw aangemaakt, omdat andere printers hem kunnen gebruiken. Geef de printer een eigen `portName`
- `Set-PrintConfiguration` voert de eigen code van de driver van de leverancier uit en blijft bij sommige universele drivers hangen, dus het draait in een job met een time-out van 2 minuten. Een fout daar is een waarschuwing — de printer is evengoed geïnstalleerd
- De standaardprinter is een instelling per gebruiker en wordt bewust niet ingesteld: als SYSTEM is er geen gebruiker om hem voor in te stellen
- Gestart als 32-bit (Intune Management Extension, sommige RMM-agents) start het script zichzelf opnieuw als 64-bit, omdat `pnputil` niet bestaat onder SysWOW64. Handmatig gestart zonder verhoogde rechten vraagt het daarom
- NinjaOne-scriptvariabelen: `configPath`, `printer`, `githubToken`, `workingDir`, `logPath`, `waitSeconds`, `proxy`, en de selectievakjes `checkOnly`, `whatIf`, `quiet`, `force`, `skipSignatureCheck`
- `-CheckOnly` is de gezondheidscontrole voor een RMM-conditie of een geplande detectietaak: exitcode `2` betekent dat het apparaat niet overeenkomt met de JSON
