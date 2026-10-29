[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Device](../readme.nl.md) › **Printer**

# Printer

Installeert printerdrivers en TCP/IP-printers aan de hand van één JSON-bestand. De drivers worden gedownload uit een GitHub-repository (release-asset of map), van een willekeurige https-URL of van een share.

[`Export-Printer.ps1`](#export-printerps1) werkt andersom: het leest de printers van een bestaande machine en schrijft diezelfde JSON, met de drivers ernaast, in één zip.

Gemaakt om zonder toezicht te draaien nadat een golden image een nieuwe server of sessiehost heeft geprovisioned. Bij de eerste opstart wacht het op de Print Spooler en het netwerk in plaats van te falen, en omdat elke run idempotent is, kan hetzelfde commando ook bij elke opstart draaien.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Export-Printer.ps1`](Export-Printer.ps1) ([docs](#export-printerps1)) | De printers van een machine exporteren - instellingen als JSON-bestand, drivers, volledige printtickets - naar één zip in `C:\Temp`, klaar voor `Install-Printer.ps1` op een andere |
| [`Install-Printer.ps1`](Install-Printer.ps1) ([docs](#install-printerps1)) | Printerdrivers (gedownload van GitHub) en printers installeren zoals beschreven in een JSON-bestand |
| [`New-PrinterGroup.ps1`](New-PrinterGroup.ps1) ([docs](#new-printergroupps1)) | Eén AD- of Entra ID-beveiligingsgroep per printer (`SG-U-PRINTM-<printer>`) vanuit een printer x gebruiker-matrix, leden toegevoegd, en de groeps-SID's als rechten in `printers.json` |

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
| `path` | `path` naar een map of `.zip`; relatieve paden gelden ten opzichte van waar de JSON staat - zijn map op schijf, of zijn URL (met het SAS-token of een andere query erbij). Via https moet het een bestand zijn (`.zip`), geen map |
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
| `permissions` | Wie deze wachtrij mag gebruiken: een lijst van `{ "principal": "CONTOSO\\Finance", "access": "print" }` — zie *Beveiligingsgroepen per wachtrij* hieronder. Zonder dit veld blijven de rechten ongemoeid |
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

**Beveiligingsgroepen per wachtrij**

Een wachtrij kan beperkt worden tot één of meer groepen, bijvoorbeeld een lade met cheques waar alleen Finance op mag afdrukken:

```json
{ "name": "Office - cheques", "driver": "HP Universal Printing PCL 6", "address": "10.0.5.20", "inputBin": "Tray 3",
  "permissions": [
    { "principal": "CONTOSO\\Finance",          "access": "print" },
    { "principal": "CONTOSO\\Office Managers",  "access": "manageDocuments" },
    { "principal": "S-1-12-1-1111-2222-3333-4444", "access": "print" }
  ] }
```

| `access` | Betekent |
|----------|----------|
| `print` | Afdrukken en de eigen afdruktaken beheren — wat Everyone standaard heeft |
| `manageDocuments` | Ook afdruktaken van anderen onderbreken, herstarten en verwijderen |
| `manage` | Ook de printer zelf wijzigen (instellingen, delen, rechten) |

- `principal` is `DOMEIN\Groep`, een lokale gebruiker of groep, of een **SID**. Een Entra ID-groep kan alleen via de SID worden opgegeven (`S-1-12-1-…`, uit Entra ID of uit `whoami /groups` van een lid). `Everyone`, `Authenticated Users`, `Users`, `Interactive`, `Administrators` en `SYSTEM` worden in het Engels begrepen, op Windows in elke taal
- Alleen met `-ApplyPermissions` (NinjaOne `applyPermissions`); zonder wordt het veld genegeerd en blijft de ACL met rust, zodat één JSON dient voor klanten met en zonder printgroepen. Wie in de lijst staat krijgt die toegang; **Everyone en elke andere gebruiker of groep verliest de zijne**, ook het account dat de printer toevallig heeft geïnstalleerd. Administrators, SYSTEM, CREATOR OWNER, app-pakketten en service-identiteiten houden altijd hun toegang — afdrukken vanuit apps zoals de nieuwe Outlook hangt van die vermeldingen af, dus die worden nooit verwijderd. `[]` betekent alleen beheerders. Een `print`-principal die niet te vinden is houdt die wachtrij tegen (fail closed); een `manage`- of `manageDocuments`-principal - zoals het beheerdersaccount waarmee een geëxporteerde wachtrij is aangemaakt - wordt met een waarschuwing weggelaten en de wachtrij gaat er gewoon in
- **Faalt veilig:** een groep die niet gevonden wordt houdt die wachtrij tegen (`-CheckOnly` meldt dat vooraf), en een nieuwe wachtrij waarvan de rechten niet ingesteld kunnen worden wordt weer verwijderd in plaats van open te blijven staan voor iedereen. Een bestaande wachtrij houdt zijn oude rechten en de run faalt. Bij een eerste opstart wordt een onbekende naam tot 2 minuten opnieuw geprobeerd, voor een domeincontroller die nog niet bereikbaar is
- Vergeleken op betekenis, niet op tekst: Windows mag een vermelding anders opslaan dan ze werd opgegeven, en dat telt niet als wijziging. De rechten worden bij elke run gecontroleerd, dus een groep die met de hand in de printereigenschappen is toegevoegd wordt weer weggehaald — de JSON is leidend
- De rechten bepalen wie op de wachtrij mag afdrukken. Of een gebruiker zonder toegang hem nog wel *ziet* in de lijst hangt van de applicatie af en valt buiten dit script

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-ConfigPath` | Waar de printers vandaan komen: `printers.json`, een map waarin die staat, of een exportzip - op schijf, op een share of via https (GitHub raw, een Azure Storage-blob met SAS-token, elke webserver). Een URL zonder bestandsnaam is een map met `printers.json` |
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
| `-ApplyPermissions` | De `permissions` uit de JSON toepassen (printgroepen, beheerders). Zonder deze schakelaar worden ze genegeerd: de printers gaan erin zonder groepen en de ACL van elke wachtrij blijft zoals hij is (Iedereen mag printen) |
| `-ApplySettings` | Het volledige printticket van elke printer toepassen (elke fabrikantinstelling: nieten, afwerking, kwaliteit, ...) uit zijn `printTicket`, of `settings/<printer>.xml` naast de JSON. Eerst het ticket, daarna dubbelzijdig/kleur/papier/lade uit de JSON eroverheen. Eén keer per ticket toegepast (SHA256 onder `HKLM:\SOFTWARE\M365-Scripts\InstallPrinter\PrintTickets`), zodat latere wijzigingen van gebruikers niet elke run worden overschreven |

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
- NinjaOne-scriptvariabelen: `configPath`, `printer`, `githubToken`, `workingDir`, `logPath`, `waitSeconds`, `proxy`, en de selectievakjes `checkOnly`, `whatIf`, `quiet`, `force`, `skipSignatureCheck`, `applyPermissions`, `applySettings`
- `-CheckOnly` is de gezondheidscontrole voor een RMM-conditie of een geplande detectietaak: exitcode `2` betekent dat het apparaat niet overeenkomt met de JSON

---

### Export-Printer.ps1

Het omgekeerde van `Install-Printer.ps1`: leest de printers van deze machine en schrijft één zip, standaard `C:\Temp\PrinterExport_<computer>_<yyyyMMdd-HHmmss>.zip`. Er wordt niets aan de machine gewijzigd.

| In de zip | Inhoud |
|-----------|--------|
| `printers.json` | De printers en hun instellingen in het formaat van Install-Printer: driver, adres, poortnaam/-nummer, LPR-wachtrij, SNMP, locatie, opmerking, delen, dubbelzijdig, kleur, papierformaat, standaardlade (`inputBin`) en rechten. Elke driver verwijst naar zijn eigen zip ernaast (`"type": "path"`, `drivers/<naam>.zip`) |
| `drivers\<drivernaam>.zip` | Het complete driverpakket zoals de driver store het bewaart: INF, ondertekende catalogus en binaries - waar `pnputil` van installeert. Drivers die met Windows meekomen (`ntprint.inf`, `prnms*.inf`) worden niet gekopieerd |
| `settings\<printer>.xml` | Het volledige standaard printticket per printer, met elke fabrikantinstelling die de JSON-velden niet dekken. De `printTicket` van de printer verwijst ernaar; Install-Printer past het toe met `-ApplySettings` |
| `inventory.json` | Alles wat gelezen is, ook van printers die de JSON niet kan beschrijven (WSD, USB, PDF, OneNote, fax, verbindingen met een printserver), met de reden waarom ze niet in `printers.json` staan en de ruwe `PermissionSDDL` |
| `<computer>.printerExport` | Alleen met `-PrintBrm`: de eigen Printer Migration-back-up van Windows |

Op de doelmachine: `.\Install-Printer.ps1 -ConfigPath <de zip> -CheckOnly`, daarna zonder `-CheckOnly` - de zip zoals hij is (lokaal, UNC of een URL), of uitgepakt op een share, in een GitHub-repository of een Azure Storage-container met `-ConfigPath` naar de `printers.json` daar. De relatieve driverpaden worden opgelost ten opzichte van waar de JSON staat. Drivermappen uit een oudere export werken alleen vanaf schijf: via https kan een map niet gedownload worden.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-Printer` | Alleen deze printers (namen, wildcards toegestaan). Standaard: alle |
| `-OutputPath` | Map voor de zip (standaard: `C:\Temp`), wordt aangemaakt als hij ontbreekt |
| `-NoDriver` | Alleen instellingen, geen driverbestanden - `printers.json` noemt dan geen drivers, dus het doel moet ze al hebben |
| `-PrintBrm` | Ook een `PrintBrm.exe -b`-back-up maken en in de zip zetten. Vereist beheerdersrechten |

**Voorbeelden**

```powershell
# Elke printer, met instellingen en drivers, naar C:\Temp
.\Export-Printer.ps1

# Alleen de Office-printers, alleen instellingen
.\Export-Printer.ps1 -Printer 'Office*' -NoDriver

# Ook een PrintBrm-back-up, de zip op een share
.\Export-Printer.ps1 -PrintBrm -OutputPath \\fs01\it$\exports
```

**Opmerkingen**
- Alleen TCP/IP-printers (Standard TCP/IP Port) komen in `printers.json`, want dat is wat Install-Printer aanmaakt. Een printer die het zou weigeren - een naam met `\ , !`, een adres dat geen IP-adres of hostnaam is - wordt gemeld en staat alleen in `inventory.json`
- Een printer op een poortmonitor van de fabrikant (EpsonNet Print, IPP) waarvan de poortnaam het adres bevat - `10.1.50.235:`, `http://10.0.5.20:631/ipp` - wordt geëxporteerd als Standard TCP/IP-poort (RAW 9100) naar dat adres, met zijn driver en instellingen, een waarschuwing en een notitie in zijn `description`: die poortmonitor bestaat alleen waar de software van de fabrikant geïnstalleerd is. Druk op het doel een testpagina af
- Velden worden alleen geschreven als ze afwijken van de standaarden van Install-Printer (poort `IP_<adres>`, poort 9100, SNMP uit), zodat het bestand leest alsof het met de hand geschreven is
- Rechten worden alleen geëxporteerd als ze afwijken van de Windows-standaard (Everyone mag afdrukken). Administrators, SYSTEM, CREATOR OWNER en app-identiteiten blijven weg, omdat Install-Printer die altijd behoudt. Dat geldt ook voor "manage" van het account waarmee de wachtrij is aangemaakt: Windows geeft dat aan de maker, dus het blijft weg als dat account lid is van de lokale groep Administrators (rechtstreeks of via een domeingroep), of een lokaal account is dat niet meer bestaat (het installatieaccount van een image). Een wachtrij met verder alleen Everyone die mag afdrukken krijgt dan helemaal geen `permissions`. Bekende groepen worden met hun Engelse naam geschreven, een Entra ID-gebruiker of -groep met zijn SID (een `AzureAD\naam` is nergens anders op te lossen), een lokaal account van deze machine met een waarschuwing
- Afdrukstandaarden worden gelezen in een job met een time-out van 60 seconden, omdat sommige universele drivers vastlopen in `Get-PrintConfiguration`; een fout daar is een waarschuwing en die printer wordt zonder geëxporteerd
- Beheerdersrechten zijn niet nodig, behalve voor `-PrintBrm`. 32-bits gestart herstart het script zichzelf 64-bits, omdat de driver store vanuit SysWOW64 niet zichtbaar is
- NinjaOne-scriptvariabelen: `printer`, `outputPath`, en de vinkjes `noDriver`, `printBrm`

---

### New-PrinterGroup.ps1

Zet een printer x gebruiker-matrix - zoals gekopieerd uit Excel - om in één beveiligingsgroep per printer, in **Active Directory** of **Entra ID**, met de naam `<Prefix><printernaam>` zonder spaties (`Office - Duplex` → `SG-U-PRINTM-Office-Duplex`), en voegt de aangevinkte gebruikers toe. Met `-ConfigPath` schrijft het elke groep in de vermelding van die printer in `printers.json` als `"permissions": [ { "principal": <de groep>, "access": "print" } ]`, zodat `Install-Printer.ps1` die groep laat afdrukken en Everyone dat recht afneemt; `manage`-vermeldingen blijven. Standaard een voorbeeld; `-Apply` wijzigt de directory en het bestand.

**AD of Entra ID** - een printer-ACL wordt getoetst aan het Windows-aanmeldtoken van de gebruiker, dus de groep moet er een zijn die dat token bevat:

| | Gebruikers melden aan met | Groepen | In `printers.json` |
|---|---|---|---|
| `-Directory AD` | een AD-account: domain-joined of hybrid-joined server of sessiehost, klassieke printserver | aangemaakt in `-OU` - kies er een die Entra Connect synchroniseert, dan staan ze ook in Entra ID | `DOMEIN\SG-U-PRINTM-...` |
| `-Directory Entra` | een Entra ID-account op een Entra-joined apparaat | cloudgroepen | hun SID `S-1-12-1-...`, berekend uit de object-id |
| `-Directory Entra -AdWriteback` | een AD-account (hybride), maar de groepen worden in Entra ID beheerd | cloudgroepen die Entra Cloud Sync naar AD schrijft ("Microsoft Entra ID to AD sync") | de AD-kopie, `DOMEIN\<zijn sAMAccountName>` |

In een hybride omgeving is dat `-Directory AD`: een cloudgroep komt nooit in het token van een AD-aanmelding, en de wachtrij zou iedereen weigeren. `-Directory Entra` waarschuwt als de gebruikers uit de matrix vanuit AD gesynchroniseerd worden. In beide gevallen krijgt een nieuw lid toegang bij de volgende aanmelding.

De matrix: eerste kolom de printernamen, kopregel de gebruikers (weergavenaam, UPN, of `DOMEIN\account` voor AD), elke niet-lege cel (`x`) betekent toegang. Gescheiden door tab, puntkomma of komma; een kopregel zonder de lege cel linksboven, zoals Excel kopieert, is goed. `Naam (AFDELING)` matcht ook de weergavenaam `Naam`. Een naam die geen of meerdere gebruikers oplevert wordt gemeld en overgeslagen.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-MatrixPath` | Het matrixbestand |
| `-Directory` | `AD` of `Entra` (standaard) |
| `-OU` | AD: distinguished name van de OU voor nieuwe groepen. Verplicht zodra er een groep aangemaakt moet worden |
| `-AdWriteback` | Entra: de groepen komen via Cloud Sync in AD; zet hun AD-kopieën in `printers.json` en controleer of de leden aangekomen zijn. Heeft Graph (PowerShell 7) en de ActiveDirectory-module op dezelfde domain-joined machine nodig |
| `-Prefix` | Voorvoegsel van de groepsnaam (standaard: `SG-U-PRINTM-`). AD beperkt een groepsnaam tot 64 tekens; langere stoppen de run voor er iets verandert |
| `-ConfigPath` | `printers.json` om de rechten in te schrijven. `manage`/`manageDocuments`-vermeldingen blijven, `Everyone` en andere `print`-vermeldingen worden vervangen; eerst wordt een `.bak` geschreven. Een groep wordt pas geschreven als hij bestaat - Install-Printer weigert een wachtrij aan te maken waarvan de groep niet gevonden wordt. Install-Printer past ze alleen toe met `-ApplyPermissions` |
| `-RemoveExtraMembers` | Ook groepsleden verwijderen die de matrix niet aanvinkt (standaard: gemeld, met rust gelaten) |
| `-TenantId` | Entra: doeltenant; `-TenantId` is standaard de GDAP-klanttenant (`$global:cid` / `$env:M365_CUSTOMER_TENANTID`) |
| `-Apply` | Groepen aanmaken, leden wijzigen, de JSON schrijven. Zonder: alleen rapporteren |

**Voorbeelden**

```powershell
# Hybride / on-premises: wat er in AD aangemaakt zou worden
.\New-PrinterGroup.ps1 -MatrixPath .\printer-matrix.tsv -Directory AD

# De AD-groepen aanmaken en er het printrecht van de wachtrijen van maken
.\New-PrinterGroup.ps1 -MatrixPath .\printer-matrix.tsv -Directory AD -OU 'OU=Printers,OU=Groups,DC=contoso,DC=local' -ConfigPath .\printers.json -Apply

# Entra-joined apparaten: cloudgroepen
.\New-PrinterGroup.ps1 -MatrixPath .\printer-matrix.tsv -TenantId contoso.com -ConfigPath .\printers.json -Apply
```

**Opmerkingen**
- **Cloud Sync voor `-AdWriteback`**: zet de groepen in de scope van de configuratie (geselecteerde groepen, of alle groepen met een filter zoals `displayName` REGEX MATCH `^SG-U-PRINTM-`); zet **Provision membership to on-premises users** aan - dat staat standaard uit, en zonder blijven de AD-kopieën leeg; en voeg een `sAMAccountName`-mapping toe (bv. `[displayName]`), anders krijgt de AD-kopie een willekeurige naam - dat werkt, maar leest slecht in een ACL. De AD-kopie wordt gevonden op `displayName` en zijn `msDS-ExternalDirectoryObjectId`; een groep die nog niet teruggeschreven is wordt gemeld en blijft uit de JSON - draai opnieuw na de volgende provisioning-cyclus
- **Waar**: de run noemt eerst het doel - het AD-domein en de domeincontroller die hij gebruikt (de PDC-emulator, zodat een groep die net is aangemaakt er is wanneer leden worden toegevoegd), of de Entra-tenant - en `-Apply` vraagt bevestiging voor er iets wordt aangemaakt (`-Confirm:$false` slaat dat over). Zonder `-TenantId` gebruikt Entra de GDAP-klant, anders de tenant waarin je aanmeldt, of die waarin een bestaande Graph-sessie in dat venster al zit
- `-Directory AD` heeft de ActiveDirectory-module (RSAT) nodig op een domain-joined machine en draait in Windows PowerShell 5.1; `-Directory Entra` heeft PowerShell 7 en Microsoft Graph nodig (`Group.ReadWrite.All`, `User.Read.All`)
- `-Directory Entra` gebruikt `Connect-M365.ps1` niet: aanmelden (GDAP of een gewoon beheerdersaccount) gebeurt vooraf. Een bestaande Graph-sessie voor de doeltenant wordt hergebruikt; anders meldt het script zelf interactief aan met `Connect-MgGraph` (bijvoorbeeld als `admin@tenant.onmicrosoft.com`) en aan het eind weer af
- Hou de matrix buiten deze repository: er staan echte mensen in. De door git genegeerde map `local/` is een prima plek
