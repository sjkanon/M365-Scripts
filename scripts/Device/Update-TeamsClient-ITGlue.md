# Teams-update op een werkplek (Update-TeamsClient.ps1)

**Voor IT Glue — servicedeskdocumentatie.** Bedoeld voor alle supportniveaus: level 1 kan hiermee controleren en melden, level 2 kan het uitvoeren, level 3 vindt onderaan de details.

Technische referentie voor beheerders: `Update-TeamsClient.md` in de scriptrepo.

---

## In het kort

| Vraag | Antwoord |
|-------|----------|
| Wat doet het? | Controleert of de nieuwe Teams-app op een werkplek achterloopt, en werkt hem alleen dan bij — inclusief de Teams-vergaderknop in Outlook |
| Wanneer gebruik je het? | Teams start niet of blijft hangen, Teams is verouderd, of de vergaderknop ontbreekt in Outlook |
| Wat merkt de gebruiker? | Teams sluit tijdens de update en moet opnieuw gestart worden; Outlook moet daarna één keer herstarten |
| Gaan er gegevens verloren? | Nee. Chats, bestanden en teams staan in de cloud en blijven staan |
| Hoe lang duurt het? | Meestal een paar minuten. Er wordt ongeveer 275 MB gedownload per werkplek |
| Is een herstart nodig? | Meestal niet. Het script meldt het expliciet als het wél moet |

---

## Wat het script doet

In gewone taal:

1. Het kijkt welke Teams-versie op de werkplek staat.
2. Het vraagt bij Microsoft op welke versie op dit moment de nieuwste is.
3. **Is de werkplek al bij?** Dan gebeurt er niets. Het script stopt zonder iets aan te raken.
4. **Loopt de werkplek achter?** Dan haalt het de installer op, controleert of die echt van Microsoft komt, verwijdert de oude Teams **en elke kopie van de vergader-add-in**, installeert de nieuwe en zet de vergaderknop in Outlook terug.
5. Daarna controleert het of alles er ook echt staat, en meldt of het gelukt is.

### Wat het script **niet** doet

- Het raakt de oude ("classic") Teams niet aan.
- Het verwijdert geen chats, bestanden of instellingen van de gebruiker.
- Het start de computer niet zelf opnieuw op.
- Het doet niets op een werkplek die al bij is — ook niet "voor de zekerheid".

> **Let op — "er gebeurde niets" is meestal goed nieuws.** Als het script niets doet en geen output geeft, betekent dat: deze werkplek heeft al de nieuwste Teams. Dat is geen storing en geen mislukte run. Meld dit niet als fout.

---

## Level 1 — controleren en melden

### Wanneer pak je dit erbij?

| Melding van de gebruiker | Past dit script? |
|---------------------------|------------------|
| "Teams start niet meer op" | Ja — controleer eerst |
| "Teams zegt dat ik moet updaten" | Ja |
| "De knop *Teams-vergadering* is weg in Outlook" | Ja |
| "Ik kan niet inloggen in Teams" | Nee — accountprobleem, geen versieprobleem |
| "Mijn camera/microfoon doet het niet in Teams" | Nee — apparaat/rechtenprobleem |
| "Ik mis een chat of bestand" | Nee — dat lost een herinstallatie niet op |

### Controleren zonder iets te wijzigen

In NinjaOne: draai het script op de werkplek met dit veld bij **Parameters**:

```
-CheckOnly -Quiet
```

Dit verandert **niets** op de werkplek. Je krijgt één van twee uitkomsten:

| Uitkomst | Betekenis | Wat doe je? |
|----------|-----------|-------------|
| Geen output, resultaat "geslaagd" | Teams is al bij | Geen actie. Zoek de oorzaak van het ticket elders |
| Output met `[NEW ] Update available: ... -> ...` | Er is een nieuwere versie | Meld dit in het ticket en zet door naar level 2 |

### De output lezen

Elke regel begint met een label:

| Label | Betekenis |
|-------|-----------|
| `[ OK ]` | Gecontroleerd en in orde |
| `[NEW ]` | Er is nieuws: er is een nieuwere versie beschikbaar |
| `[SKIP]` | Bewust overgeslagen — er was niets te doen |
| `[WARN]` | Let op, maar de run gaat door |
| `[FAIL]` | Mislukt — doorzetten naar level 2 |

Voorbeeld van een werkplek die al bij is:

```
  1. Preflight
  [ OK ] Found MSTeams 26225.1806.5074.1452
  [ OK ] Teams Meeting Add-in is installed

  2. Version check
  [ OK ] Latest published build (x64, ring general): 26225.1806.5074.1452
  [ OK ] Installed build is current (26225.1806.5074.1452)

  [SKIP] Teams is up to date - nothing to do.
```

Voorbeeld van een werkplek die achterloopt:

```
  2. Version check
  [ OK ] Latest published build (x64, ring general): 26240.1000.5100.2000
  [NEW ] Update available: 26225.1806.5074.1452 -> 26240.1000.5100.2000
```

### Wat vertel je de gebruiker?

> "Ik zie dat je Teams-versie verouderd is. Een collega zet de update klaar. Tijdens het bijwerken sluit Teams even af en moet je Outlook één keer opnieuw opstarten. Je chats en bestanden blijven gewoon staan."

### Wanneer escaleer je?

- Bij elke `[FAIL]`-regel.
- Als de controle zelf niet lukt (bijvoorbeeld `Could not determine the latest published build`).
- Als de gebruiker de update niet kan laten uitvoeren omdat hij in vergaderingen zit — plan het in met level 2.

---

## Level 2 — de update uitvoeren

### Vooraf

| Check | Waarom |
|-------|--------|
| Gebruiker zit niet in een vergadering | Teams sluit tijdens de update |
| Gebruiker weet dat Outlook straks herstart moet worden | Anders is de vergaderknop nog even weg |
| Werkplek staat aan en heeft internet | Er wordt ~275 MB gedownload |

### Via NinjaOne (voorkeur)

1. Zoek de werkplek op en start het script **Update-TeamsClient**.
2. Vul bij **Parameters** in:

   ```
   -Quiet -Confirm:$false
   ```

3. Wacht tot de job klaar is en beoordeel het resultaat:

   | Resultaat | Betekenis |
   |-----------|-----------|
   | Geslaagd, geen output | Werkplek was al bij — er is niets gewijzigd |
   | Geslaagd, met output tot en met `[ OK ] Done` | Update uitgevoerd en gecontroleerd |
   | Mislukt | Lees de `[FAIL]`-regel en zoek die op in de foutentabel hieronder |

> Wil je eerst zien wat er zou gebeuren zonder iets te wijzigen? Gebruik `-WhatIf -Confirm:$false`. Het script laat dan de hele update zien maar voert niets uit.

### De drie draaiwijzen

| Doel | Parameters | Wat het doet |
|------|-----------|--------------|
| Alleen controleren | `-CheckOnly -Quiet` | Wijzigt niets. Exitcode 2 als er werk ligt |
| Bijwerken als het nodig is | `-Quiet -Confirm:$false` | Doet niets op een actuele werkplek; werkt bij als er een nieuwere build is |
| Volledige herinstallatie (reparatie) | `-Force -Confirm:$false` | Herinstalleert Teams en de add-in ook als de versie al actueel is. Hiervoor kiezen bij een kapotte Teams, niet als routine |

Voeg `-AvdOptimizations` toe op AVD/VDI-sessiehosts.

### Handmatig op de werkplek

1. Open PowerShell (hoeft niet als administrator — het script vraagt zelf om rechten en gaat verder in een nieuw venster).
2. Draai:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Update-TeamsClient.ps1
   ```

3. Het script vraagt één keer om bevestiging voordat er iets verandert. Typ `y` en Enter.
4. Laat het venster open staan tot je `[ OK ] Done` ziet.

Via het beheermenu van de scriptrepo kan het ook: `menu.ps1`, toets **T**. Dat vraagt eerst of je alleen wilt controleren of echt wilt installeren.

### Na afloop controleren

1. Start Teams — hij moet normaal opstarten en ingelogd zijn.
2. Start Outlook opnieuw en controleer of **Nieuwe Teams-vergadering** in de agenda staat.
3. Ziet het script `A reboot is required`? Plan dan een herstart in met de gebruiker.

---

## Controleren of Outlook de vergaderknop ziet

De add-in machinebreed installeren is één ding; of **Outlook** hem laadt is een tweede. Outlook doet dat per gebruiker. Het script controleert dat nu ook en zet het in de output, zowel in de preflight als in de eindverificatie.

| Regel in de output | Betekenis | Actie |
|--------------------|-----------|-------|
| `[ OK ] Outlook loads the add-in for DOMEIN\gebruiker` | Outlook laadt de add-in bij het opstarten | Geen |
| `[WARN] Outlook has not registered the add-in for any signed-in user yet` | Nog niemand ingelogd, of Outlook is nog niet gestart sinds de installatie | Gebruiker laten in- en uitloggen of Outlook opnieuw starten, daarna opnieuw controleren |
| `[WARN] Outlook has the add-in switched off for ... (LoadBehavior 2)` | **Outlook heeft de add-in zelf uitgeschakeld** — meestal na een crash of een trage start | Outlook → Bestand → Opties → Invoegtoepassingen → COM-invoegtoepassingen → vinkje terugzetten. Blijft het terugvallen, doorzetten naar level 3 |
| `[WARN] Outlook knows the add-in ... but has no LoadBehavior set` | Registratie half aangelegd | Outlook opnieuw starten en opnieuw controleren |
| `[WARN] The add-in is registered for ... but its DLL is gone (...)` | De eigen registratie van die gebruiker wijst naar een bestand dat er niet meer is en overschaduwt de machinebrede installatie | **Vinkje terugzetten helpt niet.** Draai het script met `-RepairOutlookAddIn` (of vink `repairOutlookAddIn` aan): dat ruimt die verouderde registratie op, waarna de machinebrede versie het overneemt bij de volgende Outlook-start |

### De add-in laadt nog steeds niet — waarom?

Als een registratie er wél staat maar Outlook hem niet laadt, zet het script de reden eronder met `why:`. Drie oorzaken laten geen spoor na in `LoadBehavior` zelf:

| `why:`-regel | Wat er aan de hand is | Oplossing |
|--------------|------------------------|-----------|
| `Outlook is x64 but this registration points at the x86 loader` | Outlook kan alleen een DLL van zijn eigen bitness laden | De juiste versie registreren; meestal lost een volledige herinstallatie (`-Force`) dit op |
| `Outlook parked the add-in in its DisabledItems/CrashedAddins list` | Outlook heeft hem zelf uitgezet na een crash of trage start, en houdt hem uit | Outlook → Bestand → Opties → Invoegtoepassingen → Beheren: **Uitgeschakelde items** → inschakelen. Blijft het terugkomen, zet hem dan via beleid op de `DoNotDisableAddinList` |
| `Group policy sets LoadBehavior ... for this add-in` | Een GPO overschrijft de instelling van de gebruiker | Dat beleid aanpassen of verwijderen |

> Staat er **`why: nothing on this machine blocks it`**, dan is er op de werkplek zelf niets mis. Wat dan nog rest: Outlook helemaal afsluiten (ook uit de taakbalk) en opnieuw starten, en zorgen dat de gebruiker minstens één keer in Teams is ingelogd.

---

> Draait het script als System via NinjaOne, dan ziet het alleen de profielen van gebruikers die op dat moment **ingelogd** zijn. Een profiel waar niemand in zit kan het niet uitlezen. Dat is geen fout en laat de job dus ook niet mislukken.

**"Ik zie hem nog niet geladen op alle profielen"** — meestal geen storing. Het script meldt die profielen expliciet:

```
[SKIP] 2 profile(s) are not signed in, so their Outlook registration cannot be read (DOMEIN\jan, DOMEIN\piet)
       - they pick up the machine-wide registration the first time that user starts Outlook
```

Staat er in dezelfde output `[ OK ] Outlook loads the add-in for all users (machine-wide, ...)`, dan is het geregeld: die gebruikers krijgen de add-in zodra ze Outlook voor het eerst starten. Staat die regel er **niet**, dan is er niets om op terug te vallen en moet je de machinebrede installatie eerst rechtzetten (`-Force`).

---

## Waar het script naar Teams zoekt

De preflight inventariseert elke plek waar Teams kan staan, zodat je in één oogopslag ziet wat er op de werkplek leeft:

| Wat | Waar het naar kijkt |
|-----|---------------------|
| Nieuwe Teams per gebruiker | AppX-pakket `MSTeams` voor alle gebruikersprofielen |
| Nieuwe Teams in de image | Het geprovisioneerde pakket — op een sessiehost staat Teams vaak alleen daar, zonder dat een gebruiker hem al heeft |
| Classic Teams (machinebreed) | De oude *Teams Machine-Wide Installer* |
| Classic Teams per gebruiker | `Teams.exe` in het profiel van elke gebruiker |
| Vergader-add-in | Beide uninstall-hives (64-bit en 32-bit) |
| Add-in-kopieën | De machinebrede map plus de map in elk gebruikersprofiel (`%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in`) |
| Outlook-registratie | Per ingelogde gebruiker |

### Classic Teams opruimen

Classic Teams wordt **standaard alleen gemeld, niet verwijderd**. Wil je hem weg hebben, zet dan het vinkje `removeClassicTeams` aan of gebruik de parameter:

```
-RemoveClassicTeams -Quiet -Confirm:$false
```

| Wat het verwijdert | Waarom |
|--------------------|--------|
| De *Teams Machine-Wide Installer* | Zolang die er staat, zet Windows classic Teams in elk nieuw gebruikersprofiel terug |
| De installatiemap in het profiel van elke gebruiker | Daar staat de per-gebruiker-kopie |
| De autostart-regel en de verouderde uninstall-sleutel | Anders probeert Windows iets te starten dat er niet meer is |

Roaming-gegevens in `%APPDATA%\Microsoft\Teams` blijven staan; die doen niets meer zodra de client weg is.

> **Overleg dit met de klant voordat je het breed uitrolt.** Gebruikers die classic Teams nog gebruiken, verliezen hem. Draai het eerst met `-WhatIf` om te zien wat er op een toestel weg zou gaan.

Twee uitkomsten om te kennen:

| Regel in de output | Betekenis | Actie |
|--------------------|-----------|-------|
| `[FAIL] The classic Teams machine-wide installer is still present` | De uninstall is niet gelukt | Doorzetten naar level 3; de job faalt (exitcode 1) |
| `[WARN] ... is no longer registered with Windows Installer (1605)` | De vermelding in Programma's en onderdelen stond er nog, maar Windows kent het product niet meer | Geen actie. Het script ruimt die verouderde vermelding zelf op en gaat door |
| `[WARN] Classic Teams still present for ... - files in use` | Classic Teams draaide nog; de bestanden zaten vast | Gebruiker laten uitloggen; de volgende run maakt het af. Geen fout |

---

## AVD- en VDI-sessiehosts

Op een Azure Virtual Desktop- of VDI-sessiehost heeft Teams twee extra dingen nodig om beeld en geluid goed te laten lopen. Zonder die twee draait de video *in* de sessie: schokkerig beeld, hoge CPU-belasting en klachten over haperend geluid.

| Onderdeel | Waarvoor | Wanneer installeert het script het |
|-----------|----------|-------------------------------------|
| Registervlag `IsWVDEnvironment` | Vertelt Teams dat hij op een sessiehost draait en media moet doorgeven | Als de vlag nog niet op `1` staat |
| Remote Desktop WebRTC Redirector Service | Handelt beeld en geluid af op het lokale apparaat van de gebruiker in plaats van in de sessie | Als hij nog niet geïnstalleerd is |

Je krijgt dit **alleen** als je de optie expliciet aanzet — met het vinkje `avdOptimizations` of de parameter `-AvdOptimizations`:

```
-AvdOptimizations -Quiet -Confirm:$false
```

> Zet dit **niet** aan op gewone laptops en desktops. De vlag vertelt Teams daar dat hij media moet doorgeven aan een redirector die er niet is.

Draait het script op een werkplek die eruitziet als een sessiehost terwijl de optie uitstaat, dan krijg je een tip in de output:

```
  [SKIP] This looks like an AVD session host - -AvdOptimizations sets the media flag and installs WebRTC, -RemoveWebRtcRedirector drops the old stack
```

Staan beide onderdelen al goed, dan gebeurt er niets extra's en wordt er niets gedownload.

### Controleren of de optimalisatie ook echt werkt

Dat de onderdelen geïnstalleerd zijn, betekent nog niet dat gebruikers optimalisatie krijgen — dat hangt af van het lokale apparaat waarmee ze inloggen. Het enige bewijs dat de sessiehost zelf heeft, staat in het gebeurtenissenlogboek: Teams schrijft bij elke verbinding een gebeurtenis weg onder de bron **`Microsoft Teams VDI`**.

Het script leest op een sessiehost automatisch de laatste zeven dagen mee, ook zonder extra opties:

```
  [ OK ] Teams VDI logged 34 event(s) in the last 7 days, last one 2026-09-24 08:12
  [ OK ] Last event: code 24002 - SlimCore deployment not needed - the user is on the new architecture
```

Dat is de goede uitkomst: de gebruikers zitten op de nieuwe techniek. Foutregels worden vertaald, bijvoorbeeld:

```
  [WARN] Teams VDI error 2026-09-24 08:05: ... deployErrc=16002 ...
  [WARN]          code 16002 - no plugin on the endpoint - the remote desktop client has none, or it did not load
```

| Code in de output | Wat het betekent | Actie | Niveau |
|-------------------|------------------|-------|--------|
| `24002`, `24010` | Gebruiker zit op SlimCore — goed | Geen | L1 |
| `16002` | Het lokale apparaat heeft geen (of een te oude) Windows App | Windows App bijwerken op dat apparaat | L2 |
| `16026` | Citrix-beleid blokkeert de virtuele kanalen | Virtual Channel Allow List aanpassen | L3 |
| `16043` | Teams draait als RemoteApp/published app — blijft altijd op WebRTC | Geen, dit is by design | L2 |
| `16389`, `10083`, `1951` | Beleid op het lokale apparaat blokkeert de installatie van SlimCore | Intune/GPO-beleid van dat apparaat bekijken | L3 |
| `24018`, `24043`, `24058` | SlimCore is niet gedownload op het lokale apparaat | Internetverbinding/proxy van dat apparaat | L2 |
| `1722`, `15616`, `24035` | Tijdelijk — lost zichzelf meestal op | Alleen uitzoeken als het blijft terugkomen | L2 |

> Geen enkele regel in de output? Dan heeft niemand in die zeven dagen een sessie opgebouwd, of draait er nog een Teams-versie ouder dan `24123`. Dat is geen fout.

### De oude optimalisatie verwijderen (level 3)

Als élk lokaal apparaat de nieuwe techniek aankan, kan de WebRTC-redirector eraf met `-RemoveWebRtcRedirector`. Dat is **geen** standaardactie en gaat nooit samen met `-AvdOptimizations`; die twee doen het tegenovergestelde en het script weigert de combinatie.

```
-RemoveWebRtcRedirector -Quiet -Confirm:$false
```

> Doe dit pas als de controle hierboven nergens `code 16002` laat zien. Een apparaat dat SlimCore niet aankan én de redirector niet meer vindt, krijgt geen foutmelding — beeld en geluid worden dan gewoon in de sessie verwerkt, precies de belasting die we willen voorkomen.

De registervlag `IsWVDEnvironment` blijft staan: SlimCore heeft die net zo hard nodig.

---

## Melding "Teams optimization for your virtual desktop will soon be unsupported"

Gebruikers op een virtuele werkplek krijgen sinds medio 2026 deze banner in Teams:

> *Teams optimization for your virtual desktop will soon be unsupported. Starting October 1, 2026, we will no longer support the AVD Media optimization technology based on WebRTC.*

**Dit is geen storing.** Beeld en geluid werken gewoon door. Microsoft vervangt de techniek achter de media-optimalisatie en waarschuwt vooraf.

| Datum | Wat er gebeurt |
|-------|----------------|
| 1 oktober 2026 | Einde **ondersteuning** van de WebRTC-optimalisatie. Het blijft werken, maar Microsoft lost er geen problemen meer in op |
| 1 april 2027 | Einde **beschikbaarheid**: WebRTC stopt met werken. Zonder de nieuwe techniek worden gesprekken dan in de sessie zelf verwerkt — schokkerig beeld en hoge CPU-belasting op de sessiehost |

### Waar zit de oplossing? Op het lokale apparaat, niet op de sessiehost

> **SlimCore kun je niet op de sessiehost installeren.** Microsoft laat de plugin in Windows App het pakket (~50 MB) downloaden en registreren op het apparaat waarmee de gebruiker inlogt. Er is dus niets uit te rollen op de server; de enige actie zit op het endpoint.

De opvolger heet **SlimCore**. Die hoeft niemand apart te installeren: hij zit al in de nieuwe Teams-app op de sessiehost, en in **Windows App** op het apparaat waarmee de gebruiker inlogt. Wat bepaalt of de nieuwe techniek gebruikt wordt, is dus de **versie van Windows App op de lokale pc of laptop van de gebruiker**.

| Kant | Wat er moet gebeuren |
|------|----------------------|
| Sessiehost (AVD/VDI) | Niets extra's installeren. Nieuwe Teams actueel houden (dat doet `Update-TeamsClient.ps1`) en `IsWVDEnvironment` op 1 laten staan |
| Lokaal apparaat van de gebruiker | **Windows App bijwerken** naar minimaal `2.0.352.0` (macOS: `11.3.4`, de losse `.pkg`, niet die uit de App Store). De oude *Remote Desktop*-client wordt hiervoor niet meer ondersteund |
| Sessiehost | Teams minimaal `24193.1805.3040.8975` — die eis haal je ruimschoots |

Lukt het op een apparaat niet, dan blokkeert meestal een beleidsinstelling de installatie. Het script controleert deze drie en noemt de Teams-foutcode erbij:

| Beleid | Gevolg | Teams-fout |
|--------|--------|------------|
| `BlockNonAdminUserInstall = 1` | Een gebruiker zonder beheerrechten kan het pakket niet registreren | `16389` |
| `AllowAllTrustedApps = 0` | Sideloading staat uit, het MSIX kan niet installeren | `15615` |
| AppLocker actief | Heeft een uitzondering nodig voor de SlimCoreVdi-pakketten | `10083` |

> Teams heeft **één herstart** nodig om van WebRTC naar SlimCore over te stappen nadat de plugin gevonden is.

### Controleren of het al goed staat

Laat de gebruiker in de virtuele sessie Teams openen → **... → Instellingen → Over Teams**. Onderin staat een van deze meldingen:

| Melding | Betekenis |
|---------|-----------|
| `AVD SlimCore Media Optimized` | Nieuwe techniek actief — klaar, banner verdwijnt |
| `AVD Media Optimized` | Nog op WebRTC — Windows App op het lokale apparaat bijwerken |
| `AVD Media not connected` | Geen optimalisatie actief — Teams afsluiten en opnieuw starten, daarna opnieuw kijken |

### Antwoord voor de gebruiker

> Bedankt voor de melding. Dit is een aankondiging van Microsoft, geen storing: je gesprekken en vergaderingen blijven gewoon werken. De techniek erachter wordt vervangen. Wij zorgen dat de app waarmee je verbinding maakt met de virtuele werkplek wordt bijgewerkt; daarna verdwijnt de melding vanzelf. Je hoeft zelf niets te doen.

### Wat wij intern moeten doen

1. **Inventariseer de Windows App-versie** op de lokale apparaten van alle gebruikers die op de virtuele werkplek inloggen. Dat is het echte werk — de sessiehosts zijn niet het probleem.
2. Rol de nieuwste Windows App uit op die apparaten.
3. Laat `-AvdOptimizations` voorlopig aan staan: Microsoft adviseert de WebRTC-redirector te behouden als terugvaloptie voor apparaten die SlimCore nog niet aankunnen. Vóór april 2027 opnieuw beoordelen.
4. Draai `Update-TeamsClient.ps1 -CheckOnly` op een sessiehost. Zoek SlimCore daar **niet**: die wordt op het lokale apparaat geïnstalleerd, niet op de host. Wat de host wél weet, staat in de gebeurtenissen die het script meeleest — zie [Controleren of de optimalisatie ook echt werkt](#controleren-of-de-optimalisatie-ook-echt-werkt). Zolang daar `code 16002` tussen staat, is er nog minstens één apparaat niet bij.
5. Pas als die code verdwenen is, mag de oude redirector eraf met `-RemoveWebRtcRedirector`.

---

## Level 3 — details

### Parameters

| Parameter | Wanneer gebruiken |
|-----------|-------------------|
| `-WhatIf` | Alles tonen, niets wijzigen |
| `-Quiet` | Alleen output als er nieuws is — voor geplande runs |
| `-CheckOnly` | Alleen controleren en melden (exitcode 2 = update beschikbaar) |
| `-Confirm:$false` | Nooit om bevestiging vragen — verplicht bij onbeheerde runs |
| `-Force` | Herinstalleren terwijl de versie al actueel is (reparatie), of installeren op een werkplek zonder Teams |
| `-AvdOptimizations` | Alleen op AVD/VDI-sessiehosts: zet de mediavlag `IsWVDEnvironment` en installeert de WebRTC-redirector |
| `-RemoveWebRtcRedirector` | Verwijdert de oude WebRTC-optimalisatie. Alleen als élk lokaal apparaat SlimCore aankan. Gaat niet samen met `-AvdOptimizations` |
| `-RemoveClassicTeams` | Verwijdert de oude Teams-client: machine-wide installer plus de installatie in elk gebruikersprofiel |
| `-RepairOutlookAddIn` | Ruimt per-gebruiker-registraties op die naar een verdwenen add-in-DLL wijzen |
| `-WebRtcUrl` | Andere downloadlocatie voor de WebRTC-redirector |
| `-SkipMeetingAddIn` | Alleen de client, de Outlook-add-in met rust laten |
| `-TimeoutSeconds` | Standaard 900. Verhogen op trage werkplekken |
| `-Ring` | Andere update-ring dan `general` |
| `-WorkingDir` | Andere downloadmap dan `C:\IT\AVD\Teams` |
| `-LogPath` | Andere logmap dan `C:\Temp` |

### Exitcodes

| Code | Betekenis | Actie |
|------|-----------|-------|
| `0` | Gelukt, of al bij | Geen |
| `1` | Mislukt | Log lezen, foutentabel raadplegen |
| `2` | Alleen bij `-CheckOnly`: er is een update beschikbaar | Update inplannen |

### Waar staat wat

| Wat | Pad |
|-----|-----|
| Logbestand (alleen bij een run die iets wijzigt) | `C:\Temp\Update-TeamsClient_<datum-tijd>.log` |
| Gedownloade installer | `C:\IT\AVD\Teams\teamsbootstrapper.exe` |
| Geïnstalleerde Outlook-add-in | `C:\Program Files (x86)\Microsoft\TeamsMeetingAddin\<versie>\` |
| Versiebron van Microsoft | `config.teams.microsoft.com` |

### Veiligheden die zijn ingebouwd

- De installer wordt **eerst** gedownload en op Microsoft-handtekening gecontroleerd, en pas daarna wordt de oude Teams verwijderd. Mislukt de download, dan houdt de werkplek zijn werkende Teams.
- Installaties krijgen een timeout en worden afgebroken als ze blijven hangen, zodat een NinjaOne-job de agent niet blokkeert.
- Loopt er al een andere installatie op de werkplek (Windows Update, ander pakket), dan wacht het script en probeert het opnieuw.
- Bij een onverwachte fout stopt de run in plaats van half werk af te maken.

---

## Foutmeldingen

| Melding | Betekenis | Actie | Niveau |
|---------|-----------|-------|--------|
| `Teams is up to date - nothing to do` | Geen fout: werkplek is bij | Geen | L1 |
| `Administrator rights are required` | Zonder beheerrechten gestart | In NinjaOne op **System** draaien, handmatig de UAC-vraag goedkeuren | L1 |
| `Could not enumerate AppX packages` | Zelfde oorzaak: te weinig rechten | Zie hierboven | L1 |
| `Could not determine the latest published build` | `config.teams.microsoft.com` niet bereikbaar (proxy/firewall/DNS) | Netwerk controleren. In noodgeval `-Force` gebruiken om zonder controle te herinstalleren | L2 |
| `Bootstrapper signature is ...` | De download is niet (meer) van Microsoft — meestal een proxy die een foutpagina teruggeeft | Proxy/webfilter controleren. **Niet** omzeilen zonder overleg | L3 |
| `Downloaded file is only N bytes` | Zelfde oorzaak: een foutpagina in plaats van de installer | Zie hierboven | L2 |
| `timed out after 900 seconds and was killed` | Installatie bleef hangen | Werkplek herstarten en opnieuw proberen; anders `-TimeoutSeconds` verhogen | L2 |
| `Another installation is in progress (1618)` | Er loopt al een installatie | Geen actie — het script probeert het zelf opnieuw | L1 |
| `WebRTC Redirector install failed (exit code 1638)` | Er stond al een andere versie van de redirector; die MSI kan niet over zichzelf heen installeren | Hoort niet meer voor te komen: het script verwijdert de oude versie eerst. Komt het toch terug, verwijder de redirector handmatig via Programma's en onderdelen en draai opnieuw | L3 |
| `Use either -AvdOptimizations ... or -RemoveWebRtcRedirector, not both` | Beide opties tegelijk opgegeven; de een installeert wat de ander weghaalt | Kies er één | L2 |
| `Could not remove the WebRTC Redirector` | msiexec weigerde de verwijdering | Log in `C:\Temp` lezen; meestal loopt er een andere installatie | L3 |
| `Teams Meeting Add-in installation failed` | Client staat er, add-in niet | Outlook volledig sluiten en het script opnieuw draaien | L2 |
| `Outlook has the add-in switched off for ... (LoadBehavior 2)` | Outlook heeft de add-in zelf uitgeschakeld, meestal na een crash | Outlook → Bestand → Opties → Invoegtoepassingen → COM-invoegtoepassingen → vinkje terugzetten. Herinstalleren helpt hier niet | L2 |
| `The add-in is registered for ... but its DLL is gone (...)` | Verouderde registratie van die gebruiker overschaduwt de machinebrede installatie | Draaien met `-RepairOutlookAddIn`; vinkje terugzetten in Outlook helpt niet | L2 |
| `Teams installation failed` | De installatie is niet doorgekomen | Log in `C:\Temp` lezen en doorzetten | L3 |
| `A reboot is required` | Windows wil herstarten om af te ronden | Herstart inplannen met de gebruiker | L1 |

---

## Veelgestelde vragen

**Het script gaf helemaal geen output. Is het wel gelopen?**
Ja. Met `-Quiet` zwijgt het script bewust zolang er niets te melden is. Geen output plus resultaat "geslaagd" betekent: werkplek is al bij.

**Kan ik het gewoon nog een keer draaien?**
Ja. Op een werkplek die al bij is doet het niets. Het is niet schadelijk om het vaker te draaien.

**Raakt de gebruiker chats of bestanden kwijt?**
Nee. Die staan in Microsoft 365, niet in de app.

**Moet de gebruiker opnieuw inloggen in Teams?**
Meestal niet — de aanmelding komt van Windows. Gebeurt het toch, dan is dat normaal na een herinstallatie.

**De vergaderknop in Outlook is nog steeds weg.**
Outlook volledig afsluiten (ook in de taakbalk) en opnieuw openen. Outlook laadt invoegtoepassingen **per gebruiker**, dus de knop verschijnt pas bij de eerstvolgende start van Outlook van die gebruiker. Blijft het weg, dan doorzetten naar level 2.

**Wordt de add-in voor alle gebruikers geïnstalleerd?**
Ja. Het script installeert hem machinebreed (`ALLUSERS=1`) in `C:\Program Files (x86)\Microsoft\TeamsMeetingAddin\`, zodat iedereen die op die werkplek inlogt hem heeft — van belang op sessiehosts en gedeelde werkplekken. Op een gewone werkplek installeert en update Teams de add-in daarnaast zelf per gebruiker; daar is `-SkipMeetingAddIn` een verdedigbare keuze. In beide gevallen pikt Outlook hem per gebruiker op bij de volgende start.

**Moet ik dit op alle werkplekken draaien?**
Dat kan als geplande NinjaOne-taak met `-Quiet -Confirm:$false`. Werkplekken die bij zijn worden overgeslagen en verschijnen niet in de activity feed — alleen waar echt iets gebeurde of misging.

---

## Bijlage — het script in NinjaOne zetten (eenmalig, level 3)

Dit hoef je maar één keer per Ninja-omgeving te doen. Daarna kan level 1 en 2 er gewoon mee werken.

> Menupaden verschillen per Ninja-versie en per taalinstelling. De veldnamen hieronder kloppen; het pad ernaartoe kan bij jou net anders heten.

### Stap 1 — het script toevoegen

**Administration → Library → Automation → Add → New Script** (in oudere versies: *Configuration → Scripting*).

Plak de volledige inhoud van `Update-TeamsClient.ps1` in de editor en zet de velden zo:

| Veld | Waarde | Waarom |
|------|--------|--------|
| Name | `Update Teams client + Outlook add-in` | — |
| Description | `Werkt nieuwe Teams bij als Microsoft een nieuwere build publiceert. Doet niets op een werkplek die al bij is.` | Zodat collega's het niet verwarren met een herinstallatie-script |
| Categories | Bijvoorbeeld `Applications` of `Microsoft 365` | Vindbaarheid |
| Language | **PowerShell** (dus 5.1, niet PowerShell 7) | De AppX-commando's zijn daar native; onder 7 lopen ze via een compatibiliteitslaag |
| Operating System | **Windows** | — |
| Architecture | **64-bit** als je die keuze hebt, anders **All** | Het script vangt 32-bit zelf op, maar 64-bit scheelt een herstart van zichzelf |
| Run As | **System** | Nodig voor AppX, MSI en het register |

Opslaan.

### Stap 2 — script variables (optioneel, maar handig)

Met script variables krijgt de collega die het script draait vinkjes in plaats van een parameterregel. Voeg ze toe onder **Script Variables** bij het script:

| Variable name | Type | Label voor de collega |
|---------------|------|------------------------|
| `whatIf` | Checkbox | Alleen tonen wat er zou gebeuren |
| `quiet` | Checkbox | Alleen melden als er nieuws is |
| `checkOnly` | Checkbox | Alleen controleren, niets installeren |
| `force` | Checkbox | Herinstalleren ook als de versie al actueel is |
| `avdOptimizations` | Checkbox | AVD/VDI: mediavlag + WebRTC-redirector afdwingen |
| `repairOutlookAddIn` | Checkbox | Verouderde per-gebruiker-registraties van de add-in opruimen |
| `removeClassicTeams` | Checkbox | Classic Teams verwijderen (machinebreed + per gebruiker) |
| `skipMeetingAddIn` | Checkbox | Outlook-add-in met rust laten |
| `workingDir` | Text | Andere downloadmap |
| `logPath` | Text | Andere logmap |
| `ring` | Text | Andere update-ring |

De naam moet inhoudelijk kloppen (`checkOnly`, niet `check_only`), maar **hoofdletters maken niet uit**: `Quiet`, `quiet` en `QUIET` werken allemaal, want environment-lookups zijn in Windows hoofdletterongevoelig. Ninja zet de variabelen klaar als environment-variabelen, en het script leest ze alleen als dezelfde parameter niet al op de commandoregel staat.

**Controleer het één keer:** draai het script met alleen het vinkje *whatIf* aan. De eerste regel van de output moet zijn:

```
  Mode: -WhatIf - nothing will be changed
```

Staat er `Mode: APPLY`, dan komt de variabele niet door — gebruik dan gewoon het Parameters-veld (stap 3).

### Stap 3 — testen op één werkplek

Zoek een testwerkplek op, **Run Script**, en vul bij **Parameters** in:

```
-WhatIf -Confirm:$false
```

Je ziet de versievergelijking en, als er een update is, alle stappen die uitgevoerd zouden worden. Er verandert niets op de werkplek. Resultaat moet **geslaagd** zijn.

### Stap 4 — inplannen voor de hele omgeving

**Policies → (het beleid van de werkplekken) → Scheduled Automations → Add → Script**.

| Instelling | Waarde |
|------------|--------|
| Script | `Update Teams client + Outlook add-in` |
| Parameters | `-Quiet -Confirm:$false` |
| Run As | System |
| Schedule | Bijvoorbeeld wekelijks buiten kantooruren |

Door `-Quiet` blijft het stil op werkplekken die al bij zijn. Alleen werkplekken waar echt iets gebeurde of misging verschijnen in de activity feed.

Heb je AVD- of VDI-sessiehosts? Maak daarvoor een **tweede** scheduled automation aan, met hetzelfde script maar deze parameters:

```
-AvdOptimizations -Quiet -Confirm:$false
```

Richt die op het beleid van de sessiehosts, niet op gewone werkplekken. Zie [AVD- en VDI-sessiehosts](#avd--en-vdi-sessiehosts).

> `-Confirm:$false` is geen overbodige luxe: het maakt het onmogelijk dat het script op een bevestigingsvraag blijft wachten, ongeacht wat de agent over de sessie meldt.

### Stap 5 — optioneel: detectie in plaats van installatie

Wil je eerst zien op hoeveel werkplekken een update klaarstaat, zonder iets te installeren? Plan dan een tweede automation in met:

```
-CheckOnly -Quiet
```

| Uitkomst | Exitcode | Hoe Ninja het toont |
|----------|----------|---------------------|
| Werkplek is bij | `0` | Geslaagd, geen output |
| Update beschikbaar | `2` | **Mislukt** (elke exitcode ≠ 0 is voor Ninja een fout) — met de versievergelijking in de output |
| Controle lukte niet | `1` | Mislukt, met de reden in de output |

Dat "mislukt" bij code `2` is bedoeld: zo vallen precies de werkplekken op die aandacht nodig hebben. Heeft jullie Ninja-versie *script result conditions*, dan kun je die exitcode gebruiken om er een echte conditie/alert van te maken.

### Aandachtspunten

| Punt | Waar je op let |
|------|----------------|
| Script-timeout in Ninja | Moet ruimer staan dan `-TimeoutSeconds` (standaard 900 s) plus de downloadtijd van ~275 MB. Staat de Ninja-timeout korter, dan breekt Ninja de job af midden in een installatie |
| Draaien als System | Zonder System krijg je `Administrator rights are required` en stopt het script netjes met exitcode 1 |
| Netwerk | `config.teams.microsoft.com` en `statics.teams.cdn.office.net` moeten bereikbaar zijn. Blokkeert de proxy die, dan meldt het script dat en verandert het niets |
| Logbestand | Alleen bij een run die iets wijzigt: `C:\Temp\Update-TeamsClient_<datum-tijd>.log` op de werkplek zelf |
| Bijwerken van het script | Nieuwe versie uit de repo opnieuw in hetzelfde Ninja-script plakken; geplande automations blijven verwijzen naar hetzelfde script |

---

## Tekst voor de gebruiker

**Vooraf:**

> Wij gaan Microsoft Teams op je werkplek bijwerken naar de nieuwste versie. Teams sluit daarbij automatisch af en start daarna weer op. Wil je Outlook na de update één keer opnieuw opstarten? Dan staat de knop voor Teams-vergaderingen weer in je agenda. Je chats, bestanden en teams blijven ongewijzigd.

**Achteraf:**

> De update is uitgevoerd. Start Teams en Outlook opnieuw op. Zie je de knop *Nieuwe Teams-vergadering* niet in je agenda, laat het ons dan even weten.














