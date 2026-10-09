# Teams, nieuwe Outlook en Copilot op AVD — watchdog en herstel

Deze procedure beschrijft wat er automatisch gebeurt als de nieuwe Teams, de nieuwe Outlook of Copilot niet werkt op een AVD-sessiehost, welke meldingen daaruit komen, en wat je doet als een gebruiker belt.

**Voor IT Glue — servicedeskdocumentatie.** Bedoeld voor alle supportniveaus: level 1 neemt het gesprek aan en leest de meldingen, level 2 kijkt op de host, level 3 beheert de scripts.

Technische referentie voor beheerders: de sectie *Watch-M365Apps.ps1* in `scripts/RDS/readme.md` in de scriptrepo.

---

## In het kort

| Vraag | Antwoord |
|-------|----------|
| Wat is het? | Een watchdog (geplande taak) op elke AVD-sessiehost die elke 30 minuten controleert of Teams, de nieuwe Outlook en Copilot werken — bij ons eigen account `itceadmin` én bij elke aangemelde gebruiker |
| Wat doet hij als iets stuk is? | Hij herstelt het zelf: eerst de host, daarna de app in de sessie van de gebruiker. Zonder venster, en nooit terwijl de app bij de gebruiker open staat |
| Hoe weten wij het? | Elke bevinding en elk herstel komt als kaart in het **CIPP Teams-kanaal** (hetzelfde kanaal als de CIPP-meldingen), met de host en de naam van de gebruiker |
| Wat merkt de gebruiker? | Normaal niets. Bij de volgende keer openen werkt de app weer |
| Gaan er gegevens verloren? | Nee. Bij gebruikers wordt een app alleen opnieuw geregistreerd, nooit gereset of verwijderd |
| Wat doet hij níet? | Een gecrashte app opnieuw starten bij een gebruiker, inlogproblemen, licenties, mailbox- of agendaproblemen oplossen |

---

## Welke scripts er meedoen

| Script | Rol | Wie start het? |
|--------|-----|----------------|
| `Watch-M365Apps.ps1` | De watchdog: controleert, herstelt en meldt. Draait als geplande taak **M365 App Watchdog** onder System, elke 30 minuten | Automatisch (geïnstalleerd door level 3) |
| `Repair-AppxPackageStore.ps1` | Herstelt de **host**: zet Teams, Outlook en Copilot klaar voor alle gebruikers in precies de versie die nodig is, en ruimt kapotte registraties op. De watchdog roept het zelf aan, hooguit één keer per 4 uur | De watchdog; handmatig door level 2/3 |
| `Update-SessionHostImage.ps1` | Wekelijks onderhoud van de image en de hosts: zet alle apps op de nieuwste versie, op elke host gelijk | Level 3, wekelijks |
| `Update-TeamsClient.ps1` | Teams bijwerken of herinstalleren op een werkplek | Level 2 — zie de procedure *Teams-update op een werkplek* |
| n8n-flow *ITCE – M365 App Watchdog → Teams* | Zet de melding van de watchdog om in een kaart in het CIPP Teams-kanaal | Automatisch |

### Hoe het samenhangt

1. De watchdog kijkt elke 30 minuten op de host:
   - staan Teams, Outlook en Copilot klaar voor nieuwe gebruikers?
   - werken ze bij `itceadmin`? Draait een app daar niet, dan start de watchdog hem en kijkt of hij blijft draaien;
   - zijn ze geregistreerd en in orde bij **elke andere aangemelde gebruiker** (pas na 10 minuten aanmelden);
   - heeft Windows sinds de vorige keer geweigerd een van die apps te openen, of mislukte een registratie — en bij wie;
   - is een van die apps gecrasht of vastgelopen.
2. Vindt hij iets, dan herstelt hij:
   - eerst de **host** met `Repair-AppxPackageStore.ps1` (meestal de echte oorzaak);
   - daarna **per gebruiker** in de eigen sessie van die gebruiker: de app wordt opnieuw geregistreerd.
3. Daarna controleert hij opnieuw, en stuurt een kaart naar het CIPP Teams-kanaal met wat er stuk was, bij wie, en of het hersteld is.

> **Waarom gaat het mis?** De meeste gevallen komen van FSLogix: bij het aanmelden zet FSLogix de apps terug in exact de versie die de gebruiker de vorige keer had. Heeft deze host die versie niet, dan mislukt dat met fout `0x80070490` en opent de app niet. De oplossing is de host die versie te geven en de app opnieuw te registreren — precies wat de watchdog doet.

---

## De kaarten in het CIPP Teams-kanaal

Elke kaart noemt de **host** (bijvoorbeeld `LEM-AVD-4`), het tijdstip en de betrokken **gebruikers**. Achter een gebruiker staat `(gebruiker)`; `itceadmin` is ons eigen testaccount.

| Kaart | Kleur | Betekenis | Wat doe je? |
|-------|-------|-----------|-------------|
| ✅ **HERSTELD** | Groen | Er was iets stuk en het is opgelost. Per regel: welke app, bij wie, en *✅ hersteld bij deze gebruiker* | Niets. Belt die gebruiker toch, zie [Als een gebruiker belt](#als-een-gebruiker-belt) |
| 🚨 **NIET HERSTELD** | Rood | Herstel geprobeerd, maar (een deel) is nog stuk. *Nog stuk na herstel* en *Wel hersteld* staan apart | Ticket aanmaken, doorzetten naar **level 2** |
| ⚠️ **PROBLEEM** | Rood | Probleem gevonden, maar herstel staat uit op deze host (`-NoRepair`) | Ticket, **level 2** |
| 💥 **CRASH** | Oranje | Teams, Outlook of Copilot is gecrasht of vastgelopen (bij iemand op de host — Windows zegt niet bij wie) | Eén crash: niets. Steeds dezelfde app en foutcode, of meerdere kaarten per dag: ticket, **level 2** |
| ✅ **WEER GOED** | Groen | Een eerder gemeld probleem is vanzelf verdwenen | Niets; sluit een openstaand ticket als de gebruiker het bevestigt |
| ⚠️ **FOUT** | Oranje | De watchdog zelf liep vast | Ticket, **level 3** |
| 🧪 **TEST** | Blauw | Testmelding na installatie | Niets |

Een probleem dat blijft, wordt na 12 uur opnieuw gemeld. Geen kaart betekent: niets gevonden — dat is goed nieuws.

### Wat de woorden op de kaart betekenen

| Op de kaart | Betekenis |
|-------------|-----------|
| *niet klaargezet op de host* | De app staat niet klaar voor nieuwe gebruikers op deze host |
| *niet geregistreerd* | De app ontbreekt bij die gebruiker en kan dus niet geopend worden |
| *kapot* | De app is bij die gebruiker geregistreerd, maar de bestanden ontbreken of de status is niet in orde |
| *kon niet openen* | De gebruiker klikte op de app en Windows weigerde hem te openen |
| *registratie mislukt* | Windows kon de app niet registreren voor die gebruiker, meestal bij het aanmelden |
| *start niet* | Alleen bij `itceadmin`: de app is er, maar start niet of blijft niet open |
| *gecrasht* / *hing en werd gesloten* | De app crashte, of reageerde niet meer en werd afgesloten |

Onder de knop **Acties** staat wat de watchdog precies gedaan heeft, per gebruiker, bijvoorbeeld:

```
AzureAD\an.claes Teams: re-registered in their session - OK
AzureAD\jan.peeters Outlook: running for them now - nothing to do
AzureAD\piet.jans Teams: signed off - the host repair covers the next sign-in
```

---

## Als een gebruiker belt

**Level 1.** Je hoeft niets op de host te doen. Je leest de kaarten en helpt de gebruiker met één van de stappen hieronder.

### Stap 1 — vraag uit

| Vraag | Waarom |
|-------|--------|
| Welke app: Teams, nieuwe Outlook of Copilot? | Andere apps vallen hier niet onder |
| Wat gebeurt er precies: opent niet, sluit af, blijft hangen, foutmelding? | Bepaalt welke kaart je zoekt |
| Sinds wanneer, en net na het aanmelden of halverwege de dag? | Net na aanmelden wijst op FSLogix — de watchdog kijkt pas na 10 minuten |
| Werkt hij in de browser wel (teams.microsoft.com, outlook.office.com)? | Werkt het daar ook niet, dan is het een account- of dienstprobleem, geen app-probleem |

### Stap 2 — zoek de kaart

Zoek in het CIPP Teams-kanaal op de **naam van de gebruiker** van de laatste uren. Weet je niet op welke host de gebruiker zit: level 2 ziet dat in de Azure Portal bij de hostpool onder *Sessies*.

### Stap 3 — handel af

| Wat de gebruiker zegt | Wat je op de kaarten ziet | Wat je doet |
|-----------------------|---------------------------|-------------|
| "Teams / Outlook opent niet, er gebeurt niets" | **HERSTELD** met de naam van de gebruiker | Vraag de app opnieuw te openen. Werkt het niet: laat de gebruiker **afmelden** (Start → profiel → *Afmelden*, niet alleen het venster sluiten) en opnieuw aanmelden |
| | **NIET HERSTELD** met de naam van de gebruiker | Laat de gebruiker afmelden en opnieuw aanmelden. Werkt het daarna nog niet: ticket naar **level 2**, met de kaart erbij |
| | Geen kaart, gebruiker is net aangemeld | Laat de gebruiker 10–15 minuten wachten en het opnieuw proberen; de watchdog kijkt pas na 10 minuten |
| | Geen kaart, langer dan een half uur aangemeld | Laat de gebruiker afmelden en opnieuw aanmelden. Helpt dat niet: ticket naar **level 2** (*watchdog zag niets*) |
| "Teams / Outlook sluit steeds vanzelf af" | **CRASH** van die app | Vraag de app opnieuw te starten. Gebeurt het vaker vandaag, of zie je meerdere CRASH-kaarten voor dezelfde app: ticket naar **level 2** |
| "Outlook / Teams reageert niet meer" | **CRASH** met *hing en werd gesloten* | Als hierboven |
| | Geen kaart | Kort vastlopen dat vanzelf herstelt wordt niet gelogd. App sluiten (ook in de taakbalk) en opnieuw openen |
| "Het scherm van Teams / Outlook blijft wit of laadt opnieuw" | Meestal geen kaart | Dat is vaak het ingebouwde browseronderdeel (WebView2), dat de watchdog niet ziet. App afsluiten en opnieuw openen; blijft het: ticket naar **level 2** |
| "Copilot staat er niet" | **niet klaargezet op de host** voor Copilot | Ticket naar **level 2** — ligt aan de host, niet aan de gebruiker |
| "Ik kan niet inloggen in Teams / Outlook" | — | Valt hier **niet** onder: account, MFA of licentie. Volg de inlogprocedure |
| "Mijn mail / agenda / chat ontbreekt" | — | Valt hier **niet** onder: dienstprobleem, geen app-probleem |

### Wat vertel je de gebruiker?

Bij **HERSTELD**:

> "Ik zie dat Teams bij jou niet goed geregistreerd was; dat is intussen automatisch hersteld. Wil je hem opnieuw openen? Lukt het niet, meld je dan even af en opnieuw aan — je bestanden en chats blijven gewoon staan."

Bij **NIET HERSTELD** of geen kaart:

> "Wil je je even afmelden — via Start, je profiel, *Afmelden* — en opnieuw aanmelden? Dan wordt de app opnieuw klaargezet. Werkt het daarna nog niet, dan kijkt een collega op de server; je hoort van ons."

Bij **CRASH**:

> "Ik zie dat Teams bij je is vastgelopen. Start hem gerust opnieuw. Gebeurt het vaker, laat het ons weten, dan zoeken we de oorzaak."

### Wanneer escaleer je naar level 2?

- Bij elke **NIET HERSTELD**-kaart.
- Als afmelden en opnieuw aanmelden niet helpt.
- Bij meerdere **CRASH**-kaarten voor dezelfde app op een dag, of dezelfde foutcode op meerdere hosts.
- Als meerdere gebruikers op dezelfde host hetzelfde melden.
- Bij elke **FOUT**-kaart (meteen door naar level 3).

Zet in het ticket: gebruiker, host, app, wat de gebruiker ziet, sinds wanneer, en de kaart (schermafbeelding of link).

---

## Level 2 — op de host kijken

Alles hieronder in een **verhoogde PowerShell op de sessiehost** (als administrator).

### De watchdog nu laten draaien

Niet wachten op de volgende halve uur:

```powershell
Start-ScheduledTask -TaskName 'M365 App Watchdog'
```

Na een paar minuten staat het resultaat in het log en, als er iets was, als kaart in het kanaal. Volg het log terwijl hij draait:

```powershell
Get-Content "C:\IT\AppWatchdog\Logs\Watch-M365Apps_$(Get-Date -Format yyyyMMdd).log" -Tail 60 -Wait
```

Stoppen met meekijken: `Ctrl+C`.

### Het log lezen

Elke regel begint met een label:

| Label | Betekenis |
|-------|-----------|
| `[ OK ]` | Gecontroleerd en in orde |
| `[SKIP]` | Bewust overgeslagen, bijvoorbeeld een gebruiker die net is aangemeld |
| `[WARN]` | Let op, de run gaat door |
| `[FAIL]` | Probleem gevonden — de regel noemt de gebruiker en de app |

De run heeft vaste stappen: *Apps users could not open*, *Crashes and hangs*, *Host*, `itceadmin`, *Users*, en bij een probleem *Repairing* en *Read back*. Een regel als

```
  [FAIL] AzureAD\jan.peeters - Outlook - Windows could not open it 2x since 09-10 15:00, last at 15:21, error 0x80070490
```

betekent: Jan kon Outlook twee keer niet openen, met de bekende FSLogix-fout.

### Alleen kijken, niets wijzigen

```powershell
& 'C:\IT\AppWatchdog\Watch-M365Apps.ps1' -NoRepair
```

Laat alles zien wat hij zou vinden, zonder iets te herstellen of te melden als herstel.

### Wanneer was het laatste herstel?

```powershell
Get-Content C:\IT\AppWatchdog\state.json
```

| Veld | Betekenis |
|------|-----------|
| `LastRepair` | Laatste herstel van de host. `null` = nooit nodig geweest |
| `LastRun` | Laatste run van de watchdog |
| `UserRepairs` | Per gebruiker en app wanneer die laatst opnieuw geregistreerd is (vergeten na een dag) |

### De host zelf onderzoeken

```powershell
& 'C:\IT\AppWatchdog\Repair-AppxPackageStore.ps1' -Name teams,outlook,copilot -CheckOnly
```

Wijzigt niets. Laat zien welke versie FSLogix bij aanmelden vroeg, welke de host heeft, en welke apps de laatste dagen faalden en met welke foutcode.

### Wat je níet doet

- **Geen reset van de app bij een gebruiker** (`Reset-AppxPackage`, of *Herstellen/Opnieuw instellen* in Instellingen): dan moet de gebruiker opnieuw aanmelden in Teams en Outlook en is de lokale cache weg.
- **Geen apps verwijderen** voor alle gebruikers op een host met aangemelde gebruikers.
- **De watchdog niet uitschakelen** om een probleem "stil" te krijgen — meld het aan level 3.

### Na afloop

- Laat de gebruiker de app openen en bevestigen dat het werkt.
- Komt het op meerdere hosts terug: doorzetten naar **level 3**, het ligt dan aan de image of FSLogix.

---

## Wat de watchdog níet ziet

| Situatie | Waarom niet | Wat wel |
|----------|-------------|---------|
| Een crash: bij wie? | Windows schrijft bij een crash geen gebruiker in het logboek | Host, app, aantal, foutcode staan wel op de kaart |
| Een wit of herladend Teams- of Outlook-venster | Dat is het browseronderdeel WebView2 dat crasht, niet de app zelf | App opnieuw openen; bij herhaling level 2 |
| *Er is iets misgegaan* in Teams, Outlook dat zichzelf netjes afsluit | Voor Windows is dat geen crash | Gebruiker vragen; bij herhaling level 2 |
| Kort *reageert niet* dat vanzelf weer bijtrekt | Wordt niet gelogd | — |
| Een gebruiker die minder dan 10 minuten is aangemeld | Windows is de apps dan nog aan het klaarzetten | De volgende run, of afmelden en opnieuw aanmelden |
| Een gebruiker die al is afgemeld | Zonder sessie kan de app niet voor die gebruiker worden hersteld | De host is wel hersteld: bij de volgende aanmelding werkt het |
| Een app die bij de gebruiker open staat | De watchdog raakt hem dan niet aan, om hem niet te storen | Draait hij, dan werkt hij kennelijk weer |
| Hosts zonder de watchdog | Hij draait alleen waar hij geïnstalleerd is | Level 3 installeert hem |

---

## Veelvoorkomende foutcodes

| Code | Betekenis | Opmerking |
|------|-----------|-----------|
| `0x80070490` | *Niet gevonden* — FSLogix vroeg een versie van de app die deze host niet heeft | De klassieker. De watchdog herstelt dit zelf |
| `0x80073D02` | De app moet eerst gesloten worden voor een update | Geen storing; de watchdog meldt dit niet |
| `0x80073CFB` / `0x80073D06` | Deze of een nieuwere versie is al geïnstalleerd | Geen storing; de watchdog meldt dit niet |
| `0xC0000005` | Crash: de app las geheugen dat niet van hem was | Bij herhaling met dezelfde module: level 2/3 |
| `0x80070005` | Toegang geweigerd | Rechten of beleid; level 2 |

---

## Level 3 — beheer van de watchdog

### Installeren, bijwerken, verwijderen

De installatie gebeurt per host, vanaf GitHub, vastgezet op een commit en een SHA-256-hash. Het volledige commando met de actuele commit en hash staat in `scripts/RDS/readme.md` bij *Installing from GitHub*.

```powershell
& $file -NoRepair                                     # eerst kijken
& $file -Install -WebhookUrl '<n8n-webhook>' -WebhookToken '<token>' -Confirm:$false
& $file -TestNotification                             # TEST-kaart in het kanaal
```

De webhook-URL en het token staan **niet** in dit document en niet in de scriptrepo (die is openbaar). Bewaar ze als wachtwoord in IT Glue.

Bijwerken naar een nieuwere versie: het downloadcommando met de nieuwe commit en hash opnieuw draaien, inclusief `-Install` — de taak gebruikt zijn eigen kopie in `C:\IT\AppWatchdog` tot je dat doet.

Verwijderen:

```powershell
& 'C:\IT\AppWatchdog\Watch-M365Apps.ps1' -Uninstall -Confirm:$false
```

### Instellingen

Opgeslagen in `C:\IT\AppWatchdog\config.json` (alleen leesbaar voor System en Administrators). Wijzigen: `-Install` opnieuw draaien met **alle** parameters.

| Parameter | Standaard | Wanneer aanpassen |
|-----------|-----------|-------------------|
| `-Account` | `itceadmin` | Meer testaccounts: `-Account itceadmin,itce.user`. Houd elk testaccount op elke host aangemeld (verbroken sessie is genoeg) |
| `-App` | Teams, Outlook, Copilot | Een app die op deze klant niet hoort, weglaten |
| `-CrashThreshold` | `1` (elke crash) | Te veel CRASH-kaarten op een drukke pool: bijvoorbeeld `3`. `0` zet crashmeldingen uit |
| `-RepairCooldownHours` | `4` | Hoe vaak de host en een gebruiker hooguit hersteld worden |
| `-RenotifyHours` | `12` | Na hoeveel uur een blijvend probleem opnieuw gemeld wordt |
| `-NoUserRepair` | uit | Nooit iets in de sessie van een gebruiker doen, alleen melden en de host herstellen |
| `-NoRepair` | uit | Alleen melden, niets herstellen |
| `-SkipLaunchTest` | uit | De apps bij `itceadmin` niet starten, alleen de registratie controleren |

### De n8n-flow

Flow *ITCE – M365 App Watchdog → Teams* in n8n. Hij controleert de header `X-Watchdog-Token` (IF-node *Geldige sleutel?*) en post de kaart naar hetzelfde Teams-kanaal als *ITCE – CIPP Alerts*. Komt er geen kaart:

| In n8n bij de uitvoeringen | Oorzaak |
|----------------------------|---------|
| Geen uitvoering | De melding kwam niet aan: webhook-URL fout, of de host kan n8n niet bereiken |
| Stopt bij *Geldige sleutel?* | Token op de host klopt niet met de IF-node |
| Fout bij *Post to CIPP Teams* | Het Teams-kanaal of de Workflows-koppeling in Teams |

### Wekelijks onderhoud

Draai `Update-SessionHostImage.ps1` wekelijks op alle hosts tegelijk. Dat houdt Teams, Outlook en Copilot op elke host op dezelfde, nieuwste versie — de belangrijkste preventie tegen `0x80070490`. Zie de sectie *Update-SessionHostImage.ps1* in `scripts/RDS/readme.md`.
