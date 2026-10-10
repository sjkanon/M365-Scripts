# Teams, nieuwe Outlook en Copilot op AVD — watchdog en herstel

Deze procedure beschrijft wat er automatisch gebeurt als de nieuwe Teams, de nieuwe Outlook of Copilot niet werkt op een AVD-sessiehost, welke meldingen daaruit komen, en wat je doet als een gebruiker belt.

**Voor IT Glue — servicedeskdocumentatie.** Bedoeld voor alle supportniveaus: level 1 neemt het gesprek aan, helpt met de eerste stappen en zet door; level 2 leest de meldingen in het CIPP Teams-kanaal en kijkt op de host; level 3 beheert de scripts.

Technische referentie voor beheerders: de sectie *Watch-M365Apps.ps1* in `scripts/RDS/readme.md` in de scriptrepo.

---

## In het kort

| Vraag | Antwoord |
|-------|----------|
| Wat is het? | Een watchdog (geplande taak) op elke AVD-sessiehost die elke 30 minuten controleert of Teams, de nieuwe Outlook en Copilot werken — bij ons eigen account `itceadmin` én bij elke aangemelde gebruiker |
| Wat doet hij als iets stuk is? | Hij herstelt het zelf. Is het alleen bij één gebruiker, dan alleen in de sessie van die gebruiker. Is het ook bij `itceadmin` of bij meerdere gebruikers, dan eerst de host en daarna de app per gebruiker. Zonder venster, en nooit terwijl de app bij de gebruiker open staat |
| Hoe weten wij het? | Elke bevinding en elk herstel komt als kaart in het **CIPP Teams-kanaal** (hetzelfde kanaal als de CIPP-meldingen), met de host en de naam van de gebruiker. **Alleen level 2 en 3** hebben toegang tot dat kanaal |
| Wat merkt de gebruiker? | Normaal niets. Bij de volgende keer openen werkt de app weer. Probeerde hij in het laatste halfuur zelf een app te openen en lukte dat niet, dan gaat die app na het herstel vanzelf voor hem open. Verder start de watchdog bij een gebruiker nooit een app |
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
   - **per gebruiker** in de eigen sessie van die gebruiker: de app wordt opnieuw geregistreerd. Is het alleen bij één gebruiker, dan blijft het daarbij;
   - de **host** met `Repair-AppxPackageStore.ps1` alleen als het niet bij één gebruiker blijft: bij `itceadmin` (de test voor de hele host), op de host zelf, of dezelfde app bij twee of meer gebruikers. De host gaat dan eerst.
3. Daarna controleert hij opnieuw, en stuurt een kaart naar het CIPP Teams-kanaal met wat er stuk was, bij wie, en of het hersteld is.

> **Waarom gaat het mis?** De meeste gevallen komen van FSLogix: bij het aanmelden zet FSLogix de apps terug in exact de versie die de gebruiker de vorige keer had. Heeft deze host die versie niet, dan mislukt dat met fout `0x80070490` en opent de app niet. De oplossing is de host die versie te geven en de app opnieuw te registreren — precies wat de watchdog doet.

---

## De kaarten in het CIPP Teams-kanaal

> **Alleen level 2 en 3** hebben toegang tot het CIPP Teams-kanaal. Level 1 ziet de kaarten niet en werkt met [Als een gebruiker belt (level 1)](#als-een-gebruiker-belt-level-1).

Elke kaart noemt de **host** (bijvoorbeeld `LEM-AVD-4`), het tijdstip en de betrokken **gebruikers**. Achter een gebruiker staat `(gebruiker)`; `itceadmin` is ons eigen testaccount.

| Kaart | Kleur | Betekenis | Wat doet level 2? |
|-------|-------|-----------|-------------------|
| 🔧 **WORDT HERSTELD** | Blauw | Probleem gevonden, de watchdog herstelt het nu. Per regel: welke app, bij wie. *Bewijs* is een zip met wat er op dat moment op de host stond | Niets. Binnen een paar minuten volgt **HERSTELD** of **NIET HERSTELD** |
| ✅ **HERSTELD** | Groen | Er was iets stuk en het is opgelost. Per regel: welke app, bij wie, en *✅ hersteld bij deze gebruiker* | Niets. Ligt er een ticket van die gebruiker: laat de app opnieuw openen en sluit het ticket als het werkt |
| 🚨 **NIET HERSTELD** | Rood | Herstel geprobeerd, maar (een deel) is nog stuk. *Nog stuk na herstel* en *Wel hersteld* staan apart | Oppakken, ook zonder ticket — de gebruiker heeft er waarschijnlijk al last van. Zie [Level 2](#level-2-de-kaart-zoeken-en-op-de-host-kijken) |
| ⚠️ **PROBLEEM** | Rood | Probleem gevonden, maar herstel staat uit op deze host (`-NoRepair`) | Oppakken, zie [Level 2](#level-2-de-kaart-zoeken-en-op-de-host-kijken) |
| 🆕 **NIEUWE VERSIE** | Blauw | Er staat een nieuwe build van Teams, Outlook of Copilot op de host (oud → nieuw). De watchdog zet elke 6 uur de nieuwste klaar en start hem meteen bij `itceadmin`; gebruikers krijgen hem bij hun volgende aanmelding | Niets. Belt een gebruiker kort daarna over die app: noem de nieuwe versie in het ticket |
| 💥 **CRASH** | Oranje | Teams, Outlook of Copilot is gecrasht of vastgelopen (bij iemand op de host — Windows zegt niet bij wie) | Eén crash: niets. Steeds dezelfde app en foutcode, of meerdere kaarten per dag: onderzoeken |
| ✅ **WEER GOED** | Groen | Een eerder gemeld probleem is vanzelf verdwenen | Niets; sluit een openstaand ticket als de gebruiker het bevestigt |
| ⚠️ **FOUT** | Oranje | De watchdog zelf liep vast | Doorzetten naar **level 3** |
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

## Als een gebruiker belt (level 1)

**Level 1.** Je ziet de meldingen van de watchdog niet en hoeft niets op de host te doen. Je stelt een paar vragen, helpt met de eerste stappen, en zet door naar level 2 als dat niet helpt. Goed om te weten: de watchdog herstelt de meeste gevallen al zelf, vaak voordat de gebruiker belt — opnieuw openen of afmelden en opnieuw aanmelden is dan genoeg.

### Stap 1: vraag uit

| Vraag | Waarom |
|-------|--------|
| Welke app: Teams, nieuwe Outlook of Copilot? | Andere apps vallen hier niet onder |
| Wat gebeurt er precies: opent niet, sluit af, blijft hangen, wit scherm, foutmelding? | Bepaalt welke eerste stap je doet |
| Sinds wanneer, en net na het aanmelden of halverwege de dag? | Net na aanmelden wijst op het klaarzetten van de apps — dat duurt soms een kwartier |
| Werkt het in de browser wel (teams.microsoft.com, outlook.office.com)? | Werkt het daar ook niet, dan is het een account- of dienstprobleem, geen app-probleem |

### Stap 2: eerste hulp

| Wat de gebruiker zegt | Wat je doet |
|-----------------------|-------------|
| "Teams / Outlook opent niet, er gebeurt niets" | 1. App helemaal afsluiten — ook rechtsonder in de taakbalk — en opnieuw openen. 2. Net aangemeld (minder dan een kwartier)? 10–15 minuten wachten en opnieuw proberen. 3. Lukt het niet: laten **afmelden** (Start → profiel → *Afmelden*, niet alleen het venster sluiten) en opnieuw aanmelden. 4. Werkt het dan nog niet: ticket naar **level 2** |
| "Teams / Outlook sluit steeds vanzelf af" | App opnieuw laten starten. Gebeurt het vandaag vaker: ticket naar **level 2** |
| "Outlook / Teams reageert niet meer" | App sluiten (ook in de taakbalk) en opnieuw openen. Bij herhaling: ticket naar **level 2** |
| "Het scherm van Teams / Outlook blijft wit of laadt opnieuw" | App afsluiten en opnieuw openen. Blijft het: ticket naar **level 2** |
| "Copilot staat er niet" | Ticket naar **level 2** — ligt meestal aan de host, niet aan de gebruiker |
| "Ik kan niet inloggen in Teams / Outlook" | Valt hier **niet** onder: account, MFA of licentie. Volg de inlogprocedure |
| "Mijn mail / agenda / chat ontbreekt" | Valt hier **niet** onder: dienstprobleem, geen app-probleem |

### Wat vertel je de gebruiker?

Als een app niet opent:

> "Kun je de app helemaal afsluiten, ook rechtsonder in de taakbalk, en opnieuw openen? Helpt dat niet, meld je dan even af — via Start, je profiel, *Afmelden* — en meld je opnieuw aan. Dan wordt de app opnieuw klaargezet. Je bestanden en chats blijven gewoon staan."

Als je doorzet naar level 2:

> "Ik zet het door naar een collega die op de server meekijkt. Vaak is het dan al automatisch hersteld; je hoort van ons."

Als een app vastliep of afsloot:

> "Start de app gerust opnieuw. Gebeurt het vandaag vaker, laat het ons dan weten, dan zoeken we de oorzaak."

### Wanneer zet je door naar level 2?

- Als afmelden en opnieuw aanmelden niet helpt.
- Als een app vandaag meerdere keren vanzelf afsluit of vastloopt.
- Als meerdere gebruikers hetzelfde melden.
- Bij *Copilot staat er niet*, of een wit scherm dat blijft.

Zet in het ticket: naam van de gebruiker, app, wat de gebruiker ziet, sinds wanneer, wat al geprobeerd is (opnieuw openen, afmelden en aanmelden) en het tijdstip van de melding.

---

## Level 2: de kaart zoeken en op de host kijken

### Stap 1: de kaart van de gebruiker zoeken

Zoek in het CIPP Teams-kanaal op de **naam van de gebruiker**, de laatste uren. De host staat op de kaart. Is er geen kaart, zoek de host dan in de Azure Portal bij de hostpool onder *Sessies*.

| Wat je ziet | Wat je doet |
|-------------|-------------|
| **WORDT HERSTELD** met de naam van de gebruiker, nog geen vervolgkaart | De watchdog is bezig. Wacht een paar minuten op de volgende kaart; een app die de gebruiker in het laatste halfuur zelf probeerde te openen gaat daarna vanzelf open. Na een kwartier nog niets: stap 2 |
| **HERSTELD** met de naam van de gebruiker | Laat de app opnieuw openen; lukt het niet, laten afmelden en opnieuw aanmelden. Werkt het: ticket sluiten |
| **NIET HERSTELD** met de naam van de gebruiker | Laat de gebruiker afmelden en opnieuw aanmelden — bij het aanmelden wordt de app opnieuw klaargezet. Werkt het dan nog niet: [Stap 2](#stap-2-op-de-host); bij één gebruiker herstelt de watchdog de host niet zelf |
| Geen kaart, gebruiker net aangemeld | De watchdog kijkt pas na 10 minuten. Laat hem nu draaien (stap 2) en kijk wat hij vindt |
| Geen kaart, langer aangemeld | Laat de watchdog nu draaien (stap 2). Ziet hij niets, dan ligt het niet aan de registratie: Teams bijwerken of herinstalleren volgens *Teams-update op een werkplek*, of doorzetten naar **level 3** |
| Een of meer **CRASH**-kaarten | Kijk naar het patroon: steeds dezelfde app, module en foutcode? Op meerdere hosts? Dan **level 3** — dan ligt het aan een versie of de image |
| *niet klaargezet op de host* (bijvoorbeeld Copilot) | De host onderzoeken (stap 2, *De host zelf onderzoeken*); lukt herstel niet, **level 3** |
| **FOUT** | Doorzetten naar **level 3** |

Wat je de gebruiker vertelt bij **HERSTELD**:

> "Ik zie dat Teams bij jou niet goed geregistreerd was; dat is intussen automatisch hersteld. Wil je hem opnieuw openen? Lukt het niet, meld je dan even af en opnieuw aan — je bestanden en chats blijven gewoon staan."

### Stap 2: op de host

Alles hieronder in een **verhoogde PowerShell op de sessiehost** (als administrator).

#### De watchdog nu laten draaien

Niet wachten op de volgende halve uur:

```powershell
Start-ScheduledTask -TaskName 'M365 App Watchdog'
```

Na een paar minuten staat het resultaat in het log en, als er iets was, als kaart in het kanaal. Volg het log terwijl hij draait:

```powershell
Get-Content "C:\IT\AppWatchdog\Logs\Watch-M365Apps_$(Get-Date -Format yyyyMMdd).log" -Tail 60 -Wait
```

Stoppen met meekijken: `Ctrl+C`.

#### Het log lezen

Elke regel begint met een label:

| Label | Betekenis |
|-------|-----------|
| `[ OK ]` | Gecontroleerd en in orde |
| `[SKIP]` | Bewust overgeslagen, bijvoorbeeld een gebruiker die net is aangemeld |
| `[WARN]` | Let op, de run gaat door |
| `[FAIL]` | Probleem gevonden — de regel noemt de gebruiker en de app |

De run heeft vaste stappen: *Apps users could not open*, *Crashes and hangs*, *Host*, `itceadmin`, *Users*, en bij een probleem *Collecting evidence*, *Repairing* en *Read back*. Een regel als

```
  [FAIL] AzureAD\jan.peeters - Outlook - Windows could not open it 2x since 09-10 15:00, last at 15:21, error 0x80070490
```

betekent: Jan kon Outlook twee keer niet openen, met de bekende FSLogix-fout.

#### Alleen kijken, niets wijzigen

```powershell
& 'C:\IT\AppWatchdog\Watch-M365Apps.ps1' -NoRepair
```

Laat alles zien wat hij zou vinden, zonder iets te herstellen of te melden als herstel.

#### Wanneer was het laatste herstel?

```powershell
Get-Content C:\IT\AppWatchdog\state.json
```

| Veld | Betekenis |
|------|-----------|
| `LastRepair` | Laatste herstel van de host. `null` = nooit nodig geweest |
| `LastRun` | Laatste run van de watchdog |
| `UserRepairs` | Per gebruiker en app wanneer die laatst opnieuw geregistreerd is (vergeten na een dag) |

#### Het bewijs: wat stond er op de host?

Bij elk nieuw probleem bewaart de watchdog vóór het herstel een zip in `C:\IT\AppWatchdog\Diag` (14 dagen); het pad staat als *Bewijs* op de kaart. Zelf verzamelen, bijvoorbeeld als een gebruiker belt en er geen kaart is:

```powershell
& 'C:\IT\AppWatchdog\Get-M365AppsLog.ps1' -User jan.peeters -App Copilot -Hours 12
```

Wijzigt niets. De zip komt in `C:\Temp`. In `summary.txt` staat per gebruiker en app wat er geregistreerd is en draait, en wat de watchdog ervan vindt; **BLIND SPOT** betekent dat de watchdog het goedkeurt zonder het te testen. Stuur de zip mee als je doorzet naar level 3.

#### De host zelf onderzoeken

```powershell
& 'C:\IT\AppWatchdog\Repair-AppxPackageStore.ps1' -Name teams,outlook,copilot -CheckOnly
```

Wijzigt niets. Laat zien welke versie FSLogix bij aanmelden vroeg, welke de host heeft, en welke apps de laatste dagen faalden en met welke foutcode.

#### Wat je níet doet

- **Geen reset van de app bij een gebruiker** (`Reset-AppxPackage`, of *Herstellen/Opnieuw instellen* in Instellingen): dan moet de gebruiker opnieuw aanmelden in Teams en Outlook en is de lokale cache weg.
- **Geen apps verwijderen** voor alle gebruikers op een host met aangemelde gebruikers.
- **De watchdog niet uitschakelen** om een probleem "stil" te krijgen — meld het aan level 3.

#### Na afloop

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
| Een gebruiker die al is afgemeld | Zonder sessie kan de app niet voor die gebruiker worden hersteld | Bij de volgende aanmelding wordt de app opnieuw klaargezet; lukt dat niet, dan ziet de watchdog het 10 minuten later |
| Een app die bij de gebruiker open staat | De watchdog raakt hem dan niet aan, om hem niet te storen | Draait hij, dan werkt hij kennelijk weer |
| Copilot bij een klant die hem opent en dan faalt | De watchdog telt alleen de nieuwe Copilot-app (`copilotapp.exe`, via Edge Update) en start hem bij `itceadmin`; bij een klant controleert hij alleen of hij geregistreerd is | Bij klachten level 2; `Get-M365AppsLog.ps1 -App Copilot` toont wat er draait |
| De oude Microsoft 365 Copilot-app | Telt niet meer: alleen de nieuwe Copilot-app geldt als Copilot | Ontbreekt de nieuwe app, dan installeert de watchdog hem via Edge Update |
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

## Level 3: beheer van de watchdog

### Installeren, bijwerken, verwijderen

De installatie gebeurt per host, vanaf GitHub, vastgezet op een commit en een SHA-256-hash. Het volledige commando met de actuele commit en hash staat in `scripts/RDS/readme.md` bij *Installing from GitHub*.

```powershell
& $file -NoRepair                                     # eerst kijken
& $file -Install -WebhookUrl '<n8n-webhook>' -WebhookToken '<token>' -Confirm:$false
& $file -TestNotification                             # TEST-kaart in het kanaal
```

De webhook-URL en het token staan **niet** in dit document en niet in de scriptrepo (die is openbaar). Bewaar ze als wachtwoord in IT Glue.

Bijwerken naar een nieuwere versie: het downloadcommando met de nieuwe commit en hash opnieuw draaien, inclusief `-Install` — de taak gebruikt zijn eigen kopie in `C:\IT\AppWatchdog` tot je dat doet. Let op: `-Install` neemt de oude `config.json` niet over. Geef de webhook, het token en de andere instellingen opnieuw mee, of lees ze eerst uit `C:\IT\AppWatchdog\config.json` en geef ze door.

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
| `-UpdateHours` | `6` | Hoe vaak de nieuwste Teams/Outlook klaargezet worden en Edge Update op Copilot controleert. `0` zet bijwerken uit |
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
