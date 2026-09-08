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
4. **Loopt de werkplek achter?** Dan haalt het de installer op, controleert of die echt van Microsoft komt, verwijdert de oude Teams, installeert de nieuwe en zet de vergaderknop in Outlook terug.
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

## Level 3 — details

### Parameters

| Parameter | Wanneer gebruiken |
|-----------|-------------------|
| `-WhatIf` | Alles tonen, niets wijzigen |
| `-Quiet` | Alleen output als er nieuws is — voor geplande runs |
| `-CheckOnly` | Alleen controleren en melden (exitcode 2 = update beschikbaar) |
| `-Confirm:$false` | Nooit om bevestiging vragen — verplicht bij onbeheerde runs |
| `-Force` | Herinstalleren terwijl de versie al actueel is (reparatie), of installeren op een werkplek zonder Teams |
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
| `Teams Meeting Add-in installation failed` | Client staat er, add-in niet | Outlook volledig sluiten en het script opnieuw draaien | L2 |
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
Outlook volledig afsluiten (ook in de taakbalk) en opnieuw openen. Blijft het weg, dan doorzetten naar level 2.

**Moet ik dit op alle werkplekken draaien?**
Dat kan als geplande NinjaOne-taak met `-Quiet -Confirm:$false`. Werkplekken die bij zijn worden overgeslagen en verschijnen niet in de activity feed — alleen waar echt iets gebeurde of misging.

---

## Tekst voor de gebruiker

**Vooraf:**

> Wij gaan Microsoft Teams op je werkplek bijwerken naar de nieuwste versie. Teams sluit daarbij automatisch af en start daarna weer op. Wil je Outlook na de update één keer opnieuw opstarten? Dan staat de knop voor Teams-vergaderingen weer in je agenda. Je chats, bestanden en teams blijven ongewijzigd.

**Achteraf:**

> De update is uitgevoerd. Start Teams en Outlook opnieuw op. Zie je de knop *Nieuwe Teams-vergadering* niet in je agenda, laat het ons dan even weten.
