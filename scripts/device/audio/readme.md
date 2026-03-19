# Detect-AudioDevices.ps1

> **BraveHub Internal Script**
> Auteur: Sjoerd Kanon | Datum: 19/03/2026

---

## Achtergrond en probleemstelling

Best Next Contact gebruikt **Adversus**, een cloud dialer die via de browser werkt. Medewerkers hoorden achtergrondgeluid tijdens gesprekken terwijl ze een headset (EPOS IMPACT DW / Yealink) ingeplugd hadden.

De oorzaak: Windows selecteert soms de **interne laptopmicrofoon** als audiobron in plaats van de headset, ook als de headset correct ingeplugd is. Dit gedrag treedt op in browsergebaseerde applicaties omdat de browser zelf de audiobron kiest op basis van Windows-instellingen, niet op basis van wat de gebruiker verwacht.

Microsoft Teams had dit probleem niet omdat Teams een eigen audiobeheer heeft dat de communicatieapparaten correct prioriteert. Adversus als webapplicatie vertrouwt volledig op de browserinstellingen.

### Wat is al geprobeerd

| Actie | Resultaat |
|---|---|
| Noise suppression in Adversus inschakelen | Onvoldoende — achtergrondgeluid bleef |
| EPOS headset software installeren + updaten | Hielp niet |
| Registry `ConsentStore\microphone = Deny` | **Te breed** — blokkeerde ook de headset, medewerkers konden helemaal niet meer bellen |

### Waarom het eerste script mislukte

Het script dat eerder werd uitgevoerd blokkeerde microfoonpermissies via het Windows Privacy / Consent Store mechanisme:

```
HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone
Value = "Deny"
```

Dit blokkeert **alle** microfooninput voor **alle** applicaties — inclusief de headset. Na uitrol kon Yllana helemaal niet meer bellen. Het script moest teruggedraaid worden.

### De correcte aanpak

De interne microfoon moet **device-specifiek** disabled worden via PnP, zodat de headset microfoon intact blijft. Om dit correct te doen moeten we eerst weten welke exacte device namen en InstanceIds de interne microfoons hebben op alle laptops — want dit kan per laptop model verschillen.

Daarom wordt eerst een **detect script** uitgerold via NinjaOne om een volledige audio inventaris op te halen.

---

## Aanpak in twee fases

```
Fase 1: Detect (dit script)
  └── Uitrollen op alle laptops via NinjaOne
  └── Output verzamelen en analyseren
  └── Bepalen welke naampatronen intern vs headset zijn

Fase 2: Disable (nog te ontwikkelen)
  └── Device-specifiek de interne microfoon disablen via Disable-PnpDevice
  └── Headset blijft actief
  └── Rollback optie ingebouwd
  └── Uitrollen via NinjaOne GPO
```

---

## Detect-AudioDevices.ps1

### Wat doet het script

Het script maakt een volledige inventaris van alle audio devices op de laptop en classificeert ze automatisch:

**Inventarisatie:**
- Alle AudioEndpoint devices met FriendlyName, InstanceId en Status
- Splitsing op basis van InstanceId: `{0.0.1...}` = microfoon (capture), `{0.0.0...}` = speaker (render)

**Classificatie microfoons:**
- Vermoedelijk intern: namen die matchen op Array, Intel, Realtek, Camera, Webcam, Integrated, etc.
- Vermoedelijk headset: namen die matchen op EPOS, Jabra, Yealink, Plantronics, USB, Headset, etc.
- Onbekend: alles wat niet geclassificeerd is

**Privacy check:**
- Controleert of het vorige foutieve script (`ConsentStore\microphone = Deny`) nog actief staat
- Als dat zo is weet je meteen welke laptop nog gereset moet worden

### Uitvoeren via NinjaOne

#### Script aanmaken

1. Ga naar NinjaOne → **Administration** → **Library** → **Scripting**
2. Klik op **Add** → **Script**
3. Vul in:
   - **Name**: `BNC - Detect Audio Devices`
   - **Language**: PowerShell
   - **Operating System**: Windows
   - **Architecture**: 64-bit
   - **Run As**: System
   - **Timeout**: 60 seconden
4. Plak de volledige inhoud van `Detect-AudioDevices.ps1` in het script veld
5. Klik **Save**

#### Script uitrollen op alle laptops

1. Ga naar NinjaOne → **Devices**
2. Filter op organisatie: **Best Next Contact**
3. Selecteer alle Windows laptops (checkboxes)
4. Klik op **Run Script** (boven in de toolbar)
5. Kies het script: `BNC - Detect Audio Devices`
6. Klik **Run**

#### Output lezen per laptop

1. Ga naar het specifieke device in NinjaOne
2. Klik op **Activities** (linkermenu)
3. Zoek de script run onder de activiteiten
4. Klik op de activiteit → **Details** → je ziet de volledige output

Of via overzicht:
1. NinjaOne → **Administration** → **Activity Log**
2. Filter op **Script**: `BNC - Detect Audio Devices`
3. Klik per device op **Details** om de output te lezen

#### Resultaten verzamelen

Maak een eenvoudige tabel van de output per laptop:

| Laptop | Gebruiker | Interne microfoon naam | Headset naam | ConsentStore status |
|---|---|---|---|---|
| LAPTOP-001 | Dean | Microfoon (Realtek Audio) | EPOS IMPACT DW | Allow |
| LAPTOP-002 | Ilias | ... | ... | ... |
| LAPTOP-003 | Salwa | ... | ... | ... |

Op basis van deze tabel bepalen we het correcte naampatroon voor het disable script.

### Output interpreteren

De output per laptop ziet er zo uit:

```
=== AUDIO DEVICE INVENTORY ===
Hostname   : LAPTOP-BNC-001
Datum      : 19/03/2026 09:00:00

--- MICROFOONS (capture devices - 0.0.1) ---
  [OK] Microfoon van hoofdtelefoon (EPOS IMPACT DW)
           SWD\MMDEVAPI\{0.0.1.00000000}.{6A4C41D0-...}
  [OK] Microfoon van hoofdtelefoon (EPOS IMPACT 60)
           SWD\MMDEVAPI\{0.0.1.00000000}.{...}
  [OK] Microfoon (Realtek(R) Audio)
           SWD\MMDEVAPI\{0.0.1.00000000}.{...}

--- VERMOEDELIJK INTERNE MICROFOONS ---
  [OK] Microfoon (Realtek(R) Audio)
           SWD\MMDEVAPI\{0.0.1.00000000}.{...}

--- VERMOEDELIJK HEADSETS ---
  [OK] Microfoon van hoofdtelefoon (EPOS IMPACT DW)
           SWD\MMDEVAPI\{0.0.1.00000000}.{...}
```

**Wat je zoekt:**
- De interne microfoon heeft een naam zoals `Microfoon (Realtek(R) Audio)`, `Microfoon (Intel SST)`, `Microfoon (High Definition Audio)`, of bevat `Array`
- De headset microfoon bevat de merknaam: `EPOS`, `Yealink`, `Jabra`, etc.
- Als een device op `[Error]` of `[Unknown]` staat was het al disabled

**Actie na analyse:**
- Noteer alle unieke namen van interne microfoons over alle laptops
- Controleer of er laptops zijn waar `ConsentStore\microphone = Deny` nog actief is (die moeten eerst gereset worden)
- Geef de namen door zodat het disable script correct geconfigureerd kan worden

---

## Fase 2: Disable script (nog te ontwikkelen)

Op basis van de detect output wordt een disable script gebouwd dat:

- De interne microfoon disabled via `Disable-PnpDevice -InstanceId "..."` (device-specifiek, niet privacy-breed)
- De headset **niet** aanraakt
- Een **rollback** optie heeft (`Enable-PnpDevice`) voor als er toch iets fout gaat
- Werkt op alle laptop modellen die bij Best Next Contact in gebruik zijn
- Uitrolbaar is via NinjaOne als geautomatiseerde actie
- Nieuwe laptops automatisch configureert bij onboarding

### Verschil met het vorige script

| | Vorig script (fout) | Nieuw script (correct) |
|---|---|---|
| Methode | Registry privacy policy | PnP device disable |
| Scope | Alle microfoons (breed) | Alleen interne microfoon |
| Headset | ❌ Ook geblokkeerd | ✅ Blijft actief |
| Rollback | Moeilijk | Eenvoudig via Enable-PnpDevice |
| Nieuwe laptops | Werkt automatisch | Vereist detect bij onboarding |

---

## Bestanden

```
scripts/
└── audio/
    └── best-next-contact/
        ├── Detect-AudioDevices.ps1     ← fase 1 (dit script)
        ├── Disable-InternalMic.ps1     ← fase 2 (nog te schrijven)
        └── README.md
```

---

## Changelog

| Datum | Versie | Wijziging |
|---|---|---|
| 18/09/2025 | — | Ticket geopend: achtergrondgeluid bij Adversus |
| 06/10/2025 | — | Testfase op 3 laptops (Dean, Ilias, Salwa) — deels succesvol |
| 17/11/2025 | — | Beslissing: interne microfoon disablen via registry |
| 04/03/2026 | — | Script uitgerold via NinjaOne — te breed, headset ook geblokkeerd |
| 05/03/2026 | — | Rollback uitgevoerd op laptop Yllana |
| 19/03/2026 | 1.0 | Detect script geschreven als fase 1 van correcte aanpak |