# Audio - Interne Microfoon Disable

> **BraveHub Internal Script**
> Auteur: Sjoerd Kanon | Datum: 19/03/2026

---

## Achtergrond en probleemstelling

De klant gebruikt een **browsergebaseerde VoIP/dialer applicatie**. Medewerkers hoorden achtergrondgeluid tijdens gesprekken terwijl ze een headset ingeplugd hadden.

De oorzaak: Windows selecteert soms de **interne laptopmicrofoon** als audiobron in plaats van de headset, ook als de headset correct ingeplugd is. Dit gedrag treedt op in browsergebaseerde applicaties omdat de browser zelf de audiobron kiest op basis van Windows-instellingen.

Applicaties met eigen audiobeheer (zoals Microsoft Teams) hebben dit probleem niet. Browsergebaseerde applicaties vertrouwen volledig op de Windows-instellingen en pakken soms de verkeerde bron.

### Tijdlijn

| Datum | Actie |
|---|---|
| Stap 1 | Ticket geopend: klacht over achtergrondgeluid bij dialer applicatie |
| Stap 2 | Testfase op enkele laptops — deels succesvol |
| Stap 3 | Beslissing: interne microfoon disablen via script |
| Stap 4 | Eerste script uitgerold via NinjaOne — te breed, headset ook geblokkeerd |
| Stap 5 | Rollback uitgevoerd op getroffen laptop |
| Stap 6 | Detect script uitgerold voor inventarisatie |
| Stap 7 | Disable + Rollback script gebouwd op basis van detect data |

---

## Waarom het eerste script mislukte

Het script dat op 04/03/2026 werd uitgerold blokkeerde microfoonpermissies via het Windows Privacy / Consent Store mechanisme:

```powershell
# FOUT - te breed
$regPath = "HKLM:\...\CapabilityAccessManager\ConsentStore\microphone"
Set-ItemProperty -Path $regPath -Name "Value" -Value "Deny"
```

Dit blokkeert **alle** microfooninput voor **alle** applicaties — inclusief de headset. Na uitrol kon Yllana helemaal niet meer bellen. Het script moest teruggedraaid worden.

### Verschil oude vs nieuwe aanpak

| | Oud script (fout) | Nieuw script (correct) |
|---|---|---|
| Methode | Registry privacy policy | PnP device disable |
| Scope | Alle microfoons | Alleen interne microfoon |
| Headset | ❌ Ook geblokkeerd | ✅ Blijft actief |
| Rollback | Moeilijk | Eenvoudig via apart script |
| Nieuwe laptops | Werkt automatisch | Heruitrol nodig bij onboarding |

---

## Aanpak in drie fases

```
Fase 1: Detect  ✅ GEDAAN (16/66 devices)
  └── Detect-AudioDevices.ps1 uitgerold via NinjaOne
  └── Custom field AudioDeviceInventory gevuld per device
  └── Patronen voor interne microfoons geïdentificeerd

Fase 2: Disable  ← VOLGENDE STAP
  └── Disable-InternalMic.ps1 uitrollen
  └── Eerst testfase op 5-10 laptops
  └── Dan uitrollen op alle devices

Fase 3: Monitoring
  └── Bevestiging van gebruikers dat headset werkt
  └── Uitrollen op resterende 50 devices zonder detect data
  └── Opnemen in onboarding procedure nieuwe laptops
```

---

## Detect analyse resultaten (19/03/2026)

Van de uitgerolde devices heeft het detect script data opgehaald. Devices die offline waren of waarop het script niet is uitgevoerd hebben geen data in het custom field.

### Gevonden interne microfoon namen

Op basis van de 16 devices zijn de volgende interne microfoons geïdentificeerd:

| Naam | Aantal devices | Actie |
|---|---|---|
| `Microfoonmatrix (Realtek High Definition Audio)` | 10 | Disablen |
| `Microphone (2- High Definition Audio Device)` | 4 | Disablen |
| `Microphone Array (Realtek High Definition Audio)` | 3 | Disablen |
| `Microphone (Realtek(R) Audio)` | 1 | Disablen |
| `Microfoonmatrix (Realtek(R) Audio)` | 1 | Disablen |
| `Microfoonmatrix (Synaptics Audio)` | 1 | Disablen |

### Gevonden headsets (worden NOOIT aangeraakt)

De volgende headset merknamen zijn aanwezig in de omgeving en worden door het script altijd overgeslagen:

- EPOS IMPACT DW
- EPOS IMPACT 60
- EPOS IMPACT D
- Yealink WH64
- Plantronics Blackwire 3225 Series

> Pas de safelist in het script aan als er nieuwe headset merken worden ingezet.

### Consent Store status

✅ Geen enkel device heeft nog `ConsentStore\microphone = Deny` — het eerste foutieve script is overal correct teruggedraaid.

### Onbekende microfoons

Op sommige devices kunnen niet-geclassificeerde devices voorkomen. Controleer de `ONBEKENDE MICROFOONS` sectie in de detect output per device en beoordeel manueel of het een intern device of een externe headset betreft.

---

## Bestanden

```
scripts/
└── audio/
    └── best-next-contact/
        ├── Detect-AudioDevices.ps1      ← fase 1: inventariseer audio devices
        ├── Disable-InternalMic.ps1      ← fase 2: disable interne microfoon
        ├── Rollback-InternalMic.ps1     ← noodstop: zet microfoon terug aan
        └── README.md
```

---

## Detect-AudioDevices.ps1

### Wat doet het

Maakt een volledige inventaris van alle audio devices op de laptop en classificeert ze automatisch als intern of headset. Schrijft het resultaat naar het NinjaOne custom field `AudioDeviceInventory`.

### Classificatielogica

Het script splitst microfoons op basis van de InstanceId:
- `{0.0.1...}` = capture device (microfoon)
- `{0.0.0...}` = render device (speaker/output)

Vervolgens worden microfoons geclassificeerd op basis van naampatronen:

**Intern** (worden gemarkeerd als te disablen):
`Array`, `Intel`, `Realtek`, `Synaptics`, `High Definition Audio`, `Camera`, `Webcam`, `Integrated`, `Microfoonmatrix`, `Microphone (`

**Headset** (worden nooit aangeraakt):
`EPOS`, `Jabra`, `Yealink`, `Plantronics`, `Poly`, `Logitech`, `Headset`, `hoofdtelefoon`, `oortelefoon`

### NinjaOne uitrol

#### Script aanmaken

1. NinjaOne → **Administration** → **Library** → **Scripting**
2. Klik **Add** → **Script**
3. Instellingen:
   - **Name**: `BNC - Detect Audio Devices`
   - **Language**: PowerShell
   - **OS**: Windows
   - **Architecture**: 64-bit
   - **Run As**: System
   - **Timeout**: 60 seconden
4. Plak inhoud van `Detect-AudioDevices.ps1`
5. **Save**

#### Uitrollen

1. NinjaOne → **Devices** → filter op **Best Next Contact**
2. Selecteer alle Windows laptops
3. **Run Script** → kies `BNC - Detect Audio Devices`
4. **Run**

#### Output lezen

Via custom field (aanbevolen voor 66 devices tegelijk):
1. NinjaOne → **Reports** → **Device Report**
2. Kolom `AudioDeviceInventory` toevoegen
3. Exporteren naar CSV → alle output in één bestand

Per device individueel:
1. Open het device → **Activities**
2. Klik op de script run → **Details**

---

## Disable-InternalMic.ps1

### Wat doet het

Disabled device-specifiek de interne microfoon op basis van de patronen uit de detect analyse. De headset wordt **nooit** aangeraakt.

### Logica per microfoon (4 checks)

```
Voor elke microfoon:
  1. Zit de naam in de headset safelist?
     → JA: skip, nooit aanraken
  2. Matcht de naam op een intern patroon?
     → NEE: skip, onbekend device
  3. Is het device al disabled (Status = Unknown/Error)?
     → JA: skip, al in orde
  4. Disable via Disable-PnpDevice -InstanceId
     → Succes: log DISABLED
     → Fout: log FAILED, exit 1
```

### Interne patronen die worden gedisabled

```
*Microfoonmatrix*
*Microphone Array*
*Microphone*Realtek*
*Microphone*High Definition Audio*
*Microfoon*Realtek*
*Microfoon*Synaptics*
*Microphone (2-*
```

### Headset safelist (worden NOOIT aangeraakt)

```
EPOS, Jabra, Plantronics, Yealink, Poly, Logitech,
Headset, hoofdtelefoon, oortelefoon, Blackwire,
Voyager, Sennheiser, Astro, SteelSeries
```

### NinjaOne uitrol

#### Script aanmaken

1. NinjaOne → **Administration** → **Library** → **Scripting**
2. Klik **Add** → **Script**
3. Instellingen:
   - **Name**: `BNC - Disable Internal Mic`
   - **Language**: PowerShell
   - **OS**: Windows
   - **Architecture**: 64-bit
   - **Run As**: System
   - **Timeout**: 60 seconden
4. Plak inhoud van `Disable-InternalMic.ps1`
5. **Save**

#### Uitrolstrategie (gefaseerd)

> **Belangrijk:** niet meteen op alle 66 devices uitrollen. Gefaseerd werken om herhaling van eerdere incidenten te vermijden.

**Testfase (5-10 laptops):**
1. Selecteer 5-10 laptops waarvan de detect data al beschikbaar is
2. Run Script → `BNC - Disable Internal Mic`
3. Bel de gebruikers op om te bevestigen dat hun headset nog werkt
4. Wacht minimaal een halve dag voor je verder gaat

**Volledige uitrol (na bevestiging testfase):**
1. Selecteer alle resterende laptops
2. Run Script → `BNC - Disable Internal Mic`
3. Controleer de output via het `AudioDeviceInventory` custom field

#### Exit codes

| Code | Betekenis |
|---|---|
| `0` | Alles succesvol |
| `1` | Eén of meer devices konden niet worden gedisabled — controleer output |

---

## Rollback-InternalMic.ps1

### Wanneer gebruiken

Meteen uitrollen op een specifiek device als een medewerker meldt dat zijn headset niet meer werkt na het disable script. Zet exact dezelfde devices terug aan als het disable script heeft uitgezet.

### NinjaOne uitrol

1. NinjaOne → open het **specifieke device** met het probleem
2. **Run Script** → kies `BNC - Rollback Internal Mic`
3. **Run**
4. Bel de gebruiker op om te bevestigen dat de headset weer werkt

> ⚠️ Rol het rollback script alleen uit op het device dat problemen heeft, niet op alle devices tegelijk.

---

## Devices zonder detect data

Devices die offline waren tijdens de detect uitrol hebben geen data in het custom field. Werkwijze:

1. NinjaOne → **Devices** → filter op de klantorganisatie
2. Filter op custom field `AudioDeviceInventory` = leeg
3. Controleer of de devices online zijn
4. Uitrol detect script opnieuw op deze devices
5. Herhaal tot alle devices data hebben

---

## Onboarding nieuwe laptops

Wanneer een nieuwe laptop wordt toegevoegd:
1. Detect script uitrollen om de interne microfoon naam te bevestigen
2. Controleer of de naam matcht op een bestaand patroon in het disable script
3. Zo ja: disable script uitrollen
4. Zo nee: patroon toevoegen aan `$internalPatterns` in het disable script en versie ophogen in de changelog

---

## Changelog

| Datum | Versie | Wijziging |
|---|---|---|
| 04/03/2026 | — | Eerste (foutief) script uitgerold — ConsentStore = Deny, te breed |
| 05/03/2026 | — | Rollback uitgevoerd op getroffen laptop |
| 19/03/2026 | 1.0 | Detect script geschreven en uitgerold |
| 19/03/2026 | 1.1 | Detect output geanalyseerd — 6 interne microfoon patronen gevonden |
| 19/03/2026 | 1.2 | Disable script gebouwd op basis van detect data |
| 19/03/2026 | 1.3 | Rollback script toegevoegd |
| 19/03/2026 | 1.4 | README generiek gemaakt voor hergebruik bij andere klanten |