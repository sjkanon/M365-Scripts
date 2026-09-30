[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **Intune**

# Intune

Autopilot-enrollment, automatisering van compliancebeleid, detectie van configuratiedrift en desktopuitrol bij klanten (achtergrond, vergrendelscherm, snelkoppeling om te vergrendelen via de taakbalk).

> De uitrol van Office-thema's en -kleuren staat in [`Custom Scripts/Intune/Desktop/`](../Custom%20Scripts/Intune/readme.nl.md) — die scripts hebben hun download-URL hard naar dat pad gecodeerd.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Compare-IntuneConfig.ps1`](Compare-IntuneConfig.ps1) ([docs](#compare-intuneconfigps1)) | Vergelijk de Intune-configuratie van een klanttenant met een MSP-baselineback-up |
| [`Repair-StuckWin32AppEnforcement.ps1`](Repair-StuckWin32AppEnforcement.ps1) ([docs](#repair-stuckwin32appenforcementps1)) | Maak Win32-apps vrij die op een apparaat vastzitten achter de GRS-retry-cooldown van Intune — handmatig en lokaal uit te voeren, met een proefdraai/rapport/`-AppId`-filter |
| [`Detect-StuckWin32AppEnforcement.ps1`](Detect-StuckWin32AppEnforcement.ps1) + [`Remediate-StuckWin32AppEnforcement.ps1`](Remediate-StuckWin32AppEnforcement.ps1) ([docs](#detect---remediate-stuckwin32appenforcementps1)) | Dezelfde oplossing, verpakt als Intune Remediation-paar — volledig vanuit de Intune-portal te starten, zonder toegang tot het apparaat |

## Mappen

| Map | Omschrijving |
|--------|-------------|
| [`Get-Autopilot/`](Get-Autopilot/readme.nl.md) | Verzamelen van Windows Autopilot-hardwarehashes |
| [`iOS-Compliance-Updater/`](iOS-Compliance-Updater/readme.nl.md) | Werkt de minimale iOS-versie in een Intune-compliancebeleid automatisch bij |
| [`Desktop/`](Desktop/readme.nl.md) | Bedrijfsachtergrond + vergrendelscherm, snelkoppeling om te vergrendelen via de taakbalk |
| [`DiskCleanup/`](DiskCleanup/readme.nl.md) | Win32-app-wrapper die het schijfopruimscript voor C:\ uitvoert en het apparaat herstart |

---

### Compare-IntuneConfig.ps1

Een wrapper rond de communitymodule `IntuneBackupAndRestore` om Intune-configuratiedrift op te sporen — compliancebeleid, configuratieprofielen en andere geëxporteerde objecten die afwijken van (of ontbreken/extra zijn ten opzichte van) een referentieback-up als "baseline". Alleen-lezen — wijzigt nooit configuratie, rapporteert alleen verschillen.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-BaselinePath` | Ja | Map met de MSP-referentieback-up (van een eerdere `Start-IntuneBackup -Path <path>`-run) |
| `-CustomerBackupPath` | Nee | Bestaande back-up van de klanttenant. Laat je deze weg, dan maakt het script eerst een back-up van de tenant waarmee je nu verbonden bent |
| `-TenantId` | Nee | Tenant-ID of domein om mee te verbinden (alleen gebruikt als `-CustomerBackupPath` is weggelaten) |
| `-OutputPath` | Nee | Map voor de automatische back-up en het verschillenrapport (standaard: `C:\Temp\` / `~/Downloads`) |

**Voorbeelden**

```powershell
# Vergelijk een bestaande klantback-up met de MSP-baseline
.\Compare-IntuneConfig.ps1 -BaselinePath C:\IntuneBaseline -CustomerBackupPath C:\Temp\CustomerBackup

# Maak een back-up van de verbonden (GDAP-)klanttenant en vergelijk die live
.\Compare-IntuneConfig.ps1 -BaselinePath C:\IntuneBaseline
```

**Opmerkingen**
- GDAP-bewust: als je `-TenantId` weglaat, wordt die automatisch bepaald uit de geselecteerde klanttenant (`$global:cid`), net als bij `Move-InboxToArchive.ps1` / `Get-SharePointStorageReport.ps1`
- Niet opgenomen in `menu.ps1` — werkt met paden naar back-upmappen en een tenantbrede export, dus voer het rechtstreeks uit

**Vereiste module**

```powershell
Install-Module IntuneBackupAndRestore -Scope CurrentUser
```

---

### Repair-StuckWin32AppEnforcement.ps1

Voer uit **op het betreffende apparaat**, als Administrator. De Win32-app-agent van Intune (IME) probeert een mislukte installatie 3 keer opnieuw, met 5 minuten ertussen, en zet de app daarna in een "GRS"-cooldown van 24 uur — zichtbaar in `AppActionProcessor.log` als `... to install is in GRS. The app will not be enforced.` Tijdens die cooldown negeert IME de app volledig, ook nadat je het eigenlijke probleem in Intune hebt opgelost (gecorrigeerde vereistenregel, nieuw pakket, gerepareerde installatieopdracht) — de cooldown is lokale apparaatstatus en kan niet worden gewist door vanuit Intune opnieuw uit te rollen.

Dit script zoekt onder `HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps` elke lokaal gecachte Win32-app met een echte laatste foutcode (niet 0/succes, niet 3010/herstart vereist). Het kan de gecachte afdwingingsstatus, de rapportagecache en de bijbehorende GRS-cooldownsleutel wissen, en herstart daarna de IME-service zodat elke app bij de volgende check-in opnieuw wordt beoordeeld.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-Apply` | Wis de vastgelopen status echt en herstart de service (standaard: alleen een proefdraairapport) |
| `-AppId` | Beperk tot één specifieke Win32-app-GUID (uit de app-URL in het Intune-beheercentrum). Standaard: elke gevonden vastgelopen app |
| `-ForceSync` | Start na het wissen ook direct een MDM-check-in (hetzelfde als de knop "Sync" in de Bedrijfsportal) |
| `-OutputPath` | Pad voor het CSV-rapport (standaard: `C:\Temp\`) |

**Voorbeelden**

```powershell
# Proefdraai — bekijk wat er nu op dit apparaat vastzit
.\Repair-StuckWin32AppEnforcement.ps1

# Wis alles wat vastzit en forceer direct een nieuwe synchronisatie
.\Repair-StuckWin32AppEnforcement.ps1 -Apply -ForceSync

# Wis alleen één specifieke app
.\Repair-StuckWin32AppEnforcement.ps1 -Apply -AppId "96ff358d-0e16-4224-b946-eb61bc930fca"
```

**Opmerkingen**
- Dit is lokale apparaatstatus — het raakt **elke** Win32-app die aan dat apparaat is toegewezen, niet maar één. Eén apparaat dat in GRS vastzit, kan er dus uitzien alsof meerdere losstaande app-uitrols tegelijk stilletjes mislukken.
- Gebaseerd op de door de community gedocumenteerde (niet officieel door Microsoft gepubliceerde) GRS-registerstructuur — zie `.NOTES` in het script voor de bronnen.

---

### Detect- / Remediate-StuckWin32AppEnforcement.ps1

Dezelfde oplossing als `Repair-StuckWin32AppEnforcement.ps1` hierboven, opgesplitst in een detectie/herstel-paar voor Intune **Devices → Scripts and remediations → Remediations**, zodat je voor het vrijmaken van een vastgelopen apparaat nooit een handmatige RDP-/consolesessie nodig hebt — alles wordt vanuit de Intune-portal gestart.

| Script | Rol |
|---|---|
| [`Detect-StuckWin32AppEnforcement.ps1`](Detect-StuckWin32AppEnforcement.ps1) | Detectiehelft — exit 1 als een Win32-app een echte laatste foutcode in de cache heeft (mogelijke GRS-blokkade), anders exit 0 |
| [`Remediate-StuckWin32AppEnforcement.ps1`](Remediate-StuckWin32AppEnforcement.ps1) | Herstelhelft — wist altijd alles wat het detectiescript vond, herstart IME en forceert direct een MDM-synchronisatie |

**Uitrollen in Intune:**

1. **Devices → Scripts and remediations → Remediations → Create**
2. Geef het een naam, bijv. `Clear Stuck Win32 App Enforcement`
3. Upload `Detect-StuckWin32AppEnforcement.ps1` als detectiescript en `Remediate-StuckWin32AppEnforcement.ps1` als herstelscript
4. Run using logged-on credentials: **No** (SYSTEM) · Run in 64-bit PowerShell: **Yes** · Enforce signature check: **No**
5. **Wijs het niet toe aan een apparaatgroep en stel geen schema in.** Laat de toewijzing leeg.

**Op aanvraag starten, per apparaat, volledig vanuit de portal** (geen schema, geen toegang tot het apparaat):

1. **Devices → All devices → [het betreffende apparaat]**
2. **… (beletselteken) → Run remediation (preview)**
3. Selecteer `Clear Stuck Win32 App Enforcement` en voer het uit

**Waarom niet toegewezen/op aanvraag in plaats van een lopend schema:** een geplande, vlootbrede toewijzing zou GRS-blokkades stilletjes blijven wissen voor *elke* app die 3 keer mislukt, om *welke* reden dan ook — waardoor een echt kapotte uitrol verborgen blijft achter eindeloze automatische nieuwe pogingen in plaats van zichtbaar te worden. Door het niet toe te wijzen en het alleen op een specifiek apparaat uit te voeren nadat je hebt vastgesteld (via `AppActionProcessor.log`/`AppWorkload.log`, of omdat het onderliggende probleem al is opgelost) dat een nieuwe poging echt zinvol is, voorkom je dat. Dit volgt het in de community gangbare patroon voor precies dit scenario (zie `.NOTES` in het script) en is niet voor deze repo verzonnen.
