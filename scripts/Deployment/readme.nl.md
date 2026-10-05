[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **Deployment**

# Setup-toolkit — USB / OOBE

> Auteur: Sjoerd Kanon

Een USB-toolkit voor Windows-installatie en Autopilot-inschrijving. Ontworpen om tijdens OOBE (Out-of-Box Experience) te gebruiken via `Shift+F10`.

---

## Scripts

| Script | Omschrijving |
|--------|--------------|
| [`start.bat`](start.bat) ([docs](#startbat)) | Hoofdmenu van de toolkit — vraagt zelf om beheerdersrechten en biedt Autopilot-inschrijving, Windows Update, hernoemen, domeinlidmaatschap en de browser voor klantinstallaties |
| [`Browse-InstallScripts.ps1`](Browse-InstallScripts.ps1) ([docs](#browse-installscriptsps1)) | Interactieve browser voor klanten en scripts achter menuopties `D` en `E` — door klantmappen bladeren en `.ps1`- / `.bat`- / `.cmd`-bestanden starten |

Ook in deze map: [`autorun.inf`](autorun.inf) — alleen het label van de USB-stick ([details](#autoruninf)).

---

## Mapstructuur op de USB

Alle bestanden moeten in **dezelfde map** op de USB-stick staan:

```
USB:\
├── start.bat                      ← Hoofdmenu — voer dit uit
├── GetAutoPilot.CMD               ← Script voor Autopilot-inschrijving
├── Get-WindowsAutoPilotInfo.ps1   ← PowerShell-module voor de hardwarehash
├── Browse-InstallScripts.ps1       ← Browser voor klantinstallaties voor optie D/E
├── Install\                        ← Lokale installatie-inhoud per klant voor optie D
└── autorun.inf                    ← Label van de USB-stick (alleen cosmetisch)
```

> `GetAutoPilot.CMD` en `Get-WindowsAutoPilotInfo.ps1` staan in `scripts/Intune/Get-Autopilot/` — kopieer ze allebei naar de USB.
> Kopieer voor menuoptie `D` zowel `Browse-InstallScripts.ps1` als de volledige map `Install` naar dezelfde USB-locatie als `start.bat`.

---

## Gebruik tijdens OOBE

Windows voert USB-scripts **niet** automatisch uit (geblokkeerd sinds Vista). Handmatige stappen:

1. Steek de USB-stick in
2. Druk tijdens OOBE op **Shift+F10** om een opdrachtprompt te openen
3. Zoek de stationsletter van de USB-stick — meestal `D:` of `E:`:
   ```
   D:
   dir
   ```
4. Start de toolkit:
   ```
   start.bat
   ```
5. Het script vraagt automatisch om beheerdersrechten en toont het menu

---

## start.bat

Het hoofdmenu van de toolkit. Het wisselt naar de eigen map (`cd /d %~dp0`), vraagt om beheerdersrechten en toont deze opties:

| Optie | Actie | Werkt in OOBE |
|---|---|---|
| `1` | Apparaatbeheer openen | ✅ |
| `2` | Autopilot-inschrijving — hash opslaan in `compHash.csv` | ✅ |
| `3` | Hashbestand verwijderen + Autopilot opnieuw uitvoeren (opslaan naar CSV) | ✅ |
| `4` | **Autopilot-inschrijving online** — hash rechtstreeks naar Intune uploaden | ✅ (internet + adminaccount nodig) |
| `5` | Windows Update via `PSWindowsUpdate` (`Install-WindowsUpdate -AcceptAll -AutoReboot`) | ✅ (internet nodig) |
| `6` | PowerShell 7 installeren via `winget` | ✅ (internet nodig) |
| `7` | Productsleutel invoeren (`slui.exe`) | ✅ |
| `8` | Lid worden van een Active Directory-domein | ✅ (domeinconnectiviteit nodig) |
| `9` | Herstarten (5 seconden vertraging) | ✅ |
| `A` | **Alles in één — Intune** — Hernoemen + Autopilot online + Windows Update + herstarten | ✅ (internet + adminaccount nodig) |
| `B` | **Apparaat hernoemen** — vraagt om een prefix en voegt het serienummer toe (`PREFIX-SERIALNUMBER`) | ✅ |
| `C` | **Alles in één — AD** — Hernoemen + domeinlidmaatschap + Windows Update + herstarten | ✅ (domeinconnectiviteit nodig) |
| `D` | **Installatiescripts per klant (lokaal)** — klantmenu openen vanuit de lokale map `Install` | ✅ |
| `E` | **Installatiescripts per klant (netwerkshare)** — klantmenu openen vanuit `\\10.222.3.94\Software` | ✅ (netwerktoegang nodig) |
| `0` | Afsluiten | ✅ |

### Browse-InstallScripts.ps1

Browser voor klantinstallaties achter menuopties `D` en `E`. Toont de klantmappen op het eerste niveau onder `-RootPath` als menu, en laat je daarna erin bladeren en `.ps1`-, `.bat`- en `.cmd`-bestanden starten (mappen `AppDeployToolkit` worden verborgen).

| Parameter | Verplicht | Omschrijving |
|-----------|-----------|--------------|
| `-RootPath` | Ja | Map waarvan de submappen de klanten zijn (lokale map `Install` of een netwerkshare) |
| `-SourceLabel` | Nee | Titel boven het klantmenu (standaard `Install Scripts`) |

```powershell
# Wat optie D uitvoert
powershell -NoProfile -ExecutionPolicy Bypass -File .\Browse-InstallScripts.ps1 -RootPath .\Install -SourceLabel "Local Install"
```

- Optie `D` heeft lokale bestanden nodig: `Browse-InstallScripts.ps1` en de volledige map `Install` naast `start.bat`.
- Optie `E` leest klantmappen uit `\\10.222.3.94\Software` en heeft netwerktoegang nodig.
- Voordat optie `D` of `E` de deploybrowser opent, bereidt `start.bat` het apparaat voor op de uitrol:
   - Maakt de lokale admingebruiker `LocalAdmin` aan of werkt die bij
   - Wachtwoord: `<wachtwoord weggelaten>`
   - Voegt `LocalAdmin` toe aan de lokale groep `Administrators`
   - Zet registervlaggen om OOBE over te slaan, zodat de rest van de OOBE-flow makkelijker kan worden overgeslagen

### Autopilot online (optie 4)

Voert `Get-WindowsAutoPilotInfo.ps1 -Online` uit — uploadt de hardwarehash rechtstreeks naar Intune zonder een CSV-bestand te maken. Vraagt om admin-inloggegevens voor Microsoft 365. Het apparaat verschijnt binnen enkele minuten in **Intune → Devices → Enroll devices → Windows enrollment → Autopilot devices**.

> Het instellingenpaneel van Windows Update is niet beschikbaar in OOBE, maar `UsoClient` start updates rechtstreeks vanaf de opdrachtregel en dat werkt prima.

### PowerShell 7 (optie 5)

Gebruikt `winget install Microsoft.PowerShell`. Vereist internet. Als `winget` niet beschikbaar is (oudere Windows 10), toont het script de URL voor handmatige download. Start na de installatie met `pwsh.exe`.

### Alles in één — Intune (optie A)

Voor Intune-/cloudbeheerde omgevingen. Voert achtereenvolgens uit:
1. Hernoemt het apparaat — vraagt om een prefix en voegt het serienummer toe (`PREFIX-SERIALNUMBER`)
2. Verwijdert de bestaande `compHash.csv`
3. Voert de online Autopilot-inschrijving uit (`Get-WindowsAutoPilotInfo.ps1 -Online`)
4. Installeert Windows-updates via `PSWindowsUpdate`
5. Herstart na 30 seconden (Ctrl+C om te annuleren)

### Alles in één — Active Directory (optie C)

Voor on-premises AD-omgevingen (geen Intune). Voert achtereenvolgens uit:
1. Hernoemt het apparaat — vraagt om een prefix en voegt het serienummer toe (`PREFIX-SERIALNUMBER`)
2. Voegt het apparaat toe aan het Active Directory-domein — vraagt om de domeinnaam en admin-inloggegevens
3. Installeert Windows-updates via `PSWindowsUpdate`
4. Herstart na 30 seconden (Ctrl+C om te annuleren)

---

## autorun.inf

Zet het label van de USB-stick op `Setup Toolkit` wanneer die wordt ingestoken. Voert **niets** automatisch uit — Windows blokkeert USB-autorun op alle moderne versies (Vista+).

---

## Wijzigingslog

| Datum | Versie | Wijziging |
|---|---|---|
| 2026-04-17 | 2.9 | Browser voor installaties per klant toegevoegd aan `start.bat`: optie `D` opent de lokale klantmappen in `Install` en optie `E` opent `\\10.222.3.94\Software`; `Browse-InstallScripts.ps1` toegevoegd om door klantmappen te bladeren en `.ps1`- / `.bat`- / `.cmd`-scripts uit te voeren; gedocumenteerd dat voor optie `D` zowel `Browse-InstallScripts.ps1` als de volledige map `Install` gekopieerd moeten worden; opties `D` en `E` maken nu de lokale admin `LocalAdmin` (`<wachtwoord weggelaten>`) aan of werken die bij, en zetten OOBE-overslaanvlaggen voordat de uitrol begint |

| Datum | Versie | Wijziging |
|---|---|---|
| 2026-03-20 | 2.8 | Alles in één gesplitst in A (Intune) en C (Active Directory); de AD-variant slaat Autopilot over |
| 2026-03-20 | 2.7 | Alles in één bijgewerkt: domeinlidmaatschap voor AD toegevoegd als stap 3 |
| 2026-03-20 | 2.6 | Alles in één bijgewerkt: apparaat hernoemen toegevoegd als eerste stap |
| 2026-03-20 | 2.5 | Optie om het apparaat te hernoemen toegevoegd (B) — vraagt om een prefix, voegt het serienummer toe (`Get-WmiObject Win32_BIOS`), max. 15 tekens |
| 2026-03-20 | 2.4 | Optie om lid te worden van een Active Directory-domein toegevoegd (8) — `Add-Computer` via PowerShell, vraagt om domein + inloggegevens |
| 2026-03-20 | 2.3 | Optie Autopilot online toegevoegd (4) — `Get-WindowsAutoPilotInfo.ps1 -Online`; Alles in één gebruikt nu online inschrijving |
| 2026-03-20 | 2.2 | Windows Update gebruikt nu de module `PSWindowsUpdate` (`Install-WindowsUpdate -AcceptAll -AutoReboot`) in plaats van `UsoClient` |
| 2026-03-20 | 2.1 | Windows Update, installatie van PowerShell 7 (`winget`) en de optie Alles in één (`A`) toegevoegd |
| 2026-03-20 | 2.0 | Herschreven naar het Engels; `cd /d %~dp0` voor het USB-pad; zelfverhoging van rechten; alleen OOBE-compatibele opties; paden tussen aanhalingstekens; `autorun.inf` en readme toegevoegd |
