[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../../../readme.nl.md) › [scripts](../../../../readme.nl.md) › [Intune](../../../readme.nl.md) › [Desktop](../../readme.nl.md) › [Background](../readme.nl.md) › **Desktop**

# Desktop

Bedrijfsachtergrond voor het bureaublad via Intune — instellen, en weer verwijderen.

## Scripts

| Script | Omschrijving |
|--------|--------------|
| [`Set-CorporateWallpaper.ps1`](Set-CorporateWallpaper.ps1) ([docs](#set-corporatewallpaperps1)) | De bedrijfsachtergrond downloaden en afdwingen voor alle gebruikers (PersonalizationCSP, huidige gebruiker, Default User) |
| [`Remove-CorporateWallpaper.ps1`](Remove-CorporateWallpaper.ps1) ([docs](#remove-corporatewallpaperps1)) | De bedrijfsachtergrond voor elke gebruiker terugdraaien — alleen wat `Set-CorporateWallpaper.ps1` schreef, de bedrijfs-lockscreen blijft |

---

## Set-CorporateWallpaper.ps1

> Auteur: Sjoerd Kanon

---

## Wat het doet

Downloadt een bedrijfsachtergrond vanaf een URL en past die via Intune toe op Windows-apparaten. Stelt de achtergrond in voor:

- De **huidige gebruiker** (WinAPI — direct van kracht)
- De **huidige gebruiker** (HKCU-register — legt de stijlinstelling vast)
- **Alle gebruikers** via MDM (PersonalizationCSP — afgedwongen via HKLM)
- **Nieuwe gebruikersaccounts** via het Default User-profiel (NTUSER.DAT)

Zo wordt de achtergrond toegepast ongeacht wie zich aanmeldt, zowel voor bestaande als voor nieuw aangemaakte accounts.

---

## Configuratie

Pas de drie variabelen in het `CONFIGURATION`-blok bovenaan het script aan voordat je het naar Intune uploadt:

```powershell
$ImageUrl       = "https://your-cdn.com/CUSTOMERNAME/wallpaper.png"
$ClientName     = "CUSTOMERNAME"
$WallpaperStyle = "10"
```

Vervang `CUSTOMERNAME` door de naam van de klant (bijv. `acme`). Al het andere ligt vast en hoeft niet te worden aangepast.

Host je de afbeelding op GitHub, gebruik dan een raw-bestands-URL (`raw.githubusercontent.com`) in plaats van een pagina-URL `github.com/.../blob/...`.
Het script kan gangbare GitHub blob-/raw-pagina-URL's automatisch normaliseren, maar direct een raw-URL gebruiken is het beste.

### Achtergrondstijlen

| Waarde | Stijl | Opmerkingen |
|---|---|---|
| `10` | Fill | Aanbevolen — vult het scherm zonder vervorming |
| `6` | Fit | Past binnen het scherm, zwarte randen mogelijk |
| `2` | Stretch | Rekt uit om te vullen, kan vervormen |
| `0` | Tile | Herhaalt de afbeelding |
| `22` | Span | Verspreidt over meerdere monitoren |

---

## Logging

Logs worden weggeschreven naar de standaard Intune-logmap:

```
C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateWallpaper-<CUSTOMERNAME>.log
```

Leesbaar via Intune Device Diagnostics of lokaal op het apparaat.

---

## Uitrol via Intune

### Als PowerShell-script

1. Intune → **Devices** → **Scripts and remediations** → **Platform scripts**
2. Klik op **Add** → **Windows 10 and later**
3. Instellingen:
   - **Script**: upload `Set-CorporateWallpaper.ps1`
   - **Run this script using the logged on credentials**: No
   - **Enforce script signature check**: No
   - **Run script in 64-bit PowerShell**: Yes
4. Wijs toe aan de gewenste apparaatgroep
5. **Save**

> Let op: het script draait als SYSTEM. Stap 5 (WinAPI) en 6 (HKCU) gelden voor de SYSTEM-context, niet voor de aangemelde gebruiker. PersonalizationCSP (stap 4) en Default User (stap 7) worden hoe dan ook correct toegepast voor alle gebruikers.

### Als Win32-app (aanbevolen)

Verpakken als Win32-app geeft je controle over opnieuw uitvoeren en detectieregels.

**Het script verpakken:**
```powershell
.\IntuneWinAppUtil.exe -c . -s Set-CorporateWallpaper.ps1 -o .
```

**Instellingen van de Intune-app:**

| Veld | Waarde |
|---|---|
| **Install command** | `powershell.exe -ExecutionPolicy Bypass -File Set-CorporateWallpaper.ps1` |
| **Uninstall command** | `cmd.exe /c echo uninstall` |
| **Detection rule** | Bestand bestaat: `C:\ProgramData\Wallpapers\corporate-background-<customername>.jpg` (of `.png` als de bronafbeelding een PNG is) |
| **Run as** | System |
| **Architecture** | 64-bit |

---

## Hoe het werkt (stap voor stap)

| Stap | Actie | Bereik |
|---|---|---|
| 1 | Maak `C:\ProgramData\Wallpapers\` aan als die nog niet bestaat | — |
| 2 | Download de afbeelding via `Invoke-WebRequest` | — |
| 3 | Laad de Windows API `user32.dll` | — |
| 4 | Schrijf de `PersonalizationCSP`-registersleutels (HKLM) | Alle gebruikers via MDM |
| 5 | Roep `SystemParametersInfo` (WinAPI) aan | Huidige gebruiker — direct |
| 6 | Schrijf `HKCU\Control Panel\Desktop` + vernieuwing via `RUNDLL32` | Huidige gebruiker — blijvend |
| 7 | Laad `Default\NTUSER.DAT` en schrijf dezelfde sleutels | Nieuwe gebruikersaccounts |

---

## Wijzigingslog

| Datum | Versie | Wijziging |
|---|---|---|
| — | 1.0 | Eerste versie |
| — | 1.2 | Generiek gemaakt voor hergebruik per klant |
| 2026-03-20 | 2.0 | Vertaald naar het Engels; `Invoke-WebRequest` vervangt `WebClient`; `#Requires -Version 5.1`; generieke placeholder voor de CDN-URL |
| 2026-04-14 | 2.1 | Validatie van de afbeeldingssignatuur en dynamische lokale extensie (`.jpg/.png/.bmp`) toegevoegd om ongeldige achtergrondbestanden te voorkomen |
| 2026-04-14 | 2.2 | Normalisatie van GitHub-URL's (`github.com/.../blob/...` naar `raw.githubusercontent.com`) en een controle op HTML-antwoorden toegevoegd om zwarte/lege achtergronden te voorkomen |
| 2026-04-14 | 2.3 | Terugvalafdwinging voor zwarte achtergronden toegevoegd: beleidssleutels voor de machineachtergrond + bijwerken van alle geladen gebruikershives; `DesktopImageUrl` gebruikt nu de bron-URL |
| 2026-04-14 | 2.4 | Fout in de opruimvolgorde opgelost waarbij het tijdelijke `.download`-bestand vóór `Move-Item` kon worden verwijderd, wat een path-not-found-fout gaf |
| 2026-04-14 | 2.5 | Stap toegevoegd die `explorer.exe` herstart, zodat wijzigingen in achtergrond/thema direct zichtbaar worden voor aangemelde gebruikers |
| 2026-04-14 | 2.6 | Generieke standaardconfiguratiewaarden (`$ImageUrl`, `$ClientName`) hersteld voor herbruikbare klantuitrols |
| 2026-04-14 | 2.7 | Failsafe-back-up van de huidige achtergrond toegevoegd en de vervangvolgorde aangepast, zodat de vorige achtergrond beschikbaar blijft als de update mislukt |

---

## Remove-CorporateWallpaper.ps1

Draait terug wat `Set-CorporateWallpaper.ps1` heeft ingesteld, en alleen dat. Uitvoeren als SYSTEM (Intune-platformscript, of de verwijderopdracht van de Win32-app).

1. **PersonalizationCSP** — verwijdert de waarden `DesktopImagePath`, `DesktopImageUrl` en `DesktopImageStatus`. De sleutel zelf verdwijnt alleen als hij daarna leeg is: [`Make-lockscreen.ps1`](../Lockscreen/Make-lockscreen.ps1) bewaart zijn `LockScreen*`-waarden in dezelfde sleutel
2. **`HKLM\...\Policies\System`** — verwijdert `Wallpaper` en `WallpaperStyle`, maar alleen als ze naar een bedrijfsachtergrond wijzen; een beleid dat iets anders zette, blijft staan
3. **Elke geladen gebruikershive** (`S-1-5-21-*`) — een achtergrond die naar een bedrijfsbestand wijst, gaat terug naar de Windows-standaard (`img0.jpg`, stijl Opvullen), en de getranscodeerde achtergrondcache van die gebruiker wordt gewist zodat de oude afbeelding niet blijft hangen
4. **`HKCU` van de huidige gebruiker** — dezelfde reset, maar alleen als het niet als SYSTEM draait (als SYSTEM is `HKCU` het eigen profiel van SYSTEM en dekte stap 3 de gebruikers al)
5. **Default User-profiel** (`C:\Users\Default\NTUSER.DAT`) — dezelfde reset, zodat nieuwe accounts de achtergrond niet meer krijgen. Overgeslagen met een melding als de hive niet geladen kan worden (niet verhoogd)
6. **Bestanden** — verwijdert de `corporate-background-*`-bestanden in `C:\ProgramData\Wallpapers`. De map verdwijnt alleen als hij leeg is — de lockscreen-afbeelding staat er ook
7. Herstart `explorer.exe` zodat de wijziging meteen zichtbaar is

"Een bedrijfsbestand" is een `corporate-background-*`-bestand in `C:\ProgramData\Wallpapers` — de naam die `Set-CorporateWallpaper.ps1` het geeft.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-WhatIf` | Toon elke waarde en elk bestand dat verwijderd of gereset zou worden; wijzig niets |

**Voorbeelden**

```powershell
# Als SYSTEM (Intune-platformscript / verwijderopdracht van de Win32-app)
powershell.exe -ExecutionPolicy Bypass -File Remove-CorporateWallpaper.ps1

# Zien wat het zou doen
.\Remove-CorporateWallpaper.ps1 -WhatIf
```

**Opmerkingen**

- Logboek: `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateWallpaper-Remove.log`
- Gebruikers waarvan het profiel niet geladen is (niet aangemeld) houden de achtergrondwaarde in hun eigen hive; Windows toont de standaard zodra het bestand weg is, en de volgende run van dit script terwijl ze aangemeld zijn, zet de waarde terug
- Eerdere versies verwijderden de hele PersonalizationCSP-sleutel en de hele map `C:\ProgramData\Wallpapers`, waardoor ook de bedrijfs-lockscreen verdween
