[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../../../readme.nl.md) › [scripts](../../../../readme.nl.md) › [Intune](../../../readme.nl.md) › [Desktop](../../readme.nl.md) › [Background](../readme.nl.md) › **Lockscreen**

# Make-lockscreen.ps1

> Auteur: Sjoerd Kanon

---

## Wat het doet

Downloadt de bedrijfsachtergrond van internet en past hetzelfde bestand via Intune toe als Windows-vergrendelscherm.

Het script schrijft de `PersonalizationCSP`-registerwaarden voor het vergrendelscherm in `HKLM`, zodat het vergrendelscherm op het hele apparaat wordt afgedwongen.

---

## Configuratie

Pas de variabelen in het `CONFIGURATION`-blok bovenaan het script aan voordat je het naar Intune uploadt:

```powershell
$ImageUrl   = "https://your-cdn.com/CUSTOMERNAME/wallpaper.png"
$ClientName = "CUSTOMERNAME"
```

Vervang `CUSTOMERNAME` door de klantnaam die je in de lokale bestandsnaam en de naam van het logbestand wilt gebruiken.

Dit script is bedoeld om dezelfde afbeelding te gebruiken als `Set-CorporateWallpaper.ps1`, zodat de huisstijl van bureaublad en vergrendelscherm op elkaar aansluit.

Host je de afbeelding op GitHub, dan heeft een raw-URL de voorkeur. Gangbare URL's van de vorm `github.com/.../blob/...` worden automatisch genormaliseerd.

---

## Logging

Logs worden weggeschreven naar:

```text
C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateLockscreen-<CUSTOMERNAME>.log
```

---

## Uitrol via Intune

### Als PowerShell-script

1. Intune -> Devices -> Scripts and remediations -> Platform scripts
2. Voeg een nieuw PowerShell-script voor Windows 10 and later toe
3. Upload `Make-lockscreen.ps1`
4. Gebruik deze instellingen:
   - Run this script using the logged on credentials: No
   - Enforce script signature check: No
   - Run script in 64-bit PowerShell: Yes
5. Wijs toe aan de gewenste apparaatgroep

### Als Win32-app

Gebruik een detectieregel op bestand voor de gedownloade afbeelding in:

```text
C:\ProgramData\Wallpapers\corporate-lockscreen-<customername>.<jpg|png|bmp>
```

---

## Hoe het werkt

| Stap | Actie |
|---|---|
| 1 | Maak `C:\ProgramData\Wallpapers\` aan als die nog niet bestaat |
| 2 | Normaliseer gangbare GitHub-download-URL's waar nodig |
| 3 | Download de afbeelding met `Invoke-WebRequest` |
| 4 | Controleer de bestandsgrootte en detecteer niet-ondersteunde of HTML-inhoud |
| 5 | Bepaal het werkelijke afbeeldingstype aan de hand van de bestandsheaders |
| 6 | Sla het bestand lokaal op met de juiste extensie |
| 7 | Schrijf `LockScreenImagePath`, `LockScreenImageUrl` en `LockScreenImageStatus` in `PersonalizationCSP` |
| 8 | Start een vernieuwing van de parameters met `RUNDLL32.EXE USER32.DLL, UpdatePerUserSystemParameters 1, True` |

---

## Versiegeschiedenis

| Datum | Versie | Wijziging |
|---|---|---|
| — | 1.0 | Eerste vergrendelschermversie met directe download via `WebClient` en vaste `.jpg`-uitvoer |
| 2026-04-16 | 2.0 | Bijgewerkt om dezelfde bron als de bedrijfsachtergrond te gebruiken; normalisatie van GitHub-/raw-URL's, afbeeldingsvalidatie, HTML-detectie, gestructureerde logging en een veiligere afhandeling van de tijdelijke download toegevoegd |
