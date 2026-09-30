[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [TenantOnboarding](../readme.nl.md) › **OneDriveManagement**

# OneDriveManagement

Onderhoudsscripts voor OneDrive for Business aan de devicekant: een watchdog voor herstarten/resetten, per bibliotheek synchronisatie afbreken, en omleiding via Known Folder Move. Staat los van [`scripts/Device/DriveMapping/`](../../Device/DriveMapping/readme.nl.md) (SharePoint/OneDrive koppelen aan een stationsletter via WebDAV), dat een andere functionaliteit is.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Register-OneDriveWatchdog.ps1`](Register-OneDriveWatchdog.ps1) ([docs](#register-onedrivewatchdogps1)) | Houd OneDrive draaiende via een geplande taak; ondersteunt ook een eenmalige `/reset` |
| [`Stop-OneDriveLibrarySync.ps1`](Stop-OneDriveLibrarySync.ps1) ([docs](#stop-onedrivelibrarysyncps1)) | Stop met het synchroniseren van één specifieke bibliotheek zonder andere aan te raken |
| [`Set-OneDriveKnownFolderRedirect.ps1`](Set-OneDriveKnownFolderRedirect.ps1) ([docs](#set-onedriveknownfolderredirectps1)) | Leid Bureaublad/Documenten/Afbeeldingen/Downloads om naar OneDrive |

---

### Register-OneDriveWatchdog.ps1

Registreert een geplande taak (draait als de huidige gebruiker, standaard elke 59 minuten) die OneDrive opnieuw start als het niet draait, of het optioneel bij elke run geforceerd herstart. `-Reset` voert in plaats daarvan een eenmalige `onedrive.exe /reset` uit voor een vastgelopen synchronisatieprofiel.

| Parameter | Omschrijving |
|-----------|-------------|
| `-IntervalMinutes` | Interval van de watchdog (standaard: `59`) |
| `-ForceRestartOnEachRun` | Stop en start OneDrive bij elke trigger, niet alleen als het niet draait |
| `-Reset` | Voer een eenmalige `/reset` uit in plaats van de watchdog te registreren |
| `-Apply` | Registreer/reset echt (standaard: alleen voorbeeldweergave) |

```powershell
.\Register-OneDriveWatchdog.ps1 -Apply
.\Register-OneDriveWatchdog.ps1 -ForceRestartOnEachRun -IntervalMinutes 30 -Apply
.\Register-OneDriveWatchdog.ps1 -Reset -Apply
```

---

### Stop-OneDriveLibrarySync.ps1

Stopt netjes met het synchroniseren van één specifieke SharePoint/OneDrive-bibliotheek: sluit OneDrive af, verwijdert het beleidsbestand en de cachevermelding van het koppelpunt van die bibliotheek, past de ini van de synchronisatiedatabase aan, verwijdert de lokale map en start OneDrive opnieuw, allemaal zonder andere gesynchroniseerde bibliotheken te verstoren.

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-MatchPattern` | Ja | Onderscheidende tekst waarmee het beleidsbestand van de bibliotheek wordt herkend |
| `-LocalFolderPath` | Ja | Lokale gesynchroniseerde map die wordt verwijderd |
| `-Apply` | Nee | Voer het afbreken echt uit (standaard: alleen voorbeeldweergave) |

```powershell
.\Stop-OneDriveLibrarySync.ps1 -MatchPattern "ProjectX" -LocalFolderPath "$env:USERPROFILE\Contoso\ProjectX - Documents" -Apply
```

> Werkt op ongedocumenteerde interne statusbestanden van OneDrive. Test op één machine voordat je breder uitrolt.

---

### Set-OneDriveKnownFolderRedirect.ps1

Leidt bekende mappen om naar de synchronisatiemap van OneDrive for Business met `SHSetKnownFolderPath`, migreert bestaande inhoud met Robocopy en verbergt (verwijdert niet) de oorspronkelijke map. Schakelt optioneel de OneDrive-vlag `Timerautomount` in, zodat eerder gesynchroniseerde bibliotheken op een nieuw device automatisch gekoppeld worden.

| Parameter | Omschrijving |
|-----------|-------------|
| `-Folders` | Welke mappen worden omgeleid (standaard: Desktop, Documents, Pictures, Downloads) |
| `-EnableAutoMountSharedLibraries` | Schakel ook `Timerautomount` in |
| `-Apply` | Leid echt om (standaard: alleen voorbeeldweergave) |

```powershell
.\Set-OneDriveKnownFolderRedirect.ps1 -Apply
.\Set-OneDriveKnownFolderRedirect.ps1 -Folders Desktop,Documents -EnableAutoMountSharedLibraries -Apply
```

> Voer uit in de context van de aangemelde gebruiker, niet als SYSTEM: het script leest de registersleutel van het OneDrive-account van die gebruiker.

---

Alle scripts doen standaard een proefdraai; geef `-Apply` mee om wijzigingen door te voeren, volgens de huisstijl.
