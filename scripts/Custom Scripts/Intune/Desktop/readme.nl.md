[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../../readme.nl.md) › [scripts](../../../readme.nl.md) › [Custom Scripts](../../readme.nl.md) › [Intune](../readme.nl.md) › **Desktop**

# Desktop (Office-thema)

Uitrol van het Office-thema en -kleurenpalet via Intune. Beide scripts krijgen de download-URL van het thema als parameter, zodat één script voor elke klant werkt; de themabestanden zelf staan niet in deze repo.

Voor de uitrol van achtergrond/vergrendelscherm/taakbalksnelkoppeling, zie [`scripts/Intune/Desktop/`](../../../Intune/Desktop/readme.nl.md).

---

## Mappen

| Map | Omschrijving |
|--------|-------------|
| [`Office Themes/`](Office%20Themes/readme.nl.md) | `Deploy-Officecolors.ps1` — installeert alleen een kleurenschema |

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Deploy-OfficeTheme.ps1`](Deploy-OfficeTheme.ps1) ([docs](#deploy-officethemeps1)) | Downloadt een Office-thema `.thmx` van een URL en installeert het voor de aangemelde gebruiker |

---

### Deploy-OfficeTheme.ps1

Downloadt de `.thmx` van `-ThemeUrl` naar `%ProgramData%\OfficeThemes` en kopieert hem naar `%APPDATA%\Microsoft\Templates\Document Themes\`, zodat hij in Office verschijnt onder de keuzelijst **Ontwerpen > Thema's** (**Design > Themes**).

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-ThemeUrl` | Directe download-URL van het `.thmx`-bestand. Verplicht |
| `-ThemeName` | Bestandsnaam om het onder op te slaan, eindigend op `.thmx` — de naam die Office toont. Standaard: het laatste deel van de URL |

**Voorbeelden**

```powershell
.\Deploy-OfficeTheme.ps1 -ThemeUrl 'https://contoso.blob.core.windows.net/branding/Contoso.thmx'

# Installatieopdracht van een Win32-app
powershell.exe -ExecutionPolicy Bypass -File .\Deploy-OfficeTheme.ps1 -ThemeUrl 'https://example.com/theme.thmx' -ThemeName 'Contoso 2026.thmx'
```

**Opmerkingen**

- Uitvoeren als de aangemelde gebruiker (schrijft naar `%APPDATA%`).
- Intune-platformscripts kunnen geen parameters meegeven: rol het uit als Win32-app met de parameters in de installatieopdracht, of upload een kopie met de standaardwaarden ingevuld.
- Zonder `-ThemeUrl`, of met een naam die niet op `.thmx` eindigt, stopt het met exitcode 1 in plaats van te vragen — onder Intune beantwoordt niemand die vraag.
