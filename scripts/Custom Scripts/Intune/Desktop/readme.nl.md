[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../../readme.nl.md) › [scripts](../../../readme.nl.md) › [Custom Scripts](../../readme.nl.md) › [Intune](../readme.nl.md) › **Desktop**

# Desktop (Office-thema)

Uitrol van het Office-thema en -kleurenpalet via Intune. Staat bewust op dit pad — beide scripts hebben hun download-URL hard naar precies deze repolocatie gecodeerd (branch `main`), dus verplaatsen zou de download breken tot de scripts zijn bijgewerkt en opnieuw naar Intune zijn uitgerold.

Voor de uitrol van achtergrond/vergrendelscherm/taakbalksnelkoppeling, zie [`scripts/Intune/Desktop/`](../../../Intune/Desktop/readme.nl.md).

---

## Mappen

| Map | Omschrijving |
|--------|-------------|
| [`Office Themes/`](Office%20Themes/readme.nl.md) | `Deploy-Officecolors.ps1` — installeert alleen het kleurenschema |

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Deploy-OfficeTheme.ps1`](Deploy-OfficeTheme.ps1) ([docs](#deploy-officethemeps1)) | Installeert het volledige Office-thema `.thmx` van VIAS Institute |
| `2026 Vias institute colours (2).thmx` | Het Office-themabestand dat `Deploy-OfficeTheme.ps1` downloadt |

---

### Deploy-OfficeTheme.ps1

Downloadt `2026 Vias institute colours (2).thmx` uit de branch `main` van deze repo op GitHub en kopieert het naar `%APPDATA%\Microsoft\Templates\Document Themes\`, zodat het in Office verschijnt onder de keuzelijst **Ontwerpen > Thema's** (**Design > Themes**).

```powershell
.\Deploy-OfficeTheme.ps1
```

> Geen parameters — bron-URL en bestandsnaam staan hard gecodeerd bovenaan het script. Rol het via Intune uit als de aangemelde gebruiker (schrijft naar `%APPDATA%`).
