[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../../../readme.nl.md) › [scripts](../../../../readme.nl.md) › [Custom Scripts](../../../readme.nl.md) › [Intune](../../readme.nl.md) › [Desktop](../readme.nl.md) › **Office Themes**

# Office-themakleuren

Rolt het Office-kleurenpalet van VIAS Institute uit (alleen de themakleuren, niet het volledige `.thmx`-thema — zie daarvoor [`Deploy-OfficeTheme.ps1`](../readme.nl.md#deploy-officethemeps1) in de bovenliggende map).

---

## Bestanden

| Bestand | Omschrijving |
|------|-------------|
| [`Deploy-Officecolors.ps1`](Deploy-Officecolors.ps1) ([docs](#deploy-officecolorsps1)) | Downloadt de XML met het kleurenschema en installeert die in de map Theme Colors van Office |
| `Test VIAS.xml` | De definitie van het kleurenschema (`<a:clrScheme>`) — donkere/lichte/accentkleuren voor het VIAS Institute-thema |

---

### Deploy-Officecolors.ps1

Downloadt `Test VIAS.xml` uit de branch `main` van deze repo op GitHub en kopieert het naar `%APPDATA%\Microsoft\Templates\Document Themes\Theme Colors\`, waarbij de map zo nodig wordt aangemaakt. Na de uitrol verschijnt "Test VIAS" als selecteerbaar kleurenschema in Office onder de keuzelijst **Ontwerpen > Kleuren** (**Design > Colors**).

```powershell
.\Deploy-Officecolors.ps1
```

> Geen parameters — bron-URL en bestandsnaam staan hard gecodeerd bovenaan het script. Rol het via Intune uit als de aangemelde gebruiker (schrijft naar `%APPDATA%`).
