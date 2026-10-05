[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../../../readme.nl.md) › [scripts](../../../../readme.nl.md) › [Custom Scripts](../../../readme.nl.md) › [Intune](../../readme.nl.md) › [Desktop](../readme.nl.md) › **Office Themes**

# Office-themakleuren

Rolt een Office-kleurenpalet uit (alleen de themakleuren, niet een volledig `.thmx`-thema — zie daarvoor [`Deploy-OfficeTheme.ps1`](../readme.nl.md#deploy-officethemeps1) in de bovenliggende map).

---

## Bestanden

| Bestand | Omschrijving |
|------|-------------|
| [`Deploy-Officecolors.ps1`](Deploy-Officecolors.ps1) ([docs](#deploy-officecolorsps1)) | Downloadt een kleurenschema-XML van een URL en installeert het in de map Theme Colors van Office |

---

### Deploy-Officecolors.ps1

Downloadt de kleurenschemadefinitie (`<a:clrScheme>`) van `-ColorsUrl` naar `%APPDATA%\Microsoft\Templates\Document Themes\Theme Colors\` en maakt de map aan als die nog niet bestaat. Na de uitrol verschijnt het schema in Office onder de keuzelijst **Ontwerpen > Kleuren** (**Design > Colors**).

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-ColorsUrl` | Directe download-URL van het kleurenschema `.xml`. Verplicht |
| `-ColorsName` | Bestandsnaam om het onder op te slaan, eindigend op `.xml`. Standaard: het laatste deel van de URL |

**Voorbeelden**

```powershell
.\Deploy-Officecolors.ps1 -ColorsUrl 'https://contoso.blob.core.windows.net/branding/Contoso.xml'
```

**Opmerkingen**

- Uitvoeren als de aangemelde gebruiker (schrijft naar `%APPDATA%`). Rol het onder Intune uit als Win32-app met de parameters in de installatieopdracht, of upload een kopie met de standaardwaarden ingevuld.
- Zonder `-ColorsUrl`, of met een naam die niet op `.xml` eindigt, stopt het met exitcode 1.
