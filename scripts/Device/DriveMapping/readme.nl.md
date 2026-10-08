[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Device](../readme.nl.md) › **DriveMapping**

# DriveMapping

Koppelt SharePoint Online- / OneDrive-documentbibliotheken via de WebDAV-redirector aan vaste stationsletters — bedoeld als aanmeldscript per gebruiker (Intune Win32-app of een geplande taak bij aanmelden), niet via [`menu.ps1`](../../../menu.ps1).

## Scripts

| Script | Omschrijving |
|--------|--------------|
| [`New-CloudDriveMapping.ps1`](New-CloudDriveMapping.ps1) ([docs](#new-clouddrivemappingps1)) | SharePoint Online- / OneDrive-documentbibliotheken via WebDAV aan stationsletters koppelen — proefdraai tenzij `-Apply` |

---

### New-CloudDriveMapping.ps1

Zet elke `https://`-URL van een documentbibliotheek om naar de WebDAV-UNC-vorm (`\\<host>@SSL\DavWWWRoot\<path>`) en koppelt die met `net use`. Draait standaard als proefdraai — zonder `-Apply` wordt er geen station gekoppeld.

Gaat ervan uit dat de aangemelde gebruiker al een geldige sessie/SSO naar de tenant heeft (op dezelfde manier waarop een browser SharePoint via WebDAV bereikt) — dit is geen app-only Graph-authenticatie. Vereist de Windows-service **WebClient** (WebDAV-redirector); standaard aanwezig op Windows 10/11, op Windows Server is de feature **Desktop Experience** nodig.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-MappingsCsv` | * | CSV met de kolommen `DriveLetter`, `Url`, optioneel `Label` — één rij per koppeling |
| `-DriveLetter` | * | Losse ad-hockoppeling: stationsletter (gebruik met `-Url`) |
| `-Url` | * | Losse ad-hockoppeling: URL van de documentbibliotheek (gebruik met `-DriveLetter`) |
| `-Label` | Nee | Weergavenaam voor de losse ad-hockoppeling |
| `-RemoveExisting` | Nee | Verwijder eerst een bestaande koppeling op de doelletter(s) (`net use <letter>: /delete /y`) |
| `-Persist` | Nee | Maak de koppeling blijvend over herstarts heen (`/persistent:yes`). Standaard uit — een aanmeldscript koppelt normaal bij elke aanmelding opnieuw |
| `-Apply` | Nee | Voer de koppeling echt uit (standaard: proefdraai) |
| `-OutputPath` | Nee | Logmap (standaard: `$env:TEMP` — aanmeldscripts draaien meestal in gebruikerscontext, niet als admin) |

*Ofwel `-MappingsCsv` ofwel `-DriveLetter` + `-Url` is verplicht.

**Voorbeeld van de koppelings-CSV**

```csv
DriveLetter,Url,Label
S,https://contoso.sharepoint.com/sites/Finance/Shared Documents,Finance Docs
O,https://contoso-my.sharepoint.com/personal/j_doe_contoso_com/Documents,My OneDrive
```

**Voorbeelden**

```powershell
# Proefdraai vanuit een CSV met koppelingen
.\New-CloudDriveMapping.ps1 -MappingsCsv .\mappings.csv

# Koppelingen uit de CSV toepassen, en eerst bestaande koppelingen op die letters opruimen
.\New-CloudDriveMapping.ps1 -MappingsCsv .\mappings.csv -RemoveExisting -Apply

# Losse ad-hockoppeling
.\New-CloudDriveMapping.ps1 -DriveLetter Z -Url "https://contoso.sharepoint.com/sites/Finance/Shared Documents" -Apply
```

**Uitrol als aanmeldscript**
- Verpak het als Intune Win32-app (of PowerShell-scriptbeleid) die in **gebruikers**context draait, gestart bij aanmelden, en die `New-CloudDriveMapping.ps1 -MappingsCsv <path> -RemoveExisting -Apply` aanroept
- Lever `mappings.csv` mee naast het script (koppelingslijsten per klant/per groep)

**Opmerkingen**
- Er worden geen referenties opgeslagen of gevraagd — of het koppelen lukt, hangt volledig af van de bestaande tenantsessie van de aangemelde gebruiker
- Na elke run wordt een logbestand geschreven met de opgeloste WebDAV-paden en de status per koppeling
