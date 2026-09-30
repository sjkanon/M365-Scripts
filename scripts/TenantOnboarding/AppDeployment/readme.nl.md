[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [TenantOnboarding](../readme.nl.md) › **AppDeployment**

# AppDeployment

Generieke, geparametriseerde scripts voor app-deployment aan de devicekant, ter vervanging van een grote familie oude scripts die elk de download-URL van één leverancier of de gegevens van één snelkoppeling/printer hardcoded bevatten. Laat elk script hier naar je eigen pakketrepository / URL wijzen; geen van de scripts bevat een hardcoded intern endpoint.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Install-Win32AppPackage.ps1`](Install-Win32AppPackage.ps1) ([docs](#install-win32apppackageps1)) | Download een gezipt PSAppDeployToolkit-pakket en voer de stille installatie uit |
| [`Install-ChocolateyPackage.ps1`](Install-ChocolateyPackage.ps1) ([docs](#install-chocolateypackageps1)) | Installeer, upgrade of verwijder een pakket via Chocolatey |
| [`Set-DefaultFileAssociation.ps1`](Set-DefaultFileAssociation.ps1) ([docs](#set-defaultfileassociationps1)) | Stel de standaard-app voor een bestandsextensie in (wrapper rond PS-SFTA) |
| [`New-DesktopShortcutsFromStartMenu.ps1`](New-DesktopShortcutsFromStartMenu.ps1) ([docs](#new-desktopshortcutsfromstartmenups1)) | Kopieer een reeks Startmenu-snelkoppelingen naar de Public Desktop |
| [`New-DesktopUrlShortcut.ps1`](New-DesktopUrlShortcut.ps1) ([docs](#new-desktopurlshortcutps1)) | Maak een `.url`-snelkoppeling op de Public Desktop |
| [`Remove-DesktopShortcut.ps1`](Remove-DesktopShortcut.ps1) ([docs](#remove-desktopshortcutps1)) | Verwijder bureaubladsnelkoppelingen die aan een naampatroon voldoen |
| [`Add-NetworkPrinterConnection.ps1`](Add-NetworkPrinterConnection.ps1) ([docs](#add-networkprinterconnectionps1)) | Voeg een netwerkprinter toe via een IP-poort en drivernaam |

---

### Install-Win32AppPackage.ps1

Downloadt een `.zip`-pakket, pakt het uit en voert het PSAppDeployToolkit-startpunt `Deploy-<App>.ps1 -DeploymentType Install -DeployMode NonInteractive` uit, en ruimt daarna op. Registreert optioneel een geplande taak die bij het aanmelden start en bij elke aanmelding opnieuw op updates controleert (voor desktopclients die zichzelf bijwerken).

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-AppName` | Ja | Herkenbare appnaam (werkmap + standaardnaam van het deployscript) |
| `-SourceUri` | Ja | URL naar het `.zip`-pakket |
| `-DeployScriptName` | Nee | Naam van het deployscript in het pakket (standaard: `Deploy-<AppName>.ps1`) |
| `-DeploymentType` | Nee | PSADT-deploymenttype (standaard: `Install`) |
| `-DeployMode` | Nee | PSADT-deploymodus (standaard: `NonInteractive`) |
| `-WorkingRoot` | Nee | Lokale werkmap (standaard: `C:\Install`) |
| `-RegisterLogonUpdateTask` | Nee | Registreer ook een geplande taak bij aanmelden die deze installatie opnieuw uitvoert |
| `-Apply` | Nee | Installeer echt (standaard: alleen voorbeeldweergave) |

**Voorbeelden**
```powershell
.\Install-Win32AppPackage.ps1 -AppName "AdobeReaderDC" -SourceUri "https://packages.contoso.com/Installers/AdobeReaderDC.zip" -Apply

.\Install-Win32AppPackage.ps1 -AppName "LineOfBusinessDesktop" `
    -SourceUri "https://packages.contoso.com/Installers/LineOfBusinessDesktop.zip" -RegisterLogonUpdateTask -Apply
```

**Opmerkingen:** vervangt een hele familie bijna identieke oude scripts (Adobe Reader, AnyDesk, Citrix Workspace, Google Drive, Jabra Direct, TeamViewer, plus een zichzelf bijwerkende line-of-business-desktopclient) die elk de URL van één leverancier hardcoded bevatten: hetzelfde patroon in vier stappen, één script.

---

### Install-ChocolateyPackage.ps1

Installeert Chocolatey als het ontbreekt, en installeert/upgradet of verwijdert daarna een opgegeven pakket.

| Parameter | Omschrijving |
|-----------|-------------|
| `-PackageName` | Chocolatey-pakket-ID (verplicht) |
| `-Uninstall` | Verwijder in plaats van installeren/upgraden |
| `-Apply` | Voer echt uit (standaard: alleen voorbeeldweergave) |

```powershell
.\Install-ChocolateyPackage.ps1 -PackageName git -Apply
.\Install-ChocolateyPackage.ps1 -PackageName git -Uninstall -Apply
```

---

### Set-DefaultFileAssociation.ps1

Stelt een standaard bestandskoppeling in via de communitytool [PS-SFTA](https://github.com/DanysysTeam/PS-SFTA), die de hash berekent die Windows 10 1803+ vereist om wijzigingen in `UserChoice` te laten beklijven.

| Parameter | Omschrijving |
|-----------|-------------|
| `-ProgId` | ProgId of `Applications\<exe>` die standaard moet worden (verplicht) |
| `-Extension` | Een of meer extensies, bijv. `.pdf` (verplicht) |
| `-Apply` | Wijzig echt (standaard: alleen voorbeeldweergave) |

```powershell
.\Set-DefaultFileAssociation.ps1 -ProgId "Applications\7zFM.exe" -Extension ".zip",".rar" -Apply
.\Set-DefaultFileAssociation.ps1 -ProgId "Acrobat.Document.DC" -Extension ".pdf" -Apply
```

> Downloadt PS-SFTA tijdens het uitvoeren van GitHub. Neem het eerst lokaal op als je beleid vooraf goedgekeurde scriptbronnen vereist.

---

### New-DesktopShortcutsFromStartMenu.ps1

Kopieert opgegeven snelkoppelingen uit het Startmenu voor alle gebruikers naar de Public Desktop.

```powershell
.\New-DesktopShortcutsFromStartMenu.ps1 -AppNames "Excel","Word","Outlook" -Apply
```

---

### New-DesktopUrlShortcut.ps1

Maakt een `.url`-bureaubladsnelkoppeling naar een willekeurig webadres.

```powershell
.\New-DesktopUrlShortcut.ps1 -Name "Company Portal" -Url "https://portal.contoso.com/" -Apply
```

---

### Remove-DesktopShortcut.ps1

Verwijdert bureaubladsnelkoppelingen die aan een wildcardpatroon voldoen, van de Public Desktop of van het bureaublad van de huidige gebruiker.

```powershell
.\Remove-DesktopShortcut.ps1 -NamePattern "*.rdp" -Apply
```

---

### Add-NetworkPrinterConnection.ps1

Voegt een TCP/IP-printerpoort + printerverbinding toe met een driver die al geïnstalleerd is.

| Parameter | Omschrijving |
|-----------|-------------|
| `-PrinterName` | Weergavenaam (verplicht) |
| `-PortAddress` | IP-adres/hostnaam van de printer (verplicht) |
| `-DriverName` | Naam van de al geïnstalleerde driver (verplicht) |
| `-PortName` | Naam van het poortobject (standaard: `IP_<PortAddress>`) |
| `-Apply` | Maak echt aan (standaard: alleen voorbeeldweergave) |

```powershell
.\Add-NetworkPrinterConnection.ps1 -PrinterName "Label Printer - Warehouse" -PortAddress "10.0.5.50" -DriverName "Dymo LabelWriter 450 Turbo" -Apply
```

> Installeer eerst de printerdriver: dit script maakt alleen de poort en de verbinding aan.

---

Alle scripts doen standaard een proefdraai; geef `-Apply` mee om wijzigingen door te voeren, volgens de huisstijl.
