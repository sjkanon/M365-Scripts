[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Intune](../readme.nl.md) › **Get-Autopilot**

# Get-Autopilot

Verzamelen van Windows Autopilot-hardwarehashes — voor enrollment via USB/OOBE, zie ook [`Deployment/`](../../Deployment/readme.nl.md), dat deze twee bestanden naar zijn USB-toolkit kopieert.

---

## Bestanden

| Bestand | Omschrijving |
|------|-------------|
| [`Get-WindowsAutoPilotInfo.ps1`](Get-WindowsAutoPilotInfo.ps1) ([docs](#get-windowsautopilotinfops1)) | Microsoft-script (Michael Niehaus, v3.5) met het online deel herschreven voor Microsoft Graph — haalt de Autopilot-hardwarehash op |
| [`GetAutoPilot.CMD`](GetAutoPilot.CMD) ([docs](#getautopilotcmd)) | Wrapper om te dubbelklikken — schakelt WinRM in en voert het script uit, met opslag naar `compHash.csv` |

---

### Get-WindowsAutoPilotInfo.ps1

Het bekende communityscript voor het verzamelen van Windows Autopilot-apparaatgegevens (hardwarehash, serienummer, Windows Product ID) en die optioneel rechtstreeks naar Intune te uploaden. Gebaseerd op v3.5 van Michael Niehaus (Microsoft, MIT-licentie) — zie de [pagina in de PowerShell Gallery](https://www.powershellgallery.com/packages/Get-WindowsAutoPilotInfo) voor de oorspronkelijke release notes.

Deze kopie is **v3.5.1**: het `-Online`-deel praat rechtstreeks met Microsoft Graph (`Invoke-MgGraphRequest` op `deviceManagement/importedWindowsAutopilotDeviceIdentities`, `windowsAutopilotDeviceIdentities`, `/devices` en `/groups/{id}/members/$ref`). v3.5 had de uitgefaseerde module **AzureAD** (`-AddToGroup`) en **Microsoft.Graph.Intune** (`Connect-MSGraph`) nodig, waardoor de online import niet meer werkte; de nieuwste versie in de gallery (3.9) leunt nog op de module WindowsAutopilotIntune. Nu is alleen `Microsoft.Graph.Authentication` nodig (wordt voor de huidige gebruiker geïnstalleerd als die ontbreekt).

**Aanmelden (`-Online`)** — **standaard gedelegeerd**: je meldt je aan als Intune-beheerder (`-DeviceCode` als er geen browser kan openen, bijv. tijdens OOBE). Scopes: `DeviceManagementServiceConfig.ReadWrite.All`, plus `GroupMember.ReadWrite.All` en `Device.Read.All` met `-AddToGroup`. **App-only** is een optie: `-AppId` met `-CertificateThumbprint` (voorkeur) of `-AppSecret`, en `-TenantId`; de app heeft dezelfde rechten als toepassingsmachtigingen nodig.

Het script blijft bewust op zichzelf staan en werkt met Windows PowerShell 5.1: het wordt naar een USB-stick gekopieerd en gestart door `GetAutoPilot.CMD` / [`Deployment/start.bat`](../../Deployment/readme.nl.md) met `powershell.exe`, dus het laadt `Connect-M365.ps1` uit de repository niet.

**Belangrijkste parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-Name` | Computernaam/-namen om van te verzamelen (standaard: `localhost`); accepteert pipeline-invoer |
| `-OutputFile` | CSV-pad om de hash naartoe te schrijven |
| `-Append` | Toevoegen aan `-OutputFile` in plaats van overschrijven |
| `-Credential` | Referenties om verbinding te maken met externe computers |
| `-Partner` | Gebruik de registratieflow via het CSP Partner Center |
| `-GroupTag` | Autopilot-groepstag om toe te wijzen |
| `-Online` | Upload de hash rechtstreeks naar Intune in plaats van (of naast) het schrijven van een CSV |
| `-TenantId` | Tenant voor `-Online` (verplicht bij app-only; gedelegeerd standaard je aanmeldtenant) |
| `-AppId` / `-CertificateThumbprint` / `-AppSecret` | App-only aanmelden voor de `-Online`-modus (certificaat heeft de voorkeur) |
| `-DeviceCode` | Gedelegeerd aanmelden met een apparaatcode (OOBE) |
| `-AssignedUser` | Wijs in Intune vooraf een gebruiker aan het apparaat toe |
| `-AssignedComputerName` | Wijs vooraf een computernaam toe (`-Online`-modus) |
| `-AddToGroup` | Voeg het apparaat na de import toe aan een Entra ID-groep (`-Online`-modus) |
| `-Assign` | Wacht op de toewijzing van het Autopilot-profiel en toon die (`-Online`-modus) |
| `-Reboot` | Herstart na een geslaagde online import + toewijzing |

**Voorbeelden**

```powershell
# Hash opslaan naar CSV
.\Get-WindowsAutoPilotInfo.ps1 -OutputFile compHash.csv

# Rechtstreeks naar Intune uploaden (interactieve aanmelding)
.\Get-WindowsAutoPilotInfo.ps1 -Online

# Vanuit OOBE: apparaatcode, aan een groep toevoegen, op het profiel wachten, herstarten
.\Get-WindowsAutoPilotInfo.ps1 -Online -DeviceCode -GroupTag "Corporate" -AddToGroup "Autopilot Devices" -Assign -Reboot

# Uploaden met een groepstag en app-only aanmelden (certificaat)
.\Get-WindowsAutoPilotInfo.ps1 -Online -GroupTag "Corporate" -TenantId "..." -AppId "..." -CertificateThumbprint "..."
```

**Opmerkingen**
- De CSV heeft nu precies de kolommen die de import van Intune accepteert (`Device Serial Number`, `Windows Product ID`, `Hardware Hash`, plus `Group Tag` / `Assigned User` als die zijn opgegeven). De eerdere kopie voegde merk/model en een tweede kolom `Hardware Hash` toe, wat `Select-Object` weigert.
- De import- en synchronisatielussen meldden voor elk apparaat het laatste apparaat en konden blijven hangen op een apparaat waarvan de import mislukte; elk apparaat wordt nu apart gecontroleerd.
- Het uitlezen van de hardwarehash vereist een verhoogde sessie.

---

### GetAutoPilot.CMD

Wrapper om te dubbelklikken, voor gebruik tijdens OOBE of door een technicus — geen PowerShell-kennis nodig:

1. Schakelt WinRM in (`Enable-PSRemoting -SkipNetworkProfileCheck -Force`)
2. Voert `Get-WindowsAutoPilotInfo.ps1 -ComputerName $env:computername -OutputFile compHash.csv -Append` uit vanuit dezelfde map
3. Pauzeert, zodat de console open blijft om het resultaat te lezen

```
GetAutoPilot.CMD
```

> Voer beide bestanden uit vanuit dezelfde map — `%~dp0` wordt bepaald ten opzichte van de eigen locatie van het `.CMD`-bestand.
