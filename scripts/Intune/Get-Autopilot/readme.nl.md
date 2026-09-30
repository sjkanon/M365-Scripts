[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Intune](../readme.nl.md) › **Get-Autopilot**

# Get-Autopilot

Verzamelen van Windows Autopilot-hardwarehashes — voor enrollment via USB/OOBE, zie ook [`Deployment/`](../../Deployment/readme.nl.md), dat deze twee bestanden naar zijn USB-toolkit kopieert.

---

## Bestanden

| Bestand | Omschrijving |
|------|-------------|
| [`Get-WindowsAutoPilotInfo.ps1`](Get-WindowsAutoPilotInfo.ps1) ([docs](#get-windowsautopilotinfops1)) | Communityscript (Michael Niehaus) — haalt de Autopilot-hardwarehash op |
| [`GetAutoPilot.CMD`](#getautopilotcmd) | Wrapper om te dubbelklikken — schakelt WinRM in en voert het script uit, met opslag naar `compHash.csv` |

---

### Get-WindowsAutoPilotInfo.ps1

Het bekende communityscript voor het verzamelen van Windows Autopilot-apparaatgegevens (hardwarehash, serienummer, Windows Product ID) en die optioneel rechtstreeks naar Intune te uploaden. Momenteel v3.5 van Michael Niehaus (Microsoft) — zie de [pagina in de PowerShell Gallery](https://www.powershellgallery.com/packages/Get-WindowsAutoPilotInfo) voor de volledige release notes.

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
| `-TenantId` / `-AppId` / `-AppSecret` | App-gebaseerde authenticatie voor de `-Online`-modus |
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

# Uploaden met een groepstag en app-gebaseerde authenticatie
.\Get-WindowsAutoPilotInfo.ps1 -Online -GroupTag "Corporate" -TenantId "..." -AppId "..." -AppSecret "..."
```

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
