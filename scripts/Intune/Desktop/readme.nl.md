[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Intune](../readme.nl.md) › **Desktop**

# Desktop

Desktopaanpassingen die via Intune worden uitgerold: bedrijfsachtergrond + vergrendelscherm, en een snelkoppeling op de taakbalk om het werkstation te vergrendelen.

> De uitrol van Office-thema's en -kleuren (`Deploy-OfficeTheme.ps1`, `Deploy-Officecolors.ps1`) staat in [`Custom Scripts/Intune/Desktop/`](../../Custom%20Scripts/Intune/Desktop/readme.nl.md) — die scripts hebben hun download-URL hard naar dat pad gecodeerd, dus ze blijven daar staan.

---

## Mappen

| Map | Omschrijving |
|--------|-------------|
| [`Background/`](Background/readme.nl.md) | Bedrijfsachtergrond (`Desktop/`) en vergrendelscherm (`Lockscreen/`) |
| [`Add Lockscreen to start and desktop/`](Add%20Lockscreen%20to%20start%20and%20desktop/readme.nl.md) | Maakt een snelkoppeling "Lock Workstation" vast aan Start |
| [`ClaudeDesktop/`](ClaudeDesktop/readme.nl.md) | Machinebrede uitrol van Claude Desktop, één script dat maandelijks draait om actueel te blijven |
| [`CoworkPrerequisites/`](CoworkPrerequisites/readme.nl.md) | Vereisten voor Cowork aan de Windows-kant (`VirtualMachinePlatform`, Snel opstarten) — een eigen, losstaande Win32-app, niet gebundeld met Claude Desktop |

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Deploy-AllIntune.ps1`](Deploy-AllIntune.ps1) ([docs](#deploy-allintuneps1)) | Voert beide bovenstaande uitrolscripts in één aanroep uit |

---

### Deploy-AllIntune.ps1

Een dunne orchestrator zonder eigen Intune-/Graph-logica — voert `CoworkPrerequisites/Deploy-CoworkPrerequisitesIntune.ps1` en daarna `ClaudeDesktop/Deploy-ClaudeDesktopIntune.ps1` uit, en geeft `-AssignmentGroupName`/`-TenantId`/`-Force` aan beide door. Elke uitrol beheert nog steeds zelfstandig zijn eigen Graph-sessie en tijdelijke App Registration — dit bespaart je alleen dat je twee opdrachten met de hand moet uitvoeren.

```powershell
# Beide apps, één opdracht
.\Deploy-AllIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop"

# Onbeheerd (bijv. geplande taak)
.\Deploy-AllIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop" -Force

# Cowork Prerequisites overslaan, alleen Claude Desktop
.\Deploy-AllIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop" -SkipCoworkPrerequisites
```

Stopt voordat Claude Desktop wordt uitgevoerd als de uitrol van Cowork Prerequisites mislukt (geef `-ContinueOnError` mee om het toch uit te voeren).
