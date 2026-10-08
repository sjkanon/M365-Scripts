[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Intune](../readme.nl.md) › **DiskCleanup**

# DiskCleanup

Intune-uitrol als Win32-app die [`Invoke-WindowsCleanup.ps1`](../../Device/readme.nl.md#invoke-windowscleanupps1) (`scripts/Device/`) als SYSTEM op de C:\-schijf uitvoert en het apparaat daarna herstart. Gebouwd als dunne wrapper, zodat alle opruimlogica op één plek blijft — niets hier dupliceert die.

## Inhoud

| Script | Rol in Intune |
|---|---|
| [`Invoke-DiskCleanupIntune.ps1`](Invoke-DiskCleanupIntune.ps1) ([docs](#invoke-diskcleanupintuneps1)) | Contentscript voor de installatieopdracht — roept het gedeelde `Invoke-WindowsCleanup.ps1 -Apply` aan, daarna `Restart-Computer -Force` |
| [`Detect-DiskCleanupIntune.ps1`](Detect-DiskCleanupIntune.ps1) ([docs](#detect-diskcleanupintuneps1)) | Aangepast detectiescript |

### Invoke-DiskCleanupIntune.ps1

Installatieopdracht van de Win32-app. Voert `Invoke-WindowsCleanup.ps1 -Apply` uit vanuit dezelfde contentmap, zet bij succes een tijdstempel in `HKLM:\SOFTWARE\DiskCleanupDeploy\LastRunUtc` en forceert daarna een herstart.

**Parameters**

| Parameter | Verplicht | Beschrijving |
|-----------|-----------|--------------|
| `-SkipDism` | Nee | Slaat het opschonen van het DISM-componentarchief (`/StartComponentCleanup /ResetBase`) over, voor een kortere, voorspelbare looptijd |
| `-NoRestart` | Nee | Ruimt op maar herstart niet — alleen voor handmatig testen buiten Intune; laat het weg in de echte installatieopdracht |

**Voorbeelden**

```powershell
# Intune-installatieopdracht
%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File Invoke-DiskCleanupIntune.ps1

# Handmatige testrun zonder DISM en zonder herstart
.\Invoke-DiskCleanupIntune.ps1 -SkipDism -NoRestart
```

### Detect-DiskCleanupIntune.ps1

Aangepast detectiescript van de Win32-app. Meldt "installed" (exit 0) zolang de tijdstempel `LastRunUtc` hoogstens `$MaxAgeDays` (30) dagen oud is, en "not installed" (exit 1) zodra die ouder is of ontbreekt, zodat Intune het opruimen opnieuw uitvoert. Intune geeft geen parameters mee aan detectiescripts: pas `$MaxAgeDays` in het script aan vóór het verpakken om de cyclus te wijzigen.

## Gedrag

- **Herstart**: direct en geforceerd (`Restart-Computer -Force`) meteen na het opruimen — geen waarschuwing of aftelling voor de gebruiker. Zorg dat de toewijzing/melding eindgebruikers vooraf informeert.
- **Opschonen van het DISM-componentarchief**: standaard inbegrepen (levert de meeste ruimte op, maar kan tientallen minuten duren). Geef `-SkipDism` mee in de installatieopdracht als je een voorspelbaar korte looptijd nodig hebt, en verhoog de installatietime-out van de Win32-app (standaard 60 min) als je het aan laat staan.
- **Bewust terugkerend**: het installatiescript zet bij succes een tijdstempel in `HKLM:\SOFTWARE\DiskCleanupDeploy\LastRunUtc`. Het detectiescript meldt "not installed" zodra die tijdstempel ouder is dan `$MaxAgeDays` (standaard 30; pas de constante in `Detect-DiskCleanupIntune.ps1` aan vóór het verpakken om dit te wijzigen), zodat Intune het opruimen elke cyclus vanzelf opnieuw uitvoert — geen maandelijkse contentverhoging nodig zoals bij de ClaudeDesktop-app.
- Logt naar `%ProgramData%\DiskCleanupDeploy\cleanup.log`, inclusief de volledige uitvoer en het CSV-rapport van `Invoke-WindowsCleanup.ps1`.

## Verpakken als Win32-app

Voor beide bestanden hier moet `Invoke-WindowsCleanup.ps1` vóór het verpakken ernaast worden gekopieerd — de content van een Win32-app is een platte map, dus de wrapper vindt het script tijdens de installatie via `$PSScriptRoot`.

```powershell
Install-Module IntuneWin32App -Scope CurrentUser   # als het nog niet is geïnstalleerd
Connect-MgGraph -Scopes "DeviceManagementApps.ReadWrite.All"

$source = "C:\Temp\DiskCleanupContent"
New-Item -ItemType Directory -Path $source -Force | Out-Null
Copy-Item ".\Invoke-DiskCleanupIntune.ps1" $source
Copy-Item "..\..\Device\Invoke-WindowsCleanup.ps1" $source

$package = New-IntuneWin32AppPackage -SourceFolder $source -SetupFile "Invoke-DiskCleanupIntune.ps1" -OutputFolder "C:\Temp\DiskCleanupOutput"

$detection = New-IntuneWin32AppDetectionRuleScript -ScriptFile ".\Detect-DiskCleanupIntune.ps1"

New-IntuneWin32App -FilePath $package.Path `
    -DisplayName "Disk Cleanup (C:)" `
    -Publisher "IT" `
    -InstallCommandLine "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File Invoke-DiskCleanupIntune.ps1" `
    -UninstallCommandLine "cmd.exe /c echo not applicable" `
    -InstallExperience "system" `
    -RestartBehavior "suppress" `
    -DetectionRule $detection
```

Wijs het als **Required** toe aan de doelgroep van apparaten. `-RestartBehavior "suppress"` zegt tegen Intune dat het er geen eigen herstartprompt bovenop moet zetten — het script forceert er al een.

### Vereisten

- De PowerShell-modules `Microsoft.Graph.Authentication` en `IntuneWin32App`
- Uitvoeren vanaf Windows (verpakken werkt alleen op Windows)
