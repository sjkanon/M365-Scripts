[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Intune](../readme.fr.md) › **DiskCleanup**

# DiskCleanup

Déploiement Intune sous forme d'application Win32 qui exécute [`Invoke-WindowsCleanup.ps1`](../../Device/readme.fr.md#invoke-windowscleanupps1) (`scripts/Device/`) sur le lecteur C:\ en tant que SYSTEM, puis redémarre l'appareil. Conçu comme un wrapper minimal afin que toute la logique de nettoyage reste à un seul endroit — rien ici ne la duplique.

## Contenu

| Script | Rôle dans Intune |
|---|---|
| [`Invoke-DiskCleanupIntune.ps1`](Invoke-DiskCleanupIntune.ps1) ([docs](#invoke-diskcleanupintuneps1)) | Script de contenu de la commande d'installation — appelle le script partagé `Invoke-WindowsCleanup.ps1 -Apply`, puis `Restart-Computer -Force` |
| [`Detect-DiskCleanupIntune.ps1`](Detect-DiskCleanupIntune.ps1) ([docs](#detect-diskcleanupintuneps1)) | Script de détection personnalisé |

### Invoke-DiskCleanupIntune.ps1

Commande d'installation de l'application Win32. Exécute `Invoke-WindowsCleanup.ps1 -Apply` depuis le même dossier de contenu, inscrit en cas de succès un horodatage dans `HKLM:\SOFTWARE\DiskCleanupDeploy\LastRunUtc`, puis force un redémarrage.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|-------------|-------------|
| `-SkipDism` | Non | Ignore le nettoyage du magasin de composants DISM (`/StartComponentCleanup /ResetBase`) pour une durée d'exécution plus courte et prévisible |
| `-NoRestart` | Non | Nettoie sans redémarrer — uniquement pour des tests manuels hors Intune ; à omettre dans la vraie commande d'installation |

**Exemples**

```powershell
# Commande d'installation Intune
%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File Invoke-DiskCleanupIntune.ps1

# Test manuel sans DISM et sans redémarrage
.\Invoke-DiskCleanupIntune.ps1 -SkipDism -NoRestart
```

### Detect-DiskCleanupIntune.ps1

Script de détection personnalisé de l'application Win32. Signale « installed » (exit 0) tant que l'horodatage `LastRunUtc` date d'au plus `$MaxAgeDays` (30) jours, et « not installed » (exit 1) dès qu'il est plus ancien ou absent, afin qu'Intune relance le nettoyage. Intune ne transmet aucun paramètre aux scripts de détection : modifiez `$MaxAgeDays` dans le script avant l'empaquetage pour changer le cycle.

## Comportement

- **Redémarrage** : immédiat et forcé (`Restart-Computer -Force`) juste après le nettoyage — sans avertissement ni compte à rebours pour l'utilisateur. Assurez-vous que l'attribution/la notification prévient les utilisateurs finaux à l'avance.
- **Nettoyage du magasin de composants DISM** : inclus par défaut (le plus gros gain d'espace, mais peut prendre des dizaines de minutes). Passez `-SkipDism` dans la commande d'installation si vous avez besoin d'une durée d'exécution courte et prévisible, et augmentez le délai d'installation de l'application Win32 (60 min par défaut) si vous le conservez.
- **Récurrent par conception** : en cas de succès, le script d'installation inscrit un horodatage dans `HKLM:\SOFTWARE\DiskCleanupDeploy\LastRunUtc`. Le script de détection renvoie « not installed » dès que cet horodatage est plus ancien que `$MaxAgeDays` (30 par défaut ; modifiez la constante dans `Detect-DiskCleanupIntune.ps1` avant l'empaquetage pour changer cette valeur), si bien qu'Intune relance de lui-même le nettoyage à chaque cycle — pas besoin d'incrémenter le contenu chaque mois comme pour l'application ClaudeDesktop.
- Journalise dans `%ProgramData%\DiskCleanupDeploy\cleanup.log`, y compris la sortie complète et le rapport CSV de `Invoke-WindowsCleanup.ps1`.

## Empaquetage en application Win32

Pour les deux fichiers d'ici, `Invoke-WindowsCleanup.ps1` doit être copié à côté d'eux avant l'empaquetage — le contenu d'une application Win32 est un dossier plat, le wrapper le retrouve donc via `$PSScriptRoot` au moment de l'installation.

```powershell
Install-Module IntuneWin32App -Scope CurrentUser   # s'il n'est pas déjà installé
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

Attribuez-la en **Required** au groupe d'appareils cible. `-RestartBehavior "suppress"` indique à Intune de ne pas ajouter sa propre invite de redémarrage — le script en force déjà un.

### Prérequis

- Les modules PowerShell `Microsoft.Graph.Authentication` et `IntuneWin32App`
- Exécution depuis Windows (l'empaquetage n'existe que sous Windows)
