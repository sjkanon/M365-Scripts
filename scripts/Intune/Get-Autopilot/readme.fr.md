[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Intune](../readme.fr.md) › **Get-Autopilot**

# Get-Autopilot

Collecte des hachages matériels Windows Autopilot — pour l'inscription via USB/OOBE, voir aussi [`Deployment/`](../../Deployment/readme.fr.md), qui copie ces deux fichiers sur sa boîte à outils USB.

---

## Fichiers

| Fichier | Description |
|------|-------------|
| [`Get-WindowsAutoPilotInfo.ps1`](Get-WindowsAutoPilotInfo.ps1) ([docs](#get-windowsautopilotinfops1)) | Script Microsoft (Michael Niehaus, v3.5) dont la partie en ligne a été réécrite pour Microsoft Graph — récupère le hachage matériel Autopilot |
| [`GetAutoPilot.CMD`](GetAutoPilot.CMD) ([docs](#getautopilotcmd)) | Wrapper à double-cliquer — active WinRM et exécute le script, avec enregistrement dans `compHash.csv` |

---

### Get-WindowsAutoPilotInfo.ps1

Le script communautaire bien connu pour collecter les informations d'appareil Windows Autopilot (hachage matériel, numéro de série, Windows Product ID) et éventuellement les charger directement dans Intune. Basé sur la v3.5 de Michael Niehaus (Microsoft, licence MIT) — voir la [page PowerShell Gallery](https://www.powershellgallery.com/packages/Get-WindowsAutoPilotInfo) pour les notes de version d'origine.

Cette copie est la **v3.5.1** : la partie `-Online` dialogue directement avec Microsoft Graph (`Invoke-MgGraphRequest` sur `deviceManagement/importedWindowsAutopilotDeviceIdentities`, `windowsAutopilotDeviceIdentities`, `/devices` et `/groups/{id}/members/$ref`). La v3.5 exigeait les modules retirés **AzureAD** (`-AddToGroup`) et **Microsoft.Graph.Intune** (`Connect-MSGraph`), si bien que l'import en ligne ne fonctionnait plus ; la dernière version de la galerie (3.9) dépend encore du module WindowsAutopilotIntune. Seul `Microsoft.Graph.Authentication` est désormais nécessaire (installé pour l'utilisateur courant s'il manque).

**Connexion (`-Online`)** — **déléguée par défaut** : vous vous connectez en tant qu'administrateur Intune (`-DeviceCode` lorsqu'aucun navigateur ne peut s'ouvrir, par ex. en OOBE). Scopes : `DeviceManagementServiceConfig.ReadWrite.All`, plus `GroupMember.ReadWrite.All` et `Device.Read.All` avec `-AddToGroup`. **L'application seule** est une option : `-AppId` avec `-CertificateThumbprint` (recommandé) ou `-AppSecret`, et `-TenantId` ; l'application a besoin des mêmes droits en autorisations d'application.

Le script reste volontairement autonome et compatible Windows PowerShell 5.1 : il est copié sur une clé USB et lancé par `GetAutoPilot.CMD` / [`Deployment/start.bat`](../../Deployment/readme.fr.md) avec `powershell.exe`, il ne charge donc pas le `Connect-M365.ps1` du dépôt.

**Paramètres principaux**

| Paramètre | Description |
|-----------|-------------|
| `-Name` | Nom(s) d'ordinateur à interroger (par défaut : `localhost`) ; accepte l'entrée par pipeline |
| `-OutputFile` | Chemin du CSV dans lequel écrire le hachage |
| `-Append` | Ajouter à `-OutputFile` au lieu de l'écraser |
| `-Credential` | Identifiants pour se connecter aux ordinateurs distants |
| `-Partner` | Utiliser le processus d'enregistrement via le CSP Partner Center |
| `-GroupTag` | Balise de groupe Autopilot à attribuer |
| `-Online` | Charger le hachage directement dans Intune au lieu (ou en plus) d'écrire un CSV |
| `-TenantId` | Tenant pour `-Online` (obligatoire en application seule ; en délégué, par défaut le tenant de connexion) |
| `-AppId` / `-CertificateThumbprint` / `-AppSecret` | Connexion en application seule pour le mode `-Online` (certificat recommandé) |
| `-DeviceCode` | Connexion déléguée avec un code d'appareil (OOBE) |
| `-AssignedUser` | Pré-attribuer un utilisateur à l'appareil dans Intune |
| `-AssignedComputerName` | Pré-attribuer un nom d'ordinateur (mode `-Online`) |
| `-AddToGroup` | Ajouter l'appareil à un groupe Entra ID après l'import (mode `-Online`) |
| `-Assign` | Attendre et afficher l'attribution du profil Autopilot (mode `-Online`) |
| `-Reboot` | Redémarrer après un import en ligne + une attribution réussis |

**Exemples**

```powershell
# Enregistrer le hachage dans un CSV
.\Get-WindowsAutoPilotInfo.ps1 -OutputFile compHash.csv

# Charger directement dans Intune (connexion interactive)
.\Get-WindowsAutoPilotInfo.ps1 -Online

# Depuis l'OOBE : code d'appareil, ajout à un groupe, attente du profil, redémarrage
.\Get-WindowsAutoPilotInfo.ps1 -Online -DeviceCode -GroupTag "Corporate" -AddToGroup "Autopilot Devices" -Assign -Reboot

# Charger avec une balise de groupe et une connexion en application seule (certificat)
.\Get-WindowsAutoPilotInfo.ps1 -Online -GroupTag "Corporate" -TenantId "..." -AppId "..." -CertificateThumbprint "..."
```

**Remarques**
- Le CSV contient désormais exactement les colonnes acceptées par l'import Intune (`Device Serial Number`, `Windows Product ID`, `Hardware Hash`, plus `Group Tag` / `Assigned User` lorsqu'ils sont fournis). L'ancienne copie ajoutait fabricant/modèle et une seconde colonne `Hardware Hash`, que `Select-Object` refuse.
- Les boucles d'import et de synchronisation affichaient le dernier appareil pour chaque appareil et pouvaient se bloquer sur un appareil dont l'import avait échoué ; chaque appareil est désormais vérifié individuellement.
- La lecture du hachage matériel nécessite une session élevée.

---

### GetAutoPilot.CMD

Wrapper à double-cliquer, pour l'OOBE ou les techniciens — aucune connaissance de PowerShell requise :

1. Active WinRM (`Enable-PSRemoting -SkipNetworkProfileCheck -Force`)
2. Exécute `Get-WindowsAutoPilotInfo.ps1 -ComputerName $env:computername -OutputFile compHash.csv -Append` depuis le même dossier
3. Se met en pause pour que la console reste ouverte afin de lire le résultat

```
GetAutoPilot.CMD
```

> Exécutez les deux fichiers depuis le même dossier — `%~dp0` est résolu par rapport à l'emplacement du fichier `.CMD` lui-même.
