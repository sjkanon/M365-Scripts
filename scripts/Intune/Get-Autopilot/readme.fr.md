[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Intune](../readme.fr.md) › **Get-Autopilot**

# Get-Autopilot

Collecte des hachages matériels Windows Autopilot — pour l'inscription via USB/OOBE, voir aussi [`Deployment/`](../../Deployment/readme.fr.md), qui copie ces deux fichiers sur sa boîte à outils USB.

---

## Fichiers

| Fichier | Description |
|------|-------------|
| [`Get-WindowsAutoPilotInfo.ps1`](Get-WindowsAutoPilotInfo.ps1) ([docs](#get-windowsautopilotinfops1)) | Script communautaire (Michael Niehaus) — récupère le hachage matériel Autopilot |
| [`GetAutoPilot.CMD`](GetAutoPilot.CMD) ([docs](#getautopilotcmd)) | Wrapper à double-cliquer — active WinRM et exécute le script, avec enregistrement dans `compHash.csv` |

---

### Get-WindowsAutoPilotInfo.ps1

Le script communautaire bien connu pour collecter les informations d'appareil Windows Autopilot (hachage matériel, numéro de série, Windows Product ID) et éventuellement les charger directement dans Intune. Actuellement en v3.5, par Michael Niehaus (Microsoft) — voir la [page PowerShell Gallery](https://www.powershellgallery.com/packages/Get-WindowsAutoPilotInfo) pour les notes de version complètes.

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
| `-TenantId` / `-AppId` / `-AppSecret` | Authentification par application pour le mode `-Online` |
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

# Charger avec une balise de groupe et une authentification par application
.\Get-WindowsAutoPilotInfo.ps1 -Online -GroupTag "Corporate" -TenantId "..." -AppId "..." -AppSecret "..."
```

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
