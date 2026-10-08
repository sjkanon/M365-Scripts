[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Intune](../readme.fr.md) › **Desktop**

# Desktop

Personnalisation du poste de travail déployée via Intune : fond d'écran et écran de verrouillage de l'entreprise, ainsi qu'un raccourci de verrouillage du poste dans la barre des tâches.

> Le déploiement des thèmes et couleurs Office (`Deploy-OfficeTheme.ps1`, `Deploy-Officecolors.ps1`) se trouve dans [`Custom Scripts/Intune/Desktop/`](../../Custom%20Scripts/Intune/Desktop/readme.fr.md) — ces scripts ont leur URL de téléchargement codée en dur vers ce chemin, ils restent donc à cet endroit.

---

## Dossiers

| Dossier | Description |
|--------|-------------|
| [`Background/`](Background/readme.fr.md) | Fond d'écran (`Desktop/`) et écran de verrouillage (`Lockscreen/`) de l'entreprise |
| [`Add Lockscreen to start and desktop/`](Add%20Lockscreen%20to%20start%20and%20desktop/readme.fr.md) | Épingle un raccourci « Lock Workstation » au menu Démarrer |
| [`ClaudeDesktop/`](ClaudeDesktop/readme.fr.md) | Déploiement de Claude Desktop à l'échelle de la machine, un seul script exécuté chaque mois pour rester à jour |
| [`CoworkPrerequisites/`](CoworkPrerequisites/readme.fr.md) | Prérequis Cowork côté Windows (`VirtualMachinePlatform`, démarrage rapide) — une application Win32 indépendante, non intégrée à Claude Desktop |

## Scripts

| Script | Description |
|--------|-------------|
| [`Deploy-AllIntune.ps1`](Deploy-AllIntune.ps1) ([docs](#deploy-allintuneps1)) | Exécute les deux scripts de déploiement ci-dessus en un seul appel |

---

### Deploy-AllIntune.ps1

Un orchestrateur minimal sans logique Intune/Graph propre — il exécute `CoworkPrerequisites/Deploy-CoworkPrerequisitesIntune.ps1` puis `ClaudeDesktop/Deploy-ClaudeDesktopIntune.ps1`, en transmettant `-AssignmentGroupName`/`-TenantId`/`-ClientId`/`-CertificateThumbprint`/`-AppOnly`/`-Force` aux deux. Nécessite PowerShell 7, comme les deux scripts de déploiement. Chaque déploiement gère toujours de manière indépendante sa propre session Graph et son App Registration temporaire — cela vous évite seulement de lancer deux commandes à la main.

```powershell
# Les deux applications, une seule commande
.\Deploy-AllIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop"

# Sans surveillance (par ex. tâche planifiée)
.\Deploy-AllIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop" -Force

# Ignorer Cowork Prerequisites, uniquement Claude Desktop
.\Deploy-AllIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop" -SkipCoworkPrerequisites
```

S'arrête avant d'exécuter Claude Desktop si le déploiement de Cowork Prerequisites échoue (passez `-ContinueOnError` pour l'exécuter quand même).
