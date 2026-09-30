[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../../../readme.fr.md) › [scripts](../../../../readme.fr.md) › [Intune](../../../readme.fr.md) › [Desktop](../../readme.fr.md) › [Background](../readme.fr.md) › **Lockscreen**

# Make-lockscreen.ps1

> Auteur : Sjoerd Kanon

---

## Ce qu'il fait

Télécharge l'image du fond d'écran de l'entreprise depuis Internet et applique ce même fichier comme écran de verrouillage Windows via Intune.

Le script écrit les valeurs de registre `PersonalizationCSP` de l'écran de verrouillage dans `HKLM`, afin que l'écran de verrouillage soit imposé à l'échelle de l'appareil.

---

## Configuration

Modifiez les variables du bloc `CONFIGURATION` en haut du script avant de le charger dans Intune :

```powershell
$ImageUrl   = "https://your-cdn.com/CUSTOMERNAME/wallpaper.png"
$ClientName = "CUSTOMERNAME"
```

Remplacez `CUSTOMERNAME` par le nom du client que vous souhaitez utiliser dans le nom du fichier local et du fichier journal.

Ce script est prévu pour utiliser la même image que `Set-CorporateWallpaper.ps1`, afin que l'identité visuelle du bureau et de l'écran de verrouillage reste cohérente.

Si vous hébergez l'image sur GitHub, une URL brute (raw) est préférable. Les URL courantes de la forme `github.com/.../blob/...` sont normalisées automatiquement.

---

## Journalisation

Les journaux sont écrits dans :

```text
C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateLockscreen-<CUSTOMERNAME>.log
```

---

## Déploiement Intune

### En tant que script PowerShell

1. Intune -> Devices -> Scripts and remediations -> Platform scripts
2. Ajoutez un nouveau script PowerShell pour Windows 10 and later
3. Chargez `Make-lockscreen.ps1`
4. Utilisez ces paramètres :
   - Run this script using the logged on credentials : No
   - Enforce script signature check : No
   - Run script in 64-bit PowerShell : Yes
5. Attribuez au groupe d'appareils requis

### En tant qu'application Win32

Utilisez une règle de détection de fichier pour l'image téléchargée dans :

```text
C:\ProgramData\Wallpapers\corporate-lockscreen-<customername>.<jpg|png|bmp>
```

---

## Fonctionnement

| Étape | Action |
|---|---|
| 1 | Créer `C:\ProgramData\Wallpapers\` s'il n'existe pas |
| 2 | Normaliser les URL de téléchargement GitHub courantes si nécessaire |
| 3 | Télécharger l'image avec `Invoke-WebRequest` |
| 4 | Valider la taille du fichier et détecter un contenu non pris en charge ou HTML |
| 5 | Déterminer le type réel de l'image à partir des en-têtes du fichier |
| 6 | Enregistrer le fichier localement avec la bonne extension |
| 7 | Écrire `LockScreenImagePath`, `LockScreenImageUrl` et `LockScreenImageStatus` dans `PersonalizationCSP` |
| 8 | Déclencher une actualisation des paramètres avec `RUNDLL32.EXE USER32.DLL, UpdatePerUserSystemParameters 1, True` |

---

## Historique des versions

| Date | Version | Modification |
|---|---|---|
| — | 1.0 | Première version de l'écran de verrouillage, avec téléchargement direct via `WebClient` et sortie `.jpg` fixe |
| 2026-04-16 | 2.0 | Mise à jour pour utiliser la même source que le fond d'écran de l'entreprise ; ajout de la normalisation des URL GitHub/raw, de la validation de l'image, de la détection HTML, d'une journalisation structurée et d'une gestion plus sûre du téléchargement temporaire |
