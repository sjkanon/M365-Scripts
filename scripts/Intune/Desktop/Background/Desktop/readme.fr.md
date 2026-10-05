[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../../../readme.fr.md) › [scripts](../../../../readme.fr.md) › [Intune](../../../readme.fr.md) › [Desktop](../../readme.fr.md) › [Background](../readme.fr.md) › **Desktop**

# Desktop

Fond d'écran d'entreprise du bureau via Intune — le définir, et le retirer.

## Scripts

| Script | Description |
|--------|-------------|
| [`Set-CorporateWallpaper.ps1`](Set-CorporateWallpaper.ps1) ([docs](#set-corporatewallpaperps1)) | Télécharger le fond d'écran d'entreprise et l'imposer à tous les utilisateurs (PersonalizationCSP, utilisateur actuel, Default User) |
| [`Remove-CorporateWallpaper.ps1`](Remove-CorporateWallpaper.ps1) ([docs](#remove-corporatewallpaperps1)) | Annuler le fond d'écran d'entreprise pour chaque utilisateur — uniquement ce qu'a écrit `Set-CorporateWallpaper.ps1`, l'écran de verrouillage d'entreprise reste |

---

## Set-CorporateWallpaper.ps1

> Auteur : Sjoerd Kanon

---

## Ce qu'il fait

Télécharge une image de fond d'écran d'entreprise depuis une URL et l'applique aux appareils Windows via Intune. Définit le fond d'écran pour :

- L'**utilisateur actuel** (WinAPI — appliqué immédiatement)
- L'**utilisateur actuel** (registre HKCU — conserve le paramètre de style)
- **Tous les utilisateurs** via MDM (PersonalizationCSP — imposé via HKLM)
- Les **nouveaux comptes utilisateur** via le profil Default User (NTUSER.DAT)

Le fond d'écran est ainsi appliqué quel que soit l'utilisateur qui se connecte, pour les comptes existants comme pour les comptes nouvellement créés.

---

## Configuration

Modifiez les trois variables du bloc `CONFIGURATION` en haut du script avant de le charger dans Intune :

```powershell
$ImageUrl       = "https://your-cdn.com/CUSTOMERNAME/wallpaper.png"
$ClientName     = "CUSTOMERNAME"
$WallpaperStyle = "10"
```

Remplacez `CUSTOMERNAME` par le nom du client (par ex. `acme`). Tout le reste est fixe et n'a pas besoin d'être modifié.

Si vous hébergez l'image sur GitHub, utilisez une URL de fichier brut (`raw.githubusercontent.com`) plutôt qu'une URL de page `github.com/.../blob/...`.
Le script sait normaliser automatiquement les URL de page GitHub blob/raw courantes, mais le mieux reste d'utiliser directement une URL brute.

### Styles de fond d'écran

| Valeur | Style | Remarques |
|---|---|---|
| `10` | Fill | Recommandé — remplit l'écran sans déformation |
| `6` | Fit | S'ajuste à l'écran, bordures noires possibles |
| `2` | Stretch | Étire pour remplir, peut déformer |
| `0` | Tile | Répète l'image en mosaïque |
| `22` | Span | S'étend sur plusieurs écrans |

---

## Journalisation

Les journaux sont écrits dans le répertoire de journaux Intune standard :

```
C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateWallpaper-<CUSTOMERNAME>.log
```

Lisibles depuis Intune Device Diagnostics ou localement sur l'appareil.

---

## Déploiement Intune

### En tant que script PowerShell

1. Intune → **Devices** → **Scripts and remediations** → **Platform scripts**
2. Cliquez sur **Add** → **Windows 10 and later**
3. Paramètres :
   - **Script** : chargez `Set-CorporateWallpaper.ps1`
   - **Run this script using the logged on credentials** : No
   - **Enforce script signature check** : No
   - **Run script in 64-bit PowerShell** : Yes
4. Attribuez au groupe d'appareils souhaité
5. **Save**

> Remarque : le script s'exécute en tant que SYSTEM. Les étapes 5 (WinAPI) et 6 (HKCU) s'appliquent au contexte SYSTEM, pas à l'utilisateur connecté. PersonalizationCSP (étape 4) et Default User (étape 7) s'appliquent correctement à tous les utilisateurs dans tous les cas.

### En tant qu'application Win32 (recommandé)

L'empaquetage en application Win32 permet de contrôler les réexécutions et d'utiliser des règles de détection.

**Empaqueter le script :**
```powershell
.\IntuneWinAppUtil.exe -c . -s Set-CorporateWallpaper.ps1 -o .
```

**Paramètres de l'application Intune :**

| Champ | Valeur |
|---|---|
| **Install command** | `powershell.exe -ExecutionPolicy Bypass -File Set-CorporateWallpaper.ps1` |
| **Uninstall command** | `cmd.exe /c echo uninstall` |
| **Detection rule** | Le fichier existe : `C:\ProgramData\Wallpapers\corporate-background-<customername>.jpg` (ou `.png` si l'image source est au format PNG) |
| **Run as** | System |
| **Architecture** | 64-bit |

---

## Fonctionnement (étape par étape)

| Étape | Action | Portée |
|---|---|---|
| 1 | Créer `C:\ProgramData\Wallpapers\` s'il n'existe pas | — |
| 2 | Télécharger l'image via `Invoke-WebRequest` | — |
| 3 | Charger l'API Windows `user32.dll` | — |
| 4 | Écrire les clés de registre `PersonalizationCSP` (HKLM) | Tous les utilisateurs via MDM |
| 5 | Appeler `SystemParametersInfo` (WinAPI) | Utilisateur actuel — immédiat |
| 6 | Écrire `HKCU\Control Panel\Desktop` + actualisation via `RUNDLL32` | Utilisateur actuel — persistant |
| 7 | Charger `Default\NTUSER.DAT` et y écrire les mêmes clés | Nouveaux comptes utilisateur |

---

## Journal des modifications

| Date | Version | Modification |
|---|---|---|
| — | 1.0 | Version initiale |
| — | 1.2 | Rendu générique pour être réutilisé par client |
| 2026-03-20 | 2.0 | Traduit en anglais ; `Invoke-WebRequest` remplace `WebClient` ; `#Requires -Version 5.1` ; espace réservé générique pour l'URL du CDN |
| 2026-04-14 | 2.1 | Ajout de la validation de la signature de l'image et d'une extension locale dynamique (`.jpg/.png/.bmp`) pour éviter les fichiers de fond d'écran invalides |
| 2026-04-14 | 2.2 | Ajout de la normalisation des URL GitHub (`github.com/.../blob/...` vers `raw.githubusercontent.com`) et d'une protection contre les réponses HTML pour éviter les fonds noirs/vides |
| 2026-04-14 | 2.3 | Ajout d'une application de secours contre les fonds noirs : clés de stratégie de fond d'écran machine + mise à jour de toutes les ruches utilisateur chargées ; `DesktopImageUrl` utilise désormais l'URL source |
| 2026-04-14 | 2.4 | Correction d'un bogue d'ordre de nettoyage où le fichier temporaire `.download` pouvait être supprimé avant `Move-Item`, provoquant une erreur path-not-found |
| 2026-04-14 | 2.5 | Ajout d'une étape de redémarrage d'`explorer.exe` pour que les changements de fond d'écran/thème soient visibles immédiatement pour les utilisateurs connectés |
| 2026-04-14 | 2.6 | Rétablissement des valeurs de configuration génériques par défaut (`$ImageUrl`, `$ClientName`) pour des déploiements clients réutilisables |
| 2026-04-14 | 2.7 | Ajout d'une sauvegarde de sécurité du fond d'écran actuel et modification de l'ordre de remplacement pour que le fond d'écran précédent reste disponible si la mise à jour échoue |

---

## Remove-CorporateWallpaper.ps1

Annule ce que `Set-CorporateWallpaper.ps1` a mis en place, et uniquement cela. À exécuter en tant que SYSTEM (script de plateforme Intune, ou commande de désinstallation de l'application Win32).

1. **PersonalizationCSP** — supprime les valeurs `DesktopImagePath`, `DesktopImageUrl` et `DesktopImageStatus`. La clé elle-même n'est supprimée que si elle est alors vide : [`Make-lockscreen.ps1`](../Lockscreen/Make-lockscreen.ps1) conserve ses valeurs `LockScreen*` dans la même clé
2. **`HKLM\...\Policies\System`** — supprime `Wallpaper` et `WallpaperStyle`, mais seulement s'ils pointent vers un fond d'écran d'entreprise ; une stratégie définie par autre chose est conservée
3. **Chaque ruche utilisateur chargée** (`S-1-5-21-*`) — un fond d'écran pointant vers un fichier d'entreprise est remis au fond Windows par défaut (`img0.jpg`, style Remplir), et le cache de fond d'écran transcodé de cet utilisateur est vidé pour que l'ancienne image ne subsiste pas
4. **`HKCU` de l'utilisateur courant** — même réinitialisation, mais seulement hors exécution en SYSTEM (en SYSTEM, `HKCU` est le profil de SYSTEM et l'étape 3 a déjà couvert les utilisateurs)
5. **Profil Default User** (`C:\Users\Default\NTUSER.DAT`) — même réinitialisation, pour que les nouveaux comptes ne reçoivent plus le fond d'écran. Ignoré avec un message si la ruche ne peut pas être chargée (non élevé)
6. **Fichiers** — supprime les fichiers `corporate-background-*` de `C:\ProgramData\Wallpapers`. Le dossier n'est supprimé que s'il est vide — l'image de l'écran de verrouillage s'y trouve aussi
7. Redémarre `explorer.exe` pour que le changement soit visible immédiatement

« Un fichier d'entreprise » désigne un fichier `corporate-background-*` dans `C:\ProgramData\Wallpapers` — le nom que lui donne `Set-CorporateWallpaper.ps1`.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-WhatIf` | Afficher chaque valeur et chaque fichier qui serait supprimé ou réinitialisé ; ne rien modifier |

**Exemples**

```powershell
# En tant que SYSTEM (script de plateforme Intune / commande de désinstallation Win32)
powershell.exe -ExecutionPolicy Bypass -File Remove-CorporateWallpaper.ps1

# Voir ce qu'il ferait
.\Remove-CorporateWallpaper.ps1 -WhatIf
```

**Remarques**

- Journal : `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateWallpaper-Remove.log`
- Les utilisateurs dont le profil n'est pas chargé (non connectés) gardent la valeur du fond d'écran dans leur propre ruche ; Windows affiche le fond par défaut dès que le fichier a disparu, et la prochaine exécution de ce script pendant qu'ils sont connectés réinitialise la valeur
- Les versions précédentes supprimaient toute la clé PersonalizationCSP et tout le dossier `C:\ProgramData\Wallpapers`, ce qui supprimait aussi l'écran de verrouillage d'entreprise
