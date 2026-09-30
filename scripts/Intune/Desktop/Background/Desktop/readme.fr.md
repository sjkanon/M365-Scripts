[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../../../readme.fr.md) › [scripts](../../../../readme.fr.md) › [Intune](../../../readme.fr.md) › [Desktop](../../readme.fr.md) › [Background](../readme.fr.md) › **Desktop**

# Set-CorporateWallpaper.ps1

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
