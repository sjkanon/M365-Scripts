[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **Deployment**

# Kit d'installation — USB / OOBE

> Auteur : Sjoerd Kanon

Un kit USB pour l'installation de Windows et l'inscription Autopilot. Conçu pour être utilisé pendant l'OOBE (Out-of-Box Experience) via `Shift+F10`.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`start.bat`](start.bat) ([docs](#startbat)) | Menu principal du kit — demande lui-même les privilèges d'administrateur et propose l'inscription Autopilot, Windows Update, le renommage, la jonction au domaine et le navigateur d'installations client |
| [`start.local.example.cmd`](start.local.example.cmd) ([docs](#startlocalcmd)) | Modèle pour `start.local.cmd` — le mot de passe LocalAdmin et le partage d'installation du site, tenus hors du dépôt |
| [`Browse-InstallScripts.ps1`](Browse-InstallScripts.ps1) ([docs](#browse-installscriptsps1)) | Navigateur interactif de clients et de scripts derrière les options de menu `D` et `E` — parcourir les dossiers clients et lancer des fichiers `.ps1` / `.bat` / `.cmd` |

Également dans ce dossier : [`autorun.inf`](autorun.inf) — uniquement le nom de volume de la clé USB ([détails](#autoruninf)).

---

## Arborescence de la clé USB

Tous les fichiers doivent se trouver dans le **même dossier** de la clé USB :

```
USB:\
├── start.bat                      ← Menu principal — à exécuter
├── start.local.cmd                ← Réglages du site : mot de passe LocalAdmin, partage d'installation (hors dépôt)
├── GetAutoPilot.CMD               ← Script d'inscription Autopilot
├── Get-WindowsAutoPilotInfo.ps1   ← Module PowerShell pour le hachage matériel
├── Browse-InstallScripts.ps1       ← Navigateur d'installations client pour les options D/E
├── Install\                        ← Contenu d'installation client local pour l'option D
└── autorun.inf                    ← Nom de volume de la clé USB (purement cosmétique)
```

> `GetAutoPilot.CMD` et `Get-WindowsAutoPilotInfo.ps1` se trouvent dans `scripts/Intune/Get-Autopilot/` — copiez-les tous les deux sur la clé USB.
> Pour l'option de menu `D`, copiez à la fois `Browse-InstallScripts.ps1` et l'intégralité du dossier `Install` au même emplacement que `start.bat` sur la clé USB.

---

## Utilisation pendant l'OOBE

Windows n'exécute **pas** automatiquement les scripts d'une clé USB (bloqué depuis Vista). Étapes manuelles :

1. Branchez la clé USB
2. Pendant l'OOBE, appuyez sur **Shift+F10** pour ouvrir une invite de commandes
3. Repérez la lettre de lecteur de la clé USB — généralement `D:` ou `E:` :
   ```
   D:
   dir
   ```
4. Lancez le kit :
   ```
   start.bat
   ```
5. Le script demande automatiquement les privilèges d'administrateur et affiche le menu

---

## start.bat

Le menu principal du kit. Il se place dans son propre dossier (`cd /d %~dp0`), demande les privilèges d'administrateur et affiche ces options :

| Option | Action | Fonctionne en OOBE |
|---|---|---|
| `1` | Ouvrir le Gestionnaire de périphériques | ✅ |
| `2` | Inscription Autopilot — enregistrer le hachage dans `compHash.csv` | ✅ |
| `3` | Supprimer le fichier de hachage + relancer Autopilot (enregistrement en CSV) | ✅ |
| `4` | **Inscription Autopilot en ligne** — téléverser le hachage directement dans Intune | ✅ (nécessite Internet + un compte administrateur) |
| `5` | Windows Update via `PSWindowsUpdate` (`Install-WindowsUpdate -AcceptAll -AutoReboot`) | ✅ (nécessite Internet) |
| `6` | Installer PowerShell 7 via `winget` | ✅ (nécessite Internet) |
| `7` | Saisir la clé de produit (`slui.exe`) | ✅ |
| `8` | Joindre un domaine Active Directory | ✅ (nécessite la connectivité au domaine) |
| `9` | Redémarrer (délai de 5 secondes) | ✅ |
| `A` | **Tout en un — Intune** — Renommage + Autopilot en ligne + Windows Update + redémarrage | ✅ (nécessite Internet + un compte administrateur) |
| `B` | **Renommer l'appareil** — demande un préfixe, ajoute le numéro de série (`PREFIX-SERIALNUMBER`) | ✅ |
| `C` | **Tout en un — AD** — Renommage + jonction au domaine + Windows Update + redémarrage | ✅ (nécessite la connectivité au domaine) |
| `D` | **Scripts d'installation client (local)** — ouvrir le menu client depuis le dossier local `Install` | ✅ |
| `E` | **Scripts d'installation client (partage réseau)** — ouvrir le menu client depuis le partage de `INSTALL_SHARE` (demandé s'il n'est pas défini) | ✅ (nécessite un accès réseau) |
| `0` | Quitter | ✅ |

### Browse-InstallScripts.ps1

Navigateur d'installations client derrière les options de menu `D` et `E`. Affiche les dossiers clients de premier niveau sous `-RootPath` sous forme de menu, puis permet de les parcourir et de lancer des fichiers `.ps1`, `.bat` et `.cmd` (les dossiers `AppDeployToolkit` sont masqués).

| Paramètre | Obligatoire | Description |
|-----------|-------------|-------------|
| `-RootPath` | Oui | Dossier dont les sous-dossiers sont les clients (dossier local `Install` ou partage réseau) |
| `-SourceLabel` | Non | Titre affiché au-dessus du menu client (par défaut `Install Scripts`) |

```powershell
# Ce que lance l'option D
powershell -NoProfile -ExecutionPolicy Bypass -File .\Browse-InstallScripts.ps1 -RootPath .\Install -SourceLabel "Local Install"
```

- L'option `D` nécessite des fichiers locaux : `Browse-InstallScripts.ps1` et le dossier `Install` complet à côté de `start.bat`.
- L'option `E` lit les dossiers clients depuis le partage de `INSTALL_SHARE` (de [`start.local.cmd`](#startlocalcmd), sinon demandé) et nécessite un accès réseau.
- Avant que l'option `D` ou `E` n'ouvre le navigateur de déploiement, `start.bat` prépare l'appareil au déploiement :
   - Crée ou met à jour l'utilisateur administrateur local `LocalAdmin`
   - Mot de passe : `LOCALADMIN_PASSWORD` de [`start.local.cmd`](#startlocalcmd), sinon demandé en saisie masquée. Sans mot de passe, pas de compte : l'option s'arrête et le menu revient
   - Ajoute `LocalAdmin` au groupe local `Administrators`
   - Définit les indicateurs de registre de saut de l'OOBE afin que le reste du parcours OOBE puisse être ignoré plus facilement

### start.local.cmd

Réglages propres au site qui n'ont pas leur place dans le dépôt. `start.bat` le charge depuis son propre dossier s'il existe ; copiez [`start.local.example.cmd`](start.local.example.cmd) en `start.local.cmd` sur la clé USB et complétez-le. `start.local.cmd` est ignoré par git.

| Variable | Description |
|----------|-------------|
| `LOCALADMIN_PASSWORD` | Mot de passe du compte `LocalAdmin` créé par les options `D` et `E`. Vide ou absent : demandé en saisie masquée. Évitez `%` — batch le remplace |
| `INSTALL_SHARE` | Chemin UNC du partage d'installation client pour l'option `E`, p. ex. `\\server\Software`. Vide ou absent : demandé au choix de l'option `E` |

### Autopilot en ligne (option 4)

Exécute `Get-WindowsAutoPilotInfo.ps1 -Online -DeviceCode` — téléverse le hachage matériel directement dans Intune via Microsoft Graph sans générer de fichier CSV. Connectez-vous avec un compte d'administrateur Microsoft 365 par code d'appareil : ouvrez l'adresse affichée sur un téléphone ou un autre PC et saisissez le code, aucun navigateur n'est nécessaire pendant l'OOBE. L'appareil apparaît dans **Intune → Devices → Enroll devices → Windows enrollment → Autopilot devices** en quelques minutes.

> Le panneau des paramètres de Windows Update n'est pas disponible en OOBE, mais `UsoClient` déclenche les mises à jour directement depuis la ligne de commande et fonctionne très bien.

### PowerShell 7 (option 5)

Utilise `winget install Microsoft.PowerShell`. Nécessite Internet. Si `winget` n'est pas disponible (anciennes versions de Windows 10), le script affiche l'URL de téléchargement manuel. Après l'installation, lancez-le avec `pwsh.exe`.

### Tout en un — Intune (option A)

Pour les environnements gérés par Intune/dans le cloud. Exécute dans l'ordre :
1. Renomme l'appareil — demande un préfixe, ajoute le numéro de série (`PREFIX-SERIALNUMBER`)
2. Supprime le `compHash.csv` existant
3. Lance l'inscription Autopilot en ligne (`Get-WindowsAutoPilotInfo.ps1 -Online -DeviceCode`)
4. Installe les mises à jour Windows via `PSWindowsUpdate`
5. Redémarre après 30 secondes (Ctrl+C pour annuler)

### Tout en un — Active Directory (option C)

Pour les environnements AD on-premise (sans Intune). Exécute dans l'ordre :
1. Renomme l'appareil — demande un préfixe, ajoute le numéro de série (`PREFIX-SERIALNUMBER`)
2. Joint le domaine Active Directory — demande le nom de domaine et des identifiants d'administrateur
3. Installe les mises à jour Windows via `PSWindowsUpdate`
4. Redémarre après 30 secondes (Ctrl+C pour annuler)

---

## autorun.inf

Définit le nom de volume de la clé USB sur `Setup Toolkit` lorsqu'elle est branchée. N'exécute **rien** automatiquement — Windows bloque l'exécution automatique depuis une clé USB sur toutes les versions modernes (Vista+).

---

## Journal des modifications

| Date | Version | Modification |
|---|---|---|
| 2026-10-05 | 3.0 | Le mot de passe LocalAdmin et l'adresse du partage d'installation ne figurent plus dans `start.bat` : ils proviennent de `start.local.cmd` (ignoré par git) et sont demandés s'il est absent — le mot de passe en saisie masquée. Sans mot de passe, les options `D`/`E` s'arrêtent au lieu de créer un compte. Ajout de `start.local.example.cmd` |
| 2026-04-17 | 2.9 | Ajout d'un navigateur d'installations par client dans `start.bat` : l'option `D` ouvre les dossiers clients locaux de `Install` et l'option `E` ouvre un partage réseau ; ajout de `Browse-InstallScripts.ps1` pour parcourir les dossiers clients et exécuter des scripts `.ps1` / `.bat` / `.cmd` ; documentation du fait que l'option `D` exige de copier à la fois `Browse-InstallScripts.ps1` et le dossier `Install` complet ; les options `D` et `E` créent/mettent désormais à jour l'administrateur local `LocalAdmin` (`<mot de passe omis>`) et définissent les indicateurs de saut de l'OOBE avant le début du déploiement |

| Date | Version | Modification |
|---|---|---|
| 2026-10-08 | 2.9 | Autopilot en ligne se connecte par code d'appareil (`-DeviceCode`) : le script dialogue maintenant avec Microsoft Graph au lieu des modules retirés AzureAD/WindowsAutopilotIntune, et une connexion par navigateur peut ne pas s'ouvrir pendant l'OOBE |
| 2026-03-20 | 2.8 | Tout en un scindé en A (Intune) et C (Active Directory) ; la variante AD ignore Autopilot |
| 2026-03-20 | 2.7 | Tout en un mis à jour : jonction au domaine AD ajoutée comme étape 3 |
| 2026-03-20 | 2.6 | Tout en un mis à jour : renommage de l'appareil ajouté comme première étape |
| 2026-03-20 | 2.5 | Ajout de l'option de renommage de l'appareil (B) — demande un préfixe, ajoute le numéro de série (`Get-WmiObject Win32_BIOS`), 15 caractères max. |
| 2026-03-20 | 2.4 | Ajout de l'option de jonction à un domaine Active Directory (8) — `Add-Computer` via PowerShell, demande le domaine + les identifiants |
| 2026-03-20 | 2.3 | Ajout de l'option Autopilot en ligne (4) — `Get-WindowsAutoPilotInfo.ps1 -Online` ; Tout en un utilise désormais l'inscription en ligne |
| 2026-03-20 | 2.2 | Windows Update utilise désormais le module `PSWindowsUpdate` (`Install-WindowsUpdate -AcceptAll -AutoReboot`) au lieu de `UsoClient` |
| 2026-03-20 | 2.1 | Ajout de Windows Update, de l'installation de PowerShell 7 (`winget`) et de l'option Tout en un (`A`) |
| 2026-03-20 | 2.0 | Réécrit en anglais ; `cd /d %~dp0` pour le chemin USB ; auto-élévation ; uniquement des options compatibles OOBE ; chemins entre guillemets ; ajout de `autorun.inf` et du readme |
