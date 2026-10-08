[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **Device**

# Scripts de gestion des appareils

Scripts pour gérer et entretenir les postes Windows. Tous les scripts nécessitent les droits administrateur.

---

## Dossiers

| Dossier | Description |
|--------|-------------|
| [`Time sync/`](Time%20sync/readme.fr.md) | Corriger la synchronisation de l'heure Windows en redémarrant W32tm et en enregistrant une tâche planifiée |
| [`audio/`](audio/readme.fr.md) | Détecter et désactiver le microphone interne des ordinateurs portables |
| [`DriveMapping/`](DriveMapping/readme.fr.md) | Mapper des bibliothèques de documents SharePoint/OneDrive sur des lettres de lecteur à l'ouverture de session |
| [`TempDisk/`](TempDisk/readme.fr.md) | Rétablir le disque temporaire éphémère en `D:` à chaque démarrage et y conserver le fichier d'échange |
| [`Printer/`](Printer/readme.fr.md) | Installer des pilotes d'imprimante (téléchargés depuis GitHub) et des imprimantes TCP/IP à partir d'un fichier JSON — conçu pour un serveur tout juste issu d'une golden image |

## Scripts

| Script | Description |
|--------|-------------|
| [`Clear-TempFiles.ps1`](Clear-TempFiles.ps1) ([docs](#clear-tempfilesps1)) | Vider le dossier temporaire partagé des scripts (`C:\Temp` sous Windows, `/tmp` sous Linux/macOS) |
| [`Invoke-WindowsActivation.ps1`](Invoke-WindowsActivation.ps1) ([docs](#invoke-windowsactivationps1)) | Activer Windows, gérer les clés de produit et les paramètres KMS |
| [`Invoke-WindowsCleanup.ps1`](Invoke-WindowsCleanup.ps1) ([docs](#invoke-windowscleanupps1)) | Analyser et libérer l'espace disque récupérable |
| [`Remove-OemBloatware.ps1`](Remove-OemBloatware.ps1) ([docs](#remove-oembloatwareps1)) | Supprimer les logiciels superflus OEM (HP/Lenovo/Dell) et les applications génériques superflues du Microsoft Store |
| [`Repair-AppxPackageStore.ps1`](Repair-AppxPackageStore.ps1) ([docs](#repair-appxpackagestoreps1)) | Réparer les paquets AppX qui échouent avec `0x80070490` — entrées orphelines du magasin de paquets, et FSLogix qui rejoue une version absente de l'hôte (Teams, nouvel Outlook, tout paquet) |
| [`Test-OpenVpnDiagnostics.ps1`](Test-OpenVpnDiagnostics.ps1) ([docs](#test-openvpndiagnosticsps1)) | Diagnostiquer les problèmes d'OpenVPN Connect |
| [`Update-TeamsClient.ps1`](Update-TeamsClient.ps1) ([docs](#update-teamsclientps1)) | Mettre à jour le nouveau Teams + le complément de réunion Outlook, uniquement lorsque Microsoft a publié un build plus récent ([fonctionnement](Update-TeamsClient.md), [IT Glue](Update-TeamsClient-ITGlue.md)) |

---

### Clear-TempFiles.ps1

Vide uniquement le dossier temporaire partagé des scripts.
Le chemin cible par défaut est `C:\Temp` sous Windows et `/tmp` sous Linux/macOS.
S'exécute par défaut en mode essai à blanc et ne supprime des fichiers que si `-Apply` est fourni.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-Apply` | Effectue réellement la suppression (par défaut : essai à blanc) |
| `-TempPath` | Remplace le chemin temporaire cible |
| `-OlderThanDays` | Cible uniquement les éléments de plus de N jours (par défaut : `1`) |

**Exemples**

```powershell
# Essai à blanc
.\Clear-TempFiles.ps1

# Supprimer uniquement les fichiers temporaires de plus de 7 jours
.\Clear-TempFiles.ps1 -Apply -OlderThanDays 7

# Remplacer le chemin temporaire (exemple Linux/macOS)
.\Clear-TempFiles.ps1 -Apply -TempPath /var/tmp
```

---

## Invoke-WindowsActivation.ps1

Activez Windows ou gérez les paramètres de licence depuis la ligne de commande.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-Status` | Affiche l'état d'activation actuel (WMI + `slmgr /dli`) |
| `-ProductKey` | Installe une clé de produit (retail ou clé générique KMS) |
| `-KmsServer` | Définit l'adresse du serveur d'activation KMS |
| `-KmsPort` | Port du serveur KMS (par défaut : 1688) |
| `-Activate` | Déclenche l'activation auprès de Microsoft ou du serveur KMS configuré |
| `-RemoveKey` | Supprime la clé de produit installée (avant une réinstallation / un transfert de licence) |
| `-ReArm` | Réinitialise le compteur de la période de grâce (max. ~3-5 fois par installation de Windows) |
| `-Force` | Ignore les demandes de confirmation |

**Exemples**

```powershell
# Vérifier l'état d'activation actuel
.\Invoke-WindowsActivation.ps1 -Status

# Installer une clé retail et activer en ligne
.\Invoke-WindowsActivation.ps1 -ProductKey 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX' -Activate

# Pointer vers un serveur KMS d'entreprise et activer
.\Invoke-WindowsActivation.ps1 -KmsServer 'kms.company.local' -Activate

# Supprimer la clé avant une réinstallation
.\Invoke-WindowsActivation.ps1 -RemoveKey -Force
```

---

## Invoke-WindowsCleanup.ps1

Analyse l'espace disque récupérable et, en option, le libère. S'exécute par défaut en mode essai à blanc — aucun fichier n'est supprimé sans `-Apply`.

**Ce qui est nettoyé**

| Catégorie | Détails |
|----------|---------|
| Fichiers temporaires | Temp des utilisateurs (tous les profils) + `C:\Windows\Temp` |
| Windows Update | `SoftwareDistribution\Download` + DeliveryOptimization |
| Prefetch | `C:\Windows\Prefetch` |
| Vidages mémoire | Minidumps système + CrashDumps par utilisateur |
| WER | Files d'attente de Windows Error Reporting (système + par utilisateur) |
| Cache | Cache des miniatures, cache des shaders DirectX, cache des polices |
| Corbeille | Tous les lecteurs |
| Cache des navigateurs | Edge, Chrome (multi-profils), Firefox — tous les profils utilisateur |
| Journaux d'événements | Tous les journaux d'événements Windows |
| Journaux applicatifs et système | Analyse dynamique de tout C:\ à la recherche de dossiers `logs`/`log`/`logging` |
| DISM | Nettoyage du magasin de composants (`/StartComponentCleanup /ResetBase`) |

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-Apply` | Effectue réellement le nettoyage (par défaut : essai à blanc uniquement) |
| `-SkipBrowserCache` | Ignore le nettoyage du cache des navigateurs |
| `-SkipAppLogs` | Ignore le nettoyage des journaux applicatifs et système |
| `-SkipEventLogs` | Ignore l'effacement des journaux d'événements Windows |
| `-SkipDism` | Ignore le nettoyage du magasin de composants DISM (lent) |
| `-SkipRecycleBin` | Ignore le vidage de la Corbeille |
| `-OutputPath` | Remplace le dossier de sortie du rapport (par défaut : `C:\Temp\`) |

**Exemples**

```powershell
# Essai à blanc — voir combien d'espace peut être libéré
.\Invoke-WindowsCleanup.ps1

# Nettoyage complet
.\Invoke-WindowsCleanup.ps1 -Apply

# Nettoyage, en ignorant le cache des navigateurs et DISM
.\Invoke-WindowsCleanup.ps1 -Apply -SkipBrowserCache -SkipDism
```

Un rapport CSV avec les résultats par catégorie est enregistré dans `C:\Temp\` après chaque exécution.

---

## Remove-OemBloatware.ps1

Détecte le fabricant de l'appareil et supprime les logiciels superflus OEM connus via `winget`, ainsi qu'une liste générique d'applications grand public du Microsoft Store (Xbox, Solitaire, Bing News/Weather, Cortana, Clipchamp, etc.) via `Remove-AppxPackage`. S'exécute par défaut en mode essai à blanc — aucune application n'est supprimée sans `-Apply`.

> Les listes de logiciels superflus sont un point de départ, pas une liste exhaustive — les ID de paquets varient selon l'image de préinstallation OEM et évoluent avec le temps. Exécutez d'abord `winget list` / `Get-AppxPackage | Select Name` sur un appareil représentatif et adaptez les listes du script si nécessaire.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-Apply` | Supprime réellement les applications trouvées (par défaut : essai à blanc uniquement) |
| `-SkipOem` | Ignore la suppression propre au fabricant, ne traite que la liste générique du Microsoft Store |
| `-SkipAppx` | Ignore la liste générique du Microsoft Store, ne traite que les applications propres au fabricant |
| `-Manufacturer` | Remplace la détection automatique (`HP`, `Lenovo`, `Dell`) |
| `-OutputPath` | Remplace le dossier de sortie du rapport (par défaut : `C:\Temp\`) |

**Exemples**

```powershell
# Essai à blanc — voir ce qui serait supprimé sur cet appareil
.\Remove-OemBloatware.ps1

# Supprimer réellement les logiciels superflus OEM + génériques
.\Remove-OemBloatware.ps1 -Apply

# Supprimer uniquement les applications superflues génériques du Microsoft Store, sans toucher aux applications OEM
.\Remove-OemBloatware.ps1 -Apply -SkipOem
```

Un rapport CSV (trouvé/supprimé par application) est enregistré dans `C:\Temp\` après chaque exécution.

**Nécessite :** `winget` (App Installer du Microsoft Store) pour la suppression des paquets OEM ; droits administrateur.

---

## Repair-AppxPackageStore.ps1

Répare les paquets AppX qui échouent avec `0x80070490` ("Element not found"), généralement journalisé ainsi :

```
MSTeams installation error: Deployment Register operation with target volume C: on Package
MSTeams_26246.1604.5133.838_x64__8wekyb3d8bbwe from:  (AppxManifest.xml)  failed with error 0x80070490.
```

Le chemin vide avant `(AppxManifest.xml)` est l'indice révélateur : Windows rejoue un enregistrement qui n'a pas d'emplacement d'installation. Il y a deux causes, et le script diagnostique les deux :

- **Le magasin de paquets de l'hôte est incohérent.** `AppxAllUserStore` liste un paquet pour un SID sans profil, un paquet dont les fichiers ont disparu, ou une entrée à l'échelle de la machine sans manifeste. `Remove-AppxPackage`, DISM et les programmes d'installation lisent tous ce magasin, donc ils échouent tous, et vider l'hôte (drain) ou le redémarrer ne change rien.
- **FSLogix rejoue une version absente de l'hôte** (source d'événement `Apps (Microsoft-FSLogix-Apps)`, à l'ouverture de session). À la fermeture de session, FSLogix enregistre les paquets de l'utilisateur par nom complet — version exacte — dans `AppxPackages.xml` dans le conteneur de profil, et les réenregistre à l'ouverture de session suivante (`HKLM\SOFTWARE\FSLogix\Profiles\InstallAppxPackages`, activé par défaut). Un hôte avec un autre build, ou sans aucun, répond `0x80070490`. La solution consiste à **provisionner** le paquet pour tous les utilisateurs, au même build sur chaque hôte du pool, et à utiliser un build FSLogix qui enregistre par nom de famille (2210 HF4 pour Teams, 25.06 pour le nouvel Outlook). `AppxPackages.xml` n'est pas modifié : Microsoft indique qu'il n'est pas prévu pour cela, et FSLogix le réécrit à la fermeture de session suivante.

**Étapes**

| Étape | Ce qu'elle fait |
|------|--------------|
| 1. Diagnostic | Paquets enregistrés dont les fichiers ont disparu (*Ghost*) ou dont l'état n'est pas Ok (*Damaged*), paquets provisionnés sans fichiers, entrées `AppxAllUserStore` orphelines, erreurs de déploiement AppX récentes |
| 1b. FSLogix | Build FSLogix, `InstallAppxPackages`, ODFC `IncludeTeams`, les paquets que FSLogix n'a pas pu enregistrer au cours des `-Days` derniers jours comparés à ce que cet hôte provisionne, stratégies d'installation AppX |
| 1c. Applications en échec | **Chaque** paquet qui n'a pas pu s'installer, se mettre à jour ou s'enregistrer au cours des `-Days` derniers jours, d'après le journal de déploiement AppX et le journal FSLogix réunis : nombre, codes d'erreur avec leur signification, versions demandées et présence ou non de leurs fichiers sur cet hôte. `0x80070490` en premier, top 15 |
| 2. Provisionnés | `Remove-AppxProvisionedPackage` pour les copies provisionnées dont les fichiers ont disparu |
| 3. Réenregistrement | `Add-AppxPackage -Register` depuis le manifeste du paquet lui-même lorsque les fichiers sont encore présents. Une version plus ancienne à côté d'une plus récente du même paquet est *Superseded*, pas endommagée — signalée en gris et laissée à Windows pour suppression, car la réenregistrer ne peut qu'échouer avec `0x80073D06` |
| 4. Suppression | `Remove-AppxPackage -AllUsers` pour les fantômes, par utilisateur lorsque cela est refusé |
| 5. Magasin | Chaque clé de registre orpheline restante est exportée dans une sauvegarde `.reg`, puis seulement supprimée — pas de sauvegarde, pas de suppression |
| 6. Provisionnement | Lorsque FSLogix échoue sur un build Teams / Outlook plus récent que celui provisionné par cet hôte : **exactement ce build**, en MSIX depuis le CDN de Microsoft à son URL versionnée, signature vérifiée, puis `Add-AppxProvisionedPackage` et relecture. Sinon `-Provision` : `teamsbootstrapper.exe -p` / Outlook `Setup.exe --provision true --quiet --start-`, téléchargés depuis Microsoft et signature vérifiée. Avec `-UseWinget`, le MSIX de winget à la place (`Microsoft.Teams`, `Microsoft.Outlook`), provisionné avec `Add-AppxProvisionedPackage` avec les éventuelles dépendances apportées par winget. `-WingetId` : la même chose pour tout autre paquet. `-Source` : tout MSIX que vous fournissez |
| 6b. Anciennes builds | `-RemoveOld`, après le provisionnement : chaque build plus ancienne des paquets nommés — copies provisionnées plus anciennes, builds plus anciennes enregistrées pour un utilisateur, et leurs entrées `AppxAllUserStore` (sauvegardées en `.reg` d'abord). Uniquement si la build conservée est provisionnée ; les dossiers `WindowsApps` sont laissés à Windows, la liste du conteneur de profil à FSLogix |
| 7. Vérification | Le diagnostic est relancé ; code de sortie 1 si quelque chose subsiste |

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-Name` | Noms de paquets, caractères génériques autorisés (par défaut `*`). Par ex. `MSTeams,Microsoft.OutlookForWindows`. Raccourcis : `teams`, `outlook`, `copilot` (= `-Copilot`) — `-Name outlook,copilot` suffit |
| `-CheckOnly` | Diagnostic uniquement, aucune modification (code de sortie `2` s'il y a du travail) |
| `-Provision` | Provisionne Teams / le nouvel Outlook pour tous les utilisateurs avec le programme d'installation de Microsoft : ce que FSLogix montre comme manquant ou en retard, plus l'un ou l'autre s'il est nommé explicitement dans `-Name` |
| `-UseWinget` | Avec `-Provision` : prend Teams / le nouvel Outlook depuis winget au lieu du programme d'installation de Microsoft. winget vérifie le SHA256, le script la signature Microsoft. Les manifestes winget sont en retard (mesuré : Teams 26198 contre 26246, Outlook 1.2026.812 contre 902) ; l'exécution avertit lorsque le build est plus ancien que ce que demandent les profils |
| `-WingetId` | ID winget d'autres paquets à provisionner de la même façon pour tous les utilisateurs (uniquement si le manifeste winget correspondant est un MSIX) |
| `-Source` | Provisionne ce `.msix` / `.msixbundle` une fois le magasin nettoyé |
| `-IncludeDeprovisioned` | Efface aussi les marqueurs Deprovisioned avec un `-Name` générique (par défaut, ils ne sont effacés que pour les paquets nommés explicitement) |
| `-SkipSignatureCheck` | N'exige pas de signature Microsoft valide sur le programme d'installation ou sur `-Source` |
| `-Days` | Profondeur de lecture des journaux FSLogix Apps et de déploiement AppX (par défaut `7`) |
| `-WorkingDir` | Dossier de téléchargement des programmes d'installation (par défaut `C:\IT\AppxRepair`) |
| `-LogPath` | Transcription et sauvegardes `.reg` (par défaut `C:\Temp`) |
| `-ComputerName` | S'exécute sur ces hôtes de session au lieu de cette machine (par ex. `lem-avd-4,lem-avd-5,lem-avd-6`) : le script se copie via PowerShell remoting (WinRM) dans `C:\IT\AppxRepair` sur chaque hôte, s'y exécute avec les mêmes paramètres et se termine par un tableau du pool — code de sortie, build FSLogix, Teams / Outlook provisionnés par hôte — qui signale toute différence entre les hôtes. Une réparation est confirmée une seule fois pour tout le pool |
| `-Credential` | Identifiants pour ces sessions distantes |
| `-Copilot` | Examine Copilot (étape 1d) : l'application Microsoft 365 Copilot (`Microsoft.MicrosoftOfficeHub`) et l'application Windows Copilot (`Microsoft.Copilot`), l'**application Microsoft Copilot unifiée** qu'Edge Update installe depuis septembre 2026, et toute stratégie qui la supprime ou la bloque. Leurs marqueurs Deprovisioned sont inclus. Avec `-Provision`, la **nouvelle** application est installée à l'échelle de la machine comme Microsoft le documente : `Install{C50565E9-...}` = 5 (Force Installs), `UpdaterExperimentationAndConfigurationServiceControl` = 1 et `CopilotUnificationAllowed{...}` = 1 sous `HKLM\SOFTWARE\Policies\Microsoft\EdgeUpdate` (après une sauvegarde `.reg`), Edge Update est invité à vérifier immédiatement, et l'exécution attend l'application jusqu'à 10 minutes. Si elle n'apparaît pas, `M365CopilotDesktopInstaller.exe --quiet --start -p` (l'ancienne application, que l'unification migre) sert de solution de repli. Une stratégie qui interdit l'installation (`Install` = 0) n'est jamais outrepassée |
| `-Latest` | Provisionner la build la **plus récente** de Teams / Outlook dans le périmètre, pas seulement celle sur laquelle FSLogix a échoué. Teams : le service de configuration de Microsoft (le flux qu'utilise le client lui-même). Outlook n'a pas de tel flux — le catalogue du Store indiquait 1.2026.818.0 alors que 915.300 était déjà sortie — on utilise donc la build la plus récente dont l'existence est prouvée : la plus récente demandée par FSLogix, enregistrée pour un utilisateur sur l'hôte, ou présente dans `WindowsApps`. À utiliser avec `-Provision` |
| `-RemoveOld` | Après le provisionnement, supprimer toute référence que cet hôte garde à une build **plus ancienne** des paquets nommés : copies provisionnées plus anciennes, builds plus anciennes enregistrées pour un utilisateur, et ce que `AppxAllUserStore` en retient encore (sauvegardé en `.reg` d'abord). Uniquement pour les paquets nommés explicitement (`-Name teams,outlook`), et seulement une fois la build conservée provisionnée — les utilisateurs ne sont jamais privés de l'application. Les dossiers `WindowsApps` sont laissés à Windows ; la liste de chaque conteneur de profil est réécrite par FSLogix à la prochaine déconnexion |

Prend en charge `-WhatIf` et `-Confirm` ; demande confirmation pour chaque modification sauf avec `-Confirm:$false`. Variables de script NinjaOne : `packageName`, `checkOnly`, `provision`, `useWinget`, `wingetId`, `source`, `includeDeprovisioned`, `skipSignatureCheck`, `days`, `workingDir`, `logPath`.

**Exemples**

```powershell
# Qu'est-ce qui est cassé sur cet hôte ? Ne modifie rien.
.\Repair-AppxPackageStore.ps1 -Name MSTeams,Microsoft.OutlookForWindows -CheckOnly

# Réparer et remettre Teams + le nouvel Outlook pour tous les utilisateurs, sans surveillance
.\Repair-AppxPackageStore.ps1 -Name MSTeams,Microsoft.OutlookForWindows -Provision -Confirm:$false

# Tout le pool depuis un seul endroit : diagnostic, puis réparation + provisionnement sur chaque hôte
.\Repair-AppxPackageStore.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name MSTeams,Microsoft.OutlookForWindows -CheckOnly
.\Repair-AppxPackageStore.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name MSTeams,Microsoft.OutlookForWindows -Provision -Confirm:$false

# Nouvel Outlook et nouvelle application Copilot sur le pool : pourquoi ils manquent, puis installation pour tous les utilisateurs
.\Repair-AppxPackageStore.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name outlook,copilot -CheckOnly
.\Repair-AppxPackageStore.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name outlook,copilot -Provision -Confirm:$false

# Les toutes dernières versions de Teams et Outlook sur le pool, pas seulement la build demandée par FSLogix
.\Repair-AppxPackageStore.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name teams,outlook -Latest -Provision -RemoveOld -Confirm:$false

# Idem, avec les deux paquets pris depuis winget
.\Repair-AppxPackageStore.ps1 -Name MSTeams,Microsoft.OutlookForWindows -Provision -UseWinget -Confirm:$false

# Chaque application en échec au cours des 14 derniers jours, plus tout le magasin ; ne modifie rien
.\Repair-AppxPackageStore.ps1 -CheckOnly -Days 14
```

**Remarques**

- Codes de sortie : `0` propre, `1` échec ou quelque chose subsiste, `2` le mode vérification seule a trouvé du travail.
- Sur un pool FSLogix, exécutez-le sur **chaque** hôte — `-ComputerName` le fait depuis un seul endroit et montre si FSLogix, Teams et Outlook sont identiques partout.
- Les profils demandent un Teams / Outlook plus récent que celui provisionné par un hôte, car les deux applications se mettent à jour elles-mêmes par utilisateur, alors que les programmes d'installation de Microsoft provisionnent un build plus ancien, le dernier réputé stable (mesuré : Teams 26225 contre 26246, Outlook 1.2026.818 contre 902 et 915). Lorsque FSLogix **échoue** sur ce build plus récent, `-Provision` installe exactement ce build depuis le CDN de Microsoft. Un hôte de production a montré que c'est nécessaire même lorsque les fichiers sont sur le disque : Outlook a échoué 186× avec `0x80070490` alors que les deux builds demandés étaient présents. Comme les applications continuent à se mettre à jour, exécutez le script de façon planifiée (par ex. tous les jours depuis NinjaOne avec `-Name teams,outlook -Provision -Confirm:$false`) pour garder les hôtes au même niveau ; avec un FSLogix antérieur à 2210 HF4 (Teams) / 25.06 (Outlook), mettez aussi FSLogix à jour.
- Les codes d'erreur sont affichés avec le texte propre à Windows (par ex. `0x80073D19` = "An error occurred because a user was logged off", sans gravité) plus une note lorsque le script en a une.
- Si un paquet continue d'échouer alors que cet hôte provisionne exactement la build demandée, l'exécution ne se termine plus par « Nothing to repair » et le code de sortie 0 : elle indique qu'il ne s'agit pas d'un écart de version, se termine par 1, et l'étape 1c affiche les preuves — la dernière erreur de déploiement AppX avec le *texte d'erreur spécifique* de Windows et son `Get-AppPackageLog -ActivityID`, ainsi que les lignes concernant le paquet dans le journal de profil de FSLogix (`C:\ProgramData\FSLogix\Logs\Profile`).
- Pour un paquet qui continue d'échouer, l'étape 1c vérifie aussi ce qui détermine si cela compte : si chaque utilisateur **connecté** a l'application enregistrée. Si c'est le cas pour tous, les échecs sont la propre relecture de FSLogix — Windows enregistre de toute façon la build provisionnée à l'ouverture de session — et l'exécution le dit et se termine par 0 (`Profiles\InstallAppxPackages = 0` est la méthode documentée par Microsoft pour le faire taire ; le script ne le modifie pas). S'il manque à quelqu'un, cette personne est nommée.
- **Copilot, septembre 2026 :** Microsoft unifie l'application Microsoft 365 Copilot et l'application Windows Copilot en une seule application *Microsoft Copilot*, installée et mise à jour par Edge Update (ID d'application `{C50565E9-CCCF-44B4-BA15-5AC5C6569197}`). Ce qui la bloque est généralement une stratégie, et le script la nomme avec son chemin : `HKLM\SOFTWARE\Policies\Microsoft\EdgeUpdate` — `Install{id}` = 0 (pas d'installation), `Uninstall{id}` = 1/2 (supprimée à chaque vérification, sauf si `Install{id}` = 5 Force Installs, qui prend le dessus) ; `PauseCopilotAppUnificationRollout` sous `HKLM\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate` ; et les stratégies `WindowsCopilot` / `WindowsAI` de Windows lui-même, à l'échelle de la machine et par utilisateur connecté. Elles proviennent de GPO ou d'Intune et ne sont **pas** modifiées par le script — une modification locale serait annulée à la prochaine actualisation. Le centre d'administration des Microsoft 365 Apps peut aussi désactiver l'installation automatique (*Modern Apps settings*) ; cela n'est pas visible sur la machine.
- `-ComputerName` nécessite WinRM depuis l'endroit où vous l'exécutez vers les hôtes (hôtes joints au domaine : Kerberos fonctionne tel quel). Pour les hôtes joints à Entra, ce n'est souvent pas configuré ; exécutez alors le script sur chaque hôte via NinjaOne.
- L'erreur FSLogix d'un utilisateur cesse après que cet utilisateur s'est déconnecté une fois d'un hôte qui provisionne le paquet — FSLogix enregistre alors la version actuelle.
- Après l'étape 5, redémarrez l'hôte quand cela vous convient, pour que le moteur de déploiement relise le magasin.
- Un fichier de sauvegarde `.reg` peut être ouvert par double-clic pour remettre une entrée en place.
- Volontairement non fait : modifier `StateRepository-Machine.srd` ou `AppxPackages.xml`. Ni l'un ni l'autre n'est pris en charge, et le premier casse le menu Démarrer et toutes les applications sur un hôte multisession en cas de problème.
- Pour Teams en particulier, [`Update-TeamsClient.ps1`](Update-TeamsClient.ps1) `-RepairAppxStore` traite la partie magasin dans le cadre d'une mise à jour ; ce script couvre tous les paquets et le côté FSLogix.

---

## Test-OpenVpnDiagnostics.ps1

Collecte et évalue les informations de diagnostic sur les problèmes d'OpenVPN Connect sur une machine Windows. Vérifie chaque couche pertinente, du pilote au réseau, et signale les problèmes détectés.

**Vérifications effectuées**

| Section | Ce qui est vérifié |
|---------|----------------|
| Adaptateurs Wintun / TAP | État du périphérique PnP — signale tout ce qui n'est pas `OK` |
| Adaptateurs réseau virtuels | Visibilité des adaptateurs — avertit s'il n'y en a aucun alors que le VPN devrait être actif |
| Profils réseau | Signale les adaptateurs VPN définis sur `Public` (devrait être `Private`) |
| Logiciels VPN installés | Liste toutes les applications liées au VPN — avertit des conflits possibles avec OpenVPN Connect |
| Hyper-V / WSL / virtualisation | Liste les fonctionnalités activées — avertit si Hyper-V est actif (peut entrer en conflit avec Wintun) |
| Service OpenVPN | État et type de démarrage du service — signale s'il ne tourne pas |
| Routes actives | Routes via l'adaptateur VPN — avertit si l'adaptateur existe mais qu'aucune route n'est présente |
| Configuration DNS | Serveurs DNS par adaptateur actif |
| Journal des événements | Les 20 dernières entrées OpenVPN du journal Application |

Les résultats sont affichés à l'écran, avec à la fin un récapitulatif de tous les problèmes.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-ExportTxt` | Non | Enregistre le rapport complet dans un fichier txt |
| `-OutputPath` | Non | Chemin personnalisé pour le rapport (implique `-ExportTxt`). Par défaut : `C:\Temp\OpenVpnDiagnostics_<timestamp>.txt` |

**Exemples**

```powershell
# Lancer le diagnostic, sortie à l'écran uniquement
.\Test-OpenVpnDiagnostics.ps1

# Lancer et enregistrer le rapport dans C:\Temp\
.\Test-OpenVpnDiagnostics.ps1 -ExportTxt

# Enregistrer dans un chemin personnalisé
.\Test-OpenVpnDiagnostics.ps1 -OutputPath "C:\Support\vpn-report.txt"
```

---

## Update-TeamsClient.ps1

> Référence complète : [Update-TeamsClient.md](Update-TeamsClient.md) — logique de décision, vérification de version, choix de conception et dépannage.
>
> Version service desk pour IT Glue (en néerlandais, par niveau de support) : [Update-TeamsClient-ITGlue.md](Update-TeamsClient-ITGlue.md).

Maintient à jour le nouveau client Teams et le complément de réunion Outlook sur un poste ou un hôte de session AVD. Il demande au service de configuration Teams quel build Microsoft publie pour cette architecture et **n'agit que si ce build est plus récent que celui installé** — un appareil à jour est laissé totalement intact. Lorsqu'une mise à jour est nécessaire, il télécharge `teamsbootstrapper.exe` et le MSIX de ce build publié précisément, en vérifie les deux signatures, désinstalle le complément de réunion, supprime et déprovisionne le paquet AppX `MSTeams`, provisionne ce paquet pour tous les utilisateurs et réinstalle le MSI du complément qui l'accompagne.

« Complément présent » signifie que ses fichiers sont présents, et non qu'une clé de registre le mentionne : un enregistrement à l'échelle de la machine qui pointe vers une DLL de chargement disparue compte comme manquant, et `-CheckOnly` le signale (`the machine-wide add-in registration points at files that are gone`). Se fier uniquement à la clé, c'est ainsi qu'un appareil dont le complément a été supprimé s'entend dire qu'il n'y a rien à faire.

Chaque étape qui modifie l'état passe par `ShouldProcess`, donc `-WhatIf` parcourt tout le déroulement sans toucher à la machine. S'exécute à la main (il s'élève lui-même via l'UAC et demande une seule confirmation) ou sans surveillance depuis un RMM tel que NinjaOne.

**Étapes**

| # | Étape | Respecte `-WhatIf` |
|---|------|-------------------|
| 1 | Contrôle préalable — inventaire : AppX par utilisateur + provisionné, Teams classique (à l'échelle de la machine + par profil), complément, enregistrement Outlook, l'optimisation multimédia AVD et ce que Teams en a journalisé, Teams/Outlook en cours d'exécution | lecture seule |
| 2 | Vérification de version — build publié contre build installé | lecture seule |
| 3 | AVD uniquement (`-AvdOptimizations`) : indicateur `IsWVDEnvironment` + redirecteur WebRTC. Ou (`-RemoveWebRtcRedirector`) : désinstallation de ce redirecteur | oui |
| 4 | Teams classique uniquement (`-RemoveClassicTeams`) : désinstallation du programme d'installation à l'échelle de la machine + des installations par profil | oui |
| 5 | Création du dossier de travail, téléchargement du bootstrapper et du MSIX du build publié (le `buildLink` du service de configuration), vérification des deux signatures Microsoft. Si le paquet échoue, le bootstrapper choisit lui-même | oui |
| 6 | Suppression de l'AppX `MSTeams` pour tous les utilisateurs et déprovisionnement ; un paquet que la pile AppX refuse de supprimer est signalé, sans être bloquant. Le complément n'est pas touché ici | oui |
| 7 | Provisionnement de ce paquet (`teamsbootstrapper.exe -p -o`) ; `-p` seul avec `-UseBootstrapperBuild` ou si aucun paquet n'a pu être récupéré | oui |
| 8 | Remplacement complet du complément, une fois le MSI disponible : désinstallation de celui enregistré (`1612` retenté depuis la copie en cache), vérification que rien ne subsiste, suppression de toute autre copie, installation (`ALLUSERS=1`) | oui |
| 9 | Vérification de l'enregistrement du complément (à l'échelle de la machine + par utilisateur connecté dans Outlook), de la suppression du classique, du paquet provisionné et des composants AVD | signalée comme ignorée sous `-WhatIf` |

Seul ce qui manque est fait : un client à jour avec un complément manquant installe uniquement le complément, et sur un hôte de session avec `-AvdOptimizations`, un redirecteur WebRTC manquant installe uniquement celui-ci.

**L'optimisation multimédia, les deux générations**

`-AvdOptimizations` installe le redirecteur WebRTC, que Microsoft retire le **1er octobre 2026** (fin de disponibilité le 1er avril 2027). Son successeur SlimCore n'est jamais installé sur l'hôte de session — le plug-in intégré à Windows App le déploie sur le *poste* depuis lequel l'utilisateur se connecte — donc la question côté hôte n'est pas « SlimCore est-il ici » mais « mes utilisateurs l'obtiennent-ils ». La réponse se trouve dans le journal d'événements Application : Teams écrit un événement `Microsoft Teams VDI` à chaque connexion, et le contrôle préalable en lit les sept derniers jours sur tout hôte de session et traduit les codes (`24002`/`24010` = sur SlimCore, `16002` = le poste n'a pas de plug-in, `16389` = une stratégie a bloqué le MSIX). `-RemoveWebRtcRedirector` retire l'ancienne génération dès que plus rien ne signale `16002` ; il refuse de s'exécuter avec `-AvdOptimizations`, et laisse `IsWVDEnvironment` en place car SlimCore a aussi besoin de cet indicateur.

> Interroger ce journal nécessite `Get-WinEvent -FilterXPath`, et non `-FilterHashtable` : la forme hashtable lève directement une exception lorsque le fournisseur n'a jamais écrit d'événement, ce qui est le cas normal sur une machine qui n'est pas un hôte de session.

Sur un poste, le contrôle préalable vérifie aussi les trois stratégies qui empêchent le déploiement : `BlockNonAdminUserInstall` (erreur `16389`), `AllowAllTrustedApps` (`15615`) et AppLocker (`10083`). AppLocker est lu plutôt que simplement détecté — seule la collection des applications empaquetées (`Appx`) peut bloquer un MSIX, une collection contenant des règles dont l'application est *non configurée* est appliquée tout de même, et rien n'est appliqué tant que le service Application Identity est arrêté. Le rapport indique le chemin du registre, le mode par collection, l'état du service et les noms des règles, et le signale lorsqu'une règle autorise déjà les paquets. Une stratégie trouvée sur un *hôte de session* est listée à titre indicatif plutôt que signalée comme avertissement : elle n'y bloque rien, mais c'est généralement la même GPO qui atteint aussi les postes.

**Pourquoi l'ordre compte :** rien n'est touché tant qu'un build plus récent n'est pas confirmé, et le programme d'installation est récupéré et vérifié *avant* la première désinstallation — ainsi, un téléchargement raté ou une URL bloquée ne peut jamais laisser l'appareil sans client Teams.

**Vérification de version**

`https://config.teams.microsoft.com/config/v1/MicrosoftTeams/...` est le flux que le client Teams utilise lui-même pour déterminer qu'il est obsolète. Il renvoie le build actuel par architecture (`BuildSettings.WebView2PreAuth.<arch>.latestVersion`). Un build installé égal ou plus récent signifie qu'il n'y a rien à faire. Si le service est injoignable, l'exécution s'arrête au lieu de réinstaller à l'aveugle — `-Force` passe outre. `-Ring` sélectionne un autre anneau de mise à jour (par défaut `general`).

> La désinstallation et la vérification du complément lisent à la fois la ruche de désinstallation 64 bits et celle de `WOW6432Node`, car celle dans laquelle l'entrée atterrit n'est pas fixe (mesuré : 64 bits pour le complément 1.26.21803 sous Windows 11). La version MSI du complément provient de la table des propriétés du MSI (COM `WindowsInstaller.Installer`), et non de `Get-AppLockerFileInformation`, absent de certaines éditions et défaillant sous PowerShell 7.

**Sécurité**

- `msiexec` et le bootstrapper s'exécutent avec un délai d'expiration (`-TimeoutSeconds`, par défaut 900) et sont arrêtés s'ils se bloquent, afin qu'une tâche RMM ne puisse pas bloquer l'agent.
- Le code de sortie MSI `1618` (une autre installation est en cours) est retenté deux fois ; `3010` compte comme une réussite et signale un redémarrage en attente dans le récapitulatif.
- Les erreurs inattendues interrompent l'exécution au lieu de continuer à mi-chemin.
- Une exécution qui modifie réellement quelque chose écrit une transcription dans `C:\Temp\Update-TeamsClient_<timestamp>.log` ; une vérification qui ne trouve rien à faire ne laisse aucun journal superflu.

**Codes de sortie**

| Code | Signification |
|------|---------|
| `0` | Réussite, ou déjà à jour |
| `1` | Échec |
| `2` | Uniquement avec `-CheckOnly` : un build plus récent est disponible |

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-WhatIf` | Affiche ce que ferait une mise à jour sans l'effectuer |
| `-Quiet` | N'affiche rien sauf s'il y a du nouveau : un build plus récent, une action ou un échec |
| `-CheckOnly` | Indique seulement si un build plus récent existe (code de sortie 2), sans rien modifier |
| `-AvdOptimizations` | Hôtes de session AVD/VDI : impose l'indicateur `IsWVDEnvironment` et le redirecteur WebRTC |
| `-RemoveClassicTeams` | Supprime aussi le client Teams classique : programme d'installation à l'échelle de la machine plus installations par profil |
| `-RemoveWebRtcRedirector` | Supprime l'ancienne optimisation multimédia WebRTC, retirée le 1er octobre 2026. Ne peut pas être combiné avec `-AvdOptimizations` ; laisse `IsWVDEnvironment` en place, car SlimCore en a aussi besoin |
| `-ClearOrphanedAddInRegistration` | Dernier recours : faire oublier à Windows Installer un complément de réunion qu'il ne peut plus désinstaller (`1612` avec son MSI en cache disparu), ce qui lui fait refuser toute réinstallation avec `1638` |
| `-RepairAppxStore` | Dernier recours côté AppX : réenregistre un paquet dont les fichiers sont encore présents, puis efface les entrées `AppxAllUserStore` que Windows ne peut plus résoudre — enregistrements pour des SID sans profil, entrée à l'échelle de la machine dont le manifeste a disparu, et le marqueur `Deprovisioned`. Limité à MSTeams ; le contrôle préalable les nomme, que le commutateur soit fourni ou non |
| `-UseWinget` | Prend le MSIX de Teams dans winget au lieu du `buildLink` du service de configuration, et provisionne ce fichier (`teamsbootstrapper.exe -p -o`). winget vérifie en plus le SHA256, mais son manifeste a un ou deux builds de retard |
| `-UseBootstrapperBuild` | Laisse `teamsbootstrapper.exe -p` choisir le build lui-même, comme avant : le déploiement progressif de Microsoft décide, et le résultat peut avoir des semaines de retard sur la vérification de version. Incompatible avec `-UseWinget` |
| `-RepairOutlookAddIn` | Efface un enregistrement Outlook par utilisateur qui pointe vers une DLL de complément qui n'existe plus, afin que celui à l'échelle de la machine reprenne la main |
| `-Confirm:$false` | Ne jamais demander de confirmation (à utiliser pour les exécutions sans surveillance) |
| `-Ring` | Anneau de mise à jour interrogé auprès du service de configuration (par défaut : `general`) |
| `-WorkingDir` | Dossier de téléchargement du bootstrapper et du MSIX de Teams (par défaut : `C:\IT\AVD\Teams`) |
| `-LogPath` | Dossier de transcription (par défaut : `C:\Temp`) |
| `-BootstrapperUrl` | Remplace l'URL de téléchargement de `teamsbootstrapper.exe` (https uniquement) |
| `-WebRtcUrl` | Remplace l'URL du MSI du redirecteur WebRTC (https uniquement) |
| `-SkipMeetingAddIn` | Ne touche pas au complément de réunion, et ne considère pas un complément manquant comme du travail. Défendable sur les postes ordinaires, où le client Teams maintient lui-même le complément à jour par utilisateur |
| `-SkipSignatureCheck` | Accepte un programme d'installation ou un paquet non signé par Microsoft (miroir interne) |
| `-TimeoutSeconds` | Délai d'expiration par processus pour msiexec/bootstrapper (par défaut : `900`) |
| `-Force` | Réinstalle même si Teams est à jour, et continue sans informations sur Teams ou la version |

**Exemples**

```powershell
# Essai à blanc — rechercher un build plus récent et afficher ce que ferait une mise à jour
.\Update-TeamsClient.ps1 -WhatIf

# Mettre à jour uniquement si Microsoft a publié un build plus récent
.\Update-TeamsClient.ps1

# Exécution RMM planifiée : silencieuse sauf en cas de build plus récent ou de problème
.\Update-TeamsClient.ps1 -Quiet -Confirm:$false

# Détection uniquement : code de sortie 2 lorsqu'une mise à jour est disponible
.\Update-TeamsClient.ps1 -CheckOnly -Quiet

# Réparation : réinstallation complète quelle que soit la vérification de version
.\Update-TeamsClient.ps1 -Force -Confirm:$false

# Rapport d'état d'un hôte de session : tout ce qui concerne Teams, ne modifie rien
.\Update-TeamsClient.ps1 -CheckOnly

# Migration terminée : retirer l'optimisation WebRTC retirée
.\Update-TeamsClient.ps1 -RemoveWebRtcRedirector -WhatIf
```

**Exécution depuis NinjaOne**

1. Ajoutez le script (Language: PowerShell, Operating System: Windows, Architecture: **All**, Run As: **System**).
2. Prévisualisez d'abord un appareil : exécutez-le avec `-WhatIf -Confirm:$false` dans le champ *Parameters* — la sortie de la tâche montre la comparaison des versions et chaque étape qu'effectuerait une mise à jour, et l'appareil reste intact.
3. Planifiez l'exécution réelle avec `-Quiet -Confirm:$false`. Sur un appareil à jour, il n'affiche rien et se termine avec `0`, de sorte que le fil d'activité ne montre que les appareils sur lesquels il a réellement fait quelque chose.
4. Pour une tâche de détection/condition, utilisez `-CheckOnly -Quiet` : silencieux et `0` si à jour, sortie et code de sortie `2` lorsqu'un build plus récent est publié.
5. Les variables de script facultatives (cases à cocher `whatIf`, `quiet`, `checkOnly`, `force`, `avdOptimizations`, `removeWebRtcRedirector`, `removeClassicTeams`, `repairOutlookAddIn`, `clearOrphanedAddInRegistration`, `skipMeetingAddIn`, `skipSignatureCheck` ; champs texte `workingDir`, `logPath`, `ring`, `webRtcUrl`, `bootstrapperUrl`) sont lues depuis l'environnement lorsque le paramètre correspondant n'est pas passé, pour qu'un technicien puisse cocher *whatIf* au lieu de taper des paramètres.

Si l'agent démarre PowerShell en 32 bits, le script se relance d'abord en 64 bits via `SysNative` — sans cela, les lectures du registre sont redirigées vers `WOW6432Node` et `$env:ProgramFiles` pointe vers le dossier x86, de sorte que ni le paquet AppX ni le MSI du complément ne sont trouvés.

---

## Time sync/

Corrige les problèmes de synchronisation de l'heure Windows en redémarrant `W32tm` sur les serveurs du pool NTP néerlandais et en enregistrant une tâche planifiée qui relance la synchronisation toutes les 59 minutes. Voir [`Time sync/readme.fr.md`](Time%20sync/readme.fr.md) pour tous les détails.
