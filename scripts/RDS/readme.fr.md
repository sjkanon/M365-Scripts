[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **RDS**

# RDS

Scripts de diagnostic, de supervision et de préparation pour l'infrastructure RDP / RD Web Access et les hôtes de session AVD. Exécutez-les directement sur le serveur RDS/RDWeb pour obtenir des résultats complets — les cibles distantes ne bénéficient que de vérifications au niveau de la connectivité.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Test-RDSDiagnostics.ps1`](Test-RDSDiagnostics.ps1) ([docs](#test-rdsdiagnosticsps1)) | Contrôle d'état ponctuel — services, configuration, certificats, compte utilisateur, journaux d'événements |
| [`Watch-RDSLive.ps1`](Watch-RDSLive.ps1) ([docs](#watch-rdsliveps1)) | Surveillance en temps réel des événements de session et de licences |
| [`Get-FSlogix-errors.ps1`](Get-FSlogix-errors.ps1) ([docs](#get-fslogix-errorsps1)) | Diagnostic des profils FSLogix / Azure Files sur un hôte de session AVD |
| [`Invoke-FSLogixShrink.ps1`](Invoke-FSLogixShrink.ps1) ([docs](#invoke-fslogixshrinkps1)) | Réduire les disques de profil FSLogix d'un partage (Invoke-FslShrinkDisk), ou vérifier si FSLogix les compacte lui-même à la déconnexion |
| [`Update-SessionHostImage.ps1`](Update-SessionHostImage.ps1) ([docs](#update-sessionhostimageps1)) | Vérifier et préparer une image Windows 11 multisession ou un hôte de session AVD pour que le nouveau Teams, le nouvel Outlook et Copilot continuent de fonctionner avec FSLogix — FSLogix lui-même n'est pas modifié |

---

### Test-RDSDiagnostics.ps1

Détermine pourquoi les utilisateurs ne peuvent pas se connecter à un serveur RDP ou RD Web Access.

**Vérifications effectuées**

| Domaine | Détails |
|------|---------|
| Serveur RDP | État de `TermService`/`SessionEnv`/`UmRdpService`, RDP activé/désactivé, NLA, limites de session, mode RD Licensing, groupe Remote Desktop Users, règles de pare-feu, sessions actives (`quser`) |
| Serveur RDWeb | Accessibilité du port 443, validité/expiration du certificat HTTPS, IIS + pool d'applications RDWeb (en local uniquement), service RD Gateway (en local uniquement) |
| Compte utilisateur (facultatif, `-Username`) | Activé/verrouillé/expiré, appartenance aux groupes, restrictions de postes d'ouverture de session, dernière connexion, âge du mot de passe |
| Journaux d'événements (facultatif, `-IncludeEventLogs`) | Security 4625 (échec d'ouverture de session RDP), 4740 (verrouillage), `TerminalServices-LocalSessionManager` 20/40 (échec de session/motif de déconnexion) |

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-RdpServer` | Nom d'hôte/IP du serveur RDP/Terminal Server. À omettre pour exécuter les vérifications locales |
| `-RdWebServer` | Nom d'hôte/IP du serveur RD Web Access |
| `-Username` | Vérifier ce qui bloque la connexion d'un compte précis (nécessite le module ActiveDirectory ou le repli ADSI) |
| `-LogPath` | Dossier du fichier journal de sortie (par défaut : `C:\Temp\`) |
| `-IncludeEventLogs` | Inclure l'analyse des journaux d'événements des `-Hours` dernières heures |
| `-Hours` | Nombre d'heures d'historique des journaux d'événements à analyser (par défaut : `24`) |

**Exemples**

```powershell
# Vérification complète — RDP + RDWeb + compte utilisateur
.\Test-RDSDiagnostics.ps1 -RdpServer rdp01.company.local -RdWebServer rdweb.company.local -Username jdoe -IncludeEventLogs

# Exécution locale sur l'hôte RDS, avec vérification des journaux d'événements
.\Test-RDSDiagnostics.ps1 -IncludeEventLogs -Hours 48
```

Les résultats sont écrits dans la console et dans un fichier journal horodaté dans `C:\Temp\`.

---

### Watch-RDSLive.ps1

Interroge les journaux d'événements Windows toutes les N secondes et diffuse les nouveaux événements dans la console + un fichier journal. À exécuter directement sur chaque serveur RDS/RDWeb.

**Événements surveillés**

| Source | Événements |
|--------|--------|
| `TerminalServices-LocalSessionManager` | Ouverture de session (21), reconnexion (22/25), fermeture de session (23), déconnexion (24), échec d'ouverture de session (20), motif de déconnexion (40) — codes de motif lisibles |
| Security | Échec d'ouverture de session RDP (4625, type 10), verrouillage de compte (4740) |
| `TerminalServices-Licensing` | Événements de licence accordée/refusée/avertissement |
| System | Fournisseur `TermServLicensing` (période de grâce, erreurs du serveur de licences) |

Affiche une ligne de pulsation (heartbeat) à chaque interrogation, avec le nombre de sessions actives.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-IntervalSeconds` | Intervalle d'interrogation en secondes (par défaut : `20`) |
| `-LogPath` | Dossier du fichier journal de sortie (par défaut : `C:\Temp\`) |
| `-NoLogFile` | Sortie console uniquement, sans fichier journal |

**Exemples**

```powershell
# Exécuter sur le serveur RDS
.\Watch-RDSLive.ps1

# Interrogation plus rapide, sans fichier journal
.\Watch-RDSLive.ps1 -IntervalSeconds 10 -NoLogFile
```

> Appuyez sur `Ctrl+C` pour arrêter. Exécutez en tant qu'Administrateur pour accéder au journal Security.

---

### Get-FSlogix-errors.ps1

Rassemble, en une seule exécution, tout ce qu'il faut pour comprendre pourquoi un profil
FSLogix ne se monte pas sur un hôte de session AVD — erreurs de montage, VHDX verrouillé,
problème SMB/Azure Files ou erreurs de disque. Lecture seule : il collecte et rend compte, il ne répare rien.

**Ce qu'il collecte**

| Domaine | Détails |
|------|---------|
| Système | Nom d'hôte, build de l'OS, uptime |
| FSLogix | Version installée, la configuration complète `Profiles`/`Containers` et l'état du service |
| Conteneurs | Fichiers VHD(X) attachés, le registre des sessions FSLogix et les chemins de profil issus de `ProfileList` |
| Stockage | Connexions SMB vers Azure Files, et si le partage VHD est seulement joignable |
| Événements | Événements FSLogix des `-Days` derniers jours, comparés aux motifs d'échec critiques connus, plus les erreurs de disque/NTFS et les événements du User Profile Service |
| Restes | Profils locaux sous `C:\Users` et fichiers journaux FSLogix |

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-User` | Produire en plus une section filtrée sur un seul utilisateur — le compte dont le profil échoue |
| `-Days` | Nombre de jours d'historique d'événements à analyser (par défaut : `7`) |
| `-OutputPath` | Dossier de la transcription (par défaut : `%SystemDrive%\Temp\FSLogixDiag`) |

**Exemples**

```powershell
# Tout ce qui concerne la dernière semaine
.\Get-FSlogix-errors.ps1

# Un seul utilisateur, deux semaines en arrière, rapport ailleurs
.\Get-FSlogix-errors.ps1 -User jdoe -Days 14 -OutputPath C:\Temp
```

> Exécutez-le dans une session élevée **sur l'hôte de session lui-même** — les données de
> conteneur, SMB et événements n'existent que là. L'ensemble de l'exécution est écrit dans
> `FSLogixDiag_<host>_<timestamp>.log` dans le dossier de sortie ; c'est ce fichier qu'il
> faut joindre à un ticket.

---

### Invoke-FSLogixShrink.ps1

Récupère l'espace que les conteneurs de profil et ODFC FSLogix conservent après la
suppression de données à l'intérieur : un VHDX dynamique grandit mais ne rétrécit jamais
de lui-même. C'est une enveloppe autour de
[Invoke-FslShrinkDisk](https://github.com/FSLogix/Invoke-FslShrinkDisk), le script de
réduction de l'équipe FSLogix elle-même — toujours le meilleur outil pour cela ; les forks
et alternatives sur GitHub font la même chose avec moins de garanties.

**Ce qu'il fait**

| Étape | Détails |
|-------|---------|
| Téléchargement | Récupère Invoke-FslShrinkDisk à un **commit épinglé** dans `C:\Scripts\Invoke-FslShrinkDisk`, le débloque et vérifie le SHA-256 du script. Une version modifiée en amont n'est jamais exécutée sans être vue ; une copie locale altérée est refusée |
| Rapport | Chaque `.vhd`/`.vhdx` du partage, du plus grand au plus petit, avec dossier, taille et dernière écriture, et le total |
| Réduction | Exécute Invoke-FslShrinkDisk de façon récursive, puis résume son journal CSV : disques réduits, Go récupérés, et les disques qui n'ont pas pu être traités |
| `-CheckHost` | Sur un hôte de session : la **compaction intégrée de FSLogix à la déconnexion** peut-elle s'exécuter ? Version de FSLogix (2210 / 2.9.8361 ou ultérieure), `VHDCompactDisk`, le service Optimiser les lecteurs (`defragsvc` non désactivé) et disques dynamiques |

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-Path` | Le partage contenant les conteneurs, par ex. `\\<storageaccount>.file.core.windows.net\<share>\Profiles`. Parcouru récursivement |
| `-ReportOnly` | Lister uniquement les disques et leur taille ; ne rien réduire |
| `-IgnoreLessThanGB` | Ignorer les disques plus petits que cette valeur (défaut : `5`) |
| `-RatioFreeSpace` | Ne réduire qu'un disque ayant au moins cette fraction d'espace libre à l'intérieur (défaut : `0.1` = 10 %) |
| `-ThrottleLimit` | Disques traités simultanément (défaut : `4` ; au plus deux fois le nombre de cœurs) |
| `-LogFilePath` | Journal CSV (défaut : `C:\Temp\FslShrink_<timestamp>.csv`) ; le dossier est créé s'il manque |
| `-ToolPath` | Emplacement d'Invoke-FslShrinkDisk (défaut : `C:\Scripts\Invoke-FslShrinkDisk`) |
| `-Force` | Télécharger à nouveau Invoke-FslShrinkDisk |
| `-CheckHost` | Vérifier plutôt la compaction propre de FSLogix sur cet hôte ; pas besoin de `-Path` |

**Exemples**

```powershell
# D'abord regarder : chaque conteneur du partage, du plus grand au plus petit
.\Invoke-FSLogixShrink.ps1 -Path \\sa.file.core.windows.net\profiles\Profiles -ReportOnly

# Réduire tout ce qui fait 5 Go ou plus avec au moins 10 % libre à l'intérieur
.\Invoke-FSLogixShrink.ps1 -Path \\sa.file.core.windows.net\profiles\Profiles

# FSLogix compacte-t-il lui-même les disques sur cet hôte ?
.\Invoke-FSLogixShrink.ps1 -CheckHost
```

**Remarques**

- FSLogix 2210 et ultérieur compacte lui-même un conteneur à chaque déconnexion, si le
  disque dépasse 1 Go et qu'au moins 20 % peuvent être gagnés ([Microsoft Learn](https://learn.microsoft.com/en-us/fslogix/concepts-vhd-disk-compaction)).
  Lancez d'abord `-CheckHost` : s'il réussit, une réduction manuelle ne fait que rattraper
  les disques des utilisateurs qui se déconnectent rarement, ou sous le seuil de 20 %.
- Un disque attaché — l'utilisateur est connecté — ne peut pas être réduit ; il apparaît
  dans le résumé comme non traité et l'exécution se termine avec le code 1. Lancez-le en
  dehors des heures de bureau ou avec les hôtes vidés.
- À exécuter en mode élevé (chaque disque est monté), avec accès au partage : sur Azure
  Files via Kerberos ou la clé du compte de stockage. Hyper-V n'est pas nécessaire.
- Pour passer à une version plus récente d'Invoke-FslShrinkDisk : mettez le nouveau commit
  et le SHA-256 de son `Invoke-FslShrinkDisk.ps1` dans `$ToolCommit` / `$ToolHash` en haut
  du script, après avoir lu le diff.

---

### Update-SessionHostImage.ps1

Teams, le nouvel Outlook et Copilot sont des applications MSIX, et sur un hôte AVD mutualisé
avec FSLogix elles cassent toujours de la même façon : l'application d'un utilisateur se met
à jour sur l'hôte A, FSLogix enregistre cette version exacte dans le profil à la
déconnexion, et à la connexion suivante sur l'hôte B — qui n'a pas cette version —
l'inscription échoue avec `0x80070490`. FSLogix 2210 HF4 (Teams) et 25.06 (Outlook)
inscrivent par famille de packages, mais ce script laisse volontairement FSLogix tel quel.
Il fait en sorte que l'image contienne tout ce dont les applications ont besoin, en une
seule build, et empêche les applications de s'en écarter par utilisateur.

**Ce qu'il vérifie**

| Étape | Détails |
|------|---------|
| Windows | Édition (Enterprise multisession), build, redémarrage en attente |
| FSLogix | Build et `InstallAppxPackages` — lecture seule. Sous 25.06, Outlook est réinscrit à sa version exacte enregistrée ; c'est le blocage ci-dessous qui le maintient fonctionnel |
| Blocage des mises à jour | Microsoft Store `AutoDownload = 2` et Teams `disableAutoUpdate = 1`, pour que les applications ne changent qu'avec l'image. Les stratégies Edge Update qui bloquent WebView2 ou Edge sont signalées |
| WebView2 | Le runtime Evergreen utilisé par les trois applications, comparé à la build Edge Stable actuelle (`edgeupdates.microsoft.com`) |
| Applications | Teams, le nouvel Outlook, l'application Microsoft 365 Copilot et l'application Copilot unifiée : build provisionnée, et utilisateurs qui ont une build plus récente que celle de l'image |
| Frameworks | Chaque `PackageDependency` des manifestes de ces applications (VCLibs, UI.Xaml, WindowsAppRuntime, …) doit être présent sur la machine à la `MinVersion` demandée — souvent trop ancien sur une image de 2024 |
| Teams sur AVD | `IsWVDEnvironment`, une build Teams assez récente pour SlimCore (`24193.1805.3040.8975`), le complément Teams Meeting, et le redirecteur WebRTC : plus pris en charge depuis le **1er octobre 2026**, cesse de fonctionner le **1er avril 2027**, conservé uniquement comme solution de repli pour les postes qui ne gèrent pas encore SlimCore |
| Office | Shared Computer Activation (obligatoire en multisession) et le canal de mise à jour |
| Connexion | `Microsoft.AAD.BrokerPlugin` présent, et aucune exclusion de `AppData\Local\Packages` ou des dossiers des applications dans le `redirections.xml` de FSLogix |
| Capture | Avec `-ForCapture` : les packages installés pour un utilisateur mais non provisionnés (Sysprep s'arrête dessus) et un redémarrage en attente font échouer la vérification |

**Ce qu'il corrige** (sans `-CheckOnly`, dans cet ordre) : les deux stratégies de blocage,
Shared Computer Activation, WebView2 (programme d'installation Evergreen Standalone,
signature vérifiée), puis les applications via les scripts existants —
[`Repair-AppxPackageStore.ps1`](../Device/readme.fr.md#repair-appxpackagestoreps1)
`-Name teams,outlook -Latest -Provision -RemoveOld` et `-Name copilot -Provision`, et
[`Update-TeamsClient.ps1`](../Device/readme.fr.md#update-teamsclientps1) `-AvdOptimizations`.
Tout est relu ensuite.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-CheckOnly` | Rapport uniquement, rien n'est modifié. Le code de sortie `2` signifie qu'il reste du travail |
| `-ForCapture` | La machine est la VM d'image sur le point d'être sysprepée : échoue aussi sur un redémarrage en attente et sur les packages par utilisateur non provisionnés |
| `-SkipApps` | Ne pas appeler Repair-AppxPackageStore / Update-TeamsClient ; seulement les stratégies, Shared Computer Activation et WebView2 |
| `-ComputerName` | Hôtes de session sur lesquels s'exécuter via PowerShell remoting ; se termine par un tableau pour tout le pool et nomme chaque colonne qui diffère entre les hôtes |
| `-Credential` | Identifiants pour `-ComputerName` |
| `-WorkingDir` | Dossier des téléchargements (par défaut : `C:\IT\SessionHostImage`) |
| `-LogPath` | Dossier de la transcription d'une exécution qui modifie quelque chose (par défaut : `C:\Temp`) |

**Exemples**

```powershell
# De quoi cet hôte ou cette image a-t-il besoin ? Ne modifie rien.
.\Update-SessionHostImage.ps1 -CheckOnly

# Les trois hôtes de session côte à côte, en lecture seule
.\Update-SessionHostImage.ps1 -ComputerName avd-0,avd-1,avd-2 -CheckOnly

# Préparer la VM d'image, puis vérifier qu'elle est prête pour la capture
.\Update-SessionHostImage.ps1 -Confirm:$false
.\Update-SessionHostImage.ps1 -CheckOnly -ForCapture

# Mettre les trois hôtes au même niveau (les passer d'abord en mode drainage)
.\Update-SessionHostImage.ps1 -ComputerName avd-0,avd-1,avd-2 -Confirm:$false
```

**Remarques**

- À exécuter en mode élevé ou en tant que System. Le script se relance dans Windows
  PowerShell 64 bits, dont les cmdlets AppX ont besoin.
- Avec le blocage actif, Teams et Outlook ne sont mis à jour que par ce script (ou une
  nouvelle image) : exécutez-le chaque mois, sur tous les hôtes en même temps, après la
  mise à jour Windows.
- SlimCore nécessite aussi le côté locataire : la stratégie Teams VDI `VDI2Optimization`
  activée, et Windows App 2.0.352.0 ou plus récent sur les postes. Aucun des deux ne se
  vérifie depuis l'hôte ; `Update-TeamsClient.ps1 -CheckOnly` lit les événements Teams VDI
  de l'hôte de session, qui montrent si les utilisateurs sont réellement sur SlimCore.
- `-ComputerName` copie ce script et les deux qu'il appelle dans
  `C:\IT\SessionHostImage` sur chaque hôte.
