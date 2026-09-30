[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Device](../readme.fr.md) › **TempDisk**

# Disque temporaire

Maintient en place le disque temporaire éphémère (`D:`) d'une VM Azure ou d'un hôte de session AVD, et y conserve le fichier d'échange.

Le disque temporaire d'une VM Azure — le disque de ressources, ou le disque NVMe local sur les tailles plus récentes — est effacé chaque fois que la VM est désallouée, redimensionnée ou déplacée vers un autre hôte. Il revient vide, parfois RAW, parfois hors ligne, parfois sans lettre de lecteur. Windows lit la configuration du fichier d'échange au démarrage, donc **un fichier d'échange configuré sur une lettre de lecteur absente au démarrage n'est tout simplement pas créé** : la machine se remet à paginer sur `C:`, ou tourne sans aucun fichier d'échange. C'est ce que ces deux scripts doivent empêcher.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Init-TempDisk.ps1`](Init-TempDisk.ps1) ([docs](#init-tempdiskps1)) | Rétablit le disque temporaire en `D:` et y configure le fichier d'échange |
| [`Register-InitTempDiskTask.ps1`](Register-InitTempDiskTask.ps1) ([docs](#register-inittempdisktaskps1)) | Installe ce script sur l'appareil et l'exécute à chaque démarrage en tant que SYSTEM |

---

### Init-TempDisk.ps1

À chaque exécution :

1. **Contrôle préalable** — chaque disque, ce qui occupe `D:`, le fichier d'échange tel que configuré (registre) et le fichier d'échange réellement utilisé dans cette session
2. **Lettre** — un lecteur optique qui occupe `D:` est déplacé ; sur une image sans disque temporaire, Windows attribue `D:` au lecteur DVD et ne la rend jamais
3. **Disque** — un volume qui *est* déjà le disque temporaire récupère sa lettre de lecteur ; sinon, un disque RAW qui n'est ni de démarrage ni système est mis en ligne, initialisé en GPT, partitionné et formaté en NTFS
4. **Fichier d'échange** — gestion automatique désactivée, fichier d'échange pointé vers `D:\pagefile.sys`, entrée de tout autre lecteur supprimée
5. **Vérification** — tout est relu, en indiquant ce qui est effectif maintenant et ce qui attend le prochain redémarrage
6. **Redémarrage** — uniquement avec `-RestartIfNeeded` : redémarre la machine lorsque c'est la seule chose qui sépare encore la configuration d'un fichier d'échange réellement utilisé

Un disque temporaire qui a seulement perdu sa lettre de lecteur la récupère et n'est jamais reformaté — le volume est reconnu à son nom de volume (`Temporary Storage`) ou au fichier `DataLoss_Warning_Readme.txt` qu'Azure écrit sur le disque de ressources.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-DriveLetter` | Lettre de lecteur du disque temporaire (par défaut : `D`) |
| `-Label` | Nom de volume écrit lors du formatage, et nom auquel un volume temporaire existant est reconnu (par défaut : `Temporary Storage`) |
| `-DiskNumber` | Formate ce disque au lieu de laisser le script en choisir un — obligatoire lorsque plusieurs disques RAW sont présents |
| `-InitialSizeMB` / `-MaximumSizeMB` | Taille du fichier d'échange en Mo. `0` (la valeur par défaut) sur les deux signifie géré par le système |
| `-KeepSystemDrivePagefile` | Laisse en place un fichier d'échange existant sur `C:` au lieu de le supprimer |
| `-SkipPagefile` | Rétablit uniquement le disque ; ne touche pas à la configuration du fichier d'échange |
| `-OpticalDriveLetter` | Lettre vers laquelle un lecteur optique est déplacé lorsqu'il occupe `D:` (par défaut : `Z`) |
| `-RestartIfNeeded` | Redémarre la machine lorsque c'est la seule chose qui sépare encore la configuration d'un fichier d'échange utilisé. Désactivé par défaut |
| `-RestartDelaySeconds` | Compte à rebours avant un redémarrage alors que quelqu'un est connecté (par défaut : `60`) — `shutdown /a` l'annule. Si personne n'est connecté, le redémarrage a lieu en quelques secondes |
| `-RestartCooldownMinutes` | Intervalle minimal entre deux redémarrages déclenchés par ce script (par défaut : `60`) |
| `-RestartEvenIfUsersSignedIn` | Redémarre même si quelqu'un est connecté. Sur un hôte de session, passez-le plutôt en mode drain |
| `-CheckOnly` | Rapport uniquement, aucune modification. Le code de sortie `2` signifie qu'il y a du travail à faire |
| `-Quiet` | N'affiche rien sauf s'il y a du nouveau — le fichier journal reçoit toujours l'historique complet |
| `-LogPath` | Dossier de `Init-TempDisk.log` (par défaut : `C:\Temp`), complété et soumis à rotation au-delà de 1 Mo |
| `-Force` | Avec `-DiskNumber` : formate ce disque même s'il comporte encore des partitions |

**Exemples**

```powershell
# Rapport d'état en lecture seule : où se trouve le disque temporaire et ce que fait réellement le fichier d'échange
.\Init-TempDisk.ps1 -CheckOnly

# Parcourir tout le déroulement sans toucher à la machine
.\Init-TempDisk.ps1 -WhatIf

# Comme la tâche planifiée l'exécute — silencieux tant que tout est en ordre,
# et un redémarrage lorsque le fichier d'échange l'attend
.\Init-TempDisk.ps1 -Quiet -RestartIfNeeded

# Fichier d'échange fixe de 16 Go au lieu d'un fichier géré par le système
.\Init-TempDisk.ps1 -InitialSizeMB 16384 -MaximumSizeMB 16384

# Indiquer quel disque est le disque temporaire, même s'il comporte encore des partitions
.\Init-TempDisk.ps1 -DiskNumber 2 -Force
```

**Codes de sortie**

| Code | Signification |
|------|---------|
| `0` | Le disque temporaire et le fichier d'échange sont conformes |
| `1` | Échec |
| `2` | Uniquement avec `-CheckOnly` : il y a du travail à faire |

**Remarques**
- **Rien de ce qui n'est pas RAW n'est jamais initialisé.** Un disque temporaire vide et un disque de données non formaté se ressemblent vus de l'extérieur, donc un disque qui comporte déjà des partitions est signalé et laissé tel quel. Avec plus d'un candidat RAW, le script refuse de deviner et demande `-DiskNumber` ; `-Force` avec `-DiskNumber` est le seul moyen de formater un disque qui comporte encore des partitions
- Un disque NVMe local l'emporte sur les autres disques RAW lorsqu'il y en a plusieurs — sur les tailles de VM plus récentes, c'*est* le disque temporaire
- **Un fichier d'échange configuré par une exécution apparaît au prochain redémarrage.** Windows lit la configuration au démarrage et ne la relit jamais, donc l'exécution le signale au lieu de prétendre avoir réussi. Comme la tâche s'exécute à chaque démarrage, l'appareil se répare de lui-même même sans `-RestartIfNeeded` : le démarrage qui recrée `D:` configure le fichier d'échange, le démarrage suivant le met en service
- `-RestartIfNeeded` comble cet écart au lieu de l'attendre, et un script qui s'exécute à chaque démarrage et peut redémarrer la machine est une boucle de redémarrage en puissance — il ne se déclenche donc que si **toutes** ces conditions sont remplies : l'exécution s'est terminée sans erreur, le disque est présent, le fichier d'échange y est configuré et seule cette session ne l'utilise pas ; personne n'est connecté (session connectée *ou* déconnectée — un `explorer.exe` par bureau, ce qui est indépendant de la langue, alors qu'analyser `query.exe` ne l'est pas) ; et aucun redémarrage n'a été déclenché pendant les dernières `-RestartCooldownMinutes`, mémorisé sous `HKLM:\SOFTWARE\ICTKanon\InitTempDisk`. Une exécution en échec ne redémarre jamais — cela masquerait l'échec derrière un redémarrage
- Le redémarrage passe par `shutdown.exe` avec le motif planifié "Operating System: Reconfiguration", afin d'apparaître comme prévu et non comme inattendu. Le compte à rebours sert à avertir des personnes, il ne s'applique donc que lorsqu'il y en a : si quelqu'un est connecté, c'est `-RestartDelaySeconds` et `shutdown /a` l'arrête ; si personne n'est connecté — le cas normal au démarrage, et le cas garanti sur un hôte de session dont le pool est en mode drain — le redémarrage a lieu en quelques secondes, car attendre une minute devant une salle vide ne fait que coûter de la disponibilité
- Configurer un fichier d'échange sur un lecteur absent écrirait un paramètre que Windows ignore, donc l'exécution échoue plutôt si `D:` n'a pas pu être rétabli
- Nécessite les droits administrateur ; lancé à la main depuis une fenêtre ordinaire, il demande lui-même l'élévation

---

### Register-InitTempDiskTask.ps1

À exécuter une fois par appareil — à la main, depuis Tactical RMM / NinjaOne, ou comme script de plateforme Intune. Copie `Init-TempDisk.ps1` dans un dossier local (le dépôt n'est pas disponible au démarrage) et enregistre une tâche qui l'exécute à chaque démarrage, en tant que SYSTEM, avec les privilèges les plus élevés et sans que quiconque ait à se connecter.

Le relancer est sans risque : une tâche portant le même nom est remplacée, c'est donc aussi ainsi que l'on modifie les arguments avec lesquels la tâche s'exécute.

Par défaut, la tâche exécute `Init-TempDisk.ps1 -Quiet -RestartIfNeeded`. Windows lit la configuration du fichier d'échange au démarrage, donc un démarrage au cours duquel le disque temporaire a dû être reconstruit passe toute cette session sans fichier d'échange sur `D:`, à moins que la machine ne redémarre une fois — et ce sont les garde-fous ci-dessus qui permettent de confier cela sans risque à une tâche de démarrage. Passez `-ScriptArguments '-Quiet'` pour exclure le redémarrage.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-ScriptSourcePath` | `Init-TempDisk.ps1` à installer (par défaut : la copie située à côté de ce script) |
| `-ScriptTargetDir` | Dossier de l'appareil dans lequel il est copié (par défaut : `C:\Scripts`) |
| `-TaskName` | Nom de la tâche planifiée (par défaut : `InitTempDisk`) |
| `-ScriptArguments` | Arguments pour `Init-TempDisk.ps1` (par défaut : `-Quiet -RestartIfNeeded`) |
| `-DelaySeconds` | Délai entre le démarrage et le lancement de la tâche (par défaut : `30`) |
| `-RunNow` | Lance aussi la tâche une fois immédiatement, au lieu d'attendre un redémarrage |
| `-Unregister` | Supprime la tâche et la copie installée du script |

**Exemples**

```powershell
# Afficher ce qui serait installé et enregistré
.\Register-InitTempDiskTask.ps1 -WhatIf

# Installer, enregistrer la tâche de démarrage et l'exécuter une fois maintenant
.\Register-InitTempDiskTask.ps1 -RunNow

# Idem, mais avec un fichier d'échange fixe de 16 Go
.\Register-InitTempDiskTask.ps1 -ScriptArguments '-Quiet -RestartIfNeeded -InitialSizeMB 16384 -MaximumSizeMB 16384'

# Sans le redémarrage — le fichier d'échange arrive au prochain redémarrage, quel qu'il soit
.\Register-InitTempDiskTask.ps1 -ScriptArguments '-Quiet'

# Le retirer de l'appareil
.\Register-InitTempDiskTask.ps1 -Unregister
```

**Remarques**
- La tâche exécute la *copie* dans `C:\Scripts`, jamais la source, donc le dépôt ou le dossier de préparation du RMM peut disparaître ensuite
- Un fichier source manquant est une erreur bloquante et non un saut silencieux — une tâche enregistrée sur un fichier absent s'exécute et échoue à chaque démarrage sans que personne ne le remarque, jusqu'à ce que le fichier d'échange disparaisse
- Le déclencheur de démarrage est retardé de 30 secondes par défaut : la pile de stockage n'a pas toujours énuméré le disque temporaire au moment où le moteur de tâches est prêt
- `-Unregister` laisse volontairement la configuration du fichier d'échange intacte — supprimer la tâche ne doit pas emporter le fichier d'échange de la machine
- L'enregistrement indique clairement si la tâche peut redémarrer la machine ; l'exécuter avec `-RunNow` sur un appareil où personne n'est connecté peut donc lancer un compte à rebours de 60 secondes
