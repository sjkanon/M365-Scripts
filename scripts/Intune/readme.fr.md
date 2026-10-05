[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **Intune**

# Intune

Inscription Autopilot, automatisation des stratégies de conformité, détection des dérives de configuration et déploiement du poste de travail chez les clients (fond d'écran, écran de verrouillage, raccourci de verrouillage dans la barre des tâches).

> Le déploiement des thèmes et couleurs Office se trouve dans [`Custom Scripts/Intune/Desktop/`](../Custom%20Scripts/Intune/readme.fr.md) — ces scripts reçoivent l'URL de téléchargement du thème en paramètre.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Compare-IntuneConfig.ps1`](Compare-IntuneConfig.ps1) ([docs](#compare-intuneconfigps1)) | Compare la configuration Intune d'un tenant client à une sauvegarde de référence (baseline) du MSP |
| [`Repair-StuckWin32AppEnforcement.ps1`](Repair-StuckWin32AppEnforcement.ps1) ([docs](#repair-stuckwin32appenforcementps1)) | Débloque sur un appareil les applications Win32 coincées derrière le délai de relance GRS d'Intune — à exécuter manuellement, en local, avec un filtre essai à blanc/rapport/`-AppId` |
| [`Detect-StuckWin32AppEnforcement.ps1`](Detect-StuckWin32AppEnforcement.ps1) + [`Remediate-StuckWin32AppEnforcement.ps1`](Remediate-StuckWin32AppEnforcement.ps1) ([docs](#detect---remediate-stuckwin32appenforcementps1)) | Le même correctif, empaqueté sous forme de paire Intune Remediation — déclenché entièrement depuis le portail Intune, sans accès à l'appareil |

## Dossiers

| Dossier | Description |
|--------|-------------|
| [`Get-Autopilot/`](Get-Autopilot/readme.fr.md) | Collecte des hachages matériels Windows Autopilot |
| [`iOS-Compliance-Updater/`](iOS-Compliance-Updater/readme.fr.md) | Met à jour automatiquement la version iOS minimale dans une stratégie de conformité Intune |
| [`Desktop/`](Desktop/readme.fr.md) | Fond d'écran et écran de verrouillage de l'entreprise, raccourci de verrouillage dans la barre des tâches |
| [`DiskCleanup/`](DiskCleanup/readme.fr.md) | Wrapper d'application Win32 qui exécute le script de nettoyage du disque C:\ puis redémarre l'appareil |

---

### Compare-IntuneConfig.ps1

Encapsule le module communautaire `IntuneBackupAndRestore` pour détecter les dérives de configuration Intune — stratégies de conformité, profils de configuration et autres objets exportés qui diffèrent d'une sauvegarde de référence « baseline » (ou qui y manquent ou s'y ajoutent). En lecture seule — ne modifie jamais la configuration, se contente de signaler les différences.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-BaselinePath` | Oui | Dossier de la sauvegarde de référence du MSP (issu d'une exécution antérieure de `Start-IntuneBackup -Path <path>`) |
| `-CustomerBackupPath` | Non | Sauvegarde existante du tenant client. S'il est omis, le script sauvegarde d'abord le tenant actuellement connecté |
| `-TenantId` | Non | ID de tenant ou domaine auquel se connecter (utilisé uniquement si `-CustomerBackupPath` est omis) |
| `-OutputPath` | Non | Dossier pour la sauvegarde automatique et le rapport des différences (par défaut : `C:\Temp\` / `~/Downloads`) |

**Exemples**

```powershell
# Comparer une sauvegarde client existante à la baseline du MSP
.\Compare-IntuneConfig.ps1 -BaselinePath C:\IntuneBaseline -CustomerBackupPath C:\Temp\CustomerBackup

# Sauvegarder le tenant client actuellement connecté (GDAP) et le comparer en direct
.\Compare-IntuneConfig.ps1 -BaselinePath C:\IntuneBaseline
```

**Remarques**
- Compatible GDAP : si `-TenantId` est omis, il est déduit automatiquement du tenant client sélectionné (`$global:cid`), comme pour `Move-InboxToArchive.ps1` / `Get-SharePointStorageReport.ps1`
- Non intégré à `menu.ps1` — il travaille avec des chemins de dossiers de sauvegarde et un export à l'échelle du tenant, exécutez-le directement

**Module requis**

```powershell
Install-Module IntuneBackupAndRestore -Scope CurrentUser
```

---

### Repair-StuckWin32AppEnforcement.ps1

À exécuter **sur l'appareil concerné**, en tant qu'administrateur. L'agent d'applications Win32 d'Intune (IME) relance une installation en échec 3 fois, à 5 minutes d'intervalle, puis bloque l'application pendant un délai « GRS » de 24 heures — visible dans `AppActionProcessor.log` sous la forme `... to install is in GRS. The app will not be enforced.` Pendant ce délai, IME ignore complètement l'application, même après que vous avez corrigé le vrai problème dans Intune (règle de configuration requise corrigée, nouveau package, commande d'installation réparée) — ce délai est un état local de l'appareil, qu'un redéploiement côté Intune ne peut pas effacer.

Ce script recherche sous `HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps` chaque application Win32 en cache local ayant un véritable dernier code d'erreur (ni 0/succès, ni 3010/redémarrage en attente). Il peut effacer son état d'application en cache, son cache de rapport et la clé de délai GRS correspondante, puis redémarre le service IME afin que chaque application soit réévaluée à neuf lors du prochain check-in.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-Apply` | Efface réellement l'état bloqué et redémarre le service (par défaut : rapport d'essai à blanc uniquement) |
| `-AppId` | Limite à un GUID d'application Win32 précis (tiré de l'URL de l'application dans le centre d'administration Intune). Par défaut : toutes les applications bloquées trouvées |
| `-ForceSync` | Après l'effacement, déclenche aussi un check-in MDM immédiat (comme le bouton « Synchroniser » du Portail d'entreprise) |
| `-OutputPath` | Chemin du rapport CSV (par défaut : `C:\Temp\`) |

**Exemples**

```powershell
# Essai à blanc — voir ce qui est actuellement bloqué sur cet appareil
.\Repair-StuckWin32AppEnforcement.ps1

# Tout débloquer et forcer une resynchronisation immédiate
.\Repair-StuckWin32AppEnforcement.ps1 -Apply -ForceSync

# Débloquer une seule application précise
.\Repair-StuckWin32AppEnforcement.ps1 -Apply -AppId "96ff358d-0e16-4224-b946-eb61bc930fca"
```

**Remarques**
- Il s'agit d'un état local à l'appareil — il affecte **toutes** les applications Win32 attribuées à cet appareil, pas une seule. Un seul appareil bloqué en GRS peut donc donner l'impression que plusieurs déploiements d'applications sans rapport échouent tous en silence en même temps.
- Basé sur la structure de registre GRS documentée par la communauté (non publiée officiellement par Microsoft) — voir la section `.NOTES` du script pour les sources.

---

### Detect- / Remediate-StuckWin32AppEnforcement.ps1

Le même correctif que `Repair-StuckWin32AppEnforcement.ps1` ci-dessus, scindé en une paire détection/correction pour Intune **Devices → Scripts and remediations → Remediations**, afin que le déblocage d'un appareil ne nécessite jamais de session RDP/console manuelle — tout se déclenche depuis le portail Intune.

| Script | Rôle |
|---|---|
| [`Detect-StuckWin32AppEnforcement.ps1`](Detect-StuckWin32AppEnforcement.ps1) | Partie détection — exit 1 si une application Win32 a un véritable dernier code d'erreur en cache (blocage GRS possible), exit 0 sinon |
| [`Remediate-StuckWin32AppEnforcement.ps1`](Remediate-StuckWin32AppEnforcement.ps1) | Partie correction — efface toujours tout ce que le script de détection a trouvé, redémarre IME et force une synchronisation MDM immédiate |

**Déploiement dans Intune :**

1. **Devices → Scripts and remediations → Remediations → Create**
2. Nommez-le par exemple `Clear Stuck Win32 App Enforcement`
3. Chargez `Detect-StuckWin32AppEnforcement.ps1` comme script de détection et `Remediate-StuckWin32AppEnforcement.ps1` comme script de correction
4. Run using logged-on credentials : **No** (SYSTEM) · Run in 64-bit PowerShell : **Yes** · Enforce signature check : **No**
5. **Ne l'attribuez à aucun groupe d'appareils et ne définissez aucune planification.** Laissez l'attribution vide.

**Déclenchement à la demande, par appareil, entièrement depuis le portail** (sans planification, sans accès à l'appareil) :

1. **Devices → All devices → [l'appareil concerné]**
2. **… (points de suspension) → Run remediation (preview)**
3. Sélectionnez `Clear Stuck Win32 App Enforcement` et exécutez-le

**Pourquoi non attribué / à la demande plutôt qu'une planification permanente :** une attribution planifiée à l'échelle de tout le parc continuerait d'effacer silencieusement les blocages GRS pour *n'importe quelle* application qui échoue 3 fois, pour *n'importe quelle* raison — masquant un déploiement réellement défaillant derrière des relances automatiques sans fin au lieu de le faire apparaître. Le laisser non attribué et ne l'exécuter que sur un appareil précis, une fois que vous avez confirmé (via `AppActionProcessor.log`/`AppWorkload.log`, ou parce que le problème sous-jacent est déjà corrigé) qu'une nouvelle tentative se justifie vraiment, évite cet écueil. Cela reprend le modèle établi par la communauté pour ce scénario précis (voir la section `.NOTES` du script) ; ce n'est pas une invention propre à ce dépôt.
