[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../../readme.fr.md) › [scripts](../../../readme.fr.md) › [Intune](../../readme.fr.md) › [Desktop](../readme.fr.md) › **ClaudeDesktop**

# ClaudeDesktop

Déploiement Intune de [Claude Desktop](https://claude.com/download) pour Windows à l'échelle de la machine. Exécutez **un seul script une fois par mois** pour garder l'application Intune à jour — aucune App Registration ni aucun secret client permanent à gérer.

Basé sur [Deploy Claude Desktop for Windows](https://support.claude.com/en/articles/12622703-deploy-claude-desktop-for-windows) et [Enterprise configuration for Claude Desktop](https://support.claude.com/en/articles/12622667-enterprise-configuration-for-claude-desktop).

> **Les prérequis Windows de Cowork (VirtualMachinePlatform, démarrage rapide) forment une application Win32 distincte et indépendante** — voir [`../CoworkPrerequisites/readme.md`](../CoworkPrerequisites/readme.fr.md). Déployez les deux si vous voulez Cowork ; déployez uniquement celle-ci si vous voulez seulement Claude Desktop. **Par défaut, il n'existe aucune dépendance Intune** entre les deux : un échec des prérequis Cowork ne doit jamais bloquer Claude Desktop (qui fonctionne très bien sans Cowork), et un problème d'installation de Claude Desktop ne doit jamais être confondu avec un problème de fonctionnalité Windows — chaque application a son propre état d'installation, visible séparément dans Intune. Si votre environnement considère Cowork comme une exigence stricte plutôt qu'une option, passez `-RequireCoworkPrerequisites` à *ce* script de déploiement pour ajouter à la place une véritable dépendance Intune — voir « Option : exiger Cowork Prerequisites » ci-dessous.

---

## Pourquoi ne pas simplement charger le MSIX comme application métier (LOB) ?

Intune installe les applications MSIX métier par utilisateur. Cela échoue pour les utilisateurs standard sans droits d'administrateur. À la place, `Add-AppxProvisionedPackage` est encapsulé dans une application Win32, exactement comme le recommande l'article officiel.

## Contenu

| Script | Rôle dans Intune |
|---|---|
| [`Deploy-ClaudeDesktopIntune.ps1`](Deploy-ClaudeDesktopIntune.ps1) ([docs](#deploy-claudedesktopintuneps1)) | **Celui que vous exécutez.** Orchestre tout ce qui suit — voir la section « Exécution mensuelle ». |
| [`Install-ClaudeDesktop-Intune.ps1`](Install-ClaudeDesktop-Intune.ps1) ([docs](#install---uninstall---detect-claudedesktop-intuneps1)) | Script de contenu de la commande d'installation |
| [`Uninstall-ClaudeDesktop-Intune.ps1`](Uninstall-ClaudeDesktop-Intune.ps1) ([docs](#install---uninstall---detect-claudedesktop-intuneps1)) | Script de contenu de la commande de désinstallation |
| [`Detect-ClaudeDesktop-Intune.ps1`](Detect-ClaudeDesktop-Intune.ps1) ([docs](#install---uninstall---detect-claudedesktop-intuneps1)) | Script de détection personnalisé |

`Install-`/`Uninstall-`/`Detect-ClaudeDesktop-Intune.ps1` ne sont jamais exécutés manuellement — `Deploy-ClaudeDesktopIntune.ps1` les empaquette automatiquement dans le `.intunewin` (ou, pour le script de détection, le charge dans le cadre de la règle de détection).

## Deploy-ClaudeDesktopIntune.ps1

Ce qu'il fait :

1. Télécharge le dernier MSIX x64 de Claude Desktop depuis l'URL de redirection « latest » officielle d'Anthropic.
2. Lit la version dans le fichier `AppxManifest.xml` contenu dans le MSIX.
3. Copie les scripts d'installation/désinstallation à côté du MSIX et construit un package `.intunewin` (module `IntuneWin32App` — `IntuneWinAppUtil.exe` est téléchargé automatiquement s'il n'est pas déjà présent).
4. Se connecte à Microsoft Graph en mode délégué (connexion interactive) et crée une **App Registration temporaire** de courte durée — même modèle que [`Remove-SharePointFileVersionsByDate.ps1`](../../../Reporting/readme.fr.md) — avec uniquement l'autorisation d'application `DeviceManagementApps.ReadWrite.All`. Elle sert à authentifier le module `IntuneWin32App`, puis est supprimée à la fin de l'exécution. Rien ne persiste entre deux exécutions, hormis l'application Intune elle-même. La connexion passe par [`Connect-M365.ps1`](../../../Startup/Connect-M365.ps1) : déléguée par défaut (code d'appareil selon `$global:useDeviceCodeAuth`, client GDAP depuis `$global:cid`). Le téléversement lui-même reste confié au module `IntuneWin32App` : `Connect-MSIntuneGraph` obtient son propre jeton et ne peut pas utiliser la session Microsoft Graph PowerShell, et réécrire en appels Graph bruts le téléversement par blocs du `.intunewin` vers Azure Storage ainsi que la validation des informations de chiffrement représente beaucoup de risque pour peu de gain. Avec `-ClientId` + `-CertificateThumbprint` (ou `-AppOnly`), vous utilisez à la place votre propre application permanente : rien de temporaire n'est créé et `Connect-MSIntuneGraph` se connecte avec le même certificat (`-ClientCert`).
5. Première exécution : crée l'application Win32 « Claude Desktop (Machine-wide) » dans Intune avec les règles de détection et de configuration requise, et l'attribue en **Required** au groupe Entra ID que vous indiquez.
6. Exécutions suivantes : pousse un package mis à jour via `Update-IntuneWin32AppPackageFile` (l'attribution existante reste intacte, les appareils reçoivent simplement le nouveau contenu) si **soit** la version du MSIX téléchargé est plus récente, **soit** les scripts Install-/Uninstall-/Detect-ClaudeDesktop-Intune.ps1 eux-mêmes ont changé depuis la dernière exécution — les deux étant suivis dans le champ Notes de l'application (`ClaudeMsixVersion=...; ScriptsHash=...`), sans fichier d'état local. Les règles de détection **et de configuration requise sont reconstruites et renvoyées à chaque exécution**, pas seulement à la création (voir « Problème connu » ci-dessous). Si ni la version ni les scripts n'ont changé, seules les règles sont actualisées.

La détection est volontairement indépendante de la version (présence du package provisionné). Intune redéploie une application Win32 sur les appareils déjà ciblés dès que sa version de contenu change dans Intune, quel que soit le résultat de la règle de détection — il n'y a donc pas de `$MinimumVersion` à incrémenter à la main chaque mois, et une simple modification de script (sans nouveau MSIX) déclenche quand même un redéploiement grâce à la vérification `ScriptsHash` ci-dessus.

### Problème connu : échecs d'installation 0x80070001 — et un bug bien plus profond du module `IntuneWin32App` derrière

Les appareils échouaient à l'installation avec l'erreur `0x80070001`, quel que soit l'appareil. La comparaison de la stratégie d'application Win32 en cache sur un appareil touché (`AppWorkload.log`) avec toutes les autres applications Win32 attribuées au même tenant a révélé l'anomalie : pour toutes les autres applications, les `RequirementRules` contenaient `RequiredOSArchitecture: 3`, mais celles de Claude Desktop contenaient `RequiredOSArchitecture: 32` — une valeur qu'aucune autre application du tenant n'utilisait, et 16 fois ce que `-Architecture 'x64'` devrait produire. Comme cette valeur est stockée sur l'objet application dans Intune (et non par appareil), cela expliquait pourquoi l'échec était reproductible à 100 % sur tous les appareils, et non une corruption propre à un appareil.

**Première tentative de correction (incomplète) :** renvoyer `-RequirementRule` à chaque mise à jour, pas seulement à la création. Cela ne fonctionnait en fait pas, à cause d'un second bug bien plus grave :

**La véritable cause racine**, confirmée par le code source du module `IntuneWin32App` lui-même (v1.5.0) et par le schéma documenté par Microsoft de la [ressource `win32LobApp`](https://learn.microsoft.com/en-us/graph/api/resources/intune-apps-win32lobapp) :
- `applicableArchitectures` / `allowedArchitectures` / `minimumSupportedWindowsRelease` sont des **propriétés plates de premier niveau** de `win32LobApp` — elles n'ont besoin d'aucun `@odata.type` (celui-ci n'est requis que pour la collection polymorphe `rules` : règles de détection ou de configuration requise basées sur un fichier/le registre/un code de produit/un script).
- `New-IntuneWin32AppRequirementRule` ne définit jamais `@odata.type` sur l'objet qu'il renvoie (à juste titre — il n'en a pas besoin).
- Mais `Set-IntuneWin32App` (uniquement la cmdlet de **mise à jour** — `Add-IntuneWin32App`, utilisée à la création, a une logique interne entièrement différente et correcte) exige malgré tout, à tort, `@odata.type` sur `-RequirementRule`, et lorsqu'il manque, exécute `Write-Warning "...missing required '@odata.type'..."` suivi d'un simple `break`.
- Ce `break` n'est entouré d'aucune boucle ni d'aucun `switch` — vérifié empiriquement, il met fin à **tout le reste de la fonction**, y compris l'appel Graph PATCH proprement dit plus bas. Autrement dit : chaque appel à `Set-IntuneWin32App` comprenant `-RequirementRule` ne faisait silencieusement **absolument rien** — ni la correction de l'architecture, ni `-Notes`, ni `-DetectionRule`, ni `-Icon`, ni `-CompanyPortalFeaturedApp`. Seul `Update-IntuneWin32AppPackageFile` (un appel distinct, effectué juste avant) atteignait réellement Intune à chaque exécution.

**La vraie correction** : la branche de mise à jour ne transmet plus du tout `-RequirementRule` à `Set-IntuneWin32App` (ce qui laisse passer normalement Notes/AppVersion/DetectionRule/Icon/CompanyPortalFeaturedApp), et appelle à la place `Set-Win32AppArchitectureRequirement` — une petite fonction d'aide qui applique directement un PATCH sur `applicableArchitectures`/`allowedArchitectures`/`minimumSupportedWindowsRelease` via Microsoft Graph (en réutilisant la même session `$Global:AuthenticationHeader` déjà établie par `Connect-MSIntuneGraph`), contournant entièrement le chemin défectueux de la cmdlet pour ce seul élément. `Add-IntuneWin32App` (la création) n'est pas du tout concerné et continue d'utiliser `-RequirementRule` normalement.

**Si un appareil échoue encore après cette correction** : vérifiez s'il n'est pas plutôt bloqué derrière le **délai de relance GRS** d'Intune, sans rapport (3 tentatives échouées → blocage de 24 h, quelle que soit la configuration de l'application) — voir les [scripts de délai GRS](../../readme.fr.md#repair-stuckwin32appenforcementps1) un niveau au-dessus, soit la version manuelle sur l'appareil, soit la paire Detect-/Remediate- qui s'exécute entièrement via le portail Intune. Cela touche toutes les applications Win32 de cet appareil, pas seulement Claude ; c'est donc une première vérification utile si plusieurs applications sans rapport sont elles aussi bloquées en silence.

### Visibilité dans le Portail d'entreprise

Le script de déploiement extrait le logo de l'application directement du MSIX téléchargé (`Properties/Logo` dans `AppxManifest.xml`, avec repli sur la variante de plus haute échelle réellement présente dans le package) et le définit comme icône de l'application Win32, avec en plus `-CompanyPortalFeaturedApp $true` à chaque exécution — ainsi, au lieu d'une icône Win32 générique perdue dans la liste complète des applications, les utilisateurs voient le vrai logo Claude, mis en avant, dans le Portail d'entreprise.

### Rôle requis

Global Administrator, ou Application Administrator combiné à un rôle capable d'accorder le consentement `AppRoleAssignment.ReadWrite.All` — la même exigence que pour l'App Registration temporaire des scripts de reporting SharePoint.

### Exécution mensuelle

```powershell
.\Deploy-ClaudeDesktopIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop"
```

Vous obtiendrez une invite de connexion interactive et une confirmation « type JA to continue » avant que quoi que ce soit ne soit créé ou mis à jour dans Intune. Ajoutez `-Force` pour ignorer la confirmation (par ex. pour une exécution planifiée/sans surveillance une fois que vous maîtrisez le processus).

| Paramètre | Valeur par défaut | Description |
|---|---|---|
| `-AssignmentGroupName` | *(obligatoire)* | Groupe Entra ID attribué en Required. Utilisé uniquement à la création. |
| `-WorkingDirectory` | `C:\Temp\ClaudeDeploy` | Dossier de préparation du téléchargement/de la construction (`Source/`, `Output/`) |
| `-AppDisplayName` | `Claude Desktop (Machine-wide)` | Sert à retrouver l'application existante lors des exécutions suivantes — ne le modifiez pas sans renommer aussi l'application dans Intune |
| `-MsixDownloadUrl` | Redirection x64 « latest » officielle d'Anthropic | À remplacer pour les tests |
| `-MinimumSupportedWindowsRelease` | `W10_21H2` | Règle de configuration requise |
| `-TenantId` | client GDAP, sinon le tenant de connexion | ID ou domaine du tenant Entra ID |
| `-ClientId` | — | Application permanente facultative pour l'application seule (avec `-CertificateThumbprint`) ; aucune App Registration temporaire n'est créée |
| `-CertificateThumbprint` | — | Certificat pour `-ClientId` (CurrentUser\My ou LocalMachine\My) ; utilisé pour Graph et pour `Connect-MSIntuneGraph -ClientCert` |
| `-AppOnly` | désactivé | Application seule avec le ClientId et l'empreinte de `graph.appid.json` |
| `-IntuneWinAppUtilPath` | téléchargement automatique | Utiliser un `IntuneWinAppUtil.exe` déjà téléchargé |
| `-Force` | désactivé | Ignorer la ou les invites de confirmation |
| `-RequireCoworkPrerequisites` | désactivé | Ajouter une véritable dépendance Intune envers l'application Cowork Prerequisites — voir ci-dessous |
| `-CoworkPrerequisitesAppDisplayName` | `Cowork Windows Prerequisites (Machine-wide)` | Doit correspondre au `-AppDisplayName` utilisé dans `Deploy-CoworkPrerequisitesIntune.ps1`. Utilisé uniquement avec `-RequireCoworkPrerequisites` |

### Option : exiger Cowork Prerequisites

Par défaut, Claude Desktop et Cowork Prerequisites sont indépendants — aucune dépendance Intune ne les relie (voir la remarque en haut de ce readme pour savoir pourquoi). Si votre environnement a réellement besoin que Cowork soit toujours présent — un appareil sans prérequis Cowork fonctionnels ne doit pas recevoir Claude Desktop du tout — passez `-RequireCoworkPrerequisites` :

```powershell
.\Deploy-ClaudeDesktopIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop" -RequireCoworkPrerequisites
```

Le script recherche l'application Cowork Prerequisites (déployez-la **d'abord** — `Deploy-CoworkPrerequisitesIntune.ps1`) et appelle `Add-IntuneWin32AppDependency` avec `DependencyType 'Detect'` (et non `'AutoInstall'`) : l'application de prérequis doit toujours être attribuée indépendamment en Required et déjà détectée sur l'appareil — cette dépendance ne l'installe pas automatiquement pour le compte de Claude Desktop, elle fait seulement attendre Intune. Le compromis par rapport au comportement par défaut : un appareil bloqué sur les prérequis Cowork (par ex. en plein redémarrage, ou en échec réel) ne recevra pas non plus Claude Desktop tant que le problème n'est pas résolu, au lieu de recevoir Claude Desktop immédiatement et Cowork plus tard.

## Install- / Uninstall- / Detect-ClaudeDesktop-Intune.ps1

Les scripts de contenu de l'application Win32. `Deploy-ClaudeDesktopIntune.ps1` les empaquette et les charge ; ils ne sont jamais exécutés manuellement.

- **`Install-ClaudeDesktop-Intune.ps1`** (commande d'installation) : supprime entièrement toute installation existante de Claude Desktop, provisionne le MSIX à l'échelle de la machine avec `Add-AppxProvisionedPackage` et désactive la mise à jour automatique propre à Claude — voir les sections ci-dessous. Paramètre `-MsixFileName` (par défaut `Claude.msix`) : nom du fichier MSIX à côté du script. Journalise dans `%ProgramData%\ClaudeDeploy\install.log`.
- **`Uninstall-ClaudeDesktop-Intune.ps1`** (commande de désinstallation) : supprime le package provisionné à l'échelle de la machine et, en solution de repli, les installations Appx par utilisateur des profils déjà connectés. Ne touche pas aux prérequis Windows de Cowork. Journalise dans `%ProgramData%\ClaudeDeploy\uninstall.log`.
- **`Detect-ClaudeDesktop-Intune.ps1`** (script de détection personnalisé, 64 bits) : signale « installed » lorsque le package provisionné à l'échelle de la machine est présent. Volontairement indépendant de la version, et ne réessaie qu'en cas d'exception DISM transitoire.

### Stratégie de mise à jour automatique

Le script d'installation définit `HKLM:\SOFTWARE\Policies\Claude\disableAutoUpdates = 1` (DWord). Le programme de mise à jour propre à Claude est désactivé afin que le contrôle des versions reste entièrement entre les mains de cette exécution Intune mensuelle, au lieu de dériver appareil par appareil.

### Réinstallation propre à chaque exécution

Avant de provisionner la nouvelle version, le script d'installation supprime d'abord complètement Claude Desktop de l'appareil, en trois passes :

1. Arrête tout processus Claude en cours d'exécution.
2. Supprime chaque installation **Appx** par utilisateur (`Get-AppxPackage -AllUsers` / `Remove-AppxPackage -AllUsers`, y compris les profils déjà connectés), puis l'ancien package provisionné à l'échelle de la machine.
3. Supprime toute **installation classique (non Appx) par utilisateur** — notamment l'installateur grand public de [claude.ai/download](https://claude.ai/download), qui s'enregistre via une clé de registre Uninstall ordinaire par utilisateur plutôt que comme package Appx, si bien que `Get-AppxPackage` ne la voit jamais. Le script parcourt la clé de registre Uninstall de chaque profil local — y compris les profils qui ne sont pas connectés à ce moment-là, en chargeant temporairement leur `NTUSER.DAT` — et exécute la `QuietUninstallString` de chaque correspondance (ou la `UninstallString` si elle est absente), avec un délai d'expiration de 120 s pour qu'un programme de désinstallation bloqué ne puisse pas figer l'installation Intune.

Ce n'est qu'après ces trois passes qu'il provisionne le nouveau MSIX. C'est volontairement plus approfondi que de simplement vider la couche de provisionnement — une installation restante issue d'un autre canal peut sinon continuer à faire tourner sa propre session Claude non gérée, même après la mise à jour de la version à l'échelle de la machine. Un utilisateur qui a Claude ouvert perd cette session lorsque ce script s'exécute.

### Robustesse sur un appareil où il n'a jamais été installé

Les deux scripts de contenu sont écrits de sorte qu'un appareil sur lequel Claude n'a jamais été présent — y compris lors de toute première exécution de l'ESP Autopilot — passe sans encombre :

- **Script d'installation** : chaque passe de suppression (arrêt du processus, Appx, désinstallation classique par utilisateur) vérifie d'abord si le résultat est vide/`$null` avant d'agir, de sorte qu'une machine totalement propre se contente de journaliser « none found » à chaque étape au lieu de produire une erreur.
- **Script de détection** : un résultat vide/négatif de `Get-AppxProvisionedPackage` (le résultat attendu sur un appareil où il n'a jamais été installé) renvoie immédiatement « not installed » (exit 1), sans nouvelle tentative — les nouvelles tentatives n'interviennent que sur une véritable **exception** (par ex. un verrou DISM transitoire, plausible juste après l'intense activité Appx du script d'installation lui-même sur le même appareil), si bien qu'une machine réellement propre n'est jamais ralentie par des tentatives qui ne peuvent pas changer le résultat.

## Prérequis

- PowerShell 7 (les scripts de déploiement chargent `scripts\Startup\Connect-M365.ps1`). Le script de déploiement lui-même fait partie de `ScriptsHash` : la première exécution après la mise à jour de ce script téléverse donc le paquet une fois de plus.
- Les modules PowerShell `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications`, `Microsoft.Graph.Groups` et `IntuneWin32App` — à installer avec `.\scripts\Startup\Install-Modules.ps1`
- Exécution depuis Windows (l'outil d'empaquetage et les cmdlets MSIX/AppX n'existent que sous Windows)
