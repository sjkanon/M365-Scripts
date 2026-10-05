[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../../readme.fr.md) › [scripts](../../../readme.fr.md) › [Intune](../../readme.fr.md) › [Desktop](../readme.fr.md) › **CoworkPrerequisites**

# CoworkPrerequisites

Déploiement Intune, à l'échelle de la machine, des prérequis côté Windows pour [Claude Cowork](https://support.claude.com/en/articles/12622667-enterprise-configuration-for-claude-desktop) — la fonctionnalité facultative `VirtualMachinePlatform` et la désactivation du démarrage rapide de Windows — empaquetés dans leur **propre application Win32 indépendante**, séparée de [`../ClaudeDesktop/`](../ClaudeDesktop/readme.fr.md).

## Pourquoi une application distincte plutôt que de l'intégrer au script d'installation de Claude Desktop

Cowork est facultatif : Claude Desktop fonctionne très bien sans. Intégrer l'étape de fonctionnalité Windows au script d'installation de Claude Desktop lui-même signifie qu'un accroc DISM sans rapport avec Claude (une pile de maintenance occupée, Windows Update injoignable comme source de fonctionnalités — deux cas plus probables sur un appareil fraîchement imagé en plein ESP Autopilot) pourrait interrompre *toute* l'installation de Claude Desktop. En le séparant :

- Un échec des prérequis Cowork ne bloque jamais Claude Desktop lui-même.
- Chaque application a son propre état d'installation, visible séparément dans Intune — vous voyez d'un coup d'œil si le problème d'un appareil relève de la « fonctionnalité Windows » ou de « Claude Desktop », au lieu d'une commande d'installation combinée opaque.

Il n'y a volontairement **aucune « Dependency » Intune** configurée par défaut entre les deux applications. Une dépendance stricte signifierait que Claude Desktop ne tenterait même pas de s'installer tant que cette application n'a pas réussi — ce qui réintroduit exactement le problème ci-dessus. Attribuez les deux en **Required** au même groupe, indépendamment.

Si votre environnement considère Cowork comme une exigence stricte plutôt qu'une option, déployez d'abord cette application, puis exécutez `Deploy-ClaudeDesktopIntune.ps1 -RequireCoworkPrerequisites` — voir [`../ClaudeDesktop/readme.md`](../ClaudeDesktop/readme.fr.md#option--exiger-cowork-prerequisites) pour ce que cela ajoute et le compromis qui en découle.

## Contenu

| Script | Rôle dans Intune |
|---|---|
| [`Deploy-CoworkPrerequisitesIntune.ps1`](Deploy-CoworkPrerequisitesIntune.ps1) ([docs](#deploy-coworkprerequisitesintuneps1)) | **Celui que vous exécutez** pour la voie application Win32. Empaquette les scripts ci-dessous et crée/met à jour l'application Win32. |
| [`Install-CoworkPrerequisites-Intune.ps1`](Install-CoworkPrerequisites-Intune.ps1) ([docs](#install-coworkprerequisites-intuneps1)) | Script de contenu de la commande d'installation de l'application Win32 |
| [`Uninstall-CoworkPrerequisites-Intune.ps1`](Uninstall-CoworkPrerequisites-Intune.ps1) ([docs](#uninstall-coworkprerequisites-intuneps1)) | Script de contenu de la commande de désinstallation de l'application Win32 |
| [`Detect-CoworkPrerequisites-Intune.ps1`](Detect-CoworkPrerequisites-Intune.ps1) ([docs](#detect-coworkprerequisites-intuneps1)) | Script de détection personnalisé de l'application Win32 |
| [`CoworkPrerequisites-PlatformScript.ps1`](CoworkPrerequisites-PlatformScript.ps1) ([docs](#coworkprerequisites-platformscriptps1)) | Voie **alternative** autonome — pas de script Deploy, pas d'empaquetage, chargé directement comme « Platform script » Intune. |

Il n'y a pas de MSIX ici — contrairement à Claude Desktop, le « contenu » se limite à ces trois scripts ; une nouvelle exécution ne fait donc quelque chose que si vous avez réellement modifié l'un d'eux (suivi via un `ScriptsHash` dans le champ Notes de l'application, même modèle que `Deploy-ClaudeDesktopIntune.ps1`).

## Install-CoworkPrerequisites-Intune.ps1

Commande d'installation de l'application Win32. Ce qu'il fait :

1. **VirtualMachinePlatform** : `Enable-WindowsOptionalFeature`, avec nouvelles tentatives en cas d'échecs DISM transitoires. Ne fait rien si la fonctionnalité est déjà activée.
2. **Démarrage rapide** : définit `HiberbootEnabled = 0` sous `HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power` à chaque exécution (pas seulement lorsque VMP vient d'être activé). La documentation Cowork d'Anthropic avertit explicitement : *"Restart the machine using Restart, not shut down and power on. With Windows Fast Startup enabled, a shutdown cycle can leave the virtualization services uninitialized."* Le démarrage rapide est activé par défaut sur presque toutes les images Windows — sans le désactiver, un utilisateur qui se contente d'« arrêter » au lieu de « redémarrer » (le cas courant) n'obtient jamais de services Cowork fonctionnels, quel que soit le nombre de cycles d'alimentation, alors même qu'Intune affiche l'application comme installée.
3. Si VMP vient d'être activé : se termine avec le code **3010** (« soft reboot required »). L'application est configurée avec `-RestartBehavior 'basedOnReturnCode'`, c'est donc **Intune lui-même** qui impose le redémarrage (invite/échéance/délai de grâce) — cela ne dépend pas de la présence d'un utilisateur connecté. Par courtoisie, le script envoie aussi une notification `msg.exe` en anglais à la session console active (reconnue grâce au `SESSIONNAME` non traduit « console », et non au texte `STATE` « Active », qui dépend de la langue du système — une simple comparaison de chaîne sur « Active » ne se déclencherait jamais, en silence, sur un Windows affiché dans une autre langue que l'anglais), mais cette notification est un plus, pas le mécanisme sur lequel repose réellement le redémarrage.

   La notification n'est **pas** envoyée en appelant `msg.exe` directement depuis ce script en contexte SYSTEM — cela produit une boîte de dialogue dont le bouton OK ne réagit pas aux clics (une bizarrerie connue de `msg.exe` pour les messages inter-sessions émis par un expéditeur non interactif). À la place, `Show-UserRestartNotification` enregistre une tâche planifiée de courte durée (`LogonType Interactive`, principal = l'utilisateur de la session console) qui exécute `msg.exe` *dans la session de l'utilisateur*, puis la désinscrit une fois le message envoyé — la boîte de dialogue appartient alors à un bureau réellement interactif et se ferme normalement.

## Detect-CoworkPrerequisites-Intune.ps1

Script de détection personnalisé de l'application Win32.

Signale « installed » uniquement si **à la fois** `VirtualMachinePlatform` est `Enabled` **et** les services HCS sous-jacents (`vmcompute`, `HNS`, `vfpext`) sont présents — le fait que VMP apparaisse `Enabled` dans DISM ne garantit pas que ces services existent déjà (voir le dépannage Cowork d'Anthropic : *"Missing HCS services: HNS, vmcompute, vfpext"*), en particulier juste après l'activation de VMP mais avant le redémarrage requis.

Volontairement, **aucune vérification du `Status` du service** (par ex. `Running`) : `vmcompute` est un service à démarrage par déclencheur, censé être `Stopped` tant qu'aucune session Cowork n'est active — vérifier `Running` signalerait un appareil parfaitement sain comme « not installed ». Seule l'absence complète du service (`Get-Service` ne le trouve pas) est un signal fiable que les composants Hyper-V sous-jacents ne sont pas encore là.

## Problème connu : `-RequirementRule` sur `Set-IntuneWin32App` ne fait silencieusement rien (corrigé)

Le même bug confirmé du module `IntuneWin32App` (1.5.0) que dans [`../ClaudeDesktop/readme.md`](../ClaudeDesktop/readme.fr.md#problème-connu--échecs-dinstallation-0x80070001--et-un-bug-bien-plus-profond-du-module-intunewin32app-derrière) : le paramètre `-RequirementRule` de `Set-IntuneWin32App` exige à tort une propriété `@odata.type` que `New-IntuneWin32AppRequirementRule` ne définit jamais, et le `break` qui suit met fin à **tout le reste de la fonction** — si bien que chaque appel de mise à jour incluant `-RequirementRule` ignorait silencieusement aussi `Notes`, `DetectionRule` et `RestartBehavior`, pas seulement l'exigence d'architecture. `Add-IntuneWin32App` (la création) n'a pas ce bug.

Corrigé de la même manière : la branche de mise à jour ne transmet plus `-RequirementRule` à `Set-IntuneWin32App`, et appelle à la place `Set-Win32AppArchitectureRequirement` — un PATCH Graph direct pour `allowedArchitectures`/`minimumSupportedWindowsRelease`, qui réutilise la session déjà établie par `Connect-MSIntuneGraph`.

## Visibilité dans le Portail d'entreprise

`-CompanyPortalFeaturedApp $true` est défini à chaque exécution (à la création comme lors des mises à jour), de sorte que l'application apparaît mise en avant dans le Portail d'entreprise au lieu de rester invisible comme simple prérequis d'arrière-plan. Il n'y a pas de MSIX dont extraire un logo ; l'icône par défaut des applications Win32 d'Intune est donc utilisée — définissez ensuite `-Icon` manuellement dans le portail Intune si vous voulez une icône personnalisée.

## Deploy-CoworkPrerequisitesIntune.ps1

À exécuter chaque mois, ou dès qu'un script de contenu a changé :

```powershell
.\Deploy-CoworkPrerequisitesIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop"
```

Même comportement de confirmation/`-Force` que `Deploy-ClaudeDesktopIntune.ps1`. Comme il n'y a pas de MSIX dont la version augmente, une nouvelle exécution ne pousse une mise à jour que si vous avez modifié l'un des trois scripts de contenu — sinon, elle ne fait rien.

| Paramètre | Valeur par défaut | Description |
|---|---|---|
| `-AssignmentGroupName` | *(obligatoire)* | Groupe Entra ID attribué en Required. Utilisé uniquement à la création. |
| `-WorkingDirectory` | `C:\Temp\CoworkPrereqDeploy` | Dossier de préparation de la construction (`Source/`, `Output/`) |
| `-AppDisplayName` | `Cowork Windows Prerequisites (Machine-wide)` | Sert à retrouver l'application existante lors des exécutions suivantes — ne le modifiez pas sans renommer aussi l'application dans Intune |
| `-MinimumSupportedWindowsRelease` | `W10_21H2` | Règle de configuration requise |
| `-TenantId` | détecté automatiquement | ID du tenant Entra ID |
| `-IntuneWinAppUtilPath` | téléchargement automatique | Utiliser un `IntuneWinAppUtil.exe` déjà téléchargé |
| `-Force` | désactivé | Ignorer la ou les invites de confirmation |

## Uninstall-CoworkPrerequisites-Intune.ps1

Commande de désinstallation de l'application Win32.

| Paramètre | Obligatoire | Description |
|-----------|-------------|-------------|
| `-DisableVirtualMachinePlatform` | Non | Désactive aussi la fonctionnalité Windows `VirtualMachinePlatform` — uniquement si rien d'autre sur l'appareil n'en a besoin |

`VirtualMachinePlatform` n'est **pas** désactivé par défaut (d'autres applications — WSL, outils basés sur Hyper-V, autres applications de type Cowork — peuvent aussi en dépendre). Passez `-DisableVirtualMachinePlatform` au script de désinstallation si vous êtes certain que rien d'autre sur l'appareil n'en a besoin. Le démarrage rapide reste désactivé dans tous les cas — c'est un paramètre anodin, à l'échelle de l'appareil, et non quelque chose de propre à Cowork.

## Prérequis

- Les modules PowerShell `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications`, `Microsoft.Graph.Groups` et `IntuneWin32App` — à installer avec `.\scripts\Startup\Install-Modules.ps1`
- Exécution depuis Windows (l'outil d'empaquetage et les cmdlets DISM n'existent que sous Windows)

## CoworkPrerequisites-PlatformScript.ps1

`CoworkPrerequisites-PlatformScript.ps1` reprend la même logique d'activation de VMP et de désactivation du démarrage rapide, adaptée pour être chargée directement comme **Platform script** Intune (Devices → Scripts and remediations → Platform scripts) au lieu de passer par la mécanique d'application Win32 ci-dessus. Pas d'empaquetage `.intunewin`, pas de règle de détection/configuration requise, pas de `Deploy-*.ps1` — chargez simplement ce fichier unique.

**Pourquoi cette variante existe à côté de l'application Win32 :** la voie application Win32 s'est heurtée à de vraies difficultés en pratique (le bug `@odata.type` du module `IntuneWin32App`, des blocages GRS affectant tout l'appareil, une coquille dans `RestartBehavior`) — un simple script de plateforme contourne entièrement toute cette mécanique propre aux applications Win32. En contrepartie, on perd les deux atouts qu'offraient respectivement les applications Win32 et les Remediations :

| | Application Win32 (`Deploy-CoworkPrerequisitesIntune.ps1` de ce dossier) | Script de plateforme (`CoworkPrerequisites-PlatformScript.ps1`) |
|---|---|---|
| Redémarrage après activation de VMP | `-RestartBehavior 'basedOnReturnCode'` + exit `3010` → **Intune lui-même** impose un redémarrage (échéance/délai de grâce), aucune action de l'utilisateur requise au-delà de l'invite | Aucun mécanisme de ce type n'existe pour les scripts de plateforme — le code de sortie n'est que succès(`0`)/échec(tout le reste), ce script se termine donc toujours avec `0` et n'envoie que la notification `msg.exe` unique (via la même astuce de tâche planifiée dans la session de l'utilisateur que l'application Win32, afin que le bouton OK fonctionne vraiment) ; l'utilisateur doit redémarrer entièrement de sa propre initiative |
| Nouvelle vérification après un échec | La règle de détection est réévaluée à chaque check-in | S'exécute une fois par appareil par défaut ; une exécution « en échec » est bien relancée lors des check-ins suivants, mais une exécution *réussie* (VMP activé, redémarrage encore en attente) n'est pas revérifiée ensuite comme le ferait le Detect quotidien d'une Proactive Remediation |

Les Proactive Remediations (Detect + Remediate, relancées chaque jour) offriraient le meilleur des deux — mais cette fonctionnalité exige une licence Windows Enterprise/Education ou un VDA par utilisateur, **que Business Premium n'inclut pas** ; ce n'est donc pas une option ici.

**Déploiement :**
1. **Devices → Scripts and remediations → Platform scripts → Add → Windows 10 and later**
2. Chargez `CoworkPrerequisites-PlatformScript.ps1`
3. Paramètres du script : **Run this script using the logged on credentials** = No (SYSTEM) · **Enforce script signature check** = No · **Run script in 64-bit PowerShell Host** = Yes
4. Attribuez au même groupe d'appareils que Claude Desktop

Journalise dans `%ProgramData%\CoworkPrereqDeploy\platformscript.log` — un nom de fichier différent du `install.log` de la variante application Win32, afin que tester les deux sur le même appareil n'écrase aucun des deux journaux.
