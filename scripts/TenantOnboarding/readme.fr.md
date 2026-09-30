[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **TenantOnboarding**

# Tenant Onboarding

Scripts pour l'intégration d'un nouveau tenant M365 et le provisionnement/la gestion de ses appareils, portés et modernisés à partir d'un dépôt hérité retiré. Chaque script suit le style maison de ce dépôt : essai à blanc par défaut avec un commutateur `-Apply` pour tout ce qui modifie l'état, `[CmdletBinding(SupportsShouldProcess)]`, et aucun identifiant, nom de tenant/client ou point de terminaison interne codé en dur.

Lorsque la source héritée comportait plusieurs scripts quasi identiques réalisant de légères variantes de la même tâche (six scripts de redémarrage de OneDrive, sept téléchargeurs d'installeurs fournisseurs codés en dur, quatre scripts d'octroi/de révocation de groupes locaux...), ils ont été regroupés en un seul script bien paramétré plutôt que portés un par un.

---

## Dossiers

| Dossier | Contenu |
|--------|----------|
| [`Provisioning/`](Provisioning/readme.fr.md) | Amorçage d'un tenant unique : compte administrateur break-glass, groupes de sécurité de base, attribution de la stratégie de base Intune |
| [`MultiTenant/`](MultiTenant/readme.fr.md) | Scripts à l'échelle du MSP, multi-clients (GDAP) : rapports de licences, rotation des mots de passe break-glass, index des portails clients |
| [`AppDeployment/`](AppDeployment/readme.fr.md) | Installeurs génériques Win32/Chocolatey, associations de fichiers par défaut, raccourcis bureau, imprimantes réseau |
| [`DeviceConfig/`](DeviceConfig/readme.fr.md) | Auto-élévation via groupes locaux, durcissement du stockage des identifiants, paramètres d'alimentation kiosque, suppression d'Office, disposition du menu Démarrer, règle de pare-feu Teams |
| [`OneDriveManagement/`](OneDriveManagement/readme.fr.md) | Surveillance du redémarrage/de la réinitialisation de OneDrive, arrêt de la synchronisation par bibliothèque, redirection Known Folder Move |
| [`UserManagement/`](UserManagement/readme.fr.md) | Création de groupes de distribution dynamiques par filtre, appartenance à des groupes d'activation de fonctionnalités |

Consultez le readme de chaque sous-dossier pour la liste complète des scripts, les paramètres et les exemples.

---

## Ce qui a été volontairement laissé de côté

- **Dépôt d'exemples Microsoft Entra ID / macOS Intune** (`Install New Tenant/MacOS/`) : une vaste collection embarquée de scripts shell, de profils `.mobileconfig` et d'installeurs tiers pour macOS. Il ne s'agit pas de l'outillage PowerShell propre à ce MSP, et cela sort du périmètre d'un portage au style maison PowerShell.
- **`Manage Tenant/Compare/Compare-Intune.ps1`** : vérifié comme étant la même fonctionnalité que celle déjà couverte par `scripts/Intune/Compare-IntuneConfig.ps1` (les deux encapsulent `Compare-IntuneBackupDirectories` du module `IntuneBackupAndRestore` pour comparer la configuration Intune d'un tenant à une référence) ; non reporté.
- Plusieurs scripts de rapport/configuration déjà couverts ailleurs dans ce dépôt : fond d'écran OneDrive/écran de verrouillage (`scripts/Intune/Desktop/Background/`), « épingler à l'écran Démarrer » (`scripts/Intune/Desktop/Add Lockscreen to start and desktop/`), synchronisation de l'heure (`scripts/Device/Time sync/`) et rapports du contrôleur UniFi (`scripts/Network/UniFi/`). Les versions héritées étaient soit identiques octet pour octet, soit supplantées par une version plus générale déjà présente dans ce dépôt.
- Les scripts reposant sur les modules retirés **MSOnline** / **AzureAD** / **AzureADPreview** n'ont pas été portés tels quels ; leurs fonctionnalités ont été réimplémentées avec Microsoft Graph ou Exchange Online (voir `MultiTenant/` et `Provisioning/`), ou abandonnées lorsque la technique sous-jacente ne s'applique plus.
- Une poignée de scripts ponctuels, à usage unique ou déjà obsolètes ont été abandonnés comme de véritables impasses : un correctif de registre de l'Explorateur Windows 11 propre aux builds inférieures à 25211 (corrigé depuis longtemps), deux expériences en double/défectueuses « définir l'application PDF par défaut + importer une disposition Démarrer sans rapport », un utilitaire trivial de fichier d'identifiants DPAPI local, et deux scripts de rapport tiers (non écrits par le MSP) dont la fonctionnalité est déjà couverte par `scripts/Exchange/Test-MailboxPermissions.ps1` et `scripts/Entra/Test-M365GroupMembership.ps1`.
- Le dossier source `Setup tenant/` était vide.

Aucun nom de client, domaine de tenant, nom d'hôte interne, adresse IP ou identifiant provenant du dépôt source n'a été reproduit dans ce dossier : chaque script ici est écrit pour des entrées génériques et paramétrées.
