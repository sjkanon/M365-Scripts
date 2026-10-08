[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [PatronToolkit](../readme.fr.md) › **Intune**

# Patron Toolkit — Intune

Rapports d'affectation des stratégies Intune et inventaire des appareils Windows Autopilot via Microsoft
Graph.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-IntunePolicyAssignments.ps1`](Get-IntunePolicyAssignments.ps1) ([docs](#get-intunepolicyassignmentsps1)) | Rapport indiquant quels groupes sont affectés à quels profils/stratégies/applications Intune |
| [`Get-AutopilotDevices.ps1`](Get-AutopilotDevices.ps1) ([docs](#get-autopilotdevicesps1)) | Rapport des appareils Windows Autopilot inscrits et des profils de déploiement |

---

### Get-IntunePolicyAssignments.ps1

Énumère les profils de configuration des appareils, les stratégies de conformité, les stratégies du Settings Catalog
et les applications mobiles Intune, en résolvant les affectations de chaque objet en noms de groupes lisibles
(ou « All users »/« All devices »), et en indiquant si l'affectation est Include ou Exclude ; les lignes
d'applications indiquent aussi l'intention (required/available/uninstall). Une ligne CSV par paire
stratégie/affectation.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-PolicyType` | Non | `DeviceConfiguration`, `CompliancePolicy`, `SettingsCatalog`, `MobileApp` (par défaut : tous) |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine. Par défaut le client GDAP (`load.config.ps1`) ou votre propre tenant ; obligatoire en app-only |
| `-ClientId` | Non | Inscription d'application pour la connexion app-only (avec `-CertificateThumbprint`). Sans lui, le script se connecte en délégué, en votre nom |
| `-CertificateThumbprint` | Non | Empreinte du certificat pour la connexion app-only avec `-ClientId` |
| `-AppOnly` | Non | Connexion app-only avec le ClientId et l'empreinte du tenant depuis `graph.appid.json` |

**Exemples**

```powershell
.\Get-IntunePolicyAssignments.ps1

.\Get-IntunePolicyAssignments.ps1 -PolicyType CompliancePolicy,SettingsCatalog
```

**Remarques**
- Lecture seule. Les modifications d'affectation en masse n'ont volontairement pas été scriptées — examinez la sortie
  de ce rapport et modifiez les affectations stratégie par stratégie dans le centre d'administration Intune
- Tous les appels sont des `Invoke-MgGraphRequest` avec `$expand=assignments` et pagination
  via `@odata.nextLink`. Les stratégies du Settings Catalog n'existent que dans Graph **beta**
  (`/beta/deviceManagement/configurationPolicies`) ; la version précédente appelait
  `Get-MgDeviceManagementConfigurationPolicy`, absente du SDK Microsoft.Graph v2, et ne
  signalait silencieusement aucune stratégie Settings Catalog. Un type illisible produit
  désormais un avertissement visible au lieu d'un résultat vide
- Connexion via [`Connect-M365.ps1`](../../Startup/readme.fr.md#connect-m365ps1) : Microsoft
  Graph, en délégué par défaut (étendues `DeviceManagementConfiguration.Read.All`,
  `DeviceManagementApps.Read.All`, `Group.Read.All` plus un rôle Intune) ; app-only avec
  `-ClientId` + `-CertificateThumbprint` ou `-AppOnly` (les mêmes autorisations d'application)

---

### Get-AutopilotDevices.ps1

Liste chaque identité d'appareil Windows Autopilot inscrite dans le tenant (numéro de série,
modèle, fabricant, group tag, état d'inscription, état d'affectation du profil de déploiement),
ainsi qu'un récapitulatif des profils de déploiement existants. Il s'agit d'un rapport d'inventaire
côté tenant — à ne pas confondre avec
[`Get-Autopilot/Get-WindowsAutoPilotInfo.ps1`](../../Intune/readme.fr.md), qui collecte le
hachage matériel d'un appareil physique pour l'inscrire.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-GroupTag` | Non | Ne rapporter que les appareils portant ce group tag |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine. Par défaut le client GDAP (`load.config.ps1`) ou votre propre tenant ; obligatoire en app-only |
| `-ClientId` | Non | Inscription d'application pour la connexion app-only (avec `-CertificateThumbprint`). Sans lui, le script se connecte en délégué, en votre nom |
| `-CertificateThumbprint` | Non | Empreinte du certificat pour la connexion app-only avec `-ClientId` |
| `-AppOnly` | Non | Connexion app-only avec le ClientId et l'empreinte du tenant depuis `graph.appid.json` |

**Exemples**

```powershell
.\Get-AutopilotDevices.ps1

.\Get-AutopilotDevices.ps1 -GroupTag "Finance-Laptops"
```

**Remarques**
- Lecture seule. L'import/l'affectation/la suppression d'appareils en masse n'ont volontairement pas été scriptés — utilisez
  `Get-WindowsAutoPilotInfo.ps1` (déjà présent dans ce dépôt) et le centre d'administration Intune pour
  l'inscription, et examinez ce rapport avant toute réaffectation en masse
- Les profils de déploiement n'existent que dans Graph **beta**
  (`/beta/deviceManagement/windowsAutopilotDeploymentProfiles`) et sont lus avec
  `Invoke-MgGraphRequest` ; la version précédente utilisait une cmdlet absente du SDK
  Microsoft.Graph v2 et affichait toujours zéro profil. Les appareils viennent de Graph v1.0
- Le décompte « sans profil de déploiement affecté » considère désormais `assignedInSync`,
  `assignedOutOfSync` et `assignedUnkownSyncState` comme affectés — Graph n'a pas de valeur
  `assigned` simple, chaque appareil était donc compté comme non affecté
- Connexion via [`Connect-M365.ps1`](../../Startup/readme.fr.md#connect-m365ps1) : Microsoft
  Graph, en délégué par défaut (étendue `DeviceManagementServiceConfig.Read.All` plus un
  rôle Intune) ; app-only avec `-ClientId` + `-CertificateThumbprint` ou `-AppOnly`

**Modules requis**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```
