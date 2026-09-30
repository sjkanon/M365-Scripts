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
(ou « All users »/« All devices »), et en indiquant si l'affectation est Include ou Exclude. Une ligne CSV
par paire stratégie/affectation.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-PolicyType` | Non | `DeviceConfiguration`, `CompliancePolicy`, `SettingsCatalog`, `MobileApp` (par défaut : tous) |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine Entra ID |

**Exemples**

```powershell
.\Get-IntunePolicyAssignments.ps1

.\Get-IntunePolicyAssignments.ps1 -PolicyType CompliancePolicy,SettingsCatalog
```

**Remarques**
- Lecture seule. Les modifications d'affectation en masse n'ont volontairement pas été scriptées — examinez la sortie
  de ce rapport et modifiez les affectations stratégie par stratégie dans le centre d'administration Intune
- Étendues requises : `DeviceManagementConfiguration.Read.All`,
  `DeviceManagementApps.Read.All`, `Group.Read.All`

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
| `-TenantId` | Non | ID de tenant ou domaine Entra ID |

**Exemples**

```powershell
.\Get-AutopilotDevices.ps1

.\Get-AutopilotDevices.ps1 -GroupTag "Finance-Laptops"
```

**Remarques**
- Lecture seule. L'import/l'affectation/la suppression d'appareils en masse n'ont volontairement pas été scriptés — utilisez
  `Get-WindowsAutoPilotInfo.ps1` (déjà présent dans ce dépôt) et le centre d'administration Intune pour
  l'inscription, et examinez ce rapport avant toute réaffectation en masse
- Étendue requise : `DeviceManagementServiceConfig.Read.All`

**Modules requis**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```
