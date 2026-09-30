[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Office365Toolkit](../readme.fr.md) › **Intune**

# Office365Toolkit / Intune

Inventaire à l'échelle du tenant des stratégies Intune / Endpoint Manager.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-IntunePolicyInventory.ps1`](Get-IntunePolicyInventory.ps1) ([docs](#get-intunepolicyinventoryps1)) | Inventaire de toutes les stratégies de conformité/configuration/protection des applications/Endpoint Security |

---

### Get-IntunePolicyInventory.ps1

Liste chaque stratégie sur les principales surfaces de stratégie Intune — stratégies de conformité
des appareils, profils de configuration des appareils, stratégies de configuration du Settings Catalog,
stratégies de protection des applications et stratégies Endpoint Security (« intents ») —
avec le nombre d'affectations. Un inventaire/une liste de contrôle rapide à un instant donné, pas une
comparaison avec une base de référence. Lecture seule.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine Entra ID |

**Exemples**

```powershell
.\Get-IntunePolicyInventory.ps1

.\Get-IntunePolicyInventory.ps1 -OutputPath C:\Reports
```

**Remarques**
- Pour comparer la configuration Intune d'un tenant client à une base de référence
  du MSP (détection de dérive), utilisez plutôt
  [`scripts/Intune/Compare-IntuneConfig.ps1`](../../Intune/readme.fr.md#compare-intuneconfigps1)
  — ce script effectue une comparaison complète à partir d'une sauvegarde ; celui-ci est un inventaire
  rapide de ce qui existe actuellement.

**Étendues requises :** `DeviceManagementConfiguration.Read.All`, `DeviceManagementApps.Read.All`
**Module requis :** `Microsoft.Graph.DeviceManagement`
