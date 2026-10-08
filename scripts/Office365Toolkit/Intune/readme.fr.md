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

Se connecte via [`Connect-M365.ps1`](../../Startup/readme.fr.md#connect-m365ps1) : délégué par
défaut (vous vous connectez en tant qu'administrateur), avec le code d'appareil et le client GDAP
issus de `load.config.ps1` ; app-only avec `-ClientId` + `-CertificateThumbprint`, ou `-AppOnly`
(`graph.appid.json`). Une session Graph déjà adaptée est réutilisée et reste connectée.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-OutputPath` | Non | Chemin du fichier de rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine ; par défaut le client GDAP de `load.config.ps1` |
| `-ClientId` | Non | Inscription d'application pour la connexion app-only (avec `-CertificateThumbprint`) |
| `-CertificateThumbprint` | Non | Empreinte du certificat pour la connexion app-only |
| `-AppOnly` | Non | App-only avec ClientId et empreinte issus de `graph.appid.json` |

**Exemples**

```powershell
.\Get-IntunePolicyInventory.ps1

.\Get-IntunePolicyInventory.ps1 -OutputPath C:\Reports\intune.csv

.\Get-IntunePolicyInventory.ps1 -TenantId contoso.onmicrosoft.com -AppOnly
```

**Remarques**
- Settings Catalog (`configurationPolicies`) et Endpoint Security (`intents`) n'existent que sur
  le point de terminaison bêta de Graph. Le SDK v1.0 n'a pas d'applets de commande pour eux ; les
  versions précédentes interceptaient l'erreur et les omettaient silencieusement. Les cinq surfaces
  sont désormais lues avec `Invoke-MgGraphRequest`, en suivant `@odata.nextLink`.
- Une stratégie sans affectation affiche maintenant `0` au lieu d'un nombre vide. Les stratégies de
  protection des applications n'ont pas d'affectations sur le type de base `managedAppPolicy` ;
  leur nombre reste donc vide.
- Pour comparer la configuration Intune d'un tenant client à une base de référence
  du MSP (détection de dérive), utilisez plutôt
  [`scripts/Intune/Compare-IntuneConfig.ps1`](../../Intune/readme.fr.md#compare-intuneconfigps1)
  — ce script effectue une comparaison complète à partir d'une sauvegarde ; celui-ci est un inventaire
  rapide de ce qui existe actuellement.

**Étendues requises :** `DeviceManagementConfiguration.Read.All`, `DeviceManagementApps.Read.All`
**Module requis :** `Microsoft.Graph.Authentication`
