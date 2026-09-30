[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [PatronToolkit](../readme.fr.md) › **SharePoint**

# Patron Toolkit — SharePoint

Configuration du partage du tenant SharePoint Online et audit des utilisateurs externes.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Test-SharePointSharingConfig.ps1`](Test-SharePointSharingConfig.ps1) ([docs](#test-sharepointsharingconfigps1)) | Rapport des paramètres de partage du tenant, des exceptions par site et des utilisateurs externes |

---

### Test-SharePointSharingConfig.ps1

Rapporte les paramètres de partage externe à l'échelle du tenant qui comptent le plus lors d'une revue de sécurité
(capacité de partage, type de lien par défaut, expiration/autorisation des liens anonymes, repartage par
les utilisateurs externes, authentification héritée). En option, énumère aussi chaque utilisateur externe (invité)
sur toutes les collections de sites et signale les sites dont la capacité de partage est
plus large que celle définie par défaut pour le tenant.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-TenantName` | * | Nom du tenant SharePoint, p. ex. `contoso` pour `https://contoso-admin.sharepoint.com` |
| `-AdminUrl` | * | URL complète du centre d'administration SharePoint (alternative à `-TenantName`) |
| `-IncludeExternalUsers` | Non | Énumérer aussi les utilisateurs externes sur tous les sites (plus lent) |
| `-IncludeSiteOverrides` | Non | Signaler aussi les sites dont le partage est plus large que la valeur par défaut |
| `-OutputPath` | Non | Dossier du ou des rapports CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | Transmis à `Connect-SPOService` si pris en charge |

*Obligatoire, sauf si vous êtes déjà connecté via `Connect-SPOService`.

**Exemples**

```powershell
.\Test-SharePointSharingConfig.ps1 -TenantName "contoso"

.\Test-SharePointSharingConfig.ps1 -TenantName "contoso" -IncludeExternalUsers -IncludeSiteOverrides
```

**Remarques**
- Pour les rapports de stockage/versions SharePoint, voir
  [`Reporting/Get-SharePointStorageReport.ps1`](../../Reporting/readme.fr.md) — une fonctionnalité distincte,
  basée sur Graph, déjà présente dans ce dépôt

**Module requis**
```powershell
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
```
