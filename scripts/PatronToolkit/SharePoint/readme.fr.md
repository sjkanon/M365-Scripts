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

Rapporte les paramètres de partage externe à l'échelle du tenant qui comptent le plus lors d'une
revue de sécurité, lus dans Microsoft Graph (`GET /admin/sharepoint/settings`) : capacité de
partage, repartage par les utilisateurs externes, obligation que le compte qui accepte corresponde
au compte invité, mode d'autorisation/blocage des domaines, authentification héritée et
déconnexion en cas d'inactivité. En option, via PnP.PowerShell, aussi le type de lien / l'autorisation
de lien par défaut / l'expiration des liens anonymes, les sites dont le partage est plus large
que la valeur par défaut du tenant, et chaque utilisateur externe (invité) sur toutes les
collections de sites.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-TenantName` | Non | Nom du tenant SharePoint, p. ex. `contoso` pour `https://contoso-admin.sharepoint.com`. Uniquement pour la partie PnP ; active aussi `-IncludeLinkSettings` |
| `-AdminUrl` | Non | URL complète du centre d'administration SharePoint (alternative à `-TenantName`). Sans l'un ni l'autre, l'URL est recherchée via Graph (`/sites/root`) |
| `-IncludeLinkSettings` | Non | Rapporter aussi le type de lien par défaut, l'autorisation de lien par défaut et l'expiration des liens anonymes (PnP) |
| `-IncludeExternalUsers` | Non | Énumérer aussi les utilisateurs externes sur tous les sites (PnP, plus lent) |
| `-IncludeSiteOverrides` | Non | Signaler aussi les sites dont le partage est plus large que la valeur par défaut (PnP) |
| `-OutputPath` | Non | Dossier du ou des rapports CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine. Par défaut le client GDAP (`load.config.ps1`) ou votre propre tenant ; obligatoire en app-only |
| `-ClientId` | Non | Inscription d'application pour la connexion app-only (avec `-CertificateThumbprint`), pour Graph et PnP. Sans lui, le script se connecte en délégué, en votre nom |
| `-CertificateThumbprint` | Non | Empreinte du certificat pour la connexion app-only avec `-ClientId` |
| `-AppOnly` | Non | Connexion app-only (Graph et PnP) avec le ClientId et l'empreinte du tenant depuis `graph.appid.json` |
| `-PnPClientId` | Non | Inscription d'application PnP pour la connexion PnP déléguée (par défaut : l'entrée du tenant dans `pnp.appid.json`) |

**Exemples**

```powershell
# Graph uniquement : paramètres de partage du tenant
.\Test-SharePointSharingConfig.ps1

# Comme avant : paramètres du tenant plus paramètres de lien (PnP)
.\Test-SharePointSharingConfig.ps1 -TenantName "contoso"

.\Test-SharePointSharingConfig.ps1 -IncludeLinkSettings -IncludeExternalUsers -IncludeSiteOverrides

# App-only, Graph et PnP avec l'application de graph.appid.json
.\Test-SharePointSharingConfig.ps1 -TenantId contoso.onmicrosoft.com -AppOnly -IncludeSiteOverrides
```

**Remarques**
- L'ancien SharePoint Online Management Shell (`Connect-SPOService`, `Get-SPOTenant`,
  `Get-SPOSite`, `Get-SPOExternalUser`) n'est plus utilisé. L'ancien `-TenantId` était
  documenté mais jamais transmis
- Graph n'a pas d'API pour les valeurs par défaut des liens, la capacité de partage par site
  ni les utilisateurs externes par site ; ces parties restent donc sur PnP (`Get-PnPTenant`,
  `Get-PnPTenantSite`, `Get-PnPExternalUser`) et ne s'exécutent que sur demande. Si la
  connexion PnP échoue, elles sont ignorées avec un avertissement et la partie Graph est tout
  de même rapportée
- Graph renvoie `sharingCapability` en camelCase (`externalUserAndGuestSharing`) et l'inverse
  de l'ancien paramètre : `ResharingByExternalUsersEnabled` au lieu de
  `PreventExternalUsersFromResharing`
- Connexion via [`Connect-M365.ps1`](../../Startup/readme.fr.md#connect-m365ps1) : Graph en
  délégué par défaut (étendue `SharePointTenantSettings.Read.All`, plus `Sites.Read.All`
  uniquement pour rechercher l'URL d'administration ; rôle SharePoint Administrator) ;
  app-only avec `-ClientId` + `-CertificateThumbprint` ou `-AppOnly`. Depuis septembre 2024,
  PnP exige sa propre inscription d'application : `-PnPClientId` ou `pnp.appid.json` en
  délégué, la même application que Graph en app-only (SharePoint `Sites.FullControl.All`)
- Pour les rapports de stockage/versions SharePoint, voir
  [`Reporting/Get-SharePointStorageReport.ps1`](../../Reporting/readme.fr.md) — une fonctionnalité distincte,
  basée sur Graph, déjà présente dans ce dépôt

**Modules requis**
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module PnP.PowerShell -Scope CurrentUser   # uniquement pour la partie PnP
```
