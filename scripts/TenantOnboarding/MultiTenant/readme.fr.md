[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [TenantOnboarding](../readme.fr.md) › **MultiTenant**

# MultiTenant

Scripts à l'échelle du MSP qui agissent en une seule fois sur chaque tenant client GDAP (Granular Delegated Admin Privileges) : rapports de licences, rotation des mots de passe break-glass et un index de liens rapides vers les portails. Ce sont les remplaçants basés sur Microsoft Graph d'une famille de scripts hérités qui parcouraient `Get-MsolPartnerContract -All` avec le module MSOnline, désormais retiré.

Les trois scripts s'authentifient en mode application seule auprès de chaque tenant client, à l'aide d'une App Registration multi-tenant compatible GDAP que vous fournissez via `-ClientId` (+ `-ClientSecret` ou `-CertificateThumbprint`) ; ils ne créent ni ne réutilisent de session Graph interactive.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-MultiTenantLicenseReport.ps1`](Get-MultiTenantLicenseReport.ps1) ([docs](#get-multitenantlicensereportps1)) | Rapport CSV des utilisateurs sous licence dans tous les tenants clients (ou une sélection) |
| [`Update-BreakGlassAdminPassword.ps1`](Update-BreakGlassAdminPassword.ps1) ([docs](#update-breakglassadminpasswordps1)) | Faire tourner le mot de passe d'un compte break-glass dans un tenant ou chez tous les clients |
| [`New-CustomerPortalIndex.ps1`](New-CustomerPortalIndex.ps1) ([docs](#new-customerportalindexps1)) | Générer un index HTML de liens rapides vers les portails d'administration par tenant client |

---

### Get-MultiTenantLicenseReport.ps1

Parcourt chaque tenant client GDAP (ou une sélection) et exporte les utilisateurs sous licence + les SKU attribués dans un seul CSV.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-ClientId` | Oui | ID client de l'App Registration multi-tenant compatible GDAP |
| `-ClientSecret` | * | Secret client pour `-ClientId` |
| `-CertificateThumbprint` | * | Empreinte du certificat pour `-ClientId` (à privilégier) |
| `-CustomerTenantId` | Non | Limiter à des ID de tenants clients précis |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\`) |

*L'un des paramètres `-ClientSecret` / `-CertificateThumbprint` est obligatoire.

**Exemple**
```powershell
.\Get-MultiTenantLicenseReport.ps1 -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
```

---

### Update-BreakGlassAdminPassword.ps1

Fait tourner le mot de passe d'un compte break-glass, soit dans un seul tenant auquel vous êtes déjà connecté, soit dans chaque tenant client GDAP. Les nouveaux mots de passe sont affichés une seule fois et ne sont jamais enregistrés sur disque.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-UserPrincipalName` | * | UPN complet, mode tenant unique |
| `-UserPrincipalNameLocalPart` | * | Partie locale de l'UPN, utilisée avec `-AllCustomers` |
| `-AllCustomers` | Non | Faire tourner dans chaque tenant client GDAP |
| `-ClientId` / `-ClientSecret` / `-CertificateThumbprint` | * | Obligatoires avec `-AllCustomers` |
| `-PasswordLength` | Non | Par défaut `24` |
| `-Apply` | Non | Faire réellement tourner (par défaut : aperçu uniquement) |

**Exemples**
```powershell
.\Update-BreakGlassAdminPassword.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com" -Apply

.\Update-BreakGlassAdminPassword.ps1 -UserPrincipalNameLocalPart "breakglass-admin" -AllCustomers `
    -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" -Apply
```

---

### New-CustomerPortalIndex.ps1

Construit une seule page HTML consultable avec, pour chaque client, des liens directs vers le centre d'administration M365, Entra ID, le centre d'administration Exchange, le centre d'administration Teams et Intune. Aucune identité visuelle n'est codée en dur : définissez `-Title` avec votre propre titre.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-ClientId` | Oui | ID client de l'App Registration multi-tenant compatible GDAP |
| `-ClientSecret` / `-CertificateThumbprint` | * | L'un des deux est obligatoire |
| `-Title` | Non | Titre de la page (par défaut : `Customer Portal Index`) |
| `-OutputPath` | Non | Chemin de sortie HTML (par défaut : `C:\Temp\CustomerPortalIndex.html`) |

**Exemple**
```powershell
.\New-CustomerPortalIndex.ps1 -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" -Title "Our Customers"
```

---

**Autorisation Graph requise sur l'application du tenant partenaire/d'origine :** `DelegatedAdminRelationship.Read.All`
**Module requis :** `Microsoft.Graph.Authentication`
