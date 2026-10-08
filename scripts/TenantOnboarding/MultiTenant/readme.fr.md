[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [TenantOnboarding](../readme.fr.md) › **MultiTenant**

# MultiTenant

Scripts à l'échelle du MSP qui agissent en une seule fois sur chaque tenant client GDAP (Granular Delegated Admin Privileges) : rapports de licences, rotation des mots de passe break-glass et un index de liens rapides vers les portails. Ce sont les remplaçants basés sur Microsoft Graph d'une famille de scripts hérités qui parcouraient `Get-MsolPartnerContract -All` avec le module MSOnline, désormais retiré.

**La connexion** passe par [`Connect-M365.ps1`](../../Startup/Connect-M365.ps1) et est **déléguée par défaut** :

- **Déléguée (GDAP, par défaut)** : vous vous connectez une fois en tant qu'administrateur partenaire dans votre propre tenant, le script récupère les clients depuis `tenantRelationships/delegatedAdminCustomers` (ou `/contracts` si cette liste n'est pas lisible), puis se connecte à chaque tenant client en votre nom via GDAP. Votre relation GDAP doit contenir un rôle qui autorise le travail (voir chaque script). Chaque client est une connexion distincte : dans le navigateur, elle aboutit généralement avec le compte en cache ; avec le code d'appareil (`$global:useDeviceCodeAuth`), vous saisissez un code par client. La première fois, *Microsoft Graph Command Line Tools* peut nécessiter un consentement dans un tenant client.
- **Application seule (option)** : `-ClientId` avec `-CertificateThumbprint` (recommandé) ou `-ClientSecret`, plus `-TenantId` (votre tenant partenaire), ou `-AppOnly` pour les lire dans `graph.appid.json`. Il faut pour cela votre propre App Registration multi-tenant **consentie dans chaque tenant client**. GDAP seul ne donne pas à une application l'accès aux clients : GDAP n'accorde que des droits délégués à vos utilisateurs.

Sans `-TenantId`, la connexion partenaire déléguée se fait sur le tenant du compte utilisé, même si un client GDAP est sélectionné dans `load.config.ps1`.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-MultiTenantLicenseReport.ps1`](Get-MultiTenantLicenseReport.ps1) ([docs](#get-multitenantlicensereportps1)) | Rapport CSV des utilisateurs sous licence dans tous les tenants clients (ou une sélection) |
| [`Update-BreakGlassAdminPassword.ps1`](Update-BreakGlassAdminPassword.ps1) ([docs](#update-breakglassadminpasswordps1)) | Faire tourner le mot de passe d'un compte break-glass dans un tenant ou chez tous les clients |
| [`New-CustomerPortalIndex.ps1`](New-CustomerPortalIndex.ps1) ([docs](#new-customerportalindexps1)) | Générer un index HTML de liens rapides vers les portails d'administration par tenant client |

---

### Get-MultiTenantLicenseReport.ps1

Parcourt chaque tenant client GDAP (ou une sélection) et exporte les utilisateurs sous licence et leurs SKU attribués dans un seul fichier CSV.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-TenantId` | Application seule | Votre tenant partenaire (d'origine). Délégué : par défaut, le tenant du compte utilisé |
| `-ClientId` | Non | App Registration multi-tenant pour la connexion en application seule ; à omettre en délégué |
| `-CertificateThumbprint` | * | Empreinte du certificat pour `-ClientId` (recommandé) |
| `-ClientSecret` | * | Secret client pour `-ClientId` |
| `-AppOnly` | Non | Application seule avec le ClientId et l'empreinte de `graph.appid.json` (entrée du tenant partenaire) |
| `-CustomerTenantId` | Non | Limiter à des ID de tenants clients précis |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\`) |

*Avec `-ClientId`, l'un de `-CertificateThumbprint` / `-ClientSecret` est obligatoire.

**Exemples**
```powershell
# Délégué : connexion en tant qu'administrateur partenaire
.\Get-MultiTenantLicenseReport.ps1

# Application seule avec un certificat
.\Get-MultiTenantLicenseReport.ps1 -TenantId "partner.onmicrosoft.com" `
    -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
```

**Remarques**
- Délégué : la relation GDAP a besoin d'un rôle qui peut lire les utilisateurs (Global Reader, Directory Readers ou User Administrator) ; scope `User.Read.All` par client.
- Application seule : `User.Read.All` en autorisation d'application, consentie dans chaque tenant client.
- Un client injoignable est signalé par `[WARN] Skipped` et le rapport continue.

---

### Update-BreakGlassAdminPassword.ps1

Fait tourner le mot de passe d'un compte break-glass, dans un seul tenant ou dans chaque tenant client GDAP (compte `<partie locale>@<domaine onmicrosoft.com initial>`). Les nouveaux mots de passe sont affichés une seule fois et ne sont jamais enregistrés sur le disque.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-UserPrincipalName` | * | UPN complet, mode tenant unique |
| `-UserPrincipalNameLocalPart` | * | Partie locale de l'UPN, utilisée avec `-AllCustomers` |
| `-AllCustomers` | Non | Effectuer la rotation dans chaque tenant client GDAP |
| `-TenantId` | Non | Tenant unique : le tenant (par défaut : client GDAP, sinon le vôtre). `-AllCustomers` : votre tenant partenaire (obligatoire en application seule) |
| `-ClientId` / `-CertificateThumbprint` / `-ClientSecret` | Non | Connexion en application seule ; à omettre en délégué |
| `-AppOnly` | Non | Application seule avec le ClientId et l'empreinte de `graph.appid.json` |
| `-PasswordLength` | Non | `24` par défaut |
| `-Apply` | Non | Effectuer réellement la rotation (par défaut : aperçu uniquement) |

**Exemples**
```powershell
# Tenant unique, délégué (réutilise une session pour ce tenant avec User.ReadWrite.All)
.\Update-BreakGlassAdminPassword.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com" -Apply

# Chaque client GDAP, délégué en tant qu'administrateur partenaire
.\Update-BreakGlassAdminPassword.ps1 -UserPrincipalNameLocalPart "breakglass-admin" -AllCustomers -Apply

# Chaque client GDAP, application seule
.\Update-BreakGlassAdminPassword.ps1 -UserPrincipalNameLocalPart "breakglass-admin" -AllCustomers `
    -TenantId "partner.onmicrosoft.com" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" -Apply
```

**Remarques**
- Réinitialiser le mot de passe d'un administrateur exige Privileged Authentication Administrator (ou Global Administrator) : dans la relation GDAP en délégué, ou attribué au principal de service de l'application dans chaque client en application seule (avec les autorisations d'application `User.ReadWrite.All` et `Domain.Read.All`).
- Le compte est recherché sur le domaine **initial** du tenant (`isInitial`) ; les versions précédentes prenaient le premier domaine `*.onmicrosoft.com`, qui peut être `contoso.mail.onmicrosoft.com`.
- Le mode tenant unique ne réutilise plus n'importe quelle session Graph : elle doit porter sur le bon tenant et disposer de `User.ReadWrite.All`.

---

### New-CustomerPortalIndex.ps1

Construit une page HTML unique et consultable avec, pour chaque client, des liens directs vers le centre d'administration M365, Entra ID, le centre d'administration Exchange, le centre d'administration Teams et Intune. Aucune identité visuelle n'est codée en dur ; définissez `-Title` avec votre propre titre. Seul le tenant partenaire est lu, une seule connexion suffit donc.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-TenantId` | Application seule | Votre tenant partenaire (d'origine). Délégué : par défaut, le tenant du compte utilisé |
| `-ClientId` | Non | App Registration pour la connexion en application seule ; à omettre en délégué |
| `-CertificateThumbprint` / `-ClientSecret` | * | L'un des deux est obligatoire avec `-ClientId` |
| `-AppOnly` | Non | Application seule avec le ClientId et l'empreinte de `graph.appid.json` |
| `-Title` | Non | Titre de la page (par défaut : `Customer Portal Index`) |
| `-OutputPath` | Non | Chemin de sortie HTML (par défaut : `C:\Temp\CustomerPortalIndex.html`) |

**Exemples**
```powershell
# Délégué : connexion en tant qu'administrateur partenaire
.\New-CustomerPortalIndex.ps1 -Title "Our Customers"

.\New-CustomerPortalIndex.ps1 -TenantId "partner.onmicrosoft.com" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" -Title "Our Customers"
```

**Remarques**
- Le tableau est désormais écrit directement en HTML : `ConvertTo-Html` encodait les valeurs des cellules, si bien que les liens s'affichaient sous forme de texte littéral `<a href=...>`. Les noms des clients sont encodés en HTML et la zone de recherche ne masque plus la ligne d'en-tête.

---

**Autorisation Graph requise dans le tenant partenaire :** `DelegatedAdminRelationship.Read.All` (scope délégué ou autorisation d'application) ; `Directory.Read.All` pour le repli sur `/contracts`
**Module requis :** `Microsoft.Graph.Authentication`
