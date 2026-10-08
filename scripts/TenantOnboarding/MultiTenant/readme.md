**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [TenantOnboarding](../readme.md) › **MultiTenant**

# MultiTenant

MSP-wide scripts that operate across every GDAP (Granular Delegated Admin Privileges) customer tenant at once — license reporting, break-glass password rotation, and a quick-links portal index. These are the Microsoft Graph-based replacements for a family of legacy scripts that looped `Get-MsolPartnerContract -All` using the now-retired MSOnline module.

**Sign-in** goes through [`Connect-M365.ps1`](../../Startup/Connect-M365.ps1) and is **delegated by default**:

- **Delegated (GDAP, default)** — you sign in once as a partner admin in your own tenant, the script lists the customers from `tenantRelationships/delegatedAdminCustomers` (or `/contracts` when that list cannot be read), and then connects to each customer tenant as you, through GDAP. Your GDAP relationship must contain a role that allows the work (see each script). Every customer is a separate sign-in: in the browser it usually completes from the cached account, with device code (`$global:useDeviceCodeAuth`) you enter a code per customer. The first time, *Microsoft Graph Command Line Tools* may need consent in a customer tenant.
- **App-only (option)** — `-ClientId` with `-CertificateThumbprint` (preferred) or `-ClientSecret`, plus `-TenantId` (your partner tenant), or `-AppOnly` to read them from `graph.appid.json`. This needs a multi-tenant app registration of your own that is **consented in every customer tenant**. GDAP alone does not give an app access to customers: GDAP grants delegated rights to your users only.

Without `-TenantId`, the delegated partner sign-in goes to the tenant of the account you use, even when `load.config.ps1` has a GDAP customer selected.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-MultiTenantLicenseReport.ps1`](Get-MultiTenantLicenseReport.ps1) ([docs](#get-multitenantlicensereportps1)) | CSV report of licensed users across all (or selected) customer tenants |
| [`Update-BreakGlassAdminPassword.ps1`](Update-BreakGlassAdminPassword.ps1) ([docs](#update-breakglassadminpasswordps1)) | Rotate a break-glass account's password in one tenant or across all customers |
| [`New-CustomerPortalIndex.ps1`](New-CustomerPortalIndex.ps1) ([docs](#new-customerportalindexps1)) | Generate an HTML index of admin-portal quick links per customer tenant |

---

### Get-MultiTenantLicenseReport.ps1

Iterates every (or selected) GDAP customer tenant and exports licensed users + assigned SKUs to one CSV.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-TenantId` | App-only | Your partner (home) tenant. Delegated: defaults to the tenant of the account you sign in with |
| `-ClientId` | No | Multi-tenant app registration for app-only sign-in; omit for delegated |
| `-CertificateThumbprint` | * | Certificate thumbprint for `-ClientId` (preferred) |
| `-ClientSecret` | * | Client secret for `-ClientId` |
| `-AppOnly` | No | App-only with ClientId and thumbprint from `graph.appid.json` (partner tenant's entry) |
| `-CustomerTenantId` | No | Restrict to specific customer tenant IDs |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\`) |

*With `-ClientId`, one of `-CertificateThumbprint` / `-ClientSecret` is required.

**Examples**
```powershell
# Delegated: sign in as a partner admin
.\Get-MultiTenantLicenseReport.ps1

# App-only with a certificate
.\Get-MultiTenantLicenseReport.ps1 -TenantId "partner.onmicrosoft.com" `
    -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
```

**Notes**
- Delegated: the GDAP relationship needs a role that can read users (Global Reader, Directory Readers or User Administrator); scope `User.Read.All` per customer.
- App-only: `User.Read.All` as an application permission, consented in each customer tenant.
- A customer that cannot be reached is reported as `[WARN] Skipped` and the report continues.

---

### Update-BreakGlassAdminPassword.ps1

Rotates the password of a break-glass account, either in a single tenant or across every GDAP customer tenant (account `<local part>@<initial onmicrosoft.com domain>`). New passwords are printed once, never saved to disk.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-UserPrincipalName` | * | Full UPN, single-tenant mode |
| `-UserPrincipalNameLocalPart` | * | UPN local part, used with `-AllCustomers` |
| `-AllCustomers` | No | Rotate across every GDAP customer tenant |
| `-TenantId` | No | Single tenant: the tenant (default: GDAP customer, else your own). `-AllCustomers`: your partner tenant (required for app-only) |
| `-ClientId` / `-CertificateThumbprint` / `-ClientSecret` | No | App-only sign-in; omit for delegated |
| `-AppOnly` | No | App-only with ClientId and thumbprint from `graph.appid.json` |
| `-PasswordLength` | No | Default `24` |
| `-Apply` | No | Actually rotate (default: preview only) |

**Examples**
```powershell
# Single tenant, delegated (reuses a session for that tenant with User.ReadWrite.All)
.\Update-BreakGlassAdminPassword.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com" -Apply

# Every GDAP customer, delegated as the partner admin
.\Update-BreakGlassAdminPassword.ps1 -UserPrincipalNameLocalPart "breakglass-admin" -AllCustomers -Apply

# Every GDAP customer, app-only
.\Update-BreakGlassAdminPassword.ps1 -UserPrincipalNameLocalPart "breakglass-admin" -AllCustomers `
    -TenantId "partner.onmicrosoft.com" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" -Apply
```

**Notes**
- Resetting an administrator's password needs Privileged Authentication Administrator (or Global Administrator): in the GDAP relationship for delegated use, or assigned to the app's service principal in each customer for app-only (with `User.ReadWrite.All` and `Domain.Read.All` application permissions).
- The account is looked up on the tenant's **initial** domain (`isInitial`); earlier versions took the first `*.onmicrosoft.com` domain, which can be `contoso.mail.onmicrosoft.com`.
- The single-tenant mode no longer reuses just any Graph session: it must be for the right tenant and hold `User.ReadWrite.All`.

---

### New-CustomerPortalIndex.ps1

Builds a single searchable HTML page with per-customer deep links to the M365 Admin Center, Entra ID, Exchange Admin Center, Teams Admin Center, and Intune. No branding is hardcoded — set `-Title` to your own. Only the partner tenant is read, so one sign-in is enough.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-TenantId` | App-only | Your partner (home) tenant. Delegated: defaults to the tenant of the account you sign in with |
| `-ClientId` | No | App registration for app-only sign-in; omit for delegated |
| `-CertificateThumbprint` / `-ClientSecret` | * | One is required with `-ClientId` |
| `-AppOnly` | No | App-only with ClientId and thumbprint from `graph.appid.json` |
| `-Title` | No | Page title (default: `Customer Portal Index`) |
| `-OutputPath` | No | HTML output path (default: `C:\Temp\CustomerPortalIndex.html`) |

**Examples**
```powershell
# Delegated: sign in as a partner admin
.\New-CustomerPortalIndex.ps1 -Title "Our Customers"

.\New-CustomerPortalIndex.ps1 -TenantId "partner.onmicrosoft.com" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" -Title "Our Customers"
```

**Notes**
- The table is now written as HTML directly: `ConvertTo-Html` encoded the cell values, so the links appeared as literal `<a href=...>` text. Customer names are HTML-encoded, and the search box no longer hides the header row.

---

**Required Graph permission in the partner tenant:** `DelegatedAdminRelationship.Read.All` (delegated scope or application permission); `Directory.Read.All` for the `/contracts` fallback
**Required module:** `Microsoft.Graph.Authentication`
