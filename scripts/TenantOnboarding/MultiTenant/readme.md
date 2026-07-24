# MultiTenant

MSP-wide scripts that operate across every GDAP (Granular Delegated Admin Privileges) customer tenant at once — license reporting, break-glass password rotation, and a quick-links portal index. These are the Microsoft Graph-based replacements for a family of legacy scripts that looped `Get-MsolPartnerContract -All` using the now-retired MSOnline module.

All three scripts authenticate app-only against each customer tenant using a multi-tenant, GDAP-enabled App Registration you provide via `-ClientId` (+ `-ClientSecret` or `-CertificateThumbprint`) — they do not create or reuse an interactive Graph session.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-MultiTenantLicenseReport.ps1`](#get-multitenantlicensereportps1) | CSV report of licensed users across all (or selected) customer tenants |
| [`Update-BreakGlassAdminPassword.ps1`](#update-breakglassadminpasswordps1) | Rotate a break-glass account's password in one tenant or across all customers |
| [`New-CustomerPortalIndex.ps1`](#new-customerportalindexps1) | Generate an HTML index of admin-portal quick links per customer tenant |

---

### Get-MultiTenantLicenseReport.ps1

Iterates every (or selected) GDAP customer tenant and exports licensed users + assigned SKUs to one CSV.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-ClientId` | Yes | Multi-tenant, GDAP-enabled App Registration Client ID |
| `-ClientSecret` | * | Client secret for `-ClientId` |
| `-CertificateThumbprint` | * | Certificate thumbprint for `-ClientId` (preferred) |
| `-CustomerTenantId` | No | Restrict to specific customer tenant IDs |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\`) |

*One of `-ClientSecret` / `-CertificateThumbprint` is required.

**Example**
```powershell
.\Get-MultiTenantLicenseReport.ps1 -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
```

---

### Update-BreakGlassAdminPassword.ps1

Rotates the password of a break-glass account, either in a single already-connected tenant or across every GDAP customer tenant. New passwords are printed once, never saved to disk.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-UserPrincipalName` | * | Full UPN, single-tenant mode |
| `-UserPrincipalNameLocalPart` | * | UPN local part, used with `-AllCustomers` |
| `-AllCustomers` | No | Rotate across every GDAP customer tenant |
| `-ClientId` / `-ClientSecret` / `-CertificateThumbprint` | * | Required with `-AllCustomers` |
| `-PasswordLength` | No | Default `24` |
| `-Apply` | No | Actually rotate (default: preview only) |

**Examples**
```powershell
.\Update-BreakGlassAdminPassword.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com" -Apply

.\Update-BreakGlassAdminPassword.ps1 -UserPrincipalNameLocalPart "breakglass-admin" -AllCustomers `
    -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" -Apply
```

---

### New-CustomerPortalIndex.ps1

Builds a single searchable HTML page with per-customer deep links to the M365 Admin Center, Entra ID, Exchange Admin Center, Teams Admin Center, and Intune. No branding is hardcoded — set `-Title` to your own.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-ClientId` | Yes | Multi-tenant, GDAP-enabled App Registration Client ID |
| `-ClientSecret` / `-CertificateThumbprint` | * | One is required |
| `-Title` | No | Page title (default: `Customer Portal Index`) |
| `-OutputPath` | No | HTML output path (default: `C:\Temp\CustomerPortalIndex.html`) |

**Example**
```powershell
.\New-CustomerPortalIndex.ps1 -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" -Title "Our Customers"
```

---

**Required Graph permission on the partner/home tenant app:** `DelegatedAdminRelationship.Read.All`
**Required module:** `Microsoft.Graph.Authentication`
