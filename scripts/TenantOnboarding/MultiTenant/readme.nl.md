[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [TenantOnboarding](../readme.nl.md) › **MultiTenant**

# MultiTenant

MSP-brede scripts die in één keer over elke GDAP-klanttenant (Granular Delegated Admin Privileges) heen werken: licentierapportage, rotatie van break-glass-wachtwoorden en een index met snelkoppelingen naar portalen. Dit zijn de op Microsoft Graph gebaseerde vervangers van een familie legacy-scripts die met de inmiddels uitgefaseerde MSOnline-module door `Get-MsolPartnerContract -All` liepen.

Alle drie de scripts authenticeren app-only bij elke klanttenant met een multi-tenant App Registration met GDAP die je opgeeft via `-ClientId` (+ `-ClientSecret` of `-CertificateThumbprint`); ze maken geen interactieve Graph-sessie aan en hergebruiken er ook geen.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Get-MultiTenantLicenseReport.ps1`](Get-MultiTenantLicenseReport.ps1) ([docs](#get-multitenantlicensereportps1)) | CSV-rapport van gelicentieerde gebruikers in alle (of geselecteerde) klanttenants |
| [`Update-BreakGlassAdminPassword.ps1`](Update-BreakGlassAdminPassword.ps1) ([docs](#update-breakglassadminpasswordps1)) | Roteer het wachtwoord van een break-glass-account in één tenant of bij alle klanten |
| [`New-CustomerPortalIndex.ps1`](New-CustomerPortalIndex.ps1) ([docs](#new-customerportalindexps1)) | Genereer een HTML-index met snelkoppelingen naar beheerportalen per klanttenant |

---

### Get-MultiTenantLicenseReport.ps1

Loopt door elke (of een selectie van) GDAP-klanttenant(s) en exporteert gelicentieerde gebruikers + toegewezen SKU's naar één CSV.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-ClientId` | Ja | Client-ID van de multi-tenant App Registration met GDAP |
| `-ClientSecret` | * | Client secret voor `-ClientId` |
| `-CertificateThumbprint` | * | Certificaatvingerafdruk voor `-ClientId` (voorkeur) |
| `-CustomerTenantId` | Nee | Beperk tot specifieke klanttenant-ID's |
| `-OutputPath` | Nee | Pad voor het CSV-rapport (standaard: `C:\Temp\`) |

*Een van `-ClientSecret` / `-CertificateThumbprint` is verplicht.

**Voorbeeld**
```powershell
.\Get-MultiTenantLicenseReport.ps1 -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
```

---

### Update-BreakGlassAdminPassword.ps1

Roteert het wachtwoord van een break-glass-account, in één tenant waarmee je al verbonden bent of in elke GDAP-klanttenant. Nieuwe wachtwoorden worden één keer getoond en nooit op schijf opgeslagen.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-UserPrincipalName` | * | Volledige UPN, modus voor één tenant |
| `-UserPrincipalNameLocalPart` | * | Lokaal deel van de UPN, gebruikt met `-AllCustomers` |
| `-AllCustomers` | Nee | Roteer in elke GDAP-klanttenant |
| `-ClientId` / `-ClientSecret` / `-CertificateThumbprint` | * | Verplicht met `-AllCustomers` |
| `-PasswordLength` | Nee | Standaard `24` |
| `-Apply` | Nee | Roteer echt (standaard: alleen voorbeeldweergave) |

**Voorbeelden**
```powershell
.\Update-BreakGlassAdminPassword.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com" -Apply

.\Update-BreakGlassAdminPassword.ps1 -UserPrincipalNameLocalPart "breakglass-admin" -AllCustomers `
    -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" -Apply
```

---

### New-CustomerPortalIndex.ps1

Bouwt één doorzoekbare HTML-pagina met per klant directe links naar het M365-beheercentrum, Entra ID, het Exchange-beheercentrum, het Teams-beheercentrum en Intune. Er is geen huisstijl hardcoded; stel `-Title` in op je eigen titel.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-ClientId` | Ja | Client-ID van de multi-tenant App Registration met GDAP |
| `-ClientSecret` / `-CertificateThumbprint` | * | Een van beide is verplicht |
| `-Title` | Nee | Paginatitel (standaard: `Customer Portal Index`) |
| `-OutputPath` | Nee | Uitvoerpad voor de HTML (standaard: `C:\Temp\CustomerPortalIndex.html`) |

**Voorbeeld**
```powershell
.\New-CustomerPortalIndex.ps1 -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" -Title "Our Customers"
```

---

**Vereiste Graph-machtiging op de app in de partner-/thuistenant:** `DelegatedAdminRelationship.Read.All`
**Vereiste module:** `Microsoft.Graph.Authentication`
