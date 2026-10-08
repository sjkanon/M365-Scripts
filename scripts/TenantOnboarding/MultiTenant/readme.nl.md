[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [TenantOnboarding](../readme.nl.md) › **MultiTenant**

# MultiTenant

MSP-brede scripts die in één keer over elke GDAP-klanttenant (Granular Delegated Admin Privileges) heen werken: licentierapportage, rotatie van break-glass-wachtwoorden en een index met snelkoppelingen naar portalen. Dit zijn de op Microsoft Graph gebaseerde vervangers van een familie legacy-scripts die met de inmiddels uitgefaseerde MSOnline-module door `Get-MsolPartnerContract -All` liepen.

**Aanmelden** gaat via [`Connect-M365.ps1`](../../Startup/Connect-M365.ps1) en is **standaard gedelegeerd**:

- **Gedelegeerd (GDAP, standaard)**: je meldt je één keer aan als partnerbeheerder in je eigen tenant, het script haalt de klanten op uit `tenantRelationships/delegatedAdminCustomers` (of `/contracts` als die lijst niet leesbaar is) en maakt daarna als jou via GDAP verbinding met elke klanttenant. Je GDAP-relatie moet een rol bevatten die het werk toestaat (zie elk script). Elke klant is een aparte aanmelding: in de browser lukt dat meestal met het account in de cache, met apparaatcode (`$global:useDeviceCodeAuth`) voer je per klant een code in. De eerste keer kan *Microsoft Graph Command Line Tools* toestemming nodig hebben in een klanttenant.
- **App-only (optie)**: `-ClientId` met `-CertificateThumbprint` (voorkeur) of `-ClientSecret`, plus `-TenantId` (je partnertenant), of `-AppOnly` om ze uit `graph.appid.json` te lezen. Daarvoor is een eigen multi-tenant App Registration nodig die **in elke klanttenant toestemming heeft**. GDAP alleen geeft een app geen toegang tot klanten: GDAP verleent alleen gedelegeerde rechten aan je gebruikers.

Zonder `-TenantId` gaat de gedelegeerde partneraanmelding naar de tenant van het account waarmee je je aanmeldt, ook als in `load.config.ps1` een GDAP-klant is geselecteerd.

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
| `-TenantId` | App-only | Je partner-/thuistenant. Gedelegeerd: standaard de tenant van het account waarmee je je aanmeldt |
| `-ClientId` | Nee | Multi-tenant App Registration voor app-only aanmelden; weglaten voor gedelegeerd |
| `-CertificateThumbprint` | * | Certificaatvingerafdruk voor `-ClientId` (voorkeur) |
| `-ClientSecret` | * | Client secret voor `-ClientId` |
| `-AppOnly` | Nee | App-only met ClientId en vingerafdruk uit `graph.appid.json` (vermelding van de partnertenant) |
| `-CustomerTenantId` | Nee | Beperk tot specifieke klanttenant-ID's |
| `-OutputPath` | Nee | Pad voor het CSV-rapport (standaard: `C:\Temp\`) |

*Met `-ClientId` is een van `-CertificateThumbprint` / `-ClientSecret` verplicht.

**Voorbeelden**
```powershell
# Gedelegeerd: aanmelden als partnerbeheerder
.\Get-MultiTenantLicenseReport.ps1

# App-only met een certificaat
.\Get-MultiTenantLicenseReport.ps1 -TenantId "partner.onmicrosoft.com" `
    -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
```

**Opmerkingen**
- Gedelegeerd: de GDAP-relatie heeft een rol nodig die gebruikers mag lezen (Global Reader, Directory Readers of User Administrator); scope `User.Read.All` per klant.
- App-only: `User.Read.All` als toepassingsmachtiging, met toestemming in elke klanttenant.
- Een klant die niet bereikbaar is, wordt gemeld als `[WARN] Skipped` en het rapport gaat door.

---

### Update-BreakGlassAdminPassword.ps1

Roteert het wachtwoord van een break-glass-account, in één tenant of in elke GDAP-klanttenant (account `<lokaal deel>@<initieel onmicrosoft.com-domein>`). Nieuwe wachtwoorden worden één keer getoond en nooit op schijf opgeslagen.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-UserPrincipalName` | * | Volledige UPN, modus voor één tenant |
| `-UserPrincipalNameLocalPart` | * | Lokaal deel van de UPN, gebruikt met `-AllCustomers` |
| `-AllCustomers` | Nee | Roteer in elke GDAP-klanttenant |
| `-TenantId` | Nee | Eén tenant: de tenant (standaard: GDAP-klant, anders je eigen). `-AllCustomers`: je partnertenant (verplicht voor app-only) |
| `-ClientId` / `-CertificateThumbprint` / `-ClientSecret` | Nee | App-only aanmelden; weglaten voor gedelegeerd |
| `-AppOnly` | Nee | App-only met ClientId en vingerafdruk uit `graph.appid.json` |
| `-PasswordLength` | Nee | Standaard `24` |
| `-Apply` | Nee | Roteer echt (standaard: alleen voorbeeldweergave) |

**Voorbeelden**
```powershell
# Eén tenant, gedelegeerd (hergebruikt een sessie voor die tenant met User.ReadWrite.All)
.\Update-BreakGlassAdminPassword.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com" -Apply

# Elke GDAP-klant, gedelegeerd als partnerbeheerder
.\Update-BreakGlassAdminPassword.ps1 -UserPrincipalNameLocalPart "breakglass-admin" -AllCustomers -Apply

# Elke GDAP-klant, app-only
.\Update-BreakGlassAdminPassword.ps1 -UserPrincipalNameLocalPart "breakglass-admin" -AllCustomers `
    -TenantId "partner.onmicrosoft.com" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" -Apply
```

**Opmerkingen**
- Het wachtwoord van een beheerder opnieuw instellen vereist Privileged Authentication Administrator (of Global Administrator): in de GDAP-relatie bij gedelegeerd gebruik, of toegewezen aan de service principal van de app in elke klant bij app-only (met de toepassingsmachtigingen `User.ReadWrite.All` en `Domain.Read.All`).
- Het account wordt gezocht op het **initiële** domein van de tenant (`isInitial`); eerdere versies namen het eerste `*.onmicrosoft.com`-domein, en dat kan `contoso.mail.onmicrosoft.com` zijn.
- De modus voor één tenant hergebruikt niet meer zomaar elke Graph-sessie: die moet voor de juiste tenant zijn en `User.ReadWrite.All` hebben.

---

### New-CustomerPortalIndex.ps1

Bouwt één doorzoekbare HTML-pagina met per klant directe links naar het M365-beheercentrum, Entra ID, het Exchange-beheercentrum, het Teams-beheercentrum en Intune. Er is geen huisstijl hardcoded; stel `-Title` in op je eigen titel. Alleen de partnertenant wordt gelezen, dus één aanmelding volstaat.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-TenantId` | App-only | Je partner-/thuistenant. Gedelegeerd: standaard de tenant van het account waarmee je je aanmeldt |
| `-ClientId` | Nee | App Registration voor app-only aanmelden; weglaten voor gedelegeerd |
| `-CertificateThumbprint` / `-ClientSecret` | * | Een van beide is verplicht met `-ClientId` |
| `-AppOnly` | Nee | App-only met ClientId en vingerafdruk uit `graph.appid.json` |
| `-Title` | Nee | Paginatitel (standaard: `Customer Portal Index`) |
| `-OutputPath` | Nee | Uitvoerpad voor de HTML (standaard: `C:\Temp\CustomerPortalIndex.html`) |

**Voorbeelden**
```powershell
# Gedelegeerd: aanmelden als partnerbeheerder
.\New-CustomerPortalIndex.ps1 -Title "Our Customers"

.\New-CustomerPortalIndex.ps1 -TenantId "partner.onmicrosoft.com" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -CertificateThumbprint "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" -Title "Our Customers"
```

**Opmerkingen**
- De tabel wordt nu rechtstreeks als HTML geschreven: `ConvertTo-Html` codeerde de celwaarden, waardoor de links als letterlijke tekst `<a href=...>` verschenen. Klantnamen worden HTML-gecodeerd en het zoekvak verbergt de kopregel niet meer.

---

**Vereiste Graph-machtiging in de partnertenant:** `DelegatedAdminRelationship.Read.All` (gedelegeerde scope of toepassingsmachtiging); `Directory.Read.All` voor de terugval op `/contracts`
**Vereiste module:** `Microsoft.Graph.Authentication`
