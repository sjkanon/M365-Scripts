[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [TenantOnboarding](../readme.nl.md) › **Provisioning**

# Provisioning

Scripts voor de bootstrap van één net ge-onboarde tenant: break-glass-beheeraccount, basisbeveiligingsgroepen en het toewijzen van het Intune-basisbeleid. Vervangt een oud interactief installatiescript met menu door losse, geparametriseerde scripts die aansluiten bij de rest van deze repository. Voer ze uit in de volgorde hieronder, als onderdeel van een checklist voor nieuwe tenants.

**Aanmelden** (alle drie de scripts) gaat via [`Connect-M365.ps1`](../../Startup/Connect-M365.ps1): **standaard gedelegeerd**, je meldt je aan als beheerder van de tenant (apparaatcode als `$global:useDeviceCodeAuth` is ingesteld; onder GDAP komt de klanttenant uit `$global:cid`, tenzij `-TenantId` er een noemt). **App-only** is een optie met `-ClientId` + `-CertificateThumbprint`, of `-AppOnly` om ze uit `graph.appid.json` te lezen. Een bestaande Graph-sessie wordt alleen hergebruikt als die voor de juiste tenant is en al de scopes heeft die het script nodig heeft; het script verbreekt alleen een sessie die het zelf heeft geopend.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`New-BreakGlassAdminAccount.ps1`](New-BreakGlassAdminAccount.ps1) ([docs](#new-breakglassadminaccountps1)) | Maak een cloud-only Global Administrator-account voor noodtoegang aan |
| [`New-TenantBaselineGroups.ps1`](New-TenantBaselineGroups.ps1) ([docs](#new-tenantbaselinegroupsps1)) | Maak de CA-uitsluitingsgroep en alle groepen voor het vrijgeven van functies die een nieuwe tenant nodig heeft |
| [`Set-IntuneBaselinePolicyAssignment.ps1`](Set-IntuneBaselinePolicyAssignment.ps1) ([docs](#set-intunebaselinepolicyassignmentps1)) | Wijs op naam gefilterde Intune-basisbeleidsregels in bulk toe aan een doelgroep |

**Aanbevolen volgorde voor een nieuwe tenant:**
1. `New-BreakGlassAdminAccount.ps1`: maak het account voor noodtoegang aan
2. `New-TenantBaselineGroups.ps1`: maak de groep "alle gebruikers behalve break glass" aan (de dynamische regel laat het UPN-patroon van het break-glass-account weg) en eventuele groepen voor het vrijgeven van functies
3. Alleen als je Conditional Access-beleid een *statische* groep uitsluit: voeg het break-glass-account daaraan toe met `New-BreakGlassAdminAccount.ps1 -ExcludeFromGroupId <group id>` bij het aanmaken, of later met `Add-UserToFeatureGroup.ps1` uit `../UserManagement/`. Aan de dynamische groep uit stap 2 hoeven (en kunnen) geen leden te worden toegevoegd.
4. Importeer/configureer je Conditional Access- en Intune-basisbeleid (bijv. `Import-ConditionalAccessBaseline.ps1` in `scripts/Entra/`)
5. `Set-IntuneBaselinePolicyAssignment.ps1`: wijs het Intune-basisbeleid toe aan de groep "alle gebruikers behalve break glass"

---

### New-BreakGlassAdminAccount.ps1

Maakt een cloud-only gebruiker in Entra ID aan als account voor noodtoegang ("break-glass"), volgens de gedocumenteerde richtlijnen van Microsoft: een apart cloud-only account, een lang willekeurig wachtwoord (één keer getoond, nooit op schijf opgeslagen), en Global Administrator direct toegewezen. Standaard een proefdraai.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-UserPrincipalName` | Ja | UPN voor het nieuwe account; gebruik het `*.onmicrosoft.com`-domein van de tenant |
| `-DisplayName` | Nee | Weergavenaam (standaard: `Break Glass Admin`) |
| `-PasswordLength` | Nee | Lengte van het gegenereerde wachtwoord (standaard: `24`) |
| `-AssignGlobalAdmin` | Nee | Wijs Global Administrator toe (standaard: aan) |
| `-ExcludeFromGroupId` | Nee | Een statische CA-uitsluitingsgroep waaraan het account wordt toegevoegd (niet de dynamische groep uit `New-TenantBaselineGroups.ps1`) |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID (standaard: GDAP-klant, anders de tenant waarbij je je aanmeldt) |
| `-ClientId` / `-CertificateThumbprint` | Nee | App-only aanmelden |
| `-AppOnly` | Nee | App-only met ClientId en vingerafdruk uit `graph.appid.json` |
| `-Apply` | Nee | Maak het account echt aan (standaard: alleen voorbeeldweergave) |

**Voorbeelden**

```powershell
.\New-BreakGlassAdminAccount.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com"
.\New-BreakGlassAdminAccount.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com" -Apply
```

**Opmerkingen**
- Scopes: `User.ReadWrite.All`, `RoleManagement.ReadWrite.Directory`, `GroupMember.ReadWrite.All`.
- Global Administrator wordt toegewezen met een unified role assignment (`roleManagement/directory/roleAssignments`, rol `62e90394-69f5-4237-9190-012177145e10`). De eerdere code probeerde de rol te activeren met `New-MgDirectoryRoleTemplate -RoleTemplateId`; die cmdlet maakt een roltemplate aan en heeft geen parameter `-RoleTemplateId`, zodat een tenant waarin de rol nooit was geactiveerd bij die stap vastliep.

**Vereiste modules**
```powershell
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.Governance -Scope CurrentUser
Install-Module Microsoft.Graph.Groups -Scope CurrentUser
```

---

### New-TenantBaselineGroups.ps1

Maakt een dynamische uitsluitingsgroep "alle gebruikers behalve break-glass-accounts" aan, plus zoveel statische groepen voor het vrijgeven van functies als je opgeeft, om Conditional Access- en Intune-toewijzingen af te bakenen in een net ge-onboarde tenant.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-BreakGlassUpnPattern` | * | Patroon dat met UPN's wordt vergeleken om de dynamische regel van de uitsluitingsgroep op te bouwen |
| `-ExclusionGroupName` | Nee | Weergavenaam van de uitsluitingsgroep |
| `-SkipExclusionGroup` | Nee | Sla de uitsluitingsgroep volledig over |
| `-AdditionalGroupNames` | Nee | Namen van extra statische groepen die worden aangemaakt |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID (standaard: GDAP-klant, anders de tenant waarbij je je aanmeldt) |
| `-ClientId` / `-CertificateThumbprint` | Nee | App-only aanmelden |
| `-AppOnly` | Nee | App-only met ClientId en vingerafdruk uit `graph.appid.json` |
| `-Apply` | Nee | Maak de groepen echt aan (standaard: alleen voorbeeldweergave) |

*Verplicht, tenzij `-SkipExclusionGroup` wordt gebruikt.

**Voorbeelden**

```powershell
.\New-TenantBaselineGroups.ps1 -BreakGlassUpnPattern "breakglass-admin" `
    -AdditionalGroupNames "SG - Enable Password Manager","SG - Enable Windows 365"

.\New-TenantBaselineGroups.ps1 -BreakGlassUpnPattern "breakglass-admin" -Apply
```

**Opmerkingen**
- Scope: `Group.ReadWrite.All`. Dynamisch lidmaatschap vereist Entra ID P1.

**Vereiste module**
```powershell
Install-Module Microsoft.Graph.Groups -Scope CurrentUser
```

---

### Set-IntuneBaselinePolicyAssignment.ps1

Wijst Intune-configuratieprofielen, nalevingsbeleid, beheersjablonen, platformscripts en beveiligingsbaselines waarvan de weergavenaam aan een filter voldoet, in bulk toe aan één doelgroep. De groep wordt **toegevoegd** aan de bestaande toewijzingen van elk beleid.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-TargetGroupId` | Ja | Groep waaraan het overeenkomende beleid wordt toegewezen |
| `-NameFilter` | Nee | Wildcardfilter op de weergavenaam van het beleid (standaard: `*Default*`) |
| `-PolicyTypes` | Nee | Welke objecttypen worden meegenomen (standaard: alle) |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID (standaard: GDAP-klant, anders de tenant waarbij je je aanmeldt) |
| `-ClientId` / `-CertificateThumbprint` | Nee | App-only aanmelden |
| `-AppOnly` | Nee | App-only met ClientId en vingerafdruk uit `graph.appid.json` |
| `-Apply` | Nee | Maak de toewijzingen echt aan (standaard: alleen voorbeeldweergave) |

**Voorbeelden**

```powershell
.\Set-IntuneBaselinePolicyAssignment.ps1 -TargetGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
.\Set-IntuneBaselinePolicyAssignment.ps1 -TargetGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Apply
```

**Vereiste module**
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

**Opmerkingen**
- Gebruikt het beta-endpoint van Microsoft Graph: het toewijzen van Intune-beleid is op v1.0 niet voor elk beleidstype volledig beschikbaar.
- De actie `/assign` van Intune vervangt de hele toewijzingslijst van een beleid. Het script leest nu de huidige toewijzingen en stuurt die samen met de nieuwe groep terug; voorheen verloor elk beleid dat het aanraakte zijn andere toewijzingen. Beleid dat al aan de groep is toegewezen, wordt overgeslagen.
- Beleidslijsten volgen `@odata.nextLink`, zodat tenants met meer beleid dan één pagina volledig worden meegenomen.
- Settings catalog-beleid (`configurationPolicies`) valt erbuiten.
- Scopes: `DeviceManagementConfiguration.ReadWrite.All`, `DeviceManagementServiceConfig.ReadWrite.All`.
