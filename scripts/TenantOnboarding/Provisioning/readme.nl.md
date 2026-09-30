[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [TenantOnboarding](../readme.nl.md) › **Provisioning**

# Provisioning

Scripts voor de bootstrap van één net ge-onboarde tenant: break-glass-beheeraccount, basisbeveiligingsgroepen en het toewijzen van het Intune-basisbeleid. Vervangt een oud interactief installatiescript met menu door losse, geparametriseerde scripts die aansluiten bij de rest van deze repository. Voer ze uit in de volgorde hieronder, als onderdeel van een checklist voor nieuwe tenants.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`New-BreakGlassAdminAccount.ps1`](New-BreakGlassAdminAccount.ps1) ([docs](#new-breakglassadminaccountps1)) | Maak een cloud-only Global Administrator-account voor noodtoegang aan |
| [`New-TenantBaselineGroups.ps1`](New-TenantBaselineGroups.ps1) ([docs](#new-tenantbaselinegroupsps1)) | Maak de CA-uitsluitingsgroep en alle groepen voor het vrijgeven van functies die een nieuwe tenant nodig heeft |
| [`Set-IntuneBaselinePolicyAssignment.ps1`](Set-IntuneBaselinePolicyAssignment.ps1) ([docs](#set-intunebaselinepolicyassignmentps1)) | Wijs op naam gefilterde Intune-basisbeleidsregels in bulk toe aan een doelgroep |

**Aanbevolen volgorde voor een nieuwe tenant:**
1. `New-BreakGlassAdminAccount.ps1`: maak het account voor noodtoegang aan
2. `New-TenantBaselineGroups.ps1`: maak de CA-uitsluitingsgroep aan (die verwijst naar het UPN-patroon van het break-glass-account) en eventuele groepen voor het vrijgeven van functies
3. `New-BreakGlassAdminAccount.ps1 -ExcludeFromGroupId <exclusion group id>` (of voer `Add-UserToFeatureGroup.ps1` uit `../UserManagement/` opnieuw uit) om zeker te zijn dat het break-glass-account lid is van zijn eigen uitsluitingsgroep
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
| `-ExcludeFromGroupId` | Nee | Groep waaraan het account wordt toegevoegd (bijv. een CA-uitsluitingsgroep) |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |
| `-Apply` | Nee | Maak het account echt aan (standaard: alleen voorbeeldweergave) |

**Voorbeelden**

```powershell
.\New-BreakGlassAdminAccount.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com"
.\New-BreakGlassAdminAccount.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com" -Apply
```

**Vereiste modules**
```powershell
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
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
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |
| `-Apply` | Nee | Maak de groepen echt aan (standaard: alleen voorbeeldweergave) |

*Verplicht, tenzij `-SkipExclusionGroup` wordt gebruikt.

**Voorbeelden**

```powershell
.\New-TenantBaselineGroups.ps1 -BreakGlassUpnPattern "breakglass-admin" `
    -AdditionalGroupNames "SG - Enable Password Manager","SG - Enable Windows 365"

.\New-TenantBaselineGroups.ps1 -BreakGlassUpnPattern "breakglass-admin" -Apply
```

**Vereiste module**
```powershell
Install-Module Microsoft.Graph.Groups -Scope CurrentUser
```

---

### Set-IntuneBaselinePolicyAssignment.ps1

Wijst Intune-configuratieprofielen, nalevingsbeleid, beheersjablonen, scripts en beveiligingsbaselines waarvan de weergavenaam aan een filter voldoet, in bulk toe aan één doelgroep.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-TargetGroupId` | Ja | Groep waaraan het overeenkomende beleid wordt toegewezen |
| `-NameFilter` | Nee | Wildcardfilter op de weergavenaam van het beleid (standaard: `*Default*`) |
| `-PolicyTypes` | Nee | Welke objecttypen worden meegenomen (standaard: alle) |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |
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
- Alle scripts in deze map maken automatisch verbinding met Microsoft Graph als er geen sessie actief is, en hergebruiken een bestaande sessie als je al verbonden bent, net als `Test-M365GroupMembership.ps1` in `scripts/Entra/`.
