[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **Entra**

# Entra ID-scripts

Scripts voor het beheren van gebruikers en resources in Microsoft Entra ID (voorheen Azure AD) via Microsoft Graph.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Set-UserManager.ps1`](Set-UserManager.ps1) ([docs](#set-usermanagerps1)) | Rapporteer en stel desgewenst in bulk de manager in voor een set Entra ID-gebruikers |
| [`Remove-M365Users.ps1`](Remove-M365Users.ps1) ([docs](#remove-m365usersps1)) | M365-gebruikersaccounts in bulk verwijderen (standaard een proefdraai) |
| [`New-M365User.ps1`](New-M365User.ps1) ([docs](#new-m365userps1)) | Eén M365-gebruiker aanmaken, optioneel met licentie |
| [`Import-M365Users.ps1`](Import-M365Users.ps1) ([docs](#import-m365usersps1)) | M365-gebruikers in bulk aanmaken vanuit CSV (standaard een proefdraai) |
| [`Get-M365UserLicenses.ps1`](Get-M365UserLicenses.ps1) ([docs](#get-m365userlicensesps1)) | Rapporteer toegewezen licenties voor een lijst gebruikers |
| [`Import-ConditionalAccessBaseline.ps1`](Import-ConditionalAccessBaseline.ps1) ([docs](#import-conditionalaccessbaselineps1)) | Importeer de community-baseline voor Conditional Access |
| [`Test-M365GroupMembership.ps1`](Test-M365GroupMembership.ps1) ([docs](#test-m365groupmembershipps1)) | Controleer eigenaren en leden van M365-groepen / Teams |
| [`Copy-GroupMember.ps1`](Copy-GroupMember.ps1) ([docs](#copy-groupmemberps1)) | Kopieer leden van de ene Entra ID-groep naar een andere (standaard een proefdraai) |
| [`New-TemporaryConditionalAccessPolicy.ps1`](New-TemporaryConditionalAccessPolicy.ps1) ([docs](#new-temporaryconditionalaccesspolicyps1)) | Maak een tijdelijk CA-beleid aan voor één gebruiker of groep |
| [`Remove-TemporaryConditionalAccessPolicies.ps1`](Remove-TemporaryConditionalAccessPolicies.ps1) ([docs](#remove-temporaryconditionalaccesspoliciesps1)) | Verwijder verlopen/alle tijdelijke CA-beleidsregels |
| [`New-UserTemporaryAccessPass.ps1`](New-UserTemporaryAccessPass.ps1) ([docs](#new-usertemporaryaccesspassps1)) | Maak een TAP-code aan voor een gebruiker |
| [`Set-EntraPasskeyMigrationOptOut.ps1`](Set-EntraPasskeyMigrationOptOut.ps1) ([docs](#set-entrapasskeymigrationoptoutps1)) | Stel de automatische inschakeling van passkeys op 1 september 2026 uit (één tenant of een GDAP-lijst) |
| [`Phising-rollout.ps1`](Phising-rollout.ps1) ([docs](#phising-rolloutps1)) | Houd een uitrolgroep voor phishingbestendige MFA en een groep geregistreerden in beide richtingen synchroon |

> Het omzetten van dynamische naar statische distributiegroepen (`Set-Distributionlist-dynamic-static.ps1`) staat in [`scripts/Exchange/`](../Exchange/readme.nl.md) — dat gebruikt Exchange Online-cmdlets, geen Graph.

---

### Set-UserManager.ps1

Rapporteert de manager van een set Entra ID-gebruikers en stelt die desgewenst in bulk in. Gebruikers kun je selecteren op groep, afdeling, huidige manager of een expliciete lijst UPN's; als `-NewManager` wordt weggelaten, rapporteert het script alleen de huidige manager van elke gebruiker.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-GroupId` | * | Object-ID van een Entra ID-groep waarvan de leden verwerkt moeten worden |
| `-GroupName` | * | Weergavenaam van een Entra ID-groep (automatisch omgezet naar een ID; geeft een fout als hij niet eenduidig is) |
| `-Department` | * | Afdelingstekst om gebruikers op te filteren (exacte overeenkomst via OData) |
| `-CurrentManager` | * | UPN of object-ID van een manager — verwerkt al zijn directe ondergeschikten in de hele tenant |
| `-UserList` | * | Expliciete array van UPN's of object-ID's |
| `-NewManager` | Nee | UPN of object-ID om als manager in te stellen voor alle gevonden gebruikers. Laat weg om alleen de huidige managers te rapporteren |
| `-OutputPath` | Nee | Pad voor CSV-export |
| `-TenantId` | Nee | Entra ID-tenant-ID of -domein voor `Connect-MgGraph` |

*Precies één van `-GroupId` / `-GroupName` / `-Department` / `-CurrentManager` / `-UserList` bepaalt de bron van de gebruikers.

**Voorbeelden**

```powershell
# Toon managers voor alle leden van een groep
.\Set-UserManager.ps1 -GroupName "Sales Team"

# Stel in bulk de manager in voor een groep
.\Set-UserManager.ps1 -GroupName "Sales Team" -NewManager "jane.doe@contoso.com"

# Stel in bulk de manager in voor een afdeling, exporteer het rapport
.\Set-UserManager.ps1 -Department "Logistics" -NewManager "jane.doe@contoso.com" -OutputPath C:\Temp\ManagerReport.csv

# Vind alle directe ondergeschikten van een manager
.\Set-UserManager.ps1 -CurrentManager "old.boss@contoso.com"

# Wijs alle directe ondergeschikten van de ene manager toe aan een andere
.\Set-UserManager.ps1 -CurrentManager "old.boss@contoso.com" -NewManager "new.boss@contoso.com"
```

Ondersteunt `-WhatIf` (`SupportsShouldProcess`).

---

### Remove-M365Users.ps1

Verwijdert M365-gebruikersaccounts in bulk uit een tenant. Trekt sessies in en verwijdert licenties vóór het verwijderen. Draait standaard als proefdraai — geef `-Apply` mee om echt te verwijderen.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-UserList` | * | Array van te verwijderen UPN's |
| `-CsvPath` | * | Pad naar CSV (kolom `UserPrincipalName`) of TXT (één UPN per regel) |
| `-Apply` | Nee | Voer het verwijderen echt uit (standaard: proefdraai) |
| `-SkipLicenseRemoval` | Nee | Sla het verwijderen van licenties vóór het verwijderen over |
| `-SkipSessionRevoke` | Nee | Sla het intrekken van actieve sessies over |
| `-OutputPath` | Nee | Pad voor CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Entra ID-tenant-ID of -domein |

*`-UserList` of `-CsvPath` is verplicht.

**Voorbeelden**

```powershell
# Proefdraai — er wordt niets gewijzigd
.\Remove-M365Users.ps1 -UserList "user1@contoso.com","user2@contoso.com"

# Echt verwijderen vanuit CSV
.\Remove-M365Users.ps1 -CsvPath .\users.csv -Apply

# Invoer via de pipeline
"user1@contoso.com","user2@contoso.com" | .\Remove-M365Users.ps1 -Apply
```

**Opmerkingen**
- Verwijderen is een voorlopige verwijdering (soft delete) — accounts komen in Verwijderde gebruikers terecht en zijn 30 dagen te herstellen
- Er wordt altijd een CSV-rapport geschreven, ook bij een proefdraai
- Als er al een Graph-verbinding is (bijv. via `functies.ps1`), wordt de bestaande verbinding hergebruikt

**Vereiste module**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

---

### New-M365User.ps1

Maakt één M365-gebruiker aan via Microsoft Graph. Genereert een willekeurig wachtwoord van 16 tekens als er geen is opgegeven. Wijst na het aanmaken optioneel een licentie toe.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-UserPrincipalName` | Ja | UPN van de nieuwe gebruiker |
| `-DisplayName` | Ja | Weergavenaam |
| `-GivenName` | Nee | Voornaam |
| `-Surname` | Nee | Achternaam |
| `-Password` | Nee | Beginwachtwoord (automatisch gegenereerd als het wordt weggelaten) |
| `-UsageLocation` | Nee | ISO-landcode van twee letters (standaard: `NL`) |
| `-Department` | Nee | Afdeling |
| `-JobTitle` | Nee | Functietitel |
| `-MobilePhone` | Nee | Mobiel telefoonnummer |
| `-LicenseSkuId` | Nee | Onderdeelnummer van de licentie-SKU om toe te wijzen (bijv. `ENTERPRISEPACK`) |
| `-NoPasswordReset` | Nee | Dwing geen wachtwoordwijziging af bij de eerste aanmelding |
| `-TenantId` | Nee | Entra ID-tenant-ID of -domein |

**Voorbeelden**

```powershell
# Minimaal — automatisch gegenereerd wachtwoord
.\New-M365User.ps1 -UserPrincipalName "j.doe@contoso.com" -DisplayName "Jane Doe"

# Volledige gegevens met licentie
.\New-M365User.ps1 -UserPrincipalName "j.doe@contoso.com" -DisplayName "Jane Doe" `
    -GivenName "Jane" -Surname "Doe" -Department "Finance" -LicenseSkuId "ENTERPRISEPACK"
```

---

### Import-M365Users.ps1

Maakt M365-gebruikers in bulk aan vanuit een CSV-bestand via Microsoft Graph. Draait standaard als proefdraai — geef `-Apply` mee om accounts aan te maken. Genereert per gebruiker een uniek wachtwoord als er geen kolom Password is. Wachtwoorden worden naar de resultaten-CSV geschreven.

**Verplichte CSV-kolommen:** `UserPrincipalName`, `DisplayName`

**Optionele CSV-kolommen:** `GivenName`, `Surname` (of `LastName`), `Department`, `JobTitle`, `MobilePhone`, `UsageLocation`, `Password`, `LicenseSkuId`

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-CsvPath` | Ja | Pad naar de invoer-CSV |
| `-Apply` | Nee | Maak de accounts echt aan (standaard: proefdraai) |
| `-UsageLocation` | Nee | Standaard landcode voor alle gebruikers (standaard: `NL`) |
| `-LicenseSkuId` | Nee | Wijs deze licentie toe aan alle gebruikers (overschrijft de CSV-kolom) |
| `-NoPasswordReset` | Nee | Dwing geen wachtwoordwijziging af bij de eerste aanmelding |
| `-OutputPath` | Nee | Pad voor CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Entra ID-tenant-ID of -domein |

**Voorbeelden**

```powershell
# Eerst een proefdraai
.\Import-M365Users.ps1 -CsvPath .\users.csv

# Accounts aanmaken
.\Import-M365Users.ps1 -CsvPath .\users.csv -Apply

# Aanmaken met een licentie voor iedereen
.\Import-M365Users.ps1 -CsvPath .\users.csv -LicenseSkuId "ENTERPRISEPACK" -Apply
```

**Opmerkingen**
- Een proefdraai schrijft altijd een resultaten-CSV — controleer die voordat je met `-Apply` draait
- Gegenereerde wachtwoorden staan in de resultaten-CSV — deel ze op een veilige manier
- Een licentie vereist dat `UsageLocation` is ingesteld; het script regelt dit automatisch

---

### Get-M365UserLicenses.ps1

Controleert toegewezen licenties voor een lijst gebruikers via Microsoft Graph en exporteert een CSV-rapport. Ondersteunt een inline lijst, CSV, TXT en invoer via de pipeline.

**Invoeropties**
- `-UserList` met UPN's/e-mailadressen
- `-CsvPath` naar `.csv` (kolom: `UserPrincipalName`, `UPN` of `Mail`)
- `-CsvPath` naar `.txt` (één UPN/e-mailadres per regel)

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-UserList` | * | Array van UPN's/e-mailadressen |
| `-CsvPath` | * | Pad naar CSV/TXT met gebruikers |
| `-OutputPath` | Nee | Pad voor CSV-rapport (standaard: `C:\Temp\UserLicenseReport_<timestamp>.csv`) |
| `-TenantId` | Nee | Entra ID-tenant-ID of -domein |

*`-UserList` of `-CsvPath` is verplicht.

**Voorbeelden**

```powershell
# Specifieke gebruikers controleren
.\Get-M365UserLicenses.ps1 -UserList "user1@contoso.com","user2@contoso.com"

# Gebruikers uit een CSV controleren
.\Get-M365UserLicenses.ps1 -CsvPath .\users.csv

# Invoer via de pipeline
"user1@contoso.com","user2@contoso.com" | .\Get-M365UserLicenses.ps1
```

**Rapportuitvoer**
- Eén rij per toewijzing van een licentie aan een gebruiker
- Gebruikers zonder licentie worden opgenomen met `LicenseStatus = Unlicensed`
- Niet-gevonden gebruikers worden opgenomen met `LicenseStatus = NotFound`

---

### Import-ConditionalAccessBaseline.ps1

Importeert de nieuwste [ConditionalAccessBaseline](https://github.com/j0eyv/ConditionalAccessBaseline) in je tenant met Microsoft Graph.

Wat het doet:
- Downloadt de nieuwste baseline (of gebruikt `-SourcePath`)
- Maakt de vereiste CA-uitsluitingsgroepen aan of hergebruikt ze
- Maakt benoemde locaties aan of hergebruikt ze
- Zet oude baseline-ID's om naar de ID's van je tenant
- Importeert Conditional Access-beleidsregels voor alle persona's/platforms
- Importeert beleidsregels standaard als **Uit** (`state = disabled`)

Het ondersteunt ook een vervolgactie om geïmporteerde beleidsregels op alleen-rapporteren of ingeschakeld te zetten.

**Parameters (meest gebruikt)**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Action` | Nee | `Import` (standaard) of `SetState` |
| `-PolicyStateOnImport` | Nee | `disabled` (standaard) of `enabledForReportingButNotEnforced` |
| `-TargetState` | Nee | Voor `SetState`: `disabled`, `enabledForReportingButNotEnforced` of `enabled` |
| `-SourcePath` | Nee | Lokale baselinemap met `Config\...` |
| `-TenantId` | Nee | Tenant-ID of domein |
| `-UpdateExisting` | Nee | Werk bestaande beleidsregels met een overeenkomende weergavenaam bij |

**Voorbeelden**

```powershell
# Importeer de nieuwste baseline en laat alle CA-beleidsregels UIT
.\Import-ConditionalAccessBaseline.ps1

# Importeer de baseline in alleen-rapporteren-modus
.\Import-ConditionalAccessBaseline.ps1 -PolicyStateOnImport enabledForReportingButNotEnforced

# Later: schakel de geïmporteerde baselinebeleidsregels in
.\Import-ConditionalAccessBaseline.ps1 -Action SetState -TargetState enabled
```

**Opmerkingen**
- Houd minstens één break-glass-account uitgesloten voordat je beleidsregels inschakelt
- Controleer na het importeren de uitsluitingsgroepen en benoemde locaties

---

### New-TemporaryConditionalAccessPolicy.ps1

Maakt een tijdelijk Conditional Access-beleid aan voor één gebruiker of groep.

- Markeert de beleidsnaam met het voorvoegsel `TEMP-CA -`
- Schrijft een verloopmoment in de beleidsomschrijving (`Expires=<UTC timestamp>`)
- Ondersteunt vensters op basis van een duur of exacte lokale begin-/einddatum en -tijd
- Houdt standaard de scriptsessie open en verwijdert het beleid direct zodra het verloopmoment is bereikt

Belangrijk:
- Direct opruimen bij het verlopen vereist dat de scriptsessie open blijft
- Sluit je de sessie eerder, draai dan later het opruimscript
- Dit script maakt geen geplande Windows-taak aan; de wacht-/opruimlus draait in de huidige sessie

**Voorbeelden**

```powershell
# Tijdelijke MFA-vereiste voor 2 uur, automatisch verwijderen bij verlopen
.\New-TemporaryConditionalAccessPolicy.ps1 -TargetType User -TargetId "<object-id>" -DisplayName "Temporary MFA" -DurationHours 2

# Tijdelijk beleid met exacte lokale begin-/einddatum en -tijd
.\New-TemporaryConditionalAccessPolicy.ps1 -TargetType User -TargetId "<object-id>" -DisplayName "Install Window" -StartDateTimeLocal "2026-07-23 19:00" -EndDateTimeLocal "2026-07-23 22:00"

# Tijdelijk blokkeerbeleid, zonder wachtlus voor automatisch opruimen
.\New-TemporaryConditionalAccessPolicy.ps1 -TargetType Group -TargetId "<object-id>" -DisplayName "Temporary Block" -Action Block -NoAutoCleanup
```

---

### Remove-TemporaryConditionalAccessPolicies.ps1

Verwijdert tijdelijke beleidsregels die zijn aangemaakt met het voorvoegsel `TEMP-CA -`.

Modi:
- standaard: verwijder alleen verlopen TEMP-CA-beleidsregels
- `-PolicyId`: verwijder één specifiek beleid
- `-RemoveAllTempPolicies`: verwijder alle TEMP-CA-beleidsregels

**Voorbeelden**

```powershell
# Verwijder alleen verlopen tijdelijke beleidsregels
.\Remove-TemporaryConditionalAccessPolicies.ps1

# Verwijder één specifiek beleid
.\Remove-TemporaryConditionalAccessPolicies.ps1 -PolicyId "<policy-id>"
```

---

### New-UserTemporaryAccessPass.ps1

Maakt een Temporary Access Pass (TAP) aan voor één gebruiker.

**Voorbeeld**

```powershell
.\New-UserTemporaryAccessPass.ps1 -UserId "user@contoso.com" -LifetimeMinutes 60 -IsUsableOnce
```

**Opmerkingen**
- Gebruik bij voorkeur `-IsUsableOnce` voor support-/installatiescenario's
- Deel de TAP-code via een veilig kanaal en laat hem snel verlopen

---

### Test-M365GroupMembership.ps1

Toont alle eigenaren en leden van Microsoft 365-groepen (inclusief groepen achter Teams). De resultaten worden geëxporteerd naar CSV met één rij per eigenaar/lid. Maakt automatisch verbinding met Graph als er geen sessie actief is; hergebruikt een bestaande sessie als er al verbinding is.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Group` | Nee | Weergavenaam of object-ID van één groep. Als dit wordt weggelaten, worden alle M365-groepen gecontroleerd |
| `-OutputPath` | Nee | Pad voor CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Entra ID-tenant-ID of -domein |

**Voorbeelden**

```powershell
# Alle M365-groepen controleren
.\Test-M365GroupMembership.ps1

# Eén groep op weergavenaam
.\Test-M365GroupMembership.ps1 -Group "Team Finance"

# Eén groep op object-ID
.\Test-M365GroupMembership.ps1 -Group "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```

**Vereiste scopes**
- `Group.Read.All`
- `Directory.Read.All`

**Vereiste module**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

---

### Copy-GroupMember.ps1

Kopieert de leden van de ene Entra ID-groep naar een andere groep. Leden die al in de doelgroep zitten, worden overgeslagen, dus het script kan veilig opnieuw worden gedraaid. Draait standaard als proefdraai — geef `-Apply` mee om wijzigingen door te voeren.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-SourceGroup` | Ja | Weergavenaam of object-ID van de groep om VANUIT te kopiëren |
| `-TargetGroup` | Ja | Weergavenaam of object-ID van de groep om NAAR te kopiëren |
| `-MemberType` | Nee | `All` (standaard), `User`, `Group`, `Device` of `ServicePrincipal` |
| `-Flatten` | Nee | Klap geneste groepen uit en kopieer hun effectieve leden in plaats van het geneste groepsobject |
| `-Mirror` | Nee | Verwijder ook leden uit de doelgroep die niet in de brongroep zitten (exacte kopie in plaats van een vereniging) |
| `-Apply` | Nee | Voeg leden echt toe/verwijder ze echt (standaard: proefdraai) |
| `-Disconnect` | Nee | Meld af bij Graph als het klaar is (standaard uit — afmelden wist de tokencache en dwingt bij de volgende run een nieuwe browserprompt af) |
| `-OutputPath` | Nee | Pad voor CSV-rapport (standaard: `C:\Temp\GroupMemberCopy_<timestamp>.csv`) |
| `-TenantId` | Nee | Entra ID-tenant-ID of -domein |

**Voorbeelden**

```powershell
# Proefdraai — toon wat er gekopieerd zou worden
.\Copy-GroupMember.ps1 -SourceGroup "All Staff" -TargetGroup "MFA Rollout"

# Kopieer de leden echt
.\Copy-GroupMember.ps1 -SourceGroup "All Staff" -TargetGroup "MFA Rollout" -Apply

# Kopieer alleen gebruikers en klap geneste groepen uit
.\Copy-GroupMember.ps1 -SourceGroup "Sales" -TargetGroup "Sales Mail" -MemberType User -Flatten -Apply

# Maak de doelgroep een exacte kopie van de brongroep (toevoegen en verwijderen)
.\Copy-GroupMember.ps1 -SourceGroup "Pilot" -TargetGroup "Pilot Copy" -Mirror -Apply
```

**Opmerkingen**
- Weergavenamen worden via Graph omgezet; een niet-eenduidige naam is een harde fout — gebruik dan de object-ID
- Een doelgroep met dynamisch lidmaatschap wordt geweigerd: het lidmaatschap wordt door regels bepaald en kan niet worden bewerkt
- E-mailbeveiligingsgroepen en distributiegroepen zijn niet beschrijfbaar via Graph — gebruik daarvoor Exchange Online-cmdlets
- Ondersteunt `-WhatIf` (`SupportsShouldProcess`)

**Vereiste scopes**
- `Group.Read.All`
- `GroupMember.ReadWrite.All`

**Vereiste module**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

---

### Set-EntraPasskeyMigrationOptOut.ps1

Stelt de tijdelijke opt-out in (of heft die op) voor de automatische inschakeling van passkeys en de uitrol van de Registration Campaign in Entra ID. Het past het beleid voor verificatiemethoden van de tenant aan:

```http
PATCH https://graph.microsoft.com/beta/policies/authenticationmethodspolicy
{ "optOutSettings": { "passkeyDynamicMigration": true } }
```

Accepteert een array van tenants, zodat een GDAP-partner in één run alle klanttenants kan langsgaan. Elke tenant wordt los afgehandeld — een tenant waarbij verbinden, lezen of aanpassen mislukt, wordt gemeld en de run gaat door met de volgende.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-TenantId` | Nee | Een of meer tenant-ID's of domeinnamen. Laat weg om de huidige verbinding / standaardtenant te gebruiken |
| `-Revert` | Nee | Zet `passkeyDynamicMigration` terug op `false` (laat de tenant weer MEEDOEN aan de automatische migratie) |
| `-ReportOnly` | Nee | Lees en toon de huidige waarde zonder iets te wijzigen |

**Voorbeelden**

```powershell
# Toon de huidige instelling voor de verbonden tenant
.\Set-EntraPasskeyMigrationOptOut.ps1 -ReportOnly

# Bekijk de wijziging vooraf zonder hem te schrijven
.\Set-EntraPasskeyMigrationOptOut.ps1 -TenantId contoso.onmicrosoft.com -WhatIf

# Opt-out voor elke tenant in een tekstbestand, zonder bevestiging per tenant
.\Set-EntraPasskeyMigrationOptOut.ps1 -TenantId (Get-Content .\tenants.txt) -Confirm:$false

# Laat een tenant weer MEEDOEN aan de automatische migratie
.\Set-EntraPasskeyMigrationOptOut.ps1 -TenantId contoso.onmicrosoft.com -Revert
```

**Tijdlijn**

| Datum | Wat er gebeurt |
|------|--------------|
| 1 september 2026 | Gebruikers die voor sms of spraak zijn ingeschakeld, worden automatisch ingeschakeld voor passkeys en aangespoord door een door Microsoft beheerde registratiecampagne |
| 1 februari 2027 | Door Microsoft geleverde sms-/spraakbezorging wordt uitgefaseerd |
| Na 1 februari 2027 | Gebruikers van wie de enige MFA-methode sms of spraak is, krijgen bij het aanmelden een **blokkerende** prompt om een passkey te registreren |

**Opmerkingen**
- De opt-out stelt **alleen** het gedrag van 1 september 2026 → 1 februari 2027 uit. Er is geen opt-out voor de handhaving vanaf 1 februari 2027 — die geldt voor alle tenants
- `optOutSettings` is per augustus 2026 **alleen beta** en wordt niet getoond in het Entra-beheercentrum
- Schrijfacties worden per tenant bevestigd (`ConfirmImpact = 'High'`); geef `-Confirm:$false` mee voor onbeheerde runs over meerdere tenants
- De instelling wordt ~2 seconden na de PATCH teruggelezen; een afwijking wordt gemeld als `PatchedUnverified` in plaats van als succes behandeld
- Geeft één object per tenant terug (`Tenant`, `Before`, `After`, `Status`, `Message`), zodat je een run naar `Export-Csv` kunt pipen
- Ondersteunt `-WhatIf` (`SupportsShouldProcess`)
- Heb je na 1 februari 2027 nog sms/spraak nodig, configureer dan een door de klant beheerde telecomprovider via de Microsoft Security Store (te selecteren vanaf 30 oktober 2026)

**Vereiste scope**
- `Policy.ReadWrite.AuthenticationMethod` (`Policy.Read.All` volstaat voor `-ReportOnly`)

**Vereiste rol**
- Authentication Policy Administrator (of Global Administrator)

**Vereiste module**
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

**Referentie** — [Passkeys by default and retirement of Microsoft-provided SMS and voice authentication](https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sms-voice-retirement)

---

### Phising-rollout.ps1

Onderhoudt twee elkaar uitsluitende statische groepen die de uitrol van phishingbestendige
MFA aansturen: een **Rollout**-groep voor gebruikers die zich nog moeten registreren, en een
**Registered**-groep voor wie dat al heeft gedaan. Een gebruiker zit nooit in beide.

Voor elke gebruiker die momenteel in een van beide groepen zit:

| Situatie | Wat er gebeurt |
|-----------|--------------|
| Heeft een geaccepteerde methode en zit in Rollout | Verplaatst naar Registered — geslaagd |
| Heeft geen geaccepteerde methode meer en zit in Registered | Terugverplaatst naar Rollout — de methode is verwijderd of verlopen, dus de gebruiker moet zich opnieuw registreren |
| Al het andere | Niets; de status klopt al |

De stap terug is net zo belangrijk als de stap vooruit: zonder die stap houdt een gebruiker
die zijn passkey verwijdert stilletjes het label 'compliant'.

**Wat telt als geregistreerd**

| `-AcceptedMethod` | Betekenis |
|-------------------|---------|
| `AuthenticatorPasskey` (standaard) | Alleen een passkey in Microsoft Authenticator — een `fido2AuthenticationMethod` waarvan de AAGUID in `-AllowedAaGuids` staat. Een YubiKey, Windows Hello for Business of CBA telt **niet** mee en laat de gebruiker in Rollout |
| `AnyPhishingResistant` | Elke phishingbestendige methode telt (FIDO2-sleutels, WHfB, Platform SSO, CBA) |

> De Entra-portal noemt zo'n methode gewoon "Passkey", met een detail als
> "MS Authenticator iOS". In Graph is het geen apart methodetype maar een fido2-methode —
> alleen de AAGUID verraadt dat het Authenticator is, en daarom is de standaard een
> AAGUID-filter in plaats van een methodenaam.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-RolloutGroupId` | Object-ID van de statische Rollout-groep (moet zich nog registreren) |
| `-RegisteredGroupId` | Object-ID van de statische Registered-groep (al compliant) |
| `-AcceptedMethod` | `AuthenticatorPasskey` (standaard) of `AnyPhishingResistant` |
| `-AllowedAaGuids` | AAGUID's die tellen als passkey in Authenticator (standaard: de iOS- en Android-AAGUID's van Microsoft Authenticator). Genegeerd bij `AnyPhishingResistant` |
| `-Interactive` | Meld aan via de browser in plaats van een managed identity — om het vanaf je eigen machine te draaien |
| `-UseAppRegistration` | Gebruik een bestaande app-registratie (certificaat of secret) |
| `-UseTemporaryApp` | Maak een wegwerp-app-registratie aan, draai app-only en verwijder hem daarna weer |

**Authenticatie**

Gebouwd als **Azure Automation-runbook** op een door het systeem toegewezen managed
identity, om elke een tot vier uur te draaien. Zonder `-Interactive` of een app-registratie
probeert het een managed identity, wat buiten Azure altijd mislukt.

Voor een run over veel gebruikers is `-UseTemporaryApp` (of `-UseAppRegistration` met een
certificaat) de betrouwbare keuze: een app-only-token wordt bij elke aanroep vers
aangemaakt en kan nooit halverwege de run een browserprompt opleveren. Een interactieve
sessie kan dat wel, en doet het ook — precies als het token midden in de run verloopt.

**Vereiste Graph-applicatiemachtigingen**

```
User.Read.All
UserAuthenticationMethod.Read.All
GroupMember.ReadWrite.All      (of ruimer: Group.ReadWrite.All)
```

`-UseTemporaryApp` vereist daarnaast Application Administrator of Global Administrator,
omdat het een app aanmaakt en er app-rollen aan toewijst.

> Het bestand heet in de repository `Phising-rollout.ps1`. Het hernoemen is een aparte
> wijziging — het Automation-runbook dat het aanroept, verwijst naar deze naam.
