[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **Azure**

# Scripts voor Azure-infrastructuur

Scripts voor het rechtstreeks beheren van Azure IaaS-resources (niet de M365-tenant) — los van elke andere categorie in deze repo, die zich via Graph of Exchange Online op Microsoft 365 / Entra ID richt. Vereist de PowerShell-module `Az` en een geauthenticeerde `Connect-AzAccount`-sessie. Niet opgenomen in [`menu.ps1`](../../menu.ps1) — voer het rechtstreeks uit tegen het doelabonnement.

---

## Mappen

| Map | Omschrijving |
|--------|-------------|
| [`VM/`](VM/readme.nl.md) | Het type schijfcontroller van een Azure-VM omzetten tussen SCSI en NVMe |

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Search-AADDSUserActivity.ps1`](Search-AADDSUserActivity.ps1) ([docs](#search-aaddsuseractivityps1)) | Doorzoek in één query alle audittabellen van Azure AD Domain Services in Log Analytics op één gebruiker |

---

### Search-AADDSUserActivity.ps1

Doorzoekt de diagnostische AAD DS-tabellen in een Log Analytics-werkruimte met één `union`-query op een opgegeven gebruikersnaam, zodat je niet vooraf hoeft te weten in welke tabel (of kolom) een gebeurtenis terecht is gekomen. Vereist diagnostische instellingen op het beheerde AAD DS-domein die logs naar de doelwerkruimte sturen.

**Standaard doorzochte tabellen**
- `AADDomainServicesAccountManagement`
- `AADDomainServicesAccountLogon`
- `AADDomainServicesLogonLogoff`
- `AADDomainServicesDirectoryServiceAccess`

Overschrijf met `-Table` om andere toe te voegen (bijv. `AADDomainServicesDNSAuditsGeneral`) of de zoekopdracht te beperken.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Username` | Ja | Gebruikersnaam/SamAccountName/UPN om op te zoeken — gematcht met `has` over elke kolom |
| `-WorkspaceId` | * | ID van de Log Analytics-werkruimte (`CustomerId`-GUID) |
| `-WorkspaceName` | * | Naam van de werkruimte — wordt automatisch opgezocht; combineer met `-ResourceGroupName` als die niet eenduidig is |
| `-ResourceGroupName` | Nee | Resourcegroep van de werkruimte, om `-WorkspaceName` eenduidig te maken |
| `-HoursBack` | Nee | Aantal uren terugkijken vanaf nu (standaard: `2`). Genegeerd als `-StartTime` is ingesteld |
| `-StartTime` / `-EndTime` | Nee | Expliciet zoekvenster (lokale tijd), gaat voor `-HoursBack` |
| `-Table` | Nee | Tabellen die in de union worden opgenomen (zie standaardwaarden hierboven) |
| `-MaxRows` | Nee | Maximaal aantal geretourneerde rijen, nieuwste eerst (standaard: `5000`) |
| `-ExportPath` | Nee | Map voor het CSV-rapport (standaard: `C:\Temp`) |

*Als `-WorkspaceId` is weggelaten, probeert het script automatisch één werkruimte te vinden via `-WorkspaceName`/`-ResourceGroupName`, of de enige werkruimte in het abonnement als er maar één is.

**Voorbeelden**

```powershell
# Directe werkruimte-ID, standaardvenster van 2 uur
.\Search-AADDSUserActivity.ps1 -Username "jdoe" -WorkspaceId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

# Werkruimte op naam opzoeken, 24 uur terugkijken
.\Search-AADDSUserActivity.ps1 -Username "jdoe" -WorkspaceName "log-aadds-prod" -HoursBack 24

# Expliciet tijdvenster
.\Search-AADDSUserActivity.ps1 -Username "jdoe" -WorkspaceId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -StartTime "2026-07-25 00:00" -EndTime "2026-07-27 00:00"
```

**Opmerkingen**
- Maakt automatisch verbinding met `Connect-AzAccount` als er geen Az-sessie actief is
- De resultaatset is begrensd op `-MaxRows` (standaard 5000) — het script waarschuwt als die grens is bereikt, zodat je weet dat je het venster moet verkleinen of de grens moet verhogen
- De CSV wordt geëxporteerd naar `-ExportPath` als `AADDSUserActivity_<username>_<timestamp>.csv`

**Vereiste modules**
```powershell
Install-Module Az.Accounts, Az.OperationalInsights -Scope CurrentUser
```
