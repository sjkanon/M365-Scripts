[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **Office365Toolkit**

# Office365Toolkit

Moderne herschrijvingen op basis van Microsoft Graph / Exchange Online van nog steeds bruikbare functies uit
de uitgefaseerde PowerShell-toolkit [`directorcia/Office365`](https://github.com/directorcia/Office365)
(CIAOPS) — een bekende verzameling beheerscripts voor Office 365 die
oorspronkelijk gebouwd was op de inmiddels uitgefaseerde modules `MSOnline` / `AzureAD` en diverse
tenant-specifieke scripts met hardgecodeerde waarden.

Het bronproject (~58 scripts) is per functie beoordeeld in plaats van
1-op-1 overgezet: dubbele rapportagescripts met één doel zijn samengevoegd,
hulpscripts voor het verbinden zijn geschrapt (deze repo maakt al automatisch verbinding en hergebruikt sessies),
alles wat de uitgefaseerde modules `MSOnline`/`AzureAD` nodig had is herschreven op basis van
`Microsoft.Graph.*` of `ExchangeOnlineManagement`, en functies die elders in deze repo al
worden gedekt (audits van mailboxgrootte/-rechten, extern doorsturen,
SharePoint-opslag, licentierapportage, Conditional Access, enz. — zie
[`scripts/Entra/`](../Entra/readme.nl.md), [`scripts/Exchange/`](../Exchange/readme.nl.md),
[`scripts/Reporting/`](../Reporting/readme.nl.md)) zijn overgeslagen. Alle code hier is een
nieuwe implementatie in de eigen stijl van deze repo, niet gekopieerd uit het bronproject.

---

## Mappen

| Map | Omschrijving |
|--------|-------------|
| [`Security/`](Security/readme.nl.md) | Secure Score-rapportage, opschonen van toestemmingen voor enterprise-apps, aanmelding op gedeelde mailboxen dichtzetten, EOP-basislijn voor antispam/antimalware |
| [`Exchange/`](Exchange/readme.nl.md) | Hygiënebasislijn voor mailboxen, doorstuurrisico via inboxregels, mailbox-invoegtoepassingen, zoeken in het Unified Audit Log, message trace |
| [`Intune/`](Intune/readme.nl.md) | Tenantbrede inventaris van Intune/Endpoint Manager-beleid |

---

## Overgeslagen functies (en waarom)

**Elders in deze repo al gedekt:**
- Audits van mailboxgrootte/-rechten/-agenda/DKIM, extern doorsturen via mailbox-
  instelling, SharePoint-opslaggebruik, licentierapportage, import van de Conditional Access-
  basislijn, TAP-codes — hiervoor bestaan al Graph/EXO-scripts onder `scripts/Entra/`,
  `scripts/Exchange/` en `scripts/Reporting/`.
- Opzoeken van leesbare namen voor licentie-SKU's (`o365-skus.ps1`) — een statische hashtable met
  verouderde SKU-namen; vervangen door `scripts/Entra/Get-M365UserLicenses.ps1`.

**Hulpscripts voor verbinding/infrastructuur, geen functies** (`*-connect*.ps1`,
`graph-connect.ps1`, `msgraph-connect.ps1`, `Intune-connect.ps1`, `az-connect*.ps1`,
`o365-setup.ps1`, `o365-update.ps1`, `o365-getrepo.ps1`, `save-cred-file.ps1`,
`c.ps1`, `r.ps1`, `sc-config.ps1`, `text-colour.ps1`): deze repo maakt in elk script al
automatisch verbinding en hergebruikt bestaande sessies (zie de huisstijl in
`scripts/Entra/Test-M365GroupMembership.ps1`), dus losse verbindingsscripts voegen
niets toe. `o365-setup.ps1` bevatte bovendien het hardgecodeerde OneDrive-pad van een echte klant en
installeert de uitgefaseerde modules `MSOnline`/`AzureAD`; `save-cred-file.ps1` slaat
inloggegevens op in een lokaal XML-bestand — beide vallen buiten de veiligheidsregels van dit project.

**Geen M365-tenantscope** (lokale diagnostiek van apparaat/netwerk, geen Graph/EXO):
`win10-asr-get.ps1`, `win10-audit-get.ps1`, `win10-def-get.ps1` (lokale controles van Windows
Defender/auditbeleid — op apparaatniveau, niet tenantbreed), `Cleanup AzureAD
device registration.ps1` (lokale opschoning van het register), `sec-test.ps1` (lokaal menu voor EICAR-/
malwaresimulatie), `ipget.ps1`/`ipinf.ps1` (generieke IP-geolocatiezoekopdrachten,
één met een hardgecodeerde persoonlijke API-sleutel — los van M365).

**Leverancier-/tenantspecifiek, niet generiek herbruikbaar:** `sc-config.ps1` (hardgecodeerd
op de Teams Direct Routing-configuratie van één Australische PSTN-provider).

**Verouderd of te beperkt, weinig rendement om opnieuw te bouwen:**
- `o365-addin-deploy.ps1` (Centralized Deployment van Outlook-invoegtoepassingen) — Microsoft
  stuurt beheerders richting de Integrated Apps-interface in het beheercentrum; er bestaat
  nog geen equivalente Graph/EXO-cmdlet.
- `o365-mcas-api.ps1` (Cloud App Security API) en `endpoint-api-svbm.ps1`
  (kwetsbaarheden-API van Defender for Endpoint) — beide hardcoderen placeholder-
  variabelen voor URI/token/secret en richten zich op API's die grotendeels zijn vervangen door de uniforme
  Defender-portal; past niet netjes in Graph/EXO.
- `o365-atp-timer.ps1` (eenmalige meting van de ATP-scanlatentie) — nichediagnostiek,
  geen doorlopende beheerfunctie.
- `az-sentinel-ruleget.ps1` (rapport van Azure Sentinel-analyseregels) — Azure
  Resource Manager-scope (`Az.SecurityInsights`), geen M365/Graph/EXO-functie.
- Beheerders van SharePoint-sitecollecties, lijst van externe gebruikers en instellingen
  voor delen (`o365-spo-admins.ps1`, `o365-spo-extusr.ps1`, `o365-spo-getsharing.ps1`)
  — niet netjes beschikbaar via Microsoft Graph zonder SharePoint Online Management
  Shell of PnP.PowerShell, die buiten de Graph/EXO-modulescope van dit project
  vallen; `o365-spo-getusage.ps1` (opslaggebruik) wordt al gedekt door
  `scripts/Reporting/Get-SharePointStorageReport.ps1`.
