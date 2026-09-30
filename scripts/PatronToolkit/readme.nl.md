[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **PatronToolkit**

# Patron Toolkit

Moderne herschrijvingen op basis van Microsoft Graph / Exchange Online van de nog steeds bruikbare functies uit
de uitgefaseerde toolkit [`directorcia/patron`](https://github.com/directorcia/patron) — een
grote verzameling van ~183 scripts voor beheer en rapportage van Microsoft 365-tenants, deels gebouwd op
de inmiddels uitgefaseerde PowerShell-modules `MSOnline` en `AzureAD`.

Geen van de code hieronder is uit dat project gekopieerd — het inspireerde de *lijst met functies*
(wat je moet controleren, hoe een goed rapport eruitziet), maar elk script hier is een nieuwe
implementatie in de huisstijl van deze repo: standaard een proefdraai voor alles wat iets wijzigt,
CSV-export, hergebruik van verbindingen met oog voor GDAP/`-TenantId`, en geen hardgecodeerde inloggegevens of
tenantgegevens. Omdat het bronproject per controle één klein script leverde met veel
overlap, voegt deze toolkit ~70 van die scripts met één doel samen tot 13
goed te parametriseren scripts, gegroepeerd per functie in plaats van 1-op-1 overgezet. Zie de sectie
`.NOTES` van elk script voor precies welke upstream-scripts het vervangt.

---

## Mappen

| Map | Omschrijving |
|--------|-------------|
| [`Entra/`](Entra/readme.nl.md) | Rapportage van MFA/SSPR-registratie, back-up van Conditional Access-beleid |
| [`Security/`](Security/readme.nl.md) | Audit van app-toestemmingen, verdachte inboxregels, beveiligingswaarschuwingen, beveiligingsstatus van e-mail, auditlogging, validatie van SPF/DMARC |
| [`Intune/`](Intune/readme.nl.md) | Rapportage van Intune-beleidstoewijzingen, inventaris van Autopilot-apparaten |
| [`Exchange/`](Exchange/readme.nl.md) | Message trace / diagnostiek van de mailflow |
| [`SharePoint/`](SharePoint/readme.nl.md) | Deelconfiguratie van SharePoint Online en audit van externe gebruikers |
| [`Teams/`](Teams/readme.nl.md) | Governance van de Teams-tenant en inventarisrapportage |

---

## Wat is overgeslagen, en waarom

De ~183 scripts van het bronproject zijn om verschillende redenen teruggebracht tot deze 13:

- **Al native door het platform gedekt** — de grote familie scripts `endpoint-*-set.ps1` /
  `intune-*comp-set.ps1` / `intune-*ap-set.ps1` (Intune-beveiligingsbasislijnen,
  nalevingsbeleid, app-beveiligingsbeleid voor Windows/iOS/Android/macOS) hardcodeert
  de specifieke, eigenzinnige instellingen van één MSP. Microsofts eigen Intune Security Baseline en
  sjablonen voor nalevingsbeleid in het beheercentrum dekken dit nu native en worden door
  Microsoft actueel gehouden — een vaste basislijn uit 2020 in een script gieten levert geen voordeel op.
- **Al gedekt in deze repo** — basisrapportage over licenties/gebruikers/groepen
  (`o365-sku-audit-csv.ps1`, `graph-sku-get.ps1`, `o365-NoSPO-ADAcct*`), import van de CA-basislijn
  (`ca-policy-import.ps1` — zie `scripts/Entra/Import-ConditionalAccessBaseline.ps1`) en
  detectie van configuratiedrift in Intune (`endpoint-policy-get.ps1` — zie
  `scripts/Intune/Compare-IntuneConfig.ps1`) bestaan al en worden actief onderhouden.
- **Uitgefaseerde/verouderde modules zonder veilig modern equivalent voor die specifieke, beperkte
  functie** — `o365-add-domain.ps1` (MSOnline + hardgecodeerde inrichting van een Azure DNS-zone),
  de oorspronkelijke aanmeldflow van `graph-usrreg-read.ps1` met lokale XML-inloggegevens, `mcas-*.ps1` (het aparte
  authenticatiemodel met portaltokens van Defender for Cloud Apps, dat steeds meer opgaat in de uniforme
  Defender-portal) — de onderliggende *functie* van de laatste twee is behouden maar herschreven
  (MFA-rapport; audit van app-toestemmingen), de rest is geschrapt.
- **Niet overdraagbaar tussen tenants / weinig waarde voor een MSP die veel klanttenants beheert** —
  rapporten voor afstemming met on-prem AD (`o365-NoSPO-ADAcct*`, `o365-SPO-NoADAcct*`), WHOIS-
  opzoekingen (`o365-whois-get.ps1`), een generieke dump van DNS-records (`o365-dns-get.ps1` — teruggebracht
  tot de controle van SPF/DMARC waar je echt iets mee kunt), en lokale plumbing voor Hyper-V/menu's/cache van inloggegevens
  (`hyperv-*.ps1`, `start.ps1`, `*-connect.ps1`, `*-creds-save.ps1`).
- **Buiten het thema** — `reclaimwin10.ps1` en `win10-bp-get.ps1` zijn grote scripts voor het register/GPO
  van de lokale machine (één is een meegeleverd script van een derde partij, niet eens patrons eigen code)
  die op één Windows 10-werkstation werken, niet op tenantbreed M365-beheer.
- **Riskante bulkwijzigingen bewust buiten scope gehouden** — bulkimport/-verwijdering van Conditional Access-
  beleid, bulktoevoegen/-verwijderen van Intune-beleid/app-toewijzingen, en bulkimport/-verwijdering/-hertoewijzing
  van Autopilot-apparaten hadden allemaal *rapportage*functies die we wel hebben behouden (zie
  `Export-ConditionalAccessPolicies.ps1`, `Get-IntunePolicyAssignments.ps1`,
  `Get-AutopilotDevices.ps1`), maar hun wijzigende tegenhangers hebben we niet overgezet — dat zijn
  tenantspecifieke bewerkingen met een grote impact die je beter één voor één in het beheer-
  centrum beoordeelt dan generiek in een script giet.

Alle details over welk(e) bronscript(s) elk rapport precies vervangt, staan in de sectie
`.NOTES` van dat script.
