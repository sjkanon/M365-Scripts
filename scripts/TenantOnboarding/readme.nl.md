[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **TenantOnboarding**

# Tenant Onboarding

Scripts voor het onboarden van een nieuwe M365-tenant en het inrichten/beheren van de devices ervan, overgezet en gemoderniseerd vanuit een uitgefaseerde legacy-repository. Elk script volgt de huisstijl van deze repository: standaard een proefdraai met een `-Apply`-switch voor alles wat iets wijzigt, `[CmdletBinding(SupportsShouldProcess)]`, en geen hardcoded referenties, tenant-/klantnamen of interne endpoints.

Waar de legacy-bron meerdere bijna identieke scripts had die kleine varianten van hetzelfde deden (zes scripts om OneDrive te herstarten, zeven downloaders met hardcoded installers van leveranciers, vier scripts om lokale groepslidmaatschappen toe te kennen/in te trekken...), zijn die samengevoegd tot één goed geparametriseerd script in plaats van één-op-één overgezet.

---

## Mappen

| Map | Inhoud |
|--------|----------|
| [`Provisioning/`](Provisioning/readme.nl.md) | Bootstrap van één tenant: break-glass-beheeraccount, basisbeveiligingsgroepen, toewijzing van het Intune-basisbeleid |
| [`MultiTenant/`](MultiTenant/readme.nl.md) | MSP-brede scripts over klanten heen (GDAP): licentierapportage, rotatie van break-glass-wachtwoorden, index van klantportalen |
| [`AppDeployment/`](AppDeployment/readme.nl.md) | Generieke Win32-/Chocolatey-installers, standaard bestandskoppelingen, bureaubladsnelkoppelingen, netwerkprinters |
| [`DeviceConfig/`](DeviceConfig/readme.nl.md) | Zelfverhoging via lokale groepen, verharding van referentieopslag, energie-instellingen voor kiosken, verwijderen van Office, Startmenu-indeling, Teams-firewallregel |
| [`OneDriveManagement/`](OneDriveManagement/readme.nl.md) | Watchdog voor herstarten/resetten van OneDrive, per bibliotheek synchronisatie afbreken, omleiding via Known Folder Move |
| [`UserManagement/`](UserManagement/readme.nl.md) | Dynamische distributiegroepen aanmaken op basis van een filter, groepslidmaatschap voor het vrijgeven van functies |

Zie de readme van elke submap voor de volledige lijst met scripts, parameters en voorbeelden.

---

## Wat bewust is weggelaten

- **Voorbeeldrepository voor Microsoft Entra ID / macOS Intune** (`Install New Tenant/MacOS/`): een grote meegeleverde verzameling shellscripts, `.mobileconfig`-profielen en installers van derden voor macOS. Dit is geen eigen PowerShell-tooling van deze MSP en valt buiten de scope van een port in PowerShell-huisstijl.
- **`Manage Tenant/Compare/Compare-Intune.ps1`**: gecontroleerd en het blijkt dezelfde functionaliteit te zijn die al gedekt wordt door `scripts/Intune/Compare-IntuneConfig.ps1` (beide gebruiken `Compare-IntuneBackupDirectories` uit de module `IntuneBackupAndRestore` om de Intune-configuratie van een tenant te vergelijken met een baseline); niet opnieuw overgezet.
- Diverse rapportage-/configuratiescripts die elders in deze repository al gedekt zijn: OneDrive-/vergrendelschermachtergrond (`scripts/Intune/Desktop/Background/`), "vastmaken aan Start" (`scripts/Intune/Desktop/Add Lockscreen to start and desktop/`), tijdsynchronisatie (`scripts/Device/Time sync/`) en rapportage van de UniFi-controller (`scripts/Network/UniFi/`). De legacy-versies waren byte-voor-byte identiek of achterhaald door een algemenere versie die al in deze repository staat.
- Scripts die leunen op de uitgefaseerde modules **MSOnline** / **AzureAD** / **AzureADPreview** zijn niet ongewijzigd overgezet; hun functionaliteit is opnieuw geïmplementeerd met Microsoft Graph of Exchange Online (zie `MultiTenant/` en `Provisioning/`), of geschrapt waar de onderliggende techniek niet meer van toepassing is.
- Een handvol eenmalige, enkelvoudige of al verouderde scripts is als echte doodlopende weg geschrapt: een buildspecifieke registerfix voor Verkenner in Windows 11, beperkt tot builds onder 25211 (al lang gepatcht), twee dubbele/kapotte experimenten om "de standaard-pdf-app in te stellen + een ongerelateerde Start-indeling te importeren", een triviaal hulpscript voor een lokaal DPAPI-referentiebestand, en twee rapportagescripts van derden (niet door de MSP geschreven) waarvan de functionaliteit al gedekt wordt door `scripts/Exchange/Test-MailboxPermissions.ps1` en `scripts/Entra/Test-M365GroupMembership.ps1`.
- De bronmap `Setup tenant/` was leeg.

Er zijn nergens in deze map klantnamen, tenantdomeinen, interne hostnamen, IP-adressen of referenties uit de bronrepository gereproduceerd: elk script hier is geschreven voor generieke, geparametriseerde invoer.
