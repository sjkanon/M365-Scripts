[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **LegacyUtilities**

# Legacy Utilities

Gemoderniseerde equivalenten, in de huisstijl van deze repository, van een reeks scripts uit een
uitgefaseerde interne repository. Geen van deze scripts is een letterlijke kopie van het
origineel: verschillende oude eenmalige scripts die kleine varianten van dezelfde taak
uitvoerden, zijn samengevoegd tot één goed geparametriseerd script, verouderde modules
(MSOnline/AzureAD) zijn vervangen door equivalenten in Microsoft Graph / Exchange Online, en
elke hardcoded klantnaam, tenantdomein, hostnaam, wachtwoord of secret uit de originelen is
omgezet in een parameter. Niets van die gegevens is meegenomen.

## Mappen

Deze map is per thema ingedeeld, met één submap per onderwerp:

| Map | Omschrijving |
|--------|----------|
| [`Exchange/`](Exchange/readme.nl.md) | Gedelegeerde mailboxtoegang, bulk aanmaken van gedeelde mailboxen/contacten, contactsynchronisatie, message trace, ontdubbelen van mailboxen |
| [`Entra/`](Entra/readme.nl.md) | Wijzigingen in groepslidmaatschap, back-up van Conditional Access-beleid |
| [`Teams/`](Teams/readme.nl.md) | Teams klonen, Planner-plannen kopiëren, bulk aanmaken van project-Teams |
| [`Network/`](Network/readme.nl.md) | Koppelen van Azure Files SMB-shares |
| [`Device/`](Device/readme.nl.md) | Standaardinstelling Num Lock, snelkoppeling Werkstation vergrendelen |
| [`Workspace365/`](Workspace365/readme.nl.md) | Provisioning/deprovisioning van Workspace 365-omgevingen |

Geen van deze scripts is al opgenomen in [`menu.ps1`](../../menu.ps1) of in de
documentatie op het hoogste niveau; die integratie volgt in een aparte ronde.

---

## Wat bewust is weggelaten

Een groot deel van het bronmateriaal is overgeslagen in plaats van overgezet: omdat het
elders in deze repository al volledig gedekt was, omdat het een echte doodlopende weg was,
of omdat het alleen bestond als echte klant-/tenantgegevens die nooit gereproduceerd mogen
worden. Zie het porteringsrapport voor het volledige overzicht; in het kort:

- **Elders in deze repository al gedekt**: verzamelen van Windows Autopilot-gegevens
  (`scripts/Intune/Get-Autopilot/`), algemene opschoning van schijf en logs
  (`scripts/Device/Invoke-WindowsCleanup.ps1`), verwijderen van OEM-bloatware
  (`scripts/Device/Remove-OemBloatware.ps1`), UniFi-netwerkrapportage en
  firmware-updates (`scripts/Network/UniFi/`), het CSP/GDAP-patroon voor verbinden met
  tenants (`scripts/Startup/functies.ps1`), agendamaprechten
  (`scripts/Exchange/Set-Calendar-rights.ps1`), en een reeks eenmalige device-/AppDeployment-scripts
  (URL-snelkoppelingen op het bureaublad, kopiëren van Startmenu-snelkoppelingen,
  verwijderen van Office, firewallregel voor Teams op het LAN) die al gegeneraliseerd zijn onder
  `scripts/TenantOnboarding/`.
- **Doodlopende wegen**: een complete mitigatiekit voor PrintNightmare (de spooler-CVE uit 2021),
  gebouwd rond de verouderde tool `subinacl.exe`. De kwetsbaarheid is al jaren gepatcht
  en de workaround heeft geen blijvende waarde.
- **Echte klant-/tenantgegevens, geen code**: diverse oude scripts en JSON-exports
  bevatten live tenant-ID's, GUID's van groeps-/gebruikersobjecten, toegangssleutels van
  opslagaccounts, client secrets van apps, of e-mailadressen/domeinen van klanten. Deze zijn,
  conform de regels voor gegevensverwerking van het project, nergens gereproduceerd (zelfs niet
  in rapporten); alleen de generieke *functionaliteit* erachter (bijv. "back-up maken van
  Conditional Access-beleid", "een Azure Files-share koppelen") is overgezet, waarbij alle
  identificerende waarden zijn omgezet in parameters.
