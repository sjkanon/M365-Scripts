[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../readme.nl.md) › **scripts**

# scripts/

Alle PowerShell-tooling van deze repo, gegroepeerd per workload. Start alles vanuit de root via [`.\menu.ps1`](../menu.ps1) (of [`.\load.ps1`](../load.ps1) bij de eerste keer) — zie de [root-readme](../readme.nl.md) voor de volledige menureferentie en de stappen om aan de slag te gaan.

---

## Op zoek naar één specifiek script?

[**INDEX.md**](INDEX.md) zet ze allemaal van A tot Z op één pagina — script, map en wat het doet — zodat je met Ctrl-F kunt zoeken in plaats van te raden in welke map het staat. Die lijst wordt uit de scripts zelf gegenereerd door [`Startup/Update-ScriptIndex.ps1`](Startup/Update-ScriptIndex.ps1); draai dat opnieuw nadat je een script hebt toegevoegd, hernoemd of verwijderd.

De tabel hieronder werkt andersom: waar elke categorie *voor* is.

---

## Mappen

| Map | Omschrijving |
|--------|-------------|
| [`ActiveDirectory/`](ActiveDirectory/readme.nl.md) | Monitoring van on-prem AD DS (bewaking van accountvergrendelingen) — richt zich rechtstreeks op een DC/fileserver, niet op Entra ID |
| [`Azure/`](Azure/readme.nl.md) | Beheer van Azure IaaS-VM's (conversie van de schijfcontroller) — richt zich rechtstreeks op Azure via `Az`, niet op de M365-tenant |
| [`Entra/`](Entra/readme.nl.md) | Levenscyclus van gebruikers, managers toewijzen, licentierapportage, Conditional Access-baseline, tijdelijke CA-vensters, TAP-codes, audit van M365-groepen (Microsoft Graph) |
| [`Exchange/`](Exchange/readme.nl.md) | Agendamigratie/-rechten, distributiegroepen, audits van mailboxen/agenda's/DKIM/doorsturen |
| [`Graph/`](Graph/readme.nl.md) | Beheer van Microsoft Graph-applicatierechten |
| [`Intune/`](Intune/readme.nl.md) | Autopilot-inschrijving, updater voor iOS-compliancebeleid, uitrol van bedrijfsachtergrond/-vergrendelscherm |
| [`SharePoint/`](SharePoint/readme.nl.md) | Contentbewerkingen in SharePoint Online / OneDrive — prullenbak herstellen per site of tenantbreed (PnP PowerShell, automatische app-registratie), en waar een bestand gebleven is: hernoemd, verplaatst of verwijderd (auditlog) |
| [`Reporting/`](Reporting/readme.nl.md) | Rapport laatste aanmelding van computers, SharePoint-opslagrapport, maandelijks licentierapport |
| [`Device/`](Device/readme.nl.md) | Onderhoud van Windows-endpoints — activatie, opschoning, tijdelijke bestanden, tijdsynchronisatie, audio, OpenVPN-diagnose, tijdelijke schijf + pagefile voor Azure/AVD |
| [`Network/`](Network/readme.nl.md) | TCP-poortcontroles, authenticatie-/netwerkdiagnose, stresstests voor bestands-I/O |
| [`RDS/`](RDS/readme.nl.md) | Diagnose van RDP-/RD Web Access-aanmeldingen, live sessiemonitoring, diagnose en verkleining van FSLogix-profielschijven |
| [`SMTP/`](SMTP/readme.nl.md) | Connectiviteitstests voor SMTP-relay (eenmalig en terugkerend) |
| [`Deployment/`](Deployment/readme.nl.md) | USB-toolkit voor Windows-installatie en Autopilot-inschrijving tijdens OOBE |
| [`DNS/`](DNS/readme.nl.md) | DNS-records resolven en importeren in AD-geïntegreerde DNS-zones |
| [`SAS/`](SAS/readme.nl.md) | Foutmonitoring van SAS-batchjobs met Zabbix-integratie |
| [`Teams/`](Teams/readme.nl.md) | Export en archivering van Microsoft Teams / SharePoint |
| [`Startup/`](Startup/readme.nl.md) | `functies.ps1` M365-functiebibliotheek + module-bootstrap + syntaxcontrole, dot-sourced door het menu |
| [`Custom Scripts/`](Custom%20Scripts/readme.nl.md) | Scripts met een vastgepind pad — uitrol van het Office-thema (de download-URL naar dit repopad is hardcoded) |
| [`TenantOnboarding/`](TenantOnboarding/readme.nl.md) | Inrichting van nieuwe tenants, multi-tenant-/GDAP-rapportage, app-uitrol, apparaatconfiguratie, OneDrive-beheer, gebruikersbeheer — gemoderniseerd uit een uitgefaseerde interne toolkit voor tenantinrichting |
| [`Office365Toolkit/`](Office365Toolkit/readme.nl.md) | Herschreven Security-/Exchange-/Intune-functionaliteit die nog nuttig is uit de uitgefaseerde `directorcia/Office365`-toolkit (CIAOPS) |
| [`PatronToolkit/`](PatronToolkit/readme.nl.md) | Herschreven Entra-/Exchange-/Intune-/Security-/SharePoint-/Teams-functionaliteit die nog nuttig is uit de uitgefaseerde `directorcia/patron`-toolkit |
| [`LegacyUtilities/`](LegacyUtilities/readme.nl.md) | Diverse gemoderniseerde scripts (Exchange, Entra, Teams, Network, Device, Workspace 365) uit allerlei kleine tools in de uitgefaseerde interne toolkit |
