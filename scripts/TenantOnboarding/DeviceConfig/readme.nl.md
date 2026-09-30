[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [TenantOnboarding](../readme.nl.md) › **DeviceConfig**

# DeviceConfig

Scripts voor configuratie en verharding van Windows-devices, gebruikt tijdens het onboarden van tenants/devices: zelfverhoging via lokale groepen, verharding van referentieopslag, energie-instellingen voor kiosken, verwijderen van Office, Startmenu-indeling en de Teams-firewallregel voor het LAN.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Set-LocalGroupSelfElevation.ps1`](Set-LocalGroupSelfElevation.ps1) ([docs](#set-localgroupselfelevationps1)) | Ken bij de volgende aanmelding het lidmaatschap van een lokale groep toe aan de aangemelde gebruiker, of trek het in |
| [`Disable-CredentialManagerVault.ps1`](Disable-CredentialManagerVault.ps1) ([docs](#disable-credentialmanagervaultps1)) | Schakel de lokale wachtwoordopslag van Windows Referentiebeheer uit |
| [`Set-KioskPowerSettings.ps1`](Set-KioskPowerSettings.ps1) ([docs](#set-kioskpowersettingsps1)) | Schakel slaapstand/snel opstarten uit voor kiosk- of altijd-aan-devices |
| [`Uninstall-MicrosoftOffice.ps1`](Uninstall-MicrosoftOffice.ps1) ([docs](#uninstall-microsoftofficeps1)) | Verwijder Microsoft Office / Microsoft 365 Apps op de achtergrond |
| [`Import-StartMenuLayout.ps1`](Import-StartMenuLayout.ps1) ([docs](#import-startmenulayoutps1)) | Pas een XML-bestand met Startmenu-indeling toe |
| [`Set-TeamsFirewallRule.ps1`](Set-TeamsFirewallRule.ps1) ([docs](#set-teamsfirewallruleps1)) | Maak de inkomende firewallregel die Teams nodig heeft voor schermdelen op het LAN |

---

### Set-LocalGroupSelfElevation.ps1

Voegt vier oude scripts samen tot één (lokale Administrators toekennen/intrekken, lokale "Network Configuration Operators" toekennen/intrekken). Registreert een geplande SYSTEM-taak die bij het aanmelden start en de aangemelde gebruiker, wie dat ook is, toevoegt aan of verwijdert uit de opgegeven lokale groep, en verwijdert elke nog openstaande tegengestelde taak.

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-GroupName` | Ja | Naam van de lokale groep |
| `-Action` | Ja | `Grant` of `Revoke` |
| `-TaskName` | Nee | Naam van de geplande taak (standaard: `<Action>-<GroupName>`) |
| `-Apply` | Nee | Registreer de taak echt (standaard: alleen voorbeeldweergave) |

```powershell
.\Set-LocalGroupSelfElevation.ps1 -GroupName "Administrators" -Action Grant -Apply
.\Set-LocalGroupSelfElevation.ps1 -GroupName "Administrators" -Action Revoke -Apply
.\Set-LocalGroupSelfElevation.ps1 -GroupName "Network Configuration Operators" -Action Grant -Apply
```

---

### Disable-CredentialManagerVault.ps1

Stopt de service `VaultSvc` en schakelt hem uit, zodat opgeslagen referenties niet meer lokaal bewaard worden.

```powershell
.\Disable-CredentialManagerVault.ps1 -Apply
```

---

### Set-KioskPowerSettings.ps1

Zet de time-outs voor beeldscherm en stand-by op nooit (netstroom en accu) en schakelt Snel opstarten uit, voor kiosk-, receptie- of altijd-aan-devices.

```powershell
.\Set-KioskPowerSettings.ps1 -Apply
```

---

### Uninstall-MicrosoftOffice.ps1

Zoekt vermeldingen van Microsoft Office / Microsoft 365 Apps en verwijdert ze op de achtergrond via hun geregistreerde uninstall-strings.

```powershell
.\Uninstall-MicrosoftOffice.ps1
.\Uninstall-MicrosoftOffice.ps1 -Apply
```

---

### Import-StartMenuLayout.ps1

Past een XML-bestand met Startmenu-indeling toe, vanuit een lokaal bestand of een URL.

| Parameter | Omschrijving |
|-----------|-------------|
| `-LayoutXmlPath` | Lokaal pad naar de indelings-XML |
| `-LayoutUrl` | URL waarvan de indelings-XML wordt gedownload |
| `-Apply` | Pas echt toe (standaard: alleen voorbeeldweergave) |

```powershell
.\Import-StartMenuLayout.ps1 -LayoutXmlPath "C:\Deploy\startmenu.xml" -Apply
.\Import-StartMenuLayout.ps1 -LayoutUrl "https://packages.contoso.com/config/startmenu.xml" -Apply
```

---

### Set-TeamsFirewallRule.ps1

Maakt de inkomende firewallregel die peer-to-peer schermdelen in Teams op het LAN toestaat op het domeinprofiel (geblokkeerd op openbaar/privé), beperkt tot de op dat moment aangemelde gebruiker. Voer uit als SYSTEM (Intune Win32-app of aanmeldtaak).

```powershell
.\Set-TeamsFirewallRule.ps1 -Apply
```

> Oorspronkelijk concept (c) Microsoft Corporation 2018 en Michael Mardahl (msendpointmgr.com), geleverd zoals het is; dit is een herschrijving in huisstijl.

---

Alle scripts doen standaard een proefdraai; geef `-Apply` mee om wijzigingen door te voeren, volgens de huisstijl.
