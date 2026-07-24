# DeviceConfig

Windows device-side configuration and hardening scripts used during tenant/device onboarding — local group self-elevation, credential storage hardening, kiosk power settings, Office removal, Start Menu layout, and the Teams LAN firewall rule.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Set-LocalGroupSelfElevation.ps1`](#set-localgroupselfelevationps1) | Grant/revoke a logged-on user's local group membership at next logon |
| [`Disable-CredentialManagerVault.ps1`](#disable-credentialmanagervaultps1) | Disable Windows Credential Manager's local password storage |
| [`Set-KioskPowerSettings.ps1`](#set-kioskpowersettingsps1) | Disable sleep/fast-startup for kiosk or always-on devices |
| [`Uninstall-MicrosoftOffice.ps1`](#uninstall-microsoftofficeps1) | Silently uninstall Microsoft Office / Microsoft 365 Apps |
| [`Import-StartMenuLayout.ps1`](#import-startmenulayoutps1) | Apply a Start Menu layout XML |
| [`Set-TeamsFirewallRule.ps1`](#set-teamsfirewallruleps1) | Create the inbound firewall rule Teams needs for LAN screen sharing |

---

### Set-LocalGroupSelfElevation.ps1

Consolidates four old scripts (grant/revoke local Administrators, grant/revoke local "Network Configuration Operators") into one. Registers a SYSTEM scheduled task, triggered at logon, that adds/removes whichever user is logged on to/from the named local group — and removes any pending opposing task.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-GroupName` | Yes | Local group name |
| `-Action` | Yes | `Grant` or `Revoke` |
| `-TaskName` | No | Scheduled task name (default: `<Action>-<GroupName>`) |
| `-Apply` | No | Actually register the task (default: preview only) |

```powershell
.\Set-LocalGroupSelfElevation.ps1 -GroupName "Administrators" -Action Grant -Apply
.\Set-LocalGroupSelfElevation.ps1 -GroupName "Administrators" -Action Revoke -Apply
.\Set-LocalGroupSelfElevation.ps1 -GroupName "Network Configuration Operators" -Action Grant -Apply
```

---

### Disable-CredentialManagerVault.ps1

Stops and disables the `VaultSvc` service, preventing local persistence of saved credentials.

```powershell
.\Disable-CredentialManagerVault.ps1 -Apply
```

---

### Set-KioskPowerSettings.ps1

Sets monitor/standby timeouts to never (AC+DC) and disables Fast Startup — for kiosk, reception, or always-on devices.

```powershell
.\Set-KioskPowerSettings.ps1 -Apply
```

---

### Uninstall-MicrosoftOffice.ps1

Finds and silently uninstalls Microsoft Office / Microsoft 365 Apps entries via their registered uninstall strings.

```powershell
.\Uninstall-MicrosoftOffice.ps1
.\Uninstall-MicrosoftOffice.ps1 -Apply
```

---

### Import-StartMenuLayout.ps1

Applies a Start Menu layout XML, from a local file or a URL.

| Parameter | Description |
|-----------|-------------|
| `-LayoutXmlPath` | Local layout XML path |
| `-LayoutUrl` | URL to download the layout XML from |
| `-Apply` | Actually apply (default: preview only) |

```powershell
.\Import-StartMenuLayout.ps1 -LayoutXmlPath "C:\Deploy\startmenu.xml" -Apply
.\Import-StartMenuLayout.ps1 -LayoutUrl "https://packages.contoso.com/config/startmenu.xml" -Apply
```

---

### Set-TeamsFirewallRule.ps1

Creates the inbound firewall rule allowing Teams LAN peer-to-peer screen sharing on the Domain profile (blocked on Public/Private), scoped to the currently logged-on user. Run as SYSTEM (Intune Win32 app or logon task).

```powershell
.\Set-TeamsFirewallRule.ps1 -Apply
```

> Original concept (c) Microsoft Corporation 2018 and Michael Mardahl (msendpointmgr.com), provided as-is; this is a house-style rewrite.

---

All scripts default to a dry run; pass `-Apply` to make changes, per house style.
