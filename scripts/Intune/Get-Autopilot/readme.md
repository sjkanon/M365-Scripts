**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [Intune](../readme.md) › **Get-Autopilot**

# Get-Autopilot

Windows Autopilot hardware hash collection — for USB/OOBE enrollment, see also [`Deployment/`](../../Deployment/readme.md), which copies these two files onto its USB toolkit.

---

## Files

| File | Description |
|------|-------------|
| [`Get-WindowsAutoPilotInfo.ps1`](Get-WindowsAutoPilotInfo.ps1) ([docs](#get-windowsautopilotinfops1)) | Microsoft script (Michael Niehaus, v3.5) with the online part rewritten for Microsoft Graph — retrieves the Autopilot hardware hash |
| [`GetAutoPilot.CMD`](GetAutoPilot.CMD) ([docs](#getautopilotcmd)) | Double-click wrapper — enables WinRM and runs the script, saving to `compHash.csv` |

---

### Get-WindowsAutoPilotInfo.ps1

The well-known community script for gathering Windows Autopilot device information (hardware hash, serial number, Windows Product ID) and optionally uploading it directly to Intune. Based on v3.5 by Michael Niehaus (Microsoft, MIT licence) — see the [PowerShell Gallery page](https://www.powershellgallery.com/packages/Get-WindowsAutoPilotInfo) for the original release notes.

This copy is **v3.5.1**: the `-Online` part talks to Microsoft Graph directly (`Invoke-MgGraphRequest` on `deviceManagement/importedWindowsAutopilotDeviceIdentities`, `windowsAutopilotDeviceIdentities`, `/devices` and `/groups/{id}/members/$ref`). v3.5 needed the retired **AzureAD** module (`-AddToGroup`) and **Microsoft.Graph.Intune** (`Connect-MSGraph`), so online import no longer worked; the gallery's latest version (3.9) still depends on the WindowsAutopilotIntune module. Only `Microsoft.Graph.Authentication` is needed now (installed for the current user when missing).

**Sign-in (`-Online`)** — **delegated by default**: you sign in as an Intune admin (`-DeviceCode` when no browser can open, e.g. OOBE). Scopes: `DeviceManagementServiceConfig.ReadWrite.All`, plus `GroupMember.ReadWrite.All` and `Device.Read.All` with `-AddToGroup`. **App-only** is an option: `-AppId` with `-CertificateThumbprint` (preferred) or `-AppSecret`, and `-TenantId`; the app needs the same permissions as application permissions.

The script deliberately stays standalone and Windows PowerShell 5.1-compatible: it is copied to a USB stick and started by `GetAutoPilot.CMD` / [`Deployment/start.bat`](../../Deployment/readme.md) with `powershell.exe`, so it does not load the repository's `Connect-M365.ps1`.

**Key parameters**

| Parameter | Description |
|-----------|-------------|
| `-Name` | Computer name(s) to collect from (default: `localhost`); accepts pipeline input |
| `-OutputFile` | CSV path to write the hash to |
| `-Append` | Append to `-OutputFile` instead of overwriting |
| `-Credential` | Credentials for connecting to remote computers |
| `-Partner` | Use CSP partner-center registration flow |
| `-GroupTag` | Autopilot group tag to assign |
| `-Online` | Upload the hash directly to Intune instead of (or in addition to) writing a CSV |
| `-TenantId` | Tenant for `-Online` (required for app-only; delegated defaults to your sign-in tenant) |
| `-AppId` / `-CertificateThumbprint` / `-AppSecret` | App-only sign-in for `-Online` mode (certificate preferred) |
| `-DeviceCode` | Delegated sign-in with a device code (OOBE) |
| `-AssignedUser` | Pre-assign a user to the device in Intune |
| `-AssignedComputerName` | Pre-assign a computer name (`-Online` mode) |
| `-AddToGroup` | Add the device to an Entra ID group after import (`-Online` mode) |
| `-Assign` | Wait for and display the Autopilot profile assignment (`-Online` mode) |
| `-Reboot` | Reboot after a successful online import + assign |

**Examples**

```powershell
# Save hash to CSV
.\Get-WindowsAutoPilotInfo.ps1 -OutputFile compHash.csv

# Upload directly to Intune (interactive sign-in)
.\Get-WindowsAutoPilotInfo.ps1 -Online

# From OOBE: device code, add to a group, wait for the profile, reboot
.\Get-WindowsAutoPilotInfo.ps1 -Online -DeviceCode -GroupTag "Corporate" -AddToGroup "Autopilot Devices" -Assign -Reboot

# Upload with a group tag and app-only sign-in (certificate)
.\Get-WindowsAutoPilotInfo.ps1 -Online -GroupTag "Corporate" -TenantId "..." -AppId "..." -CertificateThumbprint "..."
```

**Notes**
- The CSV now has exactly the columns Intune's import accepts (`Device Serial Number`, `Windows Product ID`, `Hardware Hash`, plus `Group Tag` / `Assigned User` when given). The earlier copy added make/model and a second `Hardware Hash` column, which `Select-Object` rejects.
- The import and sync loops reported the last device for every device and could hang on a device that failed to import; each device is now checked on its own.
- Reading the hardware hash needs an elevated session.

---

### GetAutoPilot.CMD

Double-click wrapper for OOBE / technician use — no PowerShell knowledge required:

1. Enables WinRM (`Enable-PSRemoting -SkipNetworkProfileCheck -Force`)
2. Runs `Get-WindowsAutoPilotInfo.ps1 -ComputerName $env:computername -OutputFile compHash.csv -Append` from the same folder
3. Pauses so the console stays open to read the result

```
GetAutoPilot.CMD
```

> Run both files from the same folder — `%~dp0` resolves relative to the `.CMD`'s own location.
