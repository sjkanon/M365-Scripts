# Get-Autopilot

Windows Autopilot hardware hash collection — for USB/OOBE enrollment, see also [`Custom Scripts/Save install time/`](../../Custom%20Scripts/Save%20install%20time/readme.md), which copies these two files onto its USB toolkit.

---

## Files

| File | Description |
|------|-------------|
| [`Get-WindowsAutoPilotInfo.ps1`](#get-windowsautopilotinfops1) | Community script (Michael Niehaus) — retrieves the Autopilot hardware hash |
| [`GetAutoPilot.CMD`](#getautopilotcmd) | Double-click wrapper — enables WinRM and runs the script, saving to `compHash.csv` |

---

### Get-WindowsAutoPilotInfo.ps1

The well-known community script for gathering Windows Autopilot device information (hardware hash, serial number, Windows Product ID) and optionally uploading it directly to Intune. Currently v3.5 by Michael Niehaus (Microsoft) — see the [PowerShell Gallery page](https://www.powershellgallery.com/packages/Get-WindowsAutoPilotInfo) for full release notes.

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
| `-TenantId` / `-AppId` / `-AppSecret` | App-based authentication for `-Online` mode |
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

# Upload with a group tag and app-based auth
.\Get-WindowsAutoPilotInfo.ps1 -Online -GroupTag "Corporate" -TenantId "..." -AppId "..." -AppSecret "..."
```

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
