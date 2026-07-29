# Intune

Autopilot enrollment, compliance policy automation, configuration drift detection, and customer desktop deployment (wallpaper, lockscreen, taskbar lock shortcut).

> Office theme/color deployment lives in [`Custom Scripts/Intune/Desktop/`](../Custom%20Scripts/Intune/readme.md) — those scripts hardcode their download URL to that path.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Compare-IntuneConfig.ps1`](#compare-intuneconfigps1) | Compare a customer tenant's Intune configuration against an MSP baseline backup |
| [`Repair-StuckWin32AppEnforcement.ps1`](#repair-stuckwin32appenforcementps1) | Clear Win32 apps stuck behind Intune's GRS retry cooldown on a device — run manually, locally, with a dry-run/report/`-AppId` filter |
| [`Detect-StuckWin32AppEnforcement.ps1`](#detect--remediate-stuckwin32appenforcementps1) + [`Remediate-StuckWin32AppEnforcement.ps1`](#detect--remediate-stuckwin32appenforcementps1) | Same fix, packaged as an Intune Remediation pair — trigger entirely from the Intune portal, no device access needed |

## Folders

| Folder | Description |
|--------|-------------|
| [`Get-Autopilot/`](Get-Autopilot/readme.md) | Windows Autopilot hardware hash collection |
| [`iOS-Compliance-Updater/`](iOS-Compliance-Updater/readme.md) | Auto-updates the minimum iOS version in an Intune compliance policy |
| [`Desktop/`](Desktop/readme.md) | Corporate wallpaper + lockscreen, taskbar lock shortcut |
| [`DiskCleanup/`](DiskCleanup/readme.md) | Win32-app wrapper that runs the C:\ disk cleanup script and reboots the device |

---

### Compare-IntuneConfig.ps1

Wraps the community `IntuneBackupAndRestore` module to detect Intune configuration drift — compliance policies, configuration profiles, and other exported objects that differ from (or are missing/added relative to) a reference "baseline" backup. Read-only — never changes configuration, only reports differences.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-BaselinePath` | Yes | MSP reference backup folder (from a previous `Start-IntuneBackup -Path <path>` run) |
| `-CustomerBackupPath` | No | Existing backup of the customer tenant. If omitted, the script backs up the currently connected tenant first |
| `-TenantId` | No | Tenant ID or domain to connect to (only used when `-CustomerBackupPath` is omitted) |
| `-OutputPath` | No | Folder for the auto-backup and diff report (default: `C:\Temp\` / `~/Downloads`) |

**Examples**

```powershell
# Compare an existing customer backup against the MSP baseline
.\Compare-IntuneConfig.ps1 -BaselinePath C:\IntuneBaseline -CustomerBackupPath C:\Temp\CustomerBackup

# Back up the currently connected (GDAP) customer tenant and compare it live
.\Compare-IntuneConfig.ps1 -BaselinePath C:\IntuneBaseline
```

**Notes**
- GDAP-aware: resolves `-TenantId` automatically from the selected customer tenant (`$global:cid`) if omitted, same as `Move-InboxToArchive.ps1` / `Get-SharePointStorageReport.ps1`
- Not wired into `menu.ps1` — works with backup folder paths and a tenant-wide export, run it directly

**Required module**

```powershell
Install-Module IntuneBackupAndRestore -Scope CurrentUser
```

---

### Repair-StuckWin32AppEnforcement.ps1

Run **on the affected device**, as Administrator. Intune's Win32 app agent (IME) retries a failing install 3 times, 5 minutes apart, then locks the app into a 24-hour "GRS" cooldown — visible in `AppActionProcessor.log` as `... to install is in GRS. The app will not be enforced.` During that cooldown, IME ignores the app completely, even after you fix the actual problem in Intune (corrected requirement rule, new package, fixed install command) — the cooldown is local device state, not something a redeploy from Intune's side can clear.

This script finds every locally cached Win32 app with a real last error code (not 0/success, not 3010/reboot-pending) under `HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps`, and can clear its cached enforcement state, reporting cache, and matching GRS cooldown key, then restarts the IME service so every app is re-evaluated fresh on the next check-in.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-Apply` | Actually clear the stuck state and restart the service (default: dry-run report only) |
| `-AppId` | Limit to one specific Win32 app GUID (from the Intune admin center app URL). Default: every stuck app found |
| `-ForceSync` | After clearing, also trigger an immediate MDM check-in (same as Company Portal's "Sync" button) |
| `-OutputPath` | CSV report path (default: `C:\Temp\`) |

**Examples**

```powershell
# Dry run — see what's currently stuck on this device
.\Repair-StuckWin32AppEnforcement.ps1

# Clear everything stuck and force an immediate resync
.\Repair-StuckWin32AppEnforcement.ps1 -Apply -ForceSync

# Only clear one specific app
.\Repair-StuckWin32AppEnforcement.ps1 -Apply -AppId "96ff358d-0e16-4224-b946-eb61bc930fca"
```

**Notes**
- This is device-local state — it affects **every** Win32 app assigned to that device, not just one, so a single device stuck in GRS can look like several unrelated app deployments are all silently failing at once.
- Based on the community-documented (not officially published by Microsoft) GRS registry structure — see script `.NOTES` for sources.

---

### Detect- / Remediate-StuckWin32AppEnforcement.ps1

Same fix as `Repair-StuckWin32AppEnforcement.ps1` above, split into a detection/remediation pair for Intune **Devices → Scripts and remediations → Remediations**, so clearing a stuck device never needs a manual RDP/console session — everything is triggered from the Intune portal.

| Script | Role |
|---|---|
| `Detect-StuckWin32AppEnforcement.ps1` | Detection half — exit 1 if any Win32 app has a real last error code cached (possible GRS lock), exit 0 otherwise |
| `Remediate-StuckWin32AppEnforcement.ps1` | Remediation half — always clears everything the detection script found, restarts IME, forces an immediate MDM sync |

**Deploy in Intune:**

1. **Devices → Scripts and remediations → Remediations → Create**
2. Name it e.g. `Clear Stuck Win32 App Enforcement`
3. Upload `Detect-StuckWin32AppEnforcement.ps1` as the detection script, `Remediate-StuckWin32AppEnforcement.ps1` as the remediation script
4. Run using logged-on credentials: **No** (SYSTEM) · Run in 64-bit PowerShell: **Yes** · Enforce signature check: **No**
5. **Do not assign it to a device group or schedule.** Leave assignment empty.

**Trigger on-demand, per device, entirely from the portal** (no schedule, no device access):

1. **Devices → All devices → [the affected device]**
2. **… (ellipsis) → Run remediation (preview)**
3. Select `Clear Stuck Win32 App Enforcement`, run it

**Why unassigned/on-demand instead of a running schedule:** a scheduled, fleet-wide assignment would keep silently clearing GRS lockouts for *any* app that fails 3 times, for *any* reason — masking a genuinely broken deployment behind an endless auto-retry instead of surfacing it. Keeping it unassigned and only running it on a specific device once you've confirmed (via `AppActionProcessor.log`/`AppWorkload.log`, or the underlying problem is already fixed) that a retry is actually warranted avoids that. This mirrors the community-established pattern for this exact scenario (see script `.NOTES`), not something invented for this repo.
