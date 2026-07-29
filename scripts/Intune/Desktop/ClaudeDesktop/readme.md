# ClaudeDesktop

Machine-wide Intune deployment of [Claude Desktop](https://claude.com/download) for Windows, including the Virtual Machine Platform feature required for Cowork. Run **one script once a month** to keep the Intune app current — no persistent App Registration or client secret to manage.

Based on [Deploy Claude Desktop for Windows](https://support.claude.com/en/articles/12622703-deploy-claude-desktop-for-windows) and [Enterprise configuration for Claude Desktop](https://support.claude.com/en/articles/12622667-enterprise-configuration-for-claude-desktop).

---

## Why not just upload the MSIX as a line-of-business app?

Intune installs LOB MSIX apps per-user. That fails for standard users without admin rights and doesn't satisfy Cowork's machine-wide requirement. Instead, `Add-AppxProvisionedPackage` is wrapped as a Win32 app, exactly as the official article recommends.

## Contents

| Script | Role in Intune |
|---|---|
| `Deploy-ClaudeDesktopIntune.ps1` | **The one you run.** Orchestrates everything below — see "Monthly run" section. |
| `Install-ClaudeDesktop-Intune.ps1` | Install command content script |
| `Uninstall-ClaudeDesktop-Intune.ps1` | Uninstall command content script |
| `Detect-ClaudeDesktop-Intune.ps1` | Custom detection script |

`Install-`/`Uninstall-`/`Detect-ClaudeDesktop-Intune.ps1` are never run manually — `Deploy-ClaudeDesktopIntune.ps1` packages them into the `.intunewin` (or, for the detection script, uploads it as part of the detection rule) automatically.

## What `Deploy-ClaudeDesktopIntune.ps1` does

1. Downloads the latest Claude Desktop x64 MSIX from Anthropic's official "latest" redirect URL.
2. Reads the version out of `AppxManifest.xml` inside the MSIX.
3. Copies the install/uninstall scripts next to the MSIX and builds a `.intunewin` package (`IntuneWin32App` module — `IntuneWinAppUtil.exe` is downloaded automatically if not already present).
4. Connects to Microsoft Graph delegated (interactive sign-in) and creates a short-lived **temporary App Registration** — same pattern as [`Remove-SharePointFileVersionsByDate.ps1`](../../../Reporting/readme.md) — with only the `DeviceManagementApps.ReadWrite.All` application permission. Used to authenticate the `IntuneWin32App` module, then deleted at the end of the run. Nothing persists between runs except the Intune app itself.
5. First run: creates the Win32 app "Claude Desktop (Machine-wide)" in Intune with detection/requirement rules and assigns it **Required** to the Entra ID group you pass in.
6. Later runs: pushes an updated package via `Update-IntuneWin32AppPackageFile` (existing assignment left untouched, devices just get the new content) if **either** the downloaded MSIX version is newer, **or** the Install-/Uninstall-/Detect-ClaudeDesktop-Intune.ps1 scripts themselves changed since the last run — both tracked in the app's Notes field (`ClaudeMsixVersion=...; ScriptsHash=...`), no local state file needed. Detection **and requirement rules are rebuilt and resubmitted on every run**, not just at first creation (see "Known issue" below). If neither the version nor the scripts changed, only the rules get refreshed.

Detection is deliberately version-agnostic (presence of the provisioned package + `VirtualMachinePlatform` enabled). Intune redeploys a Win32 app to already-targeted devices whenever its content version changes in Intune, regardless of what the detection rule reports — so there's no `$MinimumVersion` to bump by hand every month, and a pure script edit (no new MSIX) still triggers a redeploy via the `ScriptsHash` check above.

### Known issue: 0x80070001 install failures fixed by resubmitting the requirement rule

Devices failed installation with error `0x80070001` regardless of the device. Comparing the cached Win32 app policy on an affected device (`AppWorkload.log`) against every other Win32 app assigned to the same tenant showed the anomaly: every other app's `RequirementRules` had `RequiredOSArchitecture: 3`, but Claude Desktop's had `RequiredOSArchitecture: 32` — a value no other app in the tenant used, and 16x what `-Architecture 'x64'` should produce. Since this is stored on the app object in Intune (not per-device), it explained why the failure was 100% reproducible across every device, not device-specific corruption.

The root cause: `Set-IntuneWin32App` in the update branch never passed `-RequirementRule` (only `Add-IntuneWin32App` did, at first creation) — so whatever got recorded on that very first run, right or wrong, stayed on the app forever, immune to every later redeploy. Fixed by rebuilding and resubmitting both `-DetectionRule` and `-RequirementRule` on every run (see point 6 above). Run the script again to push the corrected requirement rule to the existing app.

**If a device still fails after that fix**: check whether it's actually stuck behind Intune's unrelated **GRS retry cooldown** instead (3 failed attempts → 24h lockout, regardless of the app config) — see [`Repair-StuckWin32AppEnforcement.ps1`](../../readme.md#repair-stuckwin32appenforcementps1) one level up. This affects every Win32 app on that device, not just Claude, so it's a useful first check if several unrelated apps are also silently stuck.

### Company Portal visibility

The deploy script extracts the app's logo directly from the downloaded MSIX (`Properties/Logo` in `AppxManifest.xml`, falling back to the highest-scale variant actually present in the package) and sets it as the Win32 app's icon, plus `-CompanyPortalFeaturedApp $true` on every run — so instead of a generic Win32 icon buried in the full app list, users see the real Claude logo, featured, in Company Portal.

### Required role

Global Administrator, or Application Administrator combined with a role that can grant `AppRoleAssignment.ReadWrite.All` consent — the same requirement as the temporary App Registration in the SharePoint reporting scripts.

### Monthly run

```powershell
.\Deploy-ClaudeDesktopIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop"
```

You'll get an interactive sign-in prompt and a "type JA to continue" confirmation before anything is created/updated in Intune. Add `-Force` to skip the confirmation (e.g. for a scheduled/unattended run once you're comfortable with the flow).

| Parameter | Default | Description |
|---|---|---|
| `-AssignmentGroupName` | *(required)* | Entra ID group assigned Required. Only used on first creation. |
| `-WorkingDirectory` | `C:\Temp\ClaudeDeploy` | Download/build staging folder (`Source/`, `Output/`) |
| `-AppDisplayName` | `Claude Desktop (Machine-wide)` | Used to find the existing app on later runs — don't change without renaming in Intune too |
| `-MsixDownloadUrl` | Anthropic's official x64 "latest" redirect | Override for testing |
| `-MinimumSupportedWindowsRelease` | `W10_21H2` | Requirement rule |
| `-TenantId` | auto-detected | Entra ID tenant ID |
| `-IntuneWinAppUtilPath` | auto-download | Use an already-downloaded `IntuneWinAppUtil.exe` |
| `-Force` | off | Skip the confirmation prompt(s) |

### Auto-update policy

The install script sets `HKLM:\SOFTWARE\Policies\Claude\disableAutoUpdates = 1` (DWord). Claude's own updater is disabled so version control stays entirely with this monthly Intune run instead of drifting per-device.

### VirtualMachinePlatform restart

If `VirtualMachinePlatform` was already enabled on the device, nothing else happens — Cowork works immediately. If the install script has to enable it for the first time, it does **not** restart the device itself: it sends an English `msg.exe` notification to the active console session (the logged-in user, not a broadcast to every session) telling them to restart when convenient. The feature only becomes fully active after that restart.

### Clean reinstall on every run

Before provisioning the new version, the install script fully removes Claude Desktop from the device first, in three passes:

1. Stops any running Claude process.
2. Removes every per-user **Appx** installation (`Get-AppxPackage -AllUsers` / `Remove-AppxPackage -AllUsers`, including already-logged-in profiles), then the old machine-wide provisioned package.
3. Removes any **classic (non-Appx) per-user installation** — most notably the consumer installer from [claude.ai/download](https://claude.ai/download), which registers itself through an ordinary per-user Uninstall registry key rather than as an Appx package, so `Get-AppxPackage` never sees it. The script scans the Uninstall registry key of every local profile — including profiles that aren't currently logged in, by temporarily loading their `NTUSER.DAT` — and runs each match's `QuietUninstallString` (or `UninstallString` if that's absent), with a 120s timeout so a stuck installer can't hang the Intune install.

Only after all three passes does it provision the new MSIX. This is deliberately more thorough than just clearing the provisioning layer — any installation left over from a different route can keep running its own, non-Cowork-registered Claude session even after the machine-wide version was updated. A user with Claude open loses that session when this runs.

### Robustness on a never-installed device

Both content scripts are written so a device where Claude has never been present — including the very first Autopilot ESP run — goes through cleanly:

- **Install script**: every removal pass (process kill, Appx, classic per-user uninstall) checks for an empty/`$null` result before acting, so a completely clean machine just logs "none found" at each step instead of erroring. If `VirtualMachinePlatform` has to be enabled for the first time and there's no active console session yet (the normal case mid-Autopilot ESP, since ESP itself restarts the device before handing over to the user), the restart notification is skipped with a log line explaining why — no forced restart is invented to compensate, since ESP's own end-of-provisioning restart already covers it.
- **Detection script**: an empty/negative result from `Get-AppxProvisionedPackage` or `Get-WindowsOptionalFeature` (the expected outcome on a never-installed device) reports "not installed" (exit 1) immediately, with no retry — retries only kick in on an actual **exception** from either cmdlet (e.g. a transient DISM lock, plausible right after the install script's own heavy Appx/DISM activity on the same device), so a genuinely clean machine is never slowed down waiting on retries that can't change the outcome.

### Prerequisites

- `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications`, `Microsoft.Graph.Groups`, `IntuneWin32App` PowerShell modules — install with `.\scripts\Startup\Install-Modules.ps1`
- Run from Windows (the packaging tool and MSIX/AppX cmdlets are Windows-only)
