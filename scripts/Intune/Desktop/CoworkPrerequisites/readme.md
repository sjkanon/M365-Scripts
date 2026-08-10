# CoworkPrerequisites

Machine-wide Intune deployment of the Windows-side prerequisites for [Claude Cowork](https://support.claude.com/en/articles/12622667-enterprise-configuration-for-claude-desktop) — the `VirtualMachinePlatform` optional feature and disabling Windows Fast Startup — packaged as its **own, independent** Win32 app, separate from [`../ClaudeDesktop/`](../ClaudeDesktop/readme.md).

## Why a separate app instead of bundling it into Claude Desktop's install script

Cowork is optional: Claude Desktop works fine without it. Bundling the Windows-feature step into Claude Desktop's own install script means a DISM hiccup unrelated to Claude (a busy servicing stack, Windows Update unreachable as a feature source — both more likely on a freshly imaged device mid-Autopilot-ESP) could abort the *entire* Claude Desktop install. Splitting it out means:

- A Cowork-prerequisites failure never blocks Claude Desktop itself.
- Each app has its own, separately visible install status in Intune — you can tell at a glance whether a device's problem is "Windows feature" or "Claude Desktop", instead of one opaque combined install command.

There is deliberately **no Intune "Dependency"** configured between the two apps by default. A hard dependency would mean Claude Desktop doesn't even attempt to install until this app succeeds — which reintroduces the exact problem above. Assign both **Required** to the same group, independently.

If your environment treats Cowork as a hard requirement rather than optional, deploy this app first, then run `Deploy-ClaudeDesktopIntune.ps1 -RequireCoworkPrerequisites` — see [`../ClaudeDesktop/readme.md`](../ClaudeDesktop/readme.md#optional-requiring-cowork-prerequisites) for what that adds and the tradeoff involved.

## Contents

| Script | Role in Intune |
|---|---|
| `Deploy-CoworkPrerequisitesIntune.ps1` | **The one you run** for the Win32-app route. Packages the scripts below and creates/updates the Win32 app. |
| `Install-CoworkPrerequisites-Intune.ps1` | Win32-app install command content script |
| `Uninstall-CoworkPrerequisites-Intune.ps1` | Win32-app uninstall command content script |
| `Detect-CoworkPrerequisites-Intune.ps1` | Win32-app custom detection script |
| `CoworkPrerequisites-PlatformScript.ps1` | **Alternative**, standalone route — no Deploy script, no packaging, uploaded directly as an Intune "Platform script". See "Platform script alternative" below. |

There's no MSIX here — unlike Claude Desktop, the "content" is just these three scripts, so a re-run only does something if you've actually edited one of them (tracked via a `ScriptsHash` in the app's Notes field, same pattern as `Deploy-ClaudeDesktopIntune.ps1`).

## What the install script does

1. **VirtualMachinePlatform**: `Enable-WindowsOptionalFeature`, with retry against transient DISM failures. No-op if already enabled.
2. **Fast Startup**: sets `HiberbootEnabled = 0` under `HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power`, on every run (not just when VMP was just enabled). Anthropic's own Cowork documentation explicitly warns: *"Restart the machine using Restart, not shut down and power on. With Windows Fast Startup enabled, a shutdown cycle can leave the virtualization services uninitialized."* Fast Startup is on by default on nearly every Windows image — without disabling it, a user who just "shuts down" instead of "restarts" (the common case) never gets working Cowork services, no matter how many power cycles happen, even though Intune shows the app as installed.
3. If VMP was just enabled: exits with code **3010** ("soft reboot required"). The app is configured with `-RestartBehavior 'basedOnReturnCode'`, so **Intune itself** enforces the restart (prompt/deadline/grace period) — this doesn't depend on a user being logged in. As a courtesy it also sends an English `msg.exe` notification to the active console session (recognized via the untranslated `SESSIONNAME` "console", not the OS-language-dependent `STATE` text "Active" — a plain string match on "Active" would silently never fire on a non-English Windows display language), but the notification is a nicety, not the mechanism the restart actually relies on.

   The notification is **not** sent by calling `msg.exe` directly from this SYSTEM-context script — that produces a dialog whose OK button doesn't respond to clicks (a known `msg.exe` quirk with cross-session messages from a non-interactive sender). Instead, `Show-UserRestartNotification` registers a short-lived scheduled task (`LogonType Interactive`, principal = the console session's own user) that runs `msg.exe` *inside the user's own session*, then unregisters it again once sent — the dialog then belongs to a genuinely interactive desktop and closes normally.

## What the detection script checks

Reports "installed" only if **both** `VirtualMachinePlatform` is `Enabled` **and** the underlying HCS services (`vmcompute`, `HNS`, `vfpext`) are present — VMP showing `Enabled` in DISM doesn't guarantee these services already exist (see Anthropic's own Cowork troubleshooting: *"Missing HCS services: HNS, vmcompute, vfpext"*), particularly right after enabling VMP but before the required restart.

Deliberately **no check on service `Status`** (e.g. `Running`): `vmcompute` is a trigger-start service and is expected to be `Stopped` whenever no Cowork session is active — checking for `Running` would report a perfectly healthy device as "not installed". Only the service being completely absent (`Get-Service` can't find it) is a reliable signal that the underlying Hyper-V components aren't there yet.

## Known issue: `-RequirementRule` on `Set-IntuneWin32App` silently no-ops (fixed)

Same confirmed `IntuneWin32App`-module (1.5.0) bug as [`../ClaudeDesktop/readme.md`](../ClaudeDesktop/readme.md#known-issue-0x80070001-install-failures--and-a-much-deeper-intunewin32app-module-bug-behind-it): `Set-IntuneWin32App`'s `-RequirementRule` parameter incorrectly demands an `@odata.type` property that `New-IntuneWin32AppRequirementRule` never sets, and the `break` that follows terminates the **entire rest of the function** — meaning every update call that included `-RequirementRule` silently skipped `Notes`, `DetectionRule`, and `RestartBehavior` too, not just the architecture requirement. `Add-IntuneWin32App` (first creation) doesn't have this bug.

Fixed the same way: the update branch no longer passes `-RequirementRule` to `Set-IntuneWin32App`, and instead calls `Set-Win32AppArchitectureRequirement` — a direct Graph PATCH for `allowedArchitectures`/`minimumSupportedWindowsRelease`, reusing the session `Connect-MSIntuneGraph` already established.

## Company Portal visibility

`-CompanyPortalFeaturedApp $true` is set on every run (both first creation and updates), so this shows up featured in Company Portal instead of staying invisible as a background-only prerequisite. There's no MSIX here to extract a logo from, so it uses Intune's default Win32 app icon — set `-Icon` manually in the Intune portal afterward if you want a custom one.

## Monthly / as-needed run

```powershell
.\Deploy-CoworkPrerequisitesIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop"
```

Same confirmation/`-Force` behavior as `Deploy-ClaudeDesktopIntune.ps1`. Since there's no MSIX to version-bump, re-running this only pushes an update if you've edited one of the three content scripts — otherwise it's a no-op.

| Parameter | Default | Description |
|---|---|---|
| `-AssignmentGroupName` | *(required)* | Entra ID group assigned Required. Only used on first creation. |
| `-WorkingDirectory` | `C:\Temp\CoworkPrereqDeploy` | Build staging folder (`Source/`, `Output/`) |
| `-AppDisplayName` | `Cowork Windows Prerequisites (Machine-wide)` | Used to find the existing app on later runs — don't change without renaming in Intune too |
| `-MinimumSupportedWindowsRelease` | `W10_21H2` | Requirement rule |
| `-TenantId` | auto-detected | Entra ID tenant ID |
| `-IntuneWinAppUtilPath` | auto-download | Use an already-downloaded `IntuneWinAppUtil.exe` |
| `-Force` | off | Skip the confirmation prompt(s) |

## Uninstall

`VirtualMachinePlatform` is **not** disabled by default (other applications — WSL, Hyper-V-based tools, other Cowork-like apps — may depend on it too). Pass `-DisableVirtualMachinePlatform` to the uninstall script if you're certain nothing else on the device needs it. Fast Startup is left disabled either way — it's a harmless, device-wide setting, not something specific to Cowork.

## Prerequisites

- `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications`, `Microsoft.Graph.Groups`, `IntuneWin32App` PowerShell modules — install with `.\scripts\Startup\Install-Modules.ps1`
- Run from Windows (the packaging tool and DISM cmdlets are Windows-only)

## Platform script alternative

`CoworkPrerequisites-PlatformScript.ps1` is the same enable-VMP-and-disable-Fast-Startup logic, adapted to be uploaded directly as an Intune **Platform script** (Devices → Scripts and remediations → Platform scripts) instead of going through the Win32-app machinery above. No `.intunewin` packaging, no detection/requirement rule, no `Deploy-*.ps1` — just upload the one file.

**Why this exists alongside the Win32 app:** the Win32-app route hit real friction in practice (the `IntuneWin32App`-module `@odata.type` bug, GRS lockouts affecting the whole device, a `RestartBehavior` typo) — a plain platform script sidesteps all of that Win32-app-specific machinery entirely. The tradeoff is losing the two things Win32 apps and Remediations each offered:

| | Win32 app (this folder's `Deploy-CoworkPrerequisitesIntune.ps1`) | Platform script (`CoworkPrerequisites-PlatformScript.ps1`) |
|---|---|---|
| Restart after enabling VMP | `-RestartBehavior 'basedOnReturnCode'` + exit `3010` → **Intune itself** enforces a restart (deadline/grace period), no user action required beyond the prompt | No such mechanism exists for platform scripts — exit code is only success(`0`)/failure(anything else), so this script always exits `0` and only sends the one-time `msg.exe` notification (via the same scheduled-task-in-the-user's-session trick as the Win32 app, so the OK button actually works); the user must restart entirely on their own initiative |
| Re-checking after failure | Detection rule re-evaluated on every check-in | Runs once per device by default; a "failed" run does get retried on subsequent check-ins, but a *successful* run (VMP enabled, restart still pending) is not re-verified afterward the way a Proactive Remediation's daily Detect would |

Proactive Remediations (Detect + Remediate, re-runs daily) would give the best of both — but that feature requires Windows Enterprise/Education licensing or per-user VDA, which **Business Premium does not include**, so it isn't an option here.

**Deploy:**
1. **Devices → Scripts and remediations → Platform scripts → Add → Windows 10 and later**
2. Upload `CoworkPrerequisites-PlatformScript.ps1`
3. Script settings: **Run this script using the logged on credentials** = No (SYSTEM) · **Enforce script signature check** = No · **Run script in 64-bit PowerShell Host** = Yes
4. Assign to the same device group as Claude Desktop

Logs to `%ProgramData%\CoworkPrereqDeploy\platformscript.log` — a different filename from the Win32-app variant's `install.log`, so testing both on the same device doesn't overwrite either log.
