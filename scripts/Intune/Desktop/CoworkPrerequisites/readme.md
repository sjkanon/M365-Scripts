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
| `Deploy-CoworkPrerequisitesIntune.ps1` | **The one you run.** Packages the scripts below and creates/updates the Win32 app. |
| `Install-CoworkPrerequisites-Intune.ps1` | Install command content script |
| `Uninstall-CoworkPrerequisites-Intune.ps1` | Uninstall command content script |
| `Detect-CoworkPrerequisites-Intune.ps1` | Custom detection script |

There's no MSIX here — unlike Claude Desktop, the "content" is just these three scripts, so a re-run only does something if you've actually edited one of them (tracked via a `ScriptsHash` in the app's Notes field, same pattern as `Deploy-ClaudeDesktopIntune.ps1`).

## What the install script does

1. **VirtualMachinePlatform**: `Enable-WindowsOptionalFeature`, with retry against transient DISM failures. No-op if already enabled.
2. **Fast Startup**: sets `HiberbootEnabled = 0` under `HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power`, on every run (not just when VMP was just enabled). Anthropic's own Cowork documentation explicitly warns: *"Restart the machine using Restart, not shut down and power on. With Windows Fast Startup enabled, a shutdown cycle can leave the virtualization services uninitialized."* Fast Startup is on by default on nearly every Windows image — without disabling it, a user who just "shuts down" instead of "restarts" (the common case) never gets working Cowork services, no matter how many power cycles happen, even though Intune shows the app as installed.
3. If VMP was just enabled: exits with code **3010** ("soft reboot required"). The app is configured with `-RestartBehavior 'basedOnReturnCode'`, so **Intune itself** enforces the restart (prompt/deadline/grace period) — this doesn't depend on a user being logged in. As a courtesy it also sends an English `msg.exe` notification to the active console session (recognized via the untranslated `SESSIONNAME` "console", not the OS-language-dependent `STATE` text "Active" — a plain string match on "Active" would silently never fire on a non-English Windows display language), but the notification is a nicety, not the mechanism the restart actually relies on.

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
