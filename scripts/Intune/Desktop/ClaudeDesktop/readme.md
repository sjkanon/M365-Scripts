# ClaudeDesktop

Machine-wide Intune deployment of [Claude Desktop](https://claude.com/download) for Windows. Run **one script once a month** to keep the Intune app current — no persistent App Registration or client secret to manage.

Based on [Deploy Claude Desktop for Windows](https://support.claude.com/en/articles/12622703-deploy-claude-desktop-for-windows) and [Enterprise configuration for Claude Desktop](https://support.claude.com/en/articles/12622667-enterprise-configuration-for-claude-desktop).

> **Cowork's Windows prerequisites (VirtualMachinePlatform, Fast Startup) are a separate, independent Win32 app** — see [`../CoworkPrerequisites/readme.md`](../CoworkPrerequisites/readme.md). Deploy both if you want Cowork; deploy only this one if you just want Claude Desktop itself. **By default there is no Intune dependency** between the two: a Cowork-prerequisites failure must never block Claude Desktop (which works fine without Cowork), and a Claude Desktop install problem must never be confused with a Windows-feature problem — each app gets its own, separately visible install status in Intune. If your environment treats Cowork as a hard requirement rather than optional, pass `-RequireCoworkPrerequisites` to *this* deploy script to add a real Intune dependency instead — see "Optional: requiring Cowork Prerequisites" below.

---

## Why not just upload the MSIX as a line-of-business app?

Intune installs LOB MSIX apps per-user. That fails for standard users without admin rights. Instead, `Add-AppxProvisionedPackage` is wrapped as a Win32 app, exactly as the official article recommends.

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

Detection is deliberately version-agnostic (presence of the provisioned package). Intune redeploys a Win32 app to already-targeted devices whenever its content version changes in Intune, regardless of what the detection rule reports — so there's no `$MinimumVersion` to bump by hand every month, and a pure script edit (no new MSIX) still triggers a redeploy via the `ScriptsHash` check above.

### Known issue: 0x80070001 install failures — and a much deeper `IntuneWin32App` module bug behind it

Devices failed installation with error `0x80070001` regardless of the device. Comparing the cached Win32 app policy on an affected device (`AppWorkload.log`) against every other Win32 app assigned to the same tenant showed the anomaly: every other app's `RequirementRules` had `RequiredOSArchitecture: 3`, but Claude Desktop's had `RequiredOSArchitecture: 32` — a value no other app in the tenant used, and 16x what `-Architecture 'x64'` should produce. Since this is stored on the app object in Intune (not per-device), it explained why the failure was 100% reproducible across every device, not device-specific corruption.

**First fix attempt (incomplete):** resubmit `-RequirementRule` on every update, not just at first creation. This turned out to not actually work, because of a second, much more serious bug:

**The real root cause**, confirmed against the `IntuneWin32App` module's own source (v1.5.0) and Microsoft's documented [`win32LobApp` resource](https://learn.microsoft.com/en-us/graph/api/resources/intune-apps-win32lobapp) schema:
- `applicableArchitectures` / `allowedArchitectures` / `minimumSupportedWindowsRelease` are **flat top-level properties** of `win32LobApp` — they need no `@odata.type` at all (that's only required for the polymorphic `rules` collection: file/registry/product-code/script-based detection or requirement rules).
- `New-IntuneWin32AppRequirementRule` never sets `@odata.type` on the object it returns (correctly — it doesn't need one).
- But `Set-IntuneWin32App` (the **update** cmdlet only — `Add-IntuneWin32App`, used at first creation, has entirely different, correct internal logic) incorrectly demands `@odata.type` on `-RequirementRule` anyway, and when it's missing, does `Write-Warning "...missing required '@odata.type'..."` followed by a bare `break`.
- That `break` has no enclosing loop or `switch` — empirically verified, it terminates the **entire rest of the function**, including the actual Graph PATCH call further down. Meaning: every single `Set-IntuneWin32App` call that included `-RequirementRule` silently did **nothing at all** — not the architecture fix, not `-Notes`, not `-DetectionRule`, not `-Icon`, not `-CompanyPortalFeaturedApp`. Only `Update-IntuneWin32AppPackageFile` (a separate call, made just before it) actually reached Intune each run.

**The actual fix**: the update branch no longer passes `-RequirementRule` to `Set-IntuneWin32App` at all (letting Notes/AppVersion/DetectionRule/Icon/CompanyPortalFeaturedApp go through normally), and instead calls `Set-Win32AppArchitectureRequirement` — a small helper that PATCHes `applicableArchitectures`/`allowedArchitectures`/`minimumSupportedWindowsRelease` directly via Microsoft Graph (reusing the same `$Global:AuthenticationHeader` session `Connect-MSIntuneGraph` already established), bypassing the broken cmdlet path entirely for just that piece. `Add-IntuneWin32App` (first creation) is unaffected by any of this and keeps using `-RequirementRule` normally.

**If a device still fails after that fix**: check whether it's actually stuck behind Intune's unrelated **GRS retry cooldown** instead (3 failed attempts → 24h lockout, regardless of the app config) — see the [GRS-cooldown scripts](../../readme.md#repair-stuckwin32appenforcementps1) one level up, either the manual on-device version or the Detect-/Remediate- pair that runs entirely via the Intune portal. This affects every Win32 app on that device, not just Claude, so it's a useful first check if several unrelated apps are also silently stuck.

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
| `-RequireCoworkPrerequisites` | off | Add a real Intune dependency on the Cowork Prerequisites app — see below |
| `-CoworkPrerequisitesAppDisplayName` | `Cowork Windows Prerequisites (Machine-wide)` | Must match `-AppDisplayName` used in `Deploy-CoworkPrerequisitesIntune.ps1`. Only used with `-RequireCoworkPrerequisites` |

### Optional: requiring Cowork Prerequisites

By default, Claude Desktop and Cowork Prerequisites are independent — no Intune dependency links them (see the note at the top of this readme for why). If your environment genuinely needs Cowork to always be present — a device without working Cowork prerequisites shouldn't get Claude Desktop at all — pass `-RequireCoworkPrerequisites`:

```powershell
.\Deploy-ClaudeDesktopIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop" -RequireCoworkPrerequisites
```

This looks up the Cowork Prerequisites app (deploy that one **first** — `Deploy-CoworkPrerequisitesIntune.ps1`) and calls `Add-IntuneWin32AppDependency` with `DependencyType 'Detect'` (not `'AutoInstall'`): the prerequisites app must still be independently assigned Required and already detected on the device — this dependency doesn't auto-install it on Claude Desktop's behalf, it only makes Intune wait for it. The tradeoff versus the default: a device stuck on Cowork prerequisites (e.g. mid-reboot, or genuinely failing) will also not get Claude Desktop until that's resolved, instead of getting Claude Desktop immediately and Cowork later.

### Auto-update policy

The install script sets `HKLM:\SOFTWARE\Policies\Claude\disableAutoUpdates = 1` (DWord). Claude's own updater is disabled so version control stays entirely with this monthly Intune run instead of drifting per-device.

### Clean reinstall on every run

Before provisioning the new version, the install script fully removes Claude Desktop from the device first, in three passes:

1. Stops any running Claude process.
2. Removes every per-user **Appx** installation (`Get-AppxPackage -AllUsers` / `Remove-AppxPackage -AllUsers`, including already-logged-in profiles), then the old machine-wide provisioned package.
3. Removes any **classic (non-Appx) per-user installation** — most notably the consumer installer from [claude.ai/download](https://claude.ai/download), which registers itself through an ordinary per-user Uninstall registry key rather than as an Appx package, so `Get-AppxPackage` never sees it. The script scans the Uninstall registry key of every local profile — including profiles that aren't currently logged in, by temporarily loading their `NTUSER.DAT` — and runs each match's `QuietUninstallString` (or `UninstallString` if that's absent), with a 120s timeout so a stuck installer can't hang the Intune install.

Only after all three passes does it provision the new MSIX. This is deliberately more thorough than just clearing the provisioning layer — any installation left over from a different route can keep running its own, unmanaged Claude session even after the machine-wide version was updated. A user with Claude open loses that session when this runs.

### Robustness on a never-installed device

Both content scripts are written so a device where Claude has never been present — including the very first Autopilot ESP run — goes through cleanly:

- **Install script**: every removal pass (process kill, Appx, classic per-user uninstall) checks for an empty/`$null` result before acting, so a completely clean machine just logs "none found" at each step instead of erroring.
- **Detection script**: an empty/negative result from `Get-AppxProvisionedPackage` (the expected outcome on a never-installed device) reports "not installed" (exit 1) immediately, with no retry — retries only kick in on an actual **exception** (e.g. a transient DISM lock, plausible right after the install script's own heavy Appx activity on the same device), so a genuinely clean machine is never slowed down waiting on retries that can't change the outcome.

### Prerequisites

- `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications`, `Microsoft.Graph.Groups`, `IntuneWin32App` PowerShell modules — install with `.\scripts\Startup\Install-Modules.ps1`
- Run from Windows (the packaging tool and MSIX/AppX cmdlets are Windows-only)
