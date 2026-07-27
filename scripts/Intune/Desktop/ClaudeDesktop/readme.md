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
6. Later runs: if the downloaded MSIX version is newer than what's recorded on the Intune app (tracked in the app's Notes field, no local state file needed), pushes an updated package via `Update-IntuneWin32AppPackageFile` — the existing assignment is left untouched, devices just get the new content. If the version hasn't changed, nothing happens.

Detection is deliberately version-agnostic (presence of the provisioned package + `VirtualMachinePlatform` enabled). Intune redeploys a Win32 app to already-targeted devices whenever its content version changes in Intune, regardless of what the detection rule reports — so there's no `$MinimumVersion` to bump by hand every month.

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

### Prerequisites

- `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications`, `Microsoft.Graph.Groups`, `IntuneWin32App` PowerShell modules — install with `.\scripts\Startup\Install-Modules.ps1`
- Run from Windows (the packaging tool and MSIX/AppX cmdlets are Windows-only)
