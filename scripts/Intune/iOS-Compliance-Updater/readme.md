**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [Intune](../readme.md) › **iOS-Compliance-Updater**

# Intune — iOS Compliance Updater

Automatically keeps the minimum iOS version requirement in an Intune compliance policy up to date via Microsoft Graph API. Runs as a scheduled task on a Windows server — no manual work required.

**How it works**

1. Fetches the latest released iOS version from Apple's RSS feed (with fallback to Apple Support page) — beta and RC entries are skipped, and the highest version wins (the feed also lists updates for older major versions)
2. Compares it to the current minimum version in the Intune compliance policy
3. Raises the policy if a newer version is available (it never lowers it)
4. Logs all actions to a monthly log file

**Sign-in** goes through [`Connect-M365.ps1`](../../Startup/Connect-M365.ps1), so keep this folder inside the repository layout. Because the updater runs unattended, **app-only is its default**: a certificate (thumbprint in `config.json`, private key in the Windows certificate store). A `ClientSecret` in an older `config.json` still works. `Setup.ps1` signs you in **delegated** to create the app. All scripts need PowerShell 7.

---

## Files

| File | Description |
|------|-------------|
| [`Update-iOSCompliancePolicy.ps1`](Update-iOSCompliancePolicy.ps1) ([docs](#usage)) | Main script — run manually or via scheduled task |
| [`Setup.ps1`](Setup.ps1) ([docs](#option-a--automatic-recommended)) | One-time setup — creates App Registration and certificate, writes config.json |
| [`Install-ScheduledTask.ps1`](Install-ScheduledTask.ps1) ([docs](#register-the-scheduled-task)) | Registers the Windows scheduled task (pwsh.exe) |
| `config.example.json` | Example configuration file |

---

## Setup

### Option A — Automatic (recommended)

Run `Setup.ps1` once, **elevated** (so the certificate lands in `LocalMachine\My`, where the SYSTEM task can use it). It handles everything:

```powershell
.\Setup.ps1
```

The script will:
- Install required PowerShell modules
- Sign you in (delegated, browser or device code per `load.config.ps1`)
- Create (or reuse) the App Registration in Entra ID
- Assign the required API permission and grant admin consent
- Create a self-signed certificate and upload its public key to the app (existing keys on a reused app are kept)
- Show available iOS compliance policies to choose from (all pages)
- Write `config.json` automatically

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-TenantId` | No | Tenant to set up (default: GDAP customer, else the tenant you sign in to) |
| `-CompliancePolicyName` | No | Policy to use; without it you pick from a list |
| `-AppName` | No | App Registration name (default: `Intune iOS Compliance Updater`) |
| `-CredentialType` | No | `Certificate` (default) or `Secret` |
| `-CertificateStoreLocation` | No | `LocalMachine` (default when elevated) or `CurrentUser` |
| `-CertificateExpiryYears` | No | Certificate validity (default: `2`) |
| `-SecretExpiryYears` | No | Secret validity with `-CredentialType Secret` (default: `2`) |
| `-ConfigPath` | No | Where to write `config.json` (default: script folder) |

```powershell
# Specify the tenant and policy name directly to skip the selection prompt
.\Setup.ps1 -TenantId "contoso.onmicrosoft.com" -CompliancePolicyName "iOS - Minimum version compliance"

# Old behaviour: a client secret in config.json
.\Setup.ps1 -CredentialType Secret
```

> Required role: **Global Administrator**, or **Application Administrator + Privileged Role Administrator** (admin consent) **+ Intune Administrator**

---

### Option B — Manual

**Step 1 — Create App Registration**

1. Entra ID → **App registrations** → **New registration**
2. Name: `Intune iOS Compliance Updater`
3. After creation → **API permissions** → **Add a permission**
4. Choose **Microsoft Graph** → **Application permissions**
5. Add: `DeviceManagementConfiguration.ReadWrite.All`
6. Click **Grant admin consent**
7. **Certificates & secrets** → **Certificates** → upload the `.cer` of a certificate whose private key is in `LocalMachine\My` on the server (or, less safe, create a client secret)

**Step 2 — Find the compliance policy ID**

1. Intune Admin Center → **Devices** → **Compliance** → open the iOS policy
2. Copy the GUID from the URL: `.../deviceCompliancePolicies/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

**Step 3 — Create config.json**

Copy `config.example.json` to `config.json` and fill in the values:

```json
{
    "TenantId":              "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientId":              "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "CertificateThumbprint": "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "CompliancePolicyId":    "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
}
```

With a client secret, use `"ClientSecret": "..."` instead of `CertificateThumbprint`.

> Never commit `config.json` to Git. It is listed in `.gitignore`.

---

### Register the scheduled task

After setup (both options), register the task:

```powershell
# Run as Administrator
.\Install-ScheduledTask.ps1
```

The task runs every **Monday at 07:00** as SYSTEM, with `pwsh.exe` (PowerShell 7 must be installed; the task used `powershell.exe` before, which cannot run the PowerShell 7 updater). Re-run it after updating to this version.

---

## Usage

```powershell
# Run manually
.\Update-iOSCompliancePolicy.ps1

# Dry run — no changes made
.\Update-iOSCompliancePolicy.ps1 -WhatIf

# Custom config or log path
.\Update-iOSCompliancePolicy.ps1 -ConfigPath "D:\configs\intune.json" -LogPath "D:\logs"

# Manual delegated run against one policy (no ClientId in config.json)
.\Update-iOSCompliancePolicy.ps1 -CompliancePolicyId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -WhatIf
```

**Parameters** (`Update-iOSCompliancePolicy.ps1`)

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-ConfigPath` | No | Configuration file (default: `config.json` next to the script) |
| `-LogPath` | No | Log folder (default: `logs\`) |
| `-CompliancePolicyId` | No | Overrides `CompliancePolicyId` from `config.json` |
| `-TenantId` / `-ClientId` / `-CertificateThumbprint` | No | Override the values from `config.json` |
| `-AppOnly` | No | App-only with ClientId and thumbprint from `graph.appid.json` |
| `-WhatIf` | No | Dry run |

**Notes**
- `-WhatIf` was declared both as a script parameter and through `SupportsShouldProcess`, so PowerShell refused to start the script at all ("A parameter with the name 'WhatIf' was defined multiple times"). It is now the common parameter only.
- The PATCH now sends `@odata.type` (`#microsoft.graph.iosCompliancePolicy`), and the script refuses a policy ID that is not an iOS compliance policy.
- The version lookup never worked: `Invoke-RestMethod` returns the RSS items themselves, so `$rss.channel.item.title` was always empty, and the fallback page (`111900`) no longer contained a version. The script now reads the items directly (verified against the live feed: 27.0.1, while the feed also lists 26.6.2 and 18.7.10) and falls back to "About iOS updates" (`support.apple.com/100100`, "The latest version of iOS and iPadOS is …").

---

## Logging

Logs are written to `logs\compliance-updater-<yyyy-MM>.log` next to the script:

```
[2026-03-24 07:00:01] [INFO   ] Script started
[2026-03-24 07:00:03] [SUCCESS] Latest iOS version: 18.3.2
[2026-03-24 07:00:04] [INFO   ] Policy 'iOS - Minimum version' — current minimum: 18.3.1
[2026-03-24 07:00:05] [SUCCESS] Compliance policy updated to iOS 18.3.2.
```

---

## Security

- Prefer the certificate: `config.json` then holds no secret, and the private key is not exportable from the server's certificate store.
- A client secret in `config.json` is plaintext: anyone who can read the file can change Intune configuration. Restrict the file's ACL to SYSTEM and administrators, or switch to a certificate by re-running `Setup.ps1`.
- Never commit `config.json` to Git (already in `.gitignore`)
- The App Registration uses only the minimum required permission: `DeviceManagementConfiguration.ReadWrite.All`
