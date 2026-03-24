# Intune — iOS Compliance Updater

Automatically keeps the minimum iOS version requirement in an Intune compliance policy up to date via Microsoft Graph API. Runs as a scheduled task on a Windows server — no manual work required.

**How it works**

1. Fetches the latest iOS version from Apple's RSS feed (with fallback to Apple Support page)
2. Compares it to the current minimum version in the Intune compliance policy
3. Updates the policy if a newer version is available
4. Logs all actions to a monthly log file

---

## Files

| File | Description |
|------|-------------|
| `Update-iOSCompliancePolicy.ps1` | Main script — run manually or via scheduled task |
| `Setup.ps1` | One-time setup — creates App Registration and writes config.json |
| `Install-ScheduledTask.ps1` | Registers the Windows scheduled task |
| `config.example.json` | Example configuration file |

---

## Setup

### Option A — Automatic (recommended)

Run `Setup.ps1` once. It handles everything:

```powershell
.\Setup.ps1
```

The script will:
- Install required PowerShell modules
- Open a browser for authentication
- Create the App Registration in Entra ID
- Assign the required API permission and grant admin consent
- Create a Client Secret
- Show available iOS compliance policies to choose from
- Write `config.json` automatically

```powershell
# Specify the policy name directly to skip the selection prompt
.\Setup.ps1 -CompliancePolicyName "iOS - Minimum version compliance"
```

> Required role: **Global Administrator** or **Application Administrator + Intune Administrator**

---

### Option B — Manual

**Step 1 — Create App Registration**

1. Entra ID → **App registrations** → **New registration**
2. Name: `Intune iOS Compliance Updater`
3. After creation → **API permissions** → **Add a permission**
4. Choose **Microsoft Graph** → **Application permissions**
5. Add: `DeviceManagementConfiguration.ReadWrite.All`
6. Click **Grant admin consent**
7. **Certificates & secrets** → **New client secret** — note the value

**Step 2 — Find the compliance policy ID**

1. Intune Admin Center → **Devices** → **Compliance** → open the iOS policy
2. Copy the GUID from the URL: `.../deviceCompliancePolicies/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

**Step 3 — Create config.json**

Copy `config.example.json` to `config.json` and fill in the values:

```json
{
    "TenantId":           "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientId":           "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientSecret":       "your-client-secret",
    "CompliancePolicyId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
}
```

> Never commit `config.json` to Git — it contains secrets. It is listed in `.gitignore`.

---

### Register the scheduled task

After setup (both options), register the task:

```powershell
# Run as Administrator
.\Install-ScheduledTask.ps1
```

The task runs every **Monday at 07:00** as SYSTEM.

---

## Usage

```powershell
# Run manually
.\Update-iOSCompliancePolicy.ps1

# Dry run — no changes made
.\Update-iOSCompliancePolicy.ps1 -WhatIf

# Custom config or log path
.\Update-iOSCompliancePolicy.ps1 -ConfigPath "D:\configs\intune.json" -LogPath "D:\logs"
```

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

- Never commit `config.json` to Git (already in `.gitignore`)
- For production environments, consider storing the Client Secret in **Windows Credential Manager** or **Azure Key Vault**
- The App Registration uses only the minimum required permission: `DeviceManagementConfiguration.ReadWrite.All`
