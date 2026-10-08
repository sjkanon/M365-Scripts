**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [LegacyUtilities](../readme.md) › **Workspace365**

# Legacy Utilities — Workspace 365

Provisioning scripts for [Workspace 365](https://workspace365.net) digital
workplace environments. Modernized rewrite of an old 600+ line interactive
script: MSAL.PS device-code login and raw bearer-token Graph calls replaced
with `Connect-MgGraph` / `Invoke-MgGraphRequest` (reusing an existing session
if already connected), and the always-interactive prompts turned into
parameters with this repo's dry-run-by-default / `-Apply` pattern. The
provisioning key and hostname are always supplied by the caller — the
original saved them to a local plaintext `.cfg` file next to the script; that
persistence was intentionally dropped.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`New-Workspace365Environment.ps1`](New-Workspace365Environment.ps1) ([docs](#new-workspace365environmentps1)) | Provision a new environment, its SSO app registration, and default Exchange/SharePoint links |
| [`Remove-Workspace365Environment.ps1`](Remove-Workspace365Environment.ps1) ([docs](#remove-workspace365environmentps1)) | Delete an environment via the Provisioning API |

---

### New-Workspace365Environment.ps1

Creates an Entra ID App Registration for Workspace 365 SSO (delegated
Graph/Power BI scopes + a client secret), provisions the environment via the
Workspace 365 Provisioning API, links the app as its SSO identity provider, and
points its default Exchange/SharePoint URLs at this tenant. Dry-run by default.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-WorkspaceHostname` | Yes | e.g. `https://yourcompany.workspace365.net` |
| `-ProvisioningKey` | Yes | Workspace 365 provisioning key (GUID) — never hardcode, pass at call time |
| `-EnvironmentName` | Yes | Lowercase alphanumeric environment name |
| `-RequestingUserUpn` | No | Defaults to the signed-in user |
| `-Apply` | No | Actually provision (default: preview) |
| `-TenantId` | No | Entra ID tenant ID or domain (default: the GDAP customer when `authMode` is GDAP) |
| `-ClientId` / `-CertificateThumbprint` | No | App-only sign-in with this app registration and certificate |
| `-AppOnly` | No | App-only sign-in with the app from `graph.appid.json` |

```powershell
.\New-Workspace365Environment.ps1 -WorkspaceHostname "https://yourcompany.workspace365.net" -ProvisioningKey $key -EnvironmentName "contoso" -Apply
```

**Notes**
- Signs in to Graph through [`Connect-M365.ps1`](../../Startup/readme.md): delegated by default, app-only with `-ClientId` + `-CertificateThumbprint` or `-AppOnly` (then `-RequestingUserUpn` is required — there is no signed-in user). A fitting session is reused; only a session the script opened is disconnected
- Delegated scopes: `Application.ReadWrite.All`, `User.Read`, `User.ReadBasic.All` (added: `-RequestingUserUpn` may name another admin), `Organization.Read.All`
- The Provisioning API itself is called with the provisioning key, not with Graph

---

### Remove-Workspace365Environment.ps1

Deletes an environment via the Provisioning API. Does **not** remove the
associated App Registration — clean that up separately (Entra admin center or
`Remove-MgApplication`) if it's no longer needed. Dry-run by default.

```powershell
.\Remove-Workspace365Environment.ps1 -WorkspaceHostname "https://yourcompany.workspace365.net" -ProvisioningKey $key -EnvironmentName "contoso" -Apply
```

---

**Required modules**

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

**Notes**
- By using these scripts you accept the
  [Workspace 365 terms and conditions](https://workspace365.net/en/term-and-conditions).
