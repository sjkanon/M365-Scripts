# Office365Toolkit

Modern Microsoft Graph / Exchange Online rewrites of still-useful capabilities from
the retired [`directorcia/Office365`](https://github.com/directorcia/Office365)
(CIAOPS) PowerShell toolkit — a well-known Office 365 admin script collection
originally built on the now-retired `MSOnline` / `AzureAD` modules and various
tenant-specific hardcoded scripts.

The source project (~58 scripts) was reviewed capability-by-capability rather than
ported 1:1: duplicate single-purpose report scripts were consolidated, connection
helper scripts were dropped (this repo already auto-connects and reuses sessions),
anything requiring the retired `MSOnline`/`AzureAD` modules was rewritten against
`Microsoft.Graph.*` or `ExchangeOnlineManagement`, and capabilities already covered
elsewhere in this repo (mailbox size/permission audits, external forwarding,
SharePoint storage, license reporting, Conditional Access, etc. — see
[`scripts/Entra/`](../Entra/readme.md), [`scripts/Exchange/`](../Exchange/readme.md),
[`scripts/Reporting/`](../Reporting/readme.md)) were skipped. All code here is a
fresh implementation in this repo's own style, not copied from the source project.

---

## Folders

| Folder | Description |
|--------|-------------|
| [`Security/`](Security/readme.md) | Secure Score reporting, enterprise app consent cleanup, shared mailbox sign-in lockdown, EOP anti-spam/anti-malware baseline |
| [`Exchange/`](Exchange/readme.md) | Mailbox hygiene baseline, inbox-rule forwarding risk, mailbox add-ins, Unified Audit Log search, message trace |
| [`Intune/`](Intune/readme.md) | Tenant-wide Intune/Endpoint Manager policy inventory |

---

## Skipped capabilities (and why)

**Already covered elsewhere in this repo:**
- Mailbox size/permission/calendar/DKIM audits, external forwarding by mailbox
  setting, SharePoint storage usage, license reporting, Conditional Access baseline
  import, TAP codes — all already have Graph/EXO scripts under `scripts/Entra/`,
  `scripts/Exchange/`, and `scripts/Reporting/`.
- License SKU friendly-name lookup (`o365-skus.ps1`) — a static hashtable of
  legacy SKU names; superseded by `scripts/Entra/Get-M365UserLicenses.ps1`.

**Connection/infrastructure helpers, not capabilities** (`*-connect*.ps1`,
`graph-connect.ps1`, `msgraph-connect.ps1`, `Intune-connect.ps1`, `az-connect*.ps1`,
`o365-setup.ps1`, `o365-update.ps1`, `o365-getrepo.ps1`, `save-cred-file.ps1`,
`c.ps1`, `r.ps1`, `sc-config.ps1`, `text-colour.ps1`): this repo already
auto-connects and reuses existing sessions in every script (see house style in
`scripts/Entra/Test-M365GroupMembership.ps1`), so standalone connector scripts add
nothing. `o365-setup.ps1` also hardcoded a real customer's OneDrive folder path and
installs the retired `MSOnline`/`AzureAD` modules; `save-cred-file.ps1` stores
credentials in a local XML file — both out of scope per this project's safety rules.

**Not M365-tenant scope** (local device/network diagnostics, not Graph/EXO):
`win10-asr-get.ps1`, `win10-audit-get.ps1`, `win10-def-get.ps1` (local Windows
Defender/audit-policy checks — device-level, not tenant-wide), `Cleanup AzureAD
device registration.ps1` (local registry cleanup), `sec-test.ps1` (local EICAR/
malware simulation menu), `ipget.ps1`/`ipinf.ps1` (generic IP geolocation lookups,
one with a hardcoded personal API key — unrelated to M365).

**Vendor/tenant-specific, not generically reusable:** `sc-config.ps1` (hardcoded
to one Australian PSTN carrier's Teams Direct Routing config).

**Deprecated or narrow, low ROI to rebuild:**
- `o365-addin-deploy.ps1` (Centralized Deployment of Outlook add-ins) — Microsoft
  is steering admins toward the Integrated Apps UI in the admin center; no
  Graph/EXO cmdlet equivalent exists yet.
- `o365-mcas-api.ps1` (Cloud App Security API) and `endpoint-api-svbm.ps1`
  (Defender for Endpoint vulnerability API) — both hardcode placeholder
  URI/token/secret variables and target APIs largely superseded by the unified
  Defender portal; not a clean Graph/EXO fit.
- `o365-atp-timer.ps1` (one-off ATP scanning latency measurement) — niche
  diagnostic, not an ongoing management capability.
- `az-sentinel-ruleget.ps1` (Azure Sentinel analytic rule report) — Azure
  Resource Manager scope (`Az.SecurityInsights`), not an M365/Graph/EXO capability.
- SharePoint site collection admins, external user list, and sharing-capability
  settings (`o365-spo-admins.ps1`, `o365-spo-extusr.ps1`, `o365-spo-getsharing.ps1`)
  — not cleanly exposed via Microsoft Graph without SharePoint Online Management
  Shell or PnP.PowerShell, which fall outside this project's Graph/EXO module
  scope; `o365-spo-getusage.ps1` (storage usage) is already covered by
  `scripts/Reporting/Get-SharePointStorageReport.ps1`.
