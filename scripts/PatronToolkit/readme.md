# Patron Toolkit

Modern Microsoft Graph / Exchange Online rewrites of the still-useful capabilities from
the retired [`directorcia/patron`](https://github.com/directorcia/patron) toolkit — a
large, ~183-script Microsoft 365 tenant management/reporting collection built partly on
the now-retired `MSOnline` and `AzureAD` PowerShell modules.

None of the code below is copied from that project — it inspired the *capability list*
(what to check, what a good report looks like), but every script here is a fresh
implementation in this repo's house style: dry-run-by-default for anything mutating,
CSV export, GDAP/`-TenantId`-aware connection reuse, and no hardcoded credentials or
tenant data. Because the source project shipped one tiny script per check with heavy
overlap, this toolkit consolidates ~70 of those single-purpose scripts into 13
well-parameterized ones, grouped by capability rather than ported 1:1. See each script's
`.NOTES` section for exactly which upstream scripts it replaces.

---

## Folders

| Folder | Description |
|--------|-------------|
| [`Entra/`](Entra/readme.md) | MFA/SSPR registration reporting, Conditional Access policy backup |
| [`Security/`](Security/readme.md) | App consent audit, suspicious inbox rules, security alerts, email security posture, audit logging, SPF/DMARC validation |
| [`Intune/`](Intune/readme.md) | Intune policy assignment reporting, Autopilot device inventory |
| [`Exchange/`](Exchange/readme.md) | Message trace / mail flow diagnostics |
| [`SharePoint/`](SharePoint/readme.md) | SharePoint Online sharing configuration and external user audit |
| [`Teams/`](Teams/readme.md) | Teams tenant governance and inventory reporting |

---

## What was skipped, and why

The source project's ~183 scripts collapsed into these 13 for several reasons:

- **Native platform coverage** — the large family of `endpoint-*-set.ps1` /
  `intune-*comp-set.ps1` / `intune-*ap-set.ps1` scripts (Intune security baselines,
  compliance policies, app protection policies for Windows/iOS/Android/macOS) hardcode
  one MSP's specific opinionated settings. Microsoft's own Intune Security Baseline and
  compliance policy templates in the admin center now cover this natively and are kept
  current by Microsoft — scripting a fixed 2020-era baseline provides no advantage.
- **Already covered in this repo** — basic license/user/group reporting
  (`o365-sku-audit-csv.ps1`, `graph-sku-get.ps1`, `o365-NoSPO-ADAcct*`), CA baseline import
  (`ca-policy-import.ps1` — see `scripts/Entra/Import-ConditionalAccessBaseline.ps1`), and
  Intune config drift detection (`endpoint-policy-get.ps1` — see
  `scripts/Intune/Compare-IntuneConfig.ps1`) already exist and are actively maintained.
- **Retired/legacy modules with no safe modern equivalent for the specific narrow
  capability** — `o365-add-domain.ps1` (MSOnline + hardcoded Azure DNS zone setup),
  `graph-usrreg-read.ps1`'s original local-XML credential flow, `mcas-*.ps1` (Defender for
  Cloud Apps' separate portal-token auth model, increasingly folded into the unified
  Defender portal) — the underlying *capability* of the last two was kept but rewritten
  (MFA report; app consent audit), the rest dropped.
- **Not tenant-portable / narrow value for an MSP managing many customer tenants** —
  on-prem AD reconciliation reports (`o365-NoSPO-ADAcct*`, `o365-SPO-NoADAcct*`), WHOIS
  lookup (`o365-whois-get.ps1`), generic DNS record dump (`o365-dns-get.ps1` — narrowed
  down to the actually-actionable SPF/DMARC check), and local Hyper-V/menu/credential-cache
  plumbing (`hyperv-*.ps1`, `start.ps1`, `*-connect.ps1`, `*-creds-save.ps1`).
- **Out of theme** — `reclaimwin10.ps1` and `win10-bp-get.ps1` are large local-machine
  registry/GPO scripts (one is a vendored third-party script, not even patron's own code)
  operating on a single Windows 10 workstation, not tenant-wide M365 management.
- **High-risk bulk mutations kept deliberately out of scope** — Conditional Access policy
  bulk import/delete, Intune policy/app assignment bulk add/remove, and Autopilot device
  bulk import/delete/reassign were all *reporting* capabilities we kept (see
  `Export-ConditionalAccessPolicies.ps1`, `Get-IntunePolicyAssignments.ps1`,
  `Get-AutopilotDevices.ps1`) but we did not port their mutating counterparts — those are
  high-blast-radius, tenant-specific operations best reviewed one at a time in the admin
  center rather than scripted generically.

Full detail on exactly which source script(s) each report replaces is in that script's
`.NOTES` section.
