# Azure Infrastructure Scripts

Scripts for managing Azure IaaS resources directly (not the M365 tenant) — separate from every other category in this repo, which targets Microsoft 365 / Entra ID via Graph or Exchange Online. Requires the `Az` PowerShell module and an authenticated `Connect-AzAccount` session. Not wired into [`menu.ps1`](../../menu.ps1) — run directly against the target subscription.

---

## Folders

| Folder | Description |
|--------|-------------|
| [`VM/`](VM/readme.md) | Convert Azure VM disk controller type between SCSI and NVMe |
