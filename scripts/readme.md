# scripts/

All PowerShell tooling for this repo, grouped by workload. Launch everything from the root via [`.\menu.ps1`](../menu.ps1) (or [`.\load.ps1`](../load.ps1) on first run) — see the [root readme](../readme.md) for the full menu reference and getting-started steps.

---

## Categories

| Folder | Description |
|--------|-------------|
| [`Entra/`](Entra/readme.md) | User lifecycle, manager assignment, license reporting, Conditional Access baseline (Microsoft Graph) |
| [`Exchange/`](Exchange/readme.md) | Calendar migration/permissions, dynamic → static distribution groups |
| [`Graph/`](Graph/readme.md) | Microsoft Graph application permission management |
| [`Intune/`](Intune/readme.md) | Autopilot enrollment, iOS compliance policy updater, corporate wallpaper/lockscreen deployment |
| [`Reporting/`](Reporting/readme.md) | Computer last-logon report, SharePoint storage report, monthly licensing report |
| [`Device/`](Device/readme.md) | Windows endpoint maintenance — activation, cleanup, temp files, time sync, audio |
| [`Deployment/`](Deployment/readme.md) | USB toolkit for Windows setup and Autopilot enrollment during OOBE |
| [`DNS/`](DNS/readme.md) | Resolve and import DNS records into AD-integrated DNS zones |
| [`SAS/`](SAS/readme.md) | SAS batch job error monitoring with Zabbix integration |
| [`Startup/`](Startup/readme.md) | `functies.ps1` M365 function library + module bootstrap, dot-sourced by the menu |
| [`Testing Scripts/`](Testing%20Scripts/readme.md) | Audit and diagnostic scripts by workload (Exchange, Entra, Network, RDS, SMTP, Device, SharePoint) |
| [`Custom Scripts/`](Custom%20Scripts/readme.md) | Path-pinned scripts — Office theme deployment (hardcodes its download URL to this repo path) |
