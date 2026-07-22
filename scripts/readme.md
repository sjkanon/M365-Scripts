# scripts/

All PowerShell tooling for this repo, grouped by workload. Launch everything from the root via [`.\menu.ps1`](../menu.ps1) (or [`.\load.ps1`](../load.ps1) on first run) — see the [root readme](../readme.md) for the full menu reference and getting-started steps.

---

## Categories

| Folder | Description |
|--------|-------------|
| [`Entra/`](Entra/readme.md) | User lifecycle, manager assignment, license reporting, Conditional Access baseline, temporary CA windows, TAP codes, M365 Group audit (Microsoft Graph) |
| [`Exchange/`](Exchange/readme.md) | Calendar migration/permissions, distribution groups, mailbox/calendar/DKIM/forwarding audits |
| [`Graph/`](Graph/readme.md) | Microsoft Graph application permission management |
| [`Intune/`](Intune/readme.md) | Autopilot enrollment, iOS compliance policy updater, corporate wallpaper/lockscreen deployment |
| [`Reporting/`](Reporting/readme.md) | Computer last-logon report, SharePoint storage report, monthly licensing report |
| [`Device/`](Device/readme.md) | Windows endpoint maintenance — activation, cleanup, temp files, time sync, audio, OpenVPN diagnostics |
| [`Network/`](Network/readme.md) | TCP port checks, auth/network diagnostics, file I/O stress testing |
| [`RDS/`](RDS/readme.md) | RDP / RD Web Access login diagnostics and live session monitoring |
| [`SMTP/`](SMTP/readme.md) | SMTP relay connectivity tests (one-time and recurring) |
| [`Deployment/`](Deployment/readme.md) | USB toolkit for Windows setup and Autopilot enrollment during OOBE |
| [`DNS/`](DNS/readme.md) | Resolve and import DNS records into AD-integrated DNS zones |
| [`SAS/`](SAS/readme.md) | SAS batch job error monitoring with Zabbix integration |
| [`Teams/`](Teams/readme.md) | Microsoft Teams / SharePoint export and archiving |
| [`Startup/`](Startup/readme.md) | `functies.ps1` M365 function library + module bootstrap + syntax checker, dot-sourced by the menu |
| [`Custom Scripts/`](Custom%20Scripts/readme.md) | Path-pinned scripts — Office theme deployment (hardcodes its download URL to this repo path) |
