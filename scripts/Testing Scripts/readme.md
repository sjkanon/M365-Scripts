# Testing Scripts

Audit and diagnostic scripts, organized by workload. Most are self-connecting (reuse an existing session or connect automatically) and export CSV/txt reports to `C:\Temp\` (Windows) or `~/Downloads/` (macOS/Linux).

---

## Categories

| Folder | Description |
|--------|-------------|
| [`Exchange/`](Exchange/readme.md) | Calendar, mailbox, DKIM, distribution group, and external-forwarding audits |
| [`Entra/`](Entra/readme.md) | M365 Group / Teams membership audit |
| [`Network/`](Network/readme.md) | TCP port checks, auth/network diagnostics, file I/O stress testing |
| [`RDS/`](RDS/readme.md) | RDP / RD Web Access login diagnostics and live session monitoring |
| [`SMTP/`](SMTP/readme.md) | SMTP relay connectivity tests (one-time and recurring) |
| [`Device/`](Device/readme.md) | OpenVPN diagnostics, Teams archiver |
| [`SharePoint/`](SharePoint/readme.md) | Pointer to the SharePoint storage report (lives in `Reporting/`) |
