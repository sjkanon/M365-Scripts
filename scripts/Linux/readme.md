**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../readme.md) › [scripts](../readme.md) › **Linux**

# Linux Scripts

Scripts for Linux servers: Debian/Ubuntu, including servers that run 3CX Phone System. These are **bash** scripts, not PowerShell: a Linux server like a 3CX appliance usually has no `pwsh`. They run on the server itself as root, and are not part of [`menu.ps1`](../../menu.ps1).

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Invoke-LinuxCleanup.sh`](Invoke-LinuxCleanup.sh) ([docs](#invoke-linuxcleanupsh)) | Scan and free disk space on a Debian/Ubuntu server: packages, journal, logs, temp, user caches, optionally Docker, plus 3CX logs and backups when 3CX is installed |

---

### Invoke-LinuxCleanup.sh

The Linux counterpart of [`Invoke-WindowsCleanup.ps1`](../Device/readme.md#invoke-windowscleanupps1). Without arguments it is a dry run: per category it shows how much can be freed. `--apply` performs the cleanup, and `--check-only` is the same scan for monitoring, with an exit code. Everything removed by age must be older than `--days` (default 14). Nothing is upgraded: on a 3CX server, 3CX updates itself and the OS through its own console.

**What it cleans**

| Category | Details |
|----------|---------|
| Packages | APT cache (`apt-get clean`); packages nobody needs any more, old kernels included (`apt-get autoremove --purge`), the list shown in the dry run; leftover configuration of removed packages (dpkg state `rc`) |
| Snap | Disabled snap revisions, when snap is installed |
| Journal | systemd journal, vacuumed to `--days` and `--journal-size`. The dry run works out what that vacuum frees, the way journald decides it: only archived files, the oldest first, by the time in the file name |
| System logs | Rotated logs in `/var/log` (`*.1`, `*.gz`, `*.xz`, `*.old`, ...), PostgreSQL and nginx included — active logs stay |
| Crash dumps | `/var/lib/systemd/coredump`, `/var/crash` |
| Temp | `/tmp`, `/var/tmp` |
| User cache | `~/.cache`, `~/.npm/_cacache` and `~/.local/share/Trash` of root and every home under `/home` |
| Docker | Only with `--docker`: `docker system prune -f` (stopped containers, unused networks, dangling images, build cache). Without it, Docker's own reclaimable figure is shown |
| 3CX logs | `<data-dir>/Logs`, 3CX's nginx logs (`/var/lib/3cxpbx/Bin/nginx/logs`) and `/var/lib/3cxpbx/Data/Logs`, left over on servers upgraded from the layout before `Instance1` — only when 3CX is installed |
| 3CX backups | Every `*.zip` in `<data-dir>/Backups` and the old `/var/lib/3cxpbx/Data/Backups` is listed with date and size, newest first, as one list. With `--keep-backups N` everything beyond the newest N is deleted |

Reported, never deleted: call recordings (with their size), 3CX backups without `--keep-backups`, the largest folders in the 3CX data folder, and the 10 largest files over 500 MB on the server.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `--apply` | Perform the cleanup (default: dry run only) |
| `--check-only` | Scan only, change nothing; exit code `2` when something can be freed, `0` when not. Cannot be combined with `--apply` |
| `--days N` | Only remove files older than N days (default: `14`). The journal, `/tmp` and user caches keep at least one day |
| `--journal-size SIZE` | Shrink the systemd journal to at most SIZE, e.g. `200M` or `1G` (default: `200M`) |
| `--docker` | Also run `docker system prune -f`. Volumes are never pruned |
| `--keep-backups N` | Delete 3CX backups beyond the newest N (N ≥ 1) |
| `--data-dir PATH` | 3CX data folder (default: `/var/lib/3cxpbx/Instance1/Data`) |
| `--skip-apt` | Skip everything APT/dpkg: cache, autoremove, leftover configuration |
| `--skip-autoremove` | Skip only `apt-get autoremove` |
| `--skip-journal` | Skip the systemd journal |
| `--skip-user-cache` | Skip the per-user caches and trash |
| `--skip-3cx` | Skip the 3CX sections |
| `-h`, `--help` | Show the help |

**Examples**

```bash
# Dry run — what can be freed
sudo ./Invoke-LinuxCleanup.sh

# Monitoring / RMM: exit code 2 when there is work, 0 when clean
sudo ./Invoke-LinuxCleanup.sh --check-only

# Clean up
sudo ./Invoke-LinuxCleanup.sh --apply

# Clean up harder: files older than 7 days, keep the 5 newest 3CX backups
sudo ./Invoke-LinuxCleanup.sh --apply --days 7 --keep-backups 5

# Also prune Docker
sudo ./Invoke-LinuxCleanup.sh --apply --docker

# Straight from GitHub, without copying the file (public repository only)
curl -fsSL https://raw.githubusercontent.com/sjkanon/M365-Scripts/main/scripts/Linux/Invoke-LinuxCleanup.sh | sudo bash -s -- --check-only
```

**Exit codes**

| Code | Meaning |
|------|---------|
| `0` | Done, or nothing to do |
| `1` | Wrong argument, or not run as root |
| `2` | `--check-only` only: something can be freed |

**Notes**

- Requires root (`sudo`). Works on Debian and Ubuntu; on a distribution without `apt-get` the package section is skipped and the rest still runs.
- 3CX is detected by the `3cxpbx` package or the data folder; without 3CX the 3CX sections are left out. `--skip-3cx` leaves them out on a 3CX server.
- **Safety:** when `apt`/`dpkg` is already running (an OS or 3CX update), the package section is skipped for that run. That is read from dpkg's own lock (`/proc/locks`), not from process names: `unattended-upgrades` keeps a process called `unattended-upgr` running all the time, which made the first version skip the packages on every run. `autoremove` is refused when its list contains a 3CX package. In `/tmp`, `systemd-private-*` (owned by running services) and PostgreSQL's socket lock file `.s.PGSQL.*` are left alone.
- Call recordings are never deleted: 3CX has its own retention setting for them in the management console. The same goes for the database, voicemail, prompts and configuration.
- "Total freed (measured)" compares the used space of `/`, `/var`, `/tmp`, `/home` and the 3CX data folder before and after; "reported" adds up the categories. Docker is not in the dry-run estimate, because what a prune frees cannot be known beforehand.
- A log file in the 3CX log folders that a service still holds open but has not written to for `--days` days is removed too. Its space is only freed once that service restarts.
- The 3CX paths (`/var/lib/3cxpbx/Instance1/Data`, `Logs`, `Backups`, `Recordings`, `Bin/nginx/logs`) follow 3CX's Linux layout; use `--data-dir` when an installation differs.
- Run as a dry run on a live 3CX server (Debian 12, 3CX from `repo.3cx.com`); `--apply` has not run on one yet. Tested on Debian 12 (bookworm) with a recreated 3CX folder layout and a running `systemd-journald`: dry run, `--check-only`, `--apply` and a second `--apply`.
