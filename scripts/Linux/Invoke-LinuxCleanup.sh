#!/usr/bin/env bash
#
# SYNOPSIS
#     Clean up reclaimable disk space on a Debian/Ubuntu server, 3CX Phone System included.
#
# DESCRIPTION
#     Scans and optionally removes, everything age-based older than --days (default 14):
#
#     System
#       - APT package cache (apt-get clean)
#       - Packages nobody needs any more, old kernels included (apt-get autoremove --purge);
#         refused when the list contains a 3CX package
#       - Leftover configuration of removed packages (dpkg state "rc")
#       - Disabled snap revisions, when snap is installed
#       - systemd journal, vacuumed to --days and --journal-size
#       - Rotated logs in /var/log (*.1, *.gz, *.old, ...)
#       - Crash dumps (systemd-coredump, /var/crash)
#       - Files in /tmp and /var/tmp
#     Users (root and every home under /home)
#       - ~/.cache (pip, thumbnails, ...), ~/.npm/_cacache, ~/.local/share/Trash
#     Docker, only with --docker
#       - Stopped containers, unused networks, dangling images, build cache (docker system prune);
#         volumes are never pruned
#     3CX, automatically when 3CX is installed (skip with --skip-3cx)
#       - 3CX's own logs (<data-dir>/Logs, 3CX's nginx logs)
#       - 3CX backups beyond the newest N, only with --keep-backups N
#       - Recordings and backups are reported with their size, recordings are never deleted
#
#     Never touched: databases, call recordings, voicemail, configuration, Docker volumes,
#     active logs. Nothing is upgraded - on a 3CX server, 3CX updates itself and the OS.
#
#     Run without --apply for a dry run, which shows per category how much can be freed.
#     Run with --apply to perform the cleanup. Needs root.
#     --check-only is the same scan for monitoring/RMM: exit code 2 when something can be
#     freed, 0 when there is nothing to do.
#
# PARAMETERS
#     --apply              Perform the cleanup. Without it, only a scan is performed.
#     --check-only         Scan only, change nothing; exit code 2 when something can be freed, 0 when not.
#     --days N             Only remove files older than N days (default: 14).
#     --journal-size SIZE  Shrink the systemd journal to at most SIZE, e.g. 200M or 1G (default: 200M).
#     --docker             Also run docker system prune (containers, networks, dangling images, build cache).
#     --keep-backups N     Delete 3CX backups (*.zip in <data-dir>/Backups) beyond the newest N (N >= 1).
#     --data-dir PATH      3CX data folder (default: /var/lib/3cxpbx/Instance1/Data).
#     --skip-apt           Skip everything APT/dpkg: cache, autoremove, leftover configuration.
#     --skip-autoremove    Skip only apt-get autoremove.
#     --skip-journal       Skip the systemd journal.
#     --skip-user-cache    Skip the per-user caches and trash.
#     --skip-3cx           Skip the 3CX sections.
#     -h, --help           Show this help.
#
# EXAMPLES
#     sudo ./Invoke-LinuxCleanup.sh                       # dry run: what can be freed
#     sudo ./Invoke-LinuxCleanup.sh --check-only          # monitoring: exit 2 when there is work
#     sudo ./Invoke-LinuxCleanup.sh --apply               # clean up
#     sudo ./Invoke-LinuxCleanup.sh --apply --days 7 --keep-backups 5
#     sudo ./Invoke-LinuxCleanup.sh --apply --docker
#     curl -fsSL https://raw.githubusercontent.com/sjkanon/M365-Scripts/main/scripts/Linux/Invoke-LinuxCleanup.sh | sudo bash -s -- --apply

set -uo pipefail

APPLY=0
CHECK_ONLY=0
DAYS=14
JOURNAL_SIZE=200M
DOCKER=0
KEEP_BACKUPS=''
DATA_DIR=/var/lib/3cxpbx/Instance1/Data
SKIP_APT=0
SKIP_AUTOREMOVE=0
SKIP_JOURNAL=0
SKIP_USER_CACHE=0
SKIP_3CX=0
LARGE_MB=500

usage() { sed -n '2,/^$/{s/^# \{0,1\}//;p}' "${BASH_SOURCE[0]}"; }

while (($#)); do
    case "$1" in
        --apply)           APPLY=1 ;;
        --check-only)      CHECK_ONLY=1 ;;
        --days)            DAYS="${2:-}"; shift ;;
        --journal-size)    JOURNAL_SIZE="${2:-}"; shift ;;
        --docker)          DOCKER=1 ;;
        --keep-backups)    KEEP_BACKUPS="${2:-}"; shift ;;
        --data-dir)        DATA_DIR="${2:-}"; shift ;;
        --skip-apt)        SKIP_APT=1 ;;
        --skip-autoremove) SKIP_AUTOREMOVE=1 ;;
        --skip-journal)    SKIP_JOURNAL=1 ;;
        --skip-user-cache) SKIP_USER_CACHE=1 ;;
        --skip-3cx)        SKIP_3CX=1 ;;
        -h|--help)         usage; exit 0 ;;
        *)                 echo "Unknown option: $1 (see --help)" >&2; exit 1 ;;
    esac
    shift
done

((APPLY && CHECK_ONLY)) && { echo '--apply and --check-only cannot be combined' >&2; exit 1; }
[[ $DAYS =~ ^[0-9]+$ ]]                    || { echo '--days needs a whole number' >&2; exit 1; }
[[ $JOURNAL_SIZE =~ ^[0-9]+[KMG]$ ]]        || { echo '--journal-size needs a size like 200M or 1G' >&2; exit 1; }
[[ -z $KEEP_BACKUPS || $KEEP_BACKUPS =~ ^[1-9][0-9]*$ ]] \
                                            || { echo '--keep-backups needs a number of 1 or more' >&2; exit 1; }
[[ -n $DATA_DIR ]]                          || { echo '--data-dir needs a path' >&2; exit 1; }

# ── Output ─────────────────────────────────────────────────────────────────────
if [[ -t 1 ]]; then
    C_CYAN=$'\e[36m'; C_YEL=$'\e[33m'; C_GRN=$'\e[32m'; C_RED=$'\e[31m'; C_GRY=$'\e[90m'; C_OFF=$'\e[0m'
else
    C_CYAN=''; C_YEL=''; C_GRN=''; C_RED=''; C_GRY=''; C_OFF=''
fi

section() { printf '\n  %s%s%s\n' "$C_CYAN" "$1" "$C_OFF"; }
note()    { printf '    %s%s%s\n' "$C_GRY" "$1" "$C_OFF"; }
warn()    { printf '    %s[WARN] %s%s\n' "$C_YEL" "$1" "$C_OFF"; }

fmt() { numfmt --to=iec --suffix=B --format='%.1f' "${1:-0}" 2>/dev/null || printf '%s B' "${1:-0}"; }
sum_sizes() { awk '{ s += $1 } END { printf "%.0f\n", s + 0 }'; }

# Paths among the arguments that exist, so find does not trip over a missing one.
existing() { local p; for p in "$@"; do [[ -e $p ]] && printf '%s\n' "$p"; done; }

# Bytes used on the filesystems this script cleans, each counted once.
disk_used() {
    df -B1 --output=source,used / /var /tmp /home "$DATA_DIR" 2>/dev/null |
        tail -n +2 | sort -u | awk '{ s += $2 } END { printf "%.0f\n", s + 0 }'
}

R_CAT=(); R_DETAIL=(); R_BYTES=(); R_STATUS=()

# add_result <category> <detail> <bytes> <status>   status: Run | Skipped | Info
add_result() {
    local cat=$1 detail=$2 bytes=${3:-0} status=$4 label color size
    case $status in
        Skipped) label=SKIP; color=$C_GRY; size='-' ;;
        Info)    label=INFO; color=$C_GRY; size=$(fmt "$bytes") ;;
        *)       label=$( ((APPLY)) && echo DONE || echo SCAN )
                 color=$( ((bytes > 0)) && echo "$C_GRN" || echo "$C_GRY" )
                 size=$(fmt "$bytes")
                 status=$( ((APPLY)) && echo Cleaned || echo 'Dry run' ) ;;
    esac
    printf '    %s[%s] %-46s %10s%s\n' "$color" "$label" "$detail" "$size" "$C_OFF"
    R_CAT+=("$cat"); R_DETAIL+=("$detail"); R_BYTES+=("$bytes"); R_STATUS+=("$status")
}

# clean_files <category> <detail> <find paths and tests...>
# Sizes what the find expression matches, deletes it under --apply, and reports what went.
clean_files() {
    local cat=$1 detail=$2 before after; shift 2
    before=$(find "$@" -printf '%s\n' 2>/dev/null | sum_sizes)
    if ((APPLY && before > 0)); then
        find "$@" -delete 2>/dev/null
        after=$(find "$@" -printf '%s\n' 2>/dev/null | sum_sizes)
        add_result "$cat" "$detail" $((before - after)) Run
    else
        add_result "$cat" "$detail" "$before" Run
    fi
}

apt_busy() { pgrep -x 'apt|apt-get|aptitude|dpkg|unattended-upgr' >/dev/null 2>&1; }

AGE=(-mmin "+$((DAYS * 1440))")
DAYS_MIN1=$((DAYS > 0 ? DAYS : 1))   # journald, /tmp and user caches never get "everything"
AGE_MIN1=(-mmin "+$((DAYS_MIN1 * 1440))")

# ── Header ─────────────────────────────────────────────────────────────────────
printf '\n  %s================================================%s\n' "$C_CYAN" "$C_OFF"
printf '  %s Invoke-LinuxCleanup%s\n' "$C_CYAN" "$C_OFF"
printf '  %s================================================%s\n\n' "$C_CYAN" "$C_OFF"

if ((EUID != 0)); then
    printf '  %sRun as root: sudo %s%s\n\n' "$C_RED" "$0" "$C_OFF" >&2
    exit 1
fi

for cmd in find du df awk numfmt; do
    command -v "$cmd" >/dev/null 2>&1 || { printf '  %s%s not found%s\n' "$C_RED" "$cmd" "$C_OFF" >&2; exit 1; }
done

if ((!APPLY)); then
    printf '  %s================================================%s\n' "$C_YEL" "$C_OFF"
    printf '  %s %s - no files will be deleted%s\n' "$C_YEL" "$( ((CHECK_ONLY)) && echo 'CHECK ONLY' || echo 'DRY RUN' )" "$C_OFF"
    printf '  %s Add --apply to perform the actual cleanup.%s\n' "$C_YEL" "$C_OFF"
    printf '  %s================================================%s\n\n' "$C_YEL" "$C_OFF"
fi

OS_NAME=$( . /etc/os-release 2>/dev/null && echo "${PRETTY_NAME:-unknown}" )
PBX_VERSION=$(dpkg-query -W -f='${Version}' 3cxpbx 2>/dev/null)
HAS_3CX=0
if ((!SKIP_3CX)) && [[ -n $PBX_VERSION || -d $DATA_DIR ]]; then HAS_3CX=1; fi

note "OS        : ${OS_NAME:-unknown}"
if ((SKIP_3CX)); then
    note '3CX       : skipped (--skip-3cx)'
elif ((HAS_3CX)); then
    note "3CX       : ${PBX_VERSION:-installed}, data in $DATA_DIR$([[ -d $DATA_DIR ]] || echo ' (not found)')"
else
    note '3CX       : not installed - 3CX sections are left out'
fi
note "Older than: $DAYS day(s)"

section 'Disk'
df -h --output=target,size,used,avail,pcent / /var /home "$DATA_DIR" 2>/dev/null | awk '!seen[$0]++' |
    while IFS= read -r line; do note "$line"; done

DISK_BEFORE=$(disk_used)

# ── 1. APT / dpkg ──────────────────────────────────────────────────────────────
section 'Packages'

if ((SKIP_APT)); then
    add_result 'Packages' 'APT cache, autoremove, leftover config' 0 Skipped
elif ! command -v apt-get >/dev/null 2>&1; then
    add_result 'Packages' 'No apt-get (not Debian/Ubuntu)' 0 Skipped
elif apt_busy; then
    warn 'apt/dpkg is running (an OS or 3CX update?) - packages left alone this run.'
    add_result 'Packages' 'APT cache, autoremove, leftover config' 0 Skipped
else
    cache_size() { find /var/cache/apt -type f \( -name '*.deb' -o -name '*.bin' \) -printf '%s\n' 2>/dev/null | sum_sizes; }
    before=$(cache_size)
    if ((APPLY)); then
        apt-get clean >/dev/null 2>&1 || warn 'apt-get clean failed'
        add_result 'Packages' 'APT cache (/var/cache/apt)' $((before - $(cache_size))) Run
    else
        add_result 'Packages' 'APT cache (/var/cache/apt)' "$before" Run
    fi

    if ((SKIP_AUTOREMOVE)); then
        add_result 'Packages' 'autoremove' 0 Skipped
    else
        mapfile -t AR_PKGS < <(apt-get -s autoremove 2>/dev/null | awk '/^Remv /{ print $2 }')
        if ((${#AR_PKGS[@]} == 0)); then
            add_result 'Packages' 'autoremove (nothing to remove)' 0 Run
        elif printf '%s\n' "${AR_PKGS[@]}" | grep -qi '3cx'; then
            warn "autoremove would remove a 3CX package ($(printf '%s\n' "${AR_PKGS[@]}" | grep -i 3cx | paste -sd' ')) - refused."
            add_result 'Packages' 'autoremove (lists a 3CX package)' 0 Skipped
        else
            kb=$(dpkg-query -W -f='${Installed-Size}\n' "${AR_PKGS[@]}" 2>/dev/null | sum_sizes)
            note "Unneeded: $(printf '%s ' "${AR_PKGS[@]}")"
            if ((APPLY)); then
                DEBIAN_FRONTEND=noninteractive apt-get -y autoremove --purge >/dev/null 2>&1 || warn 'apt-get autoremove failed'
            fi
            add_result 'Packages' "autoremove (${#AR_PKGS[@]} package(s))" $((kb * 1024)) Run
        fi
    fi

    # Removed packages whose configuration stayed behind. Small, but it is clutter in dpkg -l.
    mapfile -t RC_PKGS < <(dpkg-query -W -f='${db:Status-Abbrev} ${binary:Package}\n' 2>/dev/null | awk '$1 == "rc" { print $2 }')
    if ((${#RC_PKGS[@]})); then
        rc_bytes=$(dpkg-query -L "${RC_PKGS[@]}" 2>/dev/null | while IFS= read -r f; do [[ -f $f ]] && stat -c %s "$f"; done | sum_sizes)
        note "Leftover config: $(printf '%s ' "${RC_PKGS[@]}")"
        if ((APPLY)); then
            DEBIAN_FRONTEND=noninteractive dpkg --purge "${RC_PKGS[@]}" >/dev/null 2>&1 || warn 'dpkg --purge failed'
        fi
        add_result 'Packages' "Leftover config (${#RC_PKGS[@]} package(s))" "$rc_bytes" Run
    else
        add_result 'Packages' 'Leftover config (none)' 0 Run
    fi
fi

# ── 2. Snap ────────────────────────────────────────────────────────────────────
if command -v snap >/dev/null 2>&1; then
    section 'Snap'
    mapfile -t SNAP_OLD < <(LANG=C snap list --all 2>/dev/null | awk '/disabled/ { print $1 "\t" $3 }')
    bytes=0
    for row in "${SNAP_OLD[@]}"; do
        IFS=$'\t' read -r name rev <<<"$row"
        file=/var/lib/snapd/snaps/${name}_${rev}.snap
        [[ -f $file ]] && bytes=$((bytes + $(stat -c %s "$file")))
        note "$( ((APPLY)) && echo 'Removed' || echo 'Would remove' ): $name revision $rev"
        if ((APPLY)); then snap remove "$name" --revision="$rev" >/dev/null 2>&1 || warn "snap remove $name --revision=$rev failed"; fi
    done
    add_result 'Snap' "Disabled revisions (${#SNAP_OLD[@]})" "$bytes" Run
fi

# ── 3. systemd journal ─────────────────────────────────────────────────────────
section 'systemd journal'

mapfile -t JOURNAL_DIRS < <(existing /var/log/journal /run/log/journal)
journal_size() { ((${#JOURNAL_DIRS[@]})) && du -scb "${JOURNAL_DIRS[@]}" 2>/dev/null | tail -1 | cut -f1 || echo 0; }

if ((SKIP_JOURNAL)); then
    add_result 'Journal' 'systemd journal' 0 Skipped
elif ! command -v journalctl >/dev/null 2>&1; then
    add_result 'Journal' 'systemd journal (no journalctl)' 0 Skipped
else
    before=$(journal_size)
    if ((APPLY)); then
        journalctl --vacuum-time="${DAYS_MIN1}d" --vacuum-size="$JOURNAL_SIZE" >/dev/null 2>&1 || warn 'journalctl --vacuum failed'
        add_result 'Journal' "Journal (to ${DAYS_MIN1}d / $JOURNAL_SIZE)" $((before - $(journal_size))) Run
    else
        add_result 'Journal' "Journal now - vacuum to ${DAYS_MIN1}d / $JOURNAL_SIZE" "$before" Info
    fi
fi

# ── 4. Rotated system logs ─────────────────────────────────────────────────────
section 'Rotated logs'

clean_files 'System logs' "Rotated logs in /var/log (> ${DAYS}d)" \
    /var/log -xdev -regextype posix-extended -type f -not -path '/var/log/journal/*' \
    \( -name '*.gz' -o -name '*.xz' -o -name '*.bz2' -o -name '*.zst' -o -name '*.old' -o -regex '.*\.[0-9]+$' \) \
    "${AGE[@]}"

# ── 5. Crash dumps ─────────────────────────────────────────────────────────────
section 'Crash dumps'

mapfile -t DUMP_DIRS < <(existing /var/lib/systemd/coredump /var/crash)
if ((${#DUMP_DIRS[@]})); then
    clean_files 'Crash dumps' "systemd-coredump, /var/crash (> ${DAYS}d)" "${DUMP_DIRS[@]}" -xdev -type f "${AGE[@]}"
else
    add_result 'Crash dumps' 'No crash dump folder' 0 Run
fi

# ── 6. Temp ────────────────────────────────────────────────────────────────────
section 'Temp files'

# systemd-private-* belongs to running services; .s.PGSQL.* is PostgreSQL's socket lock.
clean_files 'Temp' "/tmp, /var/tmp (> ${DAYS_MIN1}d)" \
    /tmp /var/tmp -xdev -type f -not -path '*/systemd-private-*' -not -name '.s.PGSQL.*' \
    "${AGE_MIN1[@]}"

# ── 7. User caches and trash ───────────────────────────────────────────────────
section 'User caches and trash'

if ((SKIP_USER_CACHE)); then
    add_result 'User cache' 'Caches and trash' 0 Skipped
else
    mapfile -t HOMES < <(existing /root /home/*)
    found=0
    for home in "${HOMES[@]}"; do
        [[ -d $home ]] || continue
        mapfile -t CACHE_DIRS < <(existing "$home/.cache" "$home/.npm/_cacache" "$home/.local/share/Trash")
        ((${#CACHE_DIRS[@]})) || continue
        found=1
        clean_files 'User cache' "${home} cache/trash (> ${DAYS_MIN1}d)" "${CACHE_DIRS[@]}" -xdev -type f "${AGE_MIN1[@]}"
    done
    ((found)) || add_result 'User cache' 'No user caches' 0 Run
fi

# ── 8. Docker ──────────────────────────────────────────────────────────────────
if command -v docker >/dev/null 2>&1 && docker info >/dev/null 2>&1; then
    section 'Docker'
    if ((!DOCKER)); then
        reclaimable=$(docker system df --format '{{.Type}}: {{.Reclaimable}}' 2>/dev/null | paste -sd',' | sed 's/,/, /g')
        note "Reclaimable per docker: ${reclaimable:-unknown}"
        add_result 'Docker' 'docker system prune - add --docker' 0 Skipped
    else
        before=$(docker system df --format '{{.Size}}' 2>/dev/null | numfmt --from=iec --invalid=ignore 2>/dev/null | sum_sizes)
        if ((APPLY)); then
            docker system prune -f >/dev/null 2>&1 || warn 'docker system prune failed'
            after=$(docker system df --format '{{.Size}}' 2>/dev/null | numfmt --from=iec --invalid=ignore 2>/dev/null | sum_sizes)
            add_result 'Docker' 'Containers, networks, dangling images, cache' $((before - after)) Run
        else
            note "Would prune; per docker: $(docker system df --format '{{.Type}}: {{.Reclaimable}}' 2>/dev/null | paste -sd',' | sed 's/,/, /g')"
            add_result 'Docker' 'docker system prune (size: see line above)' 0 Info
        fi
    fi
fi

# ── 9. 3CX ─────────────────────────────────────────────────────────────────────
if ((HAS_3CX)); then
    if [[ -d $DATA_DIR ]]; then
        section '3CX data (report only)'
        du -sb "$DATA_DIR"/* 2>/dev/null | sort -rn | head -8 |
            while IFS=$'\t' read -r bytes path; do note "$(printf '%-10s %s' "$(fmt "$bytes")" "${path#"$DATA_DIR"/}")"; done
        if [[ -d $DATA_DIR/Recordings ]]; then
            add_result '3CX recordings' 'Recordings (never deleted here)' "$(du -sb "$DATA_DIR/Recordings" 2>/dev/null | cut -f1)" Info
        fi
    fi

    section '3CX logs'
    mapfile -t PBX_LOG_DIRS < <(existing "$DATA_DIR/Logs" /var/lib/3cxpbx/Bin/nginx/logs)
    if ((${#PBX_LOG_DIRS[@]} == 0)); then
        add_result '3CX logs' 'No 3CX log folder found' 0 Run
    fi
    for dir in "${PBX_LOG_DIRS[@]}"; do
        clean_files '3CX logs' "${dir#/var/lib/3cxpbx/} (> ${DAYS}d)" "$dir" -xdev -type f "${AGE[@]}"
    done

    BACKUP_DIR=$DATA_DIR/Backups
    if [[ -d $BACKUP_DIR ]]; then
        section '3CX backups'
        mapfile -t BACKUPS < <(find "$BACKUP_DIR" -maxdepth 1 -type f -name '*.zip' -printf '%T@\t%s\t%p\n' 2>/dev/null | sort -rn)
        total=$( ((${#BACKUPS[@]})) && printf '%s\n' "${BACKUPS[@]}" | cut -f2 | sum_sizes || echo 0 )
        if [[ -z $KEEP_BACKUPS ]]; then
            add_result '3CX backups' "${#BACKUPS[@]} backup(s) - add --keep-backups N" "$total" Info
        else
            bytes=0; count=0
            for ((i = KEEP_BACKUPS; i < ${#BACKUPS[@]}; i++)); do
                IFS=$'\t' read -r _ size path <<<"${BACKUPS[i]}"
                note "$( ((APPLY)) && echo 'Removed' || echo 'Would remove' ): ${path##*/}"
                if ((APPLY)); then rm -f -- "$path" || { warn "could not remove $path"; continue; }; fi
                bytes=$((bytes + size)); count=$((count + 1))
            done
            add_result '3CX backups' "Beyond newest $KEEP_BACKUPS of ${#BACKUPS[@]} ($count file(s))" "$bytes" Run
        fi
    fi
fi

# ── 10. Largest files (report only) ────────────────────────────────────────────
section "Files over ${LARGE_MB} MB (report only)"

mapfile -t LARGE < <(find / /var /home "$DATA_DIR" -xdev -type f -size +"${LARGE_MB}"M -printf '%s\t%p\n' 2>/dev/null | sort -u | sort -rn | head -10)
if ((${#LARGE[@]})); then
    for row in "${LARGE[@]}"; do
        IFS=$'\t' read -r size path <<<"$row"
        note "$(printf '%-10s %s' "$(fmt "$size")" "$path")"
    done
else
    note 'None.'
fi

# ── Summary ────────────────────────────────────────────────────────────────────
DISK_AFTER=$(disk_used)

declare -A BY_CAT=()
CAT_ORDER=()
TOTAL=0
for i in "${!R_CAT[@]}"; do
    [[ ${R_STATUS[i]} == Skipped || ${R_STATUS[i]} == Info ]] && continue
    ((R_BYTES[i] > 0)) || continue
    [[ -n ${BY_CAT[${R_CAT[i]}]+x} ]] || CAT_ORDER+=("${R_CAT[i]}")
    BY_CAT[${R_CAT[i]}]=$(( ${BY_CAT[${R_CAT[i]}]:-0} + R_BYTES[i] ))
    TOTAL=$((TOTAL + R_BYTES[i]))
done

printf '\n  %s================================================%s\n' "$C_CYAN" "$C_OFF"
printf '  %s Summary%s\n' "$C_CYAN" "$C_OFF"
printf '  %s================================================%s\n\n' "$C_CYAN" "$C_OFF"

for cat in "${CAT_ORDER[@]}"; do
    printf '%s\t%s\n' "${BY_CAT[$cat]}" "$cat"
done | sort -rn | while IFS=$'\t' read -r bytes cat; do
    printf '  %s%-30s %10s%s\n' "$C_GRY" "$cat" "$(fmt "$bytes")" "$C_OFF"
done

printf '\n  %s%-30s %s%s\n' "$C_GRY" '------------------------------' '----------' "$C_OFF"
if ((APPLY)); then
    printf '  %s%-30s %10s%s\n' "$C_GRN" 'Total freed (reported)' "$(fmt "$TOTAL")" "$C_OFF"
    printf '  %s%-30s %10s%s\n' "$C_GRN" 'Total freed (measured)' "$(fmt $((DISK_BEFORE - DISK_AFTER)))" "$C_OFF"
else
    printf '  %s%-30s %10s%s\n' "$C_YEL" 'Reclaimable (estimate)' "$(fmt "$TOTAL")" "$C_OFF"
    printf '\n  %sThe journal and Docker are not in the estimate. Run with --apply to perform the cleanup.%s\n' "$C_YEL" "$C_OFF"
fi
printf '\n'

# --check-only: 2 = something can be freed, 0 = nothing to do.
((CHECK_ONLY && TOTAL > 0)) && exit 2
exit 0
