# RCA - Intermittent SAS WORK Delete Failures

Date: 2026-04-15  
Owner: SAS platform operations

## 1) Incident Summary

SAS batch processing showed intermittent failures related to WORK directory cleanup on G:\sas\work.

Observed symptoms:
- SAS Application events with hc_disk_delete_library Access is denied while deleting file
- SAS Application events with hc_disk_delete Return code from system: 5
- Directory cannot be deleted during WORK cleanup

Impact:
- Unpredictable SAS batch instability
- Potential batch aborts or partial job failure when temporary WORK artifacts cannot be removed

## 2) Evidence Collected

### 2.1 SAS Event Log Evidence

Key events in Application log:
- 2026-04-14 21:37:40
- Provider: SAS
- Errors:
  - hc_disk_delete_library Access is denied while deleting file in G:\sas\work
  - hc_disk_delete Unknown error detected while deleting file; return code 5

Interpretation:
- Return code 5 on Windows equals Access denied.
- Failure occurs during delete operation, not during initial create/read in the short health checks.

### 2.2 Security Log Evidence

Security log search did not return matching 4656/4663 entries for the WORK path in the checked windows.

Interpretation:
- This can happen when Object Access auditing and/or SACL is not configured for G:\sas\work.
- Absence of Security events does not disprove real access-denied behavior at file-system/filter layer.

### 2.3 AV/EDR Evidence

Defender/EDR diagnostics showed:
- Defender enabled (real-time, behavior monitoring, AV enabled)
- Exclusions configured for both paths:
  - G:\sas\work
  - U:\sas\userwork
- Multiple Defender Operational Event ID 1121 entries
  - Rule ID: D1E49AAC-8F56-4280-B9BA-993A6D77406C
  - Process chain: WmiPrvSE.exe spawning cmd.exe/powershell commands
  - Message: Exploit Guard blocked an operation

Interpretation:
- Rule ID D1E49AAC-8F56-4280-B9BA-993A6D77406C corresponds to ASR protection for process creation from WMI/PSExec.
- The captured 1121 events are real policy blocks, but mostly reference WMI automation commands and admin$/c$ output paths, not direct sas.exe delete actions in G:\sas\work.

### 2.4 Defender Config Change Evidence

Defender Event ID 5007 at 2026-04-14 21:21:30:
- DLP config tag changed
- WdConfigHash changed

Interpretation:
- Indicates Defender policy/config refresh.
- By itself this is not proof of malicious activity or of direct SAS blocking.
- Time proximity to the 21:37 SAS delete error makes it a contributing context item worth correlation.

### 2.5 Disk and Capacity Evidence

Health checks reported:
- G: and U: are fixed NTFS volumes
- Both are treated as ephemeral scratch in this environment
- Very high free capacity during tests
- Short synthetic I/O loops (10/10) succeeded on both paths

Interpretation:
- Not a persistent disk-capacity problem.
- Behavior is intermittent/transient and likely timing-dependent under real batch conditions.

## 3) Root Cause Assessment

### Most likely root cause

Intermittent file lock or filter-driver interference during SAS WORK cleanup, resulting in Access denied (system code 5) when SAS attempts to delete temporary files/directories.

### Why this is most likely

- SAS errors are explicit delete-time access denied failures.
- Failures are intermittent, not constant.
- Basic I/O tests can pass while historical batch windows still show cleanup failures.
- AV/EDR policy activity is present on host and can influence file/process behavior even with path exclusions.

### What is not supported as primary root cause

- Kerberos ticket expiry (tickets were valid at failure times in earlier checks).
- Disk full or structural disk failure (capacity and system disk error checks were clean).
- Stable NTFS permission misconfiguration causing permanent failure (because operations also succeed repeatedly).

## 4) Contributing Factors

- Ephemeral scratch design on both G: and U: (no persistent failover benefit between them).
- Potential policy refresh moments (Defender 5007) near incident windows.
- Limited Security auditing on target path, reducing direct forensic attribution.

## 5) Corrective Actions

### Immediate

1. Keep G:\sas\work and U:\sas\userwork exclusions in Defender and validate centrally managed policy does not override them.
2. Correlate exact incident windows (for example ±15 minutes around SAS error timestamp) against:
   - Defender Operational IDs 1121/1122/1123/1124/5007
   - Process and command-line details in event message
3. Enable Object Access auditing plus SACL on WORK path for temporary forensic period.

### Near-term hardening

1. Add monitoring that extracts and reports, per incident window:
   - Blocking rule ID
   - Blocking process name
   - Command line
   - Involved file/path when present
2. Coordinate with security team to review ASR rule D1E49AAC-8F56-4280-B9BA-993A6D77406C policy scope and exceptions for trusted SAS operational flows if needed.
3. Run extended workload-aligned tests during real batch windows, not only short synthetic loops.

### Operational decision guidance

- Moving from G: to U: is not a structural mitigation when both are ephemeral scratch.
- Priority should be reducing transient block/lock conditions and improving forensic visibility.

## 6) Verification Plan

Success criteria:
- No new hc_disk_delete access denied events in SAS Application log over at least 3 full batch cycles.
- No correlated Defender/ASR block events against SAS-related process/path during those windows.
- No unexpected batch aborts linked to WORK cleanup.

Recommended review cadence:
- Daily review for first week
- Weekly thereafter for one month

## 7) Current Status

Status: Probable root cause identified, correlation confidence medium-high.  
Next required step: tighten time-window correlation and policy review with security team to reach final closure.
