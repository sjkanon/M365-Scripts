# Audio — Disable Internal Microphone

> Author: Sjoerd Kanon | Date: 19/03/2026

---

## Background and problem statement

The customer uses a **browser-based VoIP/dialler application**. Employees were hearing background noise during calls even when a headset was plugged in.

The cause: Windows sometimes selects the **internal laptop microphone** as the audio source instead of the headset, even when the headset is correctly connected. This behaviour occurs in browser-based applications because the browser itself chooses the audio source based on Windows settings.

Applications with their own audio management (such as Microsoft Teams) do not have this problem. Browser-based applications rely entirely on Windows settings and sometimes pick up the wrong source.

### Timeline

| Step | Action |
|---|---|
| Step 1 | Ticket opened: complaint about background noise in dialler application |
| Step 2 | Test phase on a few laptops — partially successful |
| Step 3 | Decision: disable internal microphone via script |
| Step 4 | First script deployed via NinjaOne — too broad, headset also blocked |
| Step 5 | Rollback executed on affected laptop |
| Step 6 | Detect script deployed for inventory |
| Step 7 | Disable + Rollback script built based on detect data |

---

## Why the first script failed

The script deployed on 04/03/2026 blocked microphone permissions via the Windows Privacy / Consent Store mechanism:

```powershell
# WRONG — too broad
$regPath = "HKLM:\...\CapabilityAccessManager\ConsentStore\microphone"
Set-ItemProperty -Path $regPath -Name "Value" -Value "Deny"
```

This blocks **all** microphone input for **all** applications — including the headset. After deployment the affected user could not make calls at all. The script had to be rolled back.

### Old vs new approach

| | Old script (wrong) | New script (correct) |
|---|---|---|
| Method | Registry privacy policy | PnP device disable |
| Scope | All microphones | Internal microphone only |
| Headset | ❌ Also blocked | ✅ Remains active |
| Rollback | Difficult | Simple via separate script |
| New laptops | Works automatically | Re-deployment needed at onboarding |

---

## Three-phase approach

```
Phase 1: Detect  ✅ DONE (16/66 devices)
  └── Detect-AudioDevices.ps1 deployed via NinjaOne
  └── Custom field AudioDeviceInventory populated per device
  └── Patterns for internal microphones identified

Phase 2: Disable  ← NEXT STEP
  └── Deploy Disable-InternalMic.ps1
  └── Test phase on 5-10 laptops first
  └── Then deploy to all devices

Phase 3: Monitoring
  └── Confirmation from users that headset works
  └── Deploy to remaining 50 devices without detect data
  └── Include in onboarding procedure for new laptops
```

---

## Detect analysis results (19/03/2026)

Of the deployed devices, the detect script retrieved data. Devices that were offline or on which the script did not run have no data in the custom field.

### Internal microphone names found

Based on the 16 devices, the following internal microphones were identified:

| Name | Number of devices | Action |
|---|---|---|
| `Microfoonmatrix (Realtek High Definition Audio)` | 10 | Disable |
| `Microphone (2- High Definition Audio Device)` | 4 | Disable |
| `Microphone Array (Realtek High Definition Audio)` | 3 | Disable |
| `Microphone (Realtek(R) Audio)` | 1 | Disable |
| `Microfoonmatrix (Realtek(R) Audio)` | 1 | Disable |
| `Microfoonmatrix (Synaptics Audio)` | 1 | Disable |

### Headsets found (NEVER touched)

The following headset brands are present in the environment and are always skipped by the script:

- EPOS IMPACT DW
- EPOS IMPACT 60
- EPOS IMPACT D
- Yealink WH64
- Plantronics Blackwire 3225 Series

> Update the safelist in the script if new headset brands are deployed.

### Consent Store status

✅ No device still has `ConsentStore\microphone = Deny` — the first incorrect script has been correctly rolled back everywhere.

### Unknown microphones

On some devices unclassified devices may appear. Check the `UNKNOWN MICROPHONES` section in the detect output per device and manually determine whether it is an internal device or an external headset.

---

## Files

```
scripts/
└── audio/
    ├── Detect-AudioDevices.ps1      ← phase 1: inventory audio devices
    ├── Disable-InternalMic.ps1      ← phase 2: disable internal microphone
    ├── Rollback-InternalMic.ps1     ← emergency: re-enable microphone
    └── readme.md
```

---

## Detect-AudioDevices.ps1

### What it does

Creates a full inventory of all audio devices on the laptop and automatically classifies them as internal or headset. Writes the result to the NinjaOne custom field `AudioDeviceInventory`.

### Classification logic

The script splits microphones based on InstanceId:
- `{0.0.1...}` = capture device (microphone)
- `{0.0.0...}` = render device (speaker/output)

Microphones are then classified based on name patterns:

**Internal** (flagged for disabling):
`Array`, `Intel`, `Realtek`, `Synaptics`, `High Definition Audio`, `Camera`, `Webcam`, `Integrated`, `Microfoonmatrix`, `Microphone (`

**Headset** (never touched):
`EPOS`, `Jabra`, `Yealink`, `Plantronics`, `Poly`, `Logitech`, `Headset`, `hoofdtelefoon`, `oortelefoon`

### NinjaOne deployment

#### Create script

1. NinjaOne → **Administration** → **Library** → **Scripting**
2. Click **Add** → **Script**
3. Settings:
   - **Name**: `BNC - Detect Audio Devices`
   - **Language**: PowerShell
   - **OS**: Windows
   - **Architecture**: 64-bit
   - **Run As**: System
   - **Timeout**: 60 seconds
4. Paste content of `Detect-AudioDevices.ps1`
5. **Save**

#### Deploy

1. NinjaOne → **Devices** → filter on **Best Next Contact**
2. Select all Windows laptops
3. **Run Script** → select `BNC - Detect Audio Devices`
4. **Run**

#### Read output

Via custom field (recommended for 66 devices at once):
1. NinjaOne → **Reports** → **Device Report**
2. Add column `AudioDeviceInventory`
3. Export to CSV — all output in one file

Per device individually:
1. Open the device → **Activities**
2. Click on the script run → **Details**

---

## Disable-InternalMic.ps1

### What it does

Disables the internal microphone device-specifically based on the patterns from the detect analysis. The headset is **never** touched.

### Logic per microphone (4 checks)

```
For each microphone:
  1. Is the name in the headset safelist?
     → YES: skip, never touch
  2. Does the name match an internal pattern?
     → NO: skip, unknown device
  3. Is the device already disabled (Status = Unknown/Error)?
     → YES: skip, already handled
  4. Disable via Disable-PnpDevice -InstanceId
     → Success: log DISABLED
     → Error: log FAILED, exit 1
```

### Internal patterns that are disabled

```
*Microfoonmatrix*
*Microphone Array*
*Microphone*Realtek*
*Microphone*High Definition Audio*
*Microfoon*Realtek*
*Microfoon*Synaptics*
*Microphone (2-*
```

### Headset safelist (NEVER touched)

```
EPOS, Jabra, Plantronics, Yealink, Poly, Logitech,
Headset, hoofdtelefoon, oortelefoon, Blackwire,
Voyager, Sennheiser, Astro, SteelSeries
```

### NinjaOne deployment

#### Create script

1. NinjaOne → **Administration** → **Library** → **Scripting**
2. Click **Add** → **Script**
3. Settings:
   - **Name**: `BNC - Disable Internal Mic`
   - **Language**: PowerShell
   - **OS**: Windows
   - **Architecture**: 64-bit
   - **Run As**: System
   - **Timeout**: 60 seconds
4. Paste content of `Disable-InternalMic.ps1`
5. **Save**

#### Deployment strategy (phased)

> **Important:** do not deploy to all 66 devices at once. Work in phases to avoid repeating previous incidents.

**Test phase (5–10 laptops):**
1. Select 5–10 laptops for which detect data is already available
2. Run Script → `BNC - Disable Internal Mic`
3. Call the users to confirm their headset still works
4. Wait at least half a day before continuing

**Full deployment (after test phase confirmation):**
1. Select all remaining laptops
2. Run Script → `BNC - Disable Internal Mic`
3. Check output via the `AudioDeviceInventory` custom field

#### Exit codes

| Code | Meaning |
|---|---|
| `0` | All successful |
| `1` | One or more devices could not be disabled — check output |

---

## Rollback-InternalMic.ps1

### When to use

Deploy immediately on a specific device if an employee reports their headset stopped working after the disable script. Re-enables exactly the same devices that the disable script turned off.

### NinjaOne deployment

1. NinjaOne → open the **specific device** with the issue
2. **Run Script** → select `BNC - Rollback Internal Mic`
3. **Run**
4. Call the user to confirm the headset is working again

> ⚠️ Only deploy the rollback script on the device with the issue, not on all devices at once.

---

## Devices without detect data

Devices that were offline during the detect deployment have no data in the custom field. Procedure:

1. NinjaOne → **Devices** → filter on the customer organisation
2. Filter on custom field `AudioDeviceInventory` = empty
3. Check whether the devices are online
4. Re-deploy the detect script on these devices
5. Repeat until all devices have data

---

## Onboarding new laptops

When a new laptop is added:
1. Deploy the detect script to confirm the internal microphone name
2. Check whether the name matches an existing pattern in the disable script
3. If yes: deploy the disable script
4. If no: add the pattern to `$internalPatterns` in the disable script and increment the version in the changelog

---

## Changelog

| Date | Version | Change |
|---|---|---|
| 04/03/2026 | — | First (incorrect) script deployed — ConsentStore = Deny, too broad |
| 05/03/2026 | — | Rollback executed on affected laptop |
| 19/03/2026 | 1.0 | Detect script written and deployed |
| 19/03/2026 | 1.1 | Detect output analysed — 6 internal microphone patterns found |
| 19/03/2026 | 1.2 | Disable script built based on detect data |
| 19/03/2026 | 1.3 | Rollback script added |
| 19/03/2026 | 1.4 | README made generic for reuse with other customers |
