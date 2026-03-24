# Audio — Disable Internal Microphone

Scripts for detecting and disabling the internal microphone on Windows laptops where employees use a browser-based VoIP application.

**Problem:** Windows sometimes selects the internal laptop microphone as the audio input instead of the connected headset, causing background noise in browser-based diallers. This does not affect Teams or other apps with their own audio management.

**Solution:** Disable the internal microphone at the PnP device level — the headset is never touched.

---

## Files

| File | Description |
|------|-------------|
| `Detect-AudioDevices.ps1` | Phase 1 — inventory all audio devices and classify as internal or headset |
| `Disable-InternalMic.ps1` | Phase 2 — disable internal microphone based on detection patterns |
| `Rollback-InternalMic.ps1` | Emergency — re-enable the microphone if something goes wrong |

---

## Detect-AudioDevices.ps1

Creates a full inventory of all audio devices on the endpoint and classifies them as internal microphone or headset. Output can be written to a RMM custom field or printed to screen.

**Classification logic**

Microphones are classified based on name patterns:

| Type | Patterns |
|------|----------|
| Internal (disable) | `Array`, `Intel`, `Realtek`, `Synaptics`, `High Definition Audio`, `Camera`, `Webcam`, `Integrated`, `Microfoonmatrix`, `Microphone (` |
| Headset (never touch) | `EPOS`, `Jabra`, `Yealink`, `Plantronics`, `Poly`, `Logitech`, `Headset` |

**Deployment**

Run as SYSTEM with administrator privileges. Compatible with NinjaOne, Datto RMM, or manual execution.

```powershell
.\Detect-AudioDevices.ps1
```

---

## Disable-InternalMic.ps1

Disables internal microphone devices based on the patterns from the detect analysis. The headset safelist is always checked first — headsets are never disabled.

**Logic per microphone**

```
1. Is the name in the headset safelist?   → YES: skip
2. Does the name match an internal pattern? → NO: skip (unknown device)
3. Is the device already disabled?          → YES: skip
4. Disable via Disable-PnpDevice -InstanceId
```

**Internal patterns disabled**

```
*Microfoonmatrix*, *Microphone Array*, *Microphone*Realtek*,
*Microphone*High Definition Audio*, *Microfoon*Realtek*,
*Microfoon*Synaptics*, *Microphone (2-*
```

**Headset safelist (never touched)**

```
EPOS, Jabra, Plantronics, Yealink, Poly, Logitech,
Headset, Blackwire, Voyager, Sennheiser, Astro, SteelSeries
```

**Exit codes**

| Code | Meaning |
|------|---------|
| `0` | All successful |
| `1` | One or more devices could not be disabled |

> Deploy in phases — test on 5–10 devices first before rolling out to all endpoints.

---

## Rollback-InternalMic.ps1

Re-enables the internal microphone on a specific device. Deploy immediately if a user reports their audio stopped working after the disable script.

```powershell
.\Rollback-InternalMic.ps1
```

> Only deploy the rollback on the affected device, not on all devices at once.

---

## Onboarding new laptops

When a new laptop model is added to the environment:

1. Run `Detect-AudioDevices.ps1` to identify the internal microphone name
2. Check if the name matches an existing pattern in `Disable-InternalMic.ps1`
3. If yes: deploy the disable script
4. If no: add the new pattern to `$internalPatterns` in the disable script
