# ============================================
# BraveHub - Audio Device Detection Script
# Ticket: #0250981 - Best Next Contact BVBA
# Doel: Inventariseer alle audio devices op de laptop
#       zodat we kunnen bepalen welke interne microfoon
#       disabled moet worden zonder de headset te raken.
# Uitrol: NinjaOne - Run as SYSTEM
# ============================================

$output = [System.Collections.Generic.List[string]]::new()
$output.Add("=== AUDIO DEVICE INVENTORY ===")
$output.Add("Hostname   : $env:COMPUTERNAME")
$output.Add("Datum      : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')")
$output.Add("Gebruiker  : $env:USERNAME")
$output.Add("")

# ── Alle AudioEndpoint devices (microfoons EN speakers) ──────────────────────
$output.Add("--- ALLE AUDIO ENDPOINTS (AudioEndpoint class) ---")
$allAudio = Get-PnpDevice -Class AudioEndpoint -ErrorAction SilentlyContinue

if ($allAudio) {
    foreach ($dev in $allAudio) {
        $output.Add("  Status      : $($dev.Status)")
        $output.Add("  FriendlyName: $($dev.FriendlyName)")
        $output.Add("  InstanceId  : $($dev.InstanceId)")
        $output.Add("  ---")
    }
} else {
    $output.Add("  GEEN AudioEndpoint devices gevonden")
}

$output.Add("")

# ── Classificatie: microfoons vs speakers ─────────────────────────────────────
# InstanceId bevat {0.0.1...} = capture (microfoon)
#                 {0.0.0...} = render (speaker/output)
$output.Add("--- MICROFOONS (capture devices - 0.0.1) ---")
$mics = $allAudio | Where-Object { $_.InstanceId -like "*0.0.1*" }
if ($mics) {
    foreach ($mic in $mics) {
        $output.Add("  [$($mic.Status)] $($mic.FriendlyName)")
        $output.Add("           $($mic.InstanceId)")
        $output.Add("  ---")
    }
} else {
    $output.Add("  GEEN microfoons gevonden")
}

$output.Add("")
$output.Add("--- SPEAKERS / OUTPUT (render devices - 0.0.0) ---")
$speakers = $allAudio | Where-Object { $_.InstanceId -like "*0.0.0*" }
if ($speakers) {
    foreach ($spk in $speakers) {
        $output.Add("  [$($spk.Status)] $($spk.FriendlyName)")
        $output.Add("           $($spk.InstanceId)")
        $output.Add("  ---")
    }
} else {
    $output.Add("  GEEN speakers gevonden")
}

$output.Add("")

# ── Verdacht: mogelijke interne microfoons ────────────────────────────────────
# Patronen die typisch interne microfoons zijn (geen headset)
$internalPatterns = @(
    "*Array*",
    "*Intern*",
    "*Internal*",
    "*Intel*",
    "*Realtek*",
    "*HD Audio*",
    "*Laptop*",
    "*Built*",
    "*Camera*",
    "*Webcam*",
    "*Integrated*"
)

$output.Add("--- VERMOEDELIJK INTERNE MICROFOONS ---")
$suspectedInternal = $mics | Where-Object {
    $name = $_.FriendlyName
    $match = $false
    foreach ($pattern in $internalPatterns) {
        if ($name -like $pattern) { $match = $true; break }
    }
    $match
}

if ($suspectedInternal) {
    foreach ($dev in $suspectedInternal) {
        $output.Add("  [$($dev.Status)] $($dev.FriendlyName)")
        $output.Add("           $($dev.InstanceId)")
        $output.Add("  ---")
    }
} else {
    $output.Add("  Geen gevonden op basis van naampatronen")
}

$output.Add("")

# ── Headsets / externe microfoons ─────────────────────────────────────────────
$headsetPatterns = @(
    "*EPOS*",
    "*Jabra*",
    "*Plantronics*",
    "*Yealink*",
    "*Poly*",
    "*Logitech*",
    "*Headset*",
    "*Headphone*",
    "*hoofdtelefoon*",
    "*oortelefoon*",
    "*USB*"
)

$output.Add("--- VERMOEDELIJK HEADSETS / EXTERNE MICROFOONS ---")
$headsets = $mics | Where-Object {
    $name = $_.FriendlyName
    $match = $false
    foreach ($pattern in $headsetPatterns) {
        if ($name -like $pattern) { $match = $true; break }
    }
    $match
}

if ($headsets) {
    foreach ($dev in $headsets) {
        $output.Add("  [$($dev.Status)] $($dev.FriendlyName)")
        $output.Add("           $($dev.InstanceId)")
        $output.Add("  ---")
    }
} else {
    $output.Add("  Geen headset microfoons gevonden (of niet ingeplugd)")
}

$output.Add("")

# ── Huidige privacy/consent instelling microfoon ──────────────────────────────
$output.Add("--- MICROFOON PRIVACY INSTELLING (registry) ---")
$regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone"
try {
    $regValue = Get-ItemProperty -Path $regPath -Name "Value" -ErrorAction Stop
    $output.Add("  ConsentStore\microphone\Value = $($regValue.Value)")
    if ($regValue.Value -eq "Deny") {
        $output.Add("  STATUS: Microfoon GEBLOKKEERD via privacy policy (te breed - blokkeert ook headset!)")
    } else {
        $output.Add("  STATUS: Microfoon toegang TOEGESTAAN via privacy policy")
    }
} catch {
    $output.Add("  Registry key niet gevonden - standaard instelling actief (Allow)")
}

$output.Add("")

# ── Onbekende microfoons (niet intern, niet headset) ──────────────────────────
$output.Add("--- ONBEKENDE MICROFOONS (niet geclassificeerd) ---")
$unknown = $mics | Where-Object {
    $dev = $_
    $isInternal = $suspectedInternal | Where-Object { $_.InstanceId -eq $dev.InstanceId }
    $isHeadset  = $headsets         | Where-Object { $_.InstanceId -eq $dev.InstanceId }
    (-not $isInternal) -and (-not $isHeadset)
}

if ($unknown) {
    foreach ($dev in $unknown) {
        $output.Add("  [$($dev.Status)] $($dev.FriendlyName)")
        $output.Add("           $($dev.InstanceId)")
        $output.Add("  ---")
    }
} else {
    $output.Add("  Alle microfoons zijn geclassificeerd")
}

$output.Add("")
$output.Add("=== EINDE RAPPORT ===")

# ── Output naar NinjaOne ──────────────────────────────────────────────────────
$result = $output -join "`n"
Write-Host $result

# NinjaOne custom field instellen (optioneel - pas naam aan indien nodig)
# Ninja-Property-Set AudioDeviceInventory $result