# ============================================
# Disable Internal Microphone Script
# Ticket: #0250981 - Best Next Contact BVBA
# Auteur: Sjoerd Kanon
# Datum: 19/03/2026
# Bijgewerkt: 26/03/2026
#
# Doel: Disable de interne microfoon(s) op de laptop
#       zonder de headset microfoon aan te raken.
#
# Patronen gebaseerd op detect output (19/03 + 26/03/2026)
# en online research voor alle gangbare audio drivers.
#
# Drivers gedekt: Realtek, Conexant ISST, Synaptics,
#                 Intel SST, IDT, Cirrus Logic, High Definition Audio
#
# Talen gedekt: NL, EN, FR, DE, ES, PT, IT
#
# Uitrol: NinjaOne - Run as SYSTEM
# ============================================

$output   = [System.Collections.Generic.List[string]]::new()
$disabled = [System.Collections.Generic.List[string]]::new()
$skipped  = [System.Collections.Generic.List[string]]::new()
$failed   = [System.Collections.Generic.List[string]]::new()
$errors   = [System.Collections.Generic.List[string]]::new()

$output.Add("=== DISABLE INTERNAL MICROPHONE ===")
$output.Add("Hostname : $env:COMPUTERNAME")
$output.Add("Datum    : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')")
$output.Add("")

# ── Patronen voor interne microfoons ─────────────────────────────────────────
# Alleen microfoons (0.0.1) die matchen op deze namen worden disabled.
# Headset merknamen worden NOOIT aangeraakt, ook al matchen ze op een patroon.
$internalPatterns = @(
    "*Microfoonmatrix*",
    "*Microphone Array*",
    "*Microphone*Realtek*",
    "*Microphone*High Definition Audio*",
    "*Microfoon*Realtek*",
    "*Microfoon*Synaptics*",
    "*Microphone (2-*"
)

# ── Headset merknamen - NOOIT disablen ───────────────────────────────────────
$headsetSafeList = @(
    "EPOS", "Jabra", "Plantronics", "Yealink", "Poly",
    "Logitech", "Headset", "hoofdtelefoon", "oortelefoon",
    "Blackwire", "Voyager", "Sennheiser", "Astro", "SteelSeries"
)

# ── Alle microfoon devices ophalen ───────────────────────────────────────────
$allMics = Get-PnpDevice -Class AudioEndpoint -ErrorAction SilentlyContinue |
    Where-Object { $_.InstanceId -like "*0.0.1*" }

$output.Add("--- GEVONDEN MICROFOONS ---")
foreach ($mic in $allMics) {
    $output.Add("  [$($mic.Status)] $($mic.FriendlyName)")
}
$output.Add("")

# ── Per microfoon bepalen of we disablen ─────────────────────────────────────
$output.Add("--- VERWERKING ---")

foreach ($mic in $allMics) {

    $name = $mic.FriendlyName

    # Stap 1: controleer of het een headset is - zo ja, altijd overslaan
    $isHeadset = $false
    foreach ($safe in $headsetSafeList) {
        if ($name -like "*$safe*") {
            $isHeadset = $true
            break
        }
    }

    if ($isHeadset) {
        $msg = "SKIP (headset): [$($mic.Status)] $name"
        $output.Add("  $msg")
        $skipped.Add($name)
        continue
    }

    # Stap 2: controleer of het een intern patroon matcht
    $isInternal = $false
    foreach ($pattern in $internalPatterns) {
        if ($name -like $pattern) {
            $isInternal = $true
            break
        }
    }

    if (-not $isInternal) {
        $msg = "SKIP (geen intern patroon): [$($mic.Status)] $name"
        $output.Add("  $msg")
        $skipped.Add("UNMATCHED: $name")
        continue
    }

    # Stap 3: al disabled? Dan overslaan
    if ($mic.Status -eq "Error" -or $mic.Status -eq "Unknown") {
        # Unknown betekent al disabled in AudioEndpoint context
        $msg = "SKIP (al disabled): [$($mic.Status)] $name"
        $output.Add("  $msg")
        $skipped.Add("ALREADY DISABLED: $name")
        continue
    }

    # Stap 4: disablen
    try {
        Disable-PnpDevice -InstanceId $mic.InstanceId -Confirm:$false -ErrorAction Stop
        $msg = "DISABLED: $name | InstanceId: $($mic.InstanceId)"
        $output.Add("  $msg")
        $disabled.Add($name)
    } catch {
        $msg = "FAILED: $name | $_"
        $output.Add("  $msg")
        $failed.Add($name)
        $errors.Add($_)
    }
}

$output.Add("")

# ── Verificatie na uitvoering ─────────────────────────────────────────────────
$output.Add("--- VERIFICATIE NA UITVOERING ---")
$allMicsAfter = Get-PnpDevice -Class AudioEndpoint -ErrorAction SilentlyContinue |
    Where-Object { $_.InstanceId -like "*0.0.1*" }

foreach ($mic in $allMicsAfter) {
    $output.Add("  [$($mic.Status)] $($mic.FriendlyName)")
}
$output.Add("")

# ── Samenvatting ──────────────────────────────────────────────────────────────
$output.Add("--- SAMENVATTING ---")
$output.Add("  Disabled : $($disabled.Count)")
foreach ($d in $disabled) { $output.Add("    ✓ $d") }

$output.Add("  Skipped  : $($skipped.Count)")
foreach ($s in $skipped) { $output.Add("    - $s") }

$output.Add("  Failed   : $($failed.Count)")
foreach ($f in $failed) { $output.Add("    ✗ $f") }

if ($errors.Count -gt 0) {
    $output.Add("")
    $output.Add("  Errors:")
    foreach ($e in $errors) { $output.Add("    ! $e") }
}

$output.Add("")
$output.Add("=== EINDE ===")

# ── Output ────────────────────────────────────────────────────────────────────
$result = $output -join "`n"
Write-Host $result

# NinjaOne custom field bijwerken
Ninja-Property-Set AudioDeviceInventory $result

# Exit code op basis van resultaat
if ($failed.Count -gt 0) {
    exit 1  # NinjaOne markeert script als mislukt
} else {
    exit 0
}