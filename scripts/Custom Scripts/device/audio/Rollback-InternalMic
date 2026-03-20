# ============================================
# Rollback: Re-enable Internal Microphone
# Ticket: #0250981 - Best Next Contact BVBA
# Auteur: Sjoerd Kanon
# Datum: 19/03/2026
#
# Doel: Zet de interne microfoon terug aan als het
#       disable script problemen veroorzaakt heeft.
#
# Uitrol: NinjaOne - Run as SYSTEM
# ============================================

$output  = [System.Collections.Generic.List[string]]::new()
$enabled = [System.Collections.Generic.List[string]]::new()
$failed  = [System.Collections.Generic.List[string]]::new()

$output.Add("=== ROLLBACK: RE-ENABLE INTERNAL MICROPHONE ===")
$output.Add("Hostname : $env:COMPUTERNAME")
$output.Add("Datum    : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')")
$output.Add("")

# Zelfde patronen als disable script
$internalPatterns = @(
    "*Microfoonmatrix*",
    "*Microphone Array*",
    "*Microphone*Realtek*",
    "*Microphone*High Definition Audio*",
    "*Microfoon*Realtek*",
    "*Microfoon*Synaptics*",
    "*Microphone (2-*"
)

$headsetSafeList = @(
    "EPOS", "Jabra", "Plantronics", "Yealink", "Poly",
    "Logitech", "Headset", "hoofdtelefoon", "oortelefoon",
    "Blackwire", "Voyager", "Sennheiser", "Astro", "SteelSeries"
)

# Alle microfoons ophalen inclusief disabled (Unknown/Error status)
$allMics = Get-PnpDevice -Class AudioEndpoint -ErrorAction SilentlyContinue |
    Where-Object { $_.InstanceId -like "*0.0.1*" }

$output.Add("--- GEVONDEN MICROFOONS (voor rollback) ---")
foreach ($mic in $allMics) {
    $output.Add("  [$($mic.Status)] $($mic.FriendlyName)")
}
$output.Add("")
$output.Add("--- VERWERKING ---")

foreach ($mic in $allMics) {
    $name = $mic.FriendlyName

    # Headsets nooit aanraken
    $isHeadset = $false
    foreach ($safe in $headsetSafeList) {
        if ($name -like "*$safe*") { $isHeadset = $true; break }
    }
    if ($isHeadset) {
        $output.Add("  SKIP (headset): $name")
        continue
    }

    # Intern patroon check
    $isInternal = $false
    foreach ($pattern in $internalPatterns) {
        if ($name -like $pattern) { $isInternal = $true; break }
    }
    if (-not $isInternal) {
        $output.Add("  SKIP (geen intern patroon): $name")
        continue
    }

    # Re-enablen
    try {
        Enable-PnpDevice -InstanceId $mic.InstanceId -Confirm:$false -ErrorAction Stop
        $output.Add("  ENABLED: $name")
        $enabled.Add($name)
    } catch {
        $output.Add("  FAILED: $name | $_")
        $failed.Add($name)
    }
}

$output.Add("")
$output.Add("--- SAMENVATTING ---")
$output.Add("  Re-enabled : $($enabled.Count)")
foreach ($e in $enabled) { $output.Add("    ✓ $e") }
$output.Add("  Failed     : $($failed.Count)")
foreach ($f in $failed) { $output.Add("    ✗ $f") }
$output.Add("")
$output.Add("=== EINDE ROLLBACK ===")

$result = $output -join "`n"
Write-Host $result
Ninja-Property-Set AudioDeviceInventory $result

if ($failed.Count -gt 0) { exit 1 } else { exit 0 }