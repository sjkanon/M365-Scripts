<#
.SYNOPSIS
    Intune Win32-app custom detection script voor Claude Desktop.

.DESCRIPTION
    Rapporteert "geïnstalleerd" (exit 0 + stdout) als de machine-brede Claude-provisioning
    aanwezig is. De Windows-kant van Cowork (VirtualMachinePlatform + HCS-services) heeft een
    eigen, onafhankelijke detectie in ../CoworkPrerequisites/Detect-CoworkPrerequisites.ps1
    — bewust niet hier meegenomen, zodat een Cowork-probleem niet door elkaar loopt met een
    Claude Desktop-installatieprobleem: beide apps krijgen hun eigen, apart zichtbare status in
    Intune.

    Bewust GEEN versie-check: Intune Win32-apps herinstalleren op reeds-toegewezen
    apparaten zodra de content-versie van de app in Intune wijzigt (zie
    Update-IntuneWin32AppPackageFile in Deploy-ClaudeDesktopIntune.ps1), onafhankelijk
    van wat deze detectieregel teruggeeft. Een versie-check hier zou dus alleen extra
    onderhoud betekenen (handmatig ophogen bij elke maandelijkse release) zonder functie.

    Op een apparaat waar Claude nog nooit heeft gestaan geeft de check gewoon een lege
    uitkomst terug (geen exception) — dat resulteert in exit 1 ("niet geïnstalleerd"), precies
    zoals bedoeld, zonder speciale afhandeling nodig. Losstaand daarvan: Get-AppxProvisionedPackage
    kan tijdelijk falen als DISM net door iets anders bezet wordt (bv. terwijl
    Install-ClaudeDesktop-Intune.ps1 op hetzelfde apparaat nog aan het opruimen/provisioneren is)
    — dat krijgt daarom een korte retry, zodat zo'n voorbijgaande DISM-lock niet als "niet
    geïnstalleerd" wordt geïnterpreteerd.

    Gebruik als Intune "Custom detection script" (64-bit script — RunAs32Bit $false in
    Deploy-ClaudeDesktopIntune.ps1, consistent met de x64-requirement rule van de app zelf
    —, geen signature check).

.NOTES
    Geen output/log-bestand: Intune leest alleen exit code + eventuele stdout van dit
    script rechtstreeks uit.
#>

$ErrorActionPreference = "Stop"

function Invoke-WithRetry {
    # Alleen bedoeld om voorbijgaande DISM/Appx-storingen (bv. "busy") op te vangen — een
    # normale lege/negatieve uitkomst (niet geïnstalleerd) is geen exception en wordt dus
    # nooit geretried; dit vertraagt het "echt niet geïnstalleerd"-pad niet.
    param([scriptblock]$Action, [int]$MaxAttempts = 3, [int]$DelaySeconds = 5)
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return & $Action
        } catch {
            if ($attempt -eq $MaxAttempts) { throw }
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

try {
    $provisioned = Invoke-WithRetry -Action {
        Get-AppxProvisionedPackage -Online -ErrorAction Stop | Where-Object { $_.DisplayName -like "*Claude*" }
    }
    if (-not $provisioned) {
        exit 1
    }

    Write-Output "Claude Desktop $($provisioned[0].Version) machine-wide provisioned."
    exit 0
}
catch {
    exit 1
}
