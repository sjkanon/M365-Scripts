<#
.SYNOPSIS
    Intune Win32-app custom detection script voor Claude Desktop.

.DESCRIPTION
    Rapporteert "geïnstalleerd" (exit 0 + stdout) alleen als de machine-brede
    Claude-provisioning, het VirtualMachinePlatform-feature (vereist voor Cowork) én de
    onderliggende HCS-services (vmcompute, HNS, vfpext) aanwezig zijn. Zo dekt de detectie
    precies het scenario waar Cowork stilzwijgend faalt terwijl Claude zelf wél als
    geïnstalleerd oogt — VirtualMachinePlatform op "Enabled" in DISM betekent niet
    automatisch dat deze services al bestaan (zie Anthropic's eigen Cowork-troubleshooting:
    "Missing HCS services: HNS, vmcompute, vfpext"), met name vlak na het inschakelen van
    VirtualMachinePlatform maar vóór de vereiste herstart.

    Bewust GEEN versie-check: Intune Win32-apps herinstalleren op reeds-toegewezen
    apparaten zodra de content-versie van de app in Intune wijzigt (zie
    Update-IntuneWin32AppPackageFile in Deploy-ClaudeDesktopIntune.ps1), onafhankelijk
    van wat deze detectieregel teruggeeft. Een versie-check hier zou dus alleen extra
    onderhoud betekenen (handmatig ophogen bij elke maandelijkse release) zonder functie.

    Op een apparaat waar Claude nog nooit heeft gestaan geven beide checks gewoon een
    lege/"Disabled"-uitkomst terug (geen exception) — dat resulteert in exit 1 ("niet
    geïnstalleerd"), precies zoals bedoeld, zonder speciale afhandeling nodig. Losstaand
    daarvan: Get-AppxProvisionedPackage/Get-WindowsOptionalFeature kunnen tijdelijk falen
    als DISM net door iets anders bezet wordt (bv. terwijl Install-ClaudeDesktop-Intune.ps1
    op hetzelfde apparaat nog aan het opruimen/provisioneren is) — beide calls krijgen
    daarom een korte retry, zodat zo'n voorbijgaande DISM-lock niet als "niet
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

    $vmp = Invoke-WithRetry -Action {
        Get-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -ErrorAction Stop
    }
    if ($vmp.State -ne "Enabled") {
        exit 1
    }

    # HCS-services die Cowork daadwerkelijk gebruikt — VirtualMachinePlatform "Enabled" in DISM
    # betekent niet automatisch dat deze al bestaan (zie Anthropic's eigen Cowork-troubleshooting:
    # "Missing HCS services: HNS, vmcompute, vfpext"), bv. vlak na het inschakelen van
    # VirtualMachinePlatform maar vóór de vereiste herstart. Bewust geen check op Status
    # "Running": vmcompute is een trigger-start service en hoort dus "Stopped" te zijn zolang er
    # geen Cowork-sessie actief is — dat zou een volkomen gezond apparaat als "niet geïnstalleerd"
    # laten zien. Alleen het volledig ontbreken van de service (Get-Service vindt hem niet) is een
    # betrouwbaar signaal dat de onderliggende Hyper-V-componenten nog niet actief zijn.
    foreach ($svcName in @('vmcompute', 'HNS', 'vfpext')) {
        if (-not (Get-Service -Name $svcName -ErrorAction SilentlyContinue)) {
            exit 1
        }
    }

    Write-Output "Claude Desktop $($provisioned[0].Version) machine-wide provisioned; VirtualMachinePlatform enabled; HCS services present."
    exit 0
}
catch {
    exit 1
}
