<#
.SYNOPSIS
    Intune Win32-app custom detection script voor de Cowork Windows-vereisten
    (VirtualMachinePlatform + onderliggende HCS-services), los van Claude Desktop zelf.

.DESCRIPTION
    Rapporteert "geïnstalleerd" (exit 0) alleen als zowel het VirtualMachinePlatform-feature als
    de onderliggende HCS-services (vmcompute, HNS, vfpext) aanwezig zijn. VirtualMachinePlatform
    op "Enabled" in DISM betekent niet automatisch dat deze services al bestaan (zie Anthropic's
    eigen Cowork-troubleshooting: "Missing HCS services: HNS, vmcompute, vfpext"), met name vlak
    na het inschakelen van VirtualMachinePlatform maar vóór de vereiste herstart.

    Bewust GEEN check op Status "Running" voor deze services: vmcompute is een trigger-start
    service en hoort dus "Stopped" te zijn zolang er geen Cowork-sessie actief is — dat zou een
    volkomen gezond apparaat als "niet geïnstalleerd" laten zien. Alleen het volledig ontbreken
    van de service (Get-Service vindt hem niet) is een betrouwbaar signaal dat de onderliggende
    Hyper-V-componenten nog niet actief zijn.

    Op een apparaat waar dit nog nooit heeft gestaan geeft de check gewoon een lege/"Disabled"-
    uitkomst terug (geen exception) — dat resulteert in exit 1 ("niet geïnstalleerd"), precies
    zoals bedoeld. Get-WindowsOptionalFeature kan tijdelijk falen als DISM net door iets anders
    bezet wordt — daarom een korte retry, zodat zo'n voorbijgaande DISM-lock niet als "niet
    geïnstalleerd" wordt geïnterpreteerd.

    Gebruik als Intune "Custom detection script" (64-bit script — RunAs32Bit $false, consistent
    met de x64-requirement rule van de app zelf —, geen signature check).

.NOTES
    Geen output/log-bestand: Intune leest alleen exit code + eventuele stdout van dit
    script rechtstreeks uit.
#>

$ErrorActionPreference = "Stop"

function Invoke-WithRetry {
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
    $vmp = Invoke-WithRetry -Action {
        Get-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -ErrorAction Stop
    }
    if ($vmp.State -ne "Enabled") {
        exit 1
    }

    foreach ($svcName in @('vmcompute', 'HNS', 'vfpext')) {
        if (-not (Get-Service -Name $svcName -ErrorAction SilentlyContinue)) {
            exit 1
        }
    }

    Write-Output "VirtualMachinePlatform enabled; HCS services (vmcompute, HNS, vfpext) present."
    exit 0
}
catch {
    exit 1
}
