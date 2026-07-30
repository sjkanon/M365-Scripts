<#
.SYNOPSIS
    Intune Proactive Remediation — detectiehelft — Cowork Windows-vereisten.

.DESCRIPTION
    Rapporteert "compliant" (exit 0) alleen als zowel het VirtualMachinePlatform-feature als de
    onderliggende HCS-services (vmcompute, HNS, vfpext) aanwezig zijn. VirtualMachinePlatform op
    "Enabled" in DISM betekent niet automatisch dat deze services al bestaan (zie Anthropic's
    eigen Cowork-troubleshooting: "Missing HCS services: HNS, vmcompute, vfpext"), met name vlak
    na het inschakelen van VirtualMachinePlatform maar vóór de vereiste herstart — in dat geval
    faalt deze detectie bewust opnieuw, en de gekoppelde Remediate-CoworkPrerequisites.ps1 stuurt
    de gebruiker (via de dagelijkse remediation-cyclus) net zolang een herinnering totdat er
    daadwerkelijk herstart is.

    Bewust GEEN check op Status "Running" voor deze services: vmcompute is een trigger-start
    service en hoort dus "Stopped" te zijn zolang er geen Cowork-sessie actief is — dat zou een
    volkomen gezond apparaat als "niet compliant" laten zien. Alleen het volledig ontbreken van
    de service (Get-Service vindt hem niet) is een betrouwbaar signaal dat de onderliggende
    Hyper-V-componenten nog niet actief zijn.

    Get-WindowsOptionalFeature kan tijdelijk falen als DISM net door iets anders bezet wordt
    (bv. Windows Update, of een andere Proactive Remediation/Win32-app-install tegelijk) — daarom
    een korte retry, zodat zo'n voorbijgaande DISM-lock niet als "niet compliant" wordt
    geïnterpreteerd.

    Gebruik als Intune "Scripts and remediations > Remediations" detectiescript — zie readme.md
    in deze map voor de volledige portal-setup (koppeling met Remediate-CoworkPrerequisites.ps1,
    schema, toewijzing).

.NOTES
    Geen apart logbestand: Intune leest bij een remediation alleen de exit code + stdout van dit
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
