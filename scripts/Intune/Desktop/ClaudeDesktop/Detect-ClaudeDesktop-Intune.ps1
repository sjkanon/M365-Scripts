<#
.SYNOPSIS
    Intune Win32-app custom detection script voor Claude Desktop.

.DESCRIPTION
    Rapporteert "geïnstalleerd" (exit 0 + stdout) alleen als zowel de machine-brede
    Claude-provisioning als het VirtualMachinePlatform-feature (vereist voor Cowork)
    aanwezig zijn. Zo dekt de detectie precies het scenario waar Cowork stilzwijgend
    faalt terwijl Claude zelf wél als geïnstalleerd oogt.

    Bewust GEEN versie-check: Intune Win32-apps herinstalleren op reeds-toegewezen
    apparaten zodra de content-versie van de app in Intune wijzigt (zie
    Update-IntuneWin32AppPackageFile in Deploy-ClaudeDesktopIntune.ps1), onafhankelijk
    van wat deze detectieregel teruggeeft. Een versie-check hier zou dus alleen extra
    onderhoud betekenen (handmatig ophogen bij elke maandelijkse release) zonder functie.

    Gebruik als Intune "Custom detection script" (32-bit script, geen signature check).

.NOTES
    Geen output/log-bestand: Intune leest alleen exit code + eventuele stdout van dit
    script rechtstreeks uit.
#>

$ErrorActionPreference = "Stop"

try {
    $provisioned = Get-AppxProvisionedPackage -Online -ErrorAction Stop |
        Where-Object { $_.DisplayName -like "*Claude*" }
    if (-not $provisioned) {
        exit 1
    }

    $vmp = Get-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -ErrorAction Stop
    if ($vmp.State -ne "Enabled") {
        exit 1
    }

    Write-Output "Claude Desktop $($provisioned[0].Version) machine-wide provisioned; VirtualMachinePlatform enabled."
    exit 0
}
catch {
    exit 1
}
