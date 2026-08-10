<#
.SYNOPSIS
    Intune Win32-app custom detection script voor de C:-schijf cleanup.

.DESCRIPTION
    Rapporteert "geïnstalleerd" (exit 0) zolang de laatst geregistreerde cleanup-run
    (HKLM:\SOFTWARE\DiskCleanupDeploy\LastRunUtc, gezet door Invoke-DiskCleanupIntune.ps1)
    niet ouder is dan $MaxAgeDays. Zodra die verlopen is, meldt dit script "niet
    geïnstalleerd" (exit 1) — Intune rolt de Win32-app dan opnieuw uit op het apparaat, en
    de cleanup draait automatisch elke $MaxAgeDays dagen opnieuw, zonder dat de content-versie
    van de app handmatig hoeft te worden opgehoogd.

    Gebruik als Intune "Custom detection script" (32-bit script, geen signature check).

.NOTES
    Intune geeft geen parameters mee aan detection scripts — pas $MaxAgeDays hieronder aan
    vóór het packagen als je een andere cyclus dan 30 dagen wilt.
    Geen output-/logbestand: Intune leest alleen exit code + eventuele stdout rechtstreeks.
#>

$MaxAgeDays = 30

$ErrorActionPreference = "Stop"

try {
    $value = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\DiskCleanupDeploy" -Name "LastRunUtc" -ErrorAction Stop
    $lastRunUtc = [datetime]::Parse($value, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
    $ageDays = ((Get-Date).ToUniversalTime() - $lastRunUtc).TotalDays

    if ($ageDays -le $MaxAgeDays) {
        Write-Output ("Laatste disk cleanup: {0:yyyy-MM-dd HH:mm} UTC ({1:N1} dag(en) geleden, binnen {2} dagen)." -f $lastRunUtc, $ageDays, $MaxAgeDays)
        exit 0
    }

    exit 1
}
catch {
    exit 1
}
