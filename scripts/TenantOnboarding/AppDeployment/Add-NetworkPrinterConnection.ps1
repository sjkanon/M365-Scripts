#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Add a network printer connection by IP-based port and driver name.

.DESCRIPTION
    Generalized replacement for a one-off script that installed a specific labeled
    printer via a hardcoded internal IP address. Adds a TCP/IP printer port and a
    printer connection using an already-installed driver (install the driver first —
    via -DriverInstallerPath if you have a silent installer, or ship it as a separate
    Win32 app/driver package).

.PARAMETER PrinterName
    Display name for the printer connection.

.PARAMETER PortAddress
    Printer's IP address or hostname.

.PARAMETER DriverName
    Exact name of an already-installed printer driver (see `Get-PrinterDriver`).

.PARAMETER PortName
    Name for the printer port object. Default: "IP_<PortAddress>".

.PARAMETER Apply
    Actually create the port and printer. Without this switch, the script only
    reports what it would do.

.EXAMPLE
    .\Add-NetworkPrinterConnection.ps1 -PrinterName "Label Printer - Warehouse" `
        -PortAddress "10.0.5.50" -DriverName "Dymo LabelWriter 450 Turbo"

.EXAMPLE
    .\Add-NetworkPrinterConnection.ps1 -PrinterName "Label Printer - Warehouse" `
        -PortAddress "10.0.5.50" -DriverName "Dymo LabelWriter 450 Turbo" -Apply

.NOTES
    Install the printer driver first (e.g. via Install-Win32AppPackage.ps1 pointing
    at the vendor's driver package) — this script only creates the port and the
    printer connection, it does not install drivers.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $PrinterName,

    [Parameter(Mandatory)]
    [string] $PortAddress,

    [Parameter(Mandatory)]
    [string] $DriverName,

    [string] $PortName,
    [switch] $Apply
)

if (-not $PortName) { $PortName = "IP_$PortAddress" }

Write-Host ""
Write-Host "  Add-NetworkPrinterConnection : $PrinterName ($PortAddress)" -ForegroundColor Cyan
Write-Host "  Mode : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

$driver = Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue
if (-not $driver) {
    Write-Host "  [WARN] Driver '$DriverName' is not installed. Install it before adding the printer." -ForegroundColor Yellow
}

if (-not $Apply) {
    Write-Host "  Would create printer port '$PortName' -> $PortAddress" -ForegroundColor Yellow
    Write-Host "  Would add printer '$PrinterName' using driver '$DriverName' on port '$PortName'" -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to perform these actions." -ForegroundColor Yellow
    Write-Host ""
    exit 0
}

if (-not $PSCmdlet.ShouldProcess($PrinterName, "Add network printer connection")) { exit 0 }

if (-not (Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue)) {
    Add-PrinterPort -Name $PortName -PrinterHostAddress $PortAddress -ErrorAction Stop
    Write-Host "  [OK]   Printer port '$PortName' created." -ForegroundColor Green
} else {
    Write-Host "  [SKIP] Printer port '$PortName' already exists." -ForegroundColor DarkGray
}

if (-not (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue)) {
    Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $PortName -ErrorAction Stop
    Write-Host "  [OK]   Printer '$PrinterName' added." -ForegroundColor Green
} else {
    Write-Host "  [SKIP] Printer '$PrinterName' already exists." -ForegroundColor DarkGray
}

Write-Host ""
