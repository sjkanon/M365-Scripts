#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Activate Windows or manage product key and KMS settings.

.DESCRIPTION
    Unified interface for Windows activation tasks using slmgr.vbs and WMI:
      - Show current activation status and license details
      - Install a product key (retail or KMS generic key)
      - Configure a KMS activation server (for volume licensing)
      - Trigger online or KMS-based activation
      - Remove the installed product key (before reimage / license transfer)
      - ReArm the grace period timer (limited to ~3-5 uses per install)

    All slmgr operations run via cscript //NoLogo to return text output
    instead of GUI dialog boxes.

.PARAMETER Status
    Show current activation status, license type, and expiration/grace info.
    Always runs, even when combined with other parameters.

.PARAMETER ProductKey
    Install a product key (format: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX).
    Use a retail key for single-device activation or a KMS generic key for
    volume-licensed environments.

.PARAMETER KmsServer
    Set the KMS host address for volume activation.
    Used in corporate environments that run an internal KMS server.
    Combine with -Activate to configure and activate in one step.

.PARAMETER KmsPort
    TCP port of the KMS server. Default: 1688.
    Only used when -KmsServer is specified.

.PARAMETER Activate
    Trigger Windows activation against Microsoft servers or the configured
    KMS server. If a KMS server is not set, activation goes online.

.PARAMETER RemoveKey
    Remove (uninstall) the product key from this machine.
    The key is scrubbed from the registry. Use before reimaging or
    transferring a retail license to another device.
    Requires confirmation unless -Force is specified.

.PARAMETER ReArm
    Reset the Windows activation grace-period counter.
    Each Windows installation allows a limited number of rearms (usually 3-5).
    Use 'slmgr /dlv' output to check remaining rearm count.
    Requires a reboot to take effect.

.PARAMETER Force
    Skip confirmation prompts (e.g. for -RemoveKey and -ReArm).

.EXAMPLE
    # Show current activation status
    .\Invoke-WindowsActivation.ps1 -Status

.EXAMPLE
    # Install a retail product key and activate online
    .\Invoke-WindowsActivation.ps1 -ProductKey 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX' -Activate

.EXAMPLE
    # Point to a corporate KMS server and activate
    .\Invoke-WindowsActivation.ps1 -KmsServer 'kms.company.local' -Activate

.EXAMPLE
    # KMS server on a non-standard port
    .\Invoke-WindowsActivation.ps1 -KmsServer 'kms.company.local' -KmsPort 2500 -Activate

.EXAMPLE
    # Remove the product key before reimaging
    .\Invoke-WindowsActivation.ps1 -RemoveKey -Force

.EXAMPLE
    # Full flow: install key, set KMS server, activate, show status
    .\Invoke-WindowsActivation.ps1 -ProductKey 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX' -KmsServer 'kms.company.local' -Activate -Status
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [switch] $Status,
    [string] $ProductKey,
    [string] $KmsServer,
    [int]    $KmsPort = 1688,
    [switch] $Activate,
    [switch] $RemoveKey,
    [switch] $ReArm,
    [switch] $Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$slmgr = "$env:SystemRoot\System32\slmgr.vbs"

# ── Helpers ────────────────────────────────────────────────────────────────────

function Invoke-Slmgr {
    param([string[]]$Arguments)
    $output = & cscript.exe //NoLogo $slmgr @Arguments 2>&1
    return ($output -join "`n").Trim()
}

function Get-ActivationStatus {
    # LicenseStatus values returned by WMI
    $statusMap = @{
        0 = 'Unlicensed'
        1 = 'Licensed'
        2 = 'Out-of-Box Grace Period'
        3 = 'Out-of-Tolerance Grace Period'
        4 = 'Non-Genuine Grace Period'
        5 = 'Notification (not activated)'
        6 = 'Extended Grace Period'
    }

    # Filter on the Windows activation application ID
    $product = Get-CimInstance -ClassName SoftwareLicensingProduct `
        -Filter "ApplicationId='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL" `
        -ErrorAction SilentlyContinue |
        Select-Object -First 1

    if (-not $product) {
        return [PSCustomObject]@{
            Name          = 'Unknown'
            Status        = 'No license found'
            StatusCode    = -1
            Activated     = $false
            LicenseFamily = ''
            GraceRemaining = ''
            KmsServer     = ''
        }
    }

    $graceMin = $product.GracePeriodRemaining
    $graceStr = if ($graceMin -gt 0) {
        $days  = [math]::Floor($graceMin / 1440)
        $hours = [math]::Floor(($graceMin % 1440) / 60)
        "$days day(s), $hours hour(s) remaining"
    } else { '' }

    [PSCustomObject]@{
        Name           = $product.Name
        Status         = $statusMap[[int]$product.LicenseStatus]
        StatusCode     = [int]$product.LicenseStatus
        Activated      = ($product.LicenseStatus -eq 1)
        LicenseFamily  = $product.ProductKeyChannel
        GraceRemaining = $graceStr
        KmsServer      = if ($product.DiscoveredKeyManagementServiceMachineName) {
                             "$($product.DiscoveredKeyManagementServiceMachineName):$($product.DiscoveredKeyManagementServicePort)"
                         } else { '' }
    }
}

function Write-Header {
    param([string]$Text)
    Write-Host ''
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host "  $('─' * $Text.Length)" -ForegroundColor DarkGray
}

function Write-Ok   { param([string]$Msg) Write-Host "  [OK]   $Msg" -ForegroundColor Green   }
function Write-Info { param([string]$Msg) Write-Host "  [INFO] $Msg" -ForegroundColor DarkGray }
function Write-Warn { param([string]$Msg) Write-Host "  [WARN] $Msg" -ForegroundColor Yellow  }
function Write-Fail { param([string]$Msg) Write-Host "  [FAIL] $Msg" -ForegroundColor Red     }

# ── Banner ─────────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host '   Invoke-WindowsActivation' -ForegroundColor Cyan
Write-Host '  ================================================' -ForegroundColor Cyan

# Require at least one action
if (-not ($Status -or $ProductKey -or $KmsServer -or $Activate -or $RemoveKey -or $ReArm)) {
    Write-Host ''
    Write-Warn 'No action specified. Use -Status to check, or -Help for usage.'
    Write-Host ''
    Write-Host '  Examples:' -ForegroundColor DarkGray
    Write-Host '    .\Invoke-WindowsActivation.ps1 -Status' -ForegroundColor DarkGray
    Write-Host '    .\Invoke-WindowsActivation.ps1 -ProductKey XXXXX-XXXXX-XXXXX-XXXXX-XXXXX -Activate' -ForegroundColor DarkGray
    Write-Host '    .\Invoke-WindowsActivation.ps1 -KmsServer kms.company.local -Activate' -ForegroundColor DarkGray
    Write-Host ''
    exit 0
}

# ── 1. Install product key ─────────────────────────────────────────────────────
if ($ProductKey) {
    Write-Header 'Installing product key'

    # Basic format validation
    if ($ProductKey -notmatch '^[A-Z0-9]{5}(-[A-Z0-9]{5}){4}$') {
        Write-Fail "Invalid product key format. Expected: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"
        exit 1
    }

    Write-Info "Installing key: $($ProductKey.Substring(0,5))-XXXXX-XXXXX-XXXXX-$($ProductKey.Substring(24,5))"
    try {
        $result = Invoke-Slmgr '/ipk', $ProductKey
        Write-Ok $result
    } catch {
        Write-Fail "Failed to install product key: $_"
        exit 1
    }
}

# ── 2. Set KMS server ─────────────────────────────────────────────────────────
if ($KmsServer) {
    Write-Header 'Configuring KMS server'
    $kmsTarget = "${KmsServer}:${KmsPort}"
    Write-Info "Setting KMS server to: $kmsTarget"
    try {
        $result = Invoke-Slmgr '/skms', $kmsTarget
        Write-Ok $result
    } catch {
        Write-Fail "Failed to set KMS server: $_"
        exit 1
    }
}

# ── 3. Activate ───────────────────────────────────────────────────────────────
if ($Activate) {
    Write-Header 'Activating Windows'
    Write-Info 'Contacting activation server...'
    try {
        $result = Invoke-Slmgr '/ato'
        Write-Ok $result
    } catch {
        Write-Fail "Activation failed: $_"
        Write-Info 'Check network connectivity and verify the product key or KMS server.'
    }
}

# ── 4. Remove product key ─────────────────────────────────────────────────────
if ($RemoveKey) {
    Write-Header 'Removing product key'

    if (-not $Force) {
        Write-Warn 'This will remove the installed product key from this machine.'
        $confirm = Read-Host '  Continue? (yes/no)'
        if ($confirm -ne 'yes') {
            Write-Info 'Cancelled.'
        } else {
            $RemoveKey = $true
        }
    }

    if ($Force -or $confirm -eq 'yes') {
        try {
            $result = Invoke-Slmgr '/upk'
            Write-Ok $result
            Write-Info 'Key removed. The system will enter an unlicensed state.'
        } catch {
            Write-Fail "Failed to remove product key: $_"
            exit 1
        }
    }
}

# ── 5. ReArm ──────────────────────────────────────────────────────────────────
if ($ReArm) {
    Write-Header 'ReArm — resetting grace period'

    if (-not $Force) {
        Write-Warn 'ReArm resets the grace period timer. This operation is limited (usually 3-5 times per install).'
        $confirm = Read-Host '  Continue? (yes/no)'
        if ($confirm -ne 'yes') {
            Write-Info 'Cancelled.'
            $ReArm = $false
        }
    }

    if ($ReArm) {
        try {
            $result = Invoke-Slmgr '/rearm'
            Write-Ok $result
            Write-Warn 'A reboot is required for the rearm to take effect.'
        } catch {
            Write-Fail "ReArm failed: $_"
            exit 1
        }
    }
}

# ── 6. Status ─────────────────────────────────────────────────────────────────
if ($Status) {
    Write-Header 'Activation status'

    $info = Get-ActivationStatus

    $statusColor = if ($info.Activated) { 'Green' } elseif ($info.StatusCode -in 2,6) { 'Yellow' } else { 'Red' }

    Write-Host ("  {0,-20} {1}" -f 'Product:',       $info.Name)          -ForegroundColor White
    Write-Host ("  {0,-20} {1}" -f 'Status:',        $info.Status)        -ForegroundColor $statusColor
    Write-Host ("  {0,-20} {1}" -f 'License type:',  $info.LicenseFamily) -ForegroundColor White

    if ($info.GraceRemaining) {
        Write-Host ("  {0,-20} {1}" -f 'Grace period:', $info.GraceRemaining) -ForegroundColor Yellow
    }
    if ($info.KmsServer) {
        Write-Host ("  {0,-20} {1}" -f 'KMS server:', $info.KmsServer) -ForegroundColor White
    }

    # Detailed output via slmgr
    Write-Host ''
    Write-Info 'Full license details (slmgr /dli):'
    Write-Host ''
    Invoke-Slmgr '/dli' | ForEach-Object {
        Write-Host "    $_" -ForegroundColor DarkGray
    }
}

Write-Host ''
