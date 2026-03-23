#Requires -Version 5.1
# Cross-platform: Windows (PS 5.1+), macOS and Linux (PS 7+)
<#
.SYNOPSIS
    Interactive launcher menu for all M365-Scripts tools.

.DESCRIPTION
    Press a number key or F-key to launch a script directly — no Enter needed.
    Required parameters are prompted before the script runs.
#>

$ROOT = $PSScriptRoot

# ── Menu items ────────────────────────────────────────────────────────────────
# Each entry: Key, FKey (1-based), Category, Label, ScriptPath, Action (optional scriptblock for params)
$menu = @(
    [PSCustomObject]@{
        Key      = '1'
        FKey     = [ConsoleKey]::F1
        Category = 'Network'
        Label    = 'Test-Ports          — check open TCP ports on any host'
        Script   = "$ROOT\scripts\Network\Test-Ports.ps1"
        Params   = {
            $target  = Read-Host "  Target (IP or hostname)"
            $ports   = Read-Host "  Ports  (e.g. 80,443 or 1294:1494 or 80,1294:1494)"
            $timeout = Read-Host "  Timeout ms [500]"
            if (-not $timeout) { $timeout = 500 }
            $show    = Read-Host "  Show closed ports? [y/N]"
            $args = @{ Target = $target; Ports = $ports; TimeoutMs = [int]$timeout }
            if ($show -match '^[Yy]') { $args['ShowClosed'] = $true }
            return $args
        }
    }
    [PSCustomObject]@{
        Key      = '2'
        FKey     = [ConsoleKey]::F2
        Category = 'Exchange'
        Label    = 'Migrate-Calendar    — migrate M365 group calendar to room mailbox'
        Script   = "$ROOT\scripts\Exchange\Migrate-Calendar.ps1"
        Params   = {
            $tenantId  = Read-Host "  TenantId"
            $adminUPN  = Read-Host "  AdminUPN"
            $groupMail = Read-Host "  Source group mail"
            return @{ TenantId = $tenantId; AdminUPN = $adminUPN; SourceGroupMail = $groupMail }
        }
    }
    [PSCustomObject]@{
        Key      = '3'
        FKey     = [ConsoleKey]::F3
        Category = 'Exchange'
        Label    = 'Set-Calendar-rights — grant calendar permissions to a user'
        Script   = "$ROOT\scripts\Exchange\Set-Calendar-rights.ps1"
        Params   = {
            $user    = Read-Host "  User (without domain)"
            $mailbox = Read-Host "  Target mailbox (without domain)"
            $rights  = Read-Host "  Access rights (e.g. Reviewer, Editor, Owner)"
            return @{ User = $user; TargetMailbox = $mailbox; AccessRights = $rights }
        }
    }
    [PSCustomObject]@{
        Key      = '4'
        FKey     = [ConsoleKey]::F4
        Category = 'Testing'
        Label    = 'Test-SMTP           — one-time SMTP connectivity test'
        Script   = "$ROOT\scripts\Testing Scripts\SMTP\testsmtp.ps1"
        Params   = {
            $server = Read-Host "  SMTP server [smtp.office365.com]"
            $from   = Read-Host "  From address"
            $to     = Read-Host "  To address"
            $args = @{}
            if ($server) { $args['SmtpServer'] = $server }
            if ($from)   { $args['From']       = $from   }
            if ($to)     { $args['To']         = $to     }
            return $args
        }
    }
    [PSCustomObject]@{
        Key      = '5'
        FKey     = [ConsoleKey]::F5
        Category = 'Testing'
        Label    = 'Test-SMTP (5 min)   — recurring SMTP test every 5 minutes'
        Script   = "$ROOT\scripts\Testing Scripts\SMTP\testsmtp_5min.ps1"
        Params   = {
            $server = Read-Host "  SMTP server [smtp.office365.com]"
            $from   = Read-Host "  From address"
            $to     = Read-Host "  To address"
            $args = @{}
            if ($server) { $args['SmtpServer'] = $server }
            if ($from)   { $args['From']       = $from   }
            if ($to)     { $args['To']         = $to     }
            return $args
        }
    }
    [PSCustomObject]@{
        Key      = '6'
        FKey     = [ConsoleKey]::F6
        Category = 'Device'
        Label    = 'Restart-Time-Sync   — force Windows time service sync'
        Script   = "$ROOT\scripts\Custom Scripts\device\Time sync\Restart-Time-Sync.ps1"
        Params   = { return @{} }
    }
    [PSCustomObject]@{
        Key      = '7'
        FKey     = [ConsoleKey]::F7
        Category = 'Device'
        Label    = 'Detect-AudioDevices — list connected audio devices'
        Script   = "$ROOT\scripts\Custom Scripts\device\audio\detect-audiodevices.ps1"
        Params   = { return @{} }
    }
    [PSCustomObject]@{
        Key      = '8'
        FKey     = [ConsoleKey]::F8
        Category = 'Device'
        Label    = 'Disable-InternalMic — disable internal microphone via policy'
        Script   = "$ROOT\scripts\Custom Scripts\device\audio\Disable-internalmic.ps1"
        Params   = { return @{} }
    }
    [PSCustomObject]@{
        Key      = '9'
        FKey     = [ConsoleKey]::F9
        Category = 'Startup'
        Label    = 'Install-Modules     — bootstrap: install all required PS modules'
        Script   = "$ROOT\scripts\Startup\Install-Modules.ps1"
        Params   = { return @{} }
    }
    [PSCustomObject]@{
        Key      = 'A'
        FKey     = [ConsoleKey]::F10
        Category = 'Reporting'
        Label    = 'Licensing-Report    — generate monthly Pax8 + Ingram Excel report'
        Script   = "$ROOT\scripts\Reporting\Licensing\genereer_rapport.ps1"
        Params   = { return @{} }
    }
)

# ── Display menu ──────────────────────────────────────────────────────────────
function Show-Menu {
    Clear-Host
    Write-Host ""
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "   M365-Scripts — Launcher" -ForegroundColor Cyan
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host ""

    $currentCategory = ""
    foreach ($item in $menu) {
        if ($item.Category -ne $currentCategory) {
            if ($currentCategory -ne "") { Write-Host "" }
            Write-Host "  $($item.Category.ToUpper())" -ForegroundColor DarkGray
            $currentCategory = $item.Category
        }
        $keyLabel = "[$($item.Key)] / [F$(([array]::IndexOf($menu, $item)) + 1)]"
        Write-Host ("  {0,-22} {1}" -f $keyLabel, $item.Label)
    }

    Write-Host ""
    Write-Host "  [0] / [Esc]  Exit"
    Write-Host ""
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Press a key to launch:" -ForegroundColor Yellow
}

# ── Run a script ──────────────────────────────────────────────────────────────
function Invoke-MenuItem {
    param([PSCustomObject] $item)

    if (-not (Test-Path $item.Script)) {
        Write-Host ""
        Write-Host "  [ERROR] Script not found: $($item.Script)" -ForegroundColor Red
        Write-Host "  Press any key to return..."
        $null = [Console]::ReadKey($true)
        return
    }

    Clear-Host
    Write-Host ""
    Write-Host "  ── $($item.Category) / $($item.Label.Split('—')[0].Trim()) ──" -ForegroundColor Cyan
    Write-Host ""

    # Collect parameters
    $params = & $item.Params

    Write-Host ""
    Write-Host "  Running..." -ForegroundColor DarkGray
    Write-Host ""

    if ($params.Count -gt 0) {
        & $item.Script @params
    } else {
        & $item.Script
    }

    Write-Host ""
    Write-Host "  Press any key to return to the menu..." -ForegroundColor DarkGray
    $null = [Console]::ReadKey($true)
}

# ── Main loop ─────────────────────────────────────────────────────────────────
while ($true) {
    Show-Menu

    $keyInfo = [Console]::ReadKey($true)
    $char    = $keyInfo.Key

    # Exit on 0 or Escape
    if ($keyInfo.KeyChar -eq '0' -or $char -eq [ConsoleKey]::Escape) {
        Clear-Host
        break
    }

    # Match by character key (1-9, A) or F-key (F1-F10)
    $selected = $menu | Where-Object {
        $_.Key -eq $keyInfo.KeyChar.ToString().ToUpper() -or $_.FKey -eq $char
    } | Select-Object -First 1

    if ($selected) {
        Invoke-MenuItem -item $selected
    }
}
