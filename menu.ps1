#Requires -Version 5.1
# Cross-platform: Windows (PS 5.1+), macOS and Linux (PS 7+)
<#
.SYNOPSIS
    Interactive launcher menu for all M365-Scripts tools.

.DESCRIPTION
    Press a number key or letter key to launch a script or M365 management function.
    M365 functions (B–E) lazy-load functies.ps1 on first use — Graph authentication
    is only triggered when you first select an M365 option.
#>

$ROOT = $PSScriptRoot
$script:FunctiesLoaded = $false

# ── Fallback: ask for UPN if not set by load.ps1 ─────────────────────────────
if (-not $global:upn) {
    Write-Host ""
    $global:upn = Read-Host "  Admin UPN"
}

# ── Lazy-load functies.ps1 ────────────────────────────────────────────────────
function Import-Functies {
    if ($script:FunctiesLoaded) { return $true }

    $path = Join-Path $ROOT 'scripts\Startup\functies.ps1'
    if (-not (Test-Path $path)) {
        Write-Host ""
        Write-Host "  [ERROR] functies.ps1 not found: $path" -ForegroundColor Red
        Write-Host "  Press any key to return..."
        $null = [Console]::ReadKey($true)
        return $false
    }

    Write-Host "  Loading M365 functions and connecting to Graph..." -ForegroundColor DarkGray

    try {
        . $path
        $script:FunctiesLoaded = $true
        return $true
    } catch {
        Write-Host "  [ERROR] Failed to load functies.ps1: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Press any key to return..."
        $null = [Console]::ReadKey($true)
        return $false
    }
}

# ── Run a submenu ─────────────────────────────────────────────────────────────
function Invoke-Submenu {
    param(
        [string] $Title,
        [array]  $Items
    )

    while ($true) {
        Clear-Host
        Write-Host ""
        Write-Host "  ================================================" -ForegroundColor Cyan
        Write-Host "   $Title" -ForegroundColor Cyan
        Write-Host "  ================================================" -ForegroundColor Cyan
        Write-Host ""

        foreach ($item in $Items) {
            Write-Host ("  [{0}] {1}" -f $item.Key, $item.Label)
        }

        Write-Host ""
        Write-Host "  [0] / [Esc]  Back"
        Write-Host ""
        Write-Host "  ================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  Press a key:" -ForegroundColor Yellow

        $keyInfo = [Console]::ReadKey($true)
        if ($keyInfo.KeyChar -eq '0' -or $keyInfo.Key -eq [ConsoleKey]::Escape) { return }

        $selected = $Items | Where-Object { $_.Key -eq $keyInfo.KeyChar.ToString().ToUpper() } | Select-Object -First 1
        if ($selected) {
            Clear-Host
            Write-Host ""
            Write-Host "  ── $($selected.Label.Split('—')[0].Trim()) ──" -ForegroundColor Cyan
            Write-Host ""
            try {
                & $selected.Action
            } catch {
                Write-Host ""
                Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
            }
            Write-Host ""
            Write-Host "  Press any key to return..." -ForegroundColor DarkGray
            $null = [Console]::ReadKey($true)
        }
    }
}

# ── M365 submenus ─────────────────────────────────────────────────────────────

$ExchangeSubmenu = @(
    @{ Key='1'; Label='Add-SharedMailboxAccess   — grant full access + SendAs'; Action={
        $user    = Read-Host "  User (UPN or alias)"
        $mailbox = Read-Host "  Mailbox (UPN or alias)"
        $auto    = Read-Host "  AutoMapping? [y/N]"
        $p = @{ User=$user; Mailbox=$mailbox }
        if ($auto -match '^[Yy]') { $p['AutoMapping'] = $true }
        Add-SharedMailboxAccess @p
    }}
    @{ Key='2'; Label='Set-MailboxLocale         — set language + timezone for all mailboxes'; Action={
        $lang = Read-Host "  Language LCID [1043=NL  2060=FR  1033=EN]"
        $tz   = Read-Host "  TimeZone [W. Europe Standard Time]"
        $p = @{}
        if ($lang) { $p['Language'] = [int]$lang }
        if ($tz)   { $p['TimeZone'] = $tz }
        Set-MailboxLocale @p
    }}
    @{ Key='3'; Label='Add-MailboxAlias          — add email alias to a mailbox'; Action={
        Add-MailboxAlias
    }}
    @{ Key='4'; Label='Get-MailboxAliases        — list all mailbox SMTP addresses'; Action={
        Get-MailboxAliases | Format-Table -AutoSize
    }}
    @{ Key='5'; Label='Export-DistributionGroups — export all DGs to CSV'; Action={
        Export-DistributionGroups
    }}
    @{ Key='6'; Label='Set-AutoReply             — configure out-of-office reply'; Action={
        Set-AutoReply
    }}
    @{ Key='7'; Label='Enable-CopyOfSentItems    — enable sent-item copy for all mailboxes'; Action={
        Enable-CopyOfSentItems
    }}
)

$EntraSubmenu = @(
    @{ Key='1'; Label='Get-TenantAdmins      — list Global Administrators'; Action={
        Get-TenantAdmins | Format-Table -AutoSize
    }}
    @{ Key='2'; Label='Add-TenantDomain      — add and verify a new domain'; Action={
        Add-TenantDomain
    }}
    @{ Key='3'; Label='Get-TenantLicenses    — show license assignment counts'; Action={
        Get-TenantLicenses | Format-Table -AutoSize
    }}
    @{ Key='4'; Label='Get-TenantUsers       — list all users'; Action={
        Get-TenantUsers | Format-Table -AutoSize
    }}
    @{ Key='5'; Label='Add-TenantAdmin       — grant Global Admin to a user'; Action={
        Add-TenantAdmin
    }}
    @{ Key='6'; Label='Reset-UserPassword    — reset a user password'; Action={
        Reset-UserPassword
    }}
    @{ Key='7'; Label='Export-SignInLogs     — export sign-in audit logs to CSV'; Action={
        $days = Read-Host "  Days back [30]"
        if ($days) { Export-SignInLogs -Days ([int]$days) } else { Export-SignInLogs }
    }}
    @{ Key='8'; Label='Get-EntraApplication  — find an Enterprise Application'; Action={
        Get-EntraApplication | Format-List
    }}
)

$MspSubmenu = @(
    @{ Key='1'; Label='New-MspAdmin              — create MSP admin account in tenant'; Action={
        New-MspAdmin
    }}
    @{ Key='2'; Label='Set-MspAdminAsGroupOwner  — set MSP admin as group owner'; Action={
        Set-MspAdminAsGroupOwner
    }}
    @{ Key='3'; Label='Reset-MspAdminPassword    — reset MSP admin password'; Action={
        Reset-MspAdminPassword
    }}
)

# ── Menu items ────────────────────────────────────────────────────────────────
# Script items  : Key, FKey, Category, Label, Script, Params
# Action items  : Key, FKey, Category, Label, Action, [NoWait]
#   NoWait=$true  — item manages its own UI loop (submenus); skip "press any key"
$menu = @(
    [PSCustomObject]@{ Key='1'; FKey=[ConsoleKey]::F1; Category='Network'
        Label='Test-Ports          — check open TCP ports on any host'
        Script="$ROOT\scripts\Network\Test-Ports.ps1"
        Params={
            $target  = Read-Host "  Target (IP or hostname)"
            $ports   = Read-Host "  Ports  (e.g. 80,443 or 1294:1494 or 80,1294:1494)"
            $timeout = Read-Host "  Timeout ms [500]"
            if (-not $timeout) { $timeout = 500 }
            $show    = Read-Host "  Show closed ports? [y/N]"
            $a = @{ Target=$target; Ports=$ports; TimeoutMs=[int]$timeout }
            if ($show -match '^[Yy]') { $a['ShowClosed']=$true }
            return $a
        }
    }
    [PSCustomObject]@{ Key='2'; FKey=[ConsoleKey]::F2; Category='Exchange'
        Label='Migrate-Calendar    — migrate M365 group calendar to room mailbox'
        Script="$ROOT\scripts\Exchange\Migrate-Calendar.ps1"
        Params={
            $tenantId  = Read-Host "  TenantId"
            $adminUPN  = Read-Host "  AdminUPN"
            $groupMail = Read-Host "  Source group mail"
            return @{ TenantId=$tenantId; AdminUPN=$adminUPN; SourceGroupMail=$groupMail }
        }
    }
    [PSCustomObject]@{ Key='3'; FKey=[ConsoleKey]::F3; Category='Exchange'
        Label='Set-Calendar-rights — grant calendar permissions to a user'
        Script="$ROOT\scripts\Exchange\Set-Calendar-rights.ps1"
        Params={
            $user    = Read-Host "  User (without domain)"
            $mailbox = Read-Host "  Target mailbox (without domain)"
            $rights  = Read-Host "  Access rights (e.g. Reviewer, Editor, Owner)"
            return @{ User=$user; TargetMailbox=$mailbox; AccessRights=$rights }
        }
    }
    [PSCustomObject]@{ Key='4'; FKey=[ConsoleKey]::F4; Category='Testing'
        Label='Test-SMTP           — one-time SMTP connectivity test'
        Script="$ROOT\scripts\Testing Scripts\SMTP\testsmtp.ps1"
        Params={
            $server = Read-Host "  SMTP server [smtp.office365.com]"
            $from   = Read-Host "  From address"
            $to     = Read-Host "  To address"
            $a = @{}
            if ($server) { $a['SmtpServer']=$server }
            if ($from)   { $a['From']=$from }
            if ($to)     { $a['To']=$to }
            return $a
        }
    }
    [PSCustomObject]@{ Key='5'; FKey=[ConsoleKey]::F5; Category='Testing'
        Label='Test-SMTP (5 min)   — recurring SMTP test every 5 minutes'
        Script="$ROOT\scripts\Testing Scripts\SMTP\testsmtp_5min.ps1"
        Params={
            $server = Read-Host "  SMTP server [smtp.office365.com]"
            $from   = Read-Host "  From address"
            $to     = Read-Host "  To address"
            $a = @{}
            if ($server) { $a['SmtpServer']=$server }
            if ($from)   { $a['From']=$from }
            if ($to)     { $a['To']=$to }
            return $a
        }
    }
    [PSCustomObject]@{ Key='6'; FKey=[ConsoleKey]::F6; Category='Device'
        Label='Restart-Time-Sync   — force Windows time service sync'
        Script="$ROOT\scripts\Custom Scripts\device\Time sync\Restart-Time-Sync.ps1"
        Params={ return @{} }
    }
    [PSCustomObject]@{ Key='7'; FKey=[ConsoleKey]::F7; Category='Device'
        Label='Detect-AudioDevices — list connected audio devices'
        Script="$ROOT\scripts\Custom Scripts\device\audio\detect-audiodevices.ps1"
        Params={ return @{} }
    }
    [PSCustomObject]@{ Key='8'; FKey=[ConsoleKey]::F8; Category='Device'
        Label='Disable-InternalMic — disable internal microphone via policy'
        Script="$ROOT\scripts\Custom Scripts\device\audio\Disable-internalmic.ps1"
        Params={ return @{} }
    }
    [PSCustomObject]@{ Key='9'; FKey=[ConsoleKey]::F9; Category='Startup'
        Label='Install-Modules     — bootstrap: install all required PS modules'
        Script="$ROOT\scripts\Startup\Install-Modules.ps1"
        Params={ return @{} }
    }
    [PSCustomObject]@{ Key='A'; FKey=[ConsoleKey]::F10; Category='Reporting'
        Label='Licensing-Report    — generate monthly Pax8 + Ingram Excel report'
        Script="$ROOT\scripts\Reporting\Licensing\genereer_rapport.ps1"
        Params={ return @{} }
    }
    # ── M365 management (via functies.ps1, lazy-loaded on first use) ──────────
    [PSCustomObject]@{ Key='B'; FKey=$null; Category='M365'
        Label='Connect-Tenant      — select customer tenant'
        Action={
            if (-not (Import-Functies)) { return }
            $domain = Read-Host "  Customer domain (e.g. contoso.com)"
            Connect-Tenant -Domain $domain
        }
    }
    [PSCustomObject]@{ Key='C'; FKey=$null; Category='M365'; NoWait=$true
        Label='Exchange Online      — mailbox management'
        Action={
            if (-not (Import-Functies)) { return }
            Invoke-Submenu -Title 'M365 — Exchange Online' -Items $ExchangeSubmenu
        }
    }
    [PSCustomObject]@{ Key='D'; FKey=$null; Category='M365'; NoWait=$true
        Label='Entra ID / Graph    — user and tenant management'
        Action={
            if (-not (Import-Functies)) { return }
            Invoke-Submenu -Title 'M365 — Entra ID / Graph' -Items $EntraSubmenu
        }
    }
    [PSCustomObject]@{ Key='E'; FKey=$null; Category='M365'; NoWait=$true
        Label='MSP Admin           — MSP admin account management'
        Action={
            if (-not (Import-Functies)) { return }
            Invoke-Submenu -Title 'M365 — MSP Admin' -Items $MspSubmenu
        }
    }
)

# ── Display main menu ─────────────────────────────────────────────────────────
# NOTE: Named Show-MainMenu to avoid conflict with Show-Menu in functies.ps1
function Show-MainMenu {
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
        if ($item.FKey) {
            $fNum     = [int]$item.FKey - [int][ConsoleKey]::F1 + 1
            $keyLabel = "[$($item.Key)] / [F$fNum]"
        } else {
            $keyLabel = "[$($item.Key)]"
        }
        Write-Host ("  {0,-22} {1}" -f $keyLabel, $item.Label)
    }

    Write-Host ""
    Write-Host "  [0] / [Esc]  Exit"
    Write-Host ""
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Press a key to launch:" -ForegroundColor Yellow
}

# ── Run a menu item ───────────────────────────────────────────────────────────
function Invoke-MenuItem {
    param([PSCustomObject] $item)

    $isSubmenu = $item.PSObject.Properties['NoWait'] -and $item.NoWait

    # Action-based item (M365 functions / submenus)
    if ($item.PSObject.Properties['Action'] -and $item.Action) {
        if (-not $isSubmenu) {
            Clear-Host
            Write-Host ""
            Write-Host "  ── $($item.Category) / $($item.Label.Split('—')[0].Trim()) ──" -ForegroundColor Cyan
            Write-Host ""
        }
        try {
            & $item.Action
        } catch {
            Write-Host ""
            Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        }
        if (-not $isSubmenu) {
            Write-Host ""
            Write-Host "  Press any key to return to the menu..." -ForegroundColor DarkGray
            $null = [Console]::ReadKey($true)
        }
        return
    }

    # Script-based item
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
    Show-MainMenu

    $keyInfo = [Console]::ReadKey($true)
    $char    = $keyInfo.Key

    if ($keyInfo.KeyChar -eq '0' -or $char -eq [ConsoleKey]::Escape) {
        Clear-Host
        break
    }

    $selected = $menu | Where-Object {
        $_.Key -eq $keyInfo.KeyChar.ToString().ToUpper() -or
        ($_.FKey -and $_.FKey -eq $char)
    } | Select-Object -First 1

    if ($selected) {
        Invoke-MenuItem -item $selected
    }
}
