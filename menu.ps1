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

function Set-StartupLauncher {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Enable
    )

    if (-not $IsWindows) {
        Write-Warning 'Startup shortcut setup is only supported on Windows.'
        return
    }

    $startupDir = [Environment]::GetFolderPath('Startup')
    $shortcutPath = Join-Path $startupDir 'M365-Scripts Launcher.lnk'

    if (-not $Enable) {
        if (Test-Path $shortcutPath) {
            Remove-Item -Path $shortcutPath -Force
            Write-Host "  Removed startup shortcut: $shortcutPath" -ForegroundColor Green
        } else {
            Write-Host '  Startup shortcut was not present.' -ForegroundColor DarkGray
        }
        return
    }

    $pwshPath = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
    if (-not $pwshPath) {
        throw 'pwsh.exe not found on PATH. Install PowerShell 7 to enable startup launcher.'
    }

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $pwshPath
    $shortcut.Arguments = "-NoLogo -NoProfile -ExecutionPolicy Bypass -File `"$ROOT\load.ps1`""
    $shortcut.WorkingDirectory = $ROOT
    $shortcut.Description = 'Start M365-Scripts launcher at sign-in'
    $shortcut.Save()

    Write-Host "  Startup shortcut created: $shortcutPath" -ForegroundColor Green
}

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

        if ($global:authMode -eq 'GDAP' -and $global:defaultCustomerDomain -and -not $global:cid) {
            Write-Host "  Selecting GDAP customer tenant: $($global:defaultCustomerDomain)" -ForegroundColor DarkGray
            try {
                Connect-Tenant -Domain $global:defaultCustomerDomain
            } catch {
                Write-Warning "Could not auto-select default customer domain '$($global:defaultCustomerDomain)': $($_.Exception.Message)"
            }
        }

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
    @{ Key='8'; Label='Test-CalendarPermissions  — audit calendar folder permissions'; Action={
        $mbx = Read-Host "  Mailbox UPN (leave blank for all)"
        $p = @{}
        if ($mbx) { $p['Mailbox'] = $mbx }
        & "$ROOT\scripts\Exchange\Test-CalendarPermissions.ps1" @p
    }}
    @{ Key='9'; Label='Test-MailboxPermissions   — audit Full Access / Send As / Send on Behalf'; Action={
        $mbx = Read-Host "  Mailbox UPN (leave blank for all)"
        $p = @{}
        if ($mbx) { $p['Mailbox'] = $mbx }
        & "$ROOT\scripts\Exchange\Test-MailboxPermissions.ps1" @p
    }}
    @{ Key='A'; Label='Test-GroupPermissions     — audit DG managers / Send As / Send on Behalf'; Action={
        $grp = Read-Host "  Group email or name (leave blank for all)"
        $inc = Read-Host "  Include member list? [y/N]"
        $p = @{}
        if ($grp) { $p['Group'] = $grp }
        if ($inc -match '^[Yy]') { $p['IncludeMembers'] = $true }
        & "$ROOT\scripts\Exchange\Test-DistributionGroupPermissions.ps1" @p
    }}
    @{ Key='B'; Label='Test-DkimConfig          — validate DKIM signing config and DNS records'; Action={
        $domain = Read-Host "  Domain (leave blank for all)"
        $p = @{}
        if ($domain) { $p['Domain'] = $domain }
        & "$ROOT\scripts\Exchange\Test-DkimConfig.ps1" @p
    }}
    @{ Key='C'; Label='Get-ExternalForwards     — audit mailboxes with external forwarding'; Action={
        $mbx = Read-Host "  Mailbox UPN (leave blank for all)"
        $p = @{}
        if ($mbx) { $p['Mailbox'] = $mbx }
        & "$ROOT\scripts\Exchange\Get-ExternalForwards.ps1" @p
    }}
    @{ Key='D'; Label='Get-MailboxSizes         — report mailbox sizes sorted by storage used'; Action={
        $mbx = Read-Host "  Mailbox UPN (leave blank for all)"
        $p = @{}
        if ($mbx) { $p['Mailbox'] = $mbx }
        & "$ROOT\scripts\Exchange\Get-MailboxSizes.ps1" @p
    }}
    @{ Key='E'; Label='Move-InboxToArchive      — archive Inbox messages to Archive folder'; Action={
        $path  = Join-Path $ROOT 'scripts\Exchange\Move-InboxToArchive.ps1'
        $mbx   = Read-Host "  Mailbox UPN"
        $after = Read-Host "  Only messages on/after this date [yyyy-MM-dd] (optional)"
        $before = Read-Host "  Only messages before this date [yyyy-MM-dd] (optional)"
        $deleg = Read-Host "  Delegated mode (Exchange Admin, no temp app) instead of automatic? [y/N]"
        $p = @{ Mailbox = $mbx }
        if ($after)  { $p['After']  = [datetime]$after }
        if ($before) { $p['Before'] = [datetime]$before }
        if ($deleg -match '^[Yy]') { $p['Delegated'] = $true }
        & $path @p
        $confirm = Read-Host "  Preview completed. Add -Apply to actually move messages? [y/N]"
        if ($confirm -match '^[Yy]') { & $path @p -Apply }
    }}
    @{ Key='F'; Label='Set-DL-Dynamic-Static    — resolve a dynamic DG into a static group'; Action={
        $path  = Join-Path $ROOT 'scripts\Exchange\Set-Distributionlist-dynamic-static.ps1'
        $dyn   = Read-Host "  Dynamic distribution group identity"
        $tgt   = Read-Host "  Target static group identity"
        $clear = Read-Host "  Clear existing target members first? [y/N]"
        $whatIf = Read-Host "  Preview only (-WhatIf), no changes made? [Y/n]"
        $p = @{ DynamicGroupIdentity = $dyn; TargetGroupIdentity = $tgt }
        if ($clear -match '^[Yy]') { $p['ClearTargetMembers'] = $true }
        if ($whatIf -notmatch '^[Nn]') { $p['WhatIf'] = $true }
        & $path @p
    }}
    @{ Key='G'; Label='Get-MessageTraceReport   — who received what, when, and where it was forwarded'; Action={
        $mbx  = Read-Host "  Mailbox UPN (traces sent + received, leave blank to filter manually)"
        $days = Read-Host "  Days back [2]"
        $p = @{}
        if ($mbx)  { $p['Mailbox'] = $mbx }
        else {
            $snd = Read-Host "  Sender address (optional)"
            $rcp = Read-Host "  Recipient address (optional)"
            if ($snd) { $p['Sender']    = $snd }
            if ($rcp) { $p['Recipient'] = $rcp }
        }
        if ($days) { $p['Days'] = [int]$days }
        $det = Read-Host "  Include per-hop delivery details (slower, shows redirects)? [y/N]"
        if ($det -match '^[Yy]') { $p['IncludeDetails'] = $true }
        & "$ROOT\scripts\Exchange\Get-MessageTraceReport.ps1" @p
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
    @{ Key='9'; Label='Remove-M365Users      — bulk delete user accounts'; Action={
        $path = Join-Path $ROOT 'scripts\Entra\Remove-M365Users.ps1'
        $csv  = Read-Host "  CSV/TXT path (or leave blank for -UserList)"
        if ($csv) {
            & $path -CsvPath $csv
            $confirm = Read-Host "  Dry run completed. Add -Apply to delete? [y/N]"
            if ($confirm -match '^[Yy]') { & $path -CsvPath $csv -Apply }
        } else {
            $users = Read-Host "  UPNs (comma-separated)"
            $list  = $users -split '\s*,\s*'
            & $path -UserList $list
            $confirm = Read-Host "  Dry run completed. Add -Apply to delete? [y/N]"
            if ($confirm -match '^[Yy]') { & $path -UserList $list -Apply }
        }
    }}
    @{ Key='A'; Label='Test-M365GroupMembership — audit M365 group owners and members'; Action={
        $grp = Read-Host "  Group name or ID (leave blank for all)"
        $p = @{}
        if ($grp) { $p['Group'] = $grp }
        & "$ROOT\scripts\Entra\Test-M365GroupMembership.ps1" @p
    }}
    @{ Key='B'; Label='New-M365User             — create a single new user'; Action={
        $upn  = Read-Host "  UPN (e.g. j.doe@contoso.com)"
        $dn   = Read-Host "  Display name"
        $fn   = Read-Host "  First name (optional)"
        $ln   = Read-Host "  Last name (optional)"
        $dept = Read-Host "  Department (optional)"
        $sku  = Read-Host "  License SKU (optional, e.g. ENTERPRISEPACK)"
        $p = @{ UserPrincipalName = $upn; DisplayName = $dn }
        if ($fn)   { $p['GivenName']   = $fn }
        if ($ln)   { $p['Surname']     = $ln }
        if ($dept) { $p['Department']  = $dept }
        if ($sku)  { $p['LicenseSkuId'] = $sku }
        & "$ROOT\scripts\Entra\New-M365User.ps1" @p
    }}
    @{ Key='C'; Label='Import-M365Users         — bulk create users from CSV (dry run first)'; Action={
        $path = Join-Path $ROOT 'scripts\Entra\Import-M365Users.ps1'
        $csv  = Read-Host "  CSV path"
        $sku  = Read-Host "  Default license SKU for all users (optional)"
        $p = @{ CsvPath = $csv }
        if ($sku) { $p['LicenseSkuId'] = $sku }
        & $path @p
        $confirm = Read-Host "  Dry run completed. Add -Apply to create accounts? [y/N]"
        if ($confirm -match '^[Yy]') { & $path @p -Apply }
    }}
    @{ Key='D'; Label='New-TemporaryCA          — create temporary CA for user/group'; Action={
        $path   = Join-Path $ROOT 'scripts\Entra\New-TemporaryConditionalAccessPolicy.ps1'
        $typeIn = Read-Host "  Target type [User/Group]"
        $target = Read-Host "  Target object ID"
        $name   = Read-Host "  Policy name (suffix)"
        $modeTime = Read-Host "  Timing mode [Duration/DateTime]"
        $mode   = Read-Host "  Action [RequireMfa/Block]"
        $state  = Read-Host "  State [enabled/reportOnly/disabled]"
        $auto   = Read-Host "  Auto cleanup immediately at expiry in this session? [Y/n]"

        $p = @{}
        $p['TargetType'] = if ($typeIn -match '^(group|g)$') { 'Group' } else { 'User' }
        $p['TargetId']   = $target
        $p['DisplayName'] = $name

        if ($modeTime -match '^(datetime|date|dt)$') {
            $startAt = Read-Host "  Start local datetime [yyyy-MM-dd HH:mm]"
            $endAt   = Read-Host "  End local datetime   [yyyy-MM-dd HH:mm]"
            $p['StartDateTimeLocal'] = [datetime]$startAt
            $p['EndDateTimeLocal']   = [datetime]$endAt
        } else {
            $hours  = Read-Host "  Duration in hours [4]"
            $p['DurationHours'] = if ($hours) { [int]$hours } else { 4 }
        }

        $p['Action'] = if ($mode -match '^(block|b)$') { 'Block' } else { 'RequireMfa' }

        if ($state -match '^(report|reportonly|enabledforreportingbutnotenforced)$') {
            $p['State'] = 'enabledForReportingButNotEnforced'
        } elseif ($state -match '^(disabled|off)$') {
            $p['State'] = 'disabled'
        } else {
            $p['State'] = 'enabled'
        }

        if ($auto -match '^[Nn]') { $p['NoAutoCleanup'] = $true }
        & $path @p

        if ($p['TargetType'] -eq 'User') {
            $makeTap = Read-Host "  Also create TAP code for this user? [y/N]"
            if ($makeTap -match '^[Yy]') {
                $tapPath = Join-Path $ROOT 'scripts\Entra\New-UserTemporaryAccessPass.ps1'
                $tapMinutes = Read-Host "  TAP lifetime in minutes [60]"
                $tapOnce = Read-Host "  TAP one-time only? [Y/n]"

                $tp = @{ UserId = $target }
                $tp['LifetimeMinutes'] = if ($tapMinutes) { [int]$tapMinutes } else { 60 }
                if ($tapOnce -notmatch '^[Nn]') { $tp['IsUsableOnce'] = $true }
                & $tapPath @tp
            }
        }
    }}
    @{ Key='E'; Label='Remove-TemporaryCA       — cleanup temporary CA policies'; Action={
        $path = Join-Path $ROOT 'scripts\Entra\Remove-TemporaryConditionalAccessPolicies.ps1'
        $id   = Read-Host "  Specific policy ID (optional)"
        if ($id) {
            & $path -PolicyId $id
            return
        }

        $all = Read-Host "  Remove ALL TEMP-CA policies? [y/N]"
        if ($all -match '^[Yy]') {
            & $path -RemoveAllTempPolicies
        } else {
            & $path
        }
    }}
    @{ Key='F'; Label='New-UserTAP              — create Temporary Access Pass'; Action={
        $path = Join-Path $ROOT 'scripts\Entra\New-UserTemporaryAccessPass.ps1'
        $user = Read-Host "  User object ID or UPN"
        $mins = Read-Host "  TAP lifetime in minutes [60]"
        $once = Read-Host "  TAP one-time only? [Y/n]"

        $p = @{ UserId = $user }
        $p['LifetimeMinutes'] = if ($mins) { [int]$mins } else { 60 }
        if ($once -notmatch '^[Nn]') { $p['IsUsableOnce'] = $true }
        & $path @p
    }}
    @{ Key='G'; Label='Get-M365UserLicenses     — report assigned licenses for a set of users'; Action={
        $path = Join-Path $ROOT 'scripts\Entra\Get-M365UserLicenses.ps1'
        $csv  = Read-Host "  CSV/TXT path (or leave blank for a UPN list)"
        if ($csv) {
            & $path -CsvPath $csv
        } else {
            $users = Read-Host "  UPNs/emails (comma-separated)"
            & $path -UserList ($users -split '\s*,\s*')
        }
    }}
    @{ Key='H'; Label='Import-CA-Baseline       — import the community CA baseline'; Action={
        $path   = Join-Path $ROOT 'scripts\Entra\Import-ConditionalAccessBaseline.ps1'
        $action = Read-Host "  Action [Import/SetState] [Import]"
        if ($action -match '^(setstate)$') {
            $target = Read-Host "  Target state [disabled/enabledForReportingButNotEnforced/enabled]"
            & $path -Action SetState -TargetState $target
        } else {
            $state = Read-Host "  Policy state on import [disabled/enabledForReportingButNotEnforced] [disabled]"
            $p = @{}
            if ($state) { $p['PolicyStateOnImport'] = $state }
            & $path @p
        }
    }}
    @{ Key='I'; Label='Set-UserManager          — report/bulk-set manager for a set of users'; Action={
        $path = Join-Path $ROOT 'scripts\Entra\Set-UserManager.ps1'
        $mode = Read-Host "  Select users by [G]roup name / [D]epartment / [C]urrent manager / [U]PN list"
        $p = @{}
        switch -Regex ($mode) {
            '^[Gg]' { $p['GroupName']      = Read-Host "  Group display name" }
            '^[Dd]' { $p['Department']     = Read-Host "  Department" }
            '^[Cc]' { $p['CurrentManager'] = Read-Host "  Current manager UPN or object ID" }
            default { $p['UserList']       = (Read-Host "  UPNs (comma-separated)") -split '\s*,\s*' }
        }
        $newMgr = Read-Host "  New manager UPN (leave blank to only report current managers)"
        if ($newMgr) { $p['NewManager'] = $newMgr }
        & $path @p
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
    [PSCustomObject]@{ Key='1'; FKey=[ConsoleKey]::F1; Category='Testing'
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
        Script="$ROOT\scripts\SMTP\testsmtp.ps1"
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
        Script="$ROOT\scripts\SMTP\testsmtp_5min.ps1"
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
        Script="$ROOT\scripts\Device\Time sync\Restart-Time-Sync.ps1"
        Params={ return @{} }
    }
    [PSCustomObject]@{ Key='7'; FKey=[ConsoleKey]::F7; Category='Device'
        Label='Detect-AudioDevices — list connected audio devices'
        Script="$ROOT\scripts\Device\audio\detect-audiodevices.ps1"
        Params={ return @{} }
    }
    [PSCustomObject]@{ Key='8'; FKey=[ConsoleKey]::F8; Category='Device'
        Label='Disable-InternalMic — disable internal microphone via policy'
        Script="$ROOT\scripts\Device\audio\Disable-internalmic.ps1"
        Params={ return @{} }
    }
    [PSCustomObject]@{ Key='I'; FKey=$null; Category='Device'
        Label='Remove-OemBloatware — remove OEM + generic Store bloatware'
        Script="$ROOT\scripts\Device\Remove-OemBloatware.ps1"
        Params={
            $apply = Read-Host "  Remove apps now (not just preview)? [y/N]"
            $a = @{}
            if ($apply -match '^[Yy]') { $a['Apply'] = $true }
            return $a
        }
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
    [PSCustomObject]@{ Key='F'; FKey=$null; Category='Startup'
        Label='Enable-LauncherStartup — run launcher at Windows sign-in'
        Action={
            Set-StartupLauncher -Enable $true
        }
    }
    [PSCustomObject]@{ Key='G'; FKey=$null; Category='Startup'
        Label='Disable-LauncherStartup — remove launcher from Windows startup'
        Action={
            Set-StartupLauncher -Enable $false
        }
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
    [PSCustomObject]@{ Key='H'; FKey=$null; Category='M365'
        Label='Test-GdapConnection — validate delegated GDAP for a customer tenant'
        Action={
            if (-not (Import-Functies)) { return }
            $domain = Read-Host "  Customer domain (optional, blank uses current/default)"
            if ($domain) { Test-GdapConnection -Domain $domain }
            else { Test-GdapConnection }
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
