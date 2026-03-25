#Requires -Version 5.1
<#
.SYNOPSIS
    Diagnose why users cannot log in to an RDP server or RD Web Access server.

.DESCRIPTION
    Performs a comprehensive health check on RDP and RD Web infrastructure.

    RDP Server checks (run on the server itself for full results):
      - TermService, SessionEnv, UmRdpService status
      - RDP enabled/disabled in registry
      - NLA and security layer settings
      - Max session limit vs active session count
      - RD Licensing mode and server configuration
      - Remote Desktop Users group members
      - Windows Firewall inbound RDP rules
      - Active sessions (quser)

    RDWeb Server checks:
      - Port 443 reachability
      - HTTPS certificate validity and expiry
      - IIS service and RDWeb application pool (local only)
      - RD Gateway service (local only)

    User account checks (optional, requires -Username):
      - Account enabled / locked / password expired / account expired
      - Remote Desktop Users group membership
      - Logon workstation restrictions
      - Last logon date and password age

    Event log analysis (optional, requires -IncludeEventLogs):
      - Security 4625 (failed logon, RDP type only)
      - Security 4740 (account lockout)
      - TerminalServices-LocalSessionManager 20/40 (session failure / disconnect reason)

    All results are written to both the console and a timestamped log file.
    If the RDP or RDWeb server is remote, only connectivity checks are performed —
    run the script directly on the target server for full diagnostics.

.PARAMETER RdpServer
    Hostname or IP of the RDP / Terminal Server to check.
    Omit to run local checks on the current machine.

.PARAMETER RdWebServer
    Hostname or IP of the RD Web Access server.
    Port 443 and the HTTPS certificate are always checked.
    IIS / app pool checks require running the script on that server.

.PARAMETER Username
    Check a specific user account for common login blockers.
    Requires the ActiveDirectory module or falls back to ADSI.

.PARAMETER LogPath
    Folder for the output log file. Default: C:\Temp\.

.PARAMETER IncludeEventLogs
    Include event log analysis for the last N hours (see -Hours).

.PARAMETER Hours
    Number of hours of event log history to analyse. Default: 24.

.EXAMPLE
    # Full check — RDP + RDWeb + user account
    .\Test-RDSDiagnostics.ps1 -RdpServer rdp01.company.local -RdWebServer rdweb.company.local -Username jdoe -IncludeEventLogs

.EXAMPLE
    # Run locally on the RDS host, check event logs
    .\Test-RDSDiagnostics.ps1 -IncludeEventLogs -Hours 48

.EXAMPLE
    # Remote connectivity + certificate check only
    .\Test-RDSDiagnostics.ps1 -RdpServer rdp01.company.local -RdWebServer rdweb.company.local

.EXAMPLE
    # Check why a specific user cannot log in (run on the RDS host)
    .\Test-RDSDiagnostics.ps1 -Username jdoe -IncludeEventLogs
#>
[CmdletBinding()]
param (
    [string] $RdpServer,
    [string] $RdWebServer,
    [string] $Username,
    [string] $LogPath = 'C:\Temp',
    [switch] $IncludeEventLogs,
    [int]    $Hours = 24
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'SilentlyContinue'

# ── Log setup ──────────────────────────────────────────────────────────────────
if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath | Out-Null }
$ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile = Join-Path $LogPath "RDS_Diagnostics_$ts.log"
$issues  = [System.Collections.Generic.List[string]]::new()

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line  = '[{0}] [{1,-4}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    $color = switch ($Level) {
        'OK'   { 'Green'   }
        'WARN' { 'Yellow'  }
        'FAIL' { 'Red'     }
        'HEAD' { 'Cyan'    }
        default{ 'Gray'    }
    }
    Write-Host $line -ForegroundColor $color
    Add-Content -Path $logFile -Value $line -Encoding UTF8
}

function Write-Section {
    param([string]$Title)
    $sep = '─' * 56
    $block = "`n  $sep`n   $Title`n  $sep"
    Write-Host $block -ForegroundColor Cyan
    Add-Content -Path $logFile -Value $block -Encoding UTF8
}

function Add-Issue {
    param([string]$Msg)
    $issues.Add($Msg)
    Write-Log $Msg 'FAIL'
}

# ── Helpers ────────────────────────────────────────────────────────────────────
function Test-PortOpen {
    param([string]$Server, [int]$Port, [int]$TimeoutMs = 3000)
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $ar  = $tcp.BeginConnect($Server, $Port, $null, $null)
        $ok  = $ar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        $tcp.Close()
        return $ok
    } catch { return $false }
}

function Get-CertExpiry {
    param([string]$Server, [int]$Port = 443)
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient($Server, $Port)
        $ssl = New-Object System.Net.Security.SslStream(
            $tcp.GetStream(), $false,
            [System.Net.Security.RemoteCertificateValidationCallback]{ $true }
        )
        $ssl.AuthenticateAsClient($Server)
        $cert   = $ssl.RemoteCertificate
        $expiry = [DateTime]::Parse($cert.GetExpirationDateString())
        $subject = $cert.Subject
        $ssl.Close(); $tcp.Close()
        return [PSCustomObject]@{ Expiry = $expiry; Subject = $subject; DaysLeft = ($expiry - (Get-Date)).Days }
    } catch { return $null }
}

# Determine if specified servers are local
$localNames = @($env:COMPUTERNAME, 'localhost', '127.0.0.1', '.')
try { $localNames += ([System.Net.Dns]::GetHostEntry('').HostName) } catch {}
$isRdpLocal   = (-not $RdpServer)  -or ($RdpServer  -in $localNames)
$isRdWebLocal = $RdWebServer -and ($RdWebServer -in $localNames)

# ── Banner ─────────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  ════════════════════════════════════════════' -ForegroundColor Cyan
Write-Host '   Test-RDSDiagnostics' -ForegroundColor Cyan
Write-Host '  ════════════════════════════════════════════' -ForegroundColor Cyan
Write-Log "Log file      : $logFile"
Write-Log "RDP Server    : $(if ($RdpServer) { $RdpServer } else { "(local — $env:COMPUTERNAME)" })"
Write-Log "RDWeb Server  : $(if ($RdWebServer) { $RdWebServer } else { '(not specified)' })"
Write-Log "Username      : $(if ($Username) { $Username } else { '(not specified)' })"
Write-Log "Event logs    : $(if ($IncludeEventLogs) { "Yes — last $Hours hours" } else { 'No (use -IncludeEventLogs)' })"

# ── 1. Network Connectivity ────────────────────────────────────────────────────
Write-Section '1. Network Connectivity'

foreach ($server in @($RdpServer, $RdWebServer) | Where-Object { $_ }) {
    try {
        $ips = [System.Net.Dns]::GetHostAddresses($server) | ForEach-Object { $_.IPAddressToString }
        Write-Log "DNS [$server] → $($ips -join ', ')" 'OK'
    } catch {
        Add-Issue "DNS resolution failed for '$server'"
    }

    $ping = Test-Connection -ComputerName $server -Count 1 -Quiet -ErrorAction SilentlyContinue
    if ($ping) { Write-Log "Ping [$server] OK" 'OK' }
    else        { Write-Log "Ping [$server] failed (may be blocked by firewall)" 'WARN' }
}

if ($RdpServer) {
    if (Test-PortOpen $RdpServer 3389) { Write-Log "Port 3389 (RDP) on $RdpServer is OPEN" 'OK' }
    else                               { Add-Issue  "Port 3389 (RDP) on $RdpServer is NOT reachable" }
}

if ($RdWebServer) {
    if (Test-PortOpen $RdWebServer 443) { Write-Log "Port 443 (HTTPS) on $RdWebServer is OPEN" 'OK' }
    else                                { Add-Issue  "Port 443 (HTTPS) on $RdWebServer is NOT reachable" }

    if (Test-PortOpen $RdWebServer 80) { Write-Log "Port 80 (HTTP) on $RdWebServer is open" 'OK' }
    else                               { Write-Log "Port 80 (HTTP) on $RdWebServer is closed" 'WARN' }
}

# ── 2. RDP Server ─────────────────────────────────────────────────────────────
Write-Section '2. RDP Server'

if (-not $isRdpLocal) {
    Write-Log "Remote server specified — run the script on $RdpServer for full local checks" 'WARN'
} else {
    # Services
    $rdpServices = @(
        @{ Name = 'TermService';  Label = 'Remote Desktop Services (TermService)' },
        @{ Name = 'SessionEnv';   Label = 'RD Configuration (SessionEnv)' },
        @{ Name = 'UmRdpService'; Label = 'RD UserMode Port Redirector (UmRdpService)' }
    )
    foreach ($s in $rdpServices) {
        $svc = Get-Service -Name $s.Name -ErrorAction SilentlyContinue
        if (-not $svc)                   { Write-Log "$($s.Label): not installed" 'WARN' }
        elseif ($svc.Status -eq 'Running') { Write-Log "$($s.Label): Running" 'OK' }
        else                             { Add-Issue "$($s.Label) is $($svc.Status)" }
    }

    # RDP enabled / disabled
    $tsReg  = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server'
    $fDeny  = (Get-ItemProperty $tsReg -ErrorAction SilentlyContinue).fDenyTSConnections
    if ($fDeny -eq 0) { Write-Log "RDP enabled (fDenyTSConnections = 0)" 'OK' }
    else              { Add-Issue "RDP is DISABLED (fDenyTSConnections = $fDeny) — check System Properties or GPO" }

    # NLA + security layer
    $rdpTcp  = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'
    $nla     = (Get-ItemProperty $rdpTcp -ErrorAction SilentlyContinue).UserAuthentication
    $secLvl  = (Get-ItemProperty $rdpTcp -ErrorAction SilentlyContinue).SecurityLayer
    $nlaStr  = if ($nla -eq 1) { 'Required' } else { 'Not required' }
    $secStr  = switch ($secLvl) { 0 { 'RDP (legacy)' } 1 { 'Negotiate' } 2 { 'SSL/TLS' } default { "Unknown ($secLvl)" } }
    Write-Log "NLA: $nlaStr | Security layer: $secStr" 'INFO'
    if ($nla -eq 1) { Write-Log "NLA is required — client must support CredSSP and user must have valid credentials before session starts" 'WARN' }

    # Session limit
    $maxSess = (Get-ItemProperty $rdpTcp -ErrorAction SilentlyContinue).MaxInstanceCount
    $activeSessions = @(quser 2>$null | Select-Object -Skip 1 | Where-Object { $_ -match '\S' })
    $sessCount = $activeSessions.Count

    if ($maxSess -and $maxSess -lt 999999) {
        if ($sessCount -ge $maxSess) { Add-Issue "Session limit reached: $sessCount / $maxSess active sessions" }
        else                        { Write-Log "Active sessions: $sessCount / $maxSess (limit)" 'OK' }
    } else {
        Write-Log "Active sessions: $sessCount (no hard limit configured)" 'OK'
    }

    if ($activeSessions) {
        Write-Log "Current sessions:" 'INFO'
        $activeSessions | ForEach-Object { Write-Log "  $_" 'INFO' }
    }

    # RD Licensing
    $licReg    = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'
    $licMode   = (Get-ItemProperty $licReg -ErrorAction SilentlyContinue).LicensingMode
    $licServer = (Get-ItemProperty $licReg -ErrorAction SilentlyContinue).LicenseServers
    $licStr    = switch ($licMode) {
        2       { 'Per Device' }
        4       { 'Per User'   }
        $null   { 'Not configured via GPO (grace period may apply)' }
        default { "Unknown ($licMode)" }
    }
    Write-Log "RD Licensing mode: $licStr" 'INFO'
    if ($licServer) { Write-Log "License server: $licServer" 'INFO' }
    else            { Write-Log "No license server configured via GPO" 'WARN' }

    $licSvc = Get-Service -Name 'TermServLicensing' -ErrorAction SilentlyContinue
    if ($licSvc) {
        $lvl = if ($licSvc.Status -eq 'Running') { 'OK' } else { 'WARN' }
        Write-Log "RD Licensing service (TermServLicensing): $($licSvc.Status)" $lvl
    }

    # Firewall rules
    $fwRules = Get-NetFirewallRule -DisplayGroup 'Remote Desktop' -ErrorAction SilentlyContinue |
        Where-Object { $_.Enabled -eq 'True' -and $_.Direction -eq 'Inbound' }
    if ($fwRules) { Write-Log "Windows Firewall: $($fwRules.Count) active inbound RDP rule(s)" 'OK' }
    else          { Add-Issue "No active inbound Windows Firewall rules found for Remote Desktop" }

    # Remote Desktop Users group
    try {
        $members = Get-LocalGroupMember -Group 'Remote Desktop Users' -ErrorAction Stop
        if ($members) { Write-Log "Remote Desktop Users group: $($members.Count) member(s) — $($members.Name -join ', ')" 'INFO' }
        else          { Write-Log "Remote Desktop Users group is empty (only local admins can connect)" 'WARN' }
    } catch {
        Write-Log "Could not read Remote Desktop Users group: $_" 'WARN'
    }

    # Port 3389 listener (local confirmation)
    $listening = Get-NetTCPConnection -LocalPort 3389 -State Listen -ErrorAction SilentlyContinue
    if ($listening) { Write-Log "Port 3389 is actively listening on this machine" 'OK' }
    else            { Add-Issue "Port 3389 is NOT listening on this machine — TermService may have failed" }
}

# ── 3. Event Logs ─────────────────────────────────────────────────────────────
if ($IncludeEventLogs) {
    Write-Section "3. Event Logs (last $Hours hours)"

    if (-not $isRdpLocal) {
        Write-Log "Cannot read event logs of a remote server from here — run the script on $RdpServer" 'WARN'
    } else {
        $since = (Get-Date).AddHours(-$Hours)

        # Failed RDP logons (4625 type 10 = RemoteInteractive)
        try {
            $failed = Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4625; StartTime = $since } `
                -MaxEvents 200 -ErrorAction Stop |
                Where-Object { $_.Properties[10].Value -eq 10 }
            if ($failed) {
                Write-Log "$($failed.Count) failed RDP logon(s) detected:" 'WARN'
                $failed | Group-Object { $_.Properties[5].Value } |
                    Sort-Object Count -Descending |
                    ForEach-Object { Write-Log "  User '$($_.Name)': $($_.Count) failure(s)" 'WARN' }
            } else { Write-Log "No failed RDP logons (Event 4625) in the last $Hours hours" 'OK' }
        } catch { Write-Log "Could not read Security log (may need admin): $_" 'WARN' }

        # Account lockouts (4740)
        try {
            $lockouts = Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4740; StartTime = $since } `
                -MaxEvents 50 -ErrorAction Stop
            if ($lockouts) {
                Write-Log "$($lockouts.Count) account lockout(s) detected:" 'WARN'
                $lockouts | ForEach-Object {
                    Write-Log "  $($_.TimeCreated.ToString('HH:mm:ss')) — '$($_.Properties[0].Value)' locked out by '$($_.Properties[1].Value)'" 'WARN'
                }
            } else { Write-Log "No account lockouts (Event 4740) in the last $Hours hours" 'OK' }
        } catch { Write-Log "Could not read lockout events: $_" 'WARN' }

        # RDS session failures and disconnects (20 = logon failed, 40 = disconnect reason)
        try {
            $rdsEvents = Get-WinEvent -FilterHashtable @{
                LogName   = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'
                Id        = @(20, 40)
                StartTime = $since
            } -MaxEvents 50 -ErrorAction Stop
            if ($rdsEvents) {
                Write-Log "$($rdsEvents.Count) RDS session failure/disconnect event(s):" 'WARN'
                $rdsEvents | Select-Object -First 15 | ForEach-Object {
                    $msg = ($_.Message -replace '\s+', ' ').Substring(0, [Math]::Min(120, $_.Message.Length))
                    Write-Log "  [$($_.TimeCreated.ToString('HH:mm:ss'))] Event $($_.Id): $msg" 'WARN'
                }
            } else { Write-Log "No session failure/disconnect events (20/40) in the last $Hours hours" 'OK' }
        } catch { Write-Log "Could not read TerminalServices-LocalSessionManager log: $_" 'WARN' }

        # Kerberos failures (4771) — common with RDP + NLA
        try {
            $kerb = Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4771; StartTime = $since } `
                -MaxEvents 30 -ErrorAction Stop
            if ($kerb) {
                Write-Log "$($kerb.Count) Kerberos pre-authentication failure(s) — may indicate bad password or clock skew:" 'WARN'
                $kerb | Group-Object { $_.Properties[0].Value } |
                    ForEach-Object { Write-Log "  User '$($_.Name)': $($_.Count) failure(s)" 'WARN' }
            } else { Write-Log "No Kerberos pre-auth failures (Event 4771) in the last $Hours hours" 'OK' }
        } catch { Write-Log "Could not read Kerberos events (may not be a DC): $_" 'INFO' }
    }
}

# ── 4. RD Web Access Server ───────────────────────────────────────────────────
if ($RdWebServer) {
    Write-Section '4. RD Web Access Server'

    # HTTPS certificate check (always, remote-safe)
    if (Test-PortOpen $RdWebServer 443) {
        $certInfo = Get-CertExpiry -Server $RdWebServer
        if ($certInfo) {
            Write-Log "Certificate subject : $($certInfo.Subject)" 'INFO'
            Write-Log "Certificate expires : $($certInfo.Expiry.ToString('yyyy-MM-dd')) ($($certInfo.DaysLeft) days)" (
                if ($certInfo.DaysLeft -lt 0)  { 'FAIL' }
                elseif ($certInfo.DaysLeft -lt 30) { 'WARN' }
                else { 'OK' }
            )
            if ($certInfo.DaysLeft -lt 0)   { Add-Issue "HTTPS certificate on $RdWebServer has EXPIRED" }
            elseif ($certInfo.DaysLeft -lt 30) { Write-Log "Certificate expires in $($certInfo.DaysLeft) days — plan renewal" 'WARN' }
        } else { Write-Log "Could not retrieve HTTPS certificate from $RdWebServer" 'WARN' }
    }

    if (-not $isRdWebLocal) {
        Write-Log "Remote RDWeb server — IIS/app pool checks skipped (run script on $RdWebServer for full diagnostics)" 'WARN'
    } else {
        # IIS service
        $iis = Get-Service -Name 'W3SVC' -ErrorAction SilentlyContinue
        if (-not $iis)                     { Add-Issue "IIS (W3SVC) not found — is RD Web Access installed?" }
        elseif ($iis.Status -eq 'Running') { Write-Log "IIS (W3SVC): Running" 'OK' }
        else                               { Add-Issue "IIS (W3SVC) is $($iis.Status)" }

        # RDWeb app + app pool via WebAdministration
        try {
            Import-Module WebAdministration -ErrorAction Stop
            $app = Get-WebApplication -Site 'Default Web Site' -Name 'RDWeb' -ErrorAction SilentlyContinue
            if ($app) {
                Write-Log "RDWeb IIS application: found (pool: $($app.applicationPool))" 'OK'
                $poolState = (Get-WebConfigurationProperty `
                    -Filter "system.applicationHost/applicationPools/add[@name='$($app.applicationPool)']" `
                    -Name 'state' -ErrorAction SilentlyContinue).Value
                if ($poolState -eq 'Started') { Write-Log "App pool '$($app.applicationPool)': Started" 'OK' }
                elseif ($poolState)           { Add-Issue "App pool '$($app.applicationPool)' is $poolState" }
            } else {
                Add-Issue "RDWeb application not found under 'Default Web Site' in IIS"
            }
        } catch {
            Write-Log "WebAdministration module unavailable — skipping IIS app pool check" 'WARN'
        }

        # RD Gateway service
        $gw = Get-Service -Name 'TSGateway' -ErrorAction SilentlyContinue
        if ($gw) {
            $lvl = if ($gw.Status -eq 'Running') { 'OK' } else { 'WARN' }
            Write-Log "RD Gateway (TSGateway): $($gw.Status)" $lvl
            if ($gw.Status -ne 'Running') { Add-Issue "RD Gateway service is $($gw.Status)" }
        } else {
            Write-Log "RD Gateway: not installed on this server" 'INFO'
        }

        # RD Connection Broker
        $broker = Get-Service -Name 'Tssdis' -ErrorAction SilentlyContinue
        if ($broker) {
            $lvl = if ($broker.Status -eq 'Running') { 'OK' } else { 'WARN' }
            Write-Log "RD Connection Broker (Tssdis): $($broker.Status)" $lvl
        } else {
            Write-Log "RD Connection Broker: not installed on this server" 'INFO'
        }

        # Event Viewer — RDWeb + IIS + Gateway logs
        if ($IncludeEventLogs) {
            $since = (Get-Date).AddHours(-$Hours)
            Write-Log "--- Event Viewer (last $Hours hours) ---" 'INFO'

            # RD Web Access Admin + Operational
            $rdWebLogs = @(
                'Microsoft-Windows-TerminalServices-WebAccess/Admin',
                'Microsoft-Windows-TerminalServices-WebAccess/Operational'
            )
            foreach ($log in $rdWebLogs) {
                try {
                    $events = Get-WinEvent -FilterHashtable @{ LogName = $log; StartTime = $since } `
                        -MaxEvents 30 -ErrorAction Stop |
                        Where-Object { $_.Level -le 3 }  # Critical, Error, Warning only
                    if ($events) {
                        Write-Log "$($events.Count) error/warning event(s) in '$log':" 'WARN'
                        $events | Select-Object -First 10 | ForEach-Object {
                            $msg = ($_.Message -replace '\s+', ' ')
                            $msg = $msg.Substring(0, [Math]::Min(120, $msg.Length))
                            Write-Log "  [$($_.TimeCreated.ToString('HH:mm:ss'))] ID $($_.Id): $msg" 'WARN'
                        }
                    } else {
                        Write-Log "No errors/warnings in '$log'" 'OK'
                    }
                } catch {
                    Write-Log "Log '$log' not found or empty (RD Web Access may not be installed)" 'INFO'
                }
            }

            # RD Gateway Admin + Operational
            $gwLogs = @(
                'Microsoft-Windows-TerminalServices-Gateway/Admin',
                'Microsoft-Windows-TerminalServices-Gateway/Operational'
            )
            foreach ($log in $gwLogs) {
                try {
                    $events = Get-WinEvent -FilterHashtable @{ LogName = $log; StartTime = $since } `
                        -MaxEvents 30 -ErrorAction Stop |
                        Where-Object { $_.Level -le 3 }
                    if ($events) {
                        Write-Log "$($events.Count) error/warning event(s) in '$log':" 'WARN'
                        $events | Select-Object -First 10 | ForEach-Object {
                            $msg = ($_.Message -replace '\s+', ' ')
                            $msg = $msg.Substring(0, [Math]::Min(120, $msg.Length))
                            Write-Log "  [$($_.TimeCreated.ToString('HH:mm:ss'))] ID $($_.Id): $msg" 'WARN'
                        }
                    } else {
                        Write-Log "No errors/warnings in '$log'" 'OK'
                    }
                } catch {
                    Write-Log "Log '$log' not found (RD Gateway not installed on this server)" 'INFO'
                }
            }

            # IIS Application log (ASP.NET / application pool crashes)
            try {
                $iisErrors = Get-WinEvent -FilterHashtable @{
                    LogName   = 'Application'
                    StartTime = $since
                } -MaxEvents 200 -ErrorAction Stop |
                    Where-Object { $_.Level -le 2 -and $_.ProviderName -match 'IIS|W3SVC|ASP|HttpErr|WAS' }
                if ($iisErrors) {
                    Write-Log "$($iisErrors.Count) IIS/ASP.NET error(s) in Application log:" 'WARN'
                    $iisErrors | Select-Object -First 10 | ForEach-Object {
                        $msg = ($_.Message -replace '\s+', ' ')
                        $msg = $msg.Substring(0, [Math]::Min(120, $msg.Length))
                        Write-Log "  [$($_.TimeCreated.ToString('HH:mm:ss'))] $($_.ProviderName): $msg" 'WARN'
                    }
                } else {
                    Write-Log "No IIS/ASP.NET errors in Application log" 'OK'
                }
            } catch {
                Write-Log "Could not read Application event log: $_" 'WARN'
            }
        }
    }
}

# ── 5. User Account ───────────────────────────────────────────────────────────
if ($Username) {
    Write-Section "5. User Account: $Username"

    $adUser = $null
    try {
        $adUser = Get-ADUser -Identity $Username -Properties `
            Enabled, LockedOut, PasswordExpired, PasswordLastSet,
            LastLogonDate, MemberOf, LogonWorkstations, AccountExpirationDate `
            -ErrorAction Stop
    } catch {
        Write-Log "ActiveDirectory module unavailable or user not found — falling back to ADSI" 'WARN'
    }

    if ($adUser) {
        if ($adUser.Enabled)         { Write-Log "Account enabled       : Yes" 'OK'   }
        else                         { Add-Issue  "Account '$Username' is DISABLED"    }

        if ($adUser.LockedOut)       { Add-Issue  "Account '$Username' is LOCKED OUT"  }
        else                         { Write-Log "Account locked out    : No" 'OK'    }

        if ($adUser.PasswordExpired) { Add-Issue  "Password for '$Username' has EXPIRED" }
        else                         { Write-Log "Password expired      : No" 'OK'    }

        if ($adUser.AccountExpirationDate -and $adUser.AccountExpirationDate -lt (Get-Date)) {
            Add-Issue "Account '$Username' EXPIRED on $($adUser.AccountExpirationDate.ToString('yyyy-MM-dd'))"
        }

        $lastLogon = if ($adUser.LastLogonDate) { $adUser.LastLogonDate.ToString('yyyy-MM-dd HH:mm') } else { 'Never / unknown' }
        $pwdSet    = if ($adUser.PasswordLastSet) { $adUser.PasswordLastSet.ToString('yyyy-MM-dd') } else { 'Never' }
        Write-Log "Last logon            : $lastLogon" 'INFO'
        Write-Log "Password last set     : $pwdSet" 'INFO'

        if ($adUser.LogonWorkstations) {
            Write-Log "Logon restricted to   : $($adUser.LogonWorkstations)" 'WARN'
        }

        # Remote Desktop Users membership
        try {
            $rdGroup = Get-ADGroup 'Remote Desktop Users' -ErrorAction Stop
            if ($adUser.MemberOf -contains $rdGroup.DistinguishedName) {
                Write-Log "Remote Desktop Users  : Member" 'OK'
            } else {
                Write-Log "Remote Desktop Users  : NOT a member (only admins can connect unless added to this group)" 'WARN'
            }
        } catch {
            Write-Log "Could not check Remote Desktop Users group membership" 'WARN'
        }

    } else {
        # ADSI fallback
        try {
            $searcher = New-Object DirectoryServices.DirectorySearcher
            $searcher.Filter = "(sAMAccountName=$Username)"
            $result = $searcher.FindOne()
            if ($result) {
                $uac      = $result.Properties['useraccountcontrol'][0]
                $disabled = ($uac -band 0x2)      -ne 0
                $locked   = ($uac -band 0x10)     -ne 0
                $noPwd    = ($uac -band 0x800000) -ne 0

                if ($disabled) { Add-Issue "Account '$Username' is DISABLED (UAC flag 0x2)" }
                else           { Write-Log "Account appears enabled (ADSI)" 'OK' }
                if ($locked)   { Add-Issue "Account '$Username' may be locked (UAC flag 0x10)" }
                Write-Log "ADSI check complete — install ActiveDirectory module (RSAT) for full account details" 'INFO'
            } else {
                Add-Issue "User '$Username' was not found in Active Directory"
            }
        } catch {
            Write-Log "ADSI lookup failed: $_" 'WARN'
        }
    }
}

# ── Summary ────────────────────────────────────────────────────────────────────
Write-Section 'Summary'

if ($issues.Count -eq 0) {
    Write-Log "No issues found — RDS infrastructure looks healthy" 'OK'
} else {
    Write-Log "$($issues.Count) issue(s) found that may prevent users from logging in:" 'FAIL'
    for ($i = 0; $i -lt $issues.Count; $i++) {
        Write-Log "  $($i + 1). $($issues[$i])" 'FAIL'
    }
}

Write-Log "Log saved to: $logFile" 'INFO'
Write-Host ''
