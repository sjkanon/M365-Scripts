#Requires -Version 5.1
<#
.SYNOPSIS
    Put messages that were moved or deleted on a given day back where they came
    from, and report who moved or deleted them. Preview by default.

.DESCRIPTION
    Answers two questions about one mailbox and one day (or window):
    "who moved or deleted these messages?" and "put them back".

    Three sources, because no single one covers every case:

      Audit log  Search-UnifiedAuditLog for Move, MoveToDeletedItems, SoftDelete
                 and HardDelete on this mailbox. This is the only record of WHO
                 did it (account, owner/delegate/admin, client, IP, app ID) and
                 the only record of which folder a moved message came FROM -
                 Graph and Exchange keep no such property for a plain move.

      Deleted    Get-/Restore-RecoverableItems over Deleted Items, Recoverable
                 Items (soft-deleted) and Purges (hard-deleted, only kept under a
                 hold), filtered on the moment of deletion. Exchange remembers the
                 original folder of a deleted item itself, so these go back to
                 where they were without needing the audit log. Who deleted each
                 one is looked up in the audit log by subject and time.

                 Without the Mailbox Import Export role those cmdlets do not
                 exist for you, and the run switches to Graph: every message in
                 Deleted Items and Recoverable Items\Deletions that changed in
                 the window goes back. Audited ones go to the folder the audit
                 log says they left; the rest, and anything deleted out of
                 Deleted Items itself, go to the Inbox. Purges (hard-deleted) is
                 out of Graph's reach and is reported as Unreachable.

      Moved      Every audited Move on the day is traced back to the first folder
                 the message left, the message is located over Graph by its
                 Internet MessageId, and moved back. Moves out of Deleted Items or
                 Recoverable Items are skipped: those were restores, and undoing
                 them would delete the message again.

    Messages moved into the Archive folder without an audit record - by default
    Exchange does not audit the owner's own moves, and archiving by an app can
    be missing too - are listed from the Archive folder by their modification
    time. They are only moved (to the Inbox) with -UnauditedArchiveToInbox,
    because a modification time is a heuristic: reading or flagging a message in
    Archive changes it as well.

    NOTHING IS MOVED WITHOUT -Apply. Every run without it does the full search
    and reports what it would restore, and who did what.

.PARAMETER Mailbox
    UPN or primary SMTP address of the mailbox.

.PARAMETER Date
    The day on which the messages were moved or deleted (local time, whole day).
    Use -After/-Before instead for a precise window.

.PARAMETER After
    Start of the window in which the messages were moved or deleted (local time).
    Without -Before this means "everything from this date until now". Aliases:
    -From, -Since.

.PARAMETER Before
    End of that window (local time). Default: now.

.PARAMETER Include
    Which part to run: Deleted, Moved, or both (default). The audit report of
    who did what is always produced.

.PARAMETER UnauditedArchiveToInbox
    Also move messages in the Archive folder that were modified in the window
    and have no audit record back to the Inbox. Without it they are listed only.

.PARAMETER Apply
    Actually restore. Without this switch nothing is moved.

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\ (Windows) or ~/Downloads.

.PARAMETER TenantId
    Entra ID tenant - tenant ID (GUID) or a verified domain. Required for
    app-only Graph auth unless it can be resolved from a GDAP customer context.

.PARAMETER ClientId
    Existing App Registration for app-only Graph auth - skips the automatic
    temporary app. Needs Mail.ReadWrite application permission (admin consent).

.PARAMETER ClientSecret
    Client secret for the app registration given in -ClientId.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for the app registration given in -ClientId.

.EXAMPLE
    # What was moved or deleted on 25 September, by whom, and what would go back?
    .\Restore-MailboxMessages.ps1 -Mailbox "user@contoso.com" -Date 2026-09-25

.EXAMPLE
    # Put it all back
    .\Restore-MailboxMessages.ps1 -Mailbox "user@contoso.com" -Date 2026-09-25 -Apply

.EXAMPLE
    # Everything moved or deleted from 20 September until now
    .\Restore-MailboxMessages.ps1 -Mailbox "user@contoso.com" -Since 2026-09-20 -Apply

.EXAMPLE
    # Only the deletions, in a precise window - needs no Graph access at all
    .\Restore-MailboxMessages.ps1 -Mailbox "user@contoso.com" -Include Deleted `
        -After "2026-09-25 14:00" -Before "2026-09-25 16:00" -Apply

.EXAMPLE
    # Undo a Move-InboxToArchive.ps1 run: also move unaudited Archive items to the Inbox
    .\Restore-MailboxMessages.ps1 -Mailbox "user@contoso.com" -Date 2026-09-25 `
        -UnauditedArchiveToInbox -Apply

.NOTES
    Permissions

      Exchange  Search-UnifiedAuditLog needs the View-Only Audit Logs or Audit
                Logs role (Organization Management or Compliance Management).
                Get-/Restore-RecoverableItems need the "Mailbox Import Export"
                role, which no role group has by default:
                  New-ManagementRoleAssignment -Role "Mailbox Import Export" -User admin@contoso.com
                then reconnect. Without it the Deleted part runs over Graph
                instead (see above) - everything but Purges.

      Graph     App-only Mail.ReadWrite, for the Moved part and for the Deleted
                part when the role is missing. Same three
                routes as Remove-PhishingMessage.ps1: an existing app-only
                session, your own app (-ClientId), or a temporary app created
                with a device code sign-in (Global Administrator or Privileged
                Role Administrator) and removed again at the end. Those last two
                run on plain REST, so they do not clash with the Exchange
                module's MSAL.

    Audit log

      Up to 30-90 minutes (occasionally 24 hours) behind. A run on the same day
      can miss the most recent actions - re-run later if counts look short.

      By default Exchange audits the owner's MoveToDeletedItems, SoftDelete and
      HardDelete, but NOT the owner's Move. The script reports which actions
      are not audited on this mailbox. Missing moves cannot be traced back, and
      cannot be attributed. Enable it for next time with:
        Set-Mailbox user@contoso.com -AuditOwner @{Add='Move'}

      Moves done by an Inbox rule run as the mailbox itself and are often not
      audited. Moves into the online archive MAILBOX (retention policies) are
      not in scope - that is a different mailbox.

    Required module: ExchangeOnlineManagement.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]   $Mailbox,
    [datetime] $Date,
    [Alias('From', 'Since')]
    [datetime] $After,
    [datetime] $Before,
    [ValidateSet('Deleted', 'Moved')]
    [string[]] $Include = @('Deleted', 'Moved'),
    [switch]   $UnauditedArchiveToInbox,
    [switch]   $Apply,
    [string]   $OutputPath,
    [string]   $TenantId,
    [string]   $ClientId,
    [string]   $ClientSecret,
    [string]   $CertificateThumbprint
)

# ── Window ────────────────────────────────────────────────────────────────────
if ($PSBoundParameters.ContainsKey('Date')) {
    if ($PSBoundParameters.ContainsKey('After') -or $PSBoundParameters.ContainsKey('Before')) {
        throw "Use either -Date or -After/-Before, not both."
    }
    $windowStart = $Date.Date
    $windowEnd   = $Date.Date.AddDays(1)
} elseif ($PSBoundParameters.ContainsKey('After')) {
    $windowStart = $After
    $windowEnd   = if ($PSBoundParameters.ContainsKey('Before')) { $Before } else { Get-Date }
} else {
    throw "Specify -Date, or -After (and optionally -Before): the moment the messages were moved or deleted."
}
if ($windowStart -ge $windowEnd) { throw "The start of the window must be before its end." }

if ($ClientId -and -not $ClientSecret -and -not $CertificateThumbprint) {
    throw "-ClientId requires -ClientSecret or -CertificateThumbprint."
}
if ($UnauditedArchiveToInbox -and $Include -notcontains 'Moved') {
    throw "-UnauditedArchiveToInbox belongs to the Moved part; add 'Moved' to -Include."
}

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Restore Moved / Deleted Messages" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Mailbox : $Mailbox"
Write-Host "  Window  : $($windowStart.ToString('yyyy-MM-dd HH:mm')) - $($windowEnd.ToString('yyyy-MM-dd HH:mm'))"
Write-Host "  Parts   : $($Include -join ', ')" -ForegroundColor DarkGray
Write-Host ("  Mode    : {0}" -f $(if ($Apply) { 'APPLY - messages will be put back' } else { 'Preview only (no changes)' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

$results              = [System.Collections.Generic.List[PSObject]]::new()
$script:AuditEvents   = [System.Collections.Generic.List[PSObject]]::new()
$script:ConnectedExo  = $false
$script:AuditTruncated = $false

# ══════════════════════════════════════════════════════════════════════════════
#  Graph connection
#
#  Same three-way pattern as Remove-PhishingMessage.ps1: reuse an app-only
#  session, use your own app, or build a short-lived one and remove it again.
#  Routes 2 (secret) and 3 are plain REST, so they survive a loaded Exchange
#  module and its different MSAL.
# ══════════════════════════════════════════════════════════════════════════════
$script:TempAppObjectId = $null
$script:AppOnlyHeaders  = $null
$script:TokenBody       = $null
$script:TokenTenantId   = $null
$script:TokenExpiry     = [datetime]::MinValue
$script:AccessToken     = $null
$script:GraphConnected  = $false
$script:AdminHeaders    = $null
# The Graph PowerShell SDK's own public client, used for the device code sign-in.
$script:GraphCliClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'

function Remove-TempApp {
    # Needs the delegated admin token - the app-only token only holds Mail.ReadWrite.
    if (-not $script:TempAppObjectId) { return }
    Write-Host "  Removing temporary App Registration..." -ForegroundColor DarkGray
    try {
        Invoke-GraphAdmin -Method DELETE -Uri "https://graph.microsoft.com/v1.0/applications/$($script:TempAppObjectId)" | Out-Null
        Write-Host "  [OK]   Temporary App Registration removed." -ForegroundColor DarkGray
    } catch {
        Write-Warning "Could not remove the temporary App Registration (object ID $($script:TempAppObjectId)). Remove it by hand in Entra ID > App registrations."
    }
    $script:TempAppObjectId = $null
}

function Update-AppOnlyToken {
    if (-not $script:TokenBody) { return }
    if ($script:AppOnlyHeaders -and (Get-Date) -lt $script:TokenExpiry) { return }
    $resp = Invoke-RestMethod -Method POST -ErrorAction Stop -Body $script:TokenBody -Uri "https://login.microsoftonline.com/$($script:TokenTenantId)/oauth2/v2.0/token"
    $script:AccessToken    = $resp.access_token
    $script:AppOnlyHeaders = @{ Authorization = "Bearer $($resp.access_token)" }
    $script:TokenExpiry    = (Get-Date).AddSeconds($resp.expires_in - 300)
}

function Invoke-Graph {
    param(
        [string] $Method = 'GET',
        [Parameter(Mandatory)] [string] $Uri,
        $Body,
        [string] $ContentType = 'application/json'
    )

    # Throttling and gateway hiccups are routine and say nothing about the
    # mailbox; a 401 or 403 never improves by asking again.
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            if ($script:AppOnlyHeaders) {
                Update-AppOnlyToken
                $params = @{ Method = $Method; Uri = $Uri; Headers = $script:AppOnlyHeaders; ErrorAction = 'Stop' }
                if ($Body) { $params['Body'] = $Body; $params['ContentType'] = $ContentType }
                return Invoke-RestMethod @params
            }
            $params = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
            if ($Body) { $params['Body'] = $Body; $params['ContentType'] = $ContentType }
            return Invoke-MgGraphRequest @params
        } catch {
            $err  = $_
            $code = 0
            try { $code = [int]$err.Exception.Response.StatusCode } catch { $code = 0 }
            if ($code -eq 0 -and "$($err.Exception.Message)" -match '(?<![0-9])(429|503|504)(?![0-9])') { $code = [int]$Matches[1] }
            if ($attempt -ge 4 -or $code -notin @(429, 503, 504)) { throw $err }
            Start-Sleep -Seconds ([Math]::Pow(2, $attempt))
        }
    }
}

function Get-TokenRole {
    param([string] $Jwt)
    try {
        $payload = $Jwt.Split('.')[1].Replace('-', '+').Replace('_', '/')
        switch ($payload.Length % 4) { 2 { $payload += '==' } 3 { $payload += '=' } }
        $json = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($payload))
        return @(($json | ConvertFrom-Json).roles)
    } catch {
        return @()
    }
}

function Confirm-AppRole {
    <#
        A token minted before the role assignment reached the token service is
        valid but powerless, and cached for an hour - so wait for the roles and
        re-mint, rather than 403 on every call for the rest of the run.
    #>
    param([string[]] $Required, [int] $TimeoutSeconds = 180)

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $reported = $false
    while ($true) {
        $have    = Get-TokenRole -Jwt $script:AccessToken
        $missing = @($Required | Where-Object { $have -notcontains $_ })
        if ($missing.Count -eq 0) {
            Write-Host "  [OK]   Token carries: $($have -join ', ')" -ForegroundColor DarkGray
            return $true
        }
        if ((Get-Date) -ge $deadline) {
            Write-Warning "The app-only token still lacks $($missing -join ', ') after $TimeoutSeconds seconds."
            return $false
        }
        if (-not $reported) {
            Write-Host "  Waiting for the permission grant to reach the token service ($($missing -join ', '))..." -ForegroundColor DarkGray
            $reported = $true
        }
        Start-Sleep -Seconds 10
        $script:TokenExpiry    = [datetime]::MinValue
        $script:AppOnlyHeaders = $null
        try { Update-AppOnlyToken } catch { }
    }
}

function Resolve-EffectiveTenantId {
    # GDAP-aware, matching Move-InboxToArchive.ps1 / Remove-PhishingMessage.ps1.
    if ($TenantId) { return $TenantId }
    try {
        $gdap = ($global:authMode -and ([string]$global:authMode).ToUpperInvariant() -eq 'GDAP') -or
                ($env:M365_AUTH_MODE -and ([string]$env:M365_AUTH_MODE).ToUpperInvariant() -eq 'GDAP')
        if ($gdap -and $global:cid)      { return [string]$global:cid }
        if ($env:M365_CUSTOMER_TENANTID) { return [string]$env:M365_CUSTOMER_TENANTID }
    } catch {}
    return $null
}

function Get-AppOnlyTokenByRest {
    param([string] $Tenant, [string] $App, [string] $Secret)
    $script:TokenBody = @{
        grant_type    = 'client_credentials'
        scope         = 'https://graph.microsoft.com/.default'
        client_id     = $App
        client_secret = $Secret
    }
    $script:TokenTenantId = $Tenant
    Update-AppOnlyToken
}

function Get-DelegatedTokenByDeviceCode {
    # Device code flow over plain REST: a delegated token for app management
    # without loading the Graph SDK and its MSAL.
    param([string] $Tenant, [string[]] $Scopes)

    $scopeString = ((@($Scopes | ForEach-Object { "https://graph.microsoft.com/$_" })) + 'offline_access') -join ' '
    $dc = Invoke-RestMethod -Method POST -ErrorAction Stop `
            -Uri  "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/devicecode" `
            -Body @{ client_id = $script:GraphCliClientId; scope = $scopeString }

    Write-Host ""
    Write-Host "  ------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "   $($dc.message)" -ForegroundColor Yellow
    Write-Host "  ------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Waiting for sign-in..." -ForegroundColor DarkGray

    $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
    $interval = [Math]::Max(5, [int]$dc.interval)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval
        try {
            $tok = Invoke-RestMethod -Method POST -ErrorAction Stop `
                    -Uri  "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" `
                    -Body @{
                        grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                        client_id   = $script:GraphCliClientId
                        device_code = $dc.device_code
                    }
            return @{ Authorization = "Bearer $($tok.access_token)" }
        } catch {
            $code = ''
            try { $code = ($_.ErrorDetails.Message | ConvertFrom-Json).error } catch {}
            if ($code -eq 'authorization_pending') { continue }
            if ($code -eq 'slow_down')             { $interval += 5; continue }
            if ($code -eq 'expired_token')         { throw "Device code expired before sign-in completed." }
            if ($code -eq 'authorization_declined'){ throw "Sign-in was declined." }
            throw
        }
    }
    throw "Device code sign-in timed out."
}

function Invoke-GraphAdmin {
    param([string] $Method = 'GET', [Parameter(Mandatory)] [string] $Uri, $Body)
    $params = @{ Method = $Method; Uri = $Uri; Headers = $script:AdminHeaders; ErrorAction = 'Stop' }
    if ($Body) {
        $params['Body']        = ($Body | ConvertTo-Json -Depth 6)
        $params['ContentType'] = 'application/json'
    }
    return Invoke-RestMethod @params
}

function Connect-GraphForMail {
    if ($script:GraphConnected -or $script:AppOnlyHeaders) { return $true }

    # 1. An app-only session the caller already established.
    $ctx = $null
    try { $ctx = Get-MgContext -ErrorAction SilentlyContinue } catch {}
    if ($ctx -and $ctx.AuthType -eq 'AppOnly') {
        Write-Host "  [OK]   Using the existing app-only Graph session." -ForegroundColor DarkGray
        $script:GraphConnected = $true
        return $true
    }

    $effectiveTenantId = Resolve-EffectiveTenantId
    if (-not $effectiveTenantId) {
        Write-Warning "-TenantId is required for app-only Graph access (or a resolvable GDAP customer tenant)."
        return $false
    }

    # 2. Your own app registration.
    if ($ClientId) {
        if ($ClientSecret) {
            try {
                Get-AppOnlyTokenByRest -Tenant $effectiveTenantId -App $ClientId -Secret $ClientSecret
            } catch {
                Write-Warning "Could not obtain a token for the supplied app: $($_.Exception.Message)"
                return $false
            }
            Write-Host "  [OK]   App-only token obtained (supplied app, REST)." -ForegroundColor DarkGray
            if ((Get-TokenRole -Jwt $script:AccessToken) -notcontains 'Mail.ReadWrite') {
                Write-Warning "App $ClientId has no Mail.ReadWrite application permission. Every Graph call will return 403 until that is granted and admin-consented."
            }
            return $true
        }
        try {
            Connect-MgGraph -ClientId $ClientId -TenantId $effectiveTenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
        } catch {
            Write-Warning "Could not connect with the supplied certificate: $($_.Exception.Message)"
            return $false
        }
        Write-Host "  [OK]   Connected with the supplied certificate." -ForegroundColor DarkGray
        $script:GraphConnected = $true
        return $true
    }

    # 3. A short-lived app, created and removed over REST.
    try {
        Write-Host "  No app-only Graph access yet - setting up a temporary App Registration." -ForegroundColor Cyan
        Write-Host "  Required role: Global Administrator or Privileged Role Administrator (one-time)" -ForegroundColor DarkGray

        $script:AdminHeaders = Get-DelegatedTokenByDeviceCode -Tenant $effectiveTenantId `
                                    -Scopes @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
        Write-Host "  [OK]   Signed in." -ForegroundColor DarkGray

        $appName = "MailRestore-Temp-$(Get-Date -Format 'yyyyMMddHHmmss')"
        Write-Host "  Creating temporary App Registration '$appName'..." -ForegroundColor Cyan
        $app = Invoke-GraphAdmin -Method POST -Uri 'https://graph.microsoft.com/v1.0/applications' `
                    -Body @{ displayName = $appName; signInAudience = 'AzureADMyOrg' }
        $script:TempAppObjectId = $app.id

        $sp = Invoke-GraphAdmin -Method POST -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Body @{ appId = $app.appId }
        $graphSp = @((Invoke-GraphAdmin -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'").value)[0]
        if (-not $graphSp) { throw "Could not resolve the Microsoft Graph service principal." }

        $appRole = @($graphSp.appRoles | Where-Object { $_.value -eq 'Mail.ReadWrite' -and $_.allowedMemberTypes -contains 'Application' })[0]
        if (-not $appRole) { throw "Could not resolve the Mail.ReadWrite application role." }
        Invoke-GraphAdmin -Method POST -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)/appRoleAssignments" `
            -Body @{ principalId = $sp.id; resourceId = $graphSp.id; appRoleId = $appRole.id } | Out-Null
        Write-Host "  [OK]   Mail.ReadWrite (application) granted." -ForegroundColor DarkGray

        $secret = Invoke-GraphAdmin -Method POST -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)/addPassword" `
                    -Body @{ passwordCredential = @{ displayName = 'temp'; endDateTime = (Get-Date).AddHours(2).ToString('o') } }

        $ok = $false
        for ($i = 1; $i -le 8; $i++) {
            try {
                Get-AppOnlyTokenByRest -Tenant $effectiveTenantId -App $app.appId -Secret $secret.secretText
                $ok = $true
                break
            } catch {
                if ($i -lt 8) {
                    Write-Host "  Waiting for the app registration to propagate (attempt $i/8)..." -ForegroundColor DarkGray
                    Start-Sleep -Seconds 5
                }
            }
        }
        if (-not $ok) { throw "Could not obtain an app-only token after propagation retries." }
        Write-Host "  [OK]   App-only token obtained (temporary app)." -ForegroundColor DarkGray

        if (-not (Confirm-AppRole -Required @('Mail.ReadWrite'))) {
            Remove-TempApp
            return $false
        }
        return $true
    } catch {
        Write-Warning "Temporary app setup failed: $($_.Exception.Message)"
        Remove-TempApp
        return $false
    }
}

function Get-GraphPaged {
    param([string] $Uri, [int] $MaxItems = 100000)
    $items = [System.Collections.Generic.List[PSObject]]::new()
    $next  = $Uri
    while ($next -and $items.Count -lt $MaxItems) {
        $resp = Invoke-Graph -Method GET -Uri $next
        foreach ($v in @($resp.value)) { $items.Add($v) }
        $next = $resp.'@odata.nextLink'
    }
    return $items
}

function Invoke-GraphBatch {
    <#
        Sends requests through $batch, 20 at a time, retrying the ones Graph
        throttled. Returns a hashtable of request id -> response. A request that
        never got an answer comes back with status 0 and the error in body.
    #>
    param([object[]] $Requests)

    $responses = @{}
    $pending   = @($Requests)
    $lastError = ''
    for ($pass = 1; $pass -le 5 -and $pending.Count -gt 0; $pass++) {
        $retry = [System.Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $pending.Count; $i += 20) {
            $chunk = @($pending[$i..([Math]::Min($i + 19, $pending.Count - 1))])
            $body = @{
                requests = @($chunk | ForEach-Object {
                    $r = @{ id = $_.id; method = $_.method; url = $_.url }
                    if ($_.body) { $r['body'] = $_.body; $r['headers'] = @{ 'Content-Type' = 'application/json' } }
                    $r
                })
            } | ConvertTo-Json -Depth 8

            try {
                $resp = Invoke-Graph -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' -Body $body
            } catch {
                $lastError = $_.Exception.Message
                foreach ($c in $chunk) { $retry.Add($c) }
                continue
            }

            $byId = @{}
            foreach ($c in $chunk) { $byId[[string]$c.id] = $c }
            foreach ($r in @($resp.responses)) {
                if ([int]$r.status -in @(429, 503, 504)) { $retry.Add($byId[[string]$r.id]) }
                else { $responses[[string]$r.id] = $r }
            }
        }
        $pending = @($retry)
        if ($pending.Count -gt 0 -and $pass -lt 5) { Start-Sleep -Seconds ([Math]::Min(30, [Math]::Pow(2, $pass))) }
    }
    foreach ($p in $pending) {
        $responses[[string]$p.id] = [PSCustomObject]@{ id = $p.id; status = 0; body = [PSCustomObject]@{ error = [PSCustomObject]@{ message = "No answer after retries. $lastError" } } }
    }
    return $responses
}

function Get-BatchError {
    param($Response)
    $m = ''
    try { $m = $Response.body.error.message } catch {}
    return "HTTP $($Response.status)$(if ($m) { ": $m" })"
}

# ══════════════════════════════════════════════════════════════════════════════
#  Audit log - who did what
# ══════════════════════════════════════════════════════════════════════════════
function Get-LogonTypeName {
    param($Value)
    switch ("$Value") {
        '0' { 'Owner' }
        '1' { 'Admin' }
        '2' { 'Delegate' }
        '3' { 'Transport' }
        '4' { 'SystemService' }
        '6' { 'DelegatedAdmin' }
        default { "$Value" }
    }
}

function Get-ClientLabel {
    # ClientInfoString is the most readable clue to HOW it was done.
    param([string] $ClientInfo, [string] $AppId)
    $label = switch -Regex ($ClientInfo) {
        'Client=OWA'                    { 'Outlook on the web'; break }
        'Client=MSExchangeRPC'          { 'Outlook desktop'; break }
        'OutlookMobile|ActiveSync|AirSync' { 'Mobile'; break }
        'Client=REST'                   { 'Graph / REST app'; break }
        'Client=WebServices'            { 'EWS app'; break }
        'Client=POP3|Client=IMAP4'      { 'POP/IMAP client'; break }
        default {
            if ($ClientInfo) { ($ClientInfo -split ';')[0] } else { 'unknown' }
        }
    }
    if ($AppId) { $label += " (app $AppId)" }
    return $label
}

function Get-AuditEvents {
    <#
        Pulls every Move / delete in the window from the Unified Audit Log and
        keeps this mailbox's. The search runs tenant-wide because -UserIds
        filters on the actor, and the actor is exactly what we do not know yet.

        One search session stops at 50,000 records, which a tenant-wide search
        over weeks easily exceeds - so the window is searched one day at a time,
        and only records mentioning this mailbox are kept in memory.
    #>
    param([string] $MailboxGuid, [string] $MailboxUpn, [string] $MailboxSmtp)

    $ops = @('Move', 'MoveToDeletedItems', 'SoftDelete', 'HardDelete')
    $raw = [System.Collections.Generic.List[object]]::new()
    $ids = @($MailboxGuid, $MailboxUpn, $MailboxSmtp) | Where-Object { $_ } | ForEach-Object { "$_".ToLowerInvariant() }
    # Cheap text pre-filter before the JSON is parsed properly below.
    $mention = ($ids | ForEach-Object { [regex]::Escape($_) }) -join '|'

    $days = [Math]::Ceiling(($windowEnd - $windowStart).TotalDays)
    Write-Host "  Searching the audit log ($($ops -join ', ')), $days day(s)..." -ForegroundColor DarkGray

    $sliceStart = $windowStart
    while ($sliceStart -lt $windowEnd) {
        $sliceEnd = $sliceStart.AddDays(1)
        if ($sliceEnd -gt $windowEnd) { $sliceEnd = $windowEnd }
        if ($days -gt 1) { Write-Progress -Activity 'Searching the audit log' -Status $sliceStart.ToString('yyyy-MM-dd') -PercentComplete ([int](($sliceStart - $windowStart).TotalDays / $days * 100)) }

        $sid   = "MailRestore_$([guid]::NewGuid())"
        $read  = 0
        $total = 0
        for ($page = 1; $page -le 12; $page++) {
            $batch = @(Search-UnifiedAuditLog -StartDate $sliceStart.ToUniversalTime() -EndDate $sliceEnd.ToUniversalTime() `
                        -Operations $ops -SessionId $sid -SessionCommand ReturnLargeSet -ResultSize 5000 -ErrorAction Stop)
            if ($batch.Count -eq 0) { break }
            if ($total -eq 0) { $total = [int]$batch[0].ResultCount }
            $read += $batch.Count
            foreach ($rec in $batch) { if ("$($rec.AuditData)" -match $mention) { $raw.Add($rec) } }
            if ($read -ge $total) { break }
        }
        if ($total -gt 50000) {
            # ReturnLargeSet stops at 50,000 records per session.
            $script:AuditTruncated = $true
            Write-Warning "$($sliceStart.ToString('yyyy-MM-dd')): the audit log holds $total records that day; only 50,000 can be read in one search, so some actions on this mailbox may be missing."
        }
        $sliceStart = $sliceEnd
    }
    if ($days -gt 1) { Write-Progress -Activity 'Searching the audit log' -Completed }

    foreach ($rec in $raw) {
        try { $d = $rec.AuditData | ConvertFrom-Json } catch { continue }
        $owner = @("$($d.MailboxGuid)", "$($d.MailboxOwnerUPN)") | ForEach-Object { $_.ToLowerInvariant() }
        if (-not ($owner | Where-Object { $ids -contains $_ })) { continue }
        if ($d.CrossMailboxOperation -eq $true) { continue }

        $time = ([datetime]$d.CreationTime)
        if ($time.Kind -ne [DateTimeKind]::Utc) { $time = [datetime]::SpecifyKind($time, [DateTimeKind]::Utc) }
        $time = $time.ToLocalTime()

        $actor = if ($d.UserId) { "$($d.UserId)" } else { "$($rec.UserIds)" }
        $appId = if ($d.AppId) { "$($d.AppId)" } elseif ($d.ClientAppId) { "$($d.ClientAppId)" } else { '' }

        foreach ($item in @($d.AffectedItems)) {
            if (-not $item) { continue }
            $source = if ($item.ParentFolder -and $item.ParentFolder.Path) { "$($item.ParentFolder.Path)" } elseif ($d.Folder) { "$($d.Folder.Path)" } else { '' }
            $script:AuditEvents.Add([PSCustomObject]@{
                Time              = $time
                Operation         = "$($d.Operation)"
                Actor             = $actor
                LogonType         = Get-LogonTypeName $d.LogonType
                Client            = Get-ClientLabel -ClientInfo "$($d.ClientInfoString)" -AppId $appId
                ClientIP          = "$($d.ClientIPAddress)$(if (-not $d.ClientIPAddress) { $d.ClientIP })"
                Subject           = "$($item.Subject)"
                InternetMessageId = "$($item.InternetMessageId)"
                SourcePath        = $source
                DestPath          = if ($d.DestFolder) { "$($d.DestFolder.Path)" } else { '' }
            })
        }
    }
    Write-Host "  $($script:AuditEvents.Count) audited action(s) on this mailbox in the window." -ForegroundColor DarkGray
}

function Find-ActionEvent {
    <#
        Recoverable items carry no Internet MessageId, so the deleting action is
        matched on subject and the closest time. Good enough to name a person;
        the CSV says it is a match by subject.
    #>
    param([string] $Subject, [datetime] $Time, [string[]] $Operations)
    $cands = @($script:AuditEvents | Where-Object { $Operations -contains $_.Operation -and $_.Subject -eq $Subject })
    if ($cands.Count -eq 0) { return $null }
    return $cands | Sort-Object { [Math]::Abs(($_.Time - $Time).TotalSeconds) } | Select-Object -First 1
}

function New-ResultRow {
    param($Phase, $Operation, $ActionTime, $Subject, $Received, $InternetMessageId,
          $CurrentFolder, $TargetFolder, $AuditEvent, $Status, $Detail)
    return [PSCustomObject]@{
        Phase             = $Phase
        Operation         = $Operation
        ActionTime        = if ($ActionTime) { ([datetime]$ActionTime).ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        Actor             = if ($AuditEvent) { $AuditEvent.Actor } else { '' }
        LogonType         = if ($AuditEvent) { $AuditEvent.LogonType } else { '' }
        Client            = if ($AuditEvent) { $AuditEvent.Client } else { '' }
        ClientIP          = if ($AuditEvent) { $AuditEvent.ClientIP } else { '' }
        Subject           = $Subject
        Received          = $Received
        InternetMessageId = $InternetMessageId
        CurrentFolder     = $CurrentFolder
        TargetFolder      = $TargetFolder
        Status            = $Status
        Detail            = $Detail
    }
}

function Write-ResultRow {
    param($Row)
    $color = switch ($Row.Status) {
        'Restored'       { 'Green' }
        'WouldRestore'   { 'Yellow' }
        'Error'          { 'Red' }
        'NotRestored'    { 'Red' }
        default          { 'DarkGray' }
    }
    $who = if ($Row.Actor) { $Row.Actor } else { 'no audit record' }
    Write-Host ("      [{0}] {1} | {2} -> {3} | {4} | {5}" -f $Row.Status, $Row.ActionTime, $Row.CurrentFolder, $Row.TargetFolder, $who,
                    ("$($Row.Subject)" -replace '^(.{45}).+$', '$1...')) -ForegroundColor $color
}

# ══════════════════════════════════════════════════════════════════════════════
#  Deleted - Deleted Items / Recoverable Items / Purges
# ══════════════════════════════════════════════════════════════════════════════
function Test-MessageClass {
    param([string] $ItemClass)
    if (-not $ItemClass) { return $true }
    return ($ItemClass -like 'IPM.Note*' -or $ItemClass -like 'IPM.Schedule.Meeting*' -or
            $ItemClass -like 'REPORT.IPM.Note*' -or $ItemClass -like 'IPM.Post*')
}

function Invoke-DeletedRestore {
    param([string] $Identity)

    # Which audit operation put an item in each place.
    $opsFor = @{
        DeletedItems     = @('MoveToDeletedItems', 'Move')
        RecoverableItems = @('SoftDelete')
        PurgedItems      = @('HardDelete')
    }

    foreach ($src in @('DeletedItems', 'RecoverableItems', 'PurgedItems')) {
        try {
            # LastModifiedTime of a deleted item is the moment it was deleted,
            # which is what these filters look at.
            $items = @(Get-RecoverableItems -Identity $Identity -SourceFolder $src `
                        -FilterStartTime $windowStart -FilterEndTime $windowEnd -ResultSize Unlimited -ErrorAction Stop)
        } catch {
            # Purges is only readable under a hold; its absence is not a failure.
            if ($src -ne 'PurgedItems') { Write-Warning "Could not read $src : $($_.Exception.Message)" }
            continue
        }
        $items = @($items | Where-Object { Test-MessageClass "$($_.ItemClass)" })
        Write-Host ("  {0,-17} {1} message(s) deleted in the window" -f $src, $items.Count) -ForegroundColor DarkGray
        if ($items.Count -eq 0) { continue }

        $rows = [System.Collections.Generic.List[PSObject]]::new()
        foreach ($it in $items) {
            $ev = Find-ActionEvent -Subject "$($it.Subject)" -Time ([datetime]$it.LastModifiedTime) -Operations $opsFor[$src]
            $target = if ("$($it.LastParentPath)") { "$($it.LastParentPath)" } else { '(original folder unknown)' }
            if ("$($it.OriginalFolderExists)" -eq 'False') { $target += ' (folder no longer exists)' }

            $row = New-ResultRow -Phase 'Deleted' -Operation $src -ActionTime $it.LastModifiedTime `
                        -Subject "$($it.Subject)" -Received '' -InternetMessageId '' `
                        -CurrentFolder $src -TargetFolder $target -AuditEvent $ev `
                        -Status $(if ($Apply) { 'Pending' } else { 'WouldRestore' }) `
                        -Detail $(if ($ev) { 'Actor matched on subject and time' } else { '' })
            Add-Member -InputObject $row -NotePropertyName EntryID -NotePropertyValue "$($it.EntryID)"

            if ($Apply) {
                try {
                    # Exchange puts the item back in LastParentPath itself.
                    Restore-RecoverableItems -Identity $Identity -SourceFolder $src -EntryID $it.EntryID -ErrorAction Stop | Out-Null
                    $row.Status = 'Restored'
                } catch {
                    $row.Status = 'Error'
                    $row.Detail = $_.Exception.Message
                }
            }
            $rows.Add($row)
        }

        if ($Apply) {
            # Verify rather than trust the cmdlet's silence: anything still in
            # this folder after the restore did not go back.
            $left = @()
            try {
                $left = @(Get-RecoverableItems -Identity $Identity -SourceFolder $src `
                            -FilterStartTime $windowStart -FilterEndTime $windowEnd -ResultSize Unlimited -ErrorAction Stop |
                            ForEach-Object { "$($_.EntryID)" })
            } catch {}
            foreach ($row in $rows) {
                if ($row.Status -eq 'Restored' -and $left -contains $row.EntryID) {
                    $row.Status = 'NotRestored'
                    $row.Detail = "Still in $src after the restore"
                }
            }
        }

        foreach ($row in $rows) {
            $row.PSObject.Properties.Remove('EntryID')
            Write-ResultRow $row
            $results.Add($row)
        }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
#  Moved - traced through the audit log, put back over Graph
# ══════════════════════════════════════════════════════════════════════════════
function ConvertTo-FolderKey {
    param([string] $Path)
    return "$Path".Trim().Trim('\').ToLowerInvariant()
}

function Get-FolderMap {
    <#
        Folder path -> id for the whole mailbox. The audit log names folders by
        path ("\Inbox\Projects"), in the mailbox's own language, which is the
        same displayName Graph returns - so walking the tree lines the two up.
    #>
    param([string] $Base)

    $map = @{ PathToId = @{}; IdToPath = @{} }
    $queue = [System.Collections.Generic.Queue[object]]::new()
    $queue.Enqueue(@{ Uri = "$Base/mailFolders?includeHiddenFolders=true&`$top=100&`$select=id,displayName,childFolderCount"; Prefix = '' })
    while ($queue.Count -gt 0) {
        $q = $queue.Dequeue()
        foreach ($f in @(Get-GraphPaged -Uri $q.Uri)) {
            $path = if ($q.Prefix) { "$($q.Prefix)\$($f.displayName)" } else { "$($f.displayName)" }
            $key  = ConvertTo-FolderKey $path
            if (-not $map.PathToId.ContainsKey($key)) { $map.PathToId[$key] = $f.id }
            $map.IdToPath[$f.id] = $path
            if ([int]$f.childFolderCount -gt 0) {
                $queue.Enqueue(@{ Uri = "$Base/mailFolders/$($f.id)/childFolders?includeHiddenFolders=true&`$top=100&`$select=id,displayName,childFolderCount"; Prefix = $path })
            }
        }
    }
    return $map
}

function Get-WellKnownFolderId {
    param([string] $Base, [string] $Name)
    try { return (Invoke-Graph -Uri "$Base/mailFolders/$Name`?`$select=id").id } catch { return $null }
}

$script:FolderCtx = $null
function Get-FolderContext {
    # Read once per run: both the deleted and the moved part need it.
    param([string] $GraphMailbox)
    if ($script:FolderCtx) { return $script:FolderCtx }

    $base = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($GraphMailbox))"
    Write-Host "  Reading the folder tree..." -ForegroundColor DarkGray
    $map = Get-FolderMap -Base $base
    $ctx = @{
        Base      = $base
        Rel       = "/users/$([uri]::EscapeDataString($GraphMailbox))"
        Map       = $map
        InboxId   = Get-WellKnownFolderId -Base $base -Name 'inbox'
        ArchiveId = Get-WellKnownFolderId -Base $base -Name 'archive'
        DeletedId = Get-WellKnownFolderId -Base $base -Name 'deleteditems'
        RecovId   = Get-WellKnownFolderId -Base $base -Name 'recoverableitemsdeletions'
    }
    # Recoverable Items sits outside the folder tree Graph lists; name it so
    # reports show where a message was found.
    if ($ctx.RecovId) { $map.IdToPath[$ctx.RecovId] = 'Recoverable Items\Deletions' }
    $ctx.DeletedKey = if ($ctx.DeletedId -and $map.IdToPath.ContainsKey($ctx.DeletedId)) { ConvertTo-FolderKey $map.IdToPath[$ctx.DeletedId] } else { '' }
    $ctx.InboxPath  = if ($ctx.InboxId -and $map.IdToPath.ContainsKey($ctx.InboxId)) { $map.IdToPath[$ctx.InboxId] } else { 'Inbox' }
    $script:FolderCtx = $ctx
    return $ctx
}

function Get-FolderLabel {
    param($Ctx, [string] $FolderId)
    if ($FolderId -and $Ctx.Map.IdToPath.ContainsKey($FolderId)) { return $Ctx.Map.IdToPath[$FolderId] }
    return $FolderId
}

function Resolve-RestoreTarget {
    <#
        Folder id to put a message back in, from the path the audit log gave.
        Falls back to the Inbox when the folder is gone - and, for a deleted
        message, when the folder it was deleted FROM was Deleted Items itself,
        since putting it back there is not recovering it.
    #>
    param($Ctx, [string] $SourcePath, [switch] $NotDeletedItems)

    $key = ConvertTo-FolderKey $SourcePath
    $id  = if ($key -and $Ctx.Map.PathToId.ContainsKey($key)) { $Ctx.Map.PathToId[$key] } else { $null }
    if ($id -and $NotDeletedItems -and ($id -eq $Ctx.DeletedId -or $id -eq $Ctx.RecovId)) {
        return @{ Id = $Ctx.InboxId; Path = $Ctx.InboxPath; Detail = 'Was deleted from Deleted Items; restored to the Inbox' }
    }
    if ($id) { return @{ Id = $id; Path = (Get-FolderLabel $Ctx $id); Detail = '' } }
    $shown = "$SourcePath".Trim().Trim('\')
    return @{ Id = $Ctx.InboxId; Path = $Ctx.InboxPath; Detail = "Original folder '$shown' not found; restored to the Inbox" }
}

function Get-AuditPlans {
    <#
        One plan per message: the first audited action of the window on it, so
        a message moved A -> B -> C (or moved and then deleted) goes back to A.
    #>
    param([string[]] $Operations)
    $plans = @{}
    $events = @($script:AuditEvents | Where-Object { $Operations -contains $_.Operation -and $_.InternetMessageId })
    foreach ($g in @($events | Group-Object InternetMessageId)) {
        $ordered = @($g.Group | Sort-Object Time)
        $plans[$g.Name] = [PSCustomObject]@{ First = $ordered[0]; Last = $ordered[-1] }
    }
    return $plans
}

function Find-MessagesById {
    <#
        Locates messages by Internet MessageId, in the mailbox and in
        Recoverable Items\Deletions - /messages does not reach the latter.
        Returns key -> @{ Hits; Error }.
    #>
    param($Ctx, [hashtable] $IdsByKey)

    $requests = [System.Collections.Generic.List[object]]::new()
    foreach ($k in $IdsByKey.Keys) {
        $filter = [uri]::EscapeDataString("internetMessageId eq '$($IdsByKey[$k].Replace("'", "''"))'")
        $sel    = '$select=id,parentFolderId,subject,receivedDateTime'
        $requests.Add(@{ id = "$k|m"; method = 'GET'; url = "$($Ctx.Rel)/messages?`$filter=$filter&$sel" })
        if ($Ctx.RecovId) {
            $requests.Add(@{ id = "$k|r"; method = 'GET'; url = "$($Ctx.Rel)/mailFolders/recoverableitemsdeletions/messages?`$filter=$filter&$sel" })
        }
    }
    $out = @{}
    if ($requests.Count -eq 0) { return $out }
    $resp = Invoke-GraphBatch -Requests $requests.ToArray()
    foreach ($k in $IdsByKey.Keys) {
        $hits = [System.Collections.Generic.List[object]]::new()
        $err  = ''
        foreach ($suffix in @('m', 'r')) {
            $r = $resp["$k|$suffix"]
            if (-not $r) { continue }
            if ([int]$r.status -eq 200) {
                foreach ($v in @($r.body.value)) { if ($v) { $hits.Add($v) } }
            } elseif ($suffix -eq 'm') {
                $err = Get-BatchError $r
            }
        }
        $out[$k] = @{ Hits = @($hits); Error = $err }
    }
    return $out
}

function Set-RestoreAction {
    # Decides what happens to one located message and queues the move.
    param($Work, [string] $Key, $Row, $Msg, [string] $TargetId)

    $Row.CurrentFolder = Get-FolderLabel $Work.Ctx $Msg.parentFolderId
    if ($Msg.receivedDateTime -and -not $Row.Received) { $Row.Received = ([datetime]$Msg.receivedDateTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm') }
    if (-not $Row.Subject) { $Row.Subject = $Msg.subject }

    if (-not $TargetId) {
        $Row.Status = 'Error'
        $Row.Detail = 'No target folder could be resolved'
    } elseif ($Msg.parentFolderId -eq $TargetId) {
        $Row.Status = 'AlreadyInPlace'
    } elseif ($Apply) {
        $Work.Requests.Add(@{ id = $Key; method = 'POST'; url = "$($Work.Ctx.Rel)/messages/$($Msg.id)/move"; body = @{ destinationId = $TargetId } })
        $Row.Status = 'Pending'
    } else {
        $Row.Status = 'WouldRestore'
    }
}

function Add-AuditedRestore {
    <#
        Puts back every message an audited action of the given kind touched,
        wherever it is now. Messages already handled (by key in $Skip) are left
        to whoever handled them.
    #>
    param($Work, [string] $Phase, [string[]] $Operations, [hashtable] $Skip = @{}, [switch] $SkipRestores, [switch] $NotDeletedItems)

    $ctx   = $Work.Ctx
    $plans = Get-AuditPlans -Operations $Operations
    $ids   = @{}
    $skippedRestores = 0
    $n = 0
    foreach ($imid in @($plans.Keys)) {
        if ($Skip.ContainsKey($imid)) { continue }
        $srcKey = ConvertTo-FolderKey $plans[$imid].First.SourcePath
        # A move OUT of Deleted Items or Recoverable Items was a restore.
        # Reversing it would delete the message again.
        if ($SkipRestores -and ($srcKey -like 'recoverable items*' -or ($ctx.DeletedKey -and $srcKey -eq $ctx.DeletedKey))) {
            $skippedRestores++
            continue
        }
        $ids["$Phase$n"] = $imid
        $n++
    }
    if ($skippedRestores -gt 0) {
        Write-Host "  $skippedRestores move(s) out of Deleted/Recoverable Items left alone - those were restores." -ForegroundColor DarkGray
    }
    if ($ids.Count -eq 0) { return }
    Write-Host "  $($ids.Count) audited message(s) to trace back." -ForegroundColor DarkGray

    $found = Find-MessagesById -Ctx $ctx -IdsByKey $ids
    foreach ($k in @($ids.Keys | Sort-Object { [int]($_ -replace '\D', '') })) {
        $imid   = $ids[$k]
        $p      = $plans[$imid]
        $target = Resolve-RestoreTarget -Ctx $ctx -SourcePath $p.First.SourcePath -NotDeletedItems:$NotDeletedItems
        $row = New-ResultRow -Phase $Phase -Operation $p.First.Operation -ActionTime $p.First.Time -Subject $p.First.Subject `
                    -Received '' -InternetMessageId $imid -CurrentFolder '' -TargetFolder $target.Path `
                    -AuditEvent $p.First -Status '' -Detail $target.Detail

        $f    = $found[$k]
        $hits = @($f.Hits)
        if ($hits.Count -gt 1) {
            # The same MessageId twice (e.g. a copy). Prefer the one where the
            # window's last action put it, then one in Deleted/Recoverable Items.
            $destKey = ConvertTo-FolderKey $p.Last.DestPath
            $pick = @($hits | Where-Object { $destKey -and (ConvertTo-FolderKey (Get-FolderLabel $ctx $_.parentFolderId)) -eq $destKey })
            if ($pick.Count -eq 0) { $pick = @($hits | Where-Object { $_.parentFolderId -in @($ctx.DeletedId, $ctx.RecovId) }) }
            if ($pick.Count -gt 0) { $hits = $pick }
        }

        if ($hits.Count -eq 0) {
            $row.Status = if ($f.Error) { 'Error' } else { 'NotFound' }
            $row.Detail = if ($f.Error) { "Lookup failed - $($f.Error)" } else { 'Not in the mailbox or Recoverable Items any more' }
        } elseif ($hits.Count -gt 1) {
            $row.Status = 'Ambiguous'
            $row.Detail = 'Several copies of this message; left alone'
        } else {
            Set-RestoreAction -Work $Work -Key $k -Row $row -Msg $hits[0] -TargetId $target.Id
        }
        $Work.Handled[$imid] = $true
        $Work.Rows[$k] = $row
    }
}

function Add-FolderWindowRestore {
    <#
        Everything in one folder that changed in the window and was not handled
        through the audit log. For Deleted Items and Recoverable Items that
        change is the deletion, so these go back too - to the Inbox, since
        without an audit record the original folder is unknown.
    #>
    param($Work, [string] $Phase, [string] $WellKnown, [string] $Label, [switch] $ListOnly, [string] $ListOnlyDetail)

    $ctx = $Work.Ctx
    $from = $windowStart.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $to   = $windowEnd.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $filter = [uri]::EscapeDataString("lastModifiedDateTime ge $from and lastModifiedDateTime lt $to")
    $msgs = @()
    try {
        $msgs = @(Get-GraphPaged -Uri "$($ctx.Base)/mailFolders/$WellKnown/messages?`$filter=$filter&`$select=id,internetMessageId,subject,receivedDateTime,lastModifiedDateTime,parentFolderId&`$top=500" |
                    Where-Object { -not $Work.Handled.ContainsKey("$($_.internetMessageId)") })
    } catch {
        Write-Warning "Could not read $Label : $($_.Exception.Message)"
        return
    }
    Write-Host ("  {0,-28} {1} message(s) without an audit record changed in the window" -f $Label, $msgs.Count) -ForegroundColor DarkGray

    for ($j = 0; $j -lt $msgs.Count; $j++) {
        $m   = $msgs[$j]
        $key = "$Phase$WellKnown$j"
        $row = New-ResultRow -Phase $Phase -Operation "$Label (no audit record)" -ActionTime ([datetime]$m.lastModifiedDateTime).ToLocalTime() `
                    -Subject $m.subject -Received '' -InternetMessageId $m.internetMessageId -CurrentFolder '' `
                    -TargetFolder $ctx.InboxPath -AuditEvent $null -Status '' -Detail 'No audit record: original folder unknown, restored to the Inbox'
        if ($ListOnly) {
            $row.CurrentFolder = $Label
            if ($m.receivedDateTime) { $row.Received = ([datetime]$m.receivedDateTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm') }
            $row.Status = 'NotAudited'
            $row.Detail = $ListOnlyDetail
        } else {
            Set-RestoreAction -Work $Work -Key $key -Row $row -Msg $m -TargetId $ctx.InboxId
        }
        if ($m.internetMessageId) { $Work.Handled["$($m.internetMessageId)"] = $true }
        $Work.Rows[$key] = $row
    }
}

function Complete-GraphRestore {
    # Sends the queued moves, then reports every row of this part.
    param($Work)
    if ($Work.Requests.Count -gt 0) {
        Write-Host "  Moving $($Work.Requests.Count) message(s) back..." -ForegroundColor Cyan
        $moved = Invoke-GraphBatch -Requests $Work.Requests.ToArray()
        foreach ($req in $Work.Requests) {
            $r   = $moved[$req.id]
            $row = $Work.Rows[$req.id]
            if ($r -and [int]$r.status -in @(200, 201)) {
                $row.Status = 'Restored'
            } else {
                $row.Status = 'Error'
                $row.Detail = "Move failed - $(Get-BatchError $r)"
            }
        }
    }
    foreach ($k in $Work.Rows.Keys) {
        Write-ResultRow $Work.Rows[$k]
        $results.Add($Work.Rows[$k])
    }
}

function New-RestoreWork {
    param([string] $GraphMailbox)
    return @{
        Ctx      = Get-FolderContext -GraphMailbox $GraphMailbox
        Rows     = [ordered]@{}
        Requests = [System.Collections.Generic.List[object]]::new()
        Handled  = @{}
    }
}

function Invoke-DeletedRestoreViaGraph {
    <#
        The Deleted part without the Mailbox Import Export role. Audited
        deletions go back to the folder the audit log says they left; everything
        else deleted in the window - found in Deleted Items and Recoverable
        Items by the moment it changed - goes to the Inbox. Only Purges
        (hard-deleted, kept under a hold) is out of Graph's reach.
    #>
    param([string] $GraphMailbox)

    Write-Host "  Restoring deleted messages over Graph instead (Deleted Items + Recoverable Items)." -ForegroundColor DarkGray
    $work = New-RestoreWork -GraphMailbox $GraphMailbox

    Add-AuditedRestore -Work $work -Phase 'Deleted' -Operations @('MoveToDeletedItems', 'SoftDelete') -NotDeletedItems
    Add-FolderWindowRestore -Work $work -Phase 'Deleted' -WellKnown 'deleteditems' -Label 'Deleted Items'
    if ($work.Ctx.RecovId) {
        Add-FolderWindowRestore -Work $work -Phase 'Deleted' -WellKnown 'recoverableitemsdeletions' -Label 'Recoverable Items\Deletions'
    }

    $hard = @($script:AuditEvents | Where-Object { $_.Operation -eq 'HardDelete' })
    foreach ($h in $hard) {
        $row = New-ResultRow -Phase 'Deleted' -Operation 'HardDelete' -ActionTime $h.Time -Subject $h.Subject -Received '' `
                    -InternetMessageId $h.InternetMessageId -CurrentFolder 'Recoverable Items\Purges' -TargetFolder '' `
                    -AuditEvent $h -Status 'Unreachable' -Detail "Hard-deleted: only Restore-RecoverableItems reaches Purges (needs the Mailbox Import Export role, and a hold for it to still be there)"
        $work.Rows["DeletedHard$($work.Rows.Count)"] = $row
    }

    Complete-GraphRestore -Work $work
}

function Invoke-MovedRestore {
    param([string] $GraphMailbox)

    Write-Host ""
    Write-Host "  -- Moved messages --------------------------------" -ForegroundColor Cyan
    $work = New-RestoreWork -GraphMailbox $GraphMailbox

    Add-AuditedRestore -Work $work -Phase 'Moved' -Operations @('Move') -SkipRestores

    # Archive without an audit record: listed, and moved only on request - a
    # message that was merely read or flagged in Archive changed as well.
    if ($work.Ctx.ArchiveId) {
        $archiveLabel = Get-FolderLabel $work.Ctx $work.Ctx.ArchiveId
        if ($UnauditedArchiveToInbox) {
            Add-FolderWindowRestore -Work $work -Phase 'Unaudited' -WellKnown 'archive' -Label $archiveLabel
        } else {
            Add-FolderWindowRestore -Work $work -Phase 'Unaudited' -WellKnown 'archive' -Label $archiveLabel -ListOnly `
                -ListOnlyDetail 'No audit record: changed in Archive in the window (moved, read or flagged). Add -UnauditedArchiveToInbox to move it'
        }
    }

    Complete-GraphRestore -Work $work
}


# ══════════════════════════════════════════════════════════════════════════════
#  Run
# ══════════════════════════════════════════════════════════════════════════════
try {
    # Graph first, so its MSAL (if the SDK is used at all) loads before Exchange's.
    $graphOk = $false
    if ($Include -contains 'Moved') {
        Write-Host "  Preparing Graph access (for the moved messages)..." -ForegroundColor DarkGray
        $graphOk = Connect-GraphForMail
        if (-not $graphOk) {
            Write-Warning "No app-only Graph access - the moved messages are skipped. The audit report and the deleted messages still run."
        }
        Write-Host ""
    }

    $exoSession = $null
    try { $exoSession = Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' } } catch {}
    if (-not $exoSession) {
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            throw "Missing required module: ExchangeOnlineManagement. Install with: .\scripts\Startup\Install-Modules.ps1"
        }
        Write-Host "  Connecting to Exchange Online..." -ForegroundColor Cyan
        $exoParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
        if ($TenantId) { $exoParams['Organization'] = $TenantId }
        Connect-ExchangeOnline @exoParams
        $script:ConnectedExo = $true
    }

    $mbx = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
    $mbxUpn  = "$($mbx.UserPrincipalName)"
    $mbxSmtp = "$($mbx.PrimarySmtpAddress)"
    $mbxGuid = "$($mbx.ExchangeGuid)"
    Write-Host "  [OK]   Mailbox: $($mbx.DisplayName) <$mbxSmtp>" -ForegroundColor DarkGray

    # ── Is the evidence there at all? ──────────────────────────────────────────
    $ingest = $null
    try { $ingest = (Get-AdminAuditLogConfig -ErrorAction Stop).UnifiedAuditLogIngestionEnabled } catch {}
    if ($ingest -eq $false) {
        Write-Warning "The Unified Audit Log is switched OFF for this tenant. Nobody can be named and no move can be traced back; only deleted messages can be restored."
    }
    if ($mbx.AuditEnabled -eq $false) {
        Write-Warning "Mailbox auditing is disabled on this mailbox - expect no audit records."
    }
    # A long window can reach past what the mailbox still keeps. Say so up
    # front, rather than let "0 found" read as "nothing was deleted".
    $retain = $null
    try { $retain = [timespan]"$($mbx.RetainDeletedItemsFor)" } catch {}
    $onHold = $mbx.LitigationHoldEnabled -or @($mbx.InPlaceHolds).Count -gt 0
    if ($retain -and -not $onHold -and $windowStart -lt (Get-Date).Add(-$retain)) {
        Write-Warning ("Recoverable Items keeps deleted items for {0} days on this mailbox and it is not on hold: anything removed from Recoverable Items before {1} is permanently gone. Deleted Items itself is unaffected." -f [int]$retain.TotalDays, (Get-Date).Add(-$retain).ToString('yyyy-MM-dd'))
    }
    if ($windowStart -lt (Get-Date).AddDays(-180)) {
        Write-Warning "The window starts more than 180 days ago. The audit log usually keeps 180 days (90 on older tenants); before that nothing can be traced back or attributed."
    }

    $notAudited = @()
    foreach ($pair in @(@('Owner', $mbx.AuditOwner), @('Delegate', $mbx.AuditDelegate), @('Admin', $mbx.AuditAdmin))) {
        $set = @($pair[1] | ForEach-Object { "$_" })
        $miss = @('Move', 'MoveToDeletedItems', 'SoftDelete', 'HardDelete') | Where-Object { $set -notcontains $_ }
        if ($miss) { $notAudited += "$($pair[0]): $($miss -join ', ')" }
    }
    if ($notAudited) {
        Write-Host "  Not audited on this mailbox (these actions leave no trace, and no name):" -ForegroundColor Yellow
        foreach ($n in $notAudited) { Write-Host "    $n" -ForegroundColor Yellow }
        if ($notAudited -match '^Owner: .*\bMove\b') {
            Write-Host "    To audit the owner's own moves from now on: Set-Mailbox $mbxSmtp -AuditOwner @{Add='Move'}" -ForegroundColor DarkGray
        }
    }
    Write-Host ""

    try {
        Get-AuditEvents -MailboxGuid $mbxGuid -MailboxUpn $mbxUpn -MailboxSmtp $mbxSmtp
    } catch {
        Write-Warning "Audit log search failed: $($_.Exception.Message). Without it nobody can be named and moves cannot be traced back (needs the View-Only Audit Logs role)."
    }

    if ($Include -contains 'Deleted') {
        Write-Host ""
        Write-Host "  -- Deleted messages ------------------------------" -ForegroundColor Cyan
        # Exchange RBAC only exposes the cmdlets your roles allow - a missing
        # cmdlet here means the role, not the module.
        if (Get-Command Get-RecoverableItems -ErrorAction SilentlyContinue) {
            Invoke-DeletedRestore -Identity $mbxSmtp
        } else {
            Write-Host "  No 'Mailbox Import Export' role, so Restore-RecoverableItems is not available." -ForegroundColor Yellow
            if (-not $graphOk) { $graphOk = Connect-GraphForMail }
            if ($graphOk) {
                Invoke-DeletedRestoreViaGraph -GraphMailbox $mbxUpn
            } else {
                Write-Warning "No Graph access either - the deleted messages are skipped. Assign the role (New-ManagementRoleAssignment -Role 'Mailbox Import Export' -User <you>) or supply Graph access, and re-run."
            }
        }
    }
    if ($Include -contains 'Moved' -and $graphOk) { Invoke-MovedRestore -GraphMailbox $mbxUpn }
} finally {
    # The temporary app must go before the delegated token that can delete it expires.
    Remove-TempApp
    if ($script:GraphConnected) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    if ($script:ConnectedExo) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
}

# ── Who did it ────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  -- Who moved / deleted what ----------------------" -ForegroundColor Cyan
if ($script:AuditEvents.Count -eq 0) {
    Write-Host "  No audited actions found. See the 'Not audited' list above; the log can also lag up to a few hours." -ForegroundColor Yellow
} else {
    $who = $script:AuditEvents | Group-Object Actor, LogonType, Client, Operation | ForEach-Object {
        $first = $_.Group[0]
        [PSCustomObject]@{
            Actor     = $first.Actor
            As        = $first.LogonType
            Client    = $first.Client
            Operation = $first.Operation
            Items     = $_.Count
            # Sorted rather than Measure-Object: PS 5.1 cannot take the minimum of a DateTime.
            From      = @($_.Group | Sort-Object Time)[0].Time.ToString('HH:mm')
            To        = @($_.Group | Sort-Object Time)[-1].Time.ToString('HH:mm')
            IP        = (@($_.Group | ForEach-Object { $_.ClientIP } | Where-Object { $_ } | Select-Object -Unique) -join ', ')
        }
    } | Sort-Object Items -Descending
    $who | Format-Table -AutoSize | Out-String -Width 200 | ForEach-Object { Write-Host $_.TrimEnd() }
}

# ── Summary ───────────────────────────────────────────────────────────────────
$count = { param($s) @($results | Where-Object { $_.Status -eq $s }).Count }
Write-Host ""
Write-Host "  ------------------------------------------------" -ForegroundColor Cyan
Write-Host "  Deleted found       : $(@($results | Where-Object { $_.Phase -eq 'Deleted' }).Count)" -ForegroundColor Cyan
Write-Host "  Moved (audited)     : $(@($results | Where-Object { $_.Phase -eq 'Moved' }).Count)" -ForegroundColor Cyan
Write-Host "  Archive, unaudited  : $(@($results | Where-Object { $_.Phase -eq 'Unaudited' }).Count)" -ForegroundColor Cyan
Write-Host "  Already in place    : $(& $count 'AlreadyInPlace')" -ForegroundColor DarkGray
if ($Apply) {
    Write-Host "  Restored            : $(& $count 'Restored')" -ForegroundColor Green
} else {
    Write-Host "  Would restore       : $(& $count 'WouldRestore') - preview. Re-run with -Apply." -ForegroundColor Yellow
}
$problems = @($results | Where-Object { $_.Status -in @('Error', 'NotRestored', 'NotFound', 'Ambiguous', 'Unreachable') }).Count
if ($problems -gt 0)         { Write-Host "  Needs a look        : $problems (Error / NotRestored / NotFound / Ambiguous / Unreachable - see CSV)" -ForegroundColor Red }
if ($script:AuditTruncated)  { Write-Host "  Audit log truncated : yes - narrow the window." -ForegroundColor Yellow }

# ── Output ────────────────────────────────────────────────────────────────────
if ($results.Count -gt 0 -or $script:AuditEvents.Count -gt 0) {
    if (-not $OutputPath) {
        $safe = ($Mailbox -replace '[^A-Za-z0-9._-]', '_')
        $OutputPath = Join-Path $outputDir "MailRestore_${safe}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    }
    if ($results.Count -gt 0) { $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 }
    $auditPath = [System.IO.Path]::ChangeExtension($OutputPath, $null).TrimEnd('.') + '_Audit.csv'
    $script:AuditEvents | Sort-Object Time | Select-Object @{ n = 'Time'; e = { $_.Time.ToString('yyyy-MM-dd HH:mm:ss') } },
        Operation, Actor, LogonType, Client, ClientIP, Subject, SourcePath, DestPath, InternetMessageId |
        Export-Csv -Path $auditPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    if ($results.Count -gt 0) { Write-Host "  Report saved : $OutputPath" -ForegroundColor Green }
    Write-Host "  Audit trail  : $auditPath" -ForegroundColor Green
}
Write-Host ""
