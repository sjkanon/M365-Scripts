#Requires -Version 7.0
<#
.SYNOPSIS
    Forward all mail of a mailbox to another address: new mail from now on, and the mail
    that already arrived in the last N days.

.DESCRIPTION
    Two parts, each of which can be skipped:

      1. Forwarding rule (Exchange Online) - sets mailbox forwarding on -Mailbox so every
         new message also goes to -ForwardTo. A copy stays in the mailbox unless
         -NoKeepCopy is given. An internal -ForwardTo (a recipient in this tenant) is set
         as ForwardingAddress, which the outbound auto-forward policy does not block; an
         external one as ForwardingSmtpAddress, which needs auto-forwarding allowed in
         the outbound spam policy (the script warns when it is not).

      2. Backfill (Microsoft Graph) - forwards every message received in the last -Days
         days (or since -After), oldest first, so the target gets the history that
         arrived before the rule existed. Sent Items, Drafts, Deleted Items, Junk Email
         and Outbox are skipped. Each forward is an ordinary sent message: it is not
         subject to the auto-forward policy, and a copy lands in the mailbox's Sent Items.
         The window ends at the moment the run starts, so running both parts together
         does not forward the same message twice.

    Preview by default: the script shows the current forwarding and how many messages
    would be forwarded, and changes nothing. Only -Apply sets the rule and sends.

    Graph authentication - same routes as Move-InboxToArchive.ps1 /
    Remove-PhishingMessage.ps1:

      Temporary app (default)
          Signs you in with a device code (Global Administrator or Privileged Role
          Administrator), creates a short-lived App Registration, grants it the
          application permissions Mail.Read and Mail.Send (the delegated role does the
          consent, no separate consent screen), uses it app-only, and removes it again
          when the run ends - also when the run fails.

      Your own app registration
          -ClientId with -ClientSecret or -CertificateThumbprint (certificate in
          Cert:\CurrentUser\My or Cert:\LocalMachine\My), or -AppOnly to take ClientId
          and CertificateThumbprint from graph.appid.json in the repo root. That app needs
          the application permissions Mail.Read and Mail.Send, admin-consented. Tip: limit
          it to this mailbox with an Application Access Policy or RBAC for Applications.

      -Delegated
          Graph as yourself (Mail.Read.Shared + Mail.Send.Shared). Only works when your
          account has Full Access and Send As on -Mailbox. Not under GDAP.

    All Graph traffic is plain REST (device code, client credentials, signed client
    assertion for a certificate): the Graph SDK is never loaded, so it cannot clash with
    the Microsoft.Identity.Client that ExchangeOnlineManagement loads for part 1.

    Exchange Online sends at most 30 messages per minute per mailbox and 10,000
    recipients per day. -DelaySeconds (default 2.1) keeps the backfill under the first;
    the script warns when the backfill would come near the second.

    GDAP-aware: under a GDAP session -TenantId is resolved from the selected customer
    tenant when not supplied, and Exchange Online connects to that customer.

.PARAMETER Mailbox
    UPN or primary SMTP address of the mailbox whose mail is forwarded.

.PARAMETER ForwardTo
    Address that receives the mail. Internal or external.

.PARAMETER Days
    Backfill window in days, counted back from the start of the run. Default 30.

.PARAMETER After
    Backfill everything received from this date on, instead of -Days.

.PARAMETER NoKeepCopy
    Forwarding rule without a copy in the mailbox (DeliverToMailboxAndForward $false).
    Default is to keep a copy.

.PARAMETER Comment
    Text placed above each backfilled message. Default
    "Automatically forwarded from <Mailbox>".

.PARAMETER SkipForwardingRule
    Do only the backfill.

.PARAMETER SkipBackfill
    Set only the forwarding rule. No Graph sign-in is made.

.PARAMETER Apply
    Set the rule and forward the messages. Without it the script only previews.

.PARAMETER DelaySeconds
    Pause between two forwards. Default 2.1, just under Exchange's 30 messages per minute.

.PARAMETER LogPath
    CSV with one row per backfilled message (received, from, subject, status). Default
    C:\Temp\MailboxForward_<mailbox>_<timestamp>.csv ($HOME/Downloads outside Windows).

.PARAMETER TenantId
    Tenant ID or verified domain. Needed for the temporary app and for -ClientId, unless
    resolvable from a GDAP customer tenant context.

.PARAMETER ClientId
    Your own App Registration (application permissions Mail.Read and Mail.Send).

.PARAMETER ClientSecret
    Client secret for -ClientId.

.PARAMETER CertificateThumbprint
    Certificate for -ClientId, from Cert:\CurrentUser\My or Cert:\LocalMachine\My.

.PARAMETER AppOnly
    ClientId and CertificateThumbprint from graph.appid.json in the repo root.

.PARAMETER Delegated
    Graph as yourself instead of an app; needs Full Access and Send As on -Mailbox.

.EXAMPLE
    # Preview: current forwarding and how many messages of the last 30 days would go
    .\Send-MailboxForward.ps1 -Mailbox "jan@contoso.com" -ForwardTo "sales@fabrikam.com"

.EXAMPLE
    # Rule plus 30 days of history, with a temporary app registration
    .\Send-MailboxForward.ps1 -Mailbox "jan@contoso.com" -ForwardTo "sales@fabrikam.com" -TenantId contoso.com -Apply

.EXAMPLE
    # Only history since 1 September, with your own app (certificate)
    .\Send-MailboxForward.ps1 -Mailbox "jan@contoso.com" -ForwardTo "sales@fabrikam.com" `
        -After 2026-09-01 -SkipForwardingRule -TenantId contoso.com `
        -ClientId "00000000-0000-0000-0000-000000000000" -CertificateThumbprint "ABC123..." -Apply

.EXAMPLE
    # Only the forwarding rule, no copy in the mailbox
    .\Send-MailboxForward.ps1 -Mailbox "jan@contoso.com" -ForwardTo "piet@contoso.com" -SkipBackfill -NoKeepCopy -Apply
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $Mailbox,
    [Parameter(Mandatory)] [string] $ForwardTo,
    [ValidateRange(1, 3650)]
    [int]      $Days = 30,
    [datetime] $After,
    [switch]   $NoKeepCopy,
    [string]   $Comment,
    [switch]   $SkipForwardingRule,
    [switch]   $SkipBackfill,
    [switch]   $Apply,
    [ValidateRange(0, 60)]
    [double]   $DelaySeconds = 2.1,
    [string]   $LogPath,
    [string]   $TenantId,
    [string]   $ClientId,
    [string]   $ClientSecret,
    [string]   $CertificateThumbprint,
    [switch]   $AppOnly,
    [switch]   $Delegated
)

. (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')
$ErrorActionPreference = 'Stop'

# ── Validation ────────────────────────────────────────────────────────────────
if ($SkipForwardingRule -and $SkipBackfill) { throw "-SkipForwardingRule and -SkipBackfill together leave nothing to do." }
if ($ClientId -and -not $ClientSecret -and -not $CertificateThumbprint) { throw "-ClientId needs -ClientSecret or -CertificateThumbprint." }
if ($Delegated -and ($ClientId -or $AppOnly)) { throw "Use either -Delegated or -ClientId/-AppOnly, not both." }
if ($Delegated -and (Test-M365Gdap)) {
    throw "-Delegated does not work under GDAP: a partner account cannot be given Full Access on a customer mailbox. Drop -Delegated."
}
if ($ForwardTo -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { throw "-ForwardTo '$ForwardTo' is not an email address." }
if ($Mailbox -eq $ForwardTo) { throw "-Mailbox and -ForwardTo are the same address." }

$effectiveTenantId = Resolve-M365TenantId -TenantId $TenantId
if ($AppOnly -and -not $ClientId) {
    $reg = Get-M365AppRegistration -TenantId $effectiveTenantId
    $ClientId = $reg.ClientId
    $CertificateThumbprint = $reg.CertificateThumbprint
    if (-not $effectiveTenantId) { $effectiveTenantId = $reg.Tenant }
}
if (-not $SkipBackfill -and -not $Delegated -and -not $effectiveTenantId) {
    throw "-TenantId is required for the backfill (temporary app or -ClientId), or a resolvable GDAP customer tenant."
}

$runStart = (Get-Date).ToUniversalTime()
$since    = if ($After) { $After.ToUniversalTime() } else { $runStart.AddDays(-$Days) }
if ($since -ge $runStart) { throw "-After lies in the future." }
if (-not $Comment) { $Comment = "Automatically forwarded from $Mailbox" }

if (-not $LogPath) {
    $outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
    if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
    $safeName = $Mailbox -replace '[^\w.@-]', '_'
    $LogPath  = Join-Path $outputDir "MailboxForward_${safeName}_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
}

# ── Header ────────────────────────────────────────────────────────────────────
$authLabel = if ($SkipBackfill) { '-' }
             elseif ($Delegated) { 'Delegated (your Full Access + Send As)' }
             elseif ($ClientId)  { 'App-only (your app registration)' }
             else                { 'App-only (temporary app registration)' }
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Mailbox Forward" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Mailbox    : $Mailbox"
Write-Host "  Forward to : $ForwardTo"
Write-Host ("  Rule       : {0}" -f $(if ($SkipForwardingRule) { 'skipped' } elseif ($NoKeepCopy) { 'forward, no copy in mailbox' } else { 'forward, keep a copy in mailbox' })) -ForegroundColor DarkGray
Write-Host ("  Backfill   : {0}" -f $(if ($SkipBackfill) { 'skipped' } else { "received since $($since.ToLocalTime().ToString('yyyy-MM-dd HH:mm'))" })) -ForegroundColor DarkGray
Write-Host "  Graph auth : $authLabel" -ForegroundColor DarkGray
Write-Host ("  Mode       : {0}" -f $(if ($Apply) { 'APPLY - rule is set and messages are sent' } else { 'Preview only (no changes)' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

# ══════════════════════════════════════════════════════════════════════════════
#  Graph over plain REST
# ══════════════════════════════════════════════════════════════════════════════
$script:Headers         = $null
$script:TokenBody       = $null   # hashtable, or a scriptblock that builds one (certificate)
$script:TokenTenantId   = $null
$script:TokenExpiry     = [datetime]::MinValue
$script:AccessToken     = $null
$script:AdminHeaders    = $null
$script:TempAppObjectId = $null
$script:GraphCliClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'   # Graph PowerShell's public client

function Update-GraphToken {
    if (-not $script:TokenBody) { return }
    if ($script:Headers -and (Get-Date) -lt $script:TokenExpiry) { return }
    $body = if ($script:TokenBody -is [scriptblock]) { & $script:TokenBody } else { $script:TokenBody }
    $resp = Invoke-RestMethod -Method POST -Body $body -Uri "https://login.microsoftonline.com/$($script:TokenTenantId)/oauth2/v2.0/token"
    $script:AccessToken = $resp.access_token
    $script:Headers     = @{ Authorization = "Bearer $($resp.access_token)" }
    $script:TokenExpiry = (Get-Date).AddSeconds($resp.expires_in - 300)
    if ($resp.refresh_token -and $script:TokenBody -is [hashtable] -and $script:TokenBody.grant_type -eq 'refresh_token') {
        $script:TokenBody.refresh_token = $resp.refresh_token
    }
}

function Invoke-Graph {
    # Retries throttling and gateway errors; a 401/403 never improves by asking again.
    param(
        [string] $Method = 'GET',
        [Parameter(Mandatory)] [string] $Uri,
        $Body
    )
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            Update-GraphToken
            $p = @{ Method = $Method; Uri = $Uri; Headers = $script:Headers }
            if ($Body) { $p['Body'] = ($Body | ConvertTo-Json -Depth 8); $p['ContentType'] = 'application/json' }
            return Invoke-RestMethod @p
        } catch {
            $err  = $_
            $code = 0
            try { $code = [int]$err.Exception.Response.StatusCode } catch { }
            if ($attempt -ge 6 -or $code -notin @(429, 503, 504)) { throw $err }
            $wait = 10 * $attempt
            try { $ra = [int]"$($err.Exception.Response.Headers.RetryAfter.Delta.TotalSeconds)"; if ($ra -gt 0) { $wait = $ra } } catch { }
            Write-Host "    Throttled (HTTP $code), waiting ${wait}s..." -ForegroundColor DarkGray
            Start-Sleep -Seconds $wait
        }
    }
}

function Invoke-GraphAdmin {
    # App-management calls with the delegated device-code token.
    param([string] $Method = 'GET', [Parameter(Mandatory)] [string] $Uri, $Body)
    $p = @{ Method = $Method; Uri = $Uri; Headers = $script:AdminHeaders }
    if ($Body) { $p['Body'] = ($Body | ConvertTo-Json -Depth 6); $p['ContentType'] = 'application/json' }
    Invoke-RestMethod @p
}

function Get-DeviceCodeToken {
    # Device code flow over REST; returns the raw token response.
    param([string] $Tenant, [string[]] $Scopes)
    $scope = ((@($Scopes | ForEach-Object { "https://graph.microsoft.com/$_" })) + 'offline_access') -join ' '
    $dc = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/devicecode" `
            -Body @{ client_id = $script:GraphCliClientId; scope = $scope }
    Write-Host ""
    Write-Host "  ------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "   $($dc.message)" -ForegroundColor Yellow
    Write-Host "  ------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host ""
    $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
    $interval = [Math]::Max(5, [int]$dc.interval)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval
        try {
            return Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" -Body @{
                grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                client_id   = $script:GraphCliClientId
                device_code = $dc.device_code
            }
        } catch {
            $code = ''
            try { $code = ($_.ErrorDetails.Message | ConvertFrom-Json).error } catch { }
            if ($code -eq 'authorization_pending') { continue }
            if ($code -eq 'slow_down') { $interval += 5; continue }
            throw
        }
    }
    throw "Device code sign-in timed out."
}

function ConvertTo-Base64Url([byte[]] $Bytes) {
    [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function New-CertificateTokenBody {
    # client_credentials with a signed client assertion: certificate auth without the SDK.
    $cert = Get-Item "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
    if (-not $cert) { $cert = Get-Item "Cert:\LocalMachine\My\$CertificateThumbprint" -ErrorAction SilentlyContinue }
    if (-not $cert) { throw "Certificate $CertificateThumbprint not found in Cert:\CurrentUser\My or Cert:\LocalMachine\My." }
    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
    if (-not $rsa) { throw "Certificate $CertificateThumbprint has no usable RSA private key." }

    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $header  = @{ alg = 'RS256'; typ = 'JWT'; x5t = (ConvertTo-Base64Url $cert.GetCertHash()) } | ConvertTo-Json -Compress
    $payload = @{
        aud = "https://login.microsoftonline.com/$effectiveTenantId/oauth2/v2.0/token"
        iss = $ClientId; sub = $ClientId; jti = [guid]::NewGuid().ToString()
        nbf = $now; exp = $now + 600
    } | ConvertTo-Json -Compress
    $unsigned = "$(ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($header))).$(ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($payload)))"
    $sig = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($unsigned),
            [System.Security.Cryptography.HashAlgorithmName]::SHA256,
            [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    @{
        grant_type            = 'client_credentials'
        scope                 = 'https://graph.microsoft.com/.default'
        client_id             = $ClientId
        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        client_assertion      = "$unsigned.$(ConvertTo-Base64Url $sig)"
    }
}

function Get-TokenRole {
    try {
        $p = $script:AccessToken.Split('.')[1].Replace('-', '+').Replace('_', '/')
        switch ($p.Length % 4) { 2 { $p += '==' } 3 { $p += '=' } }
        @(([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json).roles)
    } catch { @() }
}

function Remove-TempApp {
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

function Connect-GraphForForward {
    $requiredRoles = @('Mail.Read', 'Mail.Send')

    if ($Delegated) {
        $tenant = if ($effectiveTenantId) { $effectiveTenantId } else { 'organizations' }
        $scopes = @('Mail.Read.Shared', 'Mail.Send.Shared')
        Write-Host "  Signing in as yourself for Graph ($($scopes -join ', '))..." -ForegroundColor Cyan
        $tok = Get-DeviceCodeToken -Tenant $tenant -Scopes $scopes
        $script:TokenTenantId = $tenant
        $script:AccessToken   = $tok.access_token
        $script:Headers       = @{ Authorization = "Bearer $($tok.access_token)" }
        $script:TokenExpiry   = (Get-Date).AddSeconds([int]$tok.expires_in - 300)
        $script:TokenBody     = @{
            grant_type    = 'refresh_token'
            client_id     = $script:GraphCliClientId
            refresh_token = $tok.refresh_token
            scope         = ((@($scopes | ForEach-Object { "https://graph.microsoft.com/$_" })) + 'offline_access') -join ' '
        }
        Write-Host "  [OK]   Signed in (delegated)." -ForegroundColor DarkGray
        return
    }

    $script:TokenTenantId = $effectiveTenantId

    if ($ClientId) {
        $script:TokenBody = if ($CertificateThumbprint) { ${function:New-CertificateTokenBody} } else {
            @{ grant_type = 'client_credentials'; scope = 'https://graph.microsoft.com/.default'; client_id = $ClientId; client_secret = $ClientSecret }
        }
        Update-GraphToken
        $missing = @($requiredRoles | Where-Object { (Get-TokenRole) -notcontains $_ })
        if ($missing) { throw "App $ClientId lacks the application permission(s) $($missing -join ', ') (admin consent). Grant them and re-run." }
        Write-Host "  [OK]   App-only token obtained (your app registration)." -ForegroundColor DarkGray
        return
    }

    # Temporary app: create, grant, use, remove at the end.
    Write-Host "  Setting up a temporary App Registration (Mail.Read + Mail.Send)." -ForegroundColor Cyan
    Write-Host "  Required role: Global Administrator or Privileged Role Administrator (one-time)" -ForegroundColor DarkGray
    $admin = Get-DeviceCodeToken -Tenant $effectiveTenantId -Scopes @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
    $script:AdminHeaders = @{ Authorization = "Bearer $($admin.access_token)" }
    Write-Host "  [OK]   Signed in." -ForegroundColor DarkGray

    $appName = "MailboxForward-Temp-$(Get-Date -Format 'yyyyMMddHHmmss')"
    $app = Invoke-GraphAdmin -Method POST -Uri 'https://graph.microsoft.com/v1.0/applications' `
            -Body @{ displayName = $appName; signInAudience = 'AzureADMyOrg' }
    $script:TempAppObjectId = $app.id
    Write-Host "  [OK]   Created '$appName'." -ForegroundColor DarkGray

    $sp = Invoke-GraphAdmin -Method POST -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Body @{ appId = $app.appId }
    $graphSp = @((Invoke-GraphAdmin -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'").value)[0]
    foreach ($roleName in $requiredRoles) {
        $role = @($graphSp.appRoles | Where-Object { $_.value -eq $roleName -and $_.allowedMemberTypes -contains 'Application' })[0]
        if (-not $role) { throw "Could not resolve the $roleName application role." }
        Invoke-GraphAdmin -Method POST -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)/appRoleAssignments" `
            -Body @{ principalId = $sp.id; resourceId = $graphSp.id; appRoleId = $role.id } | Out-Null
        Write-Host "  [OK]   $roleName (application) granted." -ForegroundColor DarkGray
    }

    $secret = Invoke-GraphAdmin -Method POST -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)/addPassword" `
            -Body @{ passwordCredential = @{ displayName = 'temp'; endDateTime = (Get-Date).AddHours(4).ToString('o') } }
    $script:TokenBody = @{ grant_type = 'client_credentials'; scope = 'https://graph.microsoft.com/.default'; client_id = $app.appId; client_secret = $secret.secretText }

    # A new registration takes a while to accept a token request, and a token minted
    # before the grants propagated is valid but powerless - wait for both.
    $deadline = (Get-Date).AddMinutes(4)
    Write-Host "  Waiting for the app and its permissions to propagate..." -ForegroundColor DarkGray
    while ($true) {
        try {
            $script:Headers = $null
            Update-GraphToken
            $missing = @($requiredRoles | Where-Object { (Get-TokenRole) -notcontains $_ })
            if (-not $missing) { break }
        } catch { }
        if ((Get-Date) -ge $deadline) { throw "The temporary app's token still lacks Mail.Read/Mail.Send after 4 minutes." }
        Start-Sleep -Seconds 10
    }
    Write-Host "  [OK]   App-only token obtained (temporary app)." -ForegroundColor DarkGray
}

# ══════════════════════════════════════════════════════════════════════════════
#  Run
# ══════════════════════════════════════════════════════════════════════════════
$script:Exo = $null
$exitCode   = 0

try {
    # Graph first: no SDK is involved, but the sign-in prompts are easier to follow
    # before the Exchange output starts.
    if (-not $SkipBackfill) { Connect-GraphForForward }

    # ── 1. Forwarding rule ────────────────────────────────────────────────────
    if (-not $SkipForwardingRule) {
        Write-Host ""
        Write-Host "  [1] Forwarding rule" -ForegroundColor Cyan
        $script:Exo = Connect-M365Exchange -TenantId $TenantId

        $mbx = Get-EXOMailbox -Identity $Mailbox -Properties ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward
        Write-Host ("  Now        : ForwardingAddress='{0}'  ForwardingSmtpAddress='{1}'  KeepCopy={2}" -f `
            $mbx.ForwardingAddress, $mbx.ForwardingSmtpAddress, $mbx.DeliverToMailboxAndForward) -ForegroundColor DarkGray

        $internal = Get-Recipient -Identity $ForwardTo -ErrorAction SilentlyContinue | Select-Object -First 1
        $keepCopy = -not $NoKeepCopy
        if ($internal) {
            Write-Host "  Target     : internal recipient ($($internal.RecipientTypeDetails)) -> ForwardingAddress" -ForegroundColor DarkGray
        } else {
            Write-Host "  Target     : external -> ForwardingSmtpAddress" -ForegroundColor DarkGray
            # The default outbound spam policy blocks external auto-forwarding (Automatic = Off).
            $policies = @(Get-HostedOutboundSpamFilterPolicy -ErrorAction SilentlyContinue)
            $default  = $policies | Where-Object IsDefault | Select-Object -First 1
            if ($default -and $default.AutoForwardingMode -ne 'On') {
                Write-Warning "The default outbound spam policy has AutoForwardingMode '$($default.AutoForwardingMode)': external forwards are blocked unless a custom policy covering $Mailbox sets it to 'On'."
            }
        }

        if ($Apply) {
            # Clear the other field first, so an old forward to elsewhere does not stay behind.
            if ($internal) {
                Set-Mailbox -Identity $Mailbox -ForwardingSmtpAddress $null -ForwardingAddress $internal.Identity -DeliverToMailboxAndForward $keepCopy
            } else {
                Set-Mailbox -Identity $Mailbox -ForwardingAddress $null -ForwardingSmtpAddress "smtp:$ForwardTo" -DeliverToMailboxAndForward $keepCopy
            }
            $mbx = Get-EXOMailbox -Identity $Mailbox -Properties ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward
            Write-Host ("  [OK]   Set : ForwardingAddress='{0}'  ForwardingSmtpAddress='{1}'  KeepCopy={2}" -f `
                $mbx.ForwardingAddress, $mbx.ForwardingSmtpAddress, $mbx.DeliverToMailboxAndForward) -ForegroundColor Green
        } else {
            Write-Host "  Preview    : would forward to $ForwardTo, keep copy = $keepCopy" -ForegroundColor Yellow
        }
    }

    # ── 2. Backfill ───────────────────────────────────────────────────────────
    if (-not $SkipBackfill) {
        Write-Host ""
        Write-Host "  [2] Backfill" -ForegroundColor Cyan
        $userUri = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($Mailbox))"

        $skipFolders = foreach ($wk in 'sentitems', 'drafts', 'deleteditems', 'junkemail', 'outbox') {
            try { (Invoke-Graph -Uri "$userUri/mailFolders/$wk`?`$select=id").id } catch { }
        }

        $filter = "receivedDateTime ge $($since.ToString('yyyy-MM-ddTHH:mm:ssZ')) and receivedDateTime lt $($runStart.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
        $uri = "$userUri/messages?`$filter=$([uri]::EscapeDataString($filter))&`$select=id,subject,receivedDateTime,from,parentFolderId,isDraft&`$top=200"
        $messages = [System.Collections.Generic.List[object]]::new()
        try {
            while ($uri) {
                $page = Invoke-Graph -Uri $uri
                foreach ($m in $page.value) {
                    if (-not $m.isDraft -and $m.parentFolderId -notin $skipFolders) { $messages.Add($m) }
                }
                $uri = $page.'@odata.nextLink'
            }
        } catch {
            $code = 0; try { $code = [int]$_.Exception.Response.StatusCode } catch { }
            if ($code -eq 403) {
                $hint = if ($Delegated) { "your account needs Full Access on $Mailbox." }
                        else { "an Application Access Policy or RBAC for Applications may exclude $Mailbox from this app." }
                throw "Graph refused access to $Mailbox (403): $hint"
            }
            throw
        }
        $messages = @($messages | Sort-Object { [datetime]$_.receivedDateTime })   # oldest first

        Write-Host "  Found      : $($messages.Count) message(s) received since $($since.ToLocalTime().ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor DarkGray
        if ($messages.Count -gt 9000) {
            Write-Warning "$($messages.Count) messages comes close to Exchange's 10,000 recipients per day; the rest will bounce. Split the run with -After."
        }
        if ($messages.Count -and $DelaySeconds -gt 0) {
            Write-Host ("  Duration   : about {0:N0} minute(s) at {1}s per message" -f ($messages.Count * $DelaySeconds / 60), $DelaySeconds) -ForegroundColor DarkGray
        }

        if (-not $Apply) {
            if ($messages.Count) {
                Write-Host "  Preview    : $($messages.Count) message(s) would be forwarded. Oldest / newest:" -ForegroundColor Yellow
                $shown = if ($messages.Count -le 6) { $messages } else { @($messages[0..2]) + @($messages[-3..-1]) }
                $shown | ForEach-Object { Write-Host ("    {0}  {1}  {2}" -f ([datetime]$_.receivedDateTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm'), $_.from.emailAddress.address, $_.subject) -ForegroundColor DarkGray }
            }
        } elseif ($messages.Count) {
            $body = @{
                comment      = $Comment
                toRecipients = @(@{ emailAddress = @{ address = $ForwardTo } })
            }
            $log = [System.Collections.Generic.List[object]]::new()
            $i = 0
            foreach ($m in $messages) {
                $i++
                Write-Progress -Activity "Forwarding $Mailbox -> $ForwardTo" -Status "$i / $($messages.Count)" -PercentComplete ($i / $messages.Count * 100)
                $status = 'OK'
                try {
                    Invoke-Graph -Method POST -Uri "$userUri/messages/$($m.id)/forward" -Body $body | Out-Null
                } catch {
                    $status = "FAILED: $($_.Exception.Message)"
                }
                $log.Add([pscustomobject]@{
                    Received = ([datetime]$m.receivedDateTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
                    From     = $m.from.emailAddress.address
                    Subject  = $m.subject
                    Status   = $status
                })
                if ($DelaySeconds -gt 0 -and $i -lt $messages.Count) { Start-Sleep -Milliseconds ([int]($DelaySeconds * 1000)) }
            }
            Write-Progress -Activity "Forwarding" -Completed

            $log | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
            $failed = @($log | Where-Object Status -like 'FAILED*').Count
            Write-Host ""
            Write-Host "  Forwarded  : $($log.Count - $failed)" -ForegroundColor Green
            if ($failed) {
                Write-Host "  Failed     : $failed (see the log)" -ForegroundColor Yellow
                $exitCode = 1
            }
            Write-Host "  Log        : $LogPath" -ForegroundColor DarkGray
        }
    }

    if (-not $Apply) {
        Write-Host ""
        Write-Host "  Preview done - nothing changed. Re-run with -Apply to set the rule and forward." -ForegroundColor Yellow
    }
} catch {
    Write-Host ""
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    $exitCode = 1
} finally {
    Remove-TempApp
    Disconnect-M365Exchange $script:Exo
    Write-Host ""
}

exit $exitCode
