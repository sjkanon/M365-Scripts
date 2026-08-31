#Requires -Version 5.1
<#
.SYNOPSIS
    Find and delete a phishing message from one, several, or all mailboxes in a
    Microsoft 365 tenant. Dry-run by default.

.DESCRIPTION
    Incident-response companion to Get-MessageTraceReport.ps1: the trace tells you
    who received the phish, this script removes it again.

    Two engines, because neither one covers every case:

      Purview  Content Search + New-ComplianceSearchAction -Purge. One KQL query
               sweeps every mailbox in the tenant, so you do not need to know the
               recipients up front. This is the only engine that can HardDelete
               (Recoverable Items\Purges). Three downsides: it reads the search
               index, which lags delivery by roughly 15-30 minutes; it reports
               counts per mailbox but never the individual messages (the preview
               action was retired in the cloud in May 2025); and a purge removes
               at most 10 items per mailbox per round, so the script plans the
               number of rounds from the busiest mailbox.

      Graph    Enumerates each target mailbox over the Graph mail API and deletes
               the matching messages one by one. No index lag (a message is
               findable the second it lands) and you get a per-message report of
               exactly what was removed, including folder and subject. Needs the
               recipient list, and cannot reach Recoverable Items\Purges, so
               HardDelete is not available here.

    Engine defaults to Graph when -Mailbox is given and Purview otherwise, which
    matches how these two are normally used. Override with -Engine.

    NOTHING IS DELETED WITHOUT -Apply. Every run without it performs the full
    search and reports precisely what it would have removed.

.PARAMETER Mailbox
    Target mailbox address(es). Required for the Graph engine unless
    -AllMailboxes is used. For the Purview engine, omitting this searches every
    mailbox in the tenant.

.PARAMETER AllMailboxes
    Search every mailbox in the tenant. Implicit for Purview when -Mailbox is
    omitted; for Graph this enumerates all user mailboxes, which is slow - one
    Graph query per mailbox.

.PARAMETER MessageId
    Internet MessageId of the phish, with or without angle brackets. By far the
    most precise selector - it matches that one message and nothing else. Take it
    from the message trace or from the message header.

.PARAMETER SenderAddress
    Sender address to match. Aliased as -Sender ($Sender is a PowerShell
    automatic variable).

.PARAMETER Subject
    A fragment of the subject is enough - it does not have to be the whole thing.

    Purview matches it as a phrase, which already matches anywhere in the subject:
    "kick-off meeting" finds every subject containing those words in that order.
    What KQL cannot do is match inside a word, so a leading wildcard is dropped
    (with a warning) and a trailing one only survives on a single word.

    Graph matches client-side and takes wildcards as written; a value without any
    wildcard is treated as *value* (substring).

.PARAMETER AttachmentName
    Attachment filename to match, wildcards allowed (e.g. "*.html").

.PARAMETER BodyContains
    Word or phrase from the message body - a distinctive sentence out of the phish
    is often the most durable selector when the sender rotates addresses. Purview
    engine only: Graph would have to download every body to test it.

.PARAMETER ReceivedAfter
    Only consider messages received at or after this moment (local time).

.PARAMETER ReceivedBefore
    Only consider messages received at or before this moment (local time).

.PARAMETER Engine
    Purview or Graph. Defaults to Graph when -Mailbox is supplied, Purview
    otherwise.

.PARAMETER DeleteType
    Recycle     Move to the user's Deleted Items. Recoverable by the user.
    SoftDelete  Move to Recoverable Items\Deletions. Out of the mailbox, still
                recoverable by the user via "Recover deleted items". Default.
    HardDelete  Move to Recoverable Items\Purges. Not user-recoverable; retained
                only if the mailbox is on hold. Purview engine only.

.PARAMETER Apply
    Actually delete. Without this switch the script only reports what it found.

.PARAMETER SearchName
    Name of the Content Search to create (Purview engine). Defaults to
    Phish_<timestamp>. Purview requires unique search names, so reusing the name
    of an existing search fails.

.PARAMETER KeepSearch
    Do not remove the Content Search and its purge actions afterwards. Useful
    when you want to inspect the search in the Purview portal.

.PARAMETER VerifyWithGraph
    After a Purview purge, check the affected mailboxes over Graph to confirm the
    messages are actually gone. This is the only lag-free verification available:
    the Content Search index keeps reporting purged items for up to ~30 minutes,
    and the purge action only reports what the service believes it did.

    Needs the same app-only Graph session as -Engine Graph. Soft- and hard-deleted
    items live in Recoverable Items, which Graph does not list, so a purged
    message correctly reads as gone. Verification never fails the run - if Graph
    is unavailable it says so and skips.

.PARAMETER MaxPurgeRounds
    Purview purges at most 10 items per mailbox per action, so the script repeats
    search + purge until a round removes nothing. Default 10 rounds (= up to 100
    items per mailbox).

.PARAMETER MaxMessagesPerMailbox
    Graph engine safety cap on how many matching messages to retrieve per
    mailbox. Default 500. Hitting the cap is reported explicitly.

.PARAMETER TimeoutMinutes
    How long to wait for a Content Search or a purge action to complete before
    giving up. Default 30.

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\ (Windows) or ~/Downloads.

.PARAMETER TenantId
    Entra ID tenant - either the tenant ID (GUID) or a verified domain of the
    tenant (e.g. contoso.com). Passed to Connect-ExchangeOnline /
    Connect-IPPSSession when the script has to connect itself, and required for
    app-only Graph auth unless it can be resolved from a GDAP customer context.

.PARAMETER ClientId
    Existing App Registration client ID for app-only Graph auth - skips the
    automatic temporary app. Use with -TenantId and -ClientSecret or
    -CertificateThumbprint. That app must already have Mail.ReadWrite application
    permission with admin consent granted.

.PARAMETER ClientSecret
    Client secret for the app registration given in -ClientId.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for the app registration given in -ClientId.

.EXAMPLE
    # What would be removed, tenant-wide? (no deletion - no -Apply)
    .\Remove-PhishingMessage.ps1 -MessageId "<abc123@evil.example>"

.EXAMPLE
    # Same, now actually purge it beyond user recovery
    .\Remove-PhishingMessage.ps1 -MessageId "<abc123@evil.example>" `
        -DeleteType HardDelete -Apply

.EXAMPLE
    # Known recipients from the message trace - immediate, no index lag
    .\Remove-PhishingMessage.ps1 `
        -Mailbox "a@contoso.com","b@contoso.com" `
        -Sender  "no-reply@evil.example" `
        -Subject "*password expires*" `
        -Apply

.EXAMPLE
    # Purge, then confirm over Graph that the messages are really gone
    .\Remove-PhishingMessage.ps1 -MessageId "<abc123@evil.example>" `
        -DeleteType HardDelete -Apply -VerifyWithGraph

.EXAMPLE
    # Campaign sweep: everything from one sender in a window, tenant-wide
    .\Remove-PhishingMessage.ps1 -Sender "no-reply@evil.example" `
        -ReceivedAfter (Get-Date "2026-08-30") -DeleteType HardDelete -Apply

.EXAMPLE
    # The usual phishing combination: the sender plus a fragment of the subject
    .\Remove-PhishingMessage.ps1 -Sender "no-reply@evil.example" `
        -Subject "kick-off meeting" -DeleteType HardDelete -Apply

.EXAMPLE
    # Sender rotates addresses - match on a sentence from the body instead
    .\Remove-PhishingMessage.ps1 -BodyContains "your password will expire today" `
        -ReceivedAfter (Get-Date "2026-08-30") -Apply

.EXAMPLE
    # HTML attachment campaign
    .\Remove-PhishingMessage.ps1 -AttachmentName "*.html" `
        -Sender "billing@evil.example" -Apply

.NOTES
    Permissions

      Purview  Membership of the "Search And Purge" role - in practice the
               Organization Management or eDiscovery Manager role group in the
               Purview compliance portal.

               Content Search runs on a backend that a plain IPPS connection no
               longer reaches: the session must be opened with
               -EnableSearchOnlySession, or Start-ComplianceSearch fails at
               initialisation even though the cmdlets are present. The script
               passes that switch when it connects itself, which needs
               ExchangeOnlineManagement 3.9.0 or higher. If you connected before
               starting the script and did not pass it, the session cannot be
               repaired from inside the process - open a new PowerShell window.

      Graph    Needs app-only Mail.ReadWrite; delegated Mail.ReadWrite only ever
               reaches your own mailbox. You do not have to arrange that
               yourself - the script handles it the same way
               Move-InboxToArchive.ps1 and the SharePoint reporting scripts do:

                 1. An app-only Graph session you already established is used
                    as-is.
                 2. -ClientId with -ClientSecret or -CertificateThumbprint uses
                    your own App Registration (needs Mail.ReadWrite application
                    permission, admin consent granted).
                 3. Otherwise the script connects interactively, creates a
                    short-lived temporary App Registration, self-grants it
                    Mail.ReadWrite application permission (the delegated role
                    does the consent, so no separate consent screen), takes an
                    app-only token with it, and removes the temporary app again
                    when the run finishes. Needs Global Administrator or
                    Privileged Role Administrator for that one-time setup, plus
                    the Microsoft.Graph.Applications module.

               Mail.ReadWrite (application) grants access to every mailbox in the
               tenant - scope the app with New-ApplicationAccessPolicy if that is
               wider than you want.

               GDAP-aware: under a GDAP session ($global:authMode -eq 'GDAP', set
               by Connect-Tenant / load.ps1) -TenantId is resolved from the
               selected customer tenant ($global:cid) when not supplied.
               $env:M365_CUSTOMER_TENANTID / $env:M365_AUTH_MODE are honored too.

    Index lag (Purview only): a message delivered minutes ago may not be
    searchable yet, so a purge run straight after delivery can report 0 hits and
    still leave the phish in place. Either wait ~30 minutes and re-run, or use the
    Graph engine against the recipients from the message trace.

    Purge covers the primary mailbox only - neither engine reaches the archive
    mailbox.

    Documented Purview limits, both handled by this script: a purge removes at
    most 10 items per mailbox per action, and a single content search purges at
    most 50,000 mailboxes. For bulk removal beyond that, Microsoft points at the
    Graph ediscoverySearch: purgeData API, which allows 100 items per location.

    Required modules: ExchangeOnlineManagement, plus Microsoft.Graph.Authentication
    for the Graph engine.
#>
[CmdletBinding()]
param(
    [string[]] $Mailbox,
    [switch]   $AllMailboxes,
    [string]   $MessageId,
    # Not named -Sender: $Sender is a PowerShell automatic variable. The alias
    # keeps -Sender working on the command line.
    [Alias('Sender')]
    [string]   $SenderAddress,
    [string]   $Subject,
    [string]   $AttachmentName,
    [string]   $BodyContains,
    [datetime] $ReceivedAfter,
    [datetime] $ReceivedBefore,
    [ValidateSet('Purview','Graph')]
    [string]   $Engine,
    [ValidateSet('Recycle','SoftDelete','HardDelete')]
    [string]   $DeleteType = 'SoftDelete',
    [switch]   $Apply,
    [string]   $SearchName,
    [switch]   $KeepSearch,
    [switch]   $VerifyWithGraph,
    [int]      $MaxPurgeRounds = 10,
    [int]      $MaxMessagesPerMailbox = 500,
    [int]      $TimeoutMinutes = 30,
    [string]   $OutputPath,
    [string]   $TenantId,
    [string]   $ClientId,
    [string]   $ClientSecret,
    [string]   $CertificateThumbprint
)

# ── Criteria validation ───────────────────────────────────────────────────────
# Without at least one content selector this would match - and delete - every
# message in every mailbox. Refuse rather than trust the next keystroke.
if (-not ($MessageId -or $SenderAddress -or $Subject -or $AttachmentName -or $BodyContains)) {
    throw "Specify at least one of -MessageId, -SenderAddress, -Subject, -AttachmentName or -BodyContains. A date range on its own matches every message."
}
if ($ReceivedAfter -and $ReceivedBefore -and $ReceivedAfter -ge $ReceivedBefore) {
    throw "-ReceivedAfter must be earlier than -ReceivedBefore."
}

if (-not $Engine) { $Engine = if ($Mailbox) { 'Graph' } else { 'Purview' } }

if ($Engine -eq 'Purview' -and $DeleteType -eq 'Recycle') {
    throw "Purview purge only supports SoftDelete and HardDelete - it cannot move items to Deleted Items. Use -DeleteType SoftDelete, or -Engine Graph with -Mailbox for a Recycle."
}

if ($Engine -eq 'Graph') {
    if (-not $Mailbox -and -not $AllMailboxes) {
        throw "The Graph engine needs targets: pass -Mailbox, or -AllMailboxes to sweep the whole tenant."
    }
    if ($DeleteType -eq 'HardDelete') {
        throw "HardDelete is not reachable over the Graph mail API - it cannot write to Recoverable Items\Purges. Use -Engine Purview for a hard delete, or -DeleteType SoftDelete here."
    }
    if ($BodyContains) {
        throw "-BodyContains is Purview-only; matching it over Graph would mean downloading every message body. Use -Engine Purview."
    }
}

# Normalise once: both engines want the angle-bracket form of the MessageId.
if ($MessageId) {
    $MessageId = $MessageId.Trim()
    if ($MessageId -notmatch '^<') { $MessageId = "<$MessageId" }
    if ($MessageId -notmatch '>$') { $MessageId = "$MessageId>" }
}

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Header ────────────────────────────────────────────────────────────────────
$modeLabel = if ($Apply) { "APPLY - messages will be deleted ($DeleteType)" } else { 'DRY RUN - nothing will be deleted' }
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Remove Phishing Message" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Mode      : $modeLabel" -ForegroundColor $(if ($Apply) { 'Red' } else { 'Yellow' })
Write-Host "  Engine    : $Engine" -ForegroundColor DarkGray
if ($Mailbox) {
    Write-Host "  Mailboxes : $($Mailbox -join ', ')" -ForegroundColor DarkGray
} else {
    Write-Host "  Mailboxes : all mailboxes in the tenant" -ForegroundColor DarkGray
}
if ($MessageId)      { Write-Host "  MessageId : $MessageId" -ForegroundColor DarkGray }
if ($SenderAddress)  { Write-Host "  Sender    : $SenderAddress" -ForegroundColor DarkGray }
if ($Subject)        { Write-Host "  Subject   : $Subject" -ForegroundColor DarkGray }
if ($AttachmentName) { Write-Host "  Attachment: $AttachmentName" -ForegroundColor DarkGray }
if ($BodyContains)   { Write-Host "  Body      : $BodyContains" -ForegroundColor DarkGray }
if ($ReceivedAfter)  { Write-Host "  After     : $($ReceivedAfter.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor DarkGray }
if ($ReceivedBefore) { Write-Host "  Before    : $($ReceivedBefore.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor DarkGray }
Write-Host ""

$results             = [System.Collections.Generic.List[PSObject]]::new()
$script:ConnectedIpps    = $false
$script:ConnectedExo     = $false
$script:ReusedIppsSession = $false
$script:TotalPurged       = 0
$script:Truncated     = $false

# ══════════════════════════════════════════════════════════════════════════════
#  Graph connection
#
#  Same three-way pattern as Move-InboxToArchive.ps1 and the SharePoint reporting
#  scripts: reuse an app-only session, or use your own app, or build a short-lived
#  one and tear it down afterwards.
# ══════════════════════════════════════════════════════════════════════════════
$script:TempAppObjectId = $null
$script:AppOnlyHeaders  = $null
$script:TokenBody       = $null
$script:TokenTenantId   = $null
$script:TokenExpiry     = [datetime]::MinValue
$script:GraphConnected  = $false

function Remove-TempApp {
    <#
        Deleting the temp app needs the delegated session's
        Application.ReadWrite.All - the app-only token only ever holds
        Mail.ReadWrite - so this must run before that session is disconnected.
    #>
    if (-not $script:TempAppObjectId) { return }
    Write-Host "  Removing temporary App Registration..." -ForegroundColor DarkGray
    try {
        Remove-MgApplication -ApplicationId $script:TempAppObjectId -ErrorAction Stop
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
    $script:AppOnlyHeaders = @{ Authorization = "Bearer $($resp.access_token)" }
    $script:TokenExpiry    = (Get-Date).AddSeconds($resp.expires_in - 300)
}

function Invoke-Graph {
    <#
        One entry point for every Graph call, so the temporary-app mode (a raw
        bearer token, because the delegated session has to stay connected to clean
        the app up afterwards) and a normal Connect-MgGraph session look identical
        to callers.
    #>
    param(
        [string] $Method = 'GET',
        [Parameter(Mandatory)] [string] $Uri,
        $Body,
        [string] $ContentType = 'application/json'
    )

    if ($script:AppOnlyHeaders) {
        Update-AppOnlyToken
        $params = @{ Method = $Method; Uri = $Uri; Headers = $script:AppOnlyHeaders; ErrorAction = 'Stop' }
        if ($Body) { $params['Body'] = $Body; $params['ContentType'] = $ContentType }
        return Invoke-RestMethod @params
    }

    $params = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
    if ($Body) { $params['Body'] = $Body; $params['ContentType'] = $ContentType }
    return Invoke-MgGraphRequest @params
}

function Resolve-EffectiveTenantId {
    # GDAP-aware, matching Get-SharePointStorageReport.ps1 / Move-InboxToArchive.ps1.
    if ($TenantId) { return $TenantId }
    try {
        $gdap = ($global:authMode -and ([string]$global:authMode).ToUpperInvariant() -eq 'GDAP') -or
                ($env:M365_AUTH_MODE -and ([string]$env:M365_AUTH_MODE).ToUpperInvariant() -eq 'GDAP')
        if ($gdap -and $global:cid)      { return [string]$global:cid }
        if ($env:M365_CUSTOMER_TENANTID) { return [string]$env:M365_CUSTOMER_TENANTID }
    } catch {}
    return $null
}

function Connect-GraphForMail {
    <#
        Establishes app-only Mail.ReadWrite, one way or another. Returns $true when
        Graph is usable. Callers decide whether failure is fatal: the Graph engine
        cannot run without it, verification just skips.
    #>
    if ($script:GraphConnected -or $script:AppOnlyHeaders) { return $true }

    if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        Write-Warning "Microsoft.Graph.Authentication is not available. Install-Module Microsoft.Graph.Authentication"
        return $false
    }

    # 1. An app-only session the caller already established.
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if ($ctx -and $ctx.AuthType -eq 'AppOnly') {
        Write-Host "  [OK]   Using the existing app-only Graph session." -ForegroundColor DarkGray
        $script:GraphConnected = $true
        return $true
    }

    $effectiveTenantId = Resolve-EffectiveTenantId

    # 2. Your own app registration.
    if ($ClientId) {
        if (-not $effectiveTenantId) {
            Write-Warning "-ClientId needs -TenantId (or a resolvable GDAP customer tenant)."
            return $false
        }
        try {
            if ($CertificateThumbprint) {
                Connect-MgGraph -ClientId $ClientId -TenantId $effectiveTenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
            } elseif ($ClientSecret) {
                $cred = [System.Management.Automation.PSCredential]::new($ClientId, (ConvertTo-SecureString $ClientSecret -AsPlainText -Force))
                Connect-MgGraph -ClientId $ClientId -TenantId $effectiveTenantId -ClientSecretCredential $cred -NoWelcome -ErrorAction Stop
            } else {
                Write-Warning "-ClientId needs -ClientSecret or -CertificateThumbprint."
                return $false
            }
        } catch {
            Write-Warning "Could not connect with the supplied app credentials: $($_.Exception.Message)"
            return $false
        }
        Write-Host "  [OK]   Connected with the supplied app credentials." -ForegroundColor DarkGray
        $script:GraphConnected = $true
        return $true
    }

    # 3. Build a short-lived app, use it, remove it at the end.
    if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Applications')) {
        Write-Warning "Microsoft.Graph.Applications is needed to create the temporary app registration. Install it, or pass -ClientId for your own app."
        return $false
    }

    try {
        Write-Host "  Connecting interactively to set up a temporary app registration..." -ForegroundColor Cyan
        Write-Host "  Required role: Global Administrator or Privileged Role Administrator (one-time setup)" -ForegroundColor DarkGray
        $connectParams = @{
            Scopes      = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
            NoWelcome   = $true
            ErrorAction = 'Stop'
        }
        if ($effectiveTenantId) { $connectParams['TenantId'] = $effectiveTenantId }
        Connect-MgGraph @connectParams
        $script:GraphConnected = $true

        if (-not $effectiveTenantId) { $effectiveTenantId = (Get-MgContext).TenantId }

        $appName = "PhishPurge-Temp-$(Get-Date -Format 'yyyyMMddHHmmss')"
        Write-Host "  Creating temporary App Registration '$appName'..." -ForegroundColor Cyan
        $app = New-MgApplication -DisplayName $appName -ErrorAction Stop
        $script:TempAppObjectId = $app.Id

        $sp      = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop
        $graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ErrorAction Stop
        $appRole = $graphSp.AppRoles | Where-Object { $_.Value -eq 'Mail.ReadWrite' -and $_.AllowedMemberTypes -contains 'Application' }
        if (-not $appRole) { throw "Could not resolve the Mail.ReadWrite application role." }

        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -ResourceId $graphSp.Id -AppRoleId $appRole.Id -ErrorAction Stop | Out-Null
        Write-Host "  [OK]   Mail.ReadWrite (application) granted." -ForegroundColor DarkGray

        $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
            displayName = 'temp'
            endDateTime = (Get-Date).AddHours(2)
        } -ErrorAction Stop

        $script:TokenBody = @{
            grant_type    = 'client_credentials'
            scope         = 'https://graph.microsoft.com/.default'
            client_id     = $app.AppId
            client_secret = $secret.SecretText
        }
        $script:TokenTenantId = $effectiveTenantId

        # A brand new app registration is not instantly usable - retry the first
        # token while it propagates.
        $ok = $false
        for ($i = 1; $i -le 6; $i++) {
            try { Update-AppOnlyToken; $ok = $true; break }
            catch {
                if ($i -lt 6) {
                    Write-Host "  Waiting for app registration propagation (attempt $i/6)..." -ForegroundColor DarkGray
                    Start-Sleep -Seconds 5
                }
            }
        }
        if (-not $ok) { throw "Could not obtain an app-only token after propagation retries." }

        Write-Host "  [OK]   App-only token obtained (temporary app)." -ForegroundColor DarkGray
        return $true

    } catch {
        Write-Warning "Temporary app setup failed: $($_.Exception.Message)"
        Remove-TempApp
        return $false
    }
}

# ══════════════════════════════════════════════════════════════════════════════
#  Purview engine
# ══════════════════════════════════════════════════════════════════════════════
function Invoke-PurviewPurge {

    # ── KQL query ─────────────────────────────────────────────────────────────
    # Each clause is ANDed; every selector the caller gave has to match.
    $clauses = [System.Collections.Generic.List[string]]::new()
    if ($MessageId)      { $clauses.Add("(InternetMessageId:`"$MessageId`")") }
    if ($SenderAddress)  { $clauses.Add("(From:`"$SenderAddress`")") }
    if ($Subject)        { $clauses.Add("(Subject:$(ConvertTo-KqlTerm $Subject -Field 'Subject'))") }
    if ($AttachmentName) { $clauses.Add("(Attachment:$(ConvertTo-KqlTerm $AttachmentName -Field 'Attachment'))") }
    if ($BodyContains)   { $clauses.Add("(Body:`"$BodyContains`")") }
    # The index stores Received in UTC; converting keeps the window honest for
    # operators who are not on UTC themselves.
    if ($ReceivedAfter)  { $clauses.Add("(Received>=$($ReceivedAfter.ToUniversalTime().ToString('yyyy-MM-dd')))") }
    if ($ReceivedBefore) { $clauses.Add("(Received<=$($ReceivedBefore.ToUniversalTime().ToString('yyyy-MM-dd')))") }
    $query = $clauses -join ' AND '

    Write-Host "  KQL       : $query" -ForegroundColor DarkGray
    Write-Host ""

    # ── Connection ────────────────────────────────────────────────────────────
    # Content Search runs on a separate backend that the plain IPPS connection no
    # longer reaches: without -EnableSearchOnlySession the cmdlets are present but
    # Start-ComplianceSearch fails at initialisation. The switch arrived in
    # ExchangeOnlineManagement 3.9.0, so it is passed only when supported.
    if (-not (Get-Command New-ComplianceSearch -ErrorAction SilentlyContinue)) {
        Write-Host "  Connecting to Security & Compliance PowerShell..." -ForegroundColor DarkGray
        $ippsParams = @{ ErrorAction = 'Stop' }
        if ($TenantId) { $ippsParams['Organization'] = $TenantId }
        $connectCmd = Get-Command Connect-IPPSSession -ErrorAction Stop
        if ($connectCmd.Parameters.ContainsKey('EnableSearchOnlySession')) {
            $ippsParams['EnableSearchOnlySession'] = $true
        } else {
            Write-Warning "ExchangeOnlineManagement $($connectCmd.Module.Version) has no -EnableSearchOnlySession. Content Search may fail to start; update with: Install-Module ExchangeOnlineManagement -Force"
        }
        Connect-IPPSSession @ippsParams
        $script:ConnectedIpps = $true
    } else {
        # Reusing someone else's session: it may well have been opened without
        # the switch, and that cannot be repaired from inside this process.
        $script:ReusedIppsSession = $true
    }

    if (-not $SearchName) { $SearchName = "Phish_$(Get-Date -Format 'yyyyMMdd_HHmmss')" }
    $purgeActionName = "$($SearchName)_Purge"

    # ── Create and run the search ─────────────────────────────────────────────
    $newParams = @{
        Name               = $SearchName
        ContentMatchQuery  = $query
        ErrorAction        = 'Stop'
    }
    if ($Mailbox) { $newParams['ExchangeLocation'] = $Mailbox }
    else          { $newParams['ExchangeLocation'] = 'All' }

    Write-Host "  Creating Content Search '$SearchName'..." -ForegroundColor DarkGray
    New-ComplianceSearch @newParams | Out-Null

    $round           = 0
    $roundsNeeded    = 1     # replaced by a real plan once the first search lands
    $purgedMailboxes = @()
    $searchFailed = $false
    # A re-started search keeps reporting the previous run's Completed status and
    # SuccessResults for a few seconds. Carrying the last JobEndTime forward lets
    # the wait tell a fresh result from a stale one.
    $lastJobEnd   = $null

    # The whole run sits in try/finally: a search that is created and then fails
    # would otherwise be left behind in the tenant, and Purview refuses to reuse
    # a name, so orphans pile up run after run.
    try {
    while ($true) {
        $round++

        try {
            Start-ComplianceSearch -Identity $SearchName -ErrorAction Stop
        } catch {
            if ("$($_.Exception.Message)" -match 'EnableSearchOnlySession') {
                $hint = if ($script:ReusedIppsSession) {
                    "This session was connected without it. Close this PowerShell window, open a new one and let the script connect itself, or run: Connect-IPPSSession -EnableSearchOnlySession"
                } else {
                    "Update the module - the switch needs ExchangeOnlineManagement 3.9.0 or higher: Install-Module ExchangeOnlineManagement -Force"
                }
                throw "Content Search needs an IPPS session opened with -EnableSearchOnlySession. $hint"
            }
            throw
        }

        $search = Wait-ForComplianceState -Getter { Get-ComplianceSearch -Identity $SearchName -ErrorAction Stop } `
                                          -Label "search round $round" -NewerThan $lastJobEnd
        if (-not $search) { $searchFailed = $true; break }
        $lastJobEnd = $search.JobEndTime

        $hits = Read-ComplianceSuccessResults -SuccessResults $search.SuccessResults
        $hitTotal = ($hits | Measure-Object -Property ItemCount -Sum).Sum
        if (-not $hitTotal) { $hitTotal = 0 }

        if ($round -eq 1) {
            Write-Host ""
            if ($hitTotal -eq 0) {
                Write-Host "  No matching items found in the index." -ForegroundColor Yellow
                Write-Host "  If the message was delivered in the last ~30 minutes it may not be" -ForegroundColor DarkGray
                Write-Host "  indexed yet. Re-run later, or use -Engine Graph with -Mailbox." -ForegroundColor DarkGray
                Write-Host ""
            } else {
                Write-Host "  Found $hitTotal item(s) across $($hits.Count) mailbox(es):" -ForegroundColor Cyan
                Write-Host ""
                foreach ($h in ($hits | Sort-Object -Property ItemCount -Descending)) {
                    Write-Host ("    {0,-45} {1,4} item(s)   {2}" -f $h.Location, $h.ItemCount, (Format-ByteSize $h.TotalSize)) -ForegroundColor Yellow
                }
                Write-Host ""

                # The purge cap is per mailbox, so the busiest mailbox decides how
                # many rounds are needed - not the tenant-wide total.
                if ($hits.Count -gt 50000) {
                    Write-Warning "A single content search purges at most 50,000 mailboxes; this search matched $($hits.Count). Narrow the scope with -Mailbox and run it in batches."
                    $script:Truncated = $true
                }

                $maxPerMailbox = ($hits | Measure-Object -Property ItemCount -Maximum).Maximum
                if (-not $maxPerMailbox) { $maxPerMailbox = 0 }
                $roundsNeeded  = [int][Math]::Ceiling($maxPerMailbox / 10.0)

                if ($Apply -and $roundsNeeded -gt 1) {
                    Write-Host "  Busiest mailbox holds $maxPerMailbox item(s) and a purge takes 10 per mailbox per round, so $roundsNeeded round(s) are needed." -ForegroundColor DarkGray
                    Write-Host ""
                }
                if ($Apply -and $roundsNeeded -gt $MaxPurgeRounds) {
                    Write-Warning "$roundsNeeded round(s) are needed but -MaxPurgeRounds is $MaxPurgeRounds. Raise it, or re-run afterwards to clear the remainder."
                    $script:Truncated = $true
                }

                # Purview can go no finer than this. Content Search reports counts
                # per mailbox, and the preview action that used to return the
                # individual messages was retired in the cloud in May 2025 - it is
                # documented as on-premises only. For subject-level detail, run the
                # Graph engine against these mailboxes.
                if (-not $Apply) {
                    Write-Host "  Purview reports counts per mailbox, not individual messages." -ForegroundColor DarkGray
                    Write-Host "  For per-message detail, re-run with: -Engine Graph -Mailbox <the addresses above>" -ForegroundColor DarkGray
                    Write-Host ""
                }
            }

            $purgedMailboxes = @($hits | Select-Object -ExpandProperty Location)

            # Record what the search saw, so the CSV is useful even on a dry run.
            foreach ($h in $hits) {
                $results.Add([PSCustomObject]@{
                    Mailbox    = $h.Location
                    Engine     = 'Purview'
                    ItemCount  = $h.ItemCount
                    TotalSize  = $h.TotalSize
                    Subject    = ''
                    Received   = ''
                    Folder     = ''
                    Action     = if ($Apply) { $DeleteType } else { 'DryRun' }
                    Status     = 'Matched'
                    Detail     = "Search '$SearchName'"
                })
            }
        } else {
            # Later rounds: show what the index still reports, so a multi-round
            # purge is not a silent wait. This is NOT a reliable count of what is
            # left - the index lags a purge by up to ~30 minutes, so items already
            # purged keep showing up here for a while.
            Write-Host ("  Index still reports {0} item(s) across {1} mailbox(es) (it lags a purge, so this is not proof anything remains)" -f $hitTotal, $hits.Count) -ForegroundColor DarkGray
        }

        if ($hitTotal -eq 0) { break }
        if (-not $Apply)     { break }
        # The index cannot say when the purge is done, so the plan made from the
        # first search decides. Looping until the index goes quiet would purge the
        # same items over and over and then report a false truncation.
        if ($round -gt $roundsNeeded)   { break }
        if ($round -gt $MaxPurgeRounds) { break }

        # ── Purge ─────────────────────────────────────────────────────────────
        # Action names are derived from the search name and only one action can
        # be outstanding, so the preview and the previous round's purge both have
        # to go before the next purge can be created.
        try {
            Remove-ComplianceSearchAction -Identity $purgeActionName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        } catch {}

        Write-Host "  Purge round $round : removing up to 10 item(s) per mailbox ($DeleteType)..." -ForegroundColor Yellow
        New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType $DeleteType -Confirm:$false -ErrorAction Stop | Out-Null

        $action = Wait-ForComplianceState -Getter { Get-ComplianceSearchAction -Identity $purgeActionName -Details -ErrorAction Stop } `
                                          -Label "purge round $round"
        if (-not $action) { $searchFailed = $true; break }

        # The action's own results are the only trustworthy account of what was
        # removed - the search index lags and cannot confirm it.
        $purgedNow = Read-CompliancePurgeResults -Results $action.Results
        if ($null -ne $purgedNow) {
            $script:TotalPurged += $purgedNow
            Write-Host "    round $round done - $purgedNow item(s) purged" -ForegroundColor DarkGray
        } else {
            Write-Host "    round $round done - service reported no item detail" -ForegroundColor DarkGray
        }
    }

    if ($Apply -and -not $searchFailed) {
        Write-Host ""
        if ($script:TotalPurged -gt 0) {
            Write-Host "  Purge finished after $($round - 1) round(s); the service reported $($script:TotalPurged) item(s) purged." -ForegroundColor Green
        } else {
            Write-Host "  Purge finished after $($round - 1) round(s)." -ForegroundColor Green
        }
        Write-Host "  The search index lags a purge by up to ~30 minutes, so re-running this" -ForegroundColor DarkGray
        Write-Host "  script straight away will still show the items. Verify later, or check" -ForegroundColor DarkGray
        Write-Host "  a mailbox directly." -ForegroundColor DarkGray
        foreach ($r in $results) {
            if ($r.Status -eq 'Matched') { $r.Status = 'Purged' }
        }
    }

        if ($VerifyWithGraph -and $Apply -and $purgedMailboxes.Count -gt 0) {
            Test-PurgeWithGraph -Mailboxes $purgedMailboxes
        }
    }
    # ── Cleanup ───────────────────────────────────────────────────────────────
    finally {
        if (-not $KeepSearch) {
            try { Remove-ComplianceSearchAction -Identity $purgeActionName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}
            try { Remove-ComplianceSearch -Identity $SearchName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}
        } else {
            Write-Host "  Content Search kept: '$SearchName'" -ForegroundColor DarkGray
        }
    }
}

function Wait-ForComplianceState {
    <#
        Polls a compliance object until it reports Completed. Returns the object,
        or $null when it failed or ran past -TimeoutMinutes.

        -NewerThan guards against a re-started search that is still serving the
        previous run's Completed status: a result whose JobEndTime has not moved
        past the given moment is treated as stale and polling continues.
    #>
    param(
        [scriptblock]        $Getter,
        [string]             $Label,
        [nullable[datetime]] $NewerThan
    )

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 5
        try { $obj = & $Getter } catch { continue }
        switch ("$($obj.Status)") {
            'Completed' {
                if ($NewerThan -and $obj.JobEndTime -and [datetime]$obj.JobEndTime -le $NewerThan) {
                    continue   # previous run's result, the restart has not landed yet
                }
                return $obj
            }
            'Failed'    {
                Write-Warning "$Label failed: $($obj.Errors)"
                return $null
            }
        }
    }
    Write-Warning "$Label did not complete within $TimeoutMinutes minute(s)."
    return $null
}

function ConvertTo-KqlTerm {
    <#
        Turns a user-supplied fragment into a KQL term.

        A quoted phrase already matches anywhere in the field, so a fragment of a
        subject needs no wildcards at all - "kick-off meeting" finds every subject
        containing those words in that order. What KQL cannot do is match inside a
        word: a leading wildcard is unsupported, and a wildcard is inert inside
        quotes. So a trailing wildcard is preserved on a single unquoted term
        (prefix match), a leading one is reported as dropped, and anything with a
        space falls back to a phrase.
    #>
    param(
        [string] $Value,
        [string] $Field = 'term'
    )

    $v           = $Value.Trim()
    $hadLeading  = $v.StartsWith('*')
    $hadTrailing = $v.EndsWith('*')
    $core        = $v.Trim('*').Trim()

    if ($hadLeading) {
        Write-Warning "KQL cannot match a leading wildcard, so -$Field '$Value' matches from the start of a word. A quoted fragment already matches anywhere in the field; for true substring matching use -Engine Graph with -Mailbox."
    }

    if ($core -match '\s') {
        if ($hadTrailing) {
            Write-Warning "A wildcard is inert inside a quoted phrase, so -$Field '$Value' is matched as the phrase '$core'."
        }
        return '"{0}"' -f ($core -replace '"', '')
    }
    if ($hadTrailing) { return "$core*" }
    return '"{0}"' -f ($core -replace '"', '')
}

function Test-PurgeWithGraph {
    <#
        Re-asks the Graph engine's question against the mailboxes the search hit,
        after the purge. Anything still returned is genuinely still in the
        mailbox - unlike the Content Search index, Graph has no lag.

        Reports rather than throws: a purge that already happened should not be
        reported as failed because the verification could not run.
    #>
    param([string[]] $Mailboxes)

    Write-Host ""
    Write-Host "  -- Verifying over Graph --" -ForegroundColor Cyan

    if (-not (Connect-GraphForMail)) {
        Write-Warning "Cannot verify: no app-only Graph access could be established (see .NOTES). The purge itself is unaffected - nothing here says it failed."
        return
    }

    $stillPresent = [System.Collections.Generic.List[PSObject]]::new()
    $unchecked    = 0

    foreach ($mbx in $Mailboxes) {
        try {
            $left = @(Get-GraphMatchingMessage -Mailbox $mbx)
        } catch {
            Write-Host ("    {0,-45} could not check: {1}" -f $mbx, $_.Exception.Message) -ForegroundColor DarkYellow
            $unchecked++
            continue
        }

        if ($left.Count -eq 0) {
            Write-Host ("    {0,-45} clean" -f $mbx) -ForegroundColor Green
        } else {
            Write-Host ("    {0,-45} {1} still present" -f $mbx, $left.Count) -ForegroundColor Red
            $stillPresent.Add([PSCustomObject]@{ Mailbox = $mbx; Count = $left.Count })
        }
    }

    Write-Host ""
    if ($stillPresent.Count -gt 0) {
        $total = ($stillPresent | Measure-Object -Property Count -Sum).Sum
        Write-Warning "$total item(s) still present in $($stillPresent.Count) mailbox(es) after the purge. Re-run to clear the remainder, or use -Engine Graph -Mailbox to delete them directly."
        $script:Truncated = $true
        foreach ($r in $results) {
            if ($r.Status -eq 'Purged' -and ($stillPresent.Mailbox -contains $r.Mailbox)) {
                $r.Status = 'StillPresent'
            }
        }
    } elseif ($unchecked -eq $Mailboxes.Count) {
        Write-Warning "No mailbox could be checked - verification proved nothing."
    } else {
        $checked = $Mailboxes.Count - $unchecked
        Write-Host "  Verified gone from $checked of $($Mailboxes.Count) mailbox(es)." -ForegroundColor Green
        if ($unchecked -gt 0) {
            Write-Warning "$unchecked mailbox(es) could not be checked; those are unverified, not confirmed clean."
        }
    }
}

function Format-ByteSize {
    param([int64] $Bytes)
    if     ($Bytes -ge 1MB) { '{0:N1} MB' -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { '{0:N0} KB' -f ($Bytes / 1KB) }
    else                    { "$Bytes B" }
}

function Read-CompliancePurgeResults {
    <#
        A purge action reports its outcome as free text. The exact shape is not
        contractual, so an unrecognised payload returns $null - reported as "no
        item detail" rather than guessed at as a number.
    #>
    param([string] $Results)

    if (-not $Results) { return $null }
    $total = 0
    $found = $false
    foreach ($m in [regex]::Matches($Results, 'Item count:\s*(?<count>\d+)')) {
        $total += [int] $m.Groups['count'].Value
        $found  = $true
    }
    if ($found) { return $total }
    return $null
}

function Read-ComplianceSuccessResults {
    <#
        SuccessResults comes back as free text, one line per location:
          {Location: user@contoso.com, Item count: 3, Total size: 12345}
        Only locations with a non-zero item count are returned.
    #>
    param([string] $SuccessResults)

    $out = [System.Collections.Generic.List[PSObject]]::new()
    if (-not $SuccessResults) { return $out }

    foreach ($m in [regex]::Matches($SuccessResults, 'Location:\s*(?<loc>[^,]+),\s*Item count:\s*(?<count>\d+),\s*Total size:\s*(?<size>\d+)')) {
        $count = [int] $m.Groups['count'].Value
        if ($count -le 0) { continue }
        $out.Add([PSCustomObject]@{
            Location  = $m.Groups['loc'].Value.Trim()
            ItemCount = $count
            TotalSize = [int64] $m.Groups['size'].Value
        })
    }
    return $out
}

# ══════════════════════════════════════════════════════════════════════════════
#  Graph engine
# ══════════════════════════════════════════════════════════════════════════════
function Invoke-GraphPurge {

    if (-not (Connect-GraphForMail)) {
        throw "The Graph engine needs app-only Mail.ReadWrite and none could be established. See .NOTES for the three ways to supply it."
    }

    # ── Target mailboxes ──────────────────────────────────────────────────────
    # Mailboxes come from Exchange Online, not Graph /users: enumerating users
    # would need User.Read.All on top of Mail.ReadWrite, and Get-EXOMailbox
    # answers the question that actually matters - which mailboxes exist.
    $targets = @($Mailbox)
    if ($AllMailboxes) {
        Write-Host "  Enumerating mailboxes..." -ForegroundColor DarkGray
        if (-not (Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue)) {
            $exoParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
            if ($TenantId) { $exoParams['Organization'] = $TenantId }
            Connect-ExchangeOnline @exoParams
            $script:ConnectedExo = $true
        }
        $targets = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox |
                        ForEach-Object { $_.PrimarySmtpAddress } | Where-Object { $_ })
        Write-Host "  $($targets.Count) mailbox(es) to check." -ForegroundColor DarkGray
        Write-Host ""
    }

    $folderNames = @{}

    foreach ($mbx in $targets) {
        try {
            $messages = @(Get-GraphMatchingMessage -Mailbox $mbx)
        } catch {
            Write-Host ("  {0,-45} ERROR  {1}" -f $mbx, $_.Exception.Message) -ForegroundColor Red
            $results.Add([PSCustomObject]@{
                Mailbox = $mbx; Engine = 'Graph'; ItemCount = 0; TotalSize = ''
                Subject = ''; Received = ''; Folder = ''
                Action  = 'None'; Status = 'Error'; Detail = $_.Exception.Message
            })
            continue
        }

        if ($messages.Count -eq 0) {
            Write-Host ("  {0,-45} no match" -f $mbx) -ForegroundColor DarkGray
            continue
        }

        Write-Host ("  {0,-45} {1} match(es)" -f $mbx, $messages.Count) -ForegroundColor Cyan

        foreach ($msg in $messages) {
            $folder = Resolve-GraphFolderName -Mailbox $mbx -FolderId $msg.parentFolderId -Cache $folderNames
            $row = [PSCustomObject]@{
                Mailbox   = $mbx
                Engine    = 'Graph'
                ItemCount = 1
                TotalSize = ''
                Subject   = $msg.subject
                Received  = if ($msg.receivedDateTime) { ([datetime]$msg.receivedDateTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
                Folder    = $folder
                Action    = if ($Apply) { $DeleteType } else { 'DryRun' }
                Status    = 'Matched'
                Detail    = $msg.internetMessageId
            }

            if ($Apply) {
                try {
                    Remove-GraphMessage -Mailbox $mbx -MessageId $msg.id
                    $row.Status = if ($DeleteType -eq 'Recycle') { 'Recycled' } else { 'SoftDeleted' }
                } catch {
                    $row.Status = 'Error'
                    $row.Detail = $_.Exception.Message
                }
            }

            $color = switch ($row.Status) {
                'Error'   { 'Red' }
                'Matched' { 'Yellow' }
                default   { 'Green' }
            }
            Write-Host ("      [{0}] {1}  |  {2}  |  {3}" -f $row.Status,
                            $row.Received,
                            $folder,
                            ("$($msg.subject)" -replace '^(.{55}).+$', '$1...')) -ForegroundColor $color

            $results.Add($row)
        }
    }
}

function Get-GraphMatchingMessage {
    <#
        Returns the messages in one mailbox that match the run's criteria.

        Shared by the Graph engine and by -VerifyWithGraph, so verification asks
        exactly the same question the deletion answered - a check built on a
        second, slightly different query would prove nothing.

        Server-side $filter carries what Graph can filter on; -Subject and
        -AttachmentName are matched client-side afterwards.
    #>
    param([string] $Mailbox)

    $filters = [System.Collections.Generic.List[string]]::new()
    if ($MessageId) {
        # OData string literals escape a single quote by doubling it.
        $filters.Add("internetMessageId eq '$($MessageId.Replace("'", "''"))'")
    }
    if ($SenderAddress) {
        $filters.Add("from/emailAddress/address eq '$($SenderAddress.Replace("'", "''"))'")
    }
    if ($AttachmentName) { $filters.Add("hasAttachments eq true") }
    if ($ReceivedAfter)  { $filters.Add("receivedDateTime ge $($ReceivedAfter.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))") }
    if ($ReceivedBefore) { $filters.Add("receivedDateTime le $($ReceivedBefore.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))") }

    $select = 'id,internetMessageId,subject,receivedDateTime,from,hasAttachments,parentFolderId'
    $uri = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($Mailbox))/messages?`$select=$select&`$top=100"
    if ($filters.Count -gt 0) {
        $uri += "&`$filter=$([uri]::EscapeDataString($filters -join ' and '))"
    }

    $messages = @(Get-GraphPaged -Uri $uri -MaxItems $MaxMessagesPerMailbox)

    if ($messages.Count -ge $MaxMessagesPerMailbox) {
        Write-Warning "$Mailbox hit the -MaxMessagesPerMailbox cap ($MaxMessagesPerMailbox); there may be more matches."
        $script:Truncated = $true
    }

    if ($Subject) {
        $pattern = if ($Subject -match '\*') { $Subject } else { "*$Subject*" }
        $messages = @($messages | Where-Object { $_.subject -like $pattern })
    }
    if ($AttachmentName) {
        $pattern = if ($AttachmentName -match '\*') { $AttachmentName } else { "*$AttachmentName*" }
        $messages = @($messages | Where-Object { Test-GraphAttachmentName -Mailbox $Mailbox -MessageId $_.id -Pattern $pattern })
    }

    return $messages
}

function Get-GraphPaged {
    param(
        [string] $Uri,
        [int]    $MaxItems = 1000
    )
    $items = [System.Collections.Generic.List[PSObject]]::new()
    $next  = $Uri
    while ($next -and $items.Count -lt $MaxItems) {
        $resp = Invoke-Graph -Method GET -Uri $next
        foreach ($v in @($resp.value)) {
            $items.Add($v)
            if ($items.Count -ge $MaxItems) { break }
        }
        $next = $resp.'@odata.nextLink'
    }
    return $items
}

function Test-GraphAttachmentName {
    param(
        [string] $Mailbox,
        [string] $MessageId,
        [string] $Pattern
    )
    try {
        $atts = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($Mailbox))/messages/$MessageId/attachments?`$select=name"
        foreach ($a in @($atts.value)) {
            if ($a.name -like $Pattern) { return $true }
        }
    } catch {}
    return $false
}

function Resolve-GraphFolderName {
    <#
        Folder ids are opaque and repeat across messages, so each one is looked
        up once per run and cached.
    #>
    param(
        [string]    $Mailbox,
        [string]    $FolderId,
        [hashtable] $Cache
    )
    if (-not $FolderId) { return '' }
    $key = "$Mailbox|$FolderId"
    if ($Cache.ContainsKey($key)) { return $Cache[$key] }

    $name = $FolderId
    try {
        $folderUri = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($Mailbox))/mailFolders/" + $FolderId + '?$select=displayName'
        $f = Invoke-Graph -Method GET -Uri $folderUri
        if ($f.displayName) { $name = $f.displayName }
    } catch {}

    $Cache[$key] = $name
    return $name
}

function Remove-GraphMessage {
    param(
        [string] $Mailbox,
        [string] $MessageId
    )
    $base = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($Mailbox))/messages/$MessageId"

    if ($DeleteType -eq 'Recycle') {
        # DELETE on a message moves it to Deleted Items rather than destroying it.
        Invoke-Graph -Method DELETE -Uri $base | Out-Null
    } else {
        # SoftDelete: straight into Recoverable Items\Deletions, so it leaves the
        # visible mailbox but the user can still restore it if this was a false
        # positive.
        Invoke-Graph -Method POST -Uri "$base/move" `
            -Body (@{ destinationId = 'recoverableitemsdeletions' } | ConvertTo-Json) | Out-Null
    }
}

# ══════════════════════════════════════════════════════════════════════════════
#  Run
# ══════════════════════════════════════════════════════════════════════════════
try {
    if ($Engine -eq 'Purview') { Invoke-PurviewPurge } else { Invoke-GraphPurge }
} finally {
    # The temporary app must go before the delegated session that can delete it.
    Remove-TempApp
    if ($script:GraphConnected) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    if ($script:ConnectedIpps -or $script:ConnectedExo) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
}

# ── Summary ───────────────────────────────────────────────────────────────────
$matchedItems = ($results | Measure-Object -Property ItemCount -Sum).Sum
if (-not $matchedItems) { $matchedItems = 0 }
$errorCount   = @($results | Where-Object { $_.Status -eq 'Error' }).Count
$deletedItems = ($results | Where-Object { $_.Status -in @('Purged','SoftDeleted','Recycled') } |
                    Measure-Object -Property ItemCount -Sum).Sum
if (-not $deletedItems) { $deletedItems = 0 }

Write-Host ""
Write-Host "  ------------------------------------------------" -ForegroundColor Cyan
Write-Host "  Mailboxes affected : $(@($results | Where-Object { $_.Status -ne 'Error' } | Select-Object -ExpandProperty Mailbox -Unique).Count)" -ForegroundColor Cyan
Write-Host "  Items matched      : $matchedItems" -ForegroundColor Cyan
if ($Apply) {
    Write-Host "  Items deleted      : $deletedItems ($DeleteType)" -ForegroundColor Green
} else {
    Write-Host "  Items deleted      : 0 - dry run. Re-run with -Apply to delete." -ForegroundColor Yellow
}
if ($errorCount -gt 0)  { Write-Host "  Errors             : $errorCount" -ForegroundColor Red }
if ($script:Truncated)  { Write-Host "  Results truncated  : yes - re-run to catch the remainder." -ForegroundColor Yellow }

# ── Output ────────────────────────────────────────────────────────────────────
if ($results.Count -gt 0) {
    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "PhishingRemoval_$ts.csv"
    }
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}
Write-Host ""
