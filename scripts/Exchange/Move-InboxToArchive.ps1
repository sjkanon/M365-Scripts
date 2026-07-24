#Requires -Version 5.1
<#
.SYNOPSIS
    Archive all Inbox messages of a mailbox to the Archive folder via Microsoft Graph.

.DESCRIPTION
    Moves every message in a user's Inbox to their Archive folder — the same folder
    Outlook's "Archive" button targets. Optionally restrict the scope to messages
    received within a date range. Messages are moved in batches of 20 via Microsoft
    Graph's $batch endpoint, with retry/backoff on throttling (429/503).

    Default behavior is safe preview mode: the script reports how many messages
    would be archived but makes no changes. Only when -Apply is specified are
    messages actually moved.

    Authentication:
      By default the script works app-only against any mailbox in the tenant,
      without requiring Full Access on the target mailbox. It connects
      interactively (delegated, Application.ReadWrite.All + AppRoleAssignment.ReadWrite.All),
      creates a short-lived temporary App Registration, self-grants it Mail.ReadWrite
      application permission (no separate admin-consent screen needed — the delegated
      role does that), reconnects app-only with it, and removes the temporary app
      again when the script finishes. This mirrors the temporary-app pattern used by
      Get-SharePointStorageReport.ps1 / Remove-SharePointFileVersionsByDate.ps1.
      Requires Global Administrator or Privileged Role Administrator for that one-time
      setup, and the Microsoft.Graph.Applications module.

      Pass -Delegated to skip all of that and use a plain delegated Mail.ReadWrite
      connection instead — no Entra app-creation rights needed, only Exchange Admin.
      When the target mailbox isn't the signed-in user's own, the script connects to
      Exchange Online, grants that account temporary Full Access on the mailbox,
      polls Get-MailboxPermission until it's visible (up to ~3 minutes), then keeps
      retrying the actual archive operation against a -MaxWaitMinutes deadline
      (default 65) while the grant stays in place, and only then removes it again.
      CAVEAT: Get-MailboxPermission reflects Exchange's own state almost immediately,
      but Microsoft Graph's authorization cache for delegate mailbox access can lag
      up to ~60 minutes behind that — a known Microsoft limitation. -MaxWaitMinutes
      covers that worst case in a single run; the default app-only mode (drop
      -Delegated) has no such delay if you'd rather not wait at all.

      To reuse your own existing App Registration instead of creating a temporary one,
      pass -ClientId + -TenantId + -ClientSecret (or -CertificateThumbprint); that app
      must already have Mail.ReadWrite application permission (admin consent granted).
      -TenantId accepts either the tenant ID (GUID) or a verified domain of that
      tenant (e.g. contoso.com) — whichever is easier to fill in.

      GDAP-aware: when running under a GDAP session ($global:authMode -eq 'GDAP',
      set via Connect-Tenant / load.ps1), -TenantId is resolved automatically from
      the currently selected customer tenant ($global:cid) if not supplied — same
      fallback used by Get-SharePointStorageReport.ps1 and
      Remove-SharePointFileVersionsByDate.ps1. $env:M365_CUSTOMER_TENANTID /
      $env:M365_AUTH_MODE are honored as well.

.PARAMETER Mailbox
    UPN or object ID of the mailbox whose Inbox to archive.

.PARAMETER After
    Only archive messages received on or after this date.

.PARAMETER Before
    Only archive messages received before this date.

.PARAMETER TenantId
    Entra ID tenant — either the tenant ID (GUID) or a verified domain of the tenant
    (e.g. contoso.com). Optional if already connected, or resolvable from a GDAP
    customer tenant context. Required when using -ClientId if not resolvable.

.PARAMETER ClientId
    Existing App Registration client ID for app-only auth — skips the automatic
    temporary app. Use with -TenantId and -ClientSecret or -CertificateThumbprint.
    That app must already have Mail.ReadWrite application permission (admin consent
    granted).

.PARAMETER ClientSecret
    Client secret for the app registration given in -ClientId.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for the app registration given in -ClientId.

.PARAMETER Delegated
    Skip the automatic temporary app-only setup and connect with a plain delegated
    Mail.ReadWrite session instead (needs Exchange Admin, not Entra app-creation
    rights). For a mailbox other than the signed-in user's own, the script grants
    that account temporary Full Access via Exchange Online, polls for it to
    propagate, archives, then removes the grant again. Note: Microsoft Graph can
    take up to ~60 minutes to honor a new Full Access grant even after Exchange
    itself shows it — a known Microsoft limitation. If this still 403s, prefer the
    default app-only mode instead of waiting longer.

.PARAMETER Apply
    Actually move the messages. Without this switch, the script only reports how
    many messages would be archived.

.PARAMETER MaxWaitMinutes
    -Delegated only. How long to keep retrying the Inbox read while waiting for
    Microsoft Graph to honor the temporary Full Access grant, before giving up and
    revoking it. Default 65 (covers Microsoft's documented worst case of ~60
    minutes). The temporary Full Access grant stays in place for the whole wait —
    re-running the script from scratch instead of raising this resets the clock,
    since each run revokes and re-grants a fresh permission.

.EXAMPLE
    # Preview — auto app-only setup, reports the count, makes no changes
    .\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com"

.EXAMPLE
    .\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Apply

.EXAMPLE
    .\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Before (Get-Date "2025-01-01") -Apply

.EXAMPLE
    .\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -After (Get-Date "2024-01-01") -Before (Get-Date "2025-01-01") -Apply

.EXAMPLE
    # Delegated — auto-grants + revokes temporary Full Access via Exchange Online
    # instead of the Entra app-only setup (needs Exchange Admin, not Global Admin)
    .\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Delegated -Apply

.EXAMPLE
    # Delegated, willing to wait out Microsoft's full ~90-minute worst case for
    # Graph to honor the Full Access grant, instead of the 65-minute default
    .\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Delegated -MaxWaitMinutes 90 -Apply

.EXAMPLE
    # Reuse an existing App Registration instead of creating a temporary one.
    # -TenantId accepts the tenant's domain instead of its GUID.
    .\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -TenantId "contoso.com" `
        -ClientId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" -ClientSecret "your-client-secret" -Apply
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $Mailbox,

    [datetime] $After,
    [datetime] $Before,
    [string] $TenantId,
    [string] $ClientId,
    [string] $ClientSecret,
    [string] $CertificateThumbprint,
    [switch] $Delegated,
    [switch] $Apply,
    [int] $MaxWaitMinutes = 65
)

if ($After -and $Before -and $After -ge $Before) {
    throw "-After must be earlier than -Before."
}

if ($ClientId -and -not $ClientSecret -and -not $CertificateThumbprint) {
    throw "-ClientId requires -ClientSecret or -CertificateThumbprint."
}

if ($ClientId -and $Delegated) {
    throw "Use either -ClientId or -Delegated, not both."
}

# ── Progress helper ──────────────────────────────────────────────────────────
function Write-ProgressHost {
    param(
        [Parameter(Mandatory)] [string] $Message,
        [ConsoleColor] $ForegroundColor = [ConsoleColor]::DarkGray
    )
    Write-Host ("  [{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message) -ForegroundColor $ForegroundColor
}

# ── Temporary App Registration cleanup ──────────────────────────────────────
# Removal needs the delegated session's Application.ReadWrite.All — the app-only
# token only ever holds Mail.ReadWrite, so this must run before that delegated
# session is disconnected.
$script:TempAppObjectId = $null
function Remove-TempApp {
    if ($script:TempAppObjectId) {
        Write-ProgressHost -Message "Removing temporary App Registration..."
        try {
            Remove-MgApplication -ApplicationId $script:TempAppObjectId -ErrorAction Stop
            Write-ProgressHost -Message "[OK] Temporary App Registration removed."
        } catch {
            Write-Host "  [WARN] Could not remove temp App Registration (ID: $($script:TempAppObjectId)). Remove it manually in Entra ID > App registrations." -ForegroundColor Yellow
        }
        $script:TempAppObjectId = $null
    }
}

# ── App-only token (temporary app mode) ─────────────────────────────────────
# The delegated session stays connected throughout (needed to remove the temp
# app at the end); mail operations use this raw bearer token instead.
function Update-AppOnlyToken {
    if (-not $script:TokenBody) { return }
    if ($script:AppOnlyHeaders -and (Get-Date) -lt $script:TokenExpiry) { return }
    $resp = Invoke-RestMethod -Method POST -ErrorAction Stop `
        -Uri  "https://login.microsoftonline.com/$($script:TokenTenantId)/oauth2/v2.0/token" `
        -Body $script:TokenBody
    $script:AppOnlyHeaders = @{ Authorization = "Bearer $($resp.access_token)" }
    $script:TokenExpiry    = (Get-Date).AddSeconds($resp.expires_in - 300)
}

function Invoke-Graph {
    param(
        [Parameter(Mandatory)] [string] $Method,
        [Parameter(Mandatory)] [string] $Uri,
        [string] $Body,
        [string] $ContentType = 'application/json'
    )

    if ($script:AppOnlyHeaders) {
        Update-AppOnlyToken
        $params = @{ Method = $Method; Uri = $Uri; Headers = $script:AppOnlyHeaders; ErrorAction = 'Stop' }
        if ($Body) { $params['Body'] = $Body; $params['ContentType'] = $ContentType }
        return Invoke-RestMethod @params
    }

    $params = @{ Method = $Method; Uri = $Uri; ErrorAction = 'Stop' }
    if ($Body) { $params['Body'] = $Body; $params['ContentType'] = $ContentType }
    return Invoke-MgGraphRequest @params
}

# ── Temporary Full Access cleanup (-Delegated mode) ─────────────────────────
# -Delegated grants the signed-in admin Full Access on the target mailbox just
# long enough to archive it, then revokes it again — no permission is left behind.
$script:GrantedFullAccessMailbox = $null
$script:GrantedFullAccessUser    = $null
$script:ConnectedExo             = $false
function Remove-TempFullAccess {
    if ($script:GrantedFullAccessMailbox -and $script:GrantedFullAccessUser) {
        Write-ProgressHost -Message "Removing temporary Full Access on '$($script:GrantedFullAccessMailbox)'..."
        $removed = $false
        for ($i = 1; $i -le 4 -and -not $removed; $i++) {
            $removeWarnings = $null
            try {
                Remove-MailboxPermission -Identity $script:GrantedFullAccessMailbox -User $script:GrantedFullAccessUser `
                    -AccessRights FullAccess -Confirm:$false -ErrorAction Stop -WarningVariable removeWarnings -WarningAction SilentlyContinue
            } catch {
                $removeWarnings = @($_.Exception.Message)
            }

            if (-not $removeWarnings) {
                $removed = $true
            } elseif ($i -lt 4) {
                # The grant may not have replicated to the domain controller this
                # request landed on yet — wait and retry rather than giving up.
                Write-ProgressHost -Message "Permission not visible yet, retrying removal ($i/4)..."
                Start-Sleep -Seconds 15
            }
        }
        if ($removed) {
            Write-ProgressHost -Message "[OK] Temporary Full Access removed."
        } else {
            Write-Host "  [WARN] Could not confirm removal of temporary Full Access for '$($script:GrantedFullAccessUser)' on '$($script:GrantedFullAccessMailbox)'. Verify manually (Get-MailboxPermission / Remove-MailboxPermission)." -ForegroundColor Yellow
        }
        $script:GrantedFullAccessMailbox = $null
        $script:GrantedFullAccessUser    = $null
    }
    if ($script:ConnectedExo) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        $script:ConnectedExo = $false
    }
}

$useTempApp = (-not $ClientId) -and (-not $Delegated)
if ($useTempApp -and -not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Applications')) {
    Write-Host "  [ERROR] Missing required module: Microsoft.Graph.Applications (needed to auto-create the temporary app registration)." -ForegroundColor Red
    Write-Host "  Install with: .\scripts\Startup\Install-Modules.ps1, or pass -Delegated to skip app-only mode." -ForegroundColor Yellow
    exit 1
}

# ── Tenant resolution (GDAP-aware, consistent with Get-SharePointStorageReport.ps1) ──
$isGdapMode = $false
try {
    if ($global:authMode -and ([string]$global:authMode).ToUpperInvariant() -eq 'GDAP') {
        $isGdapMode = $true
    } elseif ($env:M365_AUTH_MODE -and ([string]$env:M365_AUTH_MODE).ToUpperInvariant() -eq 'GDAP') {
        $isGdapMode = $true
    }
} catch {}

$effectiveTenantId = $TenantId
if (-not $effectiveTenantId) {
    try {
        if ($isGdapMode -and $global:cid) {
            $effectiveTenantId = [string]$global:cid
        } elseif ($env:M365_CUSTOMER_TENANTID) {
            $effectiveTenantId = [string]$env:M365_CUSTOMER_TENANTID
        }
    } catch {}
}

if (($ClientId -or $useTempApp) -and -not $effectiveTenantId) {
    throw "-TenantId is required (or a resolvable GDAP customer tenant context) for app-only auth."
}

if ($isGdapMode -and -not $effectiveTenantId) {
    Write-Host "  [WARN] GDAP mode detected but no customer TenantId found. Run Connect-Tenant first or pass -TenantId." -ForegroundColor Yellow
}

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    if ($ClientId) {
        # ── Bring your own app: full app-only connection, no temp app needed ────
        if ($CertificateThumbprint) {
            Connect-MgGraph -ClientId $ClientId -TenantId $effectiveTenantId `
                -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
        } else {
            $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
            $cred = [System.Management.Automation.PSCredential]::new($ClientId, $secureSecret)
            Connect-MgGraph -ClientId $ClientId -TenantId $effectiveTenantId `
                -ClientSecretCredential $cred -NoWelcome -ErrorAction Stop
        }
        $script:ConnectedHere = $true
        Write-Host "  [OK]   Connected with provided app credentials." -ForegroundColor DarkGray

    } elseif ($Delegated) {
        # ── Plain delegated session — auto-grants temporary Full Access via Exchange
        #    Online on other mailboxes (Exchange Admin role only, no Entra app rights) ──
        if (-not (Get-Module -ListAvailable -Name 'ExchangeOnlineManagement')) {
            throw "Missing required module: ExchangeOnlineManagement (needed for -Delegated's temporary Full Access grant). Install with: .\scripts\Startup\Install-Modules.ps1"
        }

        Write-Host "  Connecting to Exchange Online..." -ForegroundColor Cyan
        $eos = Get-ConnectionInformation -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $eos) {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            $script:ConnectedExo = $true
            $eos = Get-ConnectionInformation -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        $adminUpn = $eos.UserPrincipalName

        if ($adminUpn -and $adminUpn.ToLowerInvariant() -ne $Mailbox.ToLowerInvariant()) {
            Write-Host "  Granting temporary Full Access on '$Mailbox' to '$adminUpn'..." -ForegroundColor Cyan
            Add-MailboxPermission -Identity $Mailbox -User $adminUpn -AccessRights FullAccess `
                -AutoMapping:$false -Confirm:$false -ErrorAction Stop | Out-Null
            $script:GrantedFullAccessMailbox = $Mailbox
            $script:GrantedFullAccessUser    = $adminUpn
            Write-Host "  [OK]   Full Access granted (temporary — will be removed when the script finishes)." -ForegroundColor DarkGray

            # Exchange Online permission changes can take anywhere from seconds to a
            # few minutes to propagate — poll instead of a single blind sleep.
            $propagated = $false
            for ($i = 1; $i -le 12; $i++) {
                Write-ProgressHost -Message "Waiting for the permission to propagate (attempt $i/12)..."
                Start-Sleep -Seconds 15
                $perm = Get-MailboxPermission -Identity $Mailbox -User $adminUpn -ErrorAction SilentlyContinue
                if ($perm | Where-Object { $_.AccessRights -contains 'FullAccess' }) {
                    $propagated = $true
                    break
                }
            }
            if ($propagated) {
                Write-ProgressHost -Message "[OK] Permission confirmed. Giving Graph a few more seconds to catch up..."
                Start-Sleep -Seconds 15
            } else {
                Write-Host "  [WARN] Full Access not yet visible after 3 minutes — continuing anyway." -ForegroundColor Yellow
            }
        }

        $connectParams = @{ Scopes = @('Mail.ReadWrite'); NoWelcome = $true }
        if ($effectiveTenantId) { $connectParams['TenantId'] = $effectiveTenantId }
        Connect-MgGraph @connectParams -ErrorAction Stop
        $script:ConnectedHere = $true
        Write-Host "  [OK]   Connected (delegated)." -ForegroundColor DarkGray

    } else {
        # ── Auto mode: create a short-lived app-only app, use it, then remove it ──
        Write-Host "  Connecting interactively to set up a temporary app registration..." -ForegroundColor Cyan
        Write-Host "  Required role: Global Administrator or Privileged Role Administrator (one-time setup)" -ForegroundColor DarkGray
        Connect-MgGraph -Scopes @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All') `
            -TenantId $effectiveTenantId -NoWelcome -ErrorAction Stop
        $script:ConnectedHere = $true
        Write-Host "  [OK]   Connected (delegated, for app setup only)." -ForegroundColor DarkGray

        $ts = Get-Date -Format 'yyyyMMddHHmmss'
        $appName = "InboxArchive-Temp-$ts"
        Write-Host "  Creating temporary App Registration '$appName'..." -ForegroundColor Cyan
        $app = New-MgApplication -DisplayName $appName -ErrorAction Stop
        $script:TempAppObjectId = $app.Id

        $sp      = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop
        $graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ErrorAction Stop
        $appRole = $graphSp.AppRoles | Where-Object { $_.Value -eq 'Mail.ReadWrite' -and $_.AllowedMemberTypes -contains 'Application' }
        if (-not $appRole) {
            Write-Host "  [ERROR] Could not resolve app role 'Mail.ReadWrite'." -ForegroundColor Red
            Remove-TempApp
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            exit 1
        }
        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $sp.Id `
            -PrincipalId        $sp.Id `
            -ResourceId         $graphSp.Id `
            -AppRoleId          $appRole.Id `
            -ErrorAction Stop | Out-Null
        Write-Host "  [OK]   Mail.ReadWrite (application) granted." -ForegroundColor DarkGray

        $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
            displayName = 'temp'
            endDateTime = (Get-Date).AddHours(2)
        } -ErrorAction Stop

        Write-Host "  Obtaining app-only token..." -ForegroundColor Cyan
        $script:TokenBody = @{
            grant_type    = 'client_credentials'
            scope         = 'https://graph.microsoft.com/.default'
            client_id     = $app.AppId
            client_secret = $secret.SecretText
        }
        $script:TokenTenantId = $effectiveTenantId

        $tokenObtained = $false
        for ($i = 1; $i -le 6; $i++) {
            try {
                Update-AppOnlyToken
                $tokenObtained = $true
                break
            } catch {
                if ($i -lt 6) {
                    Write-ProgressHost -Message "Waiting for app registration propagation (attempt $i/6)..."
                    Start-Sleep -Seconds 5
                }
            }
        }
        if (-not $tokenObtained) {
            Write-Host "  [ERROR] Could not obtain an app-only token after propagation retries." -ForegroundColor Red
            Remove-TempApp
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            exit 1
        }

        # Delegated session (Application.ReadWrite.All) stays connected — needed to
        # remove the temp app at the end. Mail calls go through the app-only token.
        Write-Host "  [OK]   App-only token obtained (temporary app)." -ForegroundColor DarkGray
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Remove-TempApp
    Remove-TempFullAccess
    if ($script:ConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    exit 1
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Archive Inbox" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Mailbox : $Mailbox"
if ($After)  { Write-Host "  After   : $($After.ToString('yyyy-MM-dd HH:mm'))" }
if ($Before) { Write-Host "  Before  : $($Before.ToString('yyyy-MM-dd HH:mm'))" }
Write-Host ("  Auth    : {0}" -f $(if ($ClientId) { 'App-only (provided app)' } elseif ($Delegated) { 'Delegated (temporary Full Access)' } else { 'App-only (temporary app)' })) -ForegroundColor DarkGray
Write-Host ("  Mode    : {0}" -f $(if ($Apply) { 'Apply (messages will be moved)' } else { 'Preview only (no changes)' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

# ── Build $filter ────────────────────────────────────────────────────────────
$filterParts = [System.Collections.Generic.List[string]]::new()
if ($After)  { $filterParts.Add("receivedDateTime ge $($After.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))") }
if ($Before) { $filterParts.Add("receivedDateTime lt $($Before.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))") }
$filter = $filterParts -join ' and '

# ── Get Inbox message IDs (paginated) ───────────────────────────────────────────
function Get-InboxMessageIds {
    param([string] $MailboxId, [string] $Filter)

    $ids = [System.Collections.Generic.List[string]]::new()
    $url = "https://graph.microsoft.com/v1.0/users/$MailboxId/mailFolders/inbox/messages?`$select=id&`$top=999"
    if ($Filter) { $url += "&`$filter=$([uri]::EscapeDataString($Filter))" }

    while ($url) {
        $resp = Invoke-Graph -Method GET -Uri $url
        foreach ($m in $resp.value) { $ids.Add($m.id) }
        $url = $resp.'@odata.nextLink'
        if ($url) { Write-ProgressHost -Message "$($ids.Count) message ID(s) retrieved so far..." }
    }

    return $ids
}

# ── Move messages in batches of 20, with retry on throttling ───────────────────
function Move-MessagesToArchive {
    param([string] $MailboxId, [System.Collections.Generic.List[string]] $MessageIds)

    $moved   = 0
    $failed  = 0
    $pending = $MessageIds
    $total   = $MessageIds.Count
    $maxPasses = 5

    for ($pass = 1; $pass -le $maxPasses -and $pending.Count -gt 0; $pass++) {
        $retry = [System.Collections.Generic.List[string]]::new()

        for ($i = 0; $i -lt $pending.Count; $i += 20) {
            $end   = [Math]::Min($i + 19, $pending.Count - 1)
            $chunk = $pending.GetRange($i, $end - $i + 1)

            $percent = [Math]::Min(100, [int](($moved + $failed) / [Math]::Max($total, 1) * 100))
            Write-Progress -Activity "Archiving Inbox" -Status "$moved / $total moved" -PercentComplete $percent
            Write-ProgressHost -Message "$moved / $total moved ($percent%)..."

            $batchBody = @{
                requests = @($chunk | ForEach-Object {
                    @{
                        id      = $_
                        method  = 'POST'
                        url     = "/users/$MailboxId/messages/$_/move"
                        body    = @{ destinationId = 'archive' }
                        headers = @{ 'Content-Type' = 'application/json' }
                    }
                })
            } | ConvertTo-Json -Depth 6

            try {
                $resp = Invoke-Graph -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' -Body $batchBody -ContentType 'application/json'
            } catch {
                $retry.AddRange($chunk)
                continue
            }

            foreach ($r in $resp.responses) {
                if ($r.status -in 200, 201) {
                    $moved++
                } elseif ($r.status -in 429, 503, 504) {
                    $retry.Add($r.id)
                } else {
                    $failed++
                    Write-Host "  [WARN] Move failed for message $($r.id): HTTP $($r.status)" -ForegroundColor Yellow
                }
            }
        }

        $pending = $retry
        if ($pending.Count -gt 0 -and $pass -lt $maxPasses) {
            Start-Sleep -Seconds ([Math]::Min(30, [Math]::Pow(2, $pass)))
        }
    }

    Write-Progress -Activity "Archiving Inbox" -Completed

    if ($pending.Count -gt 0) {
        $failed += $pending.Count
        Write-Host "  [WARN] $($pending.Count) message(s) could not be moved after retries." -ForegroundColor Yellow
    }

    return [PSCustomObject]@{ Moved = $moved; Failed = $failed }
}

# ── Run ──────────────────────────────────────────────────────────────────────
Write-Host "  Retrieving Inbox message IDs..." -ForegroundColor DarkGray

# In -Delegated mode a 403 here is usually Microsoft Graph's delegate-permission
# cache still catching up (Exchange itself already shows the grant) — retry against
# a wall-clock deadline (not a fixed attempt count) so -MaxWaitMinutes covers
# Microsoft's documented worst case in a single run, without the temporary Full
# Access grant getting revoked and re-granted (which would reset the propagation
# clock) between attempts. Any other mode/error fails immediately since retrying
# wouldn't help.
$graphRetryDelaySeconds = 60
$graphDeadline = if ($Delegated) { (Get-Date).AddMinutes($MaxWaitMinutes) } else { Get-Date }
$ids = $null
$attempt = 0

while ($true) {
    $attempt++
    try {
        $ids = Get-InboxMessageIds -MailboxId $Mailbox -Filter $filter
        break
    } catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        $isPermissionLag = ($statusCode -eq 403 -and $Delegated)

        if ($isPermissionLag -and (Get-Date) -lt $graphDeadline) {
            $minutesLeft = [Math]::Max(0, [Math]::Round(($graphDeadline - (Get-Date)).TotalMinutes, 1))
            Write-ProgressHost -Message "Graph still returns 403 (permission cache lag) — retry $attempt, waiting ${graphRetryDelaySeconds}s (~$minutesLeft min left of the $MaxWaitMinutes-min window)..." -ForegroundColor Yellow
            Start-Sleep -Seconds $graphRetryDelaySeconds
            continue
        }

        Write-Host ""
        Write-Host "  [ERROR] Could not read the Inbox of '$Mailbox': $($_.Exception.Message)" -ForegroundColor Red
        if ($isPermissionLag) {
            Write-Host "  [HINT] Exchange Online already shows the Full Access grant (the removal below" -ForegroundColor Yellow
            Write-Host "         wouldn't otherwise succeed), but Microsoft Graph's own authorization cache" -ForegroundColor Yellow
            Write-Host "         for delegate mailbox access can lag up to ~60 minutes behind — a known" -ForegroundColor Yellow
            Write-Host "         Microsoft limitation. Already retried for $MaxWaitMinutes minutes." -ForegroundColor Yellow
            Write-Host "         Re-run with a higher -MaxWaitMinutes, or drop -Delegated to use the" -ForegroundColor Yellow
            Write-Host "         default app-only mode instead, which has no such propagation delay." -ForegroundColor Yellow
        } elseif ($statusCode -eq 403 -and -not $Delegated) {
            Write-Host "  [HINT] App-only access denied on this specific mailbox despite Mail.ReadWrite" -ForegroundColor Yellow
            Write-Host "         application permission. The most common cause is an Application Access" -ForegroundColor Yellow
            Write-Host "         Policy in this tenant that restricts which mailboxes app-only calls may" -ForegroundColor Yellow
            Write-Host "         touch. Check with (as an admin):" -ForegroundColor Yellow
            Write-Host "           Get-ApplicationAccessPolicy" -ForegroundColor Yellow
            Write-Host "         and whether '$Mailbox' is a member of the scoping group used there." -ForegroundColor Yellow
            Write-Host "         Also possible: '$Mailbox' has no Exchange Online license, or is a" -ForegroundColor Yellow
            Write-Host "         mailbox type Graph doesn't allow app-only access to." -ForegroundColor Yellow
        }
        Remove-TempApp
        Remove-TempFullAccess
        if ($script:ConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
        exit 1
    }
}
Write-Host "  Found $($ids.Count) message(s) to archive." -ForegroundColor DarkGray
Write-Host ""

if ($ids.Count -eq 0) {
    Write-Host "  Nothing to archive." -ForegroundColor DarkGray
} elseif (-not $Apply) {
    Write-Host "  $($ids.Count) message(s) would be archived. Re-run with -Apply to move them." -ForegroundColor Yellow
} else {
    $result = Move-MessagesToArchive -MailboxId $Mailbox -MessageIds $ids

    Write-Host ""
    Write-Host "  Archived : $($result.Moved)" -ForegroundColor Green
    if ($result.Failed -gt 0) {
        Write-Host "  Failed   : $($result.Failed)" -ForegroundColor Yellow
    }
}

Write-Host ""

# ── Disconnect / cleanup ─────────────────────────────────────────────────────
Remove-TempApp
Remove-TempFullAccess
if ($script:ConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
