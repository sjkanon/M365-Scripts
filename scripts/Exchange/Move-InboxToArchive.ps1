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
      connection instead (only needs the delegated scope, no admin app-creation rights)
      — but then archiving a mailbox other than the signed-in user's own requires the
      signed-in account to already have Full Access on that mailbox.

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
    Mail.ReadWrite session instead. Archiving a mailbox other than the signed-in
    user's own then requires the signed-in account to have Full Access on it.

.PARAMETER Apply
    Actually move the messages. Without this switch, the script only reports how
    many messages would be archived.

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
    # Own mailbox / already have Full Access — skip app-only setup
    .\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Delegated -Apply

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
    [switch] $Apply
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
        # ── Plain delegated session — requires Full Access on other mailboxes ───
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
            Remove-TempApp; exit 1
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
            Remove-TempApp; exit 1
        }

        # Delegated session (Application.ReadWrite.All) stays connected — needed to
        # remove the temp app at the end. Mail calls go through the app-only token.
        $script:ConnectedHere = $true
        Write-Host "  [OK]   App-only token obtained (temporary app)." -ForegroundColor DarkGray
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Remove-TempApp
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
Write-Host ("  Auth    : {0}" -f $(if ($ClientId) { 'App-only (provided app)' } elseif ($Delegated) { 'Delegated' } else { 'App-only (temporary app)' })) -ForegroundColor DarkGray
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
try {
    $ids = Get-InboxMessageIds -MailboxId $Mailbox -Filter $filter
} catch {
    $statusCode = $_.Exception.Response.StatusCode.value__
    Write-Host ""
    Write-Host "  [ERROR] Could not read the Inbox of '$Mailbox': $($_.Exception.Message)" -ForegroundColor Red
    if ($statusCode -eq 403 -and $Delegated) {
        Write-Host "  [HINT] Access denied on a delegated connection usually means the signed-in account" -ForegroundColor Yellow
        Write-Host "         does not have Full Access on this mailbox. Either grant Full Access, or drop" -ForegroundColor Yellow
        Write-Host "         -Delegated to let the script set up app-only access automatically." -ForegroundColor Yellow
    }
    Remove-TempApp
    if ($script:ConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    exit 1
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
if ($script:ConnectedHere) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
