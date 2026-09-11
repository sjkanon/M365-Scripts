#Requires -Version 5.1
<#
.SYNOPSIS
    Show where every calendar is mapped: which mailboxes have which other people's
    calendars in their calendar list, side by side with the rights behind them.
    -Search finds one calendar by keyword ("balie") and shows where it lives and
    where it is mapped.

.DESCRIPTION
    Test-CalendarPermissions.ps1 answers "who MAY open this calendar". This script
    answers "where IS this calendar mapped" - the calendars that actually sit in a
    user's calendar list in Outlook - and lines the two up.

    For every mailbox it reads:
      - the calendar list (GET /users/{id}/calendars). Every calendar in it whose
        owner is somebody else is a mapped calendar: a colleague, a shared mailbox,
        a room, a Microsoft 365 group, or someone outside the organisation.
      - the permissions on its own main calendar
        (GET /users/{id}/calendar/calendarPermissions).

    Both are folded into one row per calendar owner + user pair, with a status:

      Source              -Search only: the matching calendar lives in this mailbox.
      Mapped              In the user's list, and the user has an explicit right.
      MappedWithoutRight  In the list, but no explicit right on the owner's
                          calendar. Access then comes from the organisation-wide
                          default, a group, another calendar of the owner - or the
                          right was removed and the entry is left over.
      MappedGroupCalendar A Microsoft 365 group calendar (access = membership).
      MappedOwnerMissing  The owner no longer exists in the tenant - a stale entry.
      MappedExternal      The owner is outside the tenant.
      NotMapped           Explicit right, but the calendar is not in the user's list.
      NotChecked          Explicit right, but the user's list could not be read.
      GrantedToGroup      A right granted to a group (members are not expanded).
      GrantedToMissing    A right for an address that no longer exists.
      SharedExternally    A right for an address outside the tenant.
      OrgWideDefault      Everyone in the organisation gets more than free/busy.

    Why Graph and not Exchange Online PowerShell: the Exchange cmdlets see folders
    and the permissions on them, not the entries a user added to their own
    calendar list. Those are only readable over Graph.

    Search
    ------
    -Search balie finds every calendar where the keyword appears in the owner's
    name or address (the shared mailbox balie@, a room, a group called Balie) or
    in the calendar's own name (a secondary calendar "Balie" in somebody's
    mailbox). The keyword is matched anywhere in the text, case-insensitive;
    wildcards (* and ?) are used as given.

    For each match the report shows where the calendar lives (Source), every
    mailbox that has it in its calendar list, and everyone with an explicit
    right on it - including a secondary calendar's own rights, which a normal run
    does not read. Every calendar list is still read, because a mapping can sit in
    any mailbox; only the permissions of the matching calendars are read.

    A shared secondary calendar is recognised in someone's list by its name
    matching the keyword. If a user has it under another name, it shows up as
    NotMapped with a note naming the entry that is probably it.

    Not visible here
    ----------------
      - Full Access with AutoMapping adds a whole mailbox to Outlook, calendar
        included. That is a mailbox permission, not a calendar entry - see
        Test-MailboxPermissions.ps1.
      - A calendar opened in classic Outlook with "shared calendar improvements"
        turned off may live only in that Outlook profile's navigation pane and not
        in the list Graph returns.
      - Without -Search, rights are compared against the owner's MAIN calendar. A
        secondary calendar the owner shared shows up as MappedWithoutRight.

    Scope
    -----
    Without -Mailbox every mailbox in the tenant is scanned, which is the only way
    to find mappings that rest on the organisation-wide default or on a group.
    With -Mailbox the report is limited to rows where one of those mailboxes is the
    owner or the user: their own calendar lists are read, plus the lists of
    everyone with an explicit right on their calendar. -Mailbox and -Search
    cannot be combined.

    Graph access
    ------------
    Needs application permissions Calendars.Read and User.Read.All, plus
    Group.Read.All to recognise group calendars (without it they are reported as
    MappedOwnerMissing, with a note saying so). Obtained the same three ways as in
    Remove-PhishingMessage.ps1:

      1. An app-only Graph session you already established is used as-is.
      2. -ClientId with -ClientSecret (plain REST) or -CertificateThumbprint
         (Graph SDK) uses your own App Registration.
      3. Otherwise: device code sign-in, a short-lived App Registration that
         self-grants the three read permissions, an app-only token, and the app is
         removed again when the run ends. Needs Global Administrator or Privileged
         Role Administrator for that sign-in. No extra modules.

    Everything is read-only. No Exchange Online connection is made, so the
    Exchange/Graph MSAL clash described in the Exchange readme does not apply.

    GDAP-aware: under a GDAP session ($global:authMode -eq 'GDAP', set by
    Connect-Tenant / load.ps1) -TenantId is resolved from the selected customer
    tenant ($global:cid) when not supplied. $env:M365_CUSTOMER_TENANTID /
    $env:M365_AUTH_MODE are honoured too.

.PARAMETER Search
    Keyword to find one calendar by: matched against the owner's name and
    addresses and against the calendar's own name. Alias: -Keyword.

.PARAMETER Mailbox
    One or more mailbox addresses (UPN or SMTP). Limits the report to rows where
    one of these mailboxes is the calendar owner or the user. Omit to scan every
    mailbox in the tenant.

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\CalendarMappings_<timestamp>.csv
    (~/Downloads on macOS/Linux).

.PARAMETER TenantId
    Tenant ID or verified domain. Optional for the temporary-app route: without it
    the sign-in decides the tenant, and the script prints which one it is.

.PARAMETER ClientId
    Your own App Registration for app-only Graph access. Use with -TenantId and
    -ClientSecret or -CertificateThumbprint.

.PARAMETER ClientSecret
    Client secret for -ClientId.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for -ClientId. Goes through Connect-MgGraph.

.EXAMPLE
    # Where is the Balie calendar, and who has it mapped?
    .\Get-CalendarMappings.ps1 -Search balie

.EXAMPLE
    # Where is every calendar in the tenant mapped?
    .\Get-CalendarMappings.ps1 -TenantId contoso.com

.EXAMPLE
    # Where is Jan's calendar mapped, and which calendars has Jan mapped?
    .\Get-CalendarMappings.ps1 -Mailbox jan@contoso.com

.EXAMPLE
    # Own App Registration (Calendars.Read, User.Read.All, Group.Read.All)
    .\Get-CalendarMappings.ps1 -TenantId contoso.com -ClientId <appId> -ClientSecret <secret>

.NOTES
    Author: Sjoerd Kanon
#>
[CmdletBinding()]
param(
    [Alias('Keyword')]
    [string]   $Search,
    [string[]] $Mailbox,
    [string]   $OutputPath,
    [string]   $TenantId,
    [string]   $ClientId,
    [string]   $ClientSecret,
    [string]   $CertificateThumbprint
)

if ($Search -and $Mailbox) {
    throw "Use either -Search or -Mailbox, not both."
}
$Search = ([string]$Search).Trim()

# -- Output folder ------------------------------------------------------------
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# -- Header -------------------------------------------------------------------
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Calendar Mappings" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
if ($Search) {
    Write-Host "  Search    : '$Search' in owner name/address or calendar name" -ForegroundColor DarkGray
} elseif ($Mailbox) {
    Write-Host "  Scope     : $($Mailbox -join ', ') (as owner or as user)" -ForegroundColor DarkGray
} else {
    Write-Host "  Scope     : every mailbox in the tenant" -ForegroundColor DarkGray
}
Write-Host ""

# Calendars.Read covers both the calendar list and calendarPermissions.
# Group.Read.All is only there to tell a group calendar from a removed mailbox.
$RequiredRoles = @('Calendars.Read', 'User.Read.All')
$OptionalRoles = @('Group.Read.All')

# A supplied app is often broader than it needs to be; any of these will do.
$RoleAlternatives = @{
    'Calendars.Read' = @('Calendars.Read', 'Calendars.ReadWrite')
    'User.Read.All'  = @('User.Read.All', 'User.ReadWrite.All', 'Directory.Read.All', 'Directory.ReadWrite.All')
    'Group.Read.All' = @('Group.Read.All', 'Group.ReadWrite.All', 'Directory.Read.All', 'Directory.ReadWrite.All')
}

# ==============================================================================
#  Graph connection
#
#  Same three-way pattern as Remove-PhishingMessage.ps1: reuse an app-only
#  session, use your own app, or build a short-lived one and remove it after.
# ==============================================================================
$script:TempAppObjectId = $null
$script:AppOnlyHeaders  = $null
$script:TokenBody       = $null
$script:TokenTenantId   = $null
$script:TokenExpiry     = [datetime]::MinValue
$script:AccessToken     = $null
$script:GraphConnected  = $false
$script:AdminHeaders    = $null
$script:SkipGroups      = $false

# The Graph PowerShell SDK's own public client, so the device code sign-in looks
# exactly like Connect-MgGraph would - without loading the SDK.
$script:GraphCliClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'

function Invoke-GraphAdmin {
    # App-management calls, with the delegated token from the device code flow.
    param(
        [string] $Method = 'GET',
        [Parameter(Mandatory)] [string] $Uri,
        $Body
    )
    $params = @{ Method = $Method; Uri = $Uri; Headers = $script:AdminHeaders; ErrorAction = 'Stop' }
    if ($Body) {
        $params['Body']        = ($Body | ConvertTo-Json -Depth 6)
        $params['ContentType'] = 'application/json'
    }
    return Invoke-RestMethod @params
}

function Remove-TempApp {
    <#
        Needs the delegated session's Application.ReadWrite.All - the app-only
        token only holds read permissions - so this runs before that token goes.
    #>
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
    <#
        One entry point for every Graph call, so a raw bearer token and a
        Connect-MgGraph session look identical to callers. Throttling and gateway
        hiccups are retried; a 401/403 never improves by asking again.
    #>
    param(
        [string] $Method = 'GET',
        [Parameter(Mandatory)] [string] $Uri,
        $Body,
        [string] $ContentType = 'application/json'
    )

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
            # Captured up front: the status probe below has its own catch.
            $err  = $_
            $msg  = "$($err.Exception.Message)"
            $code = 0
            try { $code = [int]$err.Exception.Response.StatusCode } catch { $code = 0 }
            if ($code -eq 0 -and $msg -match '(?<![0-9])(429|503|504)(?![0-9])') { $code = [int]$Matches[1] }

            if ($attempt -ge 5 -or $code -notin @(429, 503, 504)) { throw $err }
            Start-Sleep -Seconds ([Math]::Pow(2, $attempt))
        }
    }
}

function Get-TokenClaim {
    # Decodes the payload of a JWT. Returns $null when it cannot.
    param([string] $Jwt)
    try {
        $payload = $Jwt.Split('.')[1].Replace('-', '+').Replace('_', '/')
        switch ($payload.Length % 4) {
            2 { $payload += '==' }
            3 { $payload += '=' }
        }
        return ([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($payload)) | ConvertFrom-Json)
    } catch {
        return $null
    }
}

function Get-MissingRole {
    <#
        Which of the wanted roles the token lacks. A client_credentials token is
        issued whether or not anything was granted - the roles are just absent -
        so without this the run would 403 on every mailbox instead of saying why.
    #>
    param([string[]] $Have, [string[]] $Wanted)
    return @($Wanted | Where-Object {
        $alternatives = $RoleAlternatives[$_]
        -not ($Have | Where-Object { $alternatives -contains $_ })
    })
}

function Confirm-AppRole {
    <#
        Waits until the token carries the roles. Role assignments take a while to
        reach the token service and a token minted too early is cached for an
        hour, so this re-mints rather than just sleeping.
    #>
    param([string[]] $Required, [int] $TimeoutSeconds = 180)

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $reported = $false
    while ($true) {
        $claims  = Get-TokenClaim -Jwt $script:AccessToken
        $have    = if ($claims) { @($claims.roles) } else { @() }
        $missing = Get-MissingRole -Have $have -Wanted $Required
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
    <# Device code flow over plain REST. Returns the raw access token. #>
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
            return $tok.access_token
        } catch {
            $code = ''
            try { $code = ($_.ErrorDetails.Message | ConvertFrom-Json).error } catch {}
            if ($code -eq 'authorization_pending')  { continue }
            if ($code -eq 'slow_down')              { $interval += 5; continue }
            if ($code -eq 'expired_token')          { throw "Device code expired before sign-in completed." }
            if ($code -eq 'authorization_declined') { throw "Sign-in was declined." }
            throw
        }
    }
    throw "Device code sign-in timed out."
}

function Connect-GraphForCalendars {
    <# Establishes app-only read access. Returns $true when Graph is usable. #>

    # 1. An app-only session the caller already established.
    $ctx = $null
    try { $ctx = Get-MgContext -ErrorAction SilentlyContinue } catch {}
    if ($ctx -and $ctx.AuthType -eq 'AppOnly') {
        $have = @($ctx.Scopes)
        $missing = Get-MissingRole -Have $have -Wanted $RequiredRoles
        if ($missing.Count -gt 0) {
            Write-Warning "The existing app-only Graph session lacks $($missing -join ', ')."
            return $false
        }
        if ((Get-MissingRole -Have $have -Wanted $OptionalRoles).Count -gt 0) { $script:SkipGroups = $true }
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
        if ($ClientSecret) {
            try {
                Get-AppOnlyTokenByRest -Tenant $effectiveTenantId -App $ClientId -Secret $ClientSecret
            } catch {
                Write-Warning "Could not obtain a token for the supplied app: $($_.Exception.Message)"
                return $false
            }
            $claims  = Get-TokenClaim -Jwt $script:AccessToken
            $have    = if ($claims) { @($claims.roles) } else { @() }
            $missing = Get-MissingRole -Have $have -Wanted $RequiredRoles
            if ($missing.Count -gt 0) {
                Write-Warning "App $ClientId has no $($missing -join ', ') application permission (token carries: $(if ($have) { $have -join ', ' } else { 'nothing' }))."
                return $false
            }
            if ((Get-MissingRole -Have $have -Wanted $OptionalRoles).Count -gt 0) { $script:SkipGroups = $true }
            Write-Host "  [OK]   App-only token obtained (supplied app, REST)." -ForegroundColor DarkGray
            return $true
        }
        if ($CertificateThumbprint) {
            try {
                Connect-MgGraph -ClientId $ClientId -TenantId $effectiveTenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
            } catch {
                Write-Warning "Could not connect with the supplied certificate: $($_.Exception.Message)"
                return $false
            }
            if ((Get-MissingRole -Have @((Get-MgContext).Scopes) -Wanted $OptionalRoles).Count -gt 0) { $script:SkipGroups = $true }
            Write-Host "  [OK]   Connected with the supplied certificate." -ForegroundColor DarkGray
            $script:GraphConnected = $true
            return $true
        }
        Write-Warning "-ClientId needs -ClientSecret or -CertificateThumbprint."
        return $false
    }

    # 3. A short-lived app, removed again at the end. Plain REST, no SDK.
    $wanted = $RequiredRoles + $OptionalRoles
    try {
        Write-Host "  No app-only Graph access yet - setting up a temporary App Registration." -ForegroundColor Cyan
        Write-Host "  Required role: Global Administrator or Privileged Role Administrator (one-time)" -ForegroundColor DarkGray

        # Without a tenant the sign-in decides - the tenant is read back from the
        # token and printed, so it is never a guess which tenant was scanned.
        $signInTenant = if ($effectiveTenantId) { $effectiveTenantId } else { 'organizations' }
        $adminToken   = Get-DelegatedTokenByDeviceCode -Tenant $signInTenant `
                            -Scopes @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
        $script:AdminHeaders = @{ Authorization = "Bearer $adminToken" }
        if (-not $effectiveTenantId) {
            $effectiveTenantId = (Get-TokenClaim -Jwt $adminToken).tid
            if (-not $effectiveTenantId) { throw "Could not read the tenant from the sign-in token. Pass -TenantId." }
        }
        Write-Host "  [OK]   Signed in (tenant $effectiveTenantId)." -ForegroundColor DarkGray

        $appName = "CalendarMap-Temp-$(Get-Date -Format 'yyyyMMddHHmmss')"
        Write-Host "  Creating temporary App Registration '$appName'..." -ForegroundColor Cyan
        $app = Invoke-GraphAdmin -Method POST -Uri 'https://graph.microsoft.com/v1.0/applications' `
                    -Body @{ displayName = $appName; signInAudience = 'AzureADMyOrg' }
        $script:TempAppObjectId = $app.id

        $sp = Invoke-GraphAdmin -Method POST -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Body @{ appId = $app.appId }
        $graphSp = @((Invoke-GraphAdmin -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'").value)[0]
        if (-not $graphSp) { throw "Could not resolve the Microsoft Graph service principal." }

        foreach ($roleName in $wanted) {
            $appRole = @($graphSp.appRoles | Where-Object { $_.value -eq $roleName -and $_.allowedMemberTypes -contains 'Application' })[0]
            if (-not $appRole) { throw "Could not resolve the $roleName application role." }
            Invoke-GraphAdmin -Method POST -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)/appRoleAssignments" `
                -Body @{ principalId = $sp.id; resourceId = $graphSp.id; appRoleId = $appRole.id } | Out-Null
            Write-Host "  [OK]   $roleName (application) granted." -ForegroundColor DarkGray
        }

        $secret = Invoke-GraphAdmin -Method POST -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)/addPassword" `
                    -Body @{ passwordCredential = @{ displayName = 'temp'; endDateTime = (Get-Date).AddHours(2).ToString('o') } }

        # A brand new registration is not instantly usable.
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

        if (-not (Confirm-AppRole -Required $wanted)) {
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

# ==============================================================================
#  Reading
# ==============================================================================

function Get-GraphPaged {
    param([Parameter(Mandatory)] [string] $Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    $next  = $Uri
    while ($next) {
        $resp = Invoke-Graph -Uri $next
        if ($resp.value) { foreach ($v in $resp.value) { $items.Add($v) } }
        $next = $resp.'@odata.nextLink'
    }
    return $items.ToArray()
}

function Invoke-GraphBatch {
    <#
        GET requests through $batch, twenty per call - one call per mailbox would
        make a tenant-wide run take many times longer. Throttled items (429/503/
        504 inside a batch) are collected and retried after the longest
        Retry-After the service asked for.

        Returns a hashtable: request key -> [PSCustomObject]@{ Status; Body }.
    #>
    param(
        [Parameter(Mandatory)] [System.Collections.IDictionary] $Requests,
        [string] $Activity = 'Graph'
    )

    $results = @{}
    $pending = @($Requests.Keys)
    $total   = $pending.Count
    $done    = 0

    for ($round = 1; $pending.Count -gt 0; $round++) {
        $retry = [System.Collections.Generic.List[string]]::new()
        $delay = 0

        for ($i = 0; $i -lt $pending.Count; $i += 20) {
            $chunk = @($pending[$i..([Math]::Min($i + 19, $pending.Count - 1))])
            $batch = for ($j = 0; $j -lt $chunk.Count; $j++) {
                @{ id = [string]$j; method = 'GET'; url = $Requests[$chunk[$j]] }
            }
            $body = @{ requests = @($batch) } | ConvertTo-Json -Depth 4
            $resp = Invoke-Graph -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' -Body $body

            foreach ($r in @($resp.responses)) {
                $key  = $chunk[[int]$r.id]
                $code = [int]$r.status
                if ($code -in @(429, 503, 504) -and $round -lt 6) {
                    $retry.Add($key)
                    $after = 0
                    if ($r.headers) { [void][int]::TryParse([string]$r.headers.'Retry-After', [ref]$after) }
                    if ($after -gt $delay) { $delay = $after }
                    continue
                }
                $results[$key] = [PSCustomObject]@{ Status = $code; Body = $r.body }
                $done++
            }
            Write-Progress -Activity $Activity -Status "$done / $total" -PercentComplete ([int](100 * $done / [Math]::Max(1, $total)))
        }

        $pending = @($retry)
        if ($pending.Count -gt 0) {
            if ($delay -lt 2) { $delay = [int][Math]::Pow(2, $round) }
            Write-Host "  Throttled on $($pending.Count) request(s) - retrying in $delay s..." -ForegroundColor DarkGray
            Start-Sleep -Seconds $delay
        }
    }
    Write-Progress -Activity $Activity -Completed
    return $results
}

function New-UserRequest {
    # The same relative path for a set of users, keyed by user id.
    param([string[]] $UserIds, [string] $Path)
    $requests = [ordered]@{}
    foreach ($id in $UserIds) { $requests[$id] = "/users/$id/$Path" }
    return $requests
}

function Get-GraphCollection {
    <#
        Reads one collection per request (a calendar list, or a calendar's
        permissions). A 404 means the user has no Exchange Online mailbox;
        anything else that is not 200 is a mailbox that could not be checked, and
        is kept apart so it is never reported as "nothing found".
    #>
    param([System.Collections.IDictionary] $Requests, [string] $Activity)

    $out = [PSCustomObject]@{
        Ok        = @{}
        NoMailbox = [System.Collections.Generic.List[string]]::new()
        Failed    = @{}
    }
    if (-not $Requests -or $Requests.Count -eq 0) { return $out }

    $raw = Invoke-GraphBatch -Requests $Requests -Activity $Activity

    foreach ($key in @($Requests.Keys)) {
        $r = $raw[$key]
        if ($r.Status -eq 200) {
            $items = [System.Collections.Generic.List[object]]::new()
            if ($r.Body.value) { foreach ($v in $r.Body.value) { $items.Add($v) } }
            $next = $r.Body.'@odata.nextLink'
            if ($next) {
                try {
                    foreach ($v in (Get-GraphPaged -Uri $next)) { $items.Add($v) }
                } catch {
                    $out.Failed[$key] = "next page failed: $($_.Exception.Message)"
                    continue
                }
            }
            $out.Ok[$key] = $items.ToArray()
        } elseif ($r.Status -eq 404) {
            $out.NoMailbox.Add($key)
        } else {
            $out.Failed[$key] = if ($r.Body.error) { "$($r.Status) $($r.Body.error.code): $($r.Body.error.message)" } else { "HTTP $($r.Status)" }
        }
    }
    return $out
}

# -- Directory index ----------------------------------------------------------
# Calendar owners and grantees come back as bare SMTP addresses. Every address of
# every user and mail-enabled group is indexed, so an alias resolves to the same
# object as the primary address.
$script:AddressIndex  = @{}
$script:TenantDomains = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$script:UserById      = @{}

function Get-SmtpAddress {
    param($Primary, $ProxyAddresses, $Upn)
    $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($a in @($Primary, $Upn)) { if ($a) { [void]$set.Add([string]$a) } }
    foreach ($p in @($ProxyAddresses)) {
        if ($p -and [string]$p -match '^(?i)smtp:(.+)$') { [void]$set.Add($Matches[1]) }
    }
    return , $set
}

function Add-Principal {
    param([string] $Kind, [string] $Id, [string] $Name, [string] $Address, $Addresses)
    $principal = [PSCustomObject]@{
        Key       = $Id
        Kind      = $Kind
        Id        = $Id
        Name      = $Name
        Address   = $Address
        Addresses = $Addresses
    }
    foreach ($a in $Addresses) {
        $script:AddressIndex[$a.ToLowerInvariant()] = $principal
        [void]$script:TenantDomains.Add(($a -split '@')[-1])
    }
    return $principal
}

function Resolve-Principal {
    <#
        Maps an address to a known mailbox or group. Anything else is either
        Missing (a tenant domain, so most likely a removed mailbox) or External.
    #>
    param([string] $Address, [string] $Name)

    $addr = ([string]$Address).Trim().ToLowerInvariant()
    if ($script:AddressIndex.ContainsKey($addr)) { return $script:AddressIndex[$addr] }

    $domain = ($addr -split '@')[-1]
    $principal = [PSCustomObject]@{
        Key       = "addr:$addr"
        Kind      = $(if ($script:TenantDomains.Contains($domain)) { 'Missing' } else { 'External' })
        Id        = $null
        Name      = $(if ($Name) { $Name } else { $Address })
        Address   = $Address
        Addresses = $null
    }
    $script:AddressIndex[$addr] = $principal
    return $principal
}

function Test-OwnCalendar {
    # Main, secondary, birthdays, holidays: calendars the user owns are not mappings.
    param($User, $Calendar)
    $ownerAddress = $Calendar.owner.address
    return (-not $ownerAddress -or $User.Addresses.Contains([string]$ownerAddress))
}

# -- Search -------------------------------------------------------------------
# A bare keyword matches anywhere in the text; wildcards are taken as given.
$searchPattern = if ($Search -match '[*?]') { $Search } else { "*$Search*" }

function Test-SearchMatch {
    param([string[]] $Value)
    foreach ($v in $Value) { if ($v -and $v -like $searchPattern) { return $true } }
    return $false
}

function Test-OwnerMatch {
    # Owner name, primary address and every alias.
    param($Principal)
    $values = @($Principal.Name, $Principal.Address)
    if ($Principal.Addresses) { $values += @($Principal.Addresses) }
    return (Test-SearchMatch -Value $values)
}

# -- Report rows --------------------------------------------------------------
# One row per owner + user + calendar of that owner. 'main' is the owner's main
# calendar; a secondary calendar is keyed by its id, and is only read by -Search.
$script:Pairs = [ordered]@{}

function Get-PairKey {
    param($Owner, $User, [string] $CalendarKey = 'main')
    return "$($Owner.Key)|$($User.Key)|$CalendarKey"
}

function Get-PairRow {
    # Shared by the mapping pass and the rights pass.
    param($Owner, $User, [string] $CalendarKey = 'main', [string] $CalendarName)
    $key = Get-PairKey -Owner $Owner -User $User -CalendarKey $CalendarKey
    if (-not $script:Pairs.Contains($key)) {
        $isMailbox = $Owner.Kind -eq 'Mailbox'
        $script:Pairs[$key] = [PSCustomObject]@{
            Owner          = $Owner.Name
            OwnerAddress   = $Owner.Address
            OwnerType      = $Owner.Kind
            Calendar       = $(if ($CalendarKey -ne 'main') { $CalendarName } elseif ($isMailbox) { 'Main' } else { '' })
            User           = $User.Name
            UserAddress    = $User.Address
            UserType       = $User.Kind
            Status         = $null
            Mapped         = 'No'
            MappedAs       = ''
            Rights         = ''
            CanEdit        = ''
            CanViewPrivate = ''
            Note           = ''
            OwnerKey       = $Owner.Key
            UserKey        = $User.Key
            # Which permission read backs this row: the owner id for the main
            # calendar, owner id + calendar id for a secondary one.
            GrantKey       = $(if (-not $isMailbox) { $null } elseif ($CalendarKey -eq 'main') { $Owner.Id } else { "$($Owner.Id)|$CalendarKey" })
        }
    }
    return $script:Pairs[$key]
}

# ==============================================================================
#  Run
# ==============================================================================
$results = @()
try {
    if (-not (Connect-GraphForCalendars)) {
        throw "No usable app-only Graph access - nothing was read."
    }

    # -- Directory ---------------------------------------------------------------
    Write-Host ""
    Write-Host "  Reading users and groups..." -ForegroundColor DarkGray
    $users = Get-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,mail,proxyAddresses,userType&$top=999'

    # Shared mailboxes and rooms are disabled accounts in Entra ID, so no
    # accountEnabled filter - every member with an address is a candidate.
    foreach ($u in $users) {
        if ($u.userType -ne 'Member' -or -not $u.mail) { continue }
        $addresses = Get-SmtpAddress -Primary $u.mail -ProxyAddresses $u.proxyAddresses -Upn $u.userPrincipalName
        $script:UserById[$u.id] = Add-Principal -Kind 'Mailbox' -Id $u.id -Name $u.displayName -Address $u.mail -Addresses $addresses
    }

    $groupCount = 0
    if ($script:SkipGroups) {
        Write-Warning "No Group.Read.All - group calendars cannot be recognised and will show up as MappedOwnerMissing."
    } else {
        try {
            $groups = Get-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/groups?$filter=mailEnabled eq true&$select=id,displayName,mail,proxyAddresses&$top=999'
            foreach ($g in $groups) {
                if (-not $g.mail) { continue }
                $addresses = Get-SmtpAddress -Primary $g.mail -ProxyAddresses $g.proxyAddresses
                [void](Add-Principal -Kind 'Group' -Id $g.id -Name $g.displayName -Address $g.mail -Addresses $addresses)
                $groupCount++
            }
        } catch {
            $script:SkipGroups = $true
            Write-Warning "Could not read groups ($($_.Exception.Message)) - group calendars will show up as MappedOwnerMissing."
        }
    }
    Write-Host "  [OK]   $($script:UserById.Count) users with an address, $groupCount mail-enabled groups." -ForegroundColor DarkGray

    # -- Scope -------------------------------------------------------------------
    $focus = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($m in $Mailbox) {
        $p = $script:AddressIndex[([string]$m).Trim().ToLowerInvariant()]
        if (-not $p -or $p.Kind -ne 'Mailbox') { throw "Mailbox '$m' was not found in the tenant." }
        [void]$focus.Add($p.Id)
    }

    # Permission reads whose every grant is reported. Other reads only fill in the
    # Rights column of rows that are already there.
    $reportEverything = ($focus.Count -eq 0) -and -not $Search
    $reportGrantKeys  = [System.Collections.Generic.HashSet[string]]::new()
    $calendarNames    = @{}
    $sources          = [System.Collections.Generic.List[object]]::new()

    # -- Read calendar lists and rights ---------------------------------------------
    if ($focus.Count -gt 0) {
        # Rights on the focus calendars first: their grantees are the people who
        # can have mapped them, so their lists are read too.
        $grants = Get-GraphCollection -Requests (New-UserRequest -UserIds @($focus) -Path 'calendar/calendarPermissions') -Activity 'Reading calendar permissions'
        foreach ($id in $focus) { [void]$reportGrantKeys.Add($id) }

        $scanIds = [System.Collections.Generic.HashSet[string]]::new($focus)
        foreach ($oid in $grants.Ok.Keys) {
            foreach ($g in $grants.Ok[$oid]) {
                if (-not $g.emailAddress.address) { continue }
                $p = Resolve-Principal -Address $g.emailAddress.address -Name $g.emailAddress.name
                if ($p.Kind -eq 'Mailbox') { [void]$scanIds.Add($p.Id) }
            }
        }
        Write-Host "  Reading the calendar lists of $($scanIds.Count) mailbox(es)..." -ForegroundColor DarkGray
        $lists = Get-GraphCollection -Requests (New-UserRequest -UserIds @($scanIds) -Path 'calendars') -Activity 'Reading calendar lists'

        # Owners the focus mailboxes mapped: their rights fill in the Rights column.
        $extraOwners = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($uid in @($focus)) {
            if (-not $lists.Ok.ContainsKey($uid)) { continue }
            foreach ($cal in $lists.Ok[$uid]) {
                if (-not $cal.owner.address) { continue }
                $p = Resolve-Principal -Address $cal.owner.address -Name $cal.owner.name
                if ($p.Kind -eq 'Mailbox' -and -not $grants.Ok.ContainsKey($p.Id) -and -not $grants.Failed.ContainsKey($p.Id)) {
                    [void]$extraOwners.Add($p.Id)
                }
            }
        }
        $more = Get-GraphCollection -Requests (New-UserRequest -UserIds @($extraOwners) -Path 'calendar/calendarPermissions') -Activity 'Reading calendar permissions'
        foreach ($k in $more.Ok.Keys)     { $grants.Ok[$k]     = $more.Ok[$k] }
        foreach ($k in $more.Failed.Keys) { $grants.Failed[$k] = $more.Failed[$k] }
    } else {
        $allIds = @($script:UserById.Keys)
        Write-Host "  Reading the calendar lists of $($allIds.Count) mailbox(es)..." -ForegroundColor DarkGray
        $lists = Get-GraphCollection -Requests (New-UserRequest -UserIds $allIds -Path 'calendars') -Activity 'Reading calendar lists'

        if ($Search) {
            # Only the permissions of what matches: the main calendar of an owner
            # whose name or address matches, and every own calendar whose name
            # matches - including a secondary one, which has rights of its own.
            $grantRequests = [ordered]@{}
            foreach ($uid in @($lists.Ok.Keys)) {
                $user = $script:UserById[$uid]
                if (Test-OwnerMatch -Principal $user) {
                    $grantRequests[$uid] = "/users/$uid/calendar/calendarPermissions"
                    [void]$reportGrantKeys.Add($uid)
                    $main = @($lists.Ok[$uid] | Where-Object { $_.isDefaultCalendar }) | Select-Object -First 1
                    $sources.Add([PSCustomObject]@{ Owner = $user; CalendarKey = 'main'; Name = $(if ($main) { $main.name } else { 'Main' }) })
                }
                foreach ($cal in $lists.Ok[$uid]) {
                    if ($cal.isDefaultCalendar -or -not (Test-OwnCalendar -User $user -Calendar $cal)) { continue }
                    if (-not (Test-SearchMatch -Value $cal.name)) { continue }
                    $key = "$uid|$($cal.id)"
                    $grantRequests[$key] = "/users/$uid/calendars/$([uri]::EscapeDataString([string]$cal.id))/calendarPermissions"
                    [void]$reportGrantKeys.Add($key)
                    $calendarNames[$key] = $cal.name
                    $sources.Add([PSCustomObject]@{ Owner = $user; CalendarKey = [string]$cal.id; Name = $cal.name })
                }
            }
            # The owner of a matching entry in somebody's list: its main rights
            # fill in the Rights column when no matching own calendar explains it.
            foreach ($uid in @($lists.Ok.Keys)) {
                $user = $script:UserById[$uid]
                foreach ($cal in $lists.Ok[$uid]) {
                    if ((Test-OwnCalendar -User $user -Calendar $cal) -or -not (Test-SearchMatch -Value $cal.name)) { continue }
                    $p = Resolve-Principal -Address $cal.owner.address -Name $cal.owner.name
                    if ($p.Kind -eq 'Mailbox' -and -not $grantRequests.Contains($p.Id)) {
                        $grantRequests[$p.Id] = "/users/$($p.Id)/calendar/calendarPermissions"
                    }
                }
            }
            Write-Host "  Reading the permissions on $($grantRequests.Count) matching calendar(s)..." -ForegroundColor DarkGray
            $grants = Get-GraphCollection -Requests $grantRequests -Activity 'Reading calendar permissions'
        } else {
            Write-Host "  Reading the permissions on $($lists.Ok.Count) main calendar(s)..." -ForegroundColor DarkGray
            $grants = Get-GraphCollection -Requests (New-UserRequest -UserIds @($lists.Ok.Keys) -Path 'calendar/calendarPermissions') -Activity 'Reading calendar permissions'
        }
    }

    # -- Pass 1: what is in each calendar list ------------------------------------
    $assignedEntries = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($uid in $lists.Ok.Keys) {
        $user = $script:UserById[$uid]
        foreach ($cal in $lists.Ok[$uid]) {
            if (Test-OwnCalendar -User $user -Calendar $cal) { continue }

            $owner = Resolve-Principal -Address $cal.owner.address -Name $cal.owner.name
            if ($owner.Key -eq $user.Key) { continue }

            $calendarKey  = 'main'
            $calendarName = $null
            if ($Search) {
                $nameHit = Test-SearchMatch -Value $cal.name
                if (-not $nameHit -and -not (Test-OwnerMatch -Principal $owner)) { continue }

                # An entry named after the keyword, from an owner who has a matching
                # secondary calendar, is taken to be that calendar - the list entry
                # itself carries no link back to the folder it came from.
                if ($nameHit -and $owner.Kind -eq 'Mailbox') {
                    $candidates = @($sources | Where-Object { $_.Owner.Id -eq $owner.Id -and $_.CalendarKey -ne 'main' })
                    if ($candidates.Count -gt 0) {
                        $pick = @($candidates | Where-Object { $_.Name -eq $cal.name })[0]
                        if (-not $pick) { $pick = $candidates[0] }
                        $calendarKey  = $pick.CalendarKey
                        $calendarName = $pick.Name
                    }
                }
            }

            $row = Get-PairRow -Owner $owner -User $user -CalendarKey $calendarKey -CalendarName $calendarName
            [void]$assignedEntries.Add("$uid|$($cal.id)")
            $row.Mapped   = 'Yes'
            $row.MappedAs = (@(@($row.MappedAs -split '; ') + $cal.name) | Where-Object { $_ } | Select-Object -Unique) -join '; '
            if ($cal.canEdit)             { $row.CanEdit        = 'Yes' } elseif (-not $row.CanEdit)        { $row.CanEdit        = 'No' }
            if ($cal.canViewPrivateItems) { $row.CanViewPrivate = 'Yes' } elseif (-not $row.CanViewPrivate) { $row.CanViewPrivate = 'No' }
        }
    }

    # -- Pass 2: the explicit rights on each calendar that was read ---------------
    $orgDefault = @{}
    foreach ($gk in @($grants.Ok.Keys)) {
        $ownerId      = $gk.Split('|')[0]
        $calendarKey  = if ($gk.Contains('|')) { $gk.Substring($gk.IndexOf('|') + 1) } else { 'main' }
        $calendarName = $calendarNames[$gk]
        $owner        = $script:UserById[$ownerId]
        $reportAll    = $reportEverything -or $reportGrantKeys.Contains($gk)

        foreach ($g in $grants.Ok[$gk]) {
            $address = $g.emailAddress.address
            $role    = [string]$g.role

            if (-not $address) {
                if (-not $g.isRemovable) {
                    # "My Organization" - the default for every internal user.
                    $orgDefault[$gk] = $role
                    if ($reportAll -and $role -notin @('none', 'freeBusyRead')) {
                        $org = [PSCustomObject]@{ Key = 'org'; Kind = 'Organization'; Name = $g.emailAddress.name; Address = '' }
                        $row = Get-PairRow -Owner $owner -User $org -CalendarKey $calendarKey -CalendarName $calendarName
                        $row.Rights = $role
                        $row.Status = 'OrgWideDefault'
                        $row.Mapped = ''
                        $row.Note   = 'Every internal user can open this calendar with this role'
                    }
                } elseif ($reportAll) {
                    # An entry without an address is an orphaned permission, typically
                    # left behind by a deleted account.
                    $orphan = [PSCustomObject]@{ Key = "orphan:$($g.id)"; Kind = 'Missing'; Name = $g.emailAddress.name; Address = '' }
                    $row = Get-PairRow -Owner $owner -User $orphan -CalendarKey $calendarKey -CalendarName $calendarName
                    $row.Rights = $role
                    $row.Status = 'GrantedToMissing'
                    $row.Mapped = ''
                }
                continue
            }

            $grantee = Resolve-Principal -Address $address -Name $g.emailAddress.name
            if ($grantee.Key -eq $owner.Key) { continue }

            if ($grantee.Kind -eq 'Mailbox') {
                $pairKey = Get-PairKey -Owner $owner -User $grantee -CalendarKey $calendarKey
                if (-not $reportAll -and -not $focus.Contains($grantee.Id) -and -not $script:Pairs.Contains($pairKey)) { continue }
                $row = Get-PairRow -Owner $owner -User $grantee -CalendarKey $calendarKey -CalendarName $calendarName
                $row.Rights = $role
                continue
            }
            if (-not $reportAll) { continue }

            $row = Get-PairRow -Owner $owner -User $grantee -CalendarKey $calendarKey -CalendarName $calendarName
            $row.Rights = $role
            $row.Mapped = ''
            switch ($grantee.Kind) {
                'Group' {
                    $row.Status = 'GrantedToGroup'
                    $row.Note   = 'Members are not expanded - check the group calendar list of each member'
                }
                'Missing' {
                    $row.Status = 'GrantedToMissing'
                    if ($script:SkipGroups) { $row.Note = 'Groups could not be read - this may be a group rather than a removed mailbox' }
                }
                'External' { $row.Status = 'SharedExternally' }
            }
        }
    }

    # -- Source: where a searched calendar lives ----------------------------------
    foreach ($src in $sources) {
        $here = [PSCustomObject]@{ Key = 'source'; Kind = ''; Name = ''; Address = '' }
        $row = Get-PairRow -Owner $src.Owner -User $here -CalendarKey $src.CalendarKey -CalendarName $src.Name
        $row.Status   = 'Source'
        $row.Mapped   = ''
        $row.MappedAs = $src.Name
        $row.Note     = if ($src.CalendarKey -eq 'main') { "The calendar lives in this mailbox (main calendar '$($src.Name)')" }
                        else { 'The calendar lives in this mailbox (secondary calendar)' }
    }

    # -- Status -------------------------------------------------------------------
    foreach ($row in $script:Pairs.Values) {
        if ($row.Status) { continue }

        if ($row.Mapped -eq 'Yes') {
            switch ($row.OwnerType) {
                'Group'    { $row.Status = 'MappedGroupCalendar' }
                'External' { $row.Status = 'MappedExternal' }
                'Missing'  {
                    $row.Status = 'MappedOwnerMissing'
                    if ($script:SkipGroups) { $row.Note = 'Groups could not be read - this may be a group calendar' }
                }
                default {
                    # Without the owner's rights there is nothing to compare with,
                    # so "without right" would be a guess - say so instead.
                    if (-not $grants.Ok.ContainsKey($row.GrantKey)) {
                        $row.Status = 'Mapped'
                        $row.Rights = 'unknown'
                        $row.Note   = if ($grants.Failed.ContainsKey($row.GrantKey)) { "Rights of this owner could not be read: $($grants.Failed[$row.GrantKey])" }
                                      else { 'Rights of this owner were not read (no mailbox, or its calendar list failed)' }
                    } elseif ($row.Rights -and $row.Rights -ne 'none') {
                        $row.Status = 'Mapped'
                    } else {
                        $row.Status = 'MappedWithoutRight'
                        $default = $orgDefault[$row.GrantKey]
                        $row.Note = if ($row.Rights -eq 'none') { "Explicit right is 'none'" }
                                    else { "No explicit right - organisation default is '$default'; a group or a secondary calendar may also explain it" }
                    }
                }
            }
            continue
        }

        # Only rights rows are left: an explicit right, and the question is whether
        # the calendar made it into the user's list.
        if ($row.Rights -eq 'none') { $row.Status = 'Drop'; continue }
        if ($lists.Ok.ContainsKey($row.UserKey)) {
            $row.Status = 'NotMapped'
            # A calendar of this owner the user does have, under a name that was
            # not matched to anything, may well be this one. Not when that name is
            # another calendar the owner has - an owner can share several - and,
            # for a secondary calendar, not when it is the owner's own name: that
            # entry is their main calendar.
            $userKey  = $row.UserKey
            $ownerKey = [string]$row.OwnerKey
            $notThis  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            if ($lists.Ok.ContainsKey($ownerKey)) {
                $ownerUser = $script:UserById[$ownerKey]
                foreach ($c in $lists.Ok[$ownerKey]) {
                    if ((Test-OwnCalendar -User $ownerUser -Calendar $c) -and [string]$c.name -ne [string]$row.Calendar) { [void]$notThis.Add([string]$c.name) }
                }
            }
            if ($row.Calendar -ne 'Main' -and $row.Owner) { [void]$notThis.Add([string]$row.Owner) }
            $other = @($lists.Ok[$userKey] | Where-Object {
                $_.owner.address -and -not $assignedEntries.Contains("$userKey|$($_.id)") -and
                -not $notThis.Contains([string]$_.name) -and
                (Resolve-Principal -Address $_.owner.address -Name $_.owner.name).Key -eq $row.OwnerKey
            } | ForEach-Object { $_.name })
            if ($other.Count -gt 0) {
                $row.Note = "Has a calendar of this owner as '$($other -join "', '")' - probably this one"
            }
        } else {
            $row.Status = 'NotChecked'
            $row.Mapped = 'Unknown'
            $row.Note   = if ($lists.Failed.ContainsKey($row.UserKey)) { "Calendar list could not be read: $($lists.Failed[$row.UserKey])" }
                          else { 'Calendar list was not read' }
        }
    }

    $results = @($script:Pairs.Values | Where-Object { $_.Status -ne 'Drop' })
    if ($focus.Count -gt 0) {
        $results = @($results | Where-Object { $focus.Contains([string]$_.OwnerKey) -or $focus.Contains([string]$_.UserKey) })
    }
    # Per calendar: where it lives first, then who has it.
    $results = @($results | Sort-Object Owner, Calendar, @{ Expression = { if ($_.Status -eq 'Source') { 0 } else { 1 } } }, Status, User)

    # -- Output -------------------------------------------------------------------
    Write-Host ""
    if ($results.Count -eq 0) {
        if ($Search) { Write-Host "  No calendar or owner matches '$Search'." -ForegroundColor DarkGray }
        else         { Write-Host "  No mapped or shared calendars found." -ForegroundColor DarkGray }
    } else {
        if ($Search) {
            $results | Format-Table Owner, Calendar, User, Status, Rights, MappedAs -AutoSize | Out-Host
        } else {
            $results | Format-Table Owner, User, Status, Rights, MappedAs -AutoSize | Out-Host
        }

        if (-not $OutputPath) {
            $OutputPath = Join-Path $outputDir ("CalendarMappings_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        }
        $results |
            Select-Object Owner, OwnerAddress, OwnerType, Calendar, User, UserAddress, UserType, Status, Mapped, MappedAs, Rights, CanEdit, CanViewPrivate, Note |
            Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
    }

    # -- Summary ------------------------------------------------------------------
    $explain = [ordered]@{
        Source              = 'where the searched calendar lives'
        Mapped              = 'in the calendar list, with an explicit right'
        MappedWithoutRight  = 'in the list without an explicit right on that calendar'
        MappedGroupCalendar = 'Microsoft 365 group calendar'
        MappedOwnerMissing  = 'owner no longer exists - stale entry'
        MappedExternal      = 'owner outside the tenant'
        NotMapped           = 'explicit right, not in the calendar list'
        NotChecked          = 'explicit right, calendar list could not be read'
        GrantedToGroup      = 'right granted to a group'
        GrantedToMissing    = 'right for an address that no longer exists'
        SharedExternally    = 'right for an address outside the tenant'
        OrgWideDefault      = 'everyone in the organisation gets more than free/busy'
    }
    Write-Host ""
    foreach ($status in $explain.Keys) {
        $n = @($results | Where-Object Status -eq $status).Count
        if ($n -gt 0) { Write-Host ("  {0,-20} {1,5}  {2}" -f $status, $n, $explain[$status]) -ForegroundColor Gray }
    }

    Write-Host ""
    $mappedCount = @($results | Where-Object Mapped -eq 'Yes').Count
    Write-Host "  Checked $($lists.Ok.Count) calendar list(s) - $mappedCount mapped calendar(s) found." -ForegroundColor Cyan
    if ($lists.NoMailbox.Count -gt 0) {
        Write-Host "  $($lists.NoMailbox.Count) user(s) have an address but no Exchange Online mailbox - skipped." -ForegroundColor DarkGray
    }

    # Permission reads are keyed by owner id, or owner id + calendar id.
    $problems = [ordered]@{}
    foreach ($k in $lists.Failed.Keys)  { $problems[$k] = $lists.Failed[$k] }
    foreach ($k in $grants.Failed.Keys) {
        $id = $k.Split('|')[0]
        if (-not $problems.Contains($id)) { $problems[$id] = $grants.Failed[$k] }
    }
    if ($problems.Count -gt 0) {
        Write-Host "  [WARN] $($problems.Count) mailbox(es) could not be read completely:" -ForegroundColor Yellow
        foreach ($id in (@($problems.Keys) | Select-Object -First 10)) {
            Write-Host "         $($script:UserById[$id].Address): $($problems[$id])" -ForegroundColor Yellow
        }
        if ($problems.Count -gt 10) { Write-Host "         ... and $($problems.Count - 10) more" -ForegroundColor Yellow }
        Write-Host "         403 here usually means an Application Access Policy or RBAC for Applications scopes the app." -ForegroundColor DarkGray
    }
    Write-Host ""
} finally {
    # The temporary app must never outlive the run, whatever happened above.
    Remove-TempApp
}
