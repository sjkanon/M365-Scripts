#Requires -Version 5.1
<#
.SYNOPSIS
    Move a shared calendar out of a user's mailbox into a resource mailbox of its
    own - every item and every permission along with it - and then remove the
    original. Preview by default; nothing changes without -Apply.

.DESCRIPTION
    A calendar like "Balie" often starts life as an extra calendar in somebody's
    mailbox, shared with the whole front desk. It then hangs on that one account:
    it leaves with the person, it cannot be booked like a room, and every
    permission change goes through the owner. This script lifts it out:

      1. Read     The calendar's items - single items, recurring series with their
                  exceptions and cancelled occurrences, attachments, categories -
                  and its permissions. A preview run stops here.
      2. Backup   Everything that was read goes to a JSON file before anything is
                  created, plus every attachment that cannot be copied as-is.
      3. Mailbox  A Room (default) or Equipment mailbox, with the source mailbox's
                  language and time zone, and calendar processing set up for a
                  shared calendar: auto-accept, overlapping items allowed.
      4. Rights   Every permission of the original calendar with its exact
                  Exchange access rights, plus the original owner.
      5. Items    Every item copied into the resource calendar.
      6. Verify   Every source item must have a complete copy.
      7. Remove   Only with -RemoveSourceCalendar and only after a clean
                  verification: the original calendar is deleted.

    Items
    -----
      - Recurring series stay series. Exceptions (a moved or edited occurrence) are
        applied to the copy and cancelled occurrences are cancelled in it, by
        comparing both series occurrence by occurrence. For series without an end
        date this is done up to -SeriesHorizonDays ahead.
      - Times keep their original time zone, so a weekly 9:00 item is still 9:00
        after the switch to or from daylight saving time.
      - Nobody is invited. A copied meeting would otherwise send a fresh
        invitation to every attendee; instead organizer and attendees are listed
        at the bottom of the item's body. The resource mailbox is the organizer of
        every copy.
      - File attachments up to 3 MB are copied, inline images included. Larger
        files and attached Outlook items are saved to the backup folder and
        listed at the end - they cannot be attached through this API in one call.
      - Categories keep their colour: the ones in use are created in the resource
        mailbox's category list.
      - Every copy carries the id of its source item in a hidden property. A run
        that stops halfway can simply be started again: complete copies are
        skipped, a half-finished series is removed and copied again.

    Rights
    ------
    The principals come from Graph, the exact access rights (including custom
    ones) from Get-MailboxFolderPermission. Default and Anonymous are carried over
    as they were. Not carried over, and reported: delegate flags (a resource
    mailbox has no delegates), people outside the tenant (re-share by hand) and
    orphaned entries of deleted accounts. The original owner gets
    -SourceOwnerRights (default Owner - what they had).

    -SendSharingInvitation sends the usual "X shared a calendar with you" mail, so
    users add the new calendar with one click. That only works for Reviewer,
    Editor, LimitedDetails and AvailabilityOnly; other rights are granted without
    an invitation.

    After the switch
    ----------------
    Users who had the original calendar in their list keep an entry that stops
    working once it is removed. Before removing, find them with
        Get-CalendarMappings.ps1 -Search "<calendar name>"
    The copy is a snapshot: run it when the calendar is quiet and tell users to
    switch right after, or changes made to the original in between are lost.

    Access
    ------
    Exchange Online: Exchange Administrator (New-Mailbox, folder permissions).
    An existing Exchange session is reused.

    Graph: application permission Calendars.ReadWrite, plus
    MailboxSettings.ReadWrite for the categories (optional - without it
    categories are copied by name, without colour). Obtained the same three ways
    as in Remove-PhishingMessage.ps1: an existing app-only session, -ClientId
    with -ClientSecret or -CertificateThumbprint, or a temporary App Registration
    that is removed again when the run ends. The Graph calls are plain REST, so
    the Exchange/Graph MSAL clash does not apply.

.PARAMETER Mailbox
    The user mailbox that holds the calendar (UPN or SMTP address).

.PARAMETER Calendar
    Name of the calendar in that mailbox, as shown in Outlook (e.g. "Balie").
    Must be a calendar the user owns and not their main calendar.

.PARAMETER ResourceName
    Display name of the new resource mailbox. Default: the calendar's name.

.PARAMETER ResourceAddress
    Primary SMTP address of the new resource mailbox. Default: the name as an
    alias at the source mailbox's domain (balie@contoso.com). An existing room or
    equipment mailbox at this address is reused, which is how a second run
    continues where the first one stopped.

.PARAMETER ResourceType
    Room (default) or Equipment.

.PARAMETER SourceOwnerRights
    What the original owner gets on the new calendar: Owner (default),
    PublishingEditor, Editor, Reviewer or None.

.PARAMETER SendSharingInvitation
    Send every user a sharing invitation for the new calendar.

.PARAMETER Apply
    Actually create, grant and copy. Without it the script only reports what it
    would do.

.PARAMETER RemoveSourceCalendar
    Delete the original calendar after a clean verification. Needs -Apply.
    Asks for the calendar name as confirmation unless -Force is given.

.PARAMETER Force
    Skip the typed confirmation before the original calendar is removed. In a
    non-interactive session (a scheduler, an RMM agent) nobody can type it, so
    there the original is only removed with -Force.

.PARAMETER PassThru
    Return a result object (ResourceAddress, Items, Verified, SourceRemoved,
    BackupPath) for a calling script. Move-SharedCalendar.ps1 uses it.

.PARAMETER SeriesHorizonDays
    How far ahead exceptions and cancellations of a series without an end date
    are compared (default 1095 days).

.PARAMETER BackupPath
    Folder for the backup and saved attachments. Default:
    C:\Temp\CalendarConvert_<calendar>_<timestamp> (~/Downloads on macOS/Linux).

.PARAMETER TenantId
    Tenant ID or domain. Optional: taken from the Exchange session otherwise.

.PARAMETER ClientId
    Your own App Registration for app-only Graph access.

.PARAMETER ClientSecret
    Client secret for -ClientId.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for -ClientId (via Connect-MgGraph).

.EXAMPLE
    # Preview: what would happen to Jan's calendar "Balie"?
    .\Convert-SharedCalendarToResource.ps1 -Mailbox jan@contoso.com -Calendar Balie

.EXAMPLE
    # Create the room mailbox, copy everything, invite the users - keep the original
    .\Convert-SharedCalendarToResource.ps1 -Mailbox jan@contoso.com -Calendar Balie -Apply -SendSharingInvitation

.EXAMPLE
    # The same, and remove the original once every item is verified
    .\Convert-SharedCalendarToResource.ps1 -Mailbox jan@contoso.com -Calendar Balie -Apply -RemoveSourceCalendar

.NOTES
    Author: Sjoerd Kanon
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $Mailbox,
    [Parameter(Mandatory)] [string] $Calendar,
    [string] $ResourceName,
    [string] $ResourceAddress,
    [ValidateSet('Room', 'Equipment')]
    [string] $ResourceType = 'Room',
    [ValidateSet('Owner', 'PublishingEditor', 'Editor', 'Reviewer', 'None')]
    [string] $SourceOwnerRights = 'Owner',
    [switch] $SendSharingInvitation,
    [switch] $Apply,
    [switch] $RemoveSourceCalendar,
    [switch] $Force,
    [switch] $PassThru,
    [ValidateRange(30, 3650)]
    [int]    $SeriesHorizonDays = 1095,
    [string] $BackupPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $ClientSecret,
    [string] $CertificateThumbprint
)

if ($RemoveSourceCalendar -and -not $Apply) {
    throw "-RemoveSourceCalendar needs -Apply: the original is only removed after its copy has been made and verified."
}
if (-not $ResourceName) { $ResourceName = $Calendar }

# -- Output -------------------------------------------------------------------
function Write-Step { param([string] $Message) Write-Host ""; Write-Host "  $Message" -ForegroundColor Cyan }
function Write-Ok   { param([string] $Message) Write-Host "  [OK]   $Message" -ForegroundColor Green }
function Write-Warn { param([string] $Message) Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
function Write-Info { param([string] $Message) Write-Host "         $Message" -ForegroundColor DarkGray }

function Test-CanPrompt {
    # Read-Host needs a person: not in a service, not under -NonInteractive.
    if (-not [Environment]::UserInteractive) { return $false }
    return -not ([Environment]::GetCommandLineArgs() | Where-Object { $_ -match '^-noni' })
}

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Convert Shared Calendar to Resource Mailbox" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
if ($Apply) {
    $modeText = if ($RemoveSourceCalendar) { 'APPLY - create, copy, verify, then REMOVE the original' } else { 'APPLY - create, copy and verify; the original stays' }
    Write-Host "  Mode      : $modeText" -ForegroundColor $(if ($RemoveSourceCalendar) { 'Red' } else { 'Yellow' })
} else {
    Write-Host "  Mode      : PREVIEW - nothing is changed" -ForegroundColor Yellow
}
Write-Host "  Source    : $Mailbox > '$Calendar'" -ForegroundColor DarkGray
Write-Host ""

$RequiredRoles = @('Calendars.ReadWrite')
$OptionalRoles = @('MailboxSettings.ReadWrite')
$RoleAlternatives = @{
    'Calendars.ReadWrite'       = @('Calendars.ReadWrite')
    'MailboxSettings.ReadWrite' = @('MailboxSettings.ReadWrite')
}

# Marks every copy with the id of its source item, so a second run knows what is
# already there. "Complete" is only set once attachments and series exceptions
# are in place too.
$PropertySet    = '{6f1c0e52-8a7d-4b39-9c64-2d5e8b1f3a70}'
$PropSourceId   = "String $PropertySet Name M365ScriptsSourceEventId"
$PropComplete   = "String $PropertySet Name M365ScriptsCopyComplete"

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
$script:AdminHeaders    = $null
$script:SkipCategories  = $false
$script:GraphCliClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'

function Invoke-GraphAdmin {
    param([string] $Method = 'GET', [Parameter(Mandatory)] [string] $Uri, $Body)
    $params = @{ Method = $Method; Uri = $Uri; Headers = $script:AdminHeaders; ErrorAction = 'Stop' }
    if ($Body) {
        $params['Body']        = ($Body | ConvertTo-Json -Depth 6)
        $params['ContentType'] = 'application/json'
    }
    return Invoke-RestMethod @params
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
        One entry point for every Graph call. Bodies go out as UTF-8 bytes:
        Windows PowerShell otherwise encodes a string body as ISO-8859-1, and
        every accented letter in a subject would arrive mangled.
    #>
    param([string] $Method = 'GET', [Parameter(Mandatory)] [string] $Uri, $Body)

    $json = $null
    if ($null -ne $Body) { $json = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 20 -Compress } }

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            if ($script:AppOnlyHeaders) {
                Update-AppOnlyToken
                $params = @{ Method = $Method; Uri = $Uri; Headers = $script:AppOnlyHeaders; ErrorAction = 'Stop' }
                if ($json) {
                    $params['Body']        = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $params['ContentType'] = 'application/json; charset=utf-8'
                }
                return Invoke-RestMethod @params
            }
            $params = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
            if ($json) { $params['Body'] = $json; $params['ContentType'] = 'application/json' }
            return Invoke-MgGraphRequest @params
        } catch {
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

function Get-GraphErrorText {
    # The Graph error code and message, rather than "The remote server returned an error".
    param($ErrorRecord)
    try {
        $detail = ($ErrorRecord.ErrorDetails.Message | ConvertFrom-Json).error
        if ($detail.code) { return "$($detail.code): $($detail.message)" }
    } catch {}
    return "$($ErrorRecord.Exception.Message)"
}

function Get-TokenClaim {
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
    param([string[]] $Have, [string[]] $Wanted)
    return @($Wanted | Where-Object {
        $alternatives = $RoleAlternatives[$_]
        -not ($Have | Where-Object { $alternatives -contains $_ })
    })
}

function Confirm-AppRole {
    # Re-mints until the roles are in the token; an early token stays powerless for an hour.
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

function Connect-GraphForCalendar {
    <# Establishes app-only Calendars.ReadWrite. Returns $true when Graph is usable. #>
    param([string] $Tenant)

    # 1. An app-only session the caller already established - unless an app was
    #    named with -ClientId, which always wins.
    $ctx = $null
    if (-not $ClientId) { try { $ctx = Get-MgContext -ErrorAction SilentlyContinue } catch {} }
    if ($ctx -and $ctx.AuthType -eq 'AppOnly') {
        $have = @($ctx.Scopes)
        if ((Get-MissingRole -Have $have -Wanted $RequiredRoles).Count -gt 0) {
            Write-Warning "The existing app-only Graph session lacks Calendars.ReadWrite."
            return $false
        }
        if ((Get-MissingRole -Have $have -Wanted $OptionalRoles).Count -gt 0) { $script:SkipCategories = $true }
        Write-Host "  [OK]   Using the existing app-only Graph session." -ForegroundColor DarkGray
        return $true
    }

    # 2. Your own app registration.
    if ($ClientId) {
        if (-not $Tenant) { Write-Warning "-ClientId needs -TenantId."; return $false }
        if ($ClientSecret) {
            try {
                Get-AppOnlyTokenByRest -Tenant $Tenant -App $ClientId -Secret $ClientSecret
            } catch {
                Write-Warning "Could not obtain a token for the supplied app: $($_.Exception.Message)"
                return $false
            }
            $claims = Get-TokenClaim -Jwt $script:AccessToken
            $have   = if ($claims) { @($claims.roles) } else { @() }
            if ((Get-MissingRole -Have $have -Wanted $RequiredRoles).Count -gt 0) {
                Write-Warning "App $ClientId has no Calendars.ReadWrite application permission (token carries: $(if ($have) { $have -join ', ' } else { 'nothing' }))."
                return $false
            }
            if ((Get-MissingRole -Have $have -Wanted $OptionalRoles).Count -gt 0) { $script:SkipCategories = $true }
            Write-Host "  [OK]   App-only token obtained (supplied app, REST)." -ForegroundColor DarkGray
            return $true
        }
        if ($CertificateThumbprint) {
            try {
                Connect-MgGraph -ClientId $ClientId -TenantId $Tenant -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
            } catch {
                Write-Warning "Could not connect with the supplied certificate: $($_.Exception.Message)"
                return $false
            }
            if ((Get-MissingRole -Have @((Get-MgContext).Scopes) -Wanted $OptionalRoles).Count -gt 0) { $script:SkipCategories = $true }
            Write-Host "  [OK]   Connected with the supplied certificate." -ForegroundColor DarkGray
            return $true
        }
        Write-Warning "-ClientId needs -ClientSecret or -CertificateThumbprint."
        return $false
    }

    # 3. A short-lived app, removed again at the end. Plain REST, no SDK.
    if (-not $Tenant) { Write-Warning "No tenant known for the temporary App Registration. Pass -TenantId."; return $false }
    $wanted = $RequiredRoles + $OptionalRoles
    try {
        Write-Host "  No app-only Graph access yet - setting up a temporary App Registration." -ForegroundColor Cyan
        Write-Host "  Required role: Global Administrator or Privileged Role Administrator (one-time)" -ForegroundColor DarkGray

        $adminToken = Get-DelegatedTokenByDeviceCode -Tenant $Tenant -Scopes @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
        $script:AdminHeaders = @{ Authorization = "Bearer $adminToken" }
        Write-Host "  [OK]   Signed in." -ForegroundColor DarkGray

        $appName = "CalendarConvert-Temp-$(Get-Date -Format 'yyyyMMddHHmmss')"
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
                    -Body @{ passwordCredential = @{ displayName = 'temp'; endDateTime = (Get-Date).AddHours(4).ToString('o') } }

        $ok = $false
        for ($i = 1; $i -le 8; $i++) {
            try {
                Get-AppOnlyTokenByRest -Tenant $Tenant -App $app.appId -Secret $secret.secretText
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

        if (-not (Confirm-AppRole -Required $wanted)) { Remove-TempApp; return $false }
        return $true
    } catch {
        Write-Warning "Temporary app setup failed: $($_.Exception.Message)"
        Remove-TempApp
        return $false
    }
}

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

function ConvertTo-UrlId { param([string] $Id) return [uri]::EscapeDataString($Id) }

# ==============================================================================
#  Dates and time zones
# ==============================================================================

function ConvertTo-UtcDate {
    <#
        Graph dates arrive as strings in Windows PowerShell, but PowerShell 7's
        JSON parser already turns them into DateTime - with a Kind that depends on
        whether the string had a "Z". Every date goes through here so both end up
        as the same UTC value.
    #>
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    if ($Value -is [DateTimeOffset]) { return $Value.UtcDateTime }
    if ($Value -is [datetime]) {
        switch ($Value.Kind) {
            'Utc'   { return $Value }
            'Local' { return $Value.ToUniversalTime() }
            default { return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc) }
        }
    }
    $styles = [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
    return [datetime]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, $styles)
}

function Format-DateOnly {
    param($Value)
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd') }
    $text = [string]$Value
    return $text.Substring(0, [Math]::Min(10, $text.Length))
}

$script:MailboxTimeZone = [PSCustomObject]@{ Info = [TimeZoneInfo]::Utc; Id = 'UTC' }
$script:TimeZoneCache   = @{}

function Resolve-TimeZone {
    <#
        The time zone an item was created in (originalStartTimeZone). Windows ids
        such as "W. Europe Standard Time" resolve directly; UTC markers map to
        UTC; anything else (a customised zone from an iCal import) falls back to
        the source mailbox's own time zone.
    #>
    param([string] $Id)
    $key = [string]$Id
    if ($script:TimeZoneCache.ContainsKey($key)) { return $script:TimeZoneCache[$key] }

    $result = $null
    if ($Id -match '^(tzone://Microsoft/Utc|UTC|Etc/UTC|Coordinated Universal Time)$') {
        $result = [PSCustomObject]@{ Info = [TimeZoneInfo]::Utc; Id = 'UTC' }
    } elseif ($Id) {
        try { $result = [PSCustomObject]@{ Info = [TimeZoneInfo]::FindSystemTimeZoneById($Id); Id = $Id } } catch { $result = $null }
    }
    if (-not $result) { $result = $script:MailboxTimeZone }
    $script:TimeZoneCache[$key] = $result
    return $result
}

function ConvertTo-EventTime {
    <#
        Graph returns every time in UTC (no Prefer header is sent). A series has
        to be created in its own time zone, or a weekly 9:00 item drifts an hour
        at every daylight saving switch - so the time is converted back.
    #>
    param($Time, [string] $TimeZoneId, [bool] $IsAllDay)
    $utc = ConvertTo-UtcDate $Time.dateTime
    if ($IsAllDay) {
        # All-day items are floating: midnight to midnight, in whatever zone is given.
        $zone = if ($Time.timeZone) { [string]$Time.timeZone } else { 'UTC' }
        return @{ dateTime = $utc.ToString('yyyy-MM-ddT00:00:00'); timeZone = $zone }
    }
    $tz    = Resolve-TimeZone -Id $TimeZoneId
    $local = [TimeZoneInfo]::ConvertTimeFromUtc($utc, $tz.Info)
    return @{ dateTime = $local.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = $tz.Id }
}

function Get-OccurrenceKey {
    # The slot an occurrence was generated for - the same in source and copy.
    param($Instance)
    $when = ConvertTo-UtcDate $Instance.originalStart
    if (-not $when) { $when = ConvertTo-UtcDate $Instance.start.dateTime }
    return $when.ToString('yyyy-MM-ddTHH:mm')
}

# ==============================================================================
#  Item payloads
# ==============================================================================

function Add-MigrationNote {
    <#
        Attendees are not copied as attendees - that would send each of them a new
        invitation from the resource mailbox. They are listed at the bottom of the
        body instead, together with the original organizer.
    #>
    param($Item)
    $content   = [string]$Item.body.content
    $attendees = @($Item.attendees | Where-Object { $_.emailAddress } | ForEach-Object {
        if ($_.emailAddress.name -and $_.emailAddress.name -ne $_.emailAddress.address) { "$($_.emailAddress.name) <$($_.emailAddress.address)>" }
        else { [string]$_.emailAddress.address }
    })
    $organizer = $Item.organizer.emailAddress
    $foreignOrganizer = $organizer -and $organizer.address -and -not $script:SourceAddresses.Contains([string]$organizer.address)
    if ($attendees.Count -eq 0 -and -not $foreignOrganizer) { return $content }

    $lines = @("Copied from the calendar '$Calendar' of $($script:SourceUpn).")
    if ($organizer -and $organizer.address) { $lines += "Organizer: $($organizer.name) <$($organizer.address)>" }
    if ($attendees.Count -gt 0) { $lines += "Attendees: $($attendees -join '; ')" }

    if ([string]$Item.body.contentType -eq 'text') {
        return $content + "`r`n`r`n----`r`n" + ($lines -join "`r`n")
    }
    $note = '<hr><p style="font-size:9pt;color:#666666">' + (($lines | ForEach-Object { [System.Net.WebUtility]::HtmlEncode($_) }) -join '<br>') + '</p>'
    $at = $content.LastIndexOf('</body>', [StringComparison]::OrdinalIgnoreCase)
    if ($at -ge 0) { return $content.Insert($at, $note) }
    return $content + $note
}

function ConvertTo-Recurrence {
    <#
        Graph returns every pattern field, filled with defaults for the ones the
        pattern type does not use. Only the fields that belong to the type are
        sent back, so the create call never trips over a month of 0.
    #>
    param($Recurrence, [string] $TimeZoneId)
    $p = $Recurrence.pattern
    $r = $Recurrence.range

    $pattern = [ordered]@{ type = [string]$p.type; interval = [int]$p.interval }
    switch ([string]$p.type) {
        'weekly'          { $pattern.daysOfWeek = @($p.daysOfWeek); $pattern.firstDayOfWeek = [string]$p.firstDayOfWeek }
        'absoluteMonthly' { $pattern.dayOfMonth = [int]$p.dayOfMonth }
        'relativeMonthly' { $pattern.daysOfWeek = @($p.daysOfWeek); $pattern.index = [string]$p.index }
        'absoluteYearly'  { $pattern.dayOfMonth = [int]$p.dayOfMonth; $pattern.month = [int]$p.month }
        'relativeYearly'  { $pattern.daysOfWeek = @($p.daysOfWeek); $pattern.index = [string]$p.index; $pattern.month = [int]$p.month }
    }

    $range = [ordered]@{ type = [string]$r.type; startDate = Format-DateOnly $r.startDate }
    if ($TimeZoneId) { $range.recurrenceTimeZone = $TimeZoneId }
    if ($r.type -eq 'endDate')  { $range.endDate = Format-DateOnly $r.endDate }
    if ($r.type -eq 'numbered') { $range.numberOfOccurrences = [int]$r.numberOfOccurrences }

    return [ordered]@{ pattern = $pattern; range = $range }
}

function New-EventPayload {
    param($Item, [switch] $ForException)

    $isAllDay = [bool]$Item.isAllDay
    $payload = [ordered]@{
        subject           = [string]$Item.subject
        body              = @{ contentType = $(if ($Item.body.contentType) { [string]$Item.body.contentType } else { 'html' }); content = (Add-MigrationNote -Item $Item) }
        isAllDay          = $isAllDay
        start             = ConvertTo-EventTime -Time $Item.start -TimeZoneId $Item.originalStartTimeZone -IsAllDay $isAllDay
        end               = ConvertTo-EventTime -Time $Item.end   -TimeZoneId $Item.originalEndTimeZone   -IsAllDay $isAllDay
        showAs            = $(if ($Item.showAs) { [string]$Item.showAs } else { 'busy' })
        sensitivity       = $(if ($Item.sensitivity) { [string]$Item.sensitivity } else { 'normal' })
        importance        = $(if ($Item.importance) { [string]$Item.importance } else { 'normal' })
        categories        = @($Item.categories | Where-Object { $_ })
        isReminderOn      = [bool]$Item.isReminderOn
        responseRequested = $false
    }
    if ($Item.isReminderOn) { $payload.reminderMinutesBeforeStart = [int]$Item.reminderMinutesBeforeStart }
    if ($Item.location -and $Item.location.displayName) { $payload.location = @{ displayName = [string]$Item.location.displayName } }
    if (-not $ForException -and $Item.recurrence) {
        $zone = if ($isAllDay) { $script:MailboxTimeZone.Id } else { (Resolve-TimeZone -Id $Item.originalStartTimeZone).Id }
        $payload.recurrence = ConvertTo-Recurrence -Recurrence $Item.recurrence -TimeZoneId $zone
    }
    return $payload
}

function Get-SeriesWindow {
    <#
        The stretch in which a series' occurrences are compared. From the first
        occurrence to the end date, or - for a series that never ends or ends
        after N occurrences - to -SeriesHorizonDays from now.
    #>
    param($Master)
    $first = (ConvertTo-UtcDate $Master.start.dateTime).AddDays(-1)
    if ([string]$Master.recurrence.range.type -eq 'endDate') {
        $last = [datetime]::ParseExact((Format-DateOnly $Master.recurrence.range.endDate), 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture).AddDays(2)
    } else {
        $from = if ($first -gt [datetime]::UtcNow) { $first } else { [datetime]::UtcNow }
        $last = $from.AddDays($SeriesHorizonDays)
    }
    return [PSCustomObject]@{ Start = $first.ToString('yyyy-MM-ddTHH:mm:ssZ'); End = $last.ToString('yyyy-MM-ddTHH:mm:ssZ') }
}

# ==============================================================================
#  Copy
# ==============================================================================

$script:Stats = [ordered]@{
    Copied = 0; AlreadyThere = 0; Redone = 0; Failed = 0
    Exceptions = 0; Cancelled = 0; UnmatchedSeries = 0
    Attachments = 0; AttachmentsSaved = 0
}
$script:Failures = [System.Collections.Generic.List[string]]::new()
$script:Notices  = [System.Collections.Generic.List[string]]::new()

function Save-Attachment {
    # Anything that cannot be attached in one call is kept on disk, never dropped.
    param($Attachment, [string] $SourceEventId, [string] $Subject)
    $safeSubject = (($Subject -replace '[\\/:*?"<>|]', '_').Trim())
    if (-not $safeSubject) { $safeSubject = 'item' }
    if ($safeSubject.Length -gt 60) { $safeSubject = $safeSubject.Substring(0, 60) }
    $folder = Join-Path (Join-Path $script:BackupDir 'attachments') $safeSubject
    if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    $name = (([string]$Attachment.name) -replace '[\\/:*?"<>|]', '_')
    if (-not $name) { $name = 'attachment' }

    switch ([string]$Attachment.'@odata.type') {
        '#microsoft.graph.fileAttachment' {
            [System.IO.File]::WriteAllBytes((Join-Path $folder $name), [Convert]::FromBase64String([string]$Attachment.contentBytes))
        }
        '#microsoft.graph.itemAttachment' {
            $mime = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/users/$($script:SourceId)/events/$(ConvertTo-UrlId $SourceEventId)/attachments/$(ConvertTo-UrlId $Attachment.id)/`$value"
            Set-Content -Path (Join-Path $folder "$name.eml") -Value ([string]$mime) -Encoding UTF8
        }
        default {
            Set-Content -Path (Join-Path $folder "$name.txt") -Value ("$($Attachment.name)`r`n$($Attachment.sourceUrl)") -Encoding UTF8
        }
    }
    $script:Stats.AttachmentsSaved++
    $script:Notices.Add("Attachment '$($Attachment.name)' of '$Subject' saved to $folder - attach it by hand")
}

function Copy-EventAttachment {
    param([string] $SourceEventId, [string] $TargetEventId, [string] $Subject)
    $attachments = Get-GraphPaged -Uri "https://graph.microsoft.com/v1.0/users/$($script:SourceId)/events/$(ConvertTo-UrlId $SourceEventId)/attachments"
    foreach ($a in $attachments) {
        $isFile = [string]$a.'@odata.type' -eq '#microsoft.graph.fileAttachment'
        if ($isFile -and -not $a.contentBytes) {
            $a = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/users/$($script:SourceId)/events/$(ConvertTo-UrlId $SourceEventId)/attachments/$(ConvertTo-UrlId $a.id)"
        }
        # 3 MB is the limit for attaching in one call; larger needs an upload session.
        if ($isFile -and [int64]$a.size -le 3MB) {
            $body = [ordered]@{
                '@odata.type' = '#microsoft.graph.fileAttachment'
                name          = [string]$a.name
                contentType   = [string]$a.contentType
                contentBytes  = [string]$a.contentBytes
                isInline      = [bool]$a.isInline
            }
            if ($a.contentId) { $body.contentId = [string]$a.contentId }
            Invoke-Graph -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($script:TargetId)/events/$(ConvertTo-UrlId $TargetEventId)/attachments" -Body $body | Out-Null
            $script:Stats.Attachments++
        } else {
            Save-Attachment -Attachment $a -SourceEventId $SourceEventId -Subject $Subject
        }
    }
}

function Sync-SeriesException {
    <#
        A new series has every occurrence its pattern generates. The source series
        may have moved or edited some (exceptions) and cancelled others. Both
        series are listed over the same window and matched on the slot each
        occurrence was generated for: an edited source occurrence is applied to
        its copy, a slot the source no longer has is cancelled in the copy.
    #>
    param($Master, [string] $TargetMasterId)

    $window  = $script:SeriesWindows[$Master.id]
    $sources = @($script:SourceInstances[$Master.id])
    $targets = @(Get-GraphPaged -Uri ("https://graph.microsoft.com/v1.0/users/$($script:TargetId)/events/$(ConvertTo-UrlId $TargetMasterId)/instances" +
                                      "?startDateTime=$($window.Start)&endDateTime=$($window.End)&`$select=id,originalStart,start,type&`$top=100"))

    $targetBySlot = @{}
    foreach ($t in $targets) { $targetBySlot[(Get-OccurrenceKey -Instance $t)] = $t }
    $sourceSlots = [System.Collections.Generic.HashSet[string]]::new()
    $matched = 0
    foreach ($s in $sources) {
        $slot = Get-OccurrenceKey -Instance $s
        [void]$sourceSlots.Add($slot)
        if ($targetBySlot.ContainsKey($slot)) { $matched++ }
    }

    # If the two series do not line up, "cancelling what the source lacks" would
    # wipe the copy. Leave it as the pattern made it and say so.
    if ($sources.Count -gt 0 -and $matched -lt [Math]::Ceiling($sources.Count / 2)) {
        $script:Stats.UnmatchedSeries++
        $script:Notices.Add("Series '$($Master.subject)': only $matched of $($sources.Count) occurrences lined up - exceptions and cancellations were NOT applied, check this series by hand")
        return
    }

    foreach ($s in $sources) {
        if ([string]$s.type -ne 'exception') { continue }
        $t = $targetBySlot[(Get-OccurrenceKey -Instance $s)]
        if (-not $t) { continue }
        Invoke-Graph -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$($script:TargetId)/events/$(ConvertTo-UrlId $t.id)" -Body (New-EventPayload -Item $s -ForException) | Out-Null
        $script:Stats.Exceptions++
        if ($s.hasAttachments) { $script:Notices.Add("Series '$($Master.subject)': the exception on $(Get-OccurrenceKey -Instance $s) has its own attachments - not copied") }
    }
    foreach ($slot in @($targetBySlot.Keys)) {
        if ($sourceSlots.Contains($slot)) { continue }
        Invoke-Graph -Method DELETE -Uri "https://graph.microsoft.com/v1.0/users/$($script:TargetId)/events/$(ConvertTo-UrlId $targetBySlot[$slot].id)" | Out-Null
        $script:Stats.Cancelled++
    }
}

function Get-ExistingCopy {
    # Copies already in the resource calendar, by source item id.
    $filter = [uri]::EscapeDataString("id eq '$PropSourceId' or id eq '$PropComplete'")
    $events = Get-GraphPaged -Uri ("https://graph.microsoft.com/v1.0/users/$($script:TargetId)/calendar/events" +
                                   "?`$select=id,subject&`$top=100&`$expand=singleValueExtendedProperties(`$filter=$filter)")
    $copies = @{}
    foreach ($e in $events) {
        $sourceId = $null; $complete = $false
        foreach ($p in @($e.singleValueExtendedProperties)) {
            if ([string]$p.id -like '*Name M365ScriptsSourceEventId') { $sourceId = [string]$p.value }
            if ([string]$p.id -like '*Name M365ScriptsCopyComplete')  { $complete = ([string]$p.value -eq 'true') }
        }
        if ($sourceId) { $copies[$sourceId] = [PSCustomObject]@{ Id = $e.id; Complete = $complete } }
    }
    return $copies
}

function Copy-CalendarEvent {
    param($Item, $Existing)

    $copy = $Existing[$Item.id]
    if ($copy -and $copy.Complete) { $script:Stats.AlreadyThere++; return }
    if ($copy) {
        # Left half-done by an earlier run: attachments or series exceptions may
        # be missing, so start this one over.
        Invoke-Graph -Method DELETE -Uri "https://graph.microsoft.com/v1.0/users/$($script:TargetId)/events/$(ConvertTo-UrlId $copy.Id)" | Out-Null
        $script:Stats.Redone++
    }

    $isSeries = [string]$Item.type -eq 'seriesMaster'
    $needsMore = $isSeries -or [bool]$Item.hasAttachments
    $payload = New-EventPayload -Item $Item
    $props = @(@{ id = $PropSourceId; value = [string]$Item.id })
    if (-not $needsMore) { $props += @{ id = $PropComplete; value = 'true' } }
    $payload.singleValueExtendedProperties = $props

    $new = Invoke-Graph -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($script:TargetId)/calendar/events" -Body $payload
    if ($needsMore) {
        if ($Item.hasAttachments) { Copy-EventAttachment -SourceEventId $Item.id -TargetEventId $new.id -Subject ([string]$Item.subject) }
        if ($isSeries)             { Sync-SeriesException -Master $Item -TargetMasterId $new.id }
        Invoke-Graph -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$($script:TargetId)/events/$(ConvertTo-UrlId $new.id)" `
            -Body @{ singleValueExtendedProperties = @(@{ id = $PropComplete; value = 'true' }) } | Out-Null
    }
    $script:Stats.Copied++
}

# ==============================================================================
#  Exchange helpers
# ==============================================================================

function ConvertTo-MailboxAlias {
    # "Balie Zuid-Oost" -> "balie-zuid-oost"; accents are dropped, not the letter.
    param([string] $Name)
    $plain = -join ($Name.Normalize([Text.NormalizationForm]::FormD).ToCharArray() |
                    Where-Object { [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne 'NonSpacingMark' })
    $alias = ($plain.ToLowerInvariant() -replace '[^a-z0-9]+', '-').Trim('-')
    if (-not $alias) { $alias = "calendar-$(Get-Date -Format 'yyyyMMddHHmm')" }
    return $alias
}

function Get-GraphRoleRight {
    # Fallback when Exchange cannot tell the exact rights: Graph's role, mapped.
    param([string] $Role)
    switch ($Role) {
        'freeBusyRead'                      { return @('AvailabilityOnly') }
        'limitedRead'                       { return @('LimitedDetails') }
        'read'                              { return @('Reviewer') }
        'write'                             { return @('Editor') }
        'delegateWithoutPrivateEventAccess' { return @('Editor') }
        'delegateWithPrivateEventAccess'    { return @('Editor') }
        'none'                              { return @('None') }
        default                             { return @('Reviewer') }
    }
}

function Wait-Until {
    # Polls a condition; a new mailbox takes minutes to be usable everywhere.
    param([scriptblock] $Condition, [string] $What, [int] $TimeoutSeconds = 900, [int] $IntervalSeconds = 20)
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $reported = $false
    while ($true) {
        $value = $null
        try { $value = & $Condition } catch { $value = $null }
        if ($value) { return $value }
        if ((Get-Date) -ge $deadline) { throw "Timed out after $TimeoutSeconds seconds waiting for $What." }
        if (-not $reported) { Write-Info "Waiting for $What..."; $reported = $true }
        Start-Sleep -Seconds $IntervalSeconds
    }
}

# ==============================================================================
#  Run
# ==============================================================================
$script:ConnectedExo = $false
try {
    # -- Exchange Online ---------------------------------------------------------
    if (-not (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
        throw "The ExchangeOnlineManagement module is required: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }
    try {
        $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
    } catch {
        $connectParams = @{ ShowBanner = $false }
        if ($TenantId) { $connectParams['Organization'] = $TenantId }
        Connect-ExchangeOnline @connectParams
        $script:ConnectedExo = $true
    }

    # The same tenant for Graph as for Exchange: explicit, GDAP customer, or the
    # tenant the Exchange session is connected to.
    $tenant = $TenantId
    if (-not $tenant) {
        try {
            $gdap = ($global:authMode -and ([string]$global:authMode).ToUpperInvariant() -eq 'GDAP') -or
                    ($env:M365_AUTH_MODE -and ([string]$env:M365_AUTH_MODE).ToUpperInvariant() -eq 'GDAP')
            if ($gdap -and $global:cid)          { $tenant = [string]$global:cid }
            elseif ($env:M365_CUSTOMER_TENANTID) { $tenant = [string]$env:M365_CUSTOMER_TENANTID }
        } catch {}
    }
    if (-not $tenant) {
        try { $tenant = [string](@(Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' })[0].TenantID) } catch {}
    }

    # -- Source mailbox ----------------------------------------------------------
    Write-Step '1. Source'
    $source = Get-EXOMailbox -Identity $Mailbox -Properties EmailAddresses, ExternalDirectoryObjectId, PrimarySmtpAddress, UserPrincipalName, DisplayName -ErrorAction Stop
    $script:SourceUpn = [string]$source.UserPrincipalName
    $script:SourceId  = [string]$source.ExternalDirectoryObjectId
    $script:SourceAddresses = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($a in @($source.EmailAddresses)) { if ([string]$a -match '^(?i)smtp:(.+)$') { [void]$script:SourceAddresses.Add($Matches[1]) } }
    [void]$script:SourceAddresses.Add([string]$source.PrimarySmtpAddress)
    [void]$script:SourceAddresses.Add($script:SourceUpn)
    Write-Ok "$($source.DisplayName) <$($source.PrimarySmtpAddress)>"

    $regional = $null
    try { $regional = Get-MailboxRegionalConfiguration -Identity $script:SourceUpn -ErrorAction Stop } catch { Write-Warn "Could not read the mailbox's language and time zone: $($_.Exception.Message)" }
    if ($regional -and $regional.TimeZone) {
        try { $script:MailboxTimeZone = [PSCustomObject]@{ Info = [TimeZoneInfo]::FindSystemTimeZoneById([string]$regional.TimeZone); Id = [string]$regional.TimeZone } } catch {}
    }
    Write-Info "Time zone: $($script:MailboxTimeZone.Id)$(if ($regional -and $regional.Language) { "; language: $($regional.Language)" })"

    # -- Graph -------------------------------------------------------------------
    if (-not (Connect-GraphForCalendar -Tenant $tenant)) {
        throw "No usable app-only Graph access - nothing was read or changed."
    }
    if ($script:SkipCategories) { Write-Warn "No MailboxSettings.ReadWrite - categories are copied by name, without their colour." }

    # -- The calendar ------------------------------------------------------------
    $calendars = Get-GraphPaged -Uri "https://graph.microsoft.com/v1.0/users/$($script:SourceId)/calendars?`$top=100"
    $cal = @($calendars | Where-Object { $_.name -eq $Calendar })
    if ($cal.Count -eq 0) {
        $names = @($calendars | ForEach-Object { "'$($_.name)'" }) -join ', '
        throw "No calendar '$Calendar' in $($script:SourceUpn). Calendars in the list: $names"
    }
    $own = @($cal | Where-Object { -not $_.owner.address -or $script:SourceAddresses.Contains([string]$_.owner.address) })
    if ($own.Count -eq 0) {
        $o = $cal[0].owner
        throw "'$Calendar' in $($script:SourceUpn)'s list belongs to $($o.name) <$($o.address)> - it is only mapped here. Run the script against that mailbox."
    }
    if ($own.Count -gt 1) { throw "$($script:SourceUpn) has $($own.Count) calendars called '$Calendar'. Rename one first." }
    $cal = $own[0]
    if ($cal.isDefaultCalendar) {
        throw "'$Calendar' is the main calendar of $($script:SourceUpn) and cannot be removed. If this whole mailbox is the shared calendar, convert it in place instead: Set-Mailbox $($script:SourceUpn) -Type $ResourceType"
    }
    $calId = ConvertTo-UrlId $cal.id
    Write-Ok "Calendar '$($cal.name)' found"

    # -- Items -------------------------------------------------------------------
    Write-Info "Reading items..."
    $sourceEvents = @(Get-GraphPaged -Uri "https://graph.microsoft.com/v1.0/users/$($script:SourceId)/calendars/$calId/events?`$top=100")
    $series  = @($sourceEvents | Where-Object { [string]$_.type -eq 'seriesMaster' })
    $singles = $sourceEvents.Count - $series.Count

    $script:SourceInstances = @{}
    $script:SeriesWindows   = @{}
    $exceptionCount = 0
    foreach ($m in $series) {
        $window = Get-SeriesWindow -Master $m
        $script:SeriesWindows[$m.id] = $window
        $instances = @(Get-GraphPaged -Uri ("https://graph.microsoft.com/v1.0/users/$($script:SourceId)/events/$(ConvertTo-UrlId $m.id)/instances" +
                                            "?startDateTime=$($window.Start)&endDateTime=$($window.End)&`$top=100"))
        $script:SourceInstances[$m.id] = $instances
        $exceptionCount += @($instances | Where-Object { [string]$_.type -eq 'exception' }).Count
    }
    $meetings    = @($sourceEvents | Where-Object { @($_.attendees).Count -gt 0 }).Count
    $withFiles   = @($sourceEvents | Where-Object { $_.hasAttachments }).Count
    $usedCats    = @(@($sourceEvents) + @($script:SourceInstances.Values | ForEach-Object { $_ }) |
                     ForEach-Object { $_.categories } | Where-Object { $_ } | Sort-Object -Unique)

    Write-Ok "$($sourceEvents.Count) item(s): $singles single, $($series.Count) series with $exceptionCount exception(s)"
    if ($meetings -gt 0)  { Write-Info "$meetings meeting(s) with attendees - nobody is invited, attendees are listed in the item body" }
    if ($withFiles -gt 0) { Write-Info "$withFiles item(s) with attachments" }
    if ($usedCats.Count -gt 0) { Write-Info "Categories in use: $($usedCats -join ', ')" }

    $sourceCategories = @()
    if (-not $script:SkipCategories -and $usedCats.Count -gt 0) {
        try { $sourceCategories = @(Get-GraphPaged -Uri "https://graph.microsoft.com/v1.0/users/$($script:SourceId)/outlook/masterCategories") }
        catch { Write-Warn "Could not read the category list: $(Get-GraphErrorText $_)" }
    }

    # -- Permissions -------------------------------------------------------------
    Write-Step '2. Permissions'
    $sourceFolder = $null
    try {
        $folders = @(Get-EXOMailboxFolderStatistics -Identity $script:SourceUpn -FolderScope Calendar -ErrorAction Stop |
                     Where-Object { $_.Name -eq $cal.name -and $_.FolderType -ne 'Calendar' })
        if ($folders.Count -eq 1) {
            # FolderPath uses "/" as separator and U+F8FF for a "/" inside a name.
            $path = ([string]$folders[0].FolderPath -replace '/', '\').Replace([string][char]0xF8FF, '/')
            $sourceFolder = "$($script:SourceUpn):$path"
        }
    } catch {}
    if (-not $sourceFolder) { Write-Warn "The Exchange folder could not be pinned down - rights are taken from Graph's roles instead of the exact access rights." }

    function Get-SourceRight {
        # Exact Exchange rights for one principal, or $null when Exchange cannot say.
        param([string] $User)
        if (-not $sourceFolder) { return $null }
        try { return (Get-MailboxFolderPermission -Identity $sourceFolder -User $User -ErrorAction Stop | Select-Object -First 1) } catch { return $null }
    }

    $plan = [System.Collections.Generic.List[object]]::new()
    $graphPermissions = @(Get-GraphPaged -Uri "https://graph.microsoft.com/v1.0/users/$($script:SourceId)/calendars/$calId/calendarPermissions")
    foreach ($g in $graphPermissions) {
        $address = [string]$g.emailAddress.address
        $name    = [string]$g.emailAddress.name
        if (-not $address) {
            if (-not $g.isRemovable) {
                $exo = Get-SourceRight -User 'Default'
                $rights = @(if ($exo) { $exo.AccessRights | ForEach-Object { [string]$_ } } else { Get-GraphRoleRight -Role ([string]$g.role) })
                $plan.Add([PSCustomObject]@{ Principal = 'Default (everyone in the organisation)'; User = 'Default'; Rights = $rights; Invite = $false; Note = '' })
            } else {
                $script:Notices.Add("Permission for '$name' belongs to an account that no longer exists - not carried over")
            }
            continue
        }

        $recipient = $null
        try { $recipient = Get-EXORecipient -Identity $address -ErrorAction Stop } catch {}
        if (-not $recipient) {
            $script:Notices.Add("$name <$address> is outside the tenant - share the new calendar with them by hand")
            continue
        }

        $exo    = Get-SourceRight -User $address
        # @() around the whole if: a one-element array assigned from an if-block is
        # unrolled to a plain string, and $rights[0] would then be its first letter.
        $rights = @(if ($exo) { $exo.AccessRights | ForEach-Object { [string]$_ } } else { Get-GraphRoleRight -Role ([string]$g.role) })
        if ($rights -contains 'None') { continue }
        $note = ''
        if ($exo -and @($exo.SharingPermissionFlags | Where-Object { "$_" -match 'Delegate' }).Count -gt 0) { $note = 'delegate flag dropped' }
        $isUser = [string]$recipient.RecipientTypeDetails -like '*Mailbox'
        $invite = $isUser -and $rights.Count -eq 1 -and $rights[0] -in @('AvailabilityOnly', 'LimitedDetails', 'Reviewer', 'Editor')
        $plan.Add([PSCustomObject]@{ Principal = "$name <$address>"; User = $address; Rights = $rights; Invite = $invite; Note = $note })
    }
    $anonymous = Get-SourceRight -User 'Anonymous'
    if ($anonymous -and @($anonymous.AccessRights | Where-Object { "$_" -ne 'None' }).Count -gt 0) {
        $plan.Add([PSCustomObject]@{ Principal = 'Anonymous (published calendar)'; User = 'Anonymous'; Rights = @($anonymous.AccessRights | ForEach-Object { [string]$_ }); Invite = $false; Note = 'the published link changes' })
    }
    if ($SourceOwnerRights -ne 'None') {
        $plan.Add([PSCustomObject]@{ Principal = "$($source.DisplayName) <$($source.PrimarySmtpAddress)>"; User = [string]$source.PrimarySmtpAddress; Rights = @($SourceOwnerRights); Invite = ($SourceOwnerRights -in @('Reviewer', 'Editor')); Note = 'owner of the original calendar' })
    }

    foreach ($p in $plan) {
        $line = "{0,-50} {1}" -f $p.Principal, ($p.Rights -join ', ')
        if ($p.Note) { $line += "  ($($p.Note))" }
        Write-Info $line
    }

    # -- Target ------------------------------------------------------------------
    Write-Step '3. Resource mailbox'
    if (-not $ResourceAddress) {
        $ResourceAddress = "$(ConvertTo-MailboxAlias -Name $ResourceName)@$(([string]$source.PrimarySmtpAddress).Split('@')[1])"
    }
    $existingTarget = $null
    try { $existingTarget = Get-EXORecipient -Identity $ResourceAddress -ErrorAction Stop } catch {}
    if ($existingTarget -and [string]$existingTarget.RecipientTypeDetails -notin @('RoomMailbox', 'EquipmentMailbox')) {
        throw "$ResourceAddress is already in use by a $($existingTarget.RecipientTypeDetails). Pick another -ResourceAddress."
    }
    if ($existingTarget) { Write-Ok "$ResourceType mailbox '$($existingTarget.DisplayName)' <$ResourceAddress> already exists - it is reused and existing copies are skipped" }
    else                 { Write-Ok "$ResourceType mailbox '$ResourceName' <$ResourceAddress> will be created" }

    if (-not $Apply) {
        Write-Host ""
        Write-Host "  Preview only - nothing was changed. Rerun with -Apply to:" -ForegroundColor Yellow
        Write-Info "- write a backup of every item and permission"
        Write-Info "- $(if ($existingTarget) { 'reuse' } else { 'create' }) the $ResourceType mailbox and set its rights$(if ($SendSharingInvitation) { ', with sharing invitations' })"
        Write-Info "- copy $($sourceEvents.Count) item(s) and verify every one"
        Write-Info "- $(if ($RemoveSourceCalendar) { 'then remove the original calendar' } else { 'keep the original (add -RemoveSourceCalendar to remove it)' })"
        Write-Host ""
        if ($PassThru) {
            [PSCustomObject]@{ Applied = $false; ResourceName = $ResourceName; ResourceAddress = $ResourceAddress; ResourceType = $ResourceType
                               Items = $sourceEvents.Count; Verified = $false; SourceRemoved = $false; BackupPath = $null }
        }
        return
    }

    # -- Backup ------------------------------------------------------------------
    Write-Step '4. Backup'
    if (-not $BackupPath) {
        $outputRoot = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
        $BackupPath = Join-Path $outputRoot ("CalendarConvert_{0}_{1}" -f (ConvertTo-MailboxAlias -Name $Calendar), (Get-Date -Format 'yyyyMMdd_HHmmss'))
    }
    $script:BackupDir = $BackupPath
    if (-not (Test-Path $BackupPath)) { New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null }
    $backupFile = Join-Path $BackupPath 'calendar-backup.json'
    [ordered]@{
        exportedAt  = (Get-Date).ToString('o')
        mailbox     = $script:SourceUpn
        calendar    = $cal
        events      = $sourceEvents
        instances   = $script:SourceInstances
        permissions = $graphPermissions
        plan        = $plan
        categories  = $sourceCategories
    } | ConvertTo-Json -Depth 30 | Set-Content -Path $backupFile -Encoding UTF8
    Write-Ok "Backup written: $backupFile"

    # -- Mailbox -----------------------------------------------------------------
    Write-Step '5. Resource mailbox'
    if (-not $existingTarget) {
        $newParams = @{
            Name               = $ResourceName
            DisplayName        = $ResourceName
            Alias              = ($ResourceAddress.Split('@')[0])
            PrimarySmtpAddress = $ResourceAddress
            ErrorAction        = 'Stop'
        }
        if ($ResourceType -eq 'Room') { $newParams['Room'] = $true } else { $newParams['Equipment'] = $true }
        New-Mailbox @newParams | Out-Null
        Write-Ok "$ResourceType mailbox created: $ResourceName <$ResourceAddress>"
    }

    $target = Wait-Until -What 'the mailbox to be provisioned' -Condition {
        $m = Get-EXOMailbox -Identity $ResourceAddress -Properties ExternalDirectoryObjectId -ErrorAction Stop
        if ($m.ExternalDirectoryObjectId) { $m }
    }
    $script:TargetId = [string]$target.ExternalDirectoryObjectId
    $null = Wait-Until -What 'the calendar to be reachable over Graph' -Condition {
        Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/users/$($script:TargetId)/calendar?`$select=id"
    }
    Write-Ok "Mailbox ready"

    # Same language and time zone as the source: the calendar folder gets the name
    # users know, and all-day series line up with the original.
    if ($regional) {
        try {
            $rc = @{ Identity = $ResourceAddress; ErrorAction = 'Stop' }
            if ($regional.TimeZone) { $rc['TimeZone'] = [string]$regional.TimeZone }
            if ($regional.Language) {
                $rc['Language'] = [string]$regional.Language
                $rc['DateFormat'] = $regional.DateFormat
                $rc['TimeFormat'] = $regional.TimeFormat
                $rc['LocalizeDefaultFolderName'] = $true
            }
            Set-MailboxRegionalConfiguration @rc
            Write-Ok "Language and time zone copied from the source mailbox"
        } catch {
            Write-Warn "Could not set language and time zone: $($_.Exception.Message)"
        }
    }

    # A shared calendar, not a meeting room: bookings are accepted, overlap is
    # fine, and nothing in an item is rewritten. 1080 days is the service maximum.
    $processingSet = $false
    for ($attempt = 1; $attempt -le 5 -and -not $processingSet; $attempt++) {
        try {
            Set-CalendarProcessing -Identity $ResourceAddress -AutomateProcessing AutoAccept -AllowConflicts $true `
                -BookingWindowInDays 1080 -EnforceSchedulingHorizon $false -MaximumDurationInMinutes 0 `
                -AddOrganizerToSubject $false -DeleteSubject $false -DeleteComments $false `
                -RemovePrivateProperty $false -DeleteAttachments $false -ErrorAction Stop
            $processingSet = $true
        } catch {
            if ($attempt -lt 5) { Write-Info "Calendar processing not accepted yet ($attempt/5) - retrying in 30 s"; Start-Sleep -Seconds 30 }
            else { Write-Warn "Calendar processing could not be set: $($_.Exception.Message)" }
        }
    }
    if ($processingSet) { Write-Ok "Calendar processing: auto-accept, overlap allowed" }

    # -- Rights ------------------------------------------------------------------
    Write-Step '6. Rights'
    $targetFolderName = Wait-Until -What 'the calendar folder' -Condition {
        (@(Get-EXOMailboxFolderStatistics -Identity $ResourceAddress -FolderScope Calendar -ErrorAction Stop | Where-Object { $_.FolderType -eq 'Calendar' })[0]).Name
    }
    $targetFolder = "$($ResourceAddress):\$targetFolderName"
    foreach ($p in $plan) {
        try {
            if ($p.User -in @('Default', 'Anonymous')) {
                Set-MailboxFolderPermission -Identity $targetFolder -User $p.User -AccessRights $p.Rights -ErrorAction Stop | Out-Null
                Write-Ok "$($p.Principal): $($p.Rights -join ', ')"
                continue
            }
            $current = $null
            try { $current = Get-MailboxFolderPermission -Identity $targetFolder -User $p.User -ErrorAction Stop } catch {}
            if ($current) {
                Set-MailboxFolderPermission -Identity $targetFolder -User $p.User -AccessRights $p.Rights -ErrorAction Stop | Out-Null
                Write-Ok "$($p.Principal): $($p.Rights -join ', ') (updated)"
                continue
            }
            $add = @{ Identity = $targetFolder; User = $p.User; AccessRights = $p.Rights; ErrorAction = 'Stop' }
            $invited = $false
            if ($SendSharingInvitation -and $p.Invite) {
                try {
                    Add-MailboxFolderPermission @add -SendNotificationToUser $true | Out-Null
                    $invited = $true
                } catch {
                    Write-Warn "No invitation for $($p.Principal): $($_.Exception.Message)"
                }
            }
            if (-not $invited) { Add-MailboxFolderPermission @add | Out-Null }
            Write-Ok "$($p.Principal): $($p.Rights -join ', ')$(if ($invited) { ' - invitation sent' })"
        } catch {
            $script:Failures.Add("Right for $($p.Principal): $($_.Exception.Message)")
            Write-Warn "$($p.Principal): $($_.Exception.Message)"
        }
    }

    # -- Categories --------------------------------------------------------------
    if (-not $script:SkipCategories -and $usedCats.Count -gt 0 -and $sourceCategories.Count -gt 0) {
        try {
            $targetCategories = @(Get-GraphPaged -Uri "https://graph.microsoft.com/v1.0/users/$($script:TargetId)/outlook/masterCategories")
            foreach ($c in $sourceCategories) {
                if ($usedCats -notcontains [string]$c.displayName) { continue }
                if (@($targetCategories | Where-Object { $_.displayName -eq $c.displayName }).Count -gt 0) { continue }
                Invoke-Graph -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($script:TargetId)/outlook/masterCategories" `
                    -Body @{ displayName = [string]$c.displayName; color = [string]$c.color } | Out-Null
            }
            Write-Ok "Categories with colour: $($usedCats -join ', ')"
        } catch {
            Write-Warn "Categories could not be created: $(Get-GraphErrorText $_)"
        }
    }

    # -- Items -------------------------------------------------------------------
    Write-Step '7. Items'
    $existing = Get-ExistingCopy
    if ($existing.Count -gt 0) { Write-Info "$($existing.Count) copy/copies from an earlier run found" }
    $total = $sourceEvents.Count
    $n = 0
    foreach ($ev in $sourceEvents) {
        $n++
        try {
            Copy-CalendarEvent -Item $ev -Existing $existing
        } catch {
            $script:Stats.Failed++
            $script:Failures.Add("'$($ev.subject)' ($(Get-OccurrenceKey -Instance $ev)): $(Get-GraphErrorText $_)")
        }
        Write-Progress -Activity 'Copying items' -Status "$n / $total" -PercentComplete ([int](100 * $n / [Math]::Max(1, $total)))
        if ($n % 25 -eq 0) { Write-Info ("[{0}] {1} / {2}" -f (Get-Date -Format 'HH:mm:ss'), $n, $total) }
    }
    Write-Progress -Activity 'Copying items' -Completed
    Write-Ok ("{0} copied, {1} already there{2}, {3} failed" -f $script:Stats.Copied, $script:Stats.AlreadyThere,
              $(if ($script:Stats.Redone) { ", $($script:Stats.Redone) redone" } else { '' }), $script:Stats.Failed)
    if ($series.Count -gt 0) { Write-Info "Series: $($script:Stats.Exceptions) exception(s) applied, $($script:Stats.Cancelled) occurrence(s) cancelled" }
    if ($withFiles -gt 0)    { Write-Info "Attachments: $($script:Stats.Attachments) copied, $($script:Stats.AttachmentsSaved) saved to the backup folder" }

    # -- Verify ------------------------------------------------------------------
    Write-Step '8. Verification'
    $after   = Get-ExistingCopy
    $missing = @($sourceEvents | Where-Object { -not ($after[$_.id] -and $after[$_.id].Complete) })
    $clean   = ($missing.Count -eq 0) -and ($script:Failures.Count -eq 0)
    if ($missing.Count -eq 0) { Write-Ok "All $total item(s) have a complete copy" }
    else {
        Write-Warn "$($missing.Count) item(s) without a complete copy:"
        foreach ($m in ($missing | Select-Object -First 10)) { Write-Info "- $($m.subject)" }
    }

    # -- Remove the original -----------------------------------------------------
    $removed = $false
    Write-Step '9. Original calendar'
    if (-not $RemoveSourceCalendar) {
        Write-Ok "Kept - add -RemoveSourceCalendar once users have switched"
    } elseif (-not $clean) {
        Write-Warn "NOT removed - the copy is incomplete. Fix the items above and run the same command again."
    } else {
        # Removal needs a typed confirmation or an explicit -Force. Without a
        # person at the keyboard there is nobody to confirm, so no removal.
        $go = [bool]$Force
        if (-not $Force) {
            if (Test-CanPrompt) {
                Write-Host ""
                $answer = Read-Host "  Type the calendar name '$($cal.name)' to remove it from $($script:SourceUpn)"
                $go = ($answer -ceq [string]$cal.name)
                if (-not $go) { Write-Warn "Not confirmed - the original calendar stays." }
            } else {
                Write-Warn "NOT removed - a non-interactive session cannot confirm it. Add -Force to remove it unattended."
            }
        }
        if ($go) {
            Invoke-Graph -Method DELETE -Uri "https://graph.microsoft.com/v1.0/users/$($script:SourceId)/calendars/$calId" | Out-Null
            $removed = $true
            Write-Ok "Calendar '$($cal.name)' removed from $($script:SourceUpn) - the backup is in $BackupPath"
        }
    }

    # -- Summary -----------------------------------------------------------------
    if ($script:Notices.Count -gt 0 -or $script:Failures.Count -gt 0) {
        Write-Step 'Needs attention'
        foreach ($f in $script:Failures) { Write-Warn $f }
        foreach ($note in $script:Notices) { Write-Info "- $note" }
    }
    Write-Step 'Next'
    $people = @($plan | Where-Object { $_.User -notin @('Default', 'Anonymous') } | ForEach-Object { $_.Principal })
    if ($SendSharingInvitation) { Write-Info "Users accept the sharing invitation, or add the calendar by hand:" }
    else                        { Write-Info "Users add the new calendar in Outlook:" }
    Write-Info "  Calendar > Add calendar > From directory > '$ResourceName'"
    if ($people.Count -gt 0) { Write-Info "With rights: $($people -join '; ')" }
    Write-Info "Who still has the old calendar in their list:"
    Write-Info "  .\Get-CalendarMappings.ps1 -Search '$($cal.name)'"
    Write-Host ""
    if ($PassThru) {
        [PSCustomObject]@{ Applied = $true; ResourceName = $ResourceName; ResourceAddress = $ResourceAddress; ResourceType = $ResourceType
                           Items = $total; Verified = $clean; SourceRemoved = $removed; BackupPath = $BackupPath }
    }
} finally {
    Remove-TempApp
    if ($script:ConnectedExo) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
}
