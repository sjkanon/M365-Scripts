#Requires -Version 7.0
#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Applications, Microsoft.Graph.Authentication, Microsoft.Graph.Calendar, Microsoft.Graph.Groups

<#
.SYNOPSIS
    Migrates an M365 Group calendar to a Room or Shared Mailbox.
    Writes with a temporary App Registration that is removed again at the end,
    unless your own app is given.

.DESCRIPTION
    Author : Sjoerd Kanon
    Date   : 19/03/2026

    Sign-in
    -------
    Two kinds of access are needed, and Microsoft decides which:

      Reading   The group calendar can only be read DELEGATED - Graph refuses
                group calendar reads for application permissions
                (https://learn.microsoft.com/en-us/graph/known-issues). So you sign
                in as yourself, and you must be a member of the group.
      Writing   The events go into the new room/shared mailbox, which is not yours.
                A delegated token only reaches it with explicit rights on it, so
                writing is APP-ONLY (Calendars.ReadWrite application permission).

    Delegated sign-in goes through scripts\Startup\Connect-M365.ps1: device code
    when load.config.ps1 sets useDeviceCodeAuth, the GDAP customer when authMode
    is GDAP. Exchange Online connects the same way.

    For writing, by default (MODE A) the script creates a short-lived App
    Registration with that delegated session (Application.ReadWrite.All +
    AppRoleAssignment.ReadWrite.All - Global Administrator or Privileged Role
    Administrator), grants it Calendars.ReadWrite (and Group.ReadWrite.All with
    -DeleteSourceGroup), takes an app-only token with a 2-hour secret that is
    never shown, and deletes the app again when the run ends - also when it fails.

    MODE B - your own app: -ClientId with -ClientSecret or -CertificateThumbprint,
    or -AppOnly to take ClientId and CertificateThumbprint from graph.appid.json.
    It needs Calendars.ReadWrite (and Group.ReadWrite.All for -DeleteSourceGroup)
    as application permissions with admin consent.

    Steps:
    1.  Connect Graph (delegated) and Exchange Online
    2.  Find the source M365 Group and read its events (delegated)
    3.  Temporary App Registration (Mode A) or your own app (Mode B) for writing
    4.  Create destination mailbox (Room or Shared)
    5.  Set calendar permissions + configure AutoAccept
    6.  Copy events to the destination mailbox (app-only)
    7.  (Optional) Delete source M365 Group (app-only)
    8.  Summary, remove the temporary app, disconnect what this script connected

    Why Room Mailbox instead of M365 Group?
    - No email notifications to the entire company on new events
    - Works like a meeting room: add as attendee = event appears on shared calendar
    - Everyone can read the calendar without being a member
    - AutoAccept so leave is automatically approved

    Modules: installed by scripts\Startup\Install-Modules.ps1.

.PARAMETER TenantId
    Tenant ID or domain. Defaults to the GDAP customer when load.config.ps1 sets
    authMode GDAP; otherwise the tenant you sign in to.

.PARAMETER AdminUPN
    UPN of the admin running the script. Must be a member of the source M365 Group.
    Optional: only used to warn when you signed in with another account.

.PARAMETER ClientId
    AppId of your own App Registration for writing (Mode B). Leave empty for a
    temporary one.

.PARAMETER ClientSecret
    Client secret for -ClientId.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for -ClientId.

.PARAMETER AppOnly
    Mode B with ClientId and CertificateThumbprint for the tenant from
    graph.appid.json in the repo root.

.PARAMETER AppName
    Name prefix of the temporary App Registration. Default: CalendarMigration-Temp

.PARAMETER SourceGroupMail
    Email address of the source M365 Group.

.PARAMETER SourceGroupDisplayName
    DisplayName of the source M365 Group (fallback if mail lookup fails).

.PARAMETER DestinationType
    Type of destination mailbox: Room (recommended) or Shared.

.PARAMETER DestinationDisplayName
    Display name for the destination mailbox.

.PARAMETER DestinationAlias
    Alias for the destination mailbox (must be unique in the tenant).

.PARAMETER DestinationEmail
    Primary SMTP address for the destination mailbox.

.PARAMETER DaysBack
    How many days back to retrieve events. Default: 365.

.PARAMETER DaysForward
    How many days forward to retrieve events. Default: 730.

.PARAMETER DeleteSourceGroup
    If $true, deletes the M365 Group after migration. Default: $false.

.EXAMPLE
    # Fully automatic (Mode A): temporary app, removed again at the end
    .\Migrate-Calendar.ps1 -TenantId "contoso.onmicrosoft.com" -AdminUPN "admin@contoso.com" -SourceGroupMail "holidays@contoso.com"

.EXAMPLE
    # Your own App Registration (Mode B), certificate from graph.appid.json
    .\Migrate-Calendar.ps1 -TenantId "contoso.onmicrosoft.com" -AppOnly -SourceGroupMail "holidays@contoso.com"

.EXAMPLE
    # Dry run
    .\Migrate-Calendar.ps1 -TenantId "contoso.onmicrosoft.com" -AdminUPN "admin@contoso.com" -SourceGroupMail "holidays@contoso.com" -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$TenantId,

    [string]$AdminUPN,

    # Your own App Registration for writing - leave empty for a temporary one
    [string]$ClientId              = "",
    [string]$ClientSecret          = "",
    [string]$CertificateThumbprint = "",
    [switch]$AppOnly,
    [string]$AppName               = "CalendarMigration-Temp",

    # Source M365 Group
    [string]$SourceGroupMail        = "holidays@domain.com",
    [string]$SourceGroupDisplayName = "Holidays",

    # Destination mailbox
    # "Room"   = Resource/Room Mailbox (recommended — works like a meeting room, no notifications)
    # "Shared" = Shared Mailbox (users add the calendar manually, no booking system)
    [ValidateSet("Room", "Shared")]
    [string]$DestinationType        = "Room",
    [string]$DestinationDisplayName = "Holidays Calendar",
    [string]$DestinationAlias       = "holidays-calendar",
    [string]$DestinationEmail       = "holidays-calendar@domain.com",

    # Migration time window
    [int]$DaysBack    = 365,
    [int]$DaysForward = 730,

    # Delete source M365 Group after migration
    [bool]$DeleteSourceGroup = $false
)

. (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')

#region Helpers

function Write-Step { param([string]$m) Write-Host "`n==> $m" -ForegroundColor Cyan }
function Write-OK   { param([string]$m) Write-Host "    [OK]   $m" -ForegroundColor Green }
function Write-Warn { param([string]$m) Write-Host "    [WARN] $m" -ForegroundColor Yellow }
function Write-Fail { param([string]$m) Write-Host "    [FAIL] $m" -ForegroundColor Red }

$script:Graph        = $null    # Connect-M365Graph results; only what this script opened is disconnected
$script:GraphApp     = $null
$script:Exo          = $null
$script:TempAppId    = $null    # object id of the temporary App Registration
$script:WriteHeaders = $null    # app-only bearer token (Mode A, or Mode B with a secret)

function Get-JwtRoles {
    param([string]$Jwt)
    try {
        $payload = $Jwt.Split('.')[1].Replace('-', '+').Replace('_', '/')
        switch ($payload.Length % 4) { 2 { $payload += '==' } 3 { $payload += '=' } }
        return @(([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($payload)) | ConvertFrom-Json).roles)
    } catch { return @() }
}

function Get-AppToken {
    # client_credentials over REST: the delegated Graph session stays as it is,
    # so it can remove the temporary app at the end.
    param([string]$Tenant, [string]$App, [string]$Secret)
    (Invoke-RestMethod -Method POST -ErrorAction Stop -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" -Body @{
        grant_type = 'client_credentials'; scope = 'https://graph.microsoft.com/.default'; client_id = $App; client_secret = $Secret
    }).access_token
}

function Invoke-AppGraph {
    # Every write goes through here: the app-only bearer token when there is one,
    # otherwise the app-only Graph SDK session (Mode B with a certificate).
    param([string]$Method, [string]$Uri, $Body)
    $p = @{ Method = $Method; Uri = $Uri; ErrorAction = 'Stop' }
    if ($Body) { $p['Body'] = ($Body | ConvertTo-Json -Depth 6); $p['ContentType'] = 'application/json' }
    if ($script:WriteHeaders) { return Invoke-RestMethod @p -Headers $script:WriteHeaders }
    return Invoke-MgGraphRequest @p
}

function Remove-TempApp {
    # Needs the delegated session (Application.ReadWrite.All).
    if (-not $script:TempAppId) { return }
    try {
        Remove-MgApplication -ApplicationId $script:TempAppId -ErrorAction Stop
        Write-OK "Temporary App Registration removed"
    } catch {
        Write-Warn "Could not remove the temporary App Registration (object ID $($script:TempAppId)). Remove it by hand in Entra ID > App registrations."
    }
    $script:TempAppId = $null
}

function Close-Connections {
    Remove-TempApp
    Disconnect-M365Graph $script:GraphApp
    Disconnect-M365Graph $script:Graph
    Disconnect-M365Exchange $script:Exo
}

#endregion

$tenant = Resolve-M365TenantId -TenantId $TenantId
if ($AppOnly -and -not $ClientId) {
    $reg = Get-M365AppRegistration -TenantId $tenant
    $ClientId = $reg.ClientId
    $CertificateThumbprint = $reg.CertificateThumbprint
    if (-not $tenant) { $tenant = $reg.Tenant }
}
if ($ClientId -and -not $ClientSecret -and -not $CertificateThumbprint) {
    throw "-ClientId needs -ClientSecret or -CertificateThumbprint."
}
$modeA = -not $ClientId

try {

#region Step 1: Connect Graph (delegated) and Exchange Online

# Graph first: the Graph SDK and Exchange Online each bundle their own MSAL, and
# the Graph SDK is the one that breaks when Exchange loaded first.
Write-Step "Connecting to Microsoft Graph (delegated)"
$scopes = @("Group.Read.All", "Calendars.Read")
if ($modeA) { $scopes += @("Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All") }
$script:Graph = Connect-M365Graph -Scopes $scopes -TenantId $tenant
$ctx = Get-MgContext
if (-not $tenant) { $tenant = $ctx.TenantId }
Write-OK "Graph connected (delegated) as $($ctx.Account)"
if ($AdminUPN -and $ctx.Account -and $ctx.Account -ne $AdminUPN) {
    Write-Warn "Signed in as $($ctx.Account), not $AdminUPN - that account must be a member of the group."
}

Write-Step "Connecting to Exchange Online"
$script:Exo = Connect-M365Exchange -TenantId $TenantId
Write-OK "Exchange Online connected"

#endregion

#region Step 2: Find M365 Group and retrieve events (delegated)

Write-Step "Looking up source M365 Group"

$group = $null

$mailLower = $SourceGroupMail.ToLower()
$group     = Get-MgGroup -Filter "mail eq '$mailLower'" -ErrorAction SilentlyContinue

if (-not $group) {
    Write-Warn "Not found on '$mailLower', trying '$SourceGroupMail'..."
    $group = Get-MgGroup -Filter "mail eq '$SourceGroupMail'" -ErrorAction SilentlyContinue
}

if (-not $group) {
    Write-Warn "Not found by mail, falling back to displayName '$SourceGroupDisplayName'..."
    $group = Get-MgGroup -Filter "displayName eq '$SourceGroupDisplayName'" -ErrorAction SilentlyContinue
}

if (-not $group) {
    Write-Warn "Not found via filter, last attempt via Search..."
    $group = Get-MgGroup -Search "`"displayName:$SourceGroupDisplayName`"" `
        -ConsistencyLevel eventual -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -eq $SourceGroupDisplayName -or $_.Mail -like "*holiday*" } |
        Select-Object -First 1
}

if (-not $group) {
    Write-Fail "Group not found after 4 attempts."
    Write-Host "    Tip: make sure the admin account is a member of the group." -ForegroundColor Yellow
    Write-Host "    Manual lookup: Get-MgGroup -Search `"`"displayName:Holidays`"`" -ConsistencyLevel eventual | Select DisplayName,Mail,Id" -ForegroundColor Yellow
    return
}

Write-OK "Group found: '$($group.DisplayName)' | Mail: $($group.Mail) | ID: $($group.Id)"

Write-Step "Retrieving events ($DaysBack days back to $DaysForward days forward)"

$startDate = (Get-Date).AddDays(-$DaysBack).ToString("yyyy-MM-ddT00:00:00")
$endDate   = (Get-Date).AddDays($DaysForward).ToString("yyyy-MM-ddT23:59:59")

$calEvents = Get-MgGroupCalendarEvent `
    -GroupId     $group.Id `
    -Filter      "start/dateTime ge '$startDate' and start/dateTime le '$endDate'" `
    -All `
    -ErrorAction Stop
Write-OK "$($calEvents.Count) events retrieved"

#endregion

#region Step 3: App-only access for writing

$writeRoles = @("Calendars.ReadWrite")
if ($DeleteSourceGroup) { $writeRoles += "Group.ReadWrite.All" }

if ($modeA) {
    Write-Step "Creating a temporary App Registration for writing (Mode A)"
    if ($PSCmdlet.ShouldProcess($AppName, "Create temporary App Registration ($($writeRoles -join ', '))")) {
        $name = "$AppName-$(Get-Date -Format 'yyyyMMddHHmmss')"
        $app  = New-MgApplication -DisplayName $name -SignInAudience AzureADMyOrg -ErrorAction Stop
        $script:TempAppId = $app.Id
        Write-OK "App created: $name | AppId: $($app.AppId)"

        $appSp   = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop
        $graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ErrorAction Stop
        foreach ($perm in $writeRoles) {
            $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $perm -and $_.AllowedMemberTypes -contains 'Application' }
            if (-not $role) { throw "Application permission '$perm' not found." }
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $appSp.Id -PrincipalId $appSp.Id `
                -ResourceId $graphSp.Id -AppRoleId $role.Id -ErrorAction Stop | Out-Null
            Write-OK "Granted (application): $perm"
        }

        # Short-lived and never printed: the app is deleted at the end of the run.
        $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
            DisplayName = 'temp'
            EndDateTime = (Get-Date).AddHours(2)
        } -ErrorAction Stop

        # A new app and its role assignments take a while to reach the token service.
        $deadline = (Get-Date).AddMinutes(4)
        while ($true) {
            $missing = $writeRoles
            try {
                $tok = Get-AppToken -Tenant $tenant -App $app.AppId -Secret $secret.SecretText
                $missing = @($writeRoles | Where-Object { (Get-JwtRoles $tok) -notcontains $_ })
            } catch { }
            if ($missing.Count -eq 0) { break }
            if ((Get-Date) -ge $deadline) { throw "The temporary app's token still lacks $($missing -join ', ') after 4 minutes." }
            Write-Host "    Waiting for the app and its permissions to propagate..." -ForegroundColor DarkGray
            Start-Sleep -Seconds 10
        }
        $script:WriteHeaders = @{ Authorization = "Bearer $tok" }
        Write-OK "App-only token obtained (temporary app)"
    }
} elseif ($ClientSecret) {
    Write-Step "Using your own App Registration (Mode B, secret)"
    $tok = Get-AppToken -Tenant $tenant -App $ClientId -Secret $ClientSecret
    $missing = @($writeRoles | Where-Object { (Get-JwtRoles $tok) -notcontains $_ })
    if ($missing.Count -gt 0) { Write-Warn "App $ClientId has no $($missing -join ', ') application permission - writes will fail with 403." }
    $script:WriteHeaders = @{ Authorization = "Bearer $tok" }
    Write-OK "App-only token obtained (ClientId $ClientId)"
} else {
    # A certificate goes through the Graph SDK; this replaces the delegated Graph
    # session in this process, which is no longer needed now the events are read.
    Write-Step "Using your own App Registration (Mode B, certificate)"
    $script:GraphApp = Connect-M365Graph -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -TenantId $tenant
    Write-OK "Graph connected (app-only) | ClientId $ClientId"
}

#endregion

#region Step 4: Create destination mailbox

Write-Step "Checking / creating destination mailbox: $DestinationEmail"

$existingMailbox = Get-Mailbox -Identity $DestinationEmail -ErrorAction SilentlyContinue

if ($existingMailbox) {
    Write-Warn "Mailbox '$DestinationEmail' already exists — step skipped"
} else {
    if ($PSCmdlet.ShouldProcess($DestinationEmail, "Create $DestinationType mailbox")) {
        if ($DestinationType -eq "Room") {
            New-Mailbox `
                -Name               $DestinationDisplayName `
                -Alias              $DestinationAlias `
                -PrimarySmtpAddress $DestinationEmail `
                -Room | Out-Null
            Write-OK "Room Mailbox created: $DestinationDisplayName ($DestinationEmail)"
        } else {
            New-Mailbox `
                -Name               $DestinationDisplayName `
                -Alias              $DestinationAlias `
                -PrimarySmtpAddress $DestinationEmail `
                -Shared | Out-Null
            Write-OK "Shared Mailbox created: $DestinationDisplayName ($DestinationEmail)"
        }
        Write-Host "    Waiting 60s for Exchange provisioning..." -ForegroundColor DarkGray
        Start-Sleep -Seconds 60
    }
}

#endregion

#region Step 5: Calendar permissions and AutoAccept

Write-Step "Setting calendar permissions on $DestinationEmail"

if ($PSCmdlet.ShouldProcess($DestinationEmail, "Set calendar permissions")) {
    Set-MailboxFolderPermission `
        -Identity     "$($DestinationEmail):\Calendar" `
        -User         Default `
        -AccessRights Reviewer `
        -ErrorAction  SilentlyContinue
    Write-OK "Default: Reviewer (read with details)"

    if ($DestinationType -eq "Room") {
        # Retry loop: newly created mailboxes can take a while to become ready
        $calProcSet = $false
        for ($attempt = 1; $attempt -le 5; $attempt++) {
            try {
                Set-CalendarProcessing `
                    -Identity                 $DestinationEmail `
                    -AutomateProcessing       AutoAccept `
                    -AllowConflicts           $true `
                    -AddOrganizerToSubject    $false `
                    -DeleteComments           $false `
                    -DeleteSubject            $false `
                    -BookingWindowInDays      1825 `
                    -EnforceSchedulingHorizon $false `
                    -MaximumDurationInMinutes 0 `
                    -ErrorAction Stop
                $calProcSet = $true
                break
            } catch {
                Write-Warn "Set-CalendarProcessing attempt $attempt/5 failed: $_"
                if ($attempt -lt 5) {
                    Write-Host "    Retrying in 30s..." -ForegroundColor DarkGray
                    Start-Sleep -Seconds 30
                }
            }
        }

        if ($calProcSet) {
            Write-OK "AutoAccept configured (overlapping bookings allowed, no duration limit)"
        } else {
            Write-Fail "Set-CalendarProcessing failed after 5 attempts — run manually after provisioning:"
            Write-Host "    Set-CalendarProcessing -Identity '$DestinationEmail' -AutomateProcessing AutoAccept -AllowConflicts `$true -BookingWindowInDays 1825 -EnforceSchedulingHorizon `$false -MaximumDurationInMinutes 0" -ForegroundColor Yellow
        }
    } else {
        Write-OK "Shared Mailbox: AutoAccept not applicable"
        Write-Host "    Users add the calendar manually via: Add calendar > Add from directory" -ForegroundColor DarkGray
    }
}

#endregion

#region Step 6: Copy events (app-only)

Write-Step "Copying events to destination mailbox calendar"

$successCount = 0
$failCount    = 0
$skippedCount = 0
$destUri      = "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($DestinationEmail))/events"

foreach ($calEvent in $calEvents) {

    if ($calEvent.IsCancelled) {
        $skippedCount++
        continue
    }

    if ($PSCmdlet.ShouldProcess($calEvent.Subject, "Copy event to $DestinationEmail")) {
        try {
            $tz = if ($calEvent.Start.TimeZone) { $calEvent.Start.TimeZone } else { "UTC" }

            $params = @{
                subject  = $calEvent.Subject
                isAllDay = [bool]$calEvent.IsAllDay
                showAs   = "oof"
                start    = @{ dateTime = $calEvent.Start.DateTime; timeZone = $tz }
                end      = @{ dateTime = $calEvent.End.DateTime;   timeZone = $tz }
                body     = @{
                    contentType = "text"
                    content     = "Migrated from M365 Group calendar. Original organizer: $($calEvent.Organizer.EmailAddress.Address)"
                }
            }

            Invoke-AppGraph -Method POST -Uri $destUri -Body $params | Out-Null
            Write-Host "    [+] $($calEvent.Subject) | $($calEvent.Start.DateTime)" -ForegroundColor DarkGreen
            $successCount++
        } catch {
            Write-Warn "Not copied: '$($calEvent.Subject)' — $_"
            $failCount++
        }
    }
}

Write-OK "Copy complete: $successCount OK | $skippedCount skipped (cancelled) | $failCount failed"

#endregion

#region Step 7: Delete source M365 Group (optional)

if ($DeleteSourceGroup) {
    Write-Step "Deleting source M365 Group: $SourceGroupDisplayName"
    if ($PSCmdlet.ShouldProcess($group.Id, "Delete M365 Group")) {
        try {
            Invoke-AppGraph -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$($group.Id)" | Out-Null
            Write-OK "M365 Group deleted"
        } catch {
            Write-Warn "Could not delete group: $_"
        }
    }
} else {
    Write-Warn "M365 Group was NOT deleted (DeleteSourceGroup = false)"
    Write-Host "    Recommendation: remove the group or disable membership notifications." -ForegroundColor Yellow
    Write-Host "    To delete manually: Remove-MgGroup -GroupId $($group.Id)" -ForegroundColor DarkYellow
}

#endregion

#region Step 8: Summary

$line = "=" * 65
Write-Host "`n$line" -ForegroundColor Cyan
Write-Host " SUMMARY  —  Calendar Migration" -ForegroundColor Cyan
Write-Host $line -ForegroundColor Cyan
Write-Host "  App Registration   : $(if ($modeA) { 'temporary (removed at the end)' } else { $ClientId })"
Write-Host "  Source calendar    : $SourceGroupMail (M365 Group)"
Write-Host "  Destination type   : $DestinationType Mailbox"
Write-Host "  Destination mailbox: $DestinationEmail"
Write-Host "  Events copied      : $successCount"
Write-Host "  Skipped            : $skippedCount (cancelled)"
Write-Host "  Failed             : $failCount"
Write-Host ""
Write-Host " ADD CALENDAR IN OUTLOOK (once per user):" -ForegroundColor Yellow
Write-Host "  1. Calendar > Add calendar > Add from directory"
Write-Host "  2. Search: '$DestinationDisplayName' or '$DestinationEmail'"
Write-Host "  3. Add — calendar appears under People's calendars"
Write-Host ""
if ($DestinationType -eq "Room") {
    Write-Host " HOW TO BOOK LEAVE — Room Mailbox:" -ForegroundColor Yellow
    Write-Host "  1. Create an appointment in Outlook (All day, status = Out of office)"
    Write-Host "  2. Add '$DestinationEmail' as an attendee (like a meeting room)"
    Write-Host "  3. Save — booking is automatically approved"
    Write-Host "  4. Event appears on the shared calendar for everyone"
} else {
    Write-Host " HOW TO BOOK LEAVE — Shared Mailbox:" -ForegroundColor Yellow
    Write-Host "  1. Create an appointment in Outlook (All day, status = Out of office)"
    Write-Host "  2. Save directly to the '$DestinationDisplayName' shared calendar"
    Write-Host "     (click the calendar icon next to your name and select '$DestinationDisplayName')"
    Write-Host "  3. Event appears on the shared calendar for everyone"
}
Write-Host $line -ForegroundColor Cyan

#endregion

} finally {
    # The temporary app never outlives the run; only what this script connected is closed.
    Close-Connections
}
