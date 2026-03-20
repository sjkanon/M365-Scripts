#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Applications, Microsoft.Graph.Authentication, Microsoft.Graph.Calendar, Microsoft.Graph.Groups, Microsoft.Graph.Users

<#
.SYNOPSIS
    Migrates an M365 Group calendar to a Room or Shared Mailbox.
    Automatically creates an App Registration if no ClientId/ClientSecret is provided.

.DESCRIPTION
    Author : Sjoerd Kanon
    Date   : 19/03/2026

    The script operates in two modes:

    MODE A — Fully automatic (recommended, first run):
        Do not provide ClientId/ClientSecret. The script creates an App Registration
        with the required application permissions and uses it immediately.

    MODE B — Existing App Registration:
        Provide ClientId and ClientSecret. The script skips the setup step.

    Steps:
    1.  Detect platform (Windows / macOS / Linux)
    2.  Connect to Exchange Online
    3.  Connect to Graph (delegated) for App Registration setup
    4.  Create App Registration + grant admin consent (Mode A only)
    5.  Reconnect Graph with application permissions (client credentials)
    6.  Create destination mailbox (Room or Shared)
    7.  Set calendar permissions + configure AutoAccept
    8.  Reconnect Graph as delegated user (required to read group calendar)
    9.  Find source M365 Group + retrieve events
    10. Switch back to app auth for writing
    11. Copy events to destination mailbox
    12. (Optional) Delete source M365 Group
    13. Summary

    Why Room Mailbox instead of M365 Group?
    - No email notifications to the entire company on new events
    - Works like a meeting room: add as attendee = event appears on shared calendar
    - Everyone can read the calendar without being a member
    - AutoAccept so leave is automatically approved

    Install modules if needed:
        Install-Module ExchangeOnlineManagement       -Scope CurrentUser
        Install-Module Microsoft.Graph.Applications   -Scope CurrentUser
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
        Install-Module Microsoft.Graph.Calendar       -Scope CurrentUser
        Install-Module Microsoft.Graph.Groups         -Scope CurrentUser
        Install-Module Microsoft.Graph.Users          -Scope CurrentUser

.PARAMETER TenantId
    Azure AD Tenant ID (Entra ID > Overview > Tenant ID).

.PARAMETER AdminUPN
    UPN of the admin running the script. Must be a member of the source M365 Group.

.PARAMETER ClientId
    AppId of an existing App Registration. Leave empty to create one automatically.

.PARAMETER ClientSecret
    Client Secret of the App Registration. Leave empty to create one automatically.

.PARAMETER AppName
    Name for the App Registration. Default: HolidaysCalendarMigration

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
    # First run — fully automatic (Mode A)
    .\Migrate-Calendar.ps1 -TenantId "xxxx" -AdminUPN "admin@domain.com" -SourceGroupMail "holidays@domain.com"

.EXAMPLE
    # Existing App Registration (Mode B)
    .\Migrate-Calendar.ps1 -TenantId "xxxx" -AdminUPN "admin@domain.com" -ClientId "yyyy" -ClientSecret "zzzz" -SourceGroupMail "holidays@domain.com"

.EXAMPLE
    # Dry run
    .\Migrate-Calendar.ps1 -TenantId "xxxx" -AdminUPN "admin@domain.com" -SourceGroupMail "holidays@domain.com" -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$AdminUPN,

    # App Registration — leave empty to create automatically
    [string]$ClientId     = "",
    [string]$ClientSecret = "",
    [string]$AppName      = "HolidaysCalendarMigration",

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

#region Helpers

function Write-Step { param([string]$m) Write-Host "`n==> $m" -ForegroundColor Cyan }
function Write-OK   { param([string]$m) Write-Host "    [OK]   $m" -ForegroundColor Green }
function Write-Warn { param([string]$m) Write-Host "    [WARN] $m" -ForegroundColor Yellow }
function Write-Fail { param([string]$m) Write-Host "    [FAIL] $m" -ForegroundColor Red }

function Connect-GraphDelegated {
    param([string]$TenantId, [string[]]$Scopes)
    if ($runOnWindows) {
        Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -NoWelcome
    } else {
        Write-Host "    Device code flow: copy the code and open the URL in your browser." -ForegroundColor DarkGray
        Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -UseDeviceAuthentication -NoWelcome
    }
}

function Connect-GraphAppAuth {
    param([string]$TenantId, [string]$ClientId, [string]$ClientSecret)
    $secure     = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($ClientId, $secure)

    try {
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome
        return
    } catch {
        Write-Warn "ClientSecretCredential failed, falling back to environment variables..."
    }

    # Fallback via environment variables (compatible with all SDK versions)
    $env:AZURE_CLIENT_ID     = $ClientId
    $env:AZURE_CLIENT_SECRET = $ClientSecret
    $env:AZURE_TENANT_ID     = $TenantId
    Connect-MgGraph -EnvironmentVariable -NoWelcome

    Remove-Item Env:AZURE_CLIENT_ID     -ErrorAction SilentlyContinue
    Remove-Item Env:AZURE_CLIENT_SECRET -ErrorAction SilentlyContinue
    Remove-Item Env:AZURE_TENANT_ID     -ErrorAction SilentlyContinue
}

#endregion

#region Step 1: Detect platform

Write-Step "Detecting platform"

$runOnMacOS   = ($IsMacOS -eq $true)
$runOnLinux   = ($IsLinux -eq $true)
$runOnWindows = ($IsWindows -eq $true) -or ($PSVersionTable.PSVersion.Major -le 5)

if (-not $runOnMacOS -and -not $runOnLinux -and -not $runOnWindows) {
    Write-Warn "Platform unknown — using device code flow as fallback"
    $runOnMacOS = $true
}

if     ($runOnWindows) { Write-OK "Platform: Windows — interactive browser login" }
elseif ($runOnMacOS)   { Write-OK "Platform: macOS   — device code flow" }
elseif ($runOnLinux)   { Write-OK "Platform: Linux   — device code flow" }

#endregion

#region Step 2: Connect to Exchange Online

Write-Step "Connecting to Exchange Online"
try {
    if ($runOnWindows) {
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false
    } else {
        Write-Host "    Device code flow: copy the code and open the URL in your browser." -ForegroundColor DarkGray
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false -Device
    }
    Write-OK "Exchange Online connected as $AdminUPN"
} catch {
    Write-Error "Exchange Online connection failed: $_"
    exit 1
}

#endregion

#region Steps 3 + 4: Create or reuse App Registration

$setupScopes = @("Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All", "Group.Read.All")

if (-not $ClientId -or -not $ClientSecret) {

    Write-Step "Creating App Registration (Mode A — no ClientId/ClientSecret provided)"

    try {
        Connect-GraphDelegated -TenantId $TenantId -Scopes $setupScopes
        Write-OK "Graph connected (delegated) for setup"
    } catch {
        Write-Error "Graph connection failed: $_"
        exit 1
    }

    # Reuse existing app or create new
    $app = Get-MgApplication -Filter "displayName eq '$AppName'" -ErrorAction SilentlyContinue

    if ($app) {
        Write-Warn "App '$AppName' already exists (AppId: $($app.AppId)) — reusing"
    } else {
        if ($PSCmdlet.ShouldProcess($AppName, "Create App Registration")) {
            $app = New-MgApplication -DisplayName $AppName
            Write-OK "App created: $($app.DisplayName) | AppId: $($app.AppId)"
        }
    }

    # Get Graph service principal for permission IDs
    $graphResourceId = "00000003-0000-0000-c000-000000000000"
    $graphSp         = Get-MgServicePrincipal -Filter "appId eq '$graphResourceId'"

    $requiredPerms = @("Calendars.ReadWrite", "Calendars.Read", "Group.Read.All", "Group.ReadWrite.All", "User.Read.All")
    $appRoles      = [System.Collections.Generic.List[object]]::new()

    foreach ($perm in $requiredPerms) {
        $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $perm }
        if ($role) {
            Write-OK "Permission: $perm"
            $appRoles.Add([PSCustomObject]@{ Id = $role.Id; Type = "Role" })
        } else {
            Write-Warn "Permission '$perm' not found"
        }
    }

    if ($PSCmdlet.ShouldProcess($app.AppId, "Set application permissions")) {
        $resourceAccess = $appRoles | ForEach-Object {
            @{ id = $_.Id.ToString(); type = $_.Type }
        }
        Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess @(
            @{
                resourceAppId  = $graphResourceId
                resourceAccess = @($resourceAccess)
            }
        )
        Write-OK "Application permissions set"
    }

    # Create Service Principal for admin consent
    $appSp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue
    if (-not $appSp) {
        if ($PSCmdlet.ShouldProcess($app.AppId, "Create Service Principal")) {
            $appSp = New-MgServicePrincipal -AppId $app.AppId
            Write-OK "Service Principal created"
            Start-Sleep -Seconds 5
        }
    } else {
        Write-OK "Service Principal already exists"
    }

    # Grant admin consent
    foreach ($role in $appRoles) {
        try {
            if ($PSCmdlet.ShouldProcess($role.Id, "Grant admin consent")) {
                New-MgServicePrincipalAppRoleAssignment `
                    -ServicePrincipalId $appSp.Id `
                    -PrincipalId        $appSp.Id `
                    -ResourceId         $graphSp.Id `
                    -AppRoleId          $role.Id `
                    -ErrorAction Stop | Out-Null
                Write-OK "Admin consent granted: $($requiredPerms[$appRoles.IndexOf($role)])"
            }
        } catch {
            if ($_ -match "already exists") {
                Write-Warn "Consent already present — skipped"
            } else {
                Write-Warn "Consent failed: $_"
            }
        }
    }

    # Create Client Secret
    if ($PSCmdlet.ShouldProcess($app.AppId, "Create Client Secret")) {
        $secretEndDate = (Get-Date).AddMonths(3)
        $secret = Add-MgApplicationPassword `
            -ApplicationId $app.Id `
            -PasswordCredential @{
                DisplayName = "CalendarMigration-$(Get-Date -Format 'yyyyMMdd')"
                EndDateTime = $secretEndDate
            }
        $ClientId     = $app.AppId
        $ClientSecret = $secret.SecretText
        Write-OK "Client Secret created (expires: $($secretEndDate.ToString('yyyy-MM-dd')))"
        Write-Host ""
        Write-Host "    Store these values securely (e.g. password manager):" -ForegroundColor Yellow
        Write-Host "    ClientId    : $ClientId"     -ForegroundColor Yellow
        Write-Host "    ClientSecret: $ClientSecret" -ForegroundColor Yellow
        Write-Host ""
    }

    # Wait for consent to propagate in Azure AD (can take 30–60s)
    Write-Host "    Waiting 60s for admin consent propagation in Azure AD..." -ForegroundColor DarkGray
    Start-Sleep -Seconds 60

    Write-Host "    Reconnecting with application permissions..." -ForegroundColor DarkGray
    Disconnect-MgGraph | Out-Null
    Start-Sleep -Seconds 5

} else {
    Write-Step "Using existing App Registration (Mode B)"
    Write-OK "ClientId: $ClientId"
}

#endregion

#region Step 5: Connect Graph with application permissions

Write-Step "Connecting Graph with application permissions"
try {
    Connect-GraphAppAuth -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    $ctx = Get-MgContext
    if (-not $ctx) { throw "Get-MgContext empty after connection" }
    Write-OK "Graph connected | App: $($ctx.AppName) | AuthType: $($ctx.AuthType)"
} catch {
    Write-Error "Graph app auth failed: $_"
    exit 1
}

#endregion

#region Step 6: Create destination mailbox

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
        Write-Host "    Waiting 15s for Exchange provisioning..." -ForegroundColor DarkGray
        Start-Sleep -Seconds 15
    }
}

#endregion

#region Step 7: Calendar permissions and AutoAccept

Write-Step "Setting calendar permissions on $DestinationEmail"

if ($PSCmdlet.ShouldProcess($DestinationEmail, "Set calendar permissions")) {
    Set-MailboxFolderPermission `
        -Identity     "$($DestinationEmail):\Calendar" `
        -User         Default `
        -AccessRights Reviewer `
        -ErrorAction  SilentlyContinue
    Write-OK "Default: Reviewer (read with details)"

    if ($DestinationType -eq "Room") {
        Set-CalendarProcessing `
            -Identity                 $DestinationEmail `
            -AutomateProcessing       AutoAccept `
            -AllowConflicts           $true `
            -AddOrganizerToSubject    $false `
            -DeleteComments           $false `
            -DeleteSubject            $false `
            -BookingWindowInDays      0 `
            -MaximumDurationInMinutes 0
        Write-OK "AutoAccept configured (overlapping bookings allowed, no duration limit)"
    } else {
        Write-OK "Shared Mailbox: AutoAccept not applicable"
        Write-Host "    Users add the calendar manually via: Add calendar > Add from directory" -ForegroundColor DarkGray
    }
}

#endregion

#region Step 8: Reconnect delegated for reading group calendar

# Microsoft Graph blocks Get-MgGroupCalendarEvent for application permissions (AppOnly).
# Known limitation: https://learn.microsoft.com/en-us/graph/known-issues#group-calendar
# Solution: temporarily reconnect as a delegated user to read the group calendar,
# then switch back to app auth to write to the destination mailbox.

Write-Step "Reconnecting as delegated user (required for group calendar)"
Write-Host "    Microsoft Graph blocks group calendar reads via app auth." -ForegroundColor DarkGray
Write-Host "    Delegated connection required for steps 9." -ForegroundColor DarkGray

# Fully disconnect and clear MSAL token cache for a clean session
try { Disconnect-MgGraph | Out-Null } catch {}
Start-Sleep -Seconds 3

# Reload module to clear token cache
Remove-Module Microsoft.Graph.Authentication -Force -ErrorAction SilentlyContinue
Import-Module Microsoft.Graph.Authentication -Force

try {
    $readScopes = @("Group.Read.All", "Calendars.Read", "Calendars.ReadWrite", "User.Read.All")
    if ($runOnWindows) {
        Connect-MgGraph -TenantId $TenantId -Scopes $readScopes -NoWelcome
    } else {
        Write-Host "    Device code flow: copy the code and open the URL in your browser." -ForegroundColor DarkGray
        Connect-MgGraph -TenantId $TenantId -Scopes $readScopes -UseDeviceAuthentication -NoWelcome
    }

    $ctx = Get-MgContext
    if (-not $ctx) { throw "No context after connection" }
    Write-OK "Graph reconnected (delegated) as $($ctx.Account)"
} catch {
    Write-Error "Delegated reconnect failed: $_"
    exit 1
}

#endregion

#region Step 9: Find M365 Group and retrieve events

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
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph | Out-Null
    exit 1
}

Write-OK "Group found: '$($group.DisplayName)' | Mail: $($group.Mail) | ID: $($group.Id)"

Write-Step "Retrieving events ($DaysBack days back to $DaysForward days forward)"

$startDate = (Get-Date).AddDays(-$DaysBack).ToString("yyyy-MM-ddT00:00:00")
$endDate   = (Get-Date).AddDays($DaysForward).ToString("yyyy-MM-ddT23:59:59")

try {
    $calEvents = Get-MgGroupCalendarEvent `
        -GroupId     $group.Id `
        -Filter      "start/dateTime ge '$startDate' and start/dateTime le '$endDate'" `
        -All `
        -ErrorAction Stop
    Write-OK "$($calEvents.Count) events retrieved"
} catch {
    Write-Error "Could not retrieve events: $_"
    exit 1
}

# Switch back to app auth for writing to the destination mailbox
Write-Step "Switching back to application permissions for writing"
Disconnect-MgGraph | Out-Null

try {
    Connect-GraphAppAuth -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    Write-OK "Graph reconnected (app auth) | AuthType: $((Get-MgContext).AuthType)"
} catch {
    Write-Error "App auth reconnect failed: $_"
    exit 1
}

#endregion

#region Step 10: Copy events

Write-Step "Copying events to destination mailbox calendar"

$successCount = 0
$failCount    = 0
$skippedCount = 0

foreach ($calEvent in $calEvents) {

    if ($calEvent.IsCancelled) {
        $skippedCount++
        continue
    }

    if ($PSCmdlet.ShouldProcess($calEvent.Subject, "Copy event to $DestinationEmail")) {
        try {
            $tz = if ($calEvent.Start.TimeZone) { $calEvent.Start.TimeZone } else { "UTC" }

            $params = @{
                Subject  = $calEvent.Subject
                IsAllDay = $calEvent.IsAllDay
                ShowAs   = "oof"
                Start    = @{ DateTime = $calEvent.Start.DateTime; TimeZone = $tz }
                End      = @{ DateTime = $calEvent.End.DateTime;   TimeZone = $tz }
                Body     = @{
                    ContentType = "text"
                    Content     = "Migrated from M365 Group calendar. Original organizer: $($calEvent.Organizer.EmailAddress.Address)"
                }
            }

            New-MgUserEvent -UserId $DestinationEmail -BodyParameter $params | Out-Null
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

#region Step 11: Delete source M365 Group (optional)

if ($DeleteSourceGroup) {
    Write-Step "Deleting source M365 Group: $SourceGroupDisplayName"
    if ($PSCmdlet.ShouldProcess($group.Id, "Delete M365 Group")) {
        try {
            Remove-MgGroup -GroupId $group.Id -Confirm:$false
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

#region Step 12: Summary

$line = "=" * 65
Write-Host "`n$line" -ForegroundColor Cyan
Write-Host " SUMMARY  —  Calendar Migration" -ForegroundColor Cyan
Write-Host $line -ForegroundColor Cyan
Write-Host "  Platform           : $(if ($runOnWindows) { 'Windows' } elseif ($runOnMacOS) { 'macOS' } else { 'Linux' })"
Write-Host "  App Registration   : $ClientId"
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

Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph | Out-Null
Write-OK "Done. Connections closed."

#endregion
