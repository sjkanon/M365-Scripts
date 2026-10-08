#Requires -Version 7.0
<#
.SYNOPSIS
    Push a shared contact list (e.g. a company phone list) into one or more
    users' personal Outlook Contacts folder via Microsoft Graph.

.DESCRIPTION
    Reads a CSV of contacts and creates a personal contact for each row in every
    target user's Contacts folder. Every contact created this way is tagged with
    a marker in its PersonalNotes field so a later run can find and remove only
    the previously-synced contacts before re-importing a refreshed list (pass
    -RemoveExisting), without touching contacts the user added themselves.

    Target users can be an explicit list (-UserList) or every member of a group
    (-GroupId). Defaults to a safe preview — pass -Apply to actually write
    contacts.

    Sign-in: APP-ONLY BY DEFAULT - the one script in this repo that does not default
    to delegated. Writing contacts into OTHER users' mailboxes needs the
    Contacts.ReadWrite application permission; a delegated token (even a Global
    Administrator's) only reaches the signed-in user's own contacts. The app comes
    from -ClientId + -CertificateThumbprint, or else from graph.appid.json in the repo
    root (as -AppOnly). With -Delegated you sign in as yourself (device code / GDAP
    per load.config.ps1) and can only target your own mailbox. Sign-in goes through
    scripts\Startup\Connect-M365.ps1; a fitting Graph session is reused and left
    connected, and only a session this script opened is disconnected.

.PARAMETER CsvPath
    Path to a CSV with columns: DisplayName, GivenName, Surname, CompanyName,
    BusinessPhone, MobilePhone, EmailAddress. Only DisplayName and EmailAddress
    are required; the rest are optional.

.PARAMETER UserList
    Explicit array of UPNs/object IDs whose Contacts folder should receive the
    list.

.PARAMETER GroupId
    Object ID of a group — every member gets the contact list pushed to them.

.PARAMETER Tag
    Marker written into each created contact's PersonalNotes, used to identify
    contacts this script created (so -RemoveExisting can clean them up without
    touching the user's own contacts). Default: "Synced-by-Sync-UserContacts".

.PARAMETER RemoveExisting
    Remove contacts previously created by this script (matched by -Tag) for each
    target user before importing the current CSV. Use this for a clean refresh
    instead of accumulating duplicates on every run.

.PARAMETER Apply
    Actually create/remove contacts. Without this switch, the script only
    reports what it would do.

.PARAMETER OutputPath
    CSV report path. Defaults to `C:\Temp\SyncUserContacts_<timestamp>.csv`
    (`~/Downloads` on Linux/macOS).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Defaults to the GDAP customer when authMode is GDAP;
    selects the tenant's entry in graph.appid.json.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint). Needs the
    application permissions Contacts.ReadWrite (and GroupMember.Read.All for -GroupId).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only with ClientId and CertificateThumbprint from graph.appid.json. This is
    already the default; the switch is accepted for consistency with other scripts.

.PARAMETER Delegated
    Sign in as yourself instead of app-only. Delegated Contacts.ReadWrite only covers
    your own mailbox, so every target user must be you; the script stops otherwise.

.EXAMPLE
    # Preview for an explicit user list
    .\Sync-UserContacts.ps1 -CsvPath .\companycontacts.csv -UserList "user1@contoso.com","user2@contoso.com"

.EXAMPLE
    # Push to every member of a group, refreshing previously-synced contacts first
    .\Sync-UserContacts.ps1 -CsvPath .\companycontacts.csv -GroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -RemoveExisting -Apply

.EXAMPLE
    # App registration passed explicitly
    .\Sync-UserContacts.ps1 -CsvPath .\companycontacts.csv -UserList "user1@contoso.com" -TenantId contoso.onmicrosoft.com -ClientId <appId> -CertificateThumbprint <thumbprint> -Apply

.EXAMPLE
    # Only your own Contacts folder, signed in as yourself
    .\Sync-UserContacts.ps1 -CsvPath .\companycontacts.csv -UserList "me@contoso.com" -Delegated -Apply

.NOTES
    CSV example:
        DisplayName,GivenName,Surname,CompanyName,BusinessPhone,MobilePhone,EmailAddress
        "Jane Doe",Jane,Doe,"Contoso Ltd","+1 555 0100","+1 555 0101",jane.doe@example.com

    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.PersonalContacts
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'UserList')]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $CsvPath,

    [Parameter(ParameterSetName = 'UserList', Mandatory)]
    [string[]] $UserList,

    [Parameter(ParameterSetName = 'Group', Mandatory)]
    [string] $GroupId,

    [string] $Tag = 'Synced-by-Sync-UserContacts',
    [switch] $RemoveExisting,
    [switch] $Apply,
    [string] $OutputPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly,
    [switch] $Delegated
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

if ($Delegated -and ($AppOnly -or $ClientId)) {
    throw '-Delegated cannot be combined with -AppOnly or -ClientId: pick one sign-in.'
}

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
if (-not $OutputPath) { $OutputPath = Join-Path $outputDir "SyncUserContacts_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" }

$rows = Import-Csv -Path $CsvPath
# -notcontains: the old "-not $names -contains 'X'" negated the list first, so the
# check never fired and a CSV without these columns was processed silently.
if (-not $rows -or $rows[0].PSObject.Properties.Name -notcontains 'DisplayName' -or $rows[0].PSObject.Properties.Name -notcontains 'EmailAddress') {
    Write-Error "CSV must have at least 'DisplayName' and 'EmailAddress' columns. Found: $($rows[0].PSObject.Properties.Name -join ', ')"
    exit 1
}

# ── Connection ────────────────────────────────────────────────────────────────
if ($Delegated) {
    $scopes = @('Contacts.ReadWrite', 'User.Read')
    if ($PSCmdlet.ParameterSetName -eq 'Group') { $scopes += 'GroupMember.Read.All' }
    $graph = Connect-M365Graph -Scopes $scopes -TenantId $TenantId
} else {
    # App-only by default: delegated Contacts.ReadWrite cannot write other users'
    # contacts, which is the whole point of this script.
    try {
        $graph = Connect-M365Graph -TenantId $TenantId -ClientId $ClientId `
            -CertificateThumbprint $CertificateThumbprint -AppOnly:(-not $ClientId)
    } catch {
        throw ("This script signs in app-only by default, because a delegated token can only write the signed-in user's own contacts. " +
               "Pass -ClientId and -CertificateThumbprint (app with the Contacts.ReadWrite application permission), add the tenant to graph.appid.json, " +
               "or use -Delegated to target only your own mailbox. Sign-in error: $($_.Exception.Message)")
    }
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Sync-UserContacts" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ("  Mode : {0}" -f $(if ($Apply) { 'Apply' } else { 'Preview only' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host "  Rows : $($rows.Count)"
Write-Host ""

# ── Resolve target users ──────────────────────────────────────────────────────
$targetUsers = [System.Collections.Generic.List[string]]::new()
if ($PSCmdlet.ParameterSetName -eq 'Group') {
    Write-Host "  Resolving group members..." -ForegroundColor DarkGray
    # Only users have a Contacts folder; the cast skips nested groups, devices, etc.
    $next = "v1.0/groups/$GroupId/members/microsoft.graph.user?`$select=id&`$top=999"
    while ($next) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType Hashtable -ErrorAction Stop
        foreach ($m in $page['value']) { $targetUsers.Add($m['id']) }
        $next = $page['@odata.nextLink']
    }
} else {
    $UserList | ForEach-Object { $targetUsers.Add($_) }
}

Write-Host "  Target user(s): $($targetUsers.Count)" -ForegroundColor DarkGray
Write-Host ""

if ($Delegated) {
    $me = Invoke-MgGraphRequest -Method GET -Uri 'v1.0/me?$select=id,userPrincipalName' -OutputType Hashtable -ErrorAction Stop
    $others = @($targetUsers | Where-Object { $_ -ne $me['id'] -and $_ -ne $me['userPrincipalName'] })
    if ($others.Count -gt 0) {
        Disconnect-M365Graph $graph
        throw "-Delegated can only write your own contacts ($($me['userPrincipalName'])); these targets are other users: $($others -join ', '). Run without -Delegated (app-only, Contacts.ReadWrite application permission)."
    }
}

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($userId in $targetUsers) {
    Write-Host "  -> $userId" -ForegroundColor Cyan

    if (-not $Apply) {
        if ($RemoveExisting) { Write-Host "     [PREVIEW] would remove previously-synced contacts (tag: $Tag)" -ForegroundColor DarkGray }
        Write-Host "     [PREVIEW] would create $($rows.Count) contact(s)" -ForegroundColor DarkGray
        $results.Add([PSCustomObject]@{ User = $userId; Removed = 'n/a'; Created = 'Preview'; Errors = 0 })
        continue
    }

    $removedCount = 0
    if ($RemoveExisting) {
        if ($PSCmdlet.ShouldProcess($userId, "Remove previously-synced contacts (tag: $Tag)")) {
            try {
                $existing = Get-MgUserContact -UserId $userId -All -ErrorAction Stop |
                    Where-Object { $_.PersonalNotes -match [regex]::Escape($Tag) }
                foreach ($c in $existing) {
                    Remove-MgUserContact -UserId $userId -ContactId $c.Id -ErrorAction SilentlyContinue
                    $removedCount++
                }
                Write-Host "     Removed $removedCount previously-synced contact(s)" -ForegroundColor DarkGray
            } catch {
                Write-Host "     [WARN] Could not list/remove existing contacts: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }

    $createdCount = 0
    $errorCount   = 0
    foreach ($row in $rows) {
        if (-not $row.DisplayName -or -not $row.EmailAddress) { continue }
        if (-not $PSCmdlet.ShouldProcess("$userId : $($row.DisplayName)", "Create contact")) { continue }

        $body = @{
            givenName      = $row.GivenName
            surname        = $row.Surname
            displayName    = $row.DisplayName
            companyName    = $row.CompanyName
            personalNotes  = "$Tag — imported $(Get-Date -Format 'yyyy-MM-dd')"
            emailAddresses = @(@{ address = $row.EmailAddress; name = $row.DisplayName })
        }
        if ($row.PSObject.Properties.Name -contains 'BusinessPhone' -and $row.BusinessPhone) { $body['businessPhones'] = @($row.BusinessPhone) }
        if ($row.PSObject.Properties.Name -contains 'MobilePhone' -and $row.MobilePhone) { $body['mobilePhone'] = $row.MobilePhone }

        try {
            New-MgUserContact -UserId $userId -BodyParameter $body -ErrorAction Stop | Out-Null
            $createdCount++
        } catch {
            $errorCount++
            Write-Host "     [WARN] $($row.DisplayName): $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    Write-Host "     Created $createdCount contact(s), $errorCount error(s)" -ForegroundColor Green
    $results.Add([PSCustomObject]@{ User = $userId; Removed = $removedCount; Created = $createdCount; Errors = $errorCount })
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
if (-not $Apply) { Write-Host "  Re-run with -Apply to write contacts." -ForegroundColor Yellow }
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
Disconnect-M365Graph $graph
