#Requires -Version 7.0
<#
.SYNOPSIS
    Report and optionally bulk-set the manager for a set of Entra ID users.

.DESCRIPTION
    Retrieves users directly from Entra ID via Microsoft Graph — by group membership,
    department, or an inline UPN list — then shows each user's current manager.

    When -NewManager is provided, every resolved user gets that manager assigned in
    one go.

    User source options (pick one):
    - -GroupId / -GroupName  : all (direct) members of an Entra group
    - -Department            : all users in a department (OData filter)
    - -CurrentManager        : all direct reports of a given manager (searches all of Entra ID)
    - -UserList              : explicit UPN / object-ID list

.PARAMETER GroupId
    Object ID of an Entra ID group whose members should be processed.

.PARAMETER GroupName
    Display name of an Entra ID group. Resolved to an ID automatically.
    When multiple groups match, the script errors and asks you to use -GroupId.

.PARAMETER Department
    Department string to filter users by (case-insensitive, exact match via OData).

.PARAMETER CurrentManager
    UPN or object ID of a manager. All their direct reports in Entra ID are retrieved
    and processed. Useful for finding everyone who reports to a specific person.

.PARAMETER UserList
    Array of UPNs or object IDs to process directly.

.PARAMETER NewManager
    UPN or object ID of the user to set as manager for all resolved users.
    When omitted the script only reports the current manager.

.PARAMETER OutputPath
    Path to export results as CSV. Optional.

.PARAMETER TenantId
    Optional tenant ID or domain. Defaults to the GDAP customer when authMode is GDAP.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint and -TenantId).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.NOTES
    Sign-in goes through scripts\Startup\Connect-M365.ps1: delegated as the admin by
    default (device code / GDAP customer per load.config.ps1), app-only on request.
    A fitting Graph session is reused and left connected; only a session this script
    opened is disconnected. Delegated scopes: User.Read.All + GroupMember.Read.All,
    User.ReadWrite.All when -NewManager is given.

.EXAMPLE
    # Show managers for all members of a group
    .\Set-UserManager.ps1 -GroupName "Sales Team"

.EXAMPLE
    # Bulk-set manager for a group
    .\Set-UserManager.ps1 -GroupName "Sales Team" -NewManager "jane.doe@contoso.com"

.EXAMPLE
    # Bulk-set manager for a department
    .\Set-UserManager.ps1 -Department "Logistics" -NewManager "jane.doe@contoso.com" -OutputPath C:\Temp\ManagerReport.csv

.EXAMPLE
    # Find and show all direct reports of a manager
    .\Set-UserManager.ps1 -CurrentManager "old.boss@contoso.com"

.EXAMPLE
    # Re-assign all direct reports of one manager to another
    .\Set-UserManager.ps1 -CurrentManager "old.boss@contoso.com" -NewManager "new.boss@contoso.com"

.EXAMPLE
    # Explicit UPN list
    .\Set-UserManager.ps1 -UserList "john@contoso.com","pete@contoso.com" -NewManager "jane.doe@contoso.com"
#>

[CmdletBinding(DefaultParameterSetName = 'ByGroup', SupportsShouldProcess)]
param (
    [Parameter(ParameterSetName = 'ByGroupId', Mandatory)]
    [string] $GroupId,

    [Parameter(ParameterSetName = 'ByGroup', Mandatory)]
    [string] $GroupName,

    [Parameter(ParameterSetName = 'ByDepartment', Mandatory)]
    [string] $Department,

    [Parameter(ParameterSetName = 'ByCurrentManager', Mandatory)]
    [string] $CurrentManager,

    [Parameter(ParameterSetName = 'ByList', Mandatory)]
    [string[]] $UserList,

    [string] $NewManager,

    [string] $OutputPath,

    [string] $TenantId,

    [string] $ClientId,

    [string] $CertificateThumbprint,

    [switch] $AppOnly
)

begin {
    $ErrorActionPreference = 'Stop'

    . (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')

    # ── Connect (delegated by default; reuses a fitting session) ──────────────
    $scopes = if ($NewManager) {
        @('User.ReadWrite.All', 'GroupMember.Read.All')
    } else {
        @('User.Read.All', 'GroupMember.Read.All')
    }
    $script:Graph = Connect-M365Graph -Scopes $scopes -TenantId $TenantId `
        -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

    # Strict mode only after connecting: the helper reads optional globals
    # ($global:authMode, $global:cid, ...) that strict mode treats as errors when
    # load.ps1 has not set them.
    Set-StrictMode -Version Latest

    # ── Helpers ───────────────────────────────────────────────────────────────
    $userSelect = 'id,displayName,userPrincipalName,department'

    function ConvertTo-UserRow($u) {
        [PSCustomObject]@{
            Id                = $u['id']
            DisplayName       = $u['displayName']
            UserPrincipalName = $u['userPrincipalName']
            Department        = $u['department']
        }
    }

    # GET with paging; the /microsoft.graph.user cast returns only users, with the
    # selected properties, so there is no Get-MgUser per member.
    function Get-GraphUserPage([string] $Uri) {
        $next = $Uri
        while ($next) {
            $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType Hashtable
            foreach ($u in $page['value']) { ConvertTo-UserRow $u }
            $next = $page['@odata.nextLink']
        }
    }

    # ── Resolve users from Entra ID ───────────────────────────────────────────
    $resolvedUsers = [System.Collections.Generic.List[PSCustomObject]]::new()

    switch ($PSCmdlet.ParameterSetName) {
        'ByGroup' {
            Write-Host "Searching for group: $GroupName" -ForegroundColor Cyan
            $escaped = $GroupName -replace "'", "''"
            $groups = @(Get-MgGroup -Filter "displayName eq '$escaped'" -Property Id, DisplayName -All)
            if ($groups.Count -eq 0) { throw "No group found with display name '$GroupName'." }
            if ($groups.Count -gt 1) { throw "Multiple groups match '$GroupName'. Use -GroupId instead." }
            $GroupId = $groups[0].Id
            Write-Host "  -> Group ID: $GroupId" -ForegroundColor Green
            foreach ($u in Get-GraphUserPage "v1.0/groups/$GroupId/members/microsoft.graph.user?`$select=$userSelect&`$top=999") {
                $resolvedUsers.Add($u)
            }
        }
        'ByGroupId' {
            Write-Host "Fetching members of group $GroupId..." -ForegroundColor Cyan
            foreach ($u in Get-GraphUserPage "v1.0/groups/$GroupId/members/microsoft.graph.user?`$select=$userSelect&`$top=999") {
                $resolvedUsers.Add($u)
            }
        }
        'ByDepartment' {
            Write-Host "Fetching users in department: $Department" -ForegroundColor Cyan
            $escaped = $Department -replace "'", "''"
            $deptUsers = Get-MgUser -Filter "department eq '$escaped'" -Property Id, DisplayName, UserPrincipalName, Department -All
            foreach ($u in $deptUsers) {
                $resolvedUsers.Add([PSCustomObject]@{ Id = $u.Id; DisplayName = $u.DisplayName; UserPrincipalName = $u.UserPrincipalName; Department = $u.Department })
            }
        }
        'ByCurrentManager' {
            Write-Host "Resolving manager: $CurrentManager" -ForegroundColor Cyan
            $mgr = Get-MgUser -UserId $CurrentManager -Property Id, DisplayName, UserPrincipalName
            Write-Host "  -> $($mgr.DisplayName) — fetching direct reports from all of Entra ID..." -ForegroundColor Green
            $reports = @(Get-GraphUserPage "v1.0/users/$($mgr.Id)/directReports/microsoft.graph.user?`$select=$userSelect")
            if ($reports.Count -eq 0) {
                Write-Warning "No direct reports found for $($mgr.DisplayName)."
            }
            foreach ($r in $reports) { $resolvedUsers.Add($r) }
        }
        'ByList' {
            Write-Host "Resolving $($UserList.Count) user(s) from list..." -ForegroundColor Cyan
            foreach ($entry in $UserList) {
                $u = Get-MgUser -UserId $entry.Trim() -Property Id, DisplayName, UserPrincipalName, Department
                $resolvedUsers.Add([PSCustomObject]@{ Id = $u.Id; DisplayName = $u.DisplayName; UserPrincipalName = $u.UserPrincipalName; Department = $u.Department })
            }
        }
    }

    Write-Host "$($resolvedUsers.Count) user(s) to process.`n" -ForegroundColor Cyan

    # ── Resolve new manager object ID (once) ──────────────────────────────────
    $newManagerId = $null
    $newManagerDisplayName = $null
    if ($NewManager) {
        Write-Host "Resolving new manager: $NewManager" -ForegroundColor Cyan
        $mgUser = Get-MgUser -UserId $NewManager -Property Id, DisplayName, UserPrincipalName
        $newManagerId = $mgUser.Id
        $newManagerDisplayName = $mgUser.DisplayName
        Write-Host "  -> $newManagerDisplayName ($newManagerId)" -ForegroundColor Green
    }

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
}

process {
    foreach ($user in $resolvedUsers) {
        Write-Host "Processing $($user.UserPrincipalName) ..." -NoNewline

        try {
            # Get current manager
            $currentManagerName = '(none)'
            $currentManagerUpn  = ''
            try {
                # One call: the manager with the properties we show (404 when none is set)
                $mgrDetails = Invoke-MgGraphRequest -Method GET -OutputType Hashtable `
                    -Uri "v1.0/users/$($user.Id)/manager?`$select=displayName,userPrincipalName"
                $currentManagerName = $mgrDetails['displayName']
                $currentManagerUpn  = $mgrDetails['userPrincipalName']
            }
            catch {
                # No manager set — keep defaults
            }

            $status = 'OK'

            # Set new manager if requested
            if ($newManagerId) {
                if ($PSCmdlet.ShouldProcess($user.UserPrincipalName, "Set manager to $newManagerDisplayName")) {
                    $body = @{
                        '@odata.id' = "https://graph.microsoft.com/v1.0/users/$newManagerId"
                    }
                    Invoke-MgGraphRequest -Method PUT `
                        -Uri "https://graph.microsoft.com/v1.0/users/$($user.Id)/manager/`$ref" `
                        -Body $body
                    $status = "Manager set to $newManagerDisplayName"
                }
            }

            Write-Host " $status" -ForegroundColor Green

            $results.Add([PSCustomObject]@{
                UserPrincipalName  = $user.UserPrincipalName
                DisplayName        = $user.DisplayName
                Department         = $user.Department
                PreviousManager    = $currentManagerName
                PreviousManagerUPN = $currentManagerUpn
                NewManager         = if ($newManagerId) { $newManagerDisplayName } else { '' }
                NewManagerUPN      = if ($newManagerId) { $NewManager } else { '' }
                Status             = $status
            })
        }
        catch {
            Write-Host " ERROR: $_" -ForegroundColor Red
            $results.Add([PSCustomObject]@{
                UserPrincipalName  = $user.UserPrincipalName
                DisplayName        = $user.DisplayName
                Department         = $user.Department
                PreviousManager    = ''
                PreviousManagerUPN = ''
                NewManager         = ''
                NewManagerUPN      = ''
                Status             = "ERROR: $_"
            })
        }
    }
}

end {
    Write-Host "`nResults:" -ForegroundColor Cyan
    $results | Format-Table UserPrincipalName, DisplayName, PreviousManager, NewManager, Status -AutoSize

    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Report exported to: $OutputPath" -ForegroundColor Green
    }

    # Only disconnect a session this script opened, never the caller's.
    Disconnect-M365Graph $script:Graph
}
