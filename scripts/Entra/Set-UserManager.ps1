#Requires -Version 5.1
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
    Optional tenant ID or domain for Connect-MgGraph.

.EXAMPLE
    # Show managers for all members of a group
    .\Set-UserManager.ps1 -GroupName "Sales Team"

.EXAMPLE
    # Bulk-set manager for a group
    .\Set-UserManager.ps1 -GroupName "Sales Team" -NewManager "jane.doe@aslgroup.eu"

.EXAMPLE
    # Bulk-set manager for a department
    .\Set-UserManager.ps1 -Department "Logistics" -NewManager "jane.doe@aslgroup.eu" -OutputPath C:\Temp\ManagerReport.csv

.EXAMPLE
    # Find and show all direct reports of a manager
    .\Set-UserManager.ps1 -CurrentManager "old.boss@aslgroup.eu"

.EXAMPLE
    # Re-assign all direct reports of one manager to another
    .\Set-UserManager.ps1 -CurrentManager "old.boss@aslgroup.eu" -NewManager "new.boss@aslgroup.eu"

.EXAMPLE
    # Explicit UPN list
    .\Set-UserManager.ps1 -UserList "john@aslgroup.eu","pete@aslgroup.eu" -NewManager "jane.doe@aslgroup.eu"
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

    [string] $TenantId
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    # ── Connect ───────────────────────────────────────────────────────────────
    $scopes = if ($NewManager) {
        @('User.Read.All', 'User.ReadWrite.All', 'GroupMember.Read.All')
    } else {
        @('User.Read.All', 'GroupMember.Read.All')
    }
    $connectParams = @{ Scopes = $scopes }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }

    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph @connectParams -NoWelcome

    # ── Resolve users from Entra ID ───────────────────────────────────────────
    $resolvedUsers = [System.Collections.Generic.List[Microsoft.Graph.PowerShell.Models.IMicrosoftGraphUser]]::new()

    switch ($PSCmdlet.ParameterSetName) {
        'ByGroup' {
            Write-Host "Searching for group: $GroupName" -ForegroundColor Cyan
            $groups = Get-MgGroup -Filter "displayName eq '$GroupName'" -Property Id, DisplayName
            if ($groups.Count -eq 0) { throw "No group found with display name '$GroupName'." }
            if ($groups.Count -gt 1) { throw "Multiple groups match '$GroupName'. Use -GroupId instead." }
            $GroupId = $groups[0].Id
            Write-Host "  -> Group ID: $GroupId" -ForegroundColor Green
            # fall through to ByGroupId logic
            $members = Get-MgGroupMember -GroupId $GroupId -All | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' }
            foreach ($m in $members) {
                $resolvedUsers.Add((Get-MgUser -UserId $m.Id -Property Id, DisplayName, UserPrincipalName, Department))
            }
        }
        'ByGroupId' {
            Write-Host "Fetching members of group $GroupId..." -ForegroundColor Cyan
            $members = Get-MgGroupMember -GroupId $GroupId -All | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' }
            foreach ($m in $members) {
                $resolvedUsers.Add((Get-MgUser -UserId $m.Id -Property Id, DisplayName, UserPrincipalName, Department))
            }
        }
        'ByDepartment' {
            Write-Host "Fetching users in department: $Department" -ForegroundColor Cyan
            $deptUsers = Get-MgUser -Filter "department eq '$Department'" -Property Id, DisplayName, UserPrincipalName, Department -All
            foreach ($u in $deptUsers) { $resolvedUsers.Add($u) }
        }
        'ByCurrentManager' {
            Write-Host "Resolving manager: $CurrentManager" -ForegroundColor Cyan
            $mgr = Get-MgUser -UserId $CurrentManager -Property Id, DisplayName, UserPrincipalName
            Write-Host "  -> $($mgr.DisplayName) — fetching direct reports from all of Entra ID..." -ForegroundColor Green
            $reports = Get-MgUserDirectReport -UserId $mgr.Id -All
            if ($reports.Count -eq 0) {
                Write-Warning "No direct reports found for $($mgr.DisplayName)."
            }
            foreach ($r in $reports) {
                $resolvedUsers.Add((Get-MgUser -UserId $r.Id -Property Id, DisplayName, UserPrincipalName, Department))
            }
        }
        'ByList' {
            Write-Host "Resolving $($UserList.Count) user(s) from list..." -ForegroundColor Cyan
            foreach ($entry in $UserList) {
                $resolvedUsers.Add((Get-MgUser -UserId $entry.Trim() -Property Id, DisplayName, UserPrincipalName, Department))
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
                $mgr = Get-MgUserManager -UserId $user.Id
                $mgrDetails = Get-MgUser -UserId $mgr.Id -Property DisplayName, UserPrincipalName
                $currentManagerName = $mgrDetails.DisplayName
                $currentManagerUpn  = $mgrDetails.UserPrincipalName
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

    Disconnect-MgGraph | Out-Null
}
