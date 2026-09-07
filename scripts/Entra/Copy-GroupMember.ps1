#Requires -Version 5.1
<#
.SYNOPSIS
    Copy the members of one Entra ID group into another group via Microsoft Graph.

.DESCRIPTION
    Reads the members of a source group and adds any that are missing from the
    target group. Members that are already present are skipped, so the script is
    safe to re-run.

    Defaults to dry-run mode — pass -Apply to actually write changes.

    Optionally:
      - -Flatten expands nested groups so only the effective users are copied
      - -MemberType limits the copy to a single object type (User, Group, ...)
      - -Mirror also removes members from the target that are not in the source
        (making the target an exact copy instead of a union)

.PARAMETER SourceGroup
    Display name or Object ID of the group to copy members FROM.

.PARAMETER TargetGroup
    Display name or Object ID of the group to copy members TO.

.PARAMETER MemberType
    Restrict which member types are copied: All (default), User, Group, Device
    or ServicePrincipal.

.PARAMETER Flatten
    Resolve the source group transitively, so members of nested groups are
    copied as individual objects instead of copying the nested group itself.

.PARAMETER Mirror
    Also remove members from the target group that are not present in the source
    group. Without this switch the copy is additive only.

.PARAMETER Apply
    Actually add (and with -Mirror remove) members. Without it the script only
    reports what it would do.

.PARAMETER Disconnect
    Sign out of Microsoft Graph when finished. Off by default: Disconnect-MgGraph
    clears the SDK token cache, which means a new browser prompt on every run.

.PARAMETER OutputPath
    CSV report path. Defaults to C:\Temp\GroupMemberCopy_<timestamp>.csv
    (~/Downloads on non-Windows).

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    # Dry run — show what would be copied
    .\Copy-GroupMember.ps1 -SourceGroup "All Staff" -TargetGroup "MFA Rollout"

.EXAMPLE
    # Actually copy the members
    .\Copy-GroupMember.ps1 -SourceGroup "All Staff" -TargetGroup "MFA Rollout" -Apply

.EXAMPLE
    # Copy only users, expanding nested groups
    .\Copy-GroupMember.ps1 -SourceGroup "Sales" -TargetGroup "Sales Mail" -MemberType User -Flatten -Apply

.EXAMPLE
    # Make the target an exact copy of the source (adds and removes)
    .\Copy-GroupMember.ps1 -SourceGroup "Pilot" -TargetGroup "Pilot Copy" -Mirror -Apply
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string] $SourceGroup,

    [Parameter(Mandatory = $true)]
    [string] $TargetGroup,

    [ValidateSet('All', 'User', 'Group', 'Device', 'ServicePrincipal')]
    [string] $MemberType = 'All',

    [switch] $Flatten,
    [switch] $Mirror,
    [switch] $Apply,
    [switch] $Disconnect,

    [string] $OutputPath,
    [string] $TenantId
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    if (-not (Get-MgContext -ErrorAction Stop)) { throw }
} catch {
    $connectParams = @{
        Scopes = @('Group.Read.All', 'GroupMember.ReadWrite.All', 'Directory.Read.All')
    }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
}

# ── Helpers ───────────────────────────────────────────────────────────────────
function Resolve-Group {
    <#
        Accepts an Object ID or a display name. Display names are not unique in
        Entra ID, so an ambiguous name is a hard error — silently picking the
        first match could write members into the wrong group.
    #>
    param(
        [string] $Identity,
        [string] $Label
    )

    if ($Identity -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
        try {
            return Get-MgGroup -GroupId $Identity -ErrorAction Stop
        } catch {
            throw "$Label group '$Identity' not found: $($_.Exception.Message)"
        }
    }

    # Not $matches — that is an automatic variable filled by the -match above
    $escaped = $Identity -replace "'", "''"
    $found = @(Get-MgGroup -Filter "displayName eq '$escaped'" -All -ErrorAction Stop)

    if ($found.Count -eq 0) { throw "$Label group '$Identity' not found." }
    if ($found.Count -gt 1) {
        throw "$Label group '$Identity' is ambiguous ($($found.Count) matches). Use the Object ID instead."
    }
    return $found[0]
}

function Get-MemberObjectType {
    # The SDK returns DirectoryObject; the real type sits in @odata.type
    param($Member)

    $odata = $null
    if ($Member.AdditionalProperties -and $Member.AdditionalProperties['@odata.type']) {
        $odata = $Member.AdditionalProperties['@odata.type']
    }
    if (-not $odata) { return 'Unknown' }

    switch ($odata) {
        '#microsoft.graph.user'             { 'User' }
        '#microsoft.graph.group'            { 'Group' }
        '#microsoft.graph.device'           { 'Device' }
        '#microsoft.graph.servicePrincipal' { 'ServicePrincipal' }
        '#microsoft.graph.orgContact'       { 'OrgContact' }
        default                             { $odata -replace '^#microsoft\.graph\.', '' }
    }
}

function Get-MemberLabel {
    param($Member)

    $props = $Member.AdditionalProperties
    if ($props) {
        foreach ($key in 'userPrincipalName', 'displayName', 'mail') {
            if ($props[$key]) { return $props[$key] }
        }
    }
    return $Member.Id
}

# ── Resolve both groups ───────────────────────────────────────────────────────
$source = Resolve-Group -Identity $SourceGroup -Label 'Source'
$target = Resolve-Group -Identity $TargetGroup -Label 'Target'

if ($source.Id -eq $target.Id) { throw "Source and target group are the same group." }

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Copy Group Members" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Source : $($source.DisplayName)  [$($source.Id)]" -ForegroundColor Gray
Write-Host "   Target : $($target.DisplayName)  [$($target.Id)]" -ForegroundColor Gray
Write-Host "   Mode   : $(if ($Apply) { 'APPLY' } else { 'DRY RUN' })$(if ($Mirror) { ' + MIRROR' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

# Dynamic groups cannot be edited by hand — membership is rule-driven
if ($target.GroupTypes -contains 'DynamicMembership') {
    throw "Target group '$($target.DisplayName)' has dynamic membership; members cannot be added manually."
}

# ── Read members ──────────────────────────────────────────────────────────────
if ($Flatten) {
    $sourceMembers = @(Get-MgGroupTransitiveMember -GroupId $source.Id -All -ErrorAction Stop)
    # Nested groups themselves are noise once their members are expanded
    $sourceMembers = @($sourceMembers | Where-Object { (Get-MemberObjectType -Member $_) -ne 'Group' })
} else {
    $sourceMembers = @(Get-MgGroupMember -GroupId $source.Id -All -ErrorAction Stop)
}

if ($MemberType -ne 'All') {
    $sourceMembers = @($sourceMembers | Where-Object { (Get-MemberObjectType -Member $_) -eq $MemberType })
}

$targetMembers   = @(Get-MgGroupMember -GroupId $target.Id -All -ErrorAction Stop)
$targetMemberIds = @($targetMembers.Id)
$sourceMemberIds = @($sourceMembers.Id)

Write-Host "  Source members : $($sourceMembers.Count)$(if ($Flatten) { ' (flattened)' })" -ForegroundColor DarkGray
Write-Host "  Target members : $($targetMembers.Count)" -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

# ── Add missing members ───────────────────────────────────────────────────────
foreach ($member in $sourceMembers) {
    $label = Get-MemberLabel -Member $member
    $type  = Get-MemberObjectType -Member $member

    if ($member.Id -in $targetMemberIds) {
        $results.Add([PSCustomObject]@{
            Action = 'Skip'; ObjectType = $type; Member = $label
            ObjectId = $member.Id; Status = 'AlreadyMember'; Message = ''
        })
        continue
    }

    if (-not $Apply) {
        Write-Host "  [DRY] Would add    : $label ($type)" -ForegroundColor DarkGray
        $results.Add([PSCustomObject]@{
            Action = 'Add'; ObjectType = $type; Member = $label
            ObjectId = $member.Id; Status = 'WouldAdd'; Message = ''
        })
        continue
    }

    if ($PSCmdlet.ShouldProcess($label, "Add to group '$($target.DisplayName)'")) {
        try {
            New-MgGroupMember -GroupId $target.Id -DirectoryObjectId $member.Id -ErrorAction Stop
            Write-Host "  [OK]  Added        : $label ($type)" -ForegroundColor Green
            $results.Add([PSCustomObject]@{
                Action = 'Add'; ObjectType = $type; Member = $label
                ObjectId = $member.Id; Status = 'Added'; Message = ''
            })
        } catch {
            Write-Host "  [ERR] Add failed   : $label — $($_.Exception.Message)" -ForegroundColor Red
            $results.Add([PSCustomObject]@{
                Action = 'Add'; ObjectType = $type; Member = $label
                ObjectId = $member.Id; Status = 'Failed'; Message = $_.Exception.Message
            })
        }
    }
}

# ── Mirror: remove members the source does not have ───────────────────────────
if ($Mirror) {
    foreach ($member in $targetMembers) {
        if ($member.Id -in $sourceMemberIds) { continue }

        $label = Get-MemberLabel -Member $member
        $type  = Get-MemberObjectType -Member $member

        # With a type filter, only mirror-remove objects of that same type —
        # otherwise -MemberType User would strip nested groups from the target
        if ($MemberType -ne 'All' -and $type -ne $MemberType) { continue }

        if (-not $Apply) {
            Write-Host "  [DRY] Would remove : $label ($type)" -ForegroundColor DarkGray
            $results.Add([PSCustomObject]@{
                Action = 'Remove'; ObjectType = $type; Member = $label
                ObjectId = $member.Id; Status = 'WouldRemove'; Message = ''
            })
            continue
        }

        if ($PSCmdlet.ShouldProcess($label, "Remove from group '$($target.DisplayName)'")) {
            try {
                Remove-MgGroupMemberByRef -GroupId $target.Id -DirectoryObjectId $member.Id -ErrorAction Stop
                Write-Host "  [OK]  Removed      : $label ($type)" -ForegroundColor Yellow
                $results.Add([PSCustomObject]@{
                    Action = 'Remove'; ObjectType = $type; Member = $label
                    ObjectId = $member.Id; Status = 'Removed'; Message = ''
                })
            } catch {
                Write-Host "  [ERR] Remove failed: $label — $($_.Exception.Message)" -ForegroundColor Red
                $results.Add([PSCustomObject]@{
                    Action = 'Remove'; ObjectType = $type; Member = $label
                    ObjectId = $member.Id; Status = 'Failed'; Message = $_.Exception.Message
                })
            }
        }
    }
}

# ── Report ────────────────────────────────────────────────────────────────────
if (-not $OutputPath) {
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path $outputDir "GroupMemberCopy_$ts.csv"
}
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

$added   = @($results | Where-Object { $_.Status -in 'Added', 'WouldAdd' }).Count
$removed = @($results | Where-Object { $_.Status -in 'Removed', 'WouldRemove' }).Count
$skipped = @($results | Where-Object { $_.Status -eq 'AlreadyMember' }).Count
$failed  = @($results | Where-Object { $_.Status -eq 'Failed' }).Count

Write-Host ""
Write-Host "  ------------------------------------------------" -ForegroundColor Cyan
Write-Host "   Added   : $added$(if (-not $Apply) { ' (dry run)' })" -ForegroundColor Cyan
if ($Mirror) { Write-Host "   Removed : $removed$(if (-not $Apply) { ' (dry run)' })" -ForegroundColor Cyan }
Write-Host "   Skipped : $skipped (already member)" -ForegroundColor Cyan
Write-Host "   Failed  : $failed" -ForegroundColor $(if ($failed) { 'Red' } else { 'Cyan' })
Write-Host "   Report  : $OutputPath" -ForegroundColor Green
if (-not $Apply) {
    Write-Host ""
    Write-Host "   Dry run — re-run with -Apply to write changes." -ForegroundColor Yellow
}
Write-Host ""

# ── Session ───────────────────────────────────────────────────────────────────
# Deliberately NOT calling Disconnect-MgGraph: it clears the SDK token cache, so
# every following run would trigger a fresh browser prompt. Use -Disconnect if
# you really want the session torn down (e.g. on a shared machine).
if ($Disconnect -and $script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
