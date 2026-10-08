#Requires -Version 7.0
<#
.SYNOPSIS
    Copy a Planner plan's buckets and tasks (with descriptions and checklists)
    into a new plan on a different group.

.DESCRIPTION
    Reads every bucket and task from a source Planner plan and recreates them on
    a new plan created under -DestinationGroupId, preserving bucket order, task
    titles, due/start dates, descriptions, and checklist items. Consolidates two
    old ad hoc variants of this pattern (one via raw Graph REST calls, one via
    PnP.PowerShell) into a single script on the Planner Graph API, so no extra
    PnP dependency is required.

    Sign-in goes through scripts\Startup\Connect-M365.ps1: delegated as the admin by
    default (device code / GDAP customer per load.config.ps1), app-only with -ClientId
    and -CertificateThumbprint or -AppOnly (Planner supports app-only with the
    Tasks.ReadWrite.All application permission). A Graph session for the right tenant
    that already has the scopes is reused and left connected; only a session this
    script opened is disconnected. Delegated: the signed-in admin must be a member of
    both groups to see and create their plans.

    Defaults to a safe preview — pass -Apply to actually create the new plan.

.PARAMETER SourcePlanId
    ID of the Planner plan to copy from.

.PARAMETER DestinationGroupId
    Object ID of the Microsoft 365 group (Team) the new plan should be created
    under.

.PARAMETER NewPlanTitle
    Title for the new plan. Defaults to the source plan's title.

.PARAMETER Apply
    Actually create the plan and its contents. Without this switch, the script
    only reports what it would copy.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Defaults to the GDAP customer when authMode is GDAP.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint and -TenantId).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.EXAMPLE
    # Preview
    .\Copy-PlannerPlan.ps1 -SourcePlanId "xqQg5FS2LkCp935s-FIFm2QAFkHM" -DestinationGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

.EXAMPLE
    .\Copy-PlannerPlan.ps1 -SourcePlanId "xqQg5FS2LkCp935s-FIFm2QAFkHM" -DestinationGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Apply

.NOTES
    Required module: Microsoft.Graph.Authentication
    Delegated scopes: Tasks.ReadWrite, Group.ReadWrite.All
    Application permissions: Tasks.ReadWrite.All (or Group.ReadWrite.All)
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $SourcePlanId,

    [Parameter(Mandatory)]
    [string] $DestinationGroupId,

    [string] $NewPlanTitle,
    [switch] $Apply,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

# ── Connection ────────────────────────────────────────────────────────────────
$graph = Connect-M365Graph -Scopes 'Tasks.ReadWrite', 'Group.ReadWrite.All' -TenantId $TenantId `
    -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

function Invoke-Graph {
    param([string] $Method = 'GET', [string] $Uri, [string] $Body)
    $params = @{ Method = $Method; Uri = $Uri; ErrorAction = 'Stop' }
    if ($Body) { $params['Body'] = $Body; $params['ContentType'] = 'application/json' }
    Invoke-MgGraphRequest @params
}

function Get-GraphAll {
    # Follows @odata.nextLink; a plan with more tasks than one page lost the rest.
    param([string] $Uri)
    $next = $Uri
    while ($next) {
        $resp = Invoke-Graph -Uri $next
        $resp.value
        $next = $resp.'@odata.nextLink'
    }
}

# ── Resolve source plan ────────────────────────────────────────────────────────
$plan = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/planner/plans/$SourcePlanId"
if (-not $NewPlanTitle) { $NewPlanTitle = $plan.title }

$buckets = @(Get-GraphAll -Uri "https://graph.microsoft.com/v1.0/planner/plans/$SourcePlanId/buckets")
$tasks   = @(Get-GraphAll -Uri "https://graph.microsoft.com/v1.0/planner/plans/$SourcePlanId/tasks")

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Copy-PlannerPlan" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Source plan  : $($plan.title) ($SourcePlanId)"
Write-Host "  New plan     : $NewPlanTitle (group: $DestinationGroupId)"
Write-Host "  Buckets      : $($buckets.Count)"
Write-Host "  Tasks        : $($tasks.Count)"
Write-Host ("  Mode         : {0}" -f $(if ($Apply) { 'Apply' } else { 'Preview only' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

if (-not $Apply) {
    Write-Host "  Would create plan '$NewPlanTitle' with $($buckets.Count) bucket(s) and $($tasks.Count) task(s)." -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to perform the copy." -ForegroundColor Yellow
    Disconnect-M365Graph $graph
    return
}

if (-not $PSCmdlet.ShouldProcess($NewPlanTitle, "Create copy of plan '$($plan.title)'")) {
    Disconnect-M365Graph $graph
    return
}

# ── Create new plan ────────────────────────────────────────────────────────────
$newPlanBody = @{ owner = $DestinationGroupId; title = $NewPlanTitle } | ConvertTo-Json
$newPlan = Invoke-Graph -Method POST -Uri 'https://graph.microsoft.com/v1.0/planner/plans' -Body $newPlanBody
Write-Host "  [OK]   Created plan '$($newPlan.title)' ($($newPlan.id))" -ForegroundColor Green

# ── Copy buckets (preserve order) ──────────────────────────────────────────────
$bucketMap = @{}
foreach ($bucket in ($buckets | Sort-Object orderHint)) {
    $body = @{ name = $bucket.name; planId = $newPlan.id } | ConvertTo-Json
    $newBucket = Invoke-Graph -Method POST -Uri 'https://graph.microsoft.com/v1.0/planner/buckets' -Body $body
    $bucketMap[$bucket.id] = $newBucket.id
    Write-Host "  [OK]   Bucket: $($bucket.name)" -ForegroundColor Green
}

# ── Copy tasks (with details + checklist) ──────────────────────────────────────
$taskCount = 0
foreach ($task in ($tasks | Sort-Object orderHint)) {
    if (-not $bucketMap.ContainsKey($task.bucketId)) { continue }

    $taskBody = @{
        planId   = $newPlan.id
        bucketId = $bucketMap[$task.bucketId]
        title    = $task.title
    }
    if ($task.dueDateTime) { $taskBody['dueDateTime'] = $task.dueDateTime }
    if ($task.startDateTime) { $taskBody['startDateTime'] = $task.startDateTime }

    try {
        $newTask = Invoke-Graph -Method POST -Uri 'https://graph.microsoft.com/v1.0/planner/tasks' -Body ($taskBody | ConvertTo-Json)
        $taskCount++

        # Also for tasks with only a checklist (no description), which were skipped.
        if ($task.hasDescription -or $task.checklistItemCount -gt 0) {
            Start-Sleep -Milliseconds 300
            $taskDetails    = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/planner/tasks/$($task.id)/details"
            $newTaskDetails = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/planner/tasks/$($newTask.id)/details"

            $detailBody = @{ description = $taskDetails.description }
            # checklist is a hashtable keyed by item id; its .PSObject.Properties are
            # the hashtable's own (Count, Keys, ...), not the items.
            if ($taskDetails.checklist -and $taskDetails.checklist.Count -gt 0) {
                $checklist = @{}
                foreach ($item in $taskDetails.checklist.Values) {
                    $checklist[[guid]::NewGuid().Guid] = @{
                        '@odata.type' = '#microsoft.graph.plannerChecklistItem'
                        title         = $item.title
                        isChecked     = $item.isChecked
                    }
                }
                $detailBody['checklist'] = $checklist
            }

            Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/planner/tasks/$($newTask.id)/details" `
                -Body ($detailBody | ConvertTo-Json -Depth 6) -ContentType 'application/json' `
                -Headers @{ 'if-match' = $newTaskDetails.'@odata.etag' } -ErrorAction Stop | Out-Null
        }
        Write-Host "  [OK]   Task: $($task.title)" -ForegroundColor Green
    } catch {
        Write-Host "  [WARN] Task '$($task.title)': $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "  Copied $($bucketMap.Count) bucket(s) and $taskCount task(s) to '$($newPlan.title)'." -ForegroundColor Cyan
Write-Host ""

# ── Disconnect only if this script connected ──────────────────────────────────
Disconnect-M365Graph $graph
