#Requires -Version 5.1
<#
.SYNOPSIS
    Audit distribution group permissions — managers, Send As, Send on Behalf, and settings.

.DESCRIPTION
    Connects to Exchange Online and reports on distribution groups and mail-enabled
    security groups:
      - ManagedBy (group owners/managers)
      - Send As       (Get-RecipientPermission on the group)
      - Send on Behalf (GrantSendOnBehalfTo)
      - Sender restrictions (RequireSenderAuthenticationEnabled, AcceptMessagesOnlyFrom)
      - Member count, join/leave restrictions

    Results are exported to CSV. Use -IncludeMembers to also list each member per group.

.PARAMETER Group
    Identity of a single group (name, alias, or email). If omitted, all distribution
    groups and mail-enabled security groups are audited.

.PARAMETER IncludeMembers
    Also include individual group members in the report.

.PARAMETER OutputPath
    Path to write the CSV report. Defaults to .\GroupPermissions_<timestamp>.csv.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Test-DistributionGroupPermissions.ps1

.EXAMPLE
    .\Test-DistributionGroupPermissions.ps1 -Group "helpdesk@contoso.com"

.EXAMPLE
    .\Test-DistributionGroupPermissions.ps1 -IncludeMembers -OutputPath "C:\Reports\groups.csv"
#>
[CmdletBinding()]
param(
    [string] $Group,
    [switch] $IncludeMembers,
    [string] $OutputPath,
    [string] $TenantId
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
} catch {
    $connectParams = @{ ShowBanner = $false }
    if ($TenantId) { $connectParams['Organization'] = $TenantId }
    Connect-ExchangeOnline @connectParams
    $script:ConnectedHere = $true
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Distribution Group Permissions Audit" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# ── Get groups ────────────────────────────────────────────────────────────────
if ($Group) {
    $groups = @(Get-DistributionGroup -Identity $Group -ErrorAction Stop)
} else {
    Write-Host "  Retrieving distribution groups..." -ForegroundColor DarkGray
    $groups = @(Get-DistributionGroup -ResultSize Unlimited)
}

Write-Host "  Auditing $($groups.Count) group(s)..." -ForegroundColor DarkGray
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($grp in $groups) {
    $grpEmail = $grp.PrimarySmtpAddress

    # Group settings row
    $memberCount = (Get-DistributionGroupMember -Identity $grp.Identity -ResultSize Unlimited).Count
    $results.Add([PSCustomObject]@{
        Group          = $grpEmail
        GroupName      = $grp.DisplayName
        RecordType     = 'Settings'
        Detail         = "Members:$memberCount | JoinRestriction:$($grp.MemberJoinRestriction) | LeaveRestriction:$($grp.MemberDepartRestriction) | ExternalSenders:$(-not $grp.RequireSenderAuthenticationEnabled)"
        GrantedTo      = ''
    })

    # ManagedBy
    foreach ($manager in $grp.ManagedBy) {
        $results.Add([PSCustomObject]@{
            Group      = $grpEmail
            GroupName  = $grp.DisplayName
            RecordType = 'ManagedBy'
            Detail     = ''
            GrantedTo  = $manager
        })
    }

    # Send on Behalf
    foreach ($delegate in $grp.GrantSendOnBehalfTo) {
        $results.Add([PSCustomObject]@{
            Group      = $grpEmail
            GroupName  = $grp.DisplayName
            RecordType = 'SendOnBehalf'
            Detail     = ''
            GrantedTo  = $delegate
        })
    }

    # Send As
    try {
        Get-RecipientPermission -Identity $grpEmail -ErrorAction SilentlyContinue |
            Where-Object { $_.Trustee -notlike '*SELF*' -and $_.Trustee -notlike 'NT AUTHORITY*' } |
            ForEach-Object {
                $results.Add([PSCustomObject]@{
                    Group      = $grpEmail
                    GroupName  = $grp.DisplayName
                    RecordType = 'SendAs'
                    Detail     = ($_.AccessRights -join ', ')
                    GrantedTo  = $_.Trustee
                })
            }
    } catch {
        Write-Host "  [WARN] SendAs — $grpEmail : $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Members (optional)
    if ($IncludeMembers) {
        try {
            Get-DistributionGroupMember -Identity $grp.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue |
                ForEach-Object {
                    $results.Add([PSCustomObject]@{
                        Group      = $grpEmail
                        GroupName  = $grp.DisplayName
                        RecordType = 'Member'
                        Detail     = $_.RecipientType
                        GrantedTo  = $_.PrimarySmtpAddress
                    })
                }
        } catch {
            Write-Host "  [WARN] Members — $grpEmail : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

# ── Output ────────────────────────────────────────────────────────────────────
if ($results.Count -eq 0) {
    Write-Host "  No groups or permissions found." -ForegroundColor DarkGray
} else {
    $results | Format-Table -AutoSize

    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "GroupPermissions_$ts.csv"
    }

    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "  Audited $($groups.Count) group(s) — $($results.Count) records written." -ForegroundColor Cyan
Write-Host ""

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
