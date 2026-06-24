#Requires -Version 7.0
#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Resolves members of a dynamic distribution group and copies them into a regular distribution group.

.DESCRIPTION
    This script does three things:
    1) Resolves all recipients currently matching the dynamic distribution group filter.
    2) Exports the resolved member list to CSV.
    3) Creates or updates a regular distribution group and adds those members.

    The script requires an active Exchange Online session.
    Connect first with: Connect-ExchangeOnline

.PARAMETER DynamicGroupIdentity
    Identity of the dynamic distribution group (name, alias, DN, or SMTP address).

.PARAMETER TargetGroupIdentity
    Identity of the target regular distribution group.
    If the group does not exist, it will be created.

.PARAMETER TargetDisplayName
    Display name for a new target group. Default: <DynamicDisplayName> Static

.PARAMETER TargetAlias
    Alias for a new target group. Default: <DynamicAlias>-static

.PARAMETER TargetPrimarySmtpAddress
    Primary SMTP address for a new target group.
    If omitted, Exchange Online policy can generate one.

.PARAMETER ClearTargetMembers
    Removes all existing target members before adding the resolved dynamic members.

.PARAMETER ExportCsvPath
    CSV export path for resolved members.

.PARAMETER SkipMemberAdd
    Only resolve and export members, do not modify the target group.

.PARAMETER RenameDynamicGroupTo
    Optional new name/display name for the source dynamic distribution group after processing.

.EXAMPLE
    .\Distributionlist.ps1 -DynamicGroupIdentity "All Sales" -TargetGroupIdentity "All Sales Static"

.EXAMPLE
    .\Distributionlist.ps1 -DynamicGroupIdentity sales@contoso.com -TargetGroupIdentity sales-static@contoso.com -ClearTargetMembers

.EXAMPLE
    .\Distributionlist.ps1 -DynamicGroupIdentity "All Staff" -TargetGroupIdentity "All Staff Static" -WhatIf

.EXAMPLE
    .\Distributionlist.ps1 -DynamicGroupIdentity "All Sales" -TargetGroupIdentity "All Sales Static" -RenameDynamicGroupTo "All Sales (Legacy Dynamic)"
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$DynamicGroupIdentity,

    [Parameter(Mandatory)]
    [string]$TargetGroupIdentity,

    [string]$TargetDisplayName,
    [string]$TargetAlias,
    [string]$TargetPrimarySmtpAddress,
    [switch]$ClearTargetMembers,
    [string]$ExportCsvPath,
    [switch]$SkipMemberAdd,
    [string]$RenameDynamicGroupTo
)

function Assert-ExchangeConnection {
    try {
        Get-OrganizationConfig -ErrorAction Stop | Out-Null
    }
    catch {
        throw "No active Exchange Online session found. Run Connect-ExchangeOnline first."
    }
}

function Get-RecipientKey {
    param([Parameter(Mandatory)]$Recipient)

    if ($Recipient.ExternalDirectoryObjectId) {
        return "id:$($Recipient.ExternalDirectoryObjectId)"
    }
    if ($Recipient.PrimarySmtpAddress) {
        return "smtp:$($Recipient.PrimarySmtpAddress.ToString().ToLowerInvariant())"
    }
    return "identity:$($Recipient.Identity.ToString().ToLowerInvariant())"
}

function New-SafeAlias {
    param([Parameter(Mandatory)][string]$RawAlias)

    $candidate = $RawAlias.ToLowerInvariant() -replace '[^a-z0-9._-]', '-'
    $candidate = $candidate -replace '-{2,}', '-'
    $candidate = $candidate.Trim('-')

    if ([string]::IsNullOrWhiteSpace($candidate)) {
        return "ddg-static"
    }

    if ($candidate.Length -gt 64) {
        return $candidate.Substring(0, 64)
    }

    return $candidate
}

Assert-ExchangeConnection

Write-Host "Resolving dynamic distribution group: $DynamicGroupIdentity" -ForegroundColor Cyan
$dynamicGroup = Get-DynamicDistributionGroup -Identity $DynamicGroupIdentity -ErrorAction Stop

$resolvedMembers = Get-Recipient -ResultSize Unlimited -RecipientPreviewFilter $dynamicGroup.RecipientFilter |
    Sort-Object -Property PrimarySmtpAddress -Unique

if (-not $ExportCsvPath) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $ExportCsvPath = Join-Path -Path (Get-Location) -ChildPath "DynamicGroupMembers_$timestamp.csv"
}

$resolvedMembers |
    Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails, Identity |
    Export-Csv -Path $ExportCsvPath -NoTypeInformation -Encoding UTF8

Write-Host "Resolved members: $($resolvedMembers.Count)" -ForegroundColor Green
Write-Host "CSV export: $ExportCsvPath" -ForegroundColor Green

if ($SkipMemberAdd) {
    Write-Host "SkipMemberAdd enabled, no target group changes applied." -ForegroundColor Yellow
    if (-not [string]::IsNullOrWhiteSpace($RenameDynamicGroupTo)) {
        if ($PSCmdlet.ShouldProcess($dynamicGroup.Identity, "Rename dynamic distribution group to $RenameDynamicGroupTo")) {
            Set-DynamicDistributionGroup -Identity $dynamicGroup.Identity -Name $RenameDynamicGroupTo -DisplayName $RenameDynamicGroupTo -ErrorAction Stop
            Write-Host "Renamed dynamic group to: $RenameDynamicGroupTo" -ForegroundColor Green
        }
    }
    return
}

$targetGroup = Get-DistributionGroup -Identity $TargetGroupIdentity -ErrorAction SilentlyContinue

if (-not $targetGroup) {
    if (-not $TargetDisplayName) {
        $TargetDisplayName = "$($dynamicGroup.DisplayName) Static"
    }

    if (-not $TargetAlias) {
        $TargetAlias = New-SafeAlias -RawAlias "$($dynamicGroup.Alias)-static"
    }

    $newGroupParams = @{
        Name        = $TargetDisplayName
        DisplayName = $TargetDisplayName
        Alias       = $TargetAlias
        Type        = "Distribution"
    }

    if ($TargetPrimarySmtpAddress) {
        $newGroupParams.PrimarySmtpAddress = $TargetPrimarySmtpAddress
    }

    if ($PSCmdlet.ShouldProcess($TargetGroupIdentity, "Create regular distribution group")) {
        New-DistributionGroup @newGroupParams -ErrorAction Stop | Out-Null
    }

    $targetGroup = Get-DistributionGroup -Identity $TargetGroupIdentity -ErrorAction SilentlyContinue
    if (-not $targetGroup -and $TargetAlias) {
        $targetGroup = Get-DistributionGroup -Identity $TargetAlias -ErrorAction Stop
    }

    Write-Host "Created target group: $($targetGroup.DisplayName)" -ForegroundColor Green
}
else {
    Write-Host "Using existing target group: $($targetGroup.DisplayName)" -ForegroundColor Cyan
}

$existingMembers = @()
try {
    $existingMembers = Get-DistributionGroupMember -Identity $targetGroup.Identity -ResultSize Unlimited -ErrorAction Stop
}
catch {
    Write-Verbose "Could not read current members on target group: $($_.Exception.Message)"
}

if ($ClearTargetMembers -and $existingMembers.Count -gt 0) {
    foreach ($member in $existingMembers) {
        if ($PSCmdlet.ShouldProcess($targetGroup.Identity, "Remove member $($member.Identity)")) {
            Remove-DistributionGroupMember -Identity $targetGroup.Identity -Member $member.Identity -BypassSecurityGroupManagerCheck -Confirm:$false
        }
    }
    $existingMembers = @()
}

$existingKeys = @{}
foreach ($member in $existingMembers) {
    $existingKeys[(Get-RecipientKey -Recipient $member)] = $true
}

$added = 0
$skipped = 0
$failed = 0

foreach ($member in $resolvedMembers) {
    $key = Get-RecipientKey -Recipient $member

    if ($existingKeys.ContainsKey($key)) {
        $skipped++
        continue
    }

    try {
        if ($PSCmdlet.ShouldProcess($targetGroup.Identity, "Add member $($member.Identity)")) {
            Add-DistributionGroupMember -Identity $targetGroup.Identity -Member $member.Identity -BypassSecurityGroupManagerCheck -ErrorAction Stop
        }
        $added++
    }
    catch {
        $failed++
        Write-Warning "Failed to add $($member.Identity): $($_.Exception.Message)"
    }
}

if (-not [string]::IsNullOrWhiteSpace($RenameDynamicGroupTo)) {
    if ($PSCmdlet.ShouldProcess($dynamicGroup.Identity, "Rename dynamic distribution group to $RenameDynamicGroupTo")) {
        Set-DynamicDistributionGroup -Identity $dynamicGroup.Identity -Name $RenameDynamicGroupTo -DisplayName $RenameDynamicGroupTo -ErrorAction Stop
        Write-Host "Renamed dynamic group to: $RenameDynamicGroupTo" -ForegroundColor Green
    }
}

Write-Host "Completed." -ForegroundColor Green
Write-Host "Added   : $added" -ForegroundColor Green
Write-Host "Skipped : $skipped" -ForegroundColor Yellow
Write-Host "Failed  : $failed" -ForegroundColor Red
