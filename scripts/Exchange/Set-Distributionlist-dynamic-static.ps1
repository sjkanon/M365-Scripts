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
    If omitted, the script uses the current primary SMTP of the dynamic group.

.PARAMETER CopyManagersFromDynamic
    Copies the ManagedBy owners from the dynamic group to the target group.
    Enabled by default.

.PARAMETER DisableCopyManagersFromDynamic
    Disables copying the ManagedBy owners from dynamic to target.

.PARAMETER MakeDynamicAddressTemporary
    When the target group should use the dynamic group's current primary SMTP,
    the dynamic group gets a temporary primary SMTP first.
    Enabled by default.

.PARAMETER DisableMakeDynamicAddressTemporary
    Disables the automatic temporary SMTP change on the dynamic group.

.PARAMETER ClearTargetMembers
    Removes all existing target members before adding the resolved dynamic members.

.PARAMETER ExportCsvPath
    CSV export path for resolved members.
    Default: C:\Temp\DynamicGroupMembers_<timestamp>.csv (Windows)
             ~/Downloads/DynamicGroupMembers_<timestamp>.csv (Linux/macOS)

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

.EXAMPLE
    .\Distributionlist.ps1 -DynamicGroupIdentity "All Sales" -TargetGroupIdentity "All Sales Static" -MakeDynamicAddressTemporary
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
    [switch]$CopyManagersFromDynamic,
    [switch]$DisableCopyManagersFromDynamic,
    [switch]$MakeDynamicAddressTemporary,
    [switch]$DisableMakeDynamicAddressTemporary,
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

function Get-UniqueRecipientIdentity {
    param([Parameter(Mandatory)]$Recipient)

    if ($Recipient.ExternalDirectoryObjectId) {
        return $Recipient.ExternalDirectoryObjectId.ToString()
    }
    if ($Recipient.PrimarySmtpAddress) {
        return $Recipient.PrimarySmtpAddress.ToString()
    }
    if ($Recipient.Guid) {
        return $Recipient.Guid.ToString()
    }
    if ($Recipient.DistinguishedName) {
        return $Recipient.DistinguishedName
    }

    return $Recipient.Identity
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

function New-TemporaryPrimarySmtpAddress {
    param(
        [Parameter(Mandatory)][string]$CurrentPrimarySmtpAddress,
        [string]$AliasFallback
    )

    $current = $CurrentPrimarySmtpAddress.Trim().ToLowerInvariant()
    if ($current -notmatch '^[^@]+@[^@]+$') {
        throw "Current primary SMTP address '$CurrentPrimarySmtpAddress' is not in a valid mail format."
    }

    $parts = $current.Split('@', 2)
    $localPart = if ([string]::IsNullOrWhiteSpace($AliasFallback)) { $parts[0] } else { $AliasFallback }
    $stamp = Get-Date -Format "yyyyMMddHHmmss"
    $tempLocal = New-SafeAlias -RawAlias "$localPart-temp-$stamp"

    return "$tempLocal@$($parts[1])"
}

function Escape-RecipientFilterValue {
    param([Parameter(Mandatory)][string]$Value)

    return $Value.Replace("'", "''")
}

function Get-RecipientConflictSummary {
    param([Parameter(Mandatory)]$Recipient)

    $display = if ($Recipient.DisplayName) { $Recipient.DisplayName } else { $Recipient.Name }
    $smtp = if ($Recipient.PrimarySmtpAddress) { $Recipient.PrimarySmtpAddress.ToString() } else { '<none>' }
    return "$display [$($Recipient.RecipientTypeDetails)] <$smtp>"
}

function Get-UniqueGroupIdentity {
    param([Parameter(Mandatory)]$Group)

    if ($Group.DistinguishedName) {
        return $Group.DistinguishedName
    }
    if ($Group.Guid) {
        return $Group.Guid.ToString()
    }
    if ($Group.PrimarySmtpAddress) {
        return $Group.PrimarySmtpAddress.ToString()
    }

    return $Group.Identity
}

function Assert-TargetGroupCreationInputs {
    param(
        [Parameter(Mandatory)][string]$TargetDisplayName,
        [Parameter(Mandatory)][string]$TargetAlias,
        [string]$TargetPrimarySmtpAddress,
        [string]$DynamicGroupPrimarySmtpAddress
    )

    $safeDisplayName = Escape-RecipientFilterValue -Value $TargetDisplayName
    $nameConflicts = @(Get-Recipient -ResultSize Unlimited -Filter "Name -eq '$safeDisplayName'" -ErrorAction SilentlyContinue)
    if ($nameConflicts.Count -gt 0) {
        $nameConflictList = $nameConflicts | ForEach-Object { Get-RecipientConflictSummary -Recipient $_ }
        throw "Target name '$TargetDisplayName' is already in use by: $($nameConflictList -join '; '). Use -TargetDisplayName with a unique value."
    }

    if (-not [string]::IsNullOrWhiteSpace($TargetAlias)) {
        $safeAlias = Escape-RecipientFilterValue -Value $TargetAlias
        $aliasConflicts = @(Get-Recipient -ResultSize Unlimited -Filter "Alias -eq '$safeAlias'" -ErrorAction SilentlyContinue)
        if ($aliasConflicts.Count -gt 0) {
            $aliasConflictList = $aliasConflicts | ForEach-Object { Get-RecipientConflictSummary -Recipient $_ }
            throw "Target alias '$TargetAlias' is already in use by: $($aliasConflictList -join '; '). Use -TargetAlias with a unique value."
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($TargetPrimarySmtpAddress)) {
        $smtpConflict = Get-Recipient -Identity $TargetPrimarySmtpAddress -ErrorAction SilentlyContinue
        if ($smtpConflict) {
            $smtpConflictText = Get-RecipientConflictSummary -Recipient $smtpConflict
            if (
                -not [string]::IsNullOrWhiteSpace($DynamicGroupPrimarySmtpAddress) -and
                $DynamicGroupPrimarySmtpAddress.Equals($TargetPrimarySmtpAddress, [System.StringComparison]::OrdinalIgnoreCase)
            ) {
                throw "Target primary SMTP '$TargetPrimarySmtpAddress' is still in use by: $smtpConflictText. The dynamic group address was not freed. Re-run with default temporary-address behavior or pass -MakeDynamicAddressTemporary."
            }

            throw "Target primary SMTP '$TargetPrimarySmtpAddress' is already in use by: $smtpConflictText. Use -TargetPrimarySmtpAddress with a unique address."
        }
    }
}

Assert-ExchangeConnection

Write-Host "Resolving dynamic distribution group: $DynamicGroupIdentity" -ForegroundColor Cyan
$dynamicGroup = Get-DynamicDistributionGroup -Identity $DynamicGroupIdentity -ErrorAction Stop
$dynamicGroupCmdletIdentity = Get-UniqueGroupIdentity -Group $dynamicGroup
$originalDynamicDisplayName = $dynamicGroup.DisplayName
$dynamicPrimarySmtpAddress = if ($dynamicGroup.PrimarySmtpAddress) { $dynamicGroup.PrimarySmtpAddress.ToString() } else { $null }

$shouldCopyManagers = $true
if ($CopyManagersFromDynamic) { $shouldCopyManagers = $true }
if ($DisableCopyManagersFromDynamic) { $shouldCopyManagers = $false }

$shouldMakeDynamicAddressTemporary = $true
if ($MakeDynamicAddressTemporary) { $shouldMakeDynamicAddressTemporary = $true }
if ($DisableMakeDynamicAddressTemporary) { $shouldMakeDynamicAddressTemporary = $false }

if (-not $TargetPrimarySmtpAddress -and $dynamicPrimarySmtpAddress) {
    $TargetPrimarySmtpAddress = $dynamicPrimarySmtpAddress
    Write-Host "TargetPrimarySmtpAddress not provided; using dynamic group's current primary SMTP: $TargetPrimarySmtpAddress" -ForegroundColor Cyan
}

if (
    $shouldMakeDynamicAddressTemporary -and
    $dynamicPrimarySmtpAddress -and
    $TargetPrimarySmtpAddress -and
    $dynamicPrimarySmtpAddress.Equals($TargetPrimarySmtpAddress, [System.StringComparison]::OrdinalIgnoreCase)
) {
    $temporaryDynamicAddress = New-TemporaryPrimarySmtpAddress -CurrentPrimarySmtpAddress $dynamicPrimarySmtpAddress -AliasFallback $dynamicGroup.Alias
    if ($PSCmdlet.ShouldProcess($dynamicGroup.DisplayName, "Set temporary primary SMTP on dynamic group: $temporaryDynamicAddress")) {
        Set-DynamicDistributionGroup -Identity $dynamicGroupCmdletIdentity -PrimarySmtpAddress $temporaryDynamicAddress -ErrorAction Stop
        Write-Host "Dynamic group primary SMTP changed to temporary address: $temporaryDynamicAddress" -ForegroundColor Green
    }

    $dynamicGroup = Get-DynamicDistributionGroup -Identity $dynamicGroupCmdletIdentity -ErrorAction Stop
    $dynamicAddressesToRemove = @()
    foreach ($address in @($dynamicGroup.EmailAddresses)) {
        $addressText = $address.ToString()
        $addressSmtp = ($addressText -replace '^(SMTP|smtp):', '')
        if ($addressSmtp.Equals($TargetPrimarySmtpAddress, [System.StringComparison]::OrdinalIgnoreCase)) {
            $dynamicAddressesToRemove += $addressText
        }
    }

    if ($dynamicAddressesToRemove.Count -gt 0) {
        if ($PSCmdlet.ShouldProcess($dynamicGroup.DisplayName, "Remove old target SMTP from dynamic group: $TargetPrimarySmtpAddress")) {
            Set-DynamicDistributionGroup -Identity $dynamicGroupCmdletIdentity -EmailAddresses @{ Remove = $dynamicAddressesToRemove } -ErrorAction Stop
            Write-Host "Removed old SMTP from dynamic group so target can use it: $TargetPrimarySmtpAddress" -ForegroundColor Green
        }
    }

    $dynamicGroup = Get-DynamicDistributionGroup -Identity $dynamicGroupCmdletIdentity -ErrorAction Stop
    $dynamicPrimarySmtpAddress = if ($dynamicGroup.PrimarySmtpAddress) { $dynamicGroup.PrimarySmtpAddress.ToString() } else { $null }
}

if (
    $shouldMakeDynamicAddressTemporary -and
    $TargetPrimarySmtpAddress
) {
    $dynamicAddressesToRemove = @()
    foreach ($address in @($dynamicGroup.EmailAddresses)) {
        $addressText = $address.ToString()
        $addressSmtp = ($addressText -replace '^(SMTP|smtp):', '')
        if ($addressSmtp.Equals($TargetPrimarySmtpAddress, [System.StringComparison]::OrdinalIgnoreCase)) {
            $dynamicAddressesToRemove += $addressText
        }
    }

    if ($dynamicAddressesToRemove.Count -gt 0) {
        if ($PSCmdlet.ShouldProcess($dynamicGroup.DisplayName, "Ensure target SMTP is freed on dynamic group: $TargetPrimarySmtpAddress")) {
            Set-DynamicDistributionGroup -Identity $dynamicGroupCmdletIdentity -EmailAddresses @{ Remove = $dynamicAddressesToRemove } -ErrorAction Stop
            Write-Host "Removed target SMTP from dynamic group address collection: $TargetPrimarySmtpAddress" -ForegroundColor Green
        }
        $dynamicGroup = Get-DynamicDistributionGroup -Identity $dynamicGroupCmdletIdentity -ErrorAction Stop
        $dynamicPrimarySmtpAddress = if ($dynamicGroup.PrimarySmtpAddress) { $dynamicGroup.PrimarySmtpAddress.ToString() } else { $null }
    }
}

$resolvedMembers = Get-Recipient -ResultSize Unlimited -RecipientPreviewFilter $dynamicGroup.RecipientFilter |
    Sort-Object -Property PrimarySmtpAddress -Unique

if (-not $ExportCsvPath) {
    $outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
    if (-not (Test-Path -LiteralPath $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $ExportCsvPath = Join-Path -Path $outputDir -ChildPath "DynamicGroupMembers_$timestamp.csv"
}

$resolvedMembers |
    Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails, Identity |
    Export-Csv -Path $ExportCsvPath -NoTypeInformation -Encoding UTF8

Write-Host "Resolved members: $($resolvedMembers.Count)" -ForegroundColor Green
Write-Host "CSV export: $ExportCsvPath" -ForegroundColor Green

if ($SkipMemberAdd) {
    Write-Host "SkipMemberAdd enabled, no target group changes applied." -ForegroundColor Yellow
    if (-not [string]::IsNullOrWhiteSpace($RenameDynamicGroupTo)) {
        if ($PSCmdlet.ShouldProcess($dynamicGroup.DisplayName, "Rename dynamic distribution group to $RenameDynamicGroupTo")) {
            Set-DynamicDistributionGroup -Identity $dynamicGroupCmdletIdentity -Name $RenameDynamicGroupTo -DisplayName $RenameDynamicGroupTo -ErrorAction Stop
            Write-Host "Renamed dynamic group to: $RenameDynamicGroupTo" -ForegroundColor Green
        }
    }
    return
}

$renamedDynamicGroup = $false
if (-not [string]::IsNullOrWhiteSpace($RenameDynamicGroupTo)) {
    if ($PSCmdlet.ShouldProcess($dynamicGroup.DisplayName, "Rename dynamic distribution group to $RenameDynamicGroupTo")) {
        Set-DynamicDistributionGroup -Identity $dynamicGroupCmdletIdentity -Name $RenameDynamicGroupTo -DisplayName $RenameDynamicGroupTo -ErrorAction Stop
        $renamedDynamicGroup = $true
        Write-Host "Renamed dynamic group to: $RenameDynamicGroupTo" -ForegroundColor Green
    }
}

$targetGroup = Get-DistributionGroup -Identity $TargetGroupIdentity -ErrorAction SilentlyContinue
$candidateTargetAlias = $TargetAlias
if (-not $candidateTargetAlias) {
    $candidateTargetAlias = New-SafeAlias -RawAlias "$($dynamicGroup.Alias)-static"
}

if (-not $targetGroup -and -not [string]::IsNullOrWhiteSpace($TargetPrimarySmtpAddress)) {
    $targetGroup = Get-DistributionGroup -Identity $TargetPrimarySmtpAddress -ErrorAction SilentlyContinue
    if ($targetGroup) {
        Write-Host "TargetGroupIdentity not found; using existing group by primary SMTP: $($targetGroup.DisplayName)" -ForegroundColor Cyan
    }
}

if (-not $targetGroup -and -not [string]::IsNullOrWhiteSpace($candidateTargetAlias)) {
    $targetGroup = Get-DistributionGroup -Identity $candidateTargetAlias -ErrorAction SilentlyContinue
    if ($targetGroup) {
        Write-Host "TargetGroupIdentity not found; using existing group by alias: $($targetGroup.DisplayName)" -ForegroundColor Cyan
    }
}

if (-not $targetGroup) {
    if (-not $TargetDisplayName) {
        if ($renamedDynamicGroup) {
            if ($TargetGroupIdentity -match '@') {
                $TargetDisplayName = ($originalDynamicDisplayName -replace '\s*\(Legacy\)\s*$', '').Trim()
                if ([string]::IsNullOrWhiteSpace($TargetDisplayName)) {
                    $TargetDisplayName = "$($dynamicGroup.DisplayName) Static"
                }
            }
            else {
                $TargetDisplayName = $TargetGroupIdentity
            }
        }
        else {
            $TargetDisplayName = "$($dynamicGroup.DisplayName) Static"
        }
    }

    if (-not $TargetAlias) {
        $TargetAlias = $candidateTargetAlias
    }

    Assert-TargetGroupCreationInputs -TargetDisplayName $TargetDisplayName -TargetAlias $TargetAlias -TargetPrimarySmtpAddress $TargetPrimarySmtpAddress -DynamicGroupPrimarySmtpAddress $dynamicPrimarySmtpAddress

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

$targetGroupCmdletIdentity = Get-UniqueGroupIdentity -Group $targetGroup

if (
    $TargetPrimarySmtpAddress -and
    $targetGroup.PrimarySmtpAddress -and
    -not $targetGroup.PrimarySmtpAddress.ToString().Equals($TargetPrimarySmtpAddress, [System.StringComparison]::OrdinalIgnoreCase)
) {
    if ($PSCmdlet.ShouldProcess($targetGroup.DisplayName, "Set target group primary SMTP to $TargetPrimarySmtpAddress")) {
        Set-DistributionGroup -Identity $targetGroupCmdletIdentity -PrimarySmtpAddress $TargetPrimarySmtpAddress -BypassSecurityGroupManagerCheck -ErrorAction Stop
        Write-Host "Updated target group primary SMTP: $TargetPrimarySmtpAddress" -ForegroundColor Green
    }
}

if ($shouldCopyManagers) {
    $dynamicManagers = @($dynamicGroup.ManagedBy | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($dynamicManagers.Count -gt 0) {
        if ($PSCmdlet.ShouldProcess($targetGroup.DisplayName, "Copy ManagedBy from dynamic group")) {
            Set-DistributionGroup -Identity $targetGroupCmdletIdentity -ManagedBy $dynamicManagers -BypassSecurityGroupManagerCheck -ErrorAction Stop
            Write-Host "Copied $($dynamicManagers.Count) manager(s) to target group." -ForegroundColor Green
        }
    }
    else {
        Write-Host "Dynamic group has no ManagedBy owners to copy." -ForegroundColor Yellow
    }
}
else {
    Write-Host "Skipping manager copy (DisableCopyManagersFromDynamic enabled)." -ForegroundColor Yellow
}

$existingMembers = @()
try {
    $existingMembers = Get-DistributionGroupMember -Identity $targetGroupCmdletIdentity -ResultSize Unlimited -ErrorAction Stop
}
catch {
    Write-Verbose "Could not read current members on target group: $($_.Exception.Message)"
}

if ($ClearTargetMembers -and $existingMembers.Count -gt 0) {
    foreach ($member in $existingMembers) {
        $memberIdentityForRemove = Get-UniqueRecipientIdentity -Recipient $member
        if ($PSCmdlet.ShouldProcess($targetGroup.DisplayName, "Remove member $memberIdentityForRemove")) {
            try {
                Remove-DistributionGroupMember -Identity $targetGroupCmdletIdentity -Member $memberIdentityForRemove -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
            }
            catch {
                Write-Warning "Failed to remove ${memberIdentityForRemove}: $($_.Exception.Message)"
            }
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
    $memberIdentityForAdd = Get-UniqueRecipientIdentity -Recipient $member

    if ($existingKeys.ContainsKey($key)) {
        $skipped++
        continue
    }

    try {
        if ($PSCmdlet.ShouldProcess($targetGroup.DisplayName, "Add member $memberIdentityForAdd")) {
            Add-DistributionGroupMember -Identity $targetGroupCmdletIdentity -Member $memberIdentityForAdd -BypassSecurityGroupManagerCheck -ErrorAction Stop
        }
        $added++
    }
    catch {
        $failed++
        Write-Warning "Failed to add ${memberIdentityForAdd}: $($_.Exception.Message)"
    }
}

Write-Host "Completed." -ForegroundColor Green
Write-Host "Added   : $added" -ForegroundColor Green
Write-Host "Skipped : $skipped" -ForegroundColor Yellow
Write-Host "Failed  : $failed" -ForegroundColor Red
