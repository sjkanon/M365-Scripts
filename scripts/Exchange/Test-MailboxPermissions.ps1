#Requires -Version 5.1
<#
.SYNOPSIS
    Audit mailbox access permissions - Full Access, Send As, and Send on Behalf.

.DESCRIPTION
    Connects to Exchange Online and checks all three delegation types for
    a single mailbox or all mailboxes in the tenant. You can also run it in
    reverse to find all mailboxes where a specific user has permissions:
      - Full Access    (Get-MailboxPermission)
      - Send As        (Get-RecipientPermission)
      - Send on Behalf (GrantSendOnBehalfTo from Get-EXOMailbox)

    Results are displayed in a table and optionally exported to CSV.

.PARAMETER Mailbox
    UPN of the mailbox to audit. If omitted, all user and shared mailboxes are checked.

.PARAMETER User
    UPN, SMTP address, or display name of a user to search for. The script then
    returns all mailboxes where that user has permissions.

.PARAMETER OutputPath
    Path to write the CSV report. Defaults to .\MailboxPermissions_<timestamp>.csv.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Test-MailboxPermissions.ps1

.EXAMPLE
    .\Test-MailboxPermissions.ps1 -Mailbox "shared@contoso.com"

.EXAMPLE
    .\Test-MailboxPermissions.ps1 -User "user@contoso.com"

.EXAMPLE
    .\Test-MailboxPermissions.ps1 -OutputPath "C:\Reports\permissions.csv"
#>
[CmdletBinding()]
param(
    [string] $Mailbox,
    [string] $User,
    [string] $OutputPath,
    [string] $TenantId
)

# -- Output folder ------------------------------------------------------------
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

function Normalize-PrincipalValue {
    param([object] $Value)

    if ($null -eq $Value) { return $null }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    $text = $text.Trim()
    if ($text -match '^(?i)smtp:') {
        $text = $text.Substring(5)
    }

    return $text.ToLowerInvariant()
}

function Get-PrincipalDisplayValue {
    param([object] $Principal)

    if ($null -eq $Principal) { return $null }

    foreach ($propertyName in @('DisplayName', 'PrimarySmtpAddress', 'UserPrincipalName', 'Alias', 'Name', 'Identity')) {
        if ($Principal.PSObject.Properties[$propertyName]) {
            $value = $Principal.$propertyName
            if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                return [string]$value
            }
        }
    }

    return [string]$Principal
}

function Get-PrincipalMatchValues {
    param([object] $Principal)

    $values = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    if ($null -eq $Principal) {
        return @()
    }

    $rawValues = @([string]$Principal)
    foreach ($propertyName in @('DisplayName', 'PrimarySmtpAddress', 'UserPrincipalName', 'Alias', 'Name', 'Identity', 'ExternalEmailAddress')) {
        if ($Principal.PSObject.Properties[$propertyName]) {
            $rawValues += [string]$Principal.$propertyName
        }
    }

    foreach ($value in $rawValues) {
        $normalized = Normalize-PrincipalValue $value
        if ($normalized) {
            $values.Add($normalized) | Out-Null
        }
    }

    return @($values)
}

function Test-PrincipalMatch {
    param(
        [object] $Principal,
        [string[]] $Needles
    )

    if (-not $Needles -or $Needles.Count -eq 0) { return $false }

    foreach ($candidate in (Get-PrincipalMatchValues -Principal $Principal)) {
        if ($Needles -contains $candidate) {
            return $true
        }
    }

    return $false
}

function Test-UserIsInTrusteeGroup {
    param(
        [object] $Trustee,
        [string[]] $Needles
    )

    if (-not $Needles -or $Needles.Count -eq 0) { return $false }

    $trusteeIdentity = Get-PrincipalDisplayValue $Trustee
    if ([string]::IsNullOrWhiteSpace($trusteeIdentity)) { return $false }

    if (-not $script:GroupMembershipCache) {
        $script:GroupMembershipCache = @{}
    }

    if ($script:GroupMembershipCache.ContainsKey($trusteeIdentity)) {
        return [bool]$script:GroupMembershipCache[$trusteeIdentity]
    }

    $isMember = $false

    try {
        $groupRecipient = Get-EXORecipient -Identity $trusteeIdentity -ErrorAction Stop
        $groupType = [string]$groupRecipient.RecipientTypeDetails

        $memberResults = @()
        if ($groupType -eq 'GroupMailbox') {
            try {
                $memberResults = @(Get-UnifiedGroupLinks -Identity $groupRecipient.Identity -LinkType Members -ResultSize Unlimited -ErrorAction Stop)
            } catch {
                $memberResults = @()
            }
        } elseif ($groupType -in @('MailUniversalSecurityGroup', 'MailUniversalDistributionGroup', 'MailNonUniversalGroup')) {
            try {
                $memberResults = @(Get-DistributionGroupMember -Identity $groupRecipient.Identity -ResultSize Unlimited -ErrorAction Stop)
            } catch {
                $memberResults = @()
            }
        }

        foreach ($member in $memberResults) {
            if (Test-PrincipalMatch -Principal $member -Needles $Needles) {
                $isMember = $true
                break
            }
        }
    } catch {
        $isMember = $false
    }

    $script:GroupMembershipCache[$trusteeIdentity] = $isMember
    return $isMember
}

function Test-UserMatchesTrustee {
    param(
        [object] $Trustee,
        [string[]] $Needles
    )

    if (-not $Needles -or $Needles.Count -eq 0) { return $false }

    if (Test-PrincipalMatch -Principal $Trustee -Needles $Needles) {
        return $true
    }

    return (Test-UserIsInTrusteeGroup -Trustee $Trustee -Needles $Needles)
}

function Get-ReverseMailboxPermissions {
    param(
        [string] $MailboxIdentity,
        [string] $UserIdentity
    )

    $permissionRows = [System.Collections.Generic.List[object]]::new()

    try {
        Get-MailboxPermission -Identity $MailboxIdentity -User $UserIdentity -IncludeUnresolvedPermissions -IncludeUserWithDisplayName -ErrorAction SilentlyContinue |
            Where-Object { $_.IsInherited -eq $false -and $_.Deny -eq $false } |
            ForEach-Object {
                $permissionRows.Add([PSCustomObject]@{
                    PermissionType = 'FullAccess'
                    GrantedTo      = Get-PrincipalDisplayValue $_.User
                    Rights         = ($_.AccessRights -join ', ')
                    Deny           = $_.Deny
                }) | Out-Null
            }
    } catch {}

    try {
        Get-RecipientPermission -Identity $MailboxIdentity -Trustee $UserIdentity -IncludeTrusteeWithPrimarySmtpAddress -ErrorAction SilentlyContinue |
            Where-Object { $_.AccessRights -contains 'SendAs' } |
            ForEach-Object {
                $permissionRows.Add([PSCustomObject]@{
                    PermissionType = 'SendAs'
                    GrantedTo      = Get-PrincipalDisplayValue $_.Trustee
                    Rights         = ($_.AccessRights -join ', ')
                    Deny           = $false
                }) | Out-Null
            }
    } catch {}

    return @($permissionRows)
}

# -- Connection ---------------------------------------------------------------
$script:ConnectedHere = $false
try {
    $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
} catch {
    $connectParams = @{ ShowBanner = $false }
    if ($TenantId) { $connectParams['Organization'] = $TenantId }
    Connect-ExchangeOnline @connectParams
    $script:ConnectedHere = $true
}

# -- Header ------------------------------------------------------------------
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Mailbox Permissions Audit" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

if ($User) {
    Write-Host "  Mode      : Reverse lookup by user" -ForegroundColor Cyan
}

# -- Get mailboxes ------------------------------------------------------------
if ($Mailbox -and $User) {
    throw "Use either -Mailbox or -User, not both."
}

$targetUser = $null
$targetUserValues = @()

if ($User) {
    try {
        $targetUser = Get-EXORecipient -Identity $User -ErrorAction Stop
    } catch {
        Write-Host "  [WARN] Could not resolve user '$User' to a recipient. Matching will use the provided value only." -ForegroundColor Yellow
    }
    $targetUserValues = @((Get-PrincipalMatchValues -Principal $User) + (Get-PrincipalMatchValues -Principal $targetUser) | Select-Object -Unique)
    Write-Host "  Searching mailbox permissions for user: $User" -ForegroundColor Cyan
}

if ($Mailbox) {
    $mailboxes = @(Get-EXOMailbox -Identity $Mailbox -Properties GrantSendOnBehalfTo -ErrorAction Stop)
} else {
    Write-Host "  Retrieving mailboxes..." -ForegroundColor DarkGray
    if ($User) {
        $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox -Properties GrantSendOnBehalfTo)
    } else {
        $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox -Properties GrantSendOnBehalfTo)
    }
}

if ($User) {
    Write-Host "  Checking where the user has permissions across $($mailboxes.Count) mailbox(es)..." -ForegroundColor DarkGray
} else {
    Write-Host "  Checking permissions for $($mailboxes.Count) mailbox(es)..." -ForegroundColor DarkGray
}
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($mbx in $mailboxes) {
    $mbxUpn = $mbx.UserPrincipalName

    if ($User) {
        $reverseMatches = Get-ReverseMailboxPermissions -MailboxIdentity $mbxUpn -UserIdentity $User

        foreach ($match in $reverseMatches) {
            $results.Add([PSCustomObject]@{
                Mailbox        = $mbxUpn
                PermissionType = $match.PermissionType
                GrantedTo      = $match.GrantedTo
                Rights         = $match.Rights
                Deny           = $match.Deny
            })
        }

        if ($mbx.GrantSendOnBehalfTo) {
            foreach ($delegate in $mbx.GrantSendOnBehalfTo) {
                if (Test-UserMatchesTrustee -Trustee $delegate -Needles $targetUserValues) {
                    $results.Add([PSCustomObject]@{
                        Mailbox        = $mbxUpn
                        PermissionType = 'SendOnBehalf'
                        GrantedTo      = Get-PrincipalDisplayValue $delegate
                        Rights         = 'SendOnBehalf'
                        Deny           = $false
                    }) | Out-Null
                }
            }
        }

        continue
    }

    # Full Access
    try {
        Get-MailboxPermission -Identity $mbxUpn -ErrorAction SilentlyContinue |
            Where-Object { $_.IsInherited -eq $false -and $_.User -notlike '*SELF*' } |
            ForEach-Object {
                if (-not $User -or (Test-UserMatchesTrustee -Trustee $_.User -Needles $targetUserValues)) {
                    $results.Add([PSCustomObject]@{
                        Mailbox        = $mbxUpn
                        PermissionType = 'FullAccess'
                        GrantedTo      = Get-PrincipalDisplayValue $_.User
                        Rights         = ($_.AccessRights -join ', ')
                        Deny           = $_.Deny
                    })
                }
            }
    } catch {
        Write-Host "  [WARN] FullAccess - $mbxUpn : $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Send As
    try {
        Get-RecipientPermission -Identity $mbxUpn -ErrorAction SilentlyContinue |
            Where-Object { $_.Trustee -notlike '*SELF*' -and $_.Trustee -notlike 'NT AUTHORITY*' } |
            ForEach-Object {
                if (-not $User -or (Test-UserMatchesTrustee -Trustee $_.Trustee -Needles $targetUserValues)) {
                    $results.Add([PSCustomObject]@{
                        Mailbox        = $mbxUpn
                        PermissionType = 'SendAs'
                        GrantedTo      = Get-PrincipalDisplayValue $_.Trustee
                        Rights         = ($_.AccessRights -join ', ')
                        Deny           = $false
                    })
                }
            }
    } catch {
        Write-Host "  [WARN] SendAs - $mbxUpn : $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Send on Behalf
    if ($mbx.GrantSendOnBehalfTo) {
        foreach ($delegate in $mbx.GrantSendOnBehalfTo) {
            if ($User -and -not (Test-UserMatchesTrustee -Trustee $delegate -Needles $targetUserValues)) {
                continue
            }
            $results.Add([PSCustomObject]@{
                Mailbox        = $mbxUpn
                PermissionType = 'SendOnBehalf'
                GrantedTo      = Get-PrincipalDisplayValue $delegate
                Rights         = 'SendOnBehalf'
                Deny           = $false
            })
        }
    }
}

# -- Output ------------------------------------------------------------------
if ($results.Count -eq 0) {
    if ($User) {
        Write-Host "  No mailbox permissions found for this user." -ForegroundColor DarkGray
    } else {
        Write-Host "  No delegated permissions found." -ForegroundColor DarkGray
    }
} else {
    $results | Format-Table -AutoSize

    if (-not $OutputPath) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $OutputPath = Join-Path $outputDir "MailboxPermissions_$ts.csv"
    }

    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "  Checked $($mailboxes.Count) mailbox(es) - $($results.Count) permission entries found." -ForegroundColor Cyan
Write-Host ""

# -- Disconnect if we connected ----------------------------------------------
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
