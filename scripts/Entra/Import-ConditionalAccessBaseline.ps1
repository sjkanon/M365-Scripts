#Requires -Version 5.1
<#
.SYNOPSIS
    Import the latest Conditional Access Baseline into a tenant.

.DESCRIPTION
    Downloads the latest baseline from:
    https://github.com/j0eyv/ConditionalAccessBaseline

    The script creates (or reuses) required groups and named locations, remaps
    old IDs to your tenant IDs, then imports/updates all Conditional Access
    policies.

    Default import state is OFF (disabled), so you can review and enable later.

    Includes a second action to change policy state later (Enable/ReportOnly/Off).

.PARAMETER Action
    Import   : download + import baseline (default)
    SetState : change state of already imported CA baseline policies

.PARAMETER PolicyStateOnImport
    State used during Action=Import.
    disabled (Off) is the default.

.PARAMETER TargetState
    State used during Action=SetState.

.PARAMETER SourcePath
    Optional local path to a ConditionalAccessBaseline folder that contains
    Config/ConditionalAccess, Config/Groups, Config/NamedLocations.

.PARAMETER TenantId
    Optional tenant ID or domain for Connect-MgGraph.

.PARAMETER UpdateExisting
    If set, existing policies are updated. If not set, existing policies are skipped.

.PARAMETER CreateMissingServicePrincipals
    If set, attempts to create missing service principals for GUID app IDs found
    in policy include/exclude application lists.

.PARAMETER EnsureIntuneEnrollmentServicePrincipal
    Ensures Microsoft Intune Enrollment service principal exists (recommended).

.EXAMPLE
    .\Import-ConditionalAccessBaseline.ps1

.EXAMPLE
    .\Import-ConditionalAccessBaseline.ps1 -PolicyStateOnImport enabledForReportingButNotEnforced

.EXAMPLE
    .\Import-ConditionalAccessBaseline.ps1 -Action SetState -TargetState enabled
#>
[CmdletBinding()]
param (
    [ValidateSet('Import', 'SetState')]
    [string]$Action = 'Import',

    [ValidateSet('disabled', 'enabledForReportingButNotEnforced')]
    [string]$PolicyStateOnImport = 'disabled',

    [ValidateSet('disabled', 'enabledForReportingButNotEnforced', 'enabled')]
    [string]$TargetState = 'enabled',

    [string]$SourcePath,

    [string]$TenantId,

    [switch]$UpdateExisting,

    [switch]$CreateMissingServicePrincipals,

    [switch]$EnsureIntuneEnrollmentServicePrincipal = $true
)

$ErrorActionPreference = 'Stop'

function Write-Step { param([string]$Message) Write-Host "`n=== $Message ===" -ForegroundColor Cyan }
function Write-Ok { param([string]$Message) Write-Host "[OK]   $Message" -ForegroundColor Green }
function Write-Warn { param([string]$Message) Write-Host "[WARN] $Message" -ForegroundColor Yellow }
function Write-Err { param([string]$Message) Write-Host "[ERR]  $Message" -ForegroundColor Red }

function New-MailNickname {
    param([string]$DisplayName)

    $nick = ($DisplayName -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($nick)) { $nick = 'cabaseline' }
    if ($nick.Length -gt 48) { $nick = $nick.Substring(0, 48) }

    $suffix = -join ((48..57) + (97..122) | Get-Random -Count 4 | ForEach-Object { [char]$_ })
    return "$nick$suffix"
}

function ConvertTo-Hashtable {
    param([Parameter(ValueFromPipeline = $true)]$InputObject)

    if ($null -eq $InputObject) { return $null }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $ht = @{}
        foreach ($key in $InputObject.Keys) {
            $ht[$key] = ConvertTo-Hashtable -InputObject $InputObject[$key]
        }
        return $ht
    }

    if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
        $arr = @()
        foreach ($item in $InputObject) {
            $arr += ,(ConvertTo-Hashtable -InputObject $item)
        }
        return $arr
    }

    if ($InputObject -is [psobject]) {
        $ht = @{}
        foreach ($prop in $InputObject.PSObject.Properties) {
            $ht[$prop.Name] = ConvertTo-Hashtable -InputObject $prop.Value
        }
        return $ht
    }

    return $InputObject
}

function Remove-GraphMetadata {
    param($Object)

    if ($null -eq $Object) { return $null }

    if ($Object -is [System.Collections.IDictionary]) {
        $clean = @{}
        foreach ($key in @($Object.Keys)) {
            if ($key -match '^@odata\.' -or $key -match '@odata\.' -or $key -like '#microsoft.graph.*') {
                continue
            }

            if ($key -in @(
                'id',
                'createdDateTime',
                'modifiedDateTime',
                'deletedDateTime',
                'createdByAppId',
                'organizationId',
                'securityIdentifier',
                'renewedDateTime',
                'onPremisesLastSyncDateTime',
                'onPremisesSyncEnabled',
                'templateId',
                'partialEnablementStrategy'
            )) {
                continue
            }

            $clean[$key] = Remove-GraphMetadata -Object $Object[$key]
        }
        return $clean
    }

    if ($Object -is [System.Collections.IEnumerable] -and -not ($Object -is [string])) {
        $arr = @()
        foreach ($item in $Object) {
            $arr += ,(Remove-GraphMetadata -Object $item)
        }
        return $arr
    }

    return $Object
}

function Remap-IdArray {
    param(
        [array]$Ids,
        [hashtable]$IdMap
    )

    if ($null -eq $Ids) { return $Ids }

    $result = @()
    foreach ($id in $Ids) {
        if ($IdMap.ContainsKey($id)) {
            $result += $IdMap[$id]
        } else {
            $result += $id
        }
    }
    return $result
}

function Test-Guid {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $guid = [Guid]::Empty
    return [Guid]::TryParse($Value, [ref]$guid)
}

function Get-JsonFileObject {
    param([string]$Path)

    $raw = Get-Content -Path $Path -Raw -Encoding UTF8
    # Strip UTF-8 BOM if present
    if ($raw.Length -gt 0 -and [int][char]$raw[0] -eq 65279) {
        $raw = $raw.Substring(1)
    }
    return ($raw | ConvertFrom-Json)
}

function Connect-GraphIfNeeded {
    param([string]$Tenant)

    $script:ConnectedHere = $false
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $ctx) {
        $connectParams = @{
            Scopes = @(
                'Policy.ReadWrite.ConditionalAccess',
                'Policy.Read.All',
                'Group.ReadWrite.All',
                'Directory.ReadWrite.All',
                'Application.ReadWrite.All'
            )
            NoWelcome = $true
        }
        if ($Tenant) { $connectParams['TenantId'] = $Tenant }

        Connect-MgGraph @connectParams
        $script:ConnectedHere = $true
        $ctx = Get-MgContext
    }

    Write-Ok "Connected as $($ctx.Account) to tenant $($ctx.TenantId)"
}

function Resolve-BaselinePath {
    param([string]$PathFromParam)

    if ($PathFromParam) {
        if (-not (Test-Path -Path $PathFromParam -PathType Container)) {
            throw "SourcePath not found: $PathFromParam"
        }
        return (Resolve-Path $PathFromParam).Path
    }

    $tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('CA-Baseline-' + (Get-Date -Format 'yyyyMMddHHmmss'))
    New-Item -Path $tmpRoot -ItemType Directory -Force | Out-Null

    $zipPath = Join-Path $tmpRoot 'baseline.zip'
    $url = 'https://github.com/j0eyv/ConditionalAccessBaseline/archive/refs/heads/main.zip'

    Write-Step 'Downloading latest baseline from GitHub'
    Invoke-WebRequest -Uri $url -OutFile $zipPath

    Write-Step 'Extracting baseline package'
    Expand-Archive -Path $zipPath -DestinationPath $tmpRoot -Force

    $repoPath = Join-Path $tmpRoot 'ConditionalAccessBaseline-main'
    if (-not (Test-Path -Path $repoPath -PathType Container)) {
        throw "Could not resolve extracted baseline folder at $repoPath"
    }

    Write-Ok "Baseline downloaded to $repoPath"
    return $repoPath
}

function Ensure-IntuneEnrollmentSp {
    $appId = 'd4ebce55-015a-49b5-a083-c84d1797ae8c'
    $sp = Get-MgServicePrincipal -Filter "appId eq '$appId'" -ConsistencyLevel eventual | Select-Object -First 1
    if ($sp) {
        Write-Ok 'Microsoft Intune Enrollment service principal exists'
        return
    }

    Write-Warn 'Microsoft Intune Enrollment service principal missing; creating...'
    New-MgServicePrincipal -AppId $appId | Out-Null
    Write-Ok 'Microsoft Intune Enrollment service principal created'
}

function Import-BaselineGroups {
    param(
        [string]$GroupsPath,
        [string]$MigrationTablePath,
        [hashtable]$IdMap
    )

    Write-Step 'Importing/reusing groups'

    $migration = $null
    if (Test-Path -Path $MigrationTablePath -PathType Leaf) {
        $migration = Get-JsonFileObject -Path $MigrationTablePath
    }

    $groupFiles = Get-ChildItem -Path $GroupsPath -Filter '*.json' -File | Sort-Object Name
    foreach ($file in $groupFiles) {
        $groupObj = Get-JsonFileObject -Path $file.FullName
        $displayName = [string]$groupObj.displayName
        $description = [string]$groupObj.description
        $oldId = [string]$groupObj.id

        if ([string]::IsNullOrWhiteSpace($displayName)) {
            Write-Warn "Skipping group file without displayName: $($file.Name)"
            continue
        }

        $safeDisplayName = $displayName.Replace("'", "''")
        $existing = Get-MgGroup -Filter "displayName eq '$safeDisplayName'" -ConsistencyLevel eventual | Select-Object -First 1

        if ($existing) {
            Write-Ok "Group exists: $displayName"
            if (-not [string]::IsNullOrWhiteSpace($oldId)) {
                $IdMap[$oldId] = $existing.Id
            }
            continue
        }

        $mailEnabled = [bool]$groupObj.mailEnabled
        $securityEnabled = [bool]$groupObj.securityEnabled
        $mailNickname = if ($groupObj.mailNickname) { [string]$groupObj.mailNickname } else { New-MailNickname -DisplayName $displayName }

        $createBody = @{
            displayName = $displayName
            mailEnabled = $mailEnabled
            mailNickname = $mailNickname
            securityEnabled = $securityEnabled
            description = $description
        }

        if ($groupObj.groupTypes -and $groupObj.groupTypes.Count -gt 0) {
            $createBody['groupTypes'] = @($groupObj.groupTypes)
        }

        if ($groupObj.membershipRule) {
            $createBody['membershipRule'] = [string]$groupObj.membershipRule
            $createBody['membershipRuleProcessingState'] = if ($groupObj.membershipRuleProcessingState) {
                [string]$groupObj.membershipRuleProcessingState
            } else {
                'On'
            }
        }

        try {
            $newGroup = New-MgGroup -BodyParameter $createBody
            Write-Ok "Created group: $displayName"
            if (-not [string]::IsNullOrWhiteSpace($oldId)) {
                $IdMap[$oldId] = $newGroup.Id
            }
        } catch {
            Write-Warn "Failed to create group '$displayName': $($_.Exception.Message)"
        }
    }

    if ($migration -and $migration.Objects) {
        # Extra safety: map old IDs from migration table by displayName
        $allGroups = Get-MgGroup -All
        foreach ($obj in $migration.Objects) {
            if ($obj.Type -ne 'Group') { continue }
            $dn = [string]$obj.DisplayName
            $oldId = [string]$obj.Id

            $match = $allGroups | Where-Object DisplayName -eq $dn | Select-Object -First 1
            if ($match -and -not [string]::IsNullOrWhiteSpace($oldId)) {
                $IdMap[$oldId] = $match.Id
            }
        }
    }
}

function Import-NamedLocations {
    param(
        [string]$NamedLocationsPath,
        [hashtable]$IdMap
    )

    Write-Step 'Importing/reusing named locations'

    $existingLocations = Get-MgIdentityConditionalAccessNamedLocation -All
    $locationFiles = Get-ChildItem -Path $NamedLocationsPath -Filter '*.json' -File | Sort-Object Name

    foreach ($file in $locationFiles) {
        $raw = Get-JsonFileObject -Path $file.FullName
        $rawHash = ConvertTo-Hashtable -InputObject $raw

        $displayName = [string]$rawHash.displayName
        $oldId = [string]$rawHash.id
        $odataType = [string]$rawHash.'@odata.type'

        if ([string]::IsNullOrWhiteSpace($displayName)) {
            Write-Warn "Skipping named location without displayName: $($file.Name)"
            continue
        }

        $existing = $existingLocations | Where-Object DisplayName -eq $displayName | Select-Object -First 1
        if ($existing) {
            Write-Ok "Named location exists: $displayName"
            if (-not [string]::IsNullOrWhiteSpace($oldId)) {
                $IdMap[$oldId] = $existing.Id
            }
            continue
        }

        $body = @{
            displayName = $displayName
        }

        if ($odataType -match 'countryNamedLocation') {
            $body['@odata.type'] = '#microsoft.graph.countryNamedLocation'
            $body['countriesAndRegions'] = @($rawHash.countriesAndRegions)
            $body['includeUnknownCountriesAndRegions'] = [bool]$rawHash.includeUnknownCountriesAndRegions
            if ($rawHash.countryLookupMethod) {
                $body['countryLookupMethod'] = [string]$rawHash.countryLookupMethod
            }
        } elseif ($odataType -match 'ipNamedLocation') {
            $body['@odata.type'] = '#microsoft.graph.ipNamedLocation'
            $body['isTrusted'] = [bool]$rawHash.isTrusted
            $ranges = @()
            foreach ($range in @($rawHash.ipRanges)) {
                $ranges += @{
                    '@odata.type' = [string]$range.'@odata.type'
                    cidrAddress = [string]$range.cidrAddress
                }
            }
            $body['ipRanges'] = $ranges
        } else {
            Write-Warn "Unsupported named location type in $($file.Name): $odataType"
            continue
        }

        try {
            $created = Invoke-MgGraphRequest -Method POST -Uri '/beta/identity/conditionalAccess/namedLocations' -Body $body
            Write-Ok "Created named location: $displayName"
            if ($created.id -and -not [string]::IsNullOrWhiteSpace($oldId)) {
                $IdMap[$oldId] = [string]$created.id
            }
            $existingLocations += $created
        } catch {
            Write-Warn "Failed to create named location '$displayName': $($_.Exception.Message)"
        }
    }
}

function Ensure-ServicePrincipalsFromPolicies {
    param([string]$PoliciesPath)

    Write-Step 'Checking service principals referenced by policies'

    $knownKeywords = @('All', 'None', 'Office365', 'MicrosoftAdminPortals')
    $allIds = New-Object System.Collections.Generic.HashSet[string]

    $files = Get-ChildItem -Path $PoliciesPath -Filter '*.json' -File
    foreach ($file in $files) {
        $obj = Get-JsonFileObject -Path $file.FullName
        $apps = @()
        if ($obj.conditions -and $obj.conditions.applications) {
            $apps += @($obj.conditions.applications.includeApplications)
            $apps += @($obj.conditions.applications.excludeApplications)
        }

        foreach ($item in $apps) {
            $value = [string]$item
            if ([string]::IsNullOrWhiteSpace($value)) { continue }
            if ($knownKeywords -contains $value) { continue }
            if (Test-Guid -Value $value) {
                [void]$allIds.Add($value)
            }
        }
    }

    foreach ($appId in $allIds) {
        $sp = Get-MgServicePrincipal -Filter "appId eq '$appId'" -ConsistencyLevel eventual | Select-Object -First 1
        if ($sp) {
            Write-Ok "Service principal exists for AppId $appId"
            continue
        }

        Write-Warn "Service principal missing for AppId $appId"
        if ($CreateMissingServicePrincipals) {
            try {
                New-MgServicePrincipal -AppId $appId | Out-Null
                Write-Ok "Created service principal for AppId $appId"
            } catch {
                Write-Warn "Could not create service principal for AppId $appId: $($_.Exception.Message)"
            }
        }
    }
}

function Build-PolicyBody {
    param(
        [hashtable]$RawPolicy,
        [string]$State,
        [hashtable]$IdMap
    )

    $policy = Remove-GraphMetadata -Object $RawPolicy

    # Enforce desired state on import
    $policy['state'] = $State

    # Remap group IDs
    if ($policy.conditions -and $policy.conditions.users) {
        if ($policy.conditions.users.includeGroups) {
            $policy.conditions.users.includeGroups = Remap-IdArray -Ids $policy.conditions.users.includeGroups -IdMap $IdMap
        }
        if ($policy.conditions.users.excludeGroups) {
            $policy.conditions.users.excludeGroups = Remap-IdArray -Ids $policy.conditions.users.excludeGroups -IdMap $IdMap
        }
    }

    # Remap named location IDs
    if ($policy.conditions -and $policy.conditions.locations) {
        if ($policy.conditions.locations.includeLocations) {
            $policy.conditions.locations.includeLocations = Remap-IdArray -Ids $policy.conditions.locations.includeLocations -IdMap $IdMap
        }
        if ($policy.conditions.locations.excludeLocations) {
            $policy.conditions.locations.excludeLocations = Remap-IdArray -Ids $policy.conditions.locations.excludeLocations -IdMap $IdMap
        }
    }

    # Normalize authentication strength payload if present
    if ($policy.grantControls -and $policy.grantControls.authenticationStrength) {
        $strength = $policy.grantControls.authenticationStrength
        if ($strength.id) {
            $policy.grantControls.authenticationStrength = @{ id = [string]$strength.id }
        } else {
            $policy.grantControls.Remove('authenticationStrength')
        }
    }

    return $policy
}

function Import-ConditionalAccessPolicies {
    param(
        [string]$PoliciesPath,
        [string]$State,
        [hashtable]$IdMap,
        [switch]$AllowUpdate
    )

    Write-Step "Importing Conditional Access policies (state: $State)"

    $policyFiles = Get-ChildItem -Path $PoliciesPath -Filter '*.json' -File | Sort-Object Name
    $existingPolicies = Get-MgIdentityConditionalAccessPolicy -All

    $created = 0
    $updated = 0
    $skipped = 0
    $failed = 0

    foreach ($file in $policyFiles) {
        try {
            $rawObj = Get-JsonFileObject -Path $file.FullName
            $rawHash = ConvertTo-Hashtable -InputObject $rawObj

            $displayName = [string]$rawHash.displayName
            if ([string]::IsNullOrWhiteSpace($displayName)) {
                Write-Warn "Skipping policy without displayName: $($file.Name)"
                $skipped++
                continue
            }

            $body = Build-PolicyBody -RawPolicy $rawHash -State $State -IdMap $IdMap
            $existing = $existingPolicies | Where-Object DisplayName -eq $displayName | Select-Object -First 1

            if ($existing) {
                if ($AllowUpdate) {
                    Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $existing.Id -BodyParameter $body | Out-Null
                    Write-Ok "Updated policy: $displayName"
                    $updated++
                } else {
                    Write-Warn "Policy exists (skipped): $displayName"
                    $skipped++
                }
            } else {
                $newPolicy = New-MgIdentityConditionalAccessPolicy -BodyParameter $body
                Write-Ok "Created policy: $displayName"
                $created++
                $existingPolicies += $newPolicy
            }
        } catch {
            Write-Warn "Failed policy $($file.Name): $($_.Exception.Message)"
            $failed++
        }
    }

    Write-Host ''
    Write-Host "Summary:" -ForegroundColor Cyan
    Write-Host "  Created : $created"
    Write-Host "  Updated : $updated"
    Write-Host "  Skipped : $skipped"
    Write-Host "  Failed  : $failed"
}

function Set-BaselinePolicyState {
    param(
        [string]$State,
        [string[]]$PolicyNames
    )

    Write-Step "Setting policy state to $State"

    $allPolicies = Get-MgIdentityConditionalAccessPolicy -All

    if ($PolicyNames -and $PolicyNames.Count -gt 0) {
        $targets = $allPolicies | Where-Object { $PolicyNames -contains $_.DisplayName }
    } else {
        $targets = $allPolicies | Where-Object { $_.DisplayName -match '^CA\d{3}-' }
    }

    $changed = 0
    $failed = 0

    foreach ($policy in $targets) {
        try {
            Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id -BodyParameter @{ state = $State } | Out-Null
            Write-Ok "Updated state: $($policy.DisplayName)"
            $changed++
        } catch {
            Write-Warn "Failed state update: $($policy.DisplayName) - $($_.Exception.Message)"
            $failed++
        }
    }

    Write-Host ''
    Write-Host "State update summary:" -ForegroundColor Cyan
    Write-Host "  Updated : $changed"
    Write-Host "  Failed  : $failed"
}

try {
    Write-Step 'Conditional Access Baseline Import'
    Connect-GraphIfNeeded -Tenant $TenantId

    if ($Action -eq 'SetState') {
        $policyNames = @()
        if ($SourcePath) {
            $caPath = Join-Path $SourcePath 'Config/ConditionalAccess'
            if (Test-Path -Path $caPath -PathType Container) {
                $policyNames = Get-ChildItem -Path $caPath -Filter '*.json' -File |
                    ForEach-Object {
                        try {
                            (Get-JsonFileObject -Path $_.FullName).displayName
                        } catch {
                            $null
                        }
                    } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            }
        }

        Set-BaselinePolicyState -State $TargetState -PolicyNames $policyNames
        return
    }

    $baselineRoot = Resolve-BaselinePath -PathFromParam $SourcePath
    $configPath = Join-Path $baselineRoot 'Config'
    $groupsPath = Join-Path $configPath 'Groups'
    $namedLocationsPath = Join-Path $configPath 'NamedLocations'
    $policiesPath = Join-Path $configPath 'ConditionalAccess'
    $migrationPath = Join-Path $configPath 'MigrationTable.json'

    foreach ($required in @($configPath, $groupsPath, $namedLocationsPath, $policiesPath)) {
        if (-not (Test-Path -Path $required -PathType Container)) {
            throw "Required folder not found: $required"
        }
    }

    if ($EnsureIntuneEnrollmentServicePrincipal) {
        Ensure-IntuneEnrollmentSp
    }

    $idMap = @{}

    Import-BaselineGroups -GroupsPath $groupsPath -MigrationTablePath $migrationPath -IdMap $idMap
    Import-NamedLocations -NamedLocationsPath $namedLocationsPath -IdMap $idMap

    Ensure-ServicePrincipalsFromPolicies -PoliciesPath $policiesPath

    Import-ConditionalAccessPolicies -PoliciesPath $policiesPath -State $PolicyStateOnImport -IdMap $idMap -AllowUpdate:$UpdateExisting

    Write-Host ''
    Write-Host 'Done. Policies are imported in OFF/report-only mode as requested.' -ForegroundColor Green
    Write-Host 'When ready, run Action=SetState to move policies to report-only or enabled.' -ForegroundColor Green
}
catch {
    Write-Err $_.Exception.Message
    exit 1
}
finally {
    if ($script:ConnectedHere) {
        Disconnect-MgGraph | Out-Null
    }
}
