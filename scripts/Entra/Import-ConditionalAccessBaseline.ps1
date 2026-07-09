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
    RemovePolicy : remove one specific CA policy or all baseline CA policies

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

.PARAMETER InstallMissingModules
    If set, installs missing required Microsoft Graph modules in CurrentUser scope.

.PARAMETER BaselineProvider
    Select baseline source:
    j0eyv  : JSON import from ConditionalAccessBaseline repo
    Daniel : DCToolbox baseline deployment (Daniel Chronlund)

.PARAMETER BaselinePrefix
    Optional prefix for policy names. For Daniel provider this maps to
    Deploy-DCConditionalAccessBaselinePoC -AddCustomPrefix.

.PARAMETER RemovePrefixes
    Optional list of policy name prefixes used for Action=RemovePolicy.
    Example: 'PILOT - ', 'GLOBAL - '

.PARAMETER DanielAutoDeployIds
    Optional list of template IDs for Daniel provider. If set, deployment uses
    Invoke-DCConditionalAccessGallery -AutoDeployIds instead of
    Deploy-DCConditionalAccessBaselinePoC.

.PARAMETER DanielUseRecommendedIds
    For Daniel provider, deploy a curated recommended list of baseline template
    IDs via Invoke-DCConditionalAccessGallery.

.PARAMETER PolicyName
    Display name of the conditional access policy to remove (Action=RemovePolicy).

.PARAMETER RemoveAllBaselinePolicies
    Remove all baseline CA policies. If SourcePath is provided, only names from
    Config/ConditionalAccess are removed. Without SourcePath, all policies with
    display name matching ^CA\d{3}- are removed.

.PARAMETER RemoveAssociatedGroups
    Also remove baseline groups associated with the CA baseline.

.PARAMETER RemoveAssociatedNamedLocations
    Also remove baseline named locations associated with the CA baseline.

.PARAMETER Force
    Skip confirmation prompts for removals.

.EXAMPLE
    .\Import-ConditionalAccessBaseline.ps1

.EXAMPLE
    .\Import-ConditionalAccessBaseline.ps1 -PolicyStateOnImport enabledForReportingButNotEnforced

.EXAMPLE
    .\Import-ConditionalAccessBaseline.ps1 -Action SetState -TargetState enabled

.EXAMPLE
    .\Import-ConditionalAccessBaseline.ps1 -InstallMissingModules

.EXAMPLE
    .\Import-ConditionalAccessBaseline.ps1 -Action RemovePolicy -PolicyName "CA402-GuestUsers-IdentityProtection-AllApps-AnyPlatform-SigninFrequency"

.EXAMPLE
    .\Import-ConditionalAccessBaseline.ps1 -Action RemovePolicy -RemoveAllBaselinePolicies -Force

.EXAMPLE
    .\Import-ConditionalAccessBaseline.ps1 -RemoveAllBaselinePolicies -RemoveAssociatedGroups -RemoveAssociatedNamedLocations -Force

.EXAMPLE
    .\Import-ConditionalAccessBaseline.ps1 -BaselineProvider Daniel -PolicyStateOnImport disabled -BaselinePrefix 'PILOT - '

.EXAMPLE
    .\Import-ConditionalAccessBaseline.ps1 -Action RemovePolicy -RemoveAllBaselinePolicies -RemovePrefixes 'PILOT - ','GLOBAL - ' -Force

.EXAMPLE
    .\Import-ConditionalAccessBaseline.ps1 -BaselineProvider Daniel -DanielUseRecommendedIds -BaselinePrefix 'PILOT - '
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
    [ValidateSet('Import', 'SetState', 'RemovePolicy')]
    [string]$Action = 'Import',

    [ValidateSet('disabled', 'enabledForReportingButNotEnforced')]
    [string]$PolicyStateOnImport = 'disabled',

    [ValidateSet('disabled', 'enabledForReportingButNotEnforced', 'enabled')]
    [string]$TargetState = 'enabled',

    [string]$SourcePath,

    [string]$TenantId,

    [switch]$UpdateExisting,

    [switch]$CreateMissingServicePrincipals,

    [switch]$EnsureIntuneEnrollmentServicePrincipal = $true,

    [switch]$InstallMissingModules,

    [ValidateSet('j0eyv', 'Daniel')]
    [string]$BaselineProvider = 'j0eyv',

    [string]$BaselinePrefix = '',

    [string[]]$RemovePrefixes,

    [int[]]$DanielAutoDeployIds,

    [switch]$DanielUseRecommendedIds,

    [string]$PolicyName,

    [switch]$RemoveAllBaselinePolicies,

    [switch]$RemoveAssociatedGroups,

    [switch]$RemoveAssociatedNamedLocations,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

function Write-Step { param([string]$Message) Write-Host "`n=== $Message ===" -ForegroundColor Cyan }
function Write-Ok { param([string]$Message) Write-Host "[OK]   $Message" -ForegroundColor Green }
function Write-Warn { param([string]$Message) Write-Host "[WARN] $Message" -ForegroundColor Yellow }
function Write-Err { param([string]$Message) Write-Host "[ERR]  $Message" -ForegroundColor Red }

function Ensure-RequiredModules {
    param(
        [switch]$InstallMissing,
        [string]$Provider
    )

    Write-Step 'Checking required modules'

    $requiredModules = @(
        @{ Name = 'Microsoft.Graph.Authentication'; MinimumVersion = '2.0.0' }
        @{ Name = 'Microsoft.Graph.Identity.SignIns'; MinimumVersion = '2.0.0' }
        @{ Name = 'Microsoft.Graph.Applications'; MinimumVersion = '2.0.0' }
        @{ Name = 'Microsoft.Graph.Groups'; MinimumVersion = '2.0.0' }
    )

    if ($Provider -eq 'Daniel') {
        $requiredModules += @{ Name = 'DCToolbox'; MinimumVersion = $null }
        $requiredModules += @{ Name = 'Microsoft.Graph.Identity.Governance'; MinimumVersion = '2.0.0' }
    }

    $missing = @()

    foreach ($mod in $requiredModules) {
        $available = Get-Module -ListAvailable -Name $mod.Name | Sort-Object Version -Descending | Select-Object -First 1
        $needsInstall = $false

        if (-not $available) {
            $needsInstall = $true
        } elseif ($mod.MinimumVersion -and $available.Version -lt [Version]$mod.MinimumVersion) {
            $needsInstall = $true
        }

        if ($needsInstall) {
            if ($InstallMissing) {
                try {
                    Write-Warn "Installing missing module: $($mod.Name)"
                    Install-Module -Name $mod.Name -MinimumVersion $mod.MinimumVersion -Scope CurrentUser -AllowClobber -Force -ErrorAction Stop
                    $available = Get-Module -ListAvailable -Name $mod.Name | Sort-Object Version -Descending | Select-Object -First 1
                } catch {
                    Write-Warn "Install failed for module '$($mod.Name)': $($_.Exception.Message)"
                }
            }
        }

        if (-not $available) {
            $missing += "$($mod.Name) (min $($mod.MinimumVersion))"
            continue
        }

        if ($mod.MinimumVersion -and $available.Version -lt [Version]$mod.MinimumVersion) {
            $missing += "$($mod.Name) (installed: $($available.Version), required: $($mod.MinimumVersion))"
            continue
        }

        Import-Module -Name $mod.Name -ErrorAction SilentlyContinue
        Write-Ok "$($mod.Name) $($available.Version)"
    }

    if ($missing.Count -gt 0) {
        $hint = @(
            'Missing required modules:'
            ($missing | ForEach-Object { " - $_" })
            'Run scripts/Startup/Install-Modules.ps1 or rerun with -InstallMissingModules.'
        ) -join "`n"
        throw $hint
    }
}

function Get-TargetPolicyMatchPattern {
    param(
        [string]$Provider,
        [string]$Prefix
    )

    if ($Provider -eq 'Daniel') {
        if (-not [string]::IsNullOrWhiteSpace($Prefix)) {
            return '^' + [Regex]::Escape($Prefix)
        }
        return '^(GLOBAL|PILOT|OVERRIDE)\s-\s'
    }

    if (-not [string]::IsNullOrWhiteSpace($Prefix)) {
        return '^' + [Regex]::Escape($Prefix)
    }

    return '^CA\d{3}-'
}

function Get-DanielRecommendedTemplateIds {
    # Version 15 baseline set from Daniel Chronlund's baseline (unique IDs).
    return @(1010, 1020, 1030, 1040, 1050, 1060, 1070, 1080, 1090, 1100, 2010, 2020, 2040, 2050, 2055, 2060, 2070, 3010, 3020, 3030, 3040, 1)
}

function Deploy-DanielBaseline {
    param(
        [string]$DesiredState,
        [string]$Prefix,
        [switch]$ForceInstallModules,
        [int[]]$AutoDeployIds,
        [switch]$UseRecommendedIds
    )

    Write-Step 'Deploying Daniel Chronlund baseline (DCToolbox)'

    if ($PSVersionTable.PSVersion.Major -lt 7) {
        throw 'Daniel baseline deployment requires PowerShell 7+ (Linux compatible path).'
    }

    if ($ForceInstallModules -or -not (Get-Module -ListAvailable -Name DCToolbox)) {
        Write-Warn 'Installing DCToolbox module...'
        Install-Module -Name DCToolbox -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }

    Import-Module DCToolbox -ErrorAction Stop

    $effectiveIds = $null

    if ($AutoDeployIds -and $AutoDeployIds.Count -gt 0) {
        $effectiveIds = @($AutoDeployIds | Select-Object -Unique)
    } elseif ($UseRecommendedIds) {
        $effectiveIds = Get-DanielRecommendedTemplateIds
    }

    # DCToolbox defaults to report-only unless -SkipReportOnlyMode is used.
    if ($effectiveIds -and $effectiveIds.Count -gt 0) {
        Write-Ok "Using Invoke-DCConditionalAccessGallery with AutoDeployIds: $($effectiveIds -join ', ')"

        if ([string]::IsNullOrWhiteSpace($Prefix)) {
            Invoke-DCConditionalAccessGallery -AutoDeployIds $effectiveIds -SkipDocumentation | Out-Null
        } else {
            Invoke-DCConditionalAccessGallery -AddCustomPrefix $Prefix -AutoDeployIds $effectiveIds -SkipDocumentation | Out-Null
        }
    } else {
        if ([string]::IsNullOrWhiteSpace($Prefix)) {
            Deploy-DCConditionalAccessBaselinePoC
        } else {
            Deploy-DCConditionalAccessBaselinePoC -AddCustomPrefix $Prefix
        }
    }

    # Align final state with this script's convention.
    $pattern = Get-TargetPolicyMatchPattern -Provider 'Daniel' -Prefix $Prefix
    $targets = Get-MgIdentityConditionalAccessPolicy -All | Where-Object { $_.DisplayName -match $pattern }

    foreach ($policy in $targets) {
        Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id -BodyParameter @{ state = $DesiredState } | Out-Null
    }

    if ($effectiveIds -and $effectiveIds.Count -gt 0) {
        $expectedCount = ($effectiveIds | Select-Object -Unique).Count
        $actualCount = @($targets).Count
        if ($actualCount -lt $expectedCount) {
            Write-Warn "Daniel deployment may be incomplete. Expected around $expectedCount policies for selected IDs, found $actualCount matching '$pattern'."
            Write-Warn 'Possible causes: missing permissions, missing dependencies, or template creation failures in DCToolbox output.'
        } else {
            Write-Ok "Daniel deployment completeness check passed: found $actualCount policies for selected IDs."
        }
    }

    Write-Ok "Daniel baseline deployed. Policies matched by '$pattern' set to state '$DesiredState'."
}

function New-MailNickname {
    param([string]$DisplayName)

    $nick = ($DisplayName -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($nick)) { $nick = 'cabaseline' }
    if ($nick.Length -gt 48) { $nick = $nick.Substring(0, 48) }

    $suffix = -join ((48..57) + (97..122) | Get-Random -Count 4 | ForEach-Object { [char]$_ })
    return "$nick$suffix"
}

function Get-SafeGroupDescription {
    param([string]$Description)

    if ([string]::IsNullOrWhiteSpace($Description)) { return $null }

    # Remove non-printable control characters that Graph can reject.
    $clean = ($Description -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', '').Trim()
    if ([string]::IsNullOrWhiteSpace($clean)) { return $null }

    if ($clean.Length -gt 1024) {
        $clean = $clean.Substring(0, 1024)
    }

    return $clean
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

function Resolve-GroupIds {
    param(
        [array]$Ids,
        [hashtable]$IdMap,
        [string]$PolicyName
    )

    if ($null -eq $Ids) { return $Ids }

    $resolved = @()
    foreach ($id in $Ids) {
        $value = [string]$id
        if ([string]::IsNullOrWhiteSpace($value)) { continue }

        if ($IdMap.ContainsKey($value)) {
            $resolved += [string]$IdMap[$value]
            continue
        }

        if (Test-Guid -Value $value) {
            Write-Warn "Policy '$PolicyName': unmapped group ID removed: $value"
            continue
        }
    }

    return @($resolved | Select-Object -Unique)
}

function Resolve-LocationIds {
    param(
        [array]$Ids,
        [hashtable]$IdMap,
        [string]$PolicyName
    )

    if ($null -eq $Ids) { return $Ids }

    $allowedKeywords = @('All', 'AllTrusted')
    $resolved = @()

    foreach ($id in $Ids) {
        $value = [string]$id
        if ([string]::IsNullOrWhiteSpace($value)) { continue }

        if ($allowedKeywords -contains $value) {
            $resolved += $value
            continue
        }

        if ($IdMap.ContainsKey($value)) {
            $resolved += [string]$IdMap[$value]
            continue
        }

        if (Test-Guid -Value $value) {
            Write-Warn "Policy '$PolicyName': unmapped named location ID removed: $value"
            continue
        }
    }

    return @($resolved | Select-Object -Unique)
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
        $safeDescription = Get-SafeGroupDescription -Description $description

        $createBody = @{
            displayName = $displayName
            mailEnabled = $mailEnabled
            mailNickname = $mailNickname
            securityEnabled = $securityEnabled
        }

        if ($safeDescription) {
            $createBody['description'] = $safeDescription
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
                Write-Warn "Could not create service principal for AppId ${appId}: $($_.Exception.Message)"
            }
        }
    }
}

function Build-PolicyBody {
    param(
        [hashtable]$RawPolicy,
        [string]$State,
        [hashtable]$IdMap,
        [string]$PolicyName
    )

    $policy = Remove-GraphMetadata -Object $RawPolicy

    # Enforce desired state on import
    $policy['state'] = $State

    # Remap group IDs
    if ($policy.conditions -and $policy.conditions.users) {
        if ($policy.conditions.users.includeGroups) {
            $policy.conditions.users.includeGroups = Resolve-GroupIds -Ids $policy.conditions.users.includeGroups -IdMap $IdMap -PolicyName $PolicyName
        }
        if ($policy.conditions.users.excludeGroups) {
            $policy.conditions.users.excludeGroups = Resolve-GroupIds -Ids $policy.conditions.users.excludeGroups -IdMap $IdMap -PolicyName $PolicyName
        }
    }

    # Remap named location IDs
    if ($policy.conditions -and $policy.conditions.locations) {
        if ($policy.conditions.locations.includeLocations) {
            $policy.conditions.locations.includeLocations = Resolve-LocationIds -Ids $policy.conditions.locations.includeLocations -IdMap $IdMap -PolicyName $PolicyName
        }
        if ($policy.conditions.locations.excludeLocations) {
            $policy.conditions.locations.excludeLocations = Resolve-LocationIds -Ids $policy.conditions.locations.excludeLocations -IdMap $IdMap -PolicyName $PolicyName
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

            $body = Build-PolicyBody -RawPolicy $rawHash -State $State -IdMap $IdMap -PolicyName $displayName
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
        [string[]]$PolicyNames,
        [string]$Provider,
        [string]$Prefix
    )

    Write-Step "Setting policy state to $State"

    $allPolicies = Get-MgIdentityConditionalAccessPolicy -All

    if ($PolicyNames -and $PolicyNames.Count -gt 0) {
        $targets = $allPolicies | Where-Object { $PolicyNames -contains $_.DisplayName }
    } else {
        $pattern = Get-TargetPolicyMatchPattern -Provider $Provider -Prefix $Prefix
        $targets = $allPolicies | Where-Object { $_.DisplayName -match $pattern }
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

function Remove-ConditionalAccessPolicies {
    param(
        [string]$SinglePolicyName,
        [string[]]$PolicyNames,
        [switch]$RemoveAll,
        [switch]$ForceDelete,
        [string]$Provider,
        [string]$Prefix,
        [string[]]$Prefixes
    )

    Write-Step 'Removing Conditional Access policies'

    $allPolicies = Get-MgIdentityConditionalAccessPolicy -All
    $targets = @()

    if (-not [string]::IsNullOrWhiteSpace($SinglePolicyName)) {
        $targets = $allPolicies | Where-Object { $_.DisplayName -eq $SinglePolicyName }
    } elseif ($RemoveAll) {
        if ($PolicyNames -and $PolicyNames.Count -gt 0) {
            $targets = $allPolicies | Where-Object { $PolicyNames -contains $_.DisplayName }
        } elseif ($Prefixes -and $Prefixes.Count -gt 0) {
            $targets = foreach ($candidate in $allPolicies) {
                foreach ($p in $Prefixes) {
                    if ([string]::IsNullOrWhiteSpace($p)) { continue }
                    if ($candidate.DisplayName.StartsWith($p)) {
                        $candidate
                        break
                    }
                }
            }
            $targets = @($targets | Sort-Object Id -Unique)
        } else {
            $pattern = Get-TargetPolicyMatchPattern -Provider $Provider -Prefix $Prefix
            $targets = $allPolicies | Where-Object { $_.DisplayName -match $pattern }
        }
    } else {
        throw 'For Action=RemovePolicy, use -PolicyName or -RemoveAllBaselinePolicies.'
    }

    if (-not $targets -or $targets.Count -eq 0) {
        Write-Warn 'No matching conditional access policies found to remove.'
        return
    }

    if (-not $ForceDelete) {
        Write-Host ''
        Write-Warn 'About to remove the following policies:'
        $targets | Sort-Object DisplayName | ForEach-Object { Write-Host " - $($_.DisplayName)" }
        $confirm = Read-Host 'Type YES to continue'
        if ($confirm -ne 'YES') {
            Write-Warn 'Removal cancelled.'
            return
        }
    }

    $removed = 0
    $failed = 0

    foreach ($policy in $targets) {
        try {
            if ($PSCmdlet.ShouldProcess($policy.DisplayName, 'Remove conditional access policy')) {
                Remove-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id
                Write-Ok "Removed policy: $($policy.DisplayName)"
                $removed++
            }
        } catch {
            Write-Warn "Failed to remove policy '$($policy.DisplayName)': $($_.Exception.Message)"
            $failed++
        }
    }

    Write-Host ''
    Write-Host 'Removal summary:' -ForegroundColor Cyan
    Write-Host "  Removed : $removed"
    Write-Host "  Failed  : $failed"
}

function Get-BaselineGroupNames {
    param([string]$RootPath)

    if (-not [string]::IsNullOrWhiteSpace($RootPath)) {
        $groupsPath = Join-Path $RootPath 'Config/Groups'
        if (Test-Path -Path $groupsPath -PathType Container) {
            $names = Get-ChildItem -Path $groupsPath -Filter '*.json' -File |
                ForEach-Object {
                    try {
                        (Get-JsonFileObject -Path $_.FullName).displayName
                    } catch {
                        $null
                    }
                } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            if ($names) { return @($names | Select-Object -Unique) }
        }
    }

    return @(
        'APP_Microsoft365_E5'
        'CA-ServiceAccounts'
        'CA-BreakGlassAccounts - Exclude'
    )
}

function Get-BaselineNamedLocationNames {
    param([string]$RootPath)

    if (-not [string]::IsNullOrWhiteSpace($RootPath)) {
        $locationsPath = Join-Path $RootPath 'Config/NamedLocations'
        if (Test-Path -Path $locationsPath -PathType Container) {
            $names = Get-ChildItem -Path $locationsPath -Filter '*.json' -File |
                ForEach-Object {
                    try {
                        (Get-JsonFileObject -Path $_.FullName).displayName
                    } catch {
                        $null
                    }
                } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            if ($names) { return @($names | Select-Object -Unique) }
        }
    }

    return @(
        'ALLOWED COUNTRIES'
        'ALLOWED COUNTRIES - SERVICE ACCOUNTS'
    )
}

function Remove-BaselineGroups {
    param(
        [string[]]$GroupNames,
        [switch]$ForceDelete
    )

    Write-Step 'Removing baseline groups'

    $allGroups = Get-MgGroup -All
    $targets = @()

    if ($GroupNames -and $GroupNames.Count -gt 0) {
        $targets += $allGroups | Where-Object { $GroupNames -contains $_.DisplayName }
    }

    # Fallback pattern for CA exclusion groups
    $targets += $allGroups | Where-Object { $_.DisplayName -match '^CA\d{3}-.* - Exclude$' }
    $targets = @($targets | Sort-Object Id -Unique)

    if (-not $targets -or $targets.Count -eq 0) {
        Write-Warn 'No matching groups found to remove.'
        return
    }

    if (-not $ForceDelete) {
        Write-Warn 'Group removal requested without -Force; groups will still require YES confirmation in remove action.'
    }

    $removed = 0
    $failed = 0

    foreach ($group in $targets) {
        try {
            if ($PSCmdlet.ShouldProcess($group.DisplayName, 'Remove group')) {
                Remove-MgGroup -GroupId $group.Id
                Write-Ok "Removed group: $($group.DisplayName)"
                $removed++
            }
        } catch {
            Write-Warn "Failed to remove group '$($group.DisplayName)': $($_.Exception.Message)"
            $failed++
        }
    }

    Write-Host ''
    Write-Host 'Group removal summary:' -ForegroundColor Cyan
    Write-Host "  Removed : $removed"
    Write-Host "  Failed  : $failed"
}

function Remove-BaselineNamedLocations {
    param([string[]]$LocationNames)

    Write-Step 'Removing baseline named locations'

    $allLocations = Get-MgIdentityConditionalAccessNamedLocation -All
    $targets = @()

    if ($LocationNames -and $LocationNames.Count -gt 0) {
        $targets += $allLocations | Where-Object { $LocationNames -contains $_.DisplayName }
    }

    # Fallback pattern
    $targets += $allLocations | Where-Object { $_.DisplayName -match '^ALLOWED COUNTRIES' }
    $targets = @($targets | Sort-Object Id -Unique)

    if (-not $targets -or $targets.Count -eq 0) {
        Write-Warn 'No matching named locations found to remove.'
        return
    }

    $removed = 0
    $failed = 0

    foreach ($location in $targets) {
        try {
            if ($PSCmdlet.ShouldProcess($location.DisplayName, 'Remove named location')) {
                Remove-MgIdentityConditionalAccessNamedLocation -NamedLocationId $location.Id
                Write-Ok "Removed named location: $($location.DisplayName)"
                $removed++
            }
        } catch {
            Write-Warn "Failed to remove named location '$($location.DisplayName)': $($_.Exception.Message)"
            $failed++
        }
    }

    Write-Host ''
    Write-Host 'Named location removal summary:' -ForegroundColor Cyan
    Write-Host "  Removed : $removed"
    Write-Host "  Failed  : $failed"
}

try {
    Write-Step 'Conditional Access Baseline Import'
    Ensure-RequiredModules -InstallMissing:$InstallMissingModules -Provider $BaselineProvider
    Connect-GraphIfNeeded -Tenant $TenantId

    # Backward-compatible convenience:
    # allow remove switches without explicitly setting -Action RemovePolicy.
    if (
        $Action -eq 'Import' -and (
            -not [string]::IsNullOrWhiteSpace($PolicyName) -or
            $RemoveAllBaselinePolicies -or
            $RemoveAssociatedGroups -or
            $RemoveAssociatedNamedLocations
        )
    ) {
        $Action = 'RemovePolicy'
    }

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

        Set-BaselinePolicyState -State $TargetState -PolicyNames $policyNames -Provider $BaselineProvider -Prefix $BaselinePrefix
        return
    }

    if ($Action -eq 'RemovePolicy') {
        $baselineRootForRemoval = $null
        if ($SourcePath) {
            $baselineRootForRemoval = Resolve-BaselinePath -PathFromParam $SourcePath
        }

        $policyNames = @()
        if ($baselineRootForRemoval) {
            $caPath = Join-Path $baselineRootForRemoval 'Config/ConditionalAccess'
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

        Remove-ConditionalAccessPolicies -SinglePolicyName $PolicyName -PolicyNames $policyNames -RemoveAll:$RemoveAllBaselinePolicies -ForceDelete:$Force -Provider $BaselineProvider -Prefix $BaselinePrefix -Prefixes $RemovePrefixes

        if ($RemoveAssociatedGroups) {
            $groupNames = Get-BaselineGroupNames -RootPath $baselineRootForRemoval
            Remove-BaselineGroups -GroupNames $groupNames -ForceDelete:$Force
        }

        if ($RemoveAssociatedNamedLocations) {
            $locationNames = Get-BaselineNamedLocationNames -RootPath $baselineRootForRemoval
            Remove-BaselineNamedLocations -LocationNames $locationNames
        }

        return
    }

    if ($Action -eq 'Import' -and $BaselineProvider -eq 'Daniel') {
        Deploy-DanielBaseline -DesiredState $PolicyStateOnImport -Prefix $BaselinePrefix -ForceInstallModules:$InstallMissingModules -AutoDeployIds $DanielAutoDeployIds -UseRecommendedIds:$DanielUseRecommendedIds
        Write-Host ''
        Write-Host 'Done. Daniel baseline deployed (Linux-compatible via PowerShell 7 + DCToolbox).' -ForegroundColor Green
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
