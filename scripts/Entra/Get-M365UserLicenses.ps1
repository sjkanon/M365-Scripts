#Requires -Version 5.1
<#
.SYNOPSIS
    Check assigned M365 licenses for a list of users.

.DESCRIPTION
    Reads users from -UserList, CSV, or TXT and retrieves assigned licenses from
    Microsoft Graph. Exports a CSV report with one row per user-license.

    Input options:
    - -UserList "user1@contoso.com","user2@contoso.com"
    - -CsvPath .\users.csv  (column: UserPrincipalName, UPN, or Mail)
    - -CsvPath .\users.txt  (one UPN/mail per line)

.PARAMETER UserList
    Array of UPNs or primary email addresses.

.PARAMETER CsvPath
    Path to CSV or TXT file containing users.

.PARAMETER OutputPath
    Path to the CSV report output.
    Default: C:\Temp\UserLicenseReport_<timestamp>.csv

.PARAMETER TenantId
    Optional tenant ID or domain for Connect-MgGraph.

.EXAMPLE
    .\Get-M365UserLicenses.ps1 -UserList "user1@contoso.com","user2@contoso.com"

.EXAMPLE
    .\Get-M365UserLicenses.ps1 -CsvPath .\users.csv

.EXAMPLE
    "user1@contoso.com","user2@contoso.com" | .\Get-M365UserLicenses.ps1
#>

[CmdletBinding(DefaultParameterSetName = 'List')]
param (
    [Parameter(ParameterSetName = 'List', ValueFromPipeline, Position = 0)]
    [string[]] $UserList,

    [Parameter(ParameterSetName = 'File', Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $CsvPath,

    [string] $OutputPath,

    [string] $TenantId
)

begin {
    $ErrorActionPreference = 'Stop'

    function Resolve-UserFromInput {
        param(
            [Parameter(Mandatory)]
            [string] $InputValue
        )

        $props = 'Id,DisplayName,UserPrincipalName,Mail,AssignedLicenses,AccountEnabled'

        try {
            return Get-MgUser -UserId $InputValue -Property $props -ErrorAction Stop
        } catch {
            if ($_.Exception.Message -notmatch 'Request_ResourceNotFound|does not exist') {
                throw
            }
        }

        $escaped = $InputValue.Replace("'", "''")
        $filter = "userPrincipalName eq '$escaped' or mail eq '$escaped'"
        $matches = Get-MgUser -Filter $filter -Property $props -Top 2 -ConsistencyLevel eventual -ErrorAction Stop

        if (-not $matches) {
            return $null
        }

        return @($matches)[0]
    }

    # Output defaults to a location that exists on most Windows endpoints.
    $outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
    if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
    if (-not $OutputPath) { $OutputPath = Join-Path $outputDir "UserLicenseReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" }

    $script:ConnectedHere = $false
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    $needsConnect = (-not $ctx)

    if ($ctx -and $TenantId -and $ctx.TenantId -ne $TenantId) {
        try { Disconnect-MgGraph | Out-Null } catch {}
        $needsConnect = $true
    }

    if ($needsConnect) {
        Write-Host ''
        Write-Host '  Connecting to Microsoft Graph...' -ForegroundColor Cyan
        $connectParams = @{
            Scopes    = @('User.Read.All', 'Organization.Read.All')
            NoWelcome = $true
        }
        if ($TenantId) { $connectParams['TenantId'] = $TenantId }
        Connect-MgGraph @connectParams
        $script:ConnectedHere = $true
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
    }

    if ($ctx) {
        Write-Host "  Graph context      : $($ctx.Account) | Tenant: $($ctx.TenantId)" -ForegroundColor DarkCyan
    }

    $allUsers = [System.Collections.Generic.List[string]]::new()

    if ($PSCmdlet.ParameterSetName -eq 'File') {
        $ext = [System.IO.Path]::GetExtension($CsvPath).ToLowerInvariant()

        if ($ext -eq '.csv') {
            $raw = Import-Csv -Path $CsvPath
            if (-not $raw -or $raw.Count -eq 0) {
                throw 'CSV is empty.'
            }

            $cols = $raw[0].PSObject.Properties.Name
            $col = $cols | Where-Object { $_ -match '^userprincipalname$|^upn$|^mail$' } | Select-Object -First 1
            if (-not $col) {
                throw "CSV must have a 'UserPrincipalName', 'UPN', or 'Mail' column. Found: $($cols -join ', ')"
            }

            foreach ($entry in $raw.$col) {
                if ($entry -and $entry.Trim() -match '@') {
                    $allUsers.Add($entry.Trim())
                }
            }
        } else {
            Get-Content -Path $CsvPath |
                Where-Object { $_ -and $_.Trim() -and $_ -notmatch '^\s*#' -and $_ -match '@' } |
                ForEach-Object { $allUsers.Add($_.Trim()) }
        }
    }

    $results  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $notFound = [System.Collections.Generic.List[string]]::new()
}

process {
    if ($PSCmdlet.ParameterSetName -eq 'List' -and $UserList) {
        foreach ($entry in $UserList) {
            if ($entry -and $entry.Trim() -match '@') {
                $allUsers.Add($entry.Trim())
            }
        }
    }
}

end {
    try {
        $uniqueUsers = $allUsers | Sort-Object -Unique
        if (-not $uniqueUsers -or $uniqueUsers.Count -eq 0) {
            Write-Warning 'No valid users found. Use -UserList or -CsvPath with valid UPN/email values.'
            return
        }

        Write-Host ''
        Write-Host '  ================================================' -ForegroundColor Cyan
        Write-Host '   M365 User License Check' -ForegroundColor Cyan
        Write-Host '  ================================================' -ForegroundColor Cyan
        Write-Host "  Users to check : $($uniqueUsers.Count)"
        Write-Host ''

        $skuMap = @{}
        foreach ($sku in (Get-MgSubscribedSku -All -ErrorAction Stop)) {
            $skuMap[$sku.SkuId.ToString()] = $sku.SkuPartNumber
        }

        foreach ($userInput in $uniqueUsers) {
            Write-Host "  -> $userInput" -NoNewline

            try {
                $user = Resolve-UserFromInput -InputValue $userInput
                if (-not $user) {
                    Write-Host '  [Not found]' -ForegroundColor DarkYellow
                    $notFound.Add($userInput)
                    $results.Add([PSCustomObject]@{
                        InputValue            = $userInput
                        UserPrincipalName     = ''
                        DisplayName           = ''
                        Mail                  = ''
                        AccountEnabled        = ''
                        LicenseSkuPartNumber  = ''
                        LicenseSkuId          = ''
                        LicenseStatus         = 'NotFound'
                        CheckedAt             = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                    })
                    continue
                }

                $assigned = @($user.AssignedLicenses)
                if ($assigned.Count -eq 0) {
                    Write-Host '  [No licenses]' -ForegroundColor DarkYellow
                    $results.Add([PSCustomObject]@{
                        InputValue            = $userInput
                        UserPrincipalName     = $user.UserPrincipalName
                        DisplayName           = $user.DisplayName
                        Mail                  = $user.Mail
                        AccountEnabled        = $user.AccountEnabled
                        LicenseSkuPartNumber  = ''
                        LicenseSkuId          = ''
                        LicenseStatus         = 'Unlicensed'
                        CheckedAt             = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                    })
                    continue
                }

                Write-Host "  [$($assigned.Count) license(s)]" -ForegroundColor Green
                foreach ($lic in $assigned) {
                    $skuId = $lic.SkuId.ToString()
                    $skuPartNumber = if ($skuMap.ContainsKey($skuId)) { $skuMap[$skuId] } else { 'UnknownSku' }

                    $results.Add([PSCustomObject]@{
                        InputValue            = $userInput
                        UserPrincipalName     = $user.UserPrincipalName
                        DisplayName           = $user.DisplayName
                        Mail                  = $user.Mail
                        AccountEnabled        = $user.AccountEnabled
                        LicenseSkuPartNumber  = $skuPartNumber
                        LicenseSkuId          = $skuId
                        LicenseStatus         = 'Licensed'
                        CheckedAt             = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                    })
                }
            } catch {
                if ($_.Exception.Message -match 'Request_ResourceNotFound|does not exist') {
                    Write-Host '  [Not found]' -ForegroundColor DarkYellow
                    $notFound.Add($userInput)
                    $results.Add([PSCustomObject]@{
                        InputValue            = $userInput
                        UserPrincipalName     = ''
                        DisplayName           = ''
                        Mail                  = ''
                        AccountEnabled        = ''
                        LicenseSkuPartNumber  = ''
                        LicenseSkuId          = ''
                        LicenseStatus         = 'NotFound'
                        CheckedAt             = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                    })
                } else {
                    Write-Host "  [ERROR: $($_.Exception.Message)]" -ForegroundColor Red
                    $results.Add([PSCustomObject]@{
                        InputValue            = $userInput
                        UserPrincipalName     = ''
                        DisplayName           = ''
                        Mail                  = ''
                        AccountEnabled        = ''
                        LicenseSkuPartNumber  = ''
                        LicenseSkuId          = ''
                        LicenseStatus         = "Error: $($_.Exception.Message)"
                        CheckedAt             = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                    })
                }
            }
        }

        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

        $licensedUsers = ($results | Where-Object { $_.LicenseStatus -eq 'Licensed' } | Select-Object -ExpandProperty UserPrincipalName -Unique).Count
        $unlicensedUsers = ($results | Where-Object { $_.LicenseStatus -eq 'Unlicensed' }).Count
        $errorCount = ($results | Where-Object { $_.LicenseStatus -like 'Error:*' }).Count

        Write-Host ''
        Write-Host '  ================================================' -ForegroundColor Cyan
        Write-Host '   Summary' -ForegroundColor Cyan
        Write-Host '  ================================================' -ForegroundColor Cyan
        Write-Host "  Checked users      : $($uniqueUsers.Count)"
        Write-Host "  Licensed users     : $licensedUsers" -ForegroundColor Green
        Write-Host "  Unlicensed users   : $unlicensedUsers" -ForegroundColor DarkYellow
        Write-Host "  Not found          : $($notFound.Count)" -ForegroundColor DarkYellow
        Write-Host "  Errors             : $errorCount" -ForegroundColor Red
        Write-Host ''
        Write-Host "  Report saved to: $OutputPath" -ForegroundColor Cyan
        Write-Host ''
    } finally {
        if ($script:ConnectedHere) {
            Disconnect-MgGraph | Out-Null
        }
    }
}
