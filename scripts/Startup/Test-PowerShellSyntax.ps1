#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Path = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),

    [switch]$Recurse,

    [switch]$IncludePsm1
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-TargetFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,
        [Parameter(Mandatory = $true)]
        [bool]$Recursive,
        [Parameter(Mandatory = $true)]
        [bool]$IncludeModules
    )

    if (-not (Test-Path -LiteralPath $InputPath)) {
        throw "Path not found: $InputPath"
    }

    $item = Get-Item -LiteralPath $InputPath
    if ($item.PSIsContainer) {
        $patterns = @('*.ps1')
        if ($IncludeModules) {
            $patterns += '*.psm1'
        }

        $files = Get-ChildItem -LiteralPath $item.FullName -File -Recurse:$Recursive |
            Where-Object {
                foreach ($pattern in $patterns) {
                    if ($_.Name -like $pattern) { return $true }
                }
                return $false
            } |
            Sort-Object FullName

        return $files
    }

    if ($item.Extension -in '.ps1', '.psm1') {
        return ,$item
    }

    throw "Unsupported file type: $($item.FullName). Use .ps1 or .psm1"
}

function Test-PowerShellFileSyntax {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    $tokens = $null
    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($FilePath, [ref]$tokens, [ref]$parseErrors)

    [PSCustomObject]@{
        FilePath = $FilePath
        Errors   = @($parseErrors)
        IsValid  = (@($parseErrors).Count -eq 0)
    }
}

try {
    $targetFiles = Get-TargetFiles -InputPath $Path -Recursive:$Recurse -IncludeModules:$IncludePsm1
} catch {
    Write-Error $_
    exit 2
}

if (-not $targetFiles -or $targetFiles.Count -eq 0) {
    Write-Warning "No PowerShell files found for: $Path"
    exit 0
}

Write-Host "Checking $($targetFiles.Count) file(s)..." -ForegroundColor Cyan

$results = foreach ($file in $targetFiles) {
    Test-PowerShellFileSyntax -FilePath $file.FullName
}

$invalid = $results | Where-Object { -not $_.IsValid }

if (-not $invalid) {
    Write-Host "No syntax errors found." -ForegroundColor Green
    exit 0
}

Write-Host "Syntax errors found in $($invalid.Count) file(s):" -ForegroundColor Red
foreach ($entry in $invalid) {
    Write-Host "`n$($entry.FilePath)" -ForegroundColor Yellow
    foreach ($err in $entry.Errors) {
        Write-Host ("  Line {0}, Col {1}: {2}" -f $err.Extent.StartLineNumber, $err.Extent.StartColumnNumber, $err.Message)
    }
}

exit 1
