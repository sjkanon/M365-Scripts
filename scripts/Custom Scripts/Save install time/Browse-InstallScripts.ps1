#Requires -Version 5.1
<#!
.SYNOPSIS
    Interactive customer/script browser for setup install folders.

.DESCRIPTION
    Shows first-level customer folders as menu items, then lets you browse
    and launch .ps1, .bat, and .cmd files.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $RootPath,

    [Parameter(Mandatory = $false)]
    [string] $SourceLabel = 'Install Scripts'
)

$ErrorActionPreference = 'Stop'
$launchableExtensions = @('.ps1', '.bat', '.cmd')
$excludedDirectories  = @('AppDeployToolkit')

function Test-IsWindowsHost {
    return ($IsWindows -eq $true) -or ($PSVersionTable.PSVersion.Major -le 5)
}

function Test-IsLaunchableFile {
    param([System.IO.FileInfo] $File)

    if (-not $File) { return $false }
    return $launchableExtensions -contains $File.Extension.ToLowerInvariant()
}

function Get-VisibleDirectories {
    param([string] $Path)

    return Get-ChildItem -LiteralPath $Path -Directory -Force -ErrorAction SilentlyContinue |
        Where-Object { $excludedDirectories -notcontains $_.Name } |
        Sort-Object Name
}

function Get-BrowseEntries {
    param([string] $Path)

    $entries = New-Object System.Collections.Generic.List[object]

    $dirs = Get-VisibleDirectories -Path $Path
    foreach ($dir in $dirs) {
        $entries.Add([PSCustomObject]@{
            EntryType = 'Directory'
            Label     = $dir.Name
            Path      = $dir.FullName
        })
    }

    $files = Get-ChildItem -LiteralPath $Path -File -Force -ErrorAction SilentlyContinue |
        Where-Object { Test-IsLaunchableFile -File $_ } |
        Sort-Object Name

    foreach ($file in $files) {
        $entries.Add([PSCustomObject]@{
            EntryType = 'File'
            Label     = $file.Name
            Path      = $file.FullName
        })
    }

    return $entries
}

function Invoke-LaunchableFile {
    param([string] $Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Bestand niet gevonden: $Path"
    }

    $parentPath = Split-Path -Path $Path -Parent
    $fileName   = Split-Path -Path $Path -Leaf
    $extension  = [System.IO.Path]::GetExtension($fileName).ToLowerInvariant()

    switch ($extension) {
        '.ps1' {
            Push-Location -LiteralPath $parentPath
            try {
                & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $Path
            } finally {
                Pop-Location
            }
            return
        }
        '.bat' {
            $command = 'pushd "{0}" && call "{1}" & popd' -f ($parentPath -replace '"', '""'), ($fileName -replace '"', '""')
            & cmd.exe /c $command
            return
        }
        '.cmd' {
            $command = 'pushd "{0}" && call "{1}" & popd' -f ($parentPath -replace '"', '""'), ($fileName -replace '"', '""')
            & cmd.exe /c $command
            return
        }
        default {
            throw "Onbekend script type: $extension"
        }
    }
}

function Show-CustomerMenu {
    param([string] $Root)

    while ($true) {
        $customers = @(Get-VisibleDirectories -Path $Root)

        Clear-Host
        Write-Host ''
        Write-Host '================================================' -ForegroundColor Cyan
        Write-Host "  $SourceLabel - klanten" -ForegroundColor Cyan
        Write-Host '================================================' -ForegroundColor Cyan
        Write-Host ''
        Write-Host "Bron: $Root" -ForegroundColor DarkGray
        Write-Host ''

        if ($customers.Count -eq 0) {
            Write-Host 'Geen klantmappen gevonden.' -ForegroundColor Yellow
            Write-Host ''
            Write-Host '[0] Terug'
            Write-Host ''
            $choice = Read-Host 'Selectie'
            if ($choice -eq '0') { return $null }
            continue
        }

        for ($i = 0; $i -lt $customers.Count; $i++) {
            Write-Host ('[{0}] {1}' -f ($i + 1), $customers[$i].Name)
        }

        Write-Host ''
        Write-Host '[0] Terug'
        Write-Host ''

        $choice = Read-Host 'Selecteer klant'
        if ($choice -eq '0') { return $null }

        $index = 0
        if (-not [int]::TryParse($choice, [ref] $index)) { continue }
        if ($index -lt 1 -or $index -gt $customers.Count) { continue }

        return $customers[$index - 1].FullName
    }
}

function Show-CustomerPathMenu {
    param([string] $CustomerPath)

    $currentPath = $CustomerPath

    while ($true) {
        $entries = @(Get-BrowseEntries -Path $currentPath)
        $relative = if ($currentPath -eq $CustomerPath) { '.' } else { $currentPath.Substring($CustomerPath.Length).TrimStart('\\') }

        Clear-Host
        Write-Host ''
        Write-Host '================================================' -ForegroundColor Cyan
        Write-Host "  Klant: $([System.IO.Path]::GetFileName($CustomerPath))" -ForegroundColor Cyan
        Write-Host '================================================' -ForegroundColor Cyan
        Write-Host ''
        Write-Host "Pad: $relative" -ForegroundColor DarkGray
        Write-Host ''

        if ($entries.Count -eq 0) {
            Write-Host 'Geen uitvoerbare scripts in deze map.' -ForegroundColor Yellow
        } else {
            for ($i = 0; $i -lt $entries.Count; $i++) {
                $kind = if ($entries[$i].EntryType -eq 'Directory') { '[DIR]' } else { '[RUN]' }
                Write-Host ('[{0}] {1} {2}' -f ($i + 1), $kind, $entries[$i].Label)
            }
        }

        Write-Host ''
        if ($currentPath -eq $CustomerPath) {
            Write-Host '[0] Terug naar klanten'
        } else {
            Write-Host '[0] Een map omhoog'
        }
        Write-Host ''

        $choice = Read-Host 'Selectie'
        if ([string]::IsNullOrWhiteSpace($choice)) { continue }

        if ($choice -eq '0') {
            if ($currentPath -eq $CustomerPath) {
                return
            }

            $currentPath = Split-Path -Path $currentPath -Parent
            continue
        }

        $index = 0
        if (-not [int]::TryParse($choice, [ref] $index)) { continue }
        if ($index -lt 1 -or $index -gt $entries.Count) { continue }

        $selected = $entries[$index - 1]
        if ($selected.EntryType -eq 'Directory') {
            $currentPath = $selected.Path
            continue
        }

        Clear-Host
        Write-Host ''
        Write-Host ('Script starten: {0}' -f $selected.Label) -ForegroundColor Cyan
        Write-Host ''
        try {
            Invoke-LaunchableFile -Path $selected.Path
        } catch {
            Write-Host ''
            Write-Host ('[ERROR] {0}' -f $_.Exception.Message) -ForegroundColor Red
        }

        Write-Host ''
        Write-Host 'Druk op een toets om terug te gaan...' -ForegroundColor DarkGray
        $null = [Console]::ReadKey($true)
    }
}

if (-not (Test-IsWindowsHost)) {
    Write-Host ''
    Write-Host '[ERROR] Dit script werkt alleen op Windows.' -ForegroundColor Red
    exit 1
}

if (-not (Test-Path -LiteralPath $RootPath -PathType Container)) {
    Write-Host ''
    Write-Host "[ERROR] Pad niet beschikbaar: $RootPath" -ForegroundColor Red
    Write-Host 'Druk op een toets om terug te gaan...' -ForegroundColor DarkGray
    $null = [Console]::ReadKey($true)
    exit 1
}

while ($true) {
    $customerPath = Show-CustomerMenu -Root $RootPath
    if (-not $customerPath) {
        break
    }

    Show-CustomerPathMenu -CustomerPath $customerPath
}
