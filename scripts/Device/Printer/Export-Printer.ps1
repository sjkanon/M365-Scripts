#Requires -Version 5.1
<#
.SYNOPSIS
    Export the printers of this machine - their settings as a JSON file and their
    drivers - to one zip in C:\Temp, ready for Install-Printer.ps1 on another one.

.DESCRIPTION
    The reverse of Install-Printer.ps1. Reads every printer (or those asked for
    with -Printer) and writes:

      printers.json   The printers and their settings in Install-Printer's own
                      format: driver, address, port name and number, LPR queue,
                      SNMP, location, comment, sharing, duplex, colour, paper size,
                      default tray and permissions. Every driver points at its
                      folder in the zip ("type": "path"), so after unpacking

                          .\Install-Printer.ps1 -ConfigPath <folder>\printers.json

                      installs the same drivers and printers on another machine.
      drivers\        One folder per driver: the complete driver package as the
                      driver store holds it - INF, catalog and binaries - which is
                      what pnputil installs from. Drivers that ship with Windows
                      (ntprint.inf, prnms*.inf) are not copied: every machine has them.
      settings\       Per printer the full default print ticket (<name>.xml), with
                      every vendor setting that the JSON fields do not cover.
      inventory.json  Everything that was read, also for printers the JSON cannot
                      describe: WSD, USB, PDF, OneNote, fax, connections to a print
                      server. Their raw settings are here, with the reason why they
                      are not in printers.json.

    Only TCP/IP printers (Standard TCP/IP Port) go into printers.json, because
    that is what Install-Printer.ps1 can create. A printer it would refuse - a
    name with \ , or !, an address that is neither an IP address nor a host
    name - is listed in inventory.json and reported, not written.

    Permissions are exported only when they differ from the Windows default
    (Everyone may print), as DOMAIN\Group, or as a SID where the name does not
    translate - an Entra ID group is only known by its SID. Administrators,
    SYSTEM, CREATOR OWNER and app identities are left out, as Install-Printer
    always keeps them.

    Nothing on the machine is changed. Print defaults are read in a job with a
    timeout: that is vendor driver code, which hangs on some universal drivers.

    RMM / NinjaOne
    --------------
    Script variables arrive as environment variables and are used when the
    matching parameter is not passed: printer (comma separated), outputPath, and
    the checkboxes noDriver and printBrm.

.PARAMETER Printer
    Only these printers (names, wildcards allowed). Default: all of them.

.PARAMETER OutputPath
    Folder the zip is written to (default: C:\Temp). Created when missing.

.PARAMETER NoDriver
    Write only the settings, no driver files. printers.json then lists no
    drivers, so the target needs them installed already.

.PARAMETER PrintBrm
    Also make a Windows Printer Migration backup (PrintBrm.exe -b) and put the
    .printerExport in the zip - Windows' own all-in-one format, restorable with
    PrintBrm -r or Print Management. Needs administrator rights.

.EXAMPLE
    .\Export-Printer.ps1

    Every printer, its settings and drivers to C:\Temp\PrinterExport_<computer>_<date>.zip.

.EXAMPLE
    .\Export-Printer.ps1 -Printer 'Office*' -NoDriver

    Only the Office printers, settings only.

.EXAMPLE
    .\Export-Printer.ps1 -PrintBrm -OutputPath \\fs01\it$\exports

    Also a PrintBrm backup, the zip written to a share.

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows 10 / Server 2016 or newer (PrintManagement module)
#>
[CmdletBinding()]
param (
    [string[]] $Printer,
    [string]   $OutputPath = 'C:\Temp',
    [switch]   $NoDriver,
    [switch]   $PrintBrm
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -- RMM: run 64-bit -----------------------------------------------------------
# A 32-bit PowerShell sees System32 redirected to SysWOW64, where the driver
# store is not, and Get-PrinterDriver reports the 64-bit drivers' paths there.
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $nativeShell = Join-Path $env:WINDIR 'SysNative\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path $nativeShell)) {
        Write-Error 'Running 32-bit and SysNative is unavailable. Start the script from 64-bit PowerShell.'
        exit 1
    }
    $quote = { param($s) "'" + ([string] $s -replace "'", "''") + "'" }
    $line  = '& ' + (& $quote $PSCommandPath)
    foreach ($entry in $PSBoundParameters.GetEnumerator()) {
        if ($entry.Value -is [switch]) { $line += ' -{0}:${1}' -f $entry.Key, ([bool] $entry.Value).ToString().ToLower() }
        elseif ($entry.Value -is [array]) { $line += ' -{0} {1}' -f $entry.Key, (($entry.Value | ForEach-Object { & $quote $_ }) -join ',') }
        else { $line += ' -{0} {1}' -f $entry.Key, (& $quote $entry.Value) }
    }
    & $nativeShell -NoProfile -ExecutionPolicy Bypass -Command "$line; exit `$LASTEXITCODE"
    exit $LASTEXITCODE
}

# -- RMM: script variables -----------------------------------------------------
$rmmTrue = @('true', '1', 'yes')
if (-not $PSBoundParameters.ContainsKey('NoDriver')   -and $env:noDriver -in $rmmTrue) { $NoDriver   = $true }
if (-not $PSBoundParameters.ContainsKey('PrintBrm')   -and $env:printBrm -in $rmmTrue) { $PrintBrm   = $true }
if (-not $PSBoundParameters.ContainsKey('OutputPath') -and $env:outputPath)            { $OutputPath = $env:outputPath }
if (-not $PSBoundParameters.ContainsKey('Printer')    -and $env:printer) { $Printer = @($env:printer -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($PrintBrm -and -not $isAdmin) {
    Write-Error '-PrintBrm needs administrator rights. Run the script from an elevated session or as System.'
    exit 1
}

$exitCode   = 0
$NativeEnv  = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'Windows ARM64' } else { 'Windows x64' }
$DuplexValues = @('OneSided', 'TwoSidedLongEdge', 'TwoSidedShortEdge')
$InputBinFeatures = @('JobInputBin', 'DocumentInputBin', 'PageInputBin')

function Write-Step { param([string] $Message) Write-Host "  $Message" -ForegroundColor Cyan }
function Write-Ok   { param([string] $Message) Write-Host "  [ OK ] $Message" -ForegroundColor Green }
function Write-Skip { param([string] $Message) Write-Host "  [SKIP] $Message" -ForegroundColor DarkGray }
function Write-Warn { param([string] $Message) Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
function Write-Bad  { param([string] $Message) Write-Host "  [FAIL] $Message" -ForegroundColor Red }

function Get-PropertyValue {
    <# Property access that returns $null instead of tripping Set-StrictMode. #>
    param($Object, [string] $Name)
    if ($null -eq $Object) { return $null }
    if ($Object.PSObject.Properties.Name -notcontains $Name) { return $null }
    return $Object.$Name
}

function Invoke-WithTimeout {
    <# Driver code in a job: a driver that hangs costs one printer's settings, not the export. #>
    param([Parameter(Mandatory)] [scriptblock] $ScriptBlock, [object[]] $ArgumentList = @(), [int] $Seconds = 60)

    $job = Start-Job -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList
    try {
        if (-not (Wait-Job -Job $job -Timeout $Seconds)) {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
            throw "did not finish within $Seconds seconds - the printer driver hangs on it"
        }
        return (Receive-Job -Job $job -ErrorAction Stop)
    } finally {
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }
}

function ConvertFrom-DriverVersion {
    <# Get-PrinterDriver reports the version as one 64-bit number: four 16-bit parts. #>
    param($Value)
    if ($null -eq $Value) { return $null }
    $v = [uint64] $Value
    return [version] ('{0}.{1}.{2}.{3}' -f (($v -shr 48) -band 0xFFFF), (($v -shr 32) -band 0xFFFF),
                                           (($v -shr 16) -band 0xFFFF), ($v -band 0xFFFF))
}

function ConvertTo-SafeName {
    <# A driver or printer name as a file or folder name. #>
    param([Parameter(Mandatory)] [string] $Name)
    $safe = ($Name -replace '[\\/:*?"<>|]', '_').Trim().TrimEnd('.')
    if (-not $safe) { $safe = 'unnamed' }
    return $safe
}

function Test-InboxDriver {
    <# Drivers that come with Windows: every target has them, copying them is only weight. #>
    param([string] $InfPath)
    if (-not $InfPath) { return $true }
    $leaf = Split-Path $InfPath -Leaf
    if ($leaf -match '^(ntprint|prnms)\d*\.inf$') { return $true }
    return ($InfPath -like "$env:WINDIR\INF\*")
}

function Get-PrintDefault {
    <#
        The queue's print defaults, read once in a job: duplex, colour and paper
        as Install-Printer compares them, the default tray by its display name,
        and the full default print ticket.
    #>
    param([Parameter(Mandatory)] [string] $Name)
    return Invoke-WithTimeout -ArgumentList $Name, $InputBinFeatures -ScriptBlock {
        param($PrinterName, $Features)
        $c = Get-PrintConfiguration -PrinterName $PrinterName -ErrorAction Stop
        $psk = 'http://schemas.microsoft.com/windows/2003/08/printing/printschemakeywords'
        $resolve = {
            param($Node, $QName)
            $parts = ([string] $QName).Split(':', 2)
            if ($parts.Count -ne 2) { return $null }
            [PSCustomObject]@{ Uri = $Node.GetNamespaceOfPrefix($parts[0]); Local = $parts[1] }
        }

        # The tray the default ticket selects, by the name the driver shows for it.
        $bin = $null
        try {
            [xml] $ticket = $c.PrintTicketXML
            $current = @{}
            foreach ($feature in $ticket.DocumentElement.ChildNodes) {
                if ($feature.LocalName -ne 'Feature') { continue }
                $f = & $resolve $feature $feature.GetAttribute('name')
                if (-not $f -or $f.Uri -ne $psk -or $f.Local -notin $Features) { continue }
                $option = @($feature.ChildNodes | Where-Object { $_.LocalName -eq 'Option' }) | Select-Object -First 1
                if ($option -and $option.GetAttribute('name')) { $current[$f.Local] = & $resolve $option $option.GetAttribute('name') }
            }
            $selected = foreach ($featureName in $Features) { if ($current.ContainsKey($featureName)) { $current[$featureName]; break } }
            if ($selected) {
                Add-Type -AssemblyName System.Printing
                $server = [System.Printing.LocalPrintServer]::new()
                try {
                    $stream = $server.GetPrintQueue($PrinterName).GetPrintCapabilitiesAsXml()
                    [xml] $caps = ([IO.StreamReader]::new($stream)).ReadToEnd()
                } finally { $server.Dispose() }
                $display = $null
                foreach ($feature in $caps.DocumentElement.ChildNodes) {
                    if ($feature.LocalName -ne 'Feature') { continue }
                    foreach ($option in $feature.ChildNodes) {
                        if ($option.LocalName -ne 'Option' -or -not $option.GetAttribute('name')) { continue }
                        $o = & $resolve $option $option.GetAttribute('name')
                        if (-not $o -or $o.Uri -ne $selected.Uri -or $o.Local -ne $selected.Local) { continue }
                        $display = @($option.ChildNodes | Where-Object { $_.LocalName -eq 'Property' -and $_.GetAttribute('name') -match ':DisplayName$' } |
                                     ForEach-Object { $_.InnerText.Trim() }) | Select-Object -First 1
                        if ($display) { break }
                    }
                    if ($display) { break }
                }
                $bin = if ($display) { $display } else { $selected.Local }
            }
        } catch { }

        [PSCustomObject]@{
            Duplex   = "$($c.DuplexingMode)"
            Color    = [bool] $c.Color
            Paper    = "$($c.PaperSize)"
            InputBin = $bin
            Ticket   = [string] $c.PrintTicketXML
        }
    }
}

function Test-ProtectedSid {
    <# The identities Install-Printer always keeps - not worth exporting. Same list. #>
    param([Parameter(Mandatory)] [string] $Sid)
    if ($Sid -in @('S-1-5-18', 'S-1-5-19', 'S-1-5-20', 'S-1-3-0', 'S-1-3-1',
                   'S-1-5-32-544', 'S-1-5-32-549', 'S-1-5-32-550')) { return $true }
    return ($Sid -like 'S-1-15-*' -or $Sid -like 'S-1-5-99-*' -or $Sid -like 'S-1-5-80-*')
}

function Get-PrinterPermission {
    <#
        The queue's ACL as Install-Printer's "permissions" list, or $null when it
        is the Windows default - only Everyone printing - so a plain queue stays a
        plain entry. Read from the meaning of the access masks, as Install-Printer
        does: PRINTER_ACCESS_ADMINISTER (0x4) manage, JOB_ACCESS_ADMINISTER (0x10)
        on the inherit-only job ACE manageDocuments, PRINTER_ACCESS_USE (0x8) print.
    #>
    param([string] $Sddl)
    if (-not $Sddl) { return $null }

    $rank = @{ print = 1; manageDocuments = 2; manage = 3 }
    $map  = [ordered]@{}
    foreach ($ace in ([Security.AccessControl.RawSecurityDescriptor]::new($Sddl)).DiscretionaryAcl) {
        if ($ace -isnot [Security.AccessControl.KnownAce] -or $ace.AceQualifier -ne 'AccessAllowed') { continue }
        $sid = $ace.SecurityIdentifier.Value
        if (Test-ProtectedSid $sid) { continue }
        $flags       = [int] $ace.AceFlags
        $inheritOnly = ($flags -band [int] [Security.AccessControl.AceFlags]::InheritOnly) -ne 0
        $objInherit  = ($flags -band [int] [Security.AccessControl.AceFlags]::ObjectInherit) -ne 0
        $level = if (-not $inheritOnly -and ($ace.AccessMask -band 0x4)) { 'manage' }
                 elseif ($inheritOnly -and $objInherit -and ($ace.AccessMask -band 0x10)) { 'manageDocuments' }
                 elseif (-not $inheritOnly -and ($ace.AccessMask -band 0x8)) { 'print' }
                 else { $null }
        if (-not $level) { continue }
        if (-not $map.Contains($sid) -or $rank[$level] -gt $rank[$map[$sid]]) { $map[$sid] = $level }
    }
    if ($map.Count -eq 1 -and $map.Contains('S-1-1-0') -and $map['S-1-1-0'] -eq 'print') { return $null }

    $list = foreach ($sid in $map.Keys) {
        # A well-known name translates differently per Windows language
        # ("Iedereen"); Install-Printer reads the English one on any of them.
        $principal = switch ($sid) {
            'S-1-1-0'      { 'Everyone' }
            'S-1-5-11'     { 'Authenticated Users' }
            'S-1-5-32-545' { 'Users' }
            'S-1-5-4'      { 'Interactive' }
            # An Entra ID user or group translates to AzureAD\name here, but that
            # name resolves nowhere else - the SID is the only portable form.
            { $sid -like 'S-1-12-1-*' } { $sid }
            default {
                try { ([Security.Principal.SecurityIdentifier]::new($sid)).Translate([Security.Principal.NTAccount]).Value }
                catch { $sid }
            }
        }
        # A local account of this machine means nothing on the target.
        if ($principal -like "$env:COMPUTERNAME\*") { Write-Warn "  '$principal' is a local account of this machine - the target needs one with the same name" }
        [ordered]@{ principal = $principal; access = $map[$sid] }
    }
    return ,@($list)
}

function Test-PrinterAddress {
    param([string] $Address)
    $ip = $null
    if ([Net.IPAddress]::TryParse($Address, [ref] $ip)) { return $true }
    return ($Address -match '^(?=.{1,253}$)[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?)*$')
}

# -- Run -----------------------------------------------------------------------
$stage = $null
try {
    Write-Host ''
    Write-Step '1. Printers'
    $spooler = Get-Service -Name Spooler -ErrorAction SilentlyContinue
    if (-not $spooler -or $spooler.Status -ne 'Running') { throw 'The Print Spooler is not running - there is nothing to read printers from.' }

    $all = @(Get-Printer -Full -ErrorAction Stop)
    $selected = @(if ($Printer) {
        $all | Where-Object { $name = $_.Name; @($Printer | Where-Object { $name -like $_ }).Count -gt 0 }
    } else { $all })
    if ($Printer -and $selected.Count -eq 0) {
        throw "No printer matches $($Printer -join ', ') (this machine has: $(($all | ForEach-Object { $_.Name }) -join ', '))"
    }
    if ($selected.Count -eq 0) { throw 'This machine has no printers.' }
    Write-Ok "$($selected.Count) printer(s)$(if ($Printer) { " of $($all.Count)" })"

    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $stage = Join-Path ([IO.Path]::GetTempPath()) "PrinterExport_$stamp"
    foreach ($sub in '', 'settings', 'drivers') { New-Item -ItemType Directory -Path (Join-Path $stage $sub) -Force | Out-Null }

    $ports      = @{}
    foreach ($port in @(Get-PrinterPort -ErrorAction SilentlyContinue)) { $ports[$port.Name] = $port }
    $jsonPrinters = [System.Collections.Generic.List[object]]::new()
    $inventory    = [System.Collections.Generic.List[object]]::new()
    $usedDrivers  = [ordered]@{}

    foreach ($p in $selected) {
        Write-Host ''
        Write-Step "Printer '$($p.Name)'"
        $port   = $ports[[string] $p.PortName]
        $reason = $null
        $address = if ($port) { foreach ($f in 'PrinterHostAddress', 'LprHostAddress', 'HostName') { $v = Get-PropertyValue $port $f; if ($v) { [string] $v; break } } }

        if ("$($p.Type)" -eq 'Connection')       { $reason = 'a connection to a print server queue (per user) - Install-Printer creates local TCP/IP queues' }
        elseif (-not $port)                      { $reason = "its port '$($p.PortName)' is not found" }
        elseif (-not $address)                   { $reason = "its port '$($p.PortName)' is not a TCP/IP port ($($port.Description))" }
        elseif (-not (Test-PrinterAddress $address)) { $reason = "address '$address' is neither an IP address nor a host name" }
        elseif ($p.Name -match '[\\,!]')         { $reason = 'its name contains \ , or ! - Install-Printer refuses those' }

        # Print defaults and the full ticket: read for every local queue, also
        # those that cannot go into printers.json - the inventory keeps them.
        $defaults = $null
        if ("$($p.Type)" -ne 'Connection') {
            try {
                $defaults = Get-PrintDefault -Name $p.Name
                $ticketFile = Join-Path $stage ("settings\{0}.xml" -f (ConvertTo-SafeName $p.Name))
                [IO.File]::WriteAllText($ticketFile, $defaults.Ticket, [Text.UTF8Encoding]::new($false))
            } catch {
                Write-Warn "Print defaults not readable: $($_.Exception.Message)"
            }
        }

        $permissions = $null
        try { $permissions = Get-PrinterPermission -Sddl ([string] (Get-PropertyValue $p 'PermissionSDDL')) }
        catch { Write-Warn "Permissions not readable: $($_.Exception.Message)" }

        $inventory.Add([ordered]@{
            name           = $p.Name
            inPrintersJson = -not $reason
            reason         = $reason
            type           = "$($p.Type)"
            driver         = $p.DriverName
            portName       = $p.PortName
            port           = if ($port) { [ordered]@{
                                 description = Get-PropertyValue $port 'Description'
                                 address     = $address
                                 portNumber  = Get-PropertyValue $port 'PortNumber'
                                 protocol    = Get-PropertyValue $port 'Protocol'
                                 lprQueue    = Get-PropertyValue $port 'LprQueueName'
                                 snmp        = Get-PropertyValue $port 'SNMPEnabled'
                             } } else { $null }
            location       = $p.Location
            comment        = $p.Comment
            shared         = [bool] $p.Shared
            shareName      = $p.ShareName
            published      = [bool] (Get-PropertyValue $p 'Published')
            duplex         = if ($defaults) { $defaults.Duplex } else { $null }
            color          = if ($defaults) { $defaults.Color } else { $null }
            paperSize      = if ($defaults) { $defaults.Paper } else { $null }
            inputBin       = if ($defaults) { $defaults.InputBin } else { $null }
            permissionSddl = Get-PropertyValue $p 'PermissionSDDL'
        })

        if ($reason) {
            Write-Skip "Not in printers.json: $reason (settings are in inventory.json)"
            continue
        }

        # The JSON entry, with only what differs from Install-Printer's defaults,
        # so the file reads like one written by hand.
        $entry = [ordered]@{ name = $p.Name; driver = $p.DriverName; address = $address }
        if ($p.PortName -ne "IP_$address") { $entry.portName = $p.PortName }
        $lpr = [string] (Get-PropertyValue $port 'LprQueueName')
        if ("$(Get-PropertyValue $port 'Protocol')" -in @('2', 'LPR') -and $lpr) { $entry.lprQueue = $lpr }
        elseif ((Get-PropertyValue $port 'PortNumber') -and [int] $port.PortNumber -ne 9100) { $entry.portNumber = [int] $port.PortNumber }
        if ((Get-PropertyValue $port 'SNMPEnabled') -eq $true) { $entry.snmp = $true }
        if ($p.Location) { $entry.location = [string] $p.Location }
        if ($p.Comment)  { $entry.comment  = [string] $p.Comment }
        if ($p.Shared) {
            $entry.shared = $true
            if ($p.ShareName -and $p.ShareName -ne $p.Name) { $entry.shareName = [string] $p.ShareName }
        }
        if ($defaults) {
            if ($defaults.Duplex -in $DuplexValues) { $entry.duplex = $defaults.Duplex }
            $entry.color = $defaults.Color
            if ($defaults.Paper -and $defaults.Paper -ne 'Unknown') { $entry.paperSize = $defaults.Paper }
            if ($defaults.InputBin) { $entry.inputBin = $defaults.InputBin }
        }
        if ($null -ne $permissions) { $entry.permissions = $permissions }
        $jsonPrinters.Add($entry)
        $usedDrivers[$p.DriverName] = $true

        $summary = @("$address", $(if ($entry.Contains('duplex')) { $entry.duplex }), $(if ($entry.Contains('paperSize')) { $entry.paperSize }),
                     $(if ($entry.Contains('inputBin')) { "tray '$($entry.inputBin)'" }), $(if ($null -ne $permissions) { "$(@($permissions).Count) permission(s)" })) |
                   Where-Object { $_ }
        Write-Ok ($summary -join ', ')
    }

    # -- 2. Drivers ------------------------------------------------------------
    Write-Host ''
    Write-Step '2. Drivers'
    $jsonDrivers = [System.Collections.Generic.List[object]]::new()
    if ($NoDriver) { Write-Skip 'Not exported (-NoDriver) - the target needs them installed already' }
    elseif ($usedDrivers.Count -eq 0) { Write-Skip 'No printer in printers.json, so no driver to export' }
    foreach ($name in @(if (-not $NoDriver) { $usedDrivers.Keys })) {
        $driver = @(Get-PrinterDriver -Name $name -ErrorAction SilentlyContinue |
                    Where-Object { -not (Get-PropertyValue $_ 'PrinterEnvironment') -or $_.PrinterEnvironment -eq $NativeEnv }) | Select-Object -First 1
        if (-not $driver) { Write-Warn "'$name' is not installed for $NativeEnv - not exported"; continue }
        $inf = [string] (Get-PropertyValue $driver 'InfPath')
        if (Test-InboxDriver $inf) { Write-Skip "'$name' comes with Windows - not copied"; continue }

        $package = Split-Path $inf -Parent
        if (-not (Test-Path -LiteralPath $package)) { Write-Bad "'$name': the driver store folder $package is gone"; $exitCode = 1; continue }
        $folder = 'drivers\' + (ConvertTo-SafeName $name)
        try {
            Copy-Item -LiteralPath $package -Destination (Join-Path $stage $folder) -Recurse -Force
        } catch {
            Write-Bad "'$name': copying $package failed - $($_.Exception.Message)"
            $exitCode = 1
            continue
        }
        $version = ConvertFrom-DriverVersion (Get-PropertyValue $driver 'DriverVersion')
        $size    = (Get-ChildItem -LiteralPath (Join-Path $stage $folder) -Recurse -File | Measure-Object Length -Sum).Sum
        $entry   = [ordered]@{ name = $name }
        if ($version) { $entry.version = $version.ToString() }
        $entry.inf    = Split-Path $inf -Leaf
        $entry.source = [ordered]@{ type = 'path'; path = $folder }
        $jsonDrivers.Add($entry)
        Write-Ok ("'{0}' {1} ({2:N1} MB, {3})" -f $name, $version, ($size / 1MB), (Split-Path $inf -Leaf))
    }

    # -- 3. PrintBrm -----------------------------------------------------------
    if ($PrintBrm) {
        Write-Host ''
        Write-Step '3. PrintBrm backup'
        $brm = Join-Path $env:WINDIR 'System32\spool\tools\PrintBrm.exe'
        if (-not (Test-Path -LiteralPath $brm)) { Write-Warn "PrintBrm.exe is not on this machine ($brm)" }
        else {
            $brmFile = Join-Path $stage "$env:COMPUTERNAME.printerExport"
            $brmLog  = & $brm -b -f $brmFile 2>&1
            if ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $brmFile)) { Write-Ok "$(Split-Path $brmFile -Leaf) - all printers, drivers and ports" }
            else {
                Write-Bad "PrintBrm failed (exit code $LASTEXITCODE): $(($brmLog | Select-Object -Last 3) -join ' ')"
                $exitCode = 1
            }
        }
    }

    # -- 4. Zip ----------------------------------------------------------------
    Write-Host ''
    Write-Step '4. Zip'
    $config = [ordered]@{
        description = "Exported from $env:COMPUTERNAME on $(Get-Date -Format 'yyyy-MM-dd HH:mm') by Export-Printer.ps1"
        drivers     = @($jsonDrivers)
        printers    = @($jsonPrinters)
    }
    $utf8 = [Text.UTF8Encoding]::new($false)
    [IO.File]::WriteAllText((Join-Path $stage 'printers.json'),  ($config | ConvertTo-Json -Depth 8), $utf8)
    [IO.File]::WriteAllText((Join-Path $stage 'inventory.json'), ([ordered]@{
        computer = $env:COMPUTERNAME
        exported = (Get-Date).ToString('s')
        printers = @($inventory)
    } | ConvertTo-Json -Depth 8), $utf8)

    if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
    $zip = Join-Path $OutputPath "PrinterExport_$($env:COMPUTERNAME)_$stamp.zip"
    Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip -CompressionLevel Optimal -Force
    Write-Ok ("{0} ({1:N1} MB)" -f $zip, ((Get-Item -LiteralPath $zip).Length / 1MB))

    Write-Host ''
    $skipped = @($inventory | Where-Object { -not $_.inPrintersJson }).Count
    Write-Host ("  {0} printer(s) and {1} driver(s) in printers.json{2}" -f $jsonPrinters.Count, $jsonDrivers.Count,
                $(if ($skipped) { ", $skipped printer(s) only in inventory.json" })) -ForegroundColor Cyan
    if ($jsonPrinters.Count -gt 0) {
        Write-Host '  On the target: unpack the zip, then' -ForegroundColor Gray
        Write-Host '      .\Install-Printer.ps1 -ConfigPath <folder>\printers.json -CheckOnly' -ForegroundColor Gray
    }
} catch {
    Write-Host ''
    Write-Bad "Aborted: $($_.Exception.Message)"
    $exitCode = 1
} finally {
    if ($stage) { Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host ''
exit $exitCode
