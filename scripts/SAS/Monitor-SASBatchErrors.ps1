<#
.SYNOPSIS
    Comprehensive SAS batch job monitoring with Event Viewer integration

.DESCRIPTION
    Scans both SAS log files AND Windows Event Viewer for:
    - "Can't spawn sas.bat" errors
    - Authorization level errors for library WORK
    - Disk and filesystem errors
    - SAS service failures
    - Permission/access denied events
    
    Designed for integration with Zabbix or other monitoring systems

.PARAMETER LogDirectory
    Root directory containing SAS logs (searches recursively)

.PARAMETER DaysToCheck
    Number of days back to check logs (default: 7)

.PARAMETER OutputFormat
    Output format: JSON, Text, or Zabbix (default: Text)

.PARAMETER IncludeEventLog
    Include Windows Event Viewer analysis

.PARAMETER EventLogHours
    Hours to check in Event Viewer (default: 24)

.EXAMPLE
    .\Monitor-SASBatchErrors-Enhanced.ps1 -LogDirectory "E:\SAS\Logs" -DaysToCheck 7

.EXAMPLE
    .\Monitor-SASBatchErrors-Enhanced.ps1 -LogDirectory "E:\SAS\Logs" -IncludeEventLog

.EXAMPLE
    .\Monitor-SASBatchErrors-Enhanced.ps1 -LogDirectory "E:\SAS\Logs" -OutputFormat Zabbix -IncludeEventLog
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$LogDirectory,
    
    [Parameter(Mandatory=$false)]
    [int]$DaysToCheck = 7,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet('JSON','Text','Zabbix')]
    [string]$OutputFormat = 'Text',
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFile = $null,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeEventLog,
    
    [Parameter(Mandatory=$false)]
    [int]$EventLogHours = 24
)

# Error patterns to detect in log files
$ErrorPatterns = @{
    'SpawnError' = @{
        Pattern = "Can't spawn.*sas\.bat"
        Severity = 'Critical'
        Description = 'SAS.BAT spawning failure'
    }
    'WorkLibAuth' = @{
        Pattern = 'User does not have appropriate authorization level for library WORK'
        Severity = 'Critical'
        Description = 'WORK library authorization error'
    }
    'SQLViewError' = @{
        Pattern = 'SQL view was not defined due to errors'
        Severity = 'High'
        Description = 'SQL view definition failure'
    }
    'SASAbort' = @{
        Pattern = 'SAS has ABORTED processing'
        Severity = 'Critical'
        Description = 'SAS processing aborted'
    }
    'GeneralError' = @{
        Pattern = '^ERROR:'
        Severity = 'Medium'
        Description = 'General SAS error'
    }
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message"
}

function Get-SASLogFiles {
    param(
        [string]$Path,
        [int]$Days
    )
    
    $cutoffDate = (Get-Date).AddDays(-$Days)
    
    Write-Log "Scanning for SAS log files in: $Path"
    Write-Log "Checking files modified after: $cutoffDate"
    
    $logFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { 
            ($_.Extension -eq '.log') -and 
            ($_.LastWriteTime -gt $cutoffDate)
        }
    
    Write-Log "Found $($logFiles.Count) log files to analyze"
    
    return $logFiles
}

function Test-ErrorPattern {
    param(
        [string]$Line,
        [hashtable]$Pattern
    )
    
    if ($Line -match $Pattern.Pattern) {
        return $true
    }
    return $false
}

function Parse-SASLog {
    param(
        [System.IO.FileInfo]$LogFile
    )
    
    $errors = @()
    $lineNumber = 0
    
    try {
        $content = Get-Content -Path $LogFile.FullName -ErrorAction Stop
        
        foreach ($line in $content) {
            $lineNumber++
            
            foreach ($errorType in $ErrorPatterns.Keys) {
                if (Test-ErrorPattern -Line $line -Pattern $ErrorPatterns[$errorType]) {
                    $errors += [PSCustomObject]@{
                        File = $LogFile.Name
                        FullPath = $LogFile.FullName
                        LineNumber = $lineNumber
                        Timestamp = $LogFile.LastWriteTime
                        ErrorType = $errorType
                        Severity = $ErrorPatterns[$errorType].Severity
                        Description = $ErrorPatterns[$errorType].Description
                        ErrorLine = $line.Trim()
                        Source = 'LogFile'
                    }
                }
            }
        }
    }
    catch {
        Write-Log "Error reading file $($LogFile.FullName): $_" "ERROR"
    }
    
    return $errors
}

function Get-EventLogErrors {
    param([int]$Hours)
    
    Write-Log "Checking Windows Event Viewer (last $Hours hours)..."
    
    $startTime = (Get-Date).AddHours(-$Hours)
    $eventErrors = @()
    
    # System Log - Disk/Storage errors
    try {
        $diskEvents = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            StartTime = $startTime
            Level = 1,2  # Critical, Error
        } -ErrorAction SilentlyContinue | Where-Object {
            $_.ProviderName -match 'disk|ntfs|storage|volsnap' -or
            $_.Message -match 'disk|volume|file system|I/O error'
        }
        
        foreach ($event in $diskEvents) {
            $eventErrors += [PSCustomObject]@{
                File = "EventLog:System"
                FullPath = "System Event Log"
                LineNumber = $event.Id
                Timestamp = $event.TimeCreated
                ErrorType = 'DiskError'
                Severity = 'Critical'
                Description = 'Disk/Storage system error'
                ErrorLine = ($event.Message -split "`n")[0]
                Source = 'EventLog'
            }
        }
        
        Write-Log "  Found $($diskEvents.Count) disk/storage errors"
        
    } catch {
        Write-Log "  Could not read System event log: $_" "WARNING"
    }
    
    # Application Log - SAS errors
    try {
        $sasEvents = Get-WinEvent -FilterHashtable @{
            LogName = 'Application'
            StartTime = $startTime
            Level = 1,2  # Critical, Error
        } -ErrorAction SilentlyContinue | Where-Object {
            $_.ProviderName -match 'SAS|Stargate' -or
            $_.Message -match 'sas\.bat|sas\.exe|WORK library|authorization'
        }
        
        foreach ($event in $sasEvents) {
            # Determine error type from message
            $errorType = 'SASEventError'
            $severity = 'High'
            
            if ($event.Message -match "Can't spawn.*sas\.bat") {
                $errorType = 'SpawnError'
                $severity = 'Critical'
            } elseif ($event.Message -match 'authorization level for library WORK') {
                $errorType = 'WorkLibAuth'
                $severity = 'Critical'
            }
            
            $eventErrors += [PSCustomObject]@{
                File = "EventLog:Application"
                FullPath = "Application Event Log"
                LineNumber = $event.Id
                Timestamp = $event.TimeCreated
                ErrorType = $errorType
                Severity = $severity
                Description = $event.ProviderName
                ErrorLine = ($event.Message -split "`n")[0]
                Source = 'EventLog'
            }
        }
        
        Write-Log "  Found $($sasEvents.Count) SAS-related errors"
        
    } catch {
        Write-Log "  Could not read Application event log: $_" "WARNING"
    }
    
    # Security Log - Access Denied
    try {
        $accessEvents = Get-WinEvent -FilterHashtable @{
            LogName = 'Security'
            StartTime = $startTime
            Id = 4656,4663  # Access failure events
        } -ErrorAction SilentlyContinue | Where-Object {
            $_.Message -match 'sas\\work|sas\\userwork|Access Denied'
        }
        
        foreach ($event in $accessEvents) {
            $eventErrors += [PSCustomObject]@{
                File = "EventLog:Security"
                FullPath = "Security Event Log"
                LineNumber = $event.Id
                Timestamp = $event.TimeCreated
                ErrorType = 'AccessDenied'
                Severity = 'Critical'
                Description = 'File access denied'
                ErrorLine = ($event.Message -split "`n")[0]
                Source = 'EventLog'
            }
        }
        
        Write-Log "  Found $($accessEvents.Count) access denied events"
        
    } catch {
        Write-Log "  Could not read Security log (requires admin)" "INFO"
    }
    
    return $eventErrors
}

function Format-Output {
    param(
        [array]$Errors,
        [string]$Format
    )
    
    switch ($Format) {
        'JSON' {
            $summary = @{
                Timestamp = Get-Date -Format "o"
                TotalErrors = $Errors.Count
                CriticalErrors = ($Errors | Where-Object { $_.Severity -eq 'Critical' }).Count
                HighErrors = ($Errors | Where-Object { $_.Severity -eq 'High' }).Count
                MediumErrors = ($Errors | Where-Object { $_.Severity -eq 'Medium' }).Count
                LogFileErrors = ($Errors | Where-Object { $_.Source -eq 'LogFile' }).Count
                EventLogErrors = ($Errors | Where-Object { $_.Source -eq 'EventLog' }).Count
                Errors = $Errors
            }
            return ($summary | ConvertTo-Json -Depth 10)
        }
        
        'Zabbix' {
            # Zabbix discovery format and error count
            $criticalCount = ($Errors | Where-Object { $_.Severity -eq 'Critical' }).Count
            return $criticalCount
        }
        
        'Text' {
            $output = @()
            $output += "=" * 80
            $output += "SAS Batch Error Report (Enhanced) - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            $output += "=" * 80
            $output += ""
            
            if ($Errors.Count -eq 0) {
                $output += "✓ No errors found in the specified time period"
                return $output -join "`n"
            }
            
            # Summary
            $output += "SUMMARY:"
            $output += "-" * 40
            $output += "Total Errors: $($Errors.Count)"
            $output += "Critical: $(($Errors | Where-Object { $_.Severity -eq 'Critical' }).Count)"
            $output += "High: $(($Errors | Where-Object { $_.Severity -eq 'High' }).Count)"
            $output += "Medium: $(($Errors | Where-Object { $_.Severity -eq 'Medium' }).Count)"
            $output += ""
            $output += "Sources:"
            $output += "  Log Files: $(($Errors | Where-Object { $_.Source -eq 'LogFile' }).Count)"
            $output += "  Event Viewer: $(($Errors | Where-Object { $_.Source -eq 'EventLog' }).Count)"
            $output += ""
            
            # Group by error type
            $output += "ERRORS BY TYPE:"
            $output += "-" * 40
            $errorsByType = $Errors | Group-Object -Property ErrorType
            foreach ($group in $errorsByType) {
                $output += ""
                $output += "[$($group.Name)] - $($group.Count) occurrence(s)"
                
                if ($ErrorPatterns.ContainsKey($group.Name)) {
                    $output += "  Description: $($ErrorPatterns[$group.Name].Description)"
                }
                
                $output += ""
                
                foreach ($error in ($group.Group | Sort-Object Timestamp -Descending | Select-Object -First 5)) {
                    $output += "  Source: $($error.Source)"
                    $output += "  File/Log: $($error.File)"
                    $output += "  Time: $($error.Timestamp.ToString('yyyy-MM-dd HH:mm:ss'))"
                    $output += "  Message: $($error.ErrorLine)"
                    $output += "  ---"
                }
                
                if ($group.Count -gt 5) {
                    $output += "  ... and $($group.Count - 5) more"
                }
            }
            
            $output += ""
            $output += "=" * 80
            
            return $output -join "`n"
        }
    }
}

# Main execution
try {
    Write-Log "Starting enhanced SAS batch error monitoring"
    Write-Log "Log Directory: $LogDirectory"
    Write-Log "Days to Check: $DaysToCheck"
    Write-Log "Include Event Log: $IncludeEventLog"
    Write-Log "Output Format: $OutputFormat"
    
    # Check if directory exists
    if (-not (Test-Path $LogDirectory)) {
        Write-Log "ERROR: Log directory does not exist: $LogDirectory" "ERROR"
        if ($OutputFormat -eq 'Zabbix') {
            Write-Output "-1"  # Zabbix error code
        }
        exit 1
    }
    
    # Get log files
    $logFiles = Get-SASLogFiles -Path $LogDirectory -Days $DaysToCheck
    
    # Parse all log files
    Write-Log "Analyzing log files..."
    $allErrors = @()
    
    foreach ($logFile in $logFiles) {
        Write-Log "Processing: $($logFile.Name)" "DEBUG"
        $errors = Parse-SASLog -LogFile $logFile
        if ($errors.Count -gt 0) {
            $allErrors += $errors
            Write-Log "  Found $($errors.Count) error(s)" "WARNING"
        }
    }
    
    # Check Event Viewer if requested
    if ($IncludeEventLog) {
        Write-Log ""
        Write-Log "Checking Windows Event Viewer..."
        $eventErrors = Get-EventLogErrors -Hours $EventLogHours
        
        if ($eventErrors.Count -gt 0) {
            $allErrors += $eventErrors
            Write-Log "  Found $($eventErrors.Count) event log error(s)" "WARNING"
        }
    }
    
    Write-Log "Analysis complete. Total errors found: $($allErrors.Count)"
    
    # Format and output results
    $output = Format-Output -Errors $allErrors -Format $OutputFormat
    
    if ($OutputFile) {
        $output | Out-File -FilePath $OutputFile -Encoding UTF8
        Write-Log "Results written to: $OutputFile"
    }
    
    Write-Output $output
    
    # Exit code based on severity
    if (($allErrors | Where-Object { $_.Severity -eq 'Critical' }).Count -gt 0) {
        exit 2  # Critical errors found
    }
    elseif (($allErrors | Where-Object { $_.Severity -eq 'High' }).Count -gt 0) {
        exit 1  # High severity errors found
    }
    else {
        exit 0  # No critical/high errors
    }
    
} catch {
    Write-Log "Script execution failed: $_" "ERROR"
    Write-Log $_.ScriptStackTrace "ERROR"
    if ($OutputFormat -eq 'Zabbix') {
        Write-Output "-1"
    }
    exit 1
}
