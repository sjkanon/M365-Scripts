<#
.SYNOPSIS
    Monitor SAS batch job logs for critical errors

.DESCRIPTION
    Scans SAS sequence and job logs for known error patterns:
    - "Can't spawn sas.bat" errors
    - Authorization level errors for library WORK
    - Other critical SAS batch failures
    
    Designed for integration with Zabbix or other monitoring systems

.PARAMETER LogDirectory
    Root directory containing SAS logs (searches recursively)

.PARAMETER DaysToCheck
    Number of days back to check logs (default: 7)

.PARAMETER OutputFormat
    Output format: JSON, Text, or Zabbix (default: Text)

.EXAMPLE
    .\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -DaysToCheck 7

.EXAMPLE
    .\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -OutputFormat Zabbix
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
    [string]$OutputFile = $null
)

# Error patterns to detect
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
            $output += "SAS Batch Error Report - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
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
            
            # Group by error type
            $output += "ERRORS BY TYPE:"
            $output += "-" * 40
            $errorsByType = $Errors | Group-Object -Property ErrorType
            foreach ($group in $errorsByType) {
                $output += ""
                $output += "[$($group.Name)] - $($group.Count) occurrence(s)"
                $output += "  Description: $($ErrorPatterns[$group.Name].Description)"
                $output += ""
                
                foreach ($error in ($group.Group | Sort-Object Timestamp -Descending | Select-Object -First 5)) {
                    $output += "  File: $($error.File)"
                    $output += "  Time: $($error.Timestamp.ToString('yyyy-MM-dd HH:mm:ss'))"
                    $output += "  Line: $($error.ErrorLine)"
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
    Write-Log "Starting SAS batch error monitoring"
    Write-Log "Log Directory: $LogDirectory"
    Write-Log "Days to Check: $DaysToCheck"
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
    
    if ($logFiles.Count -eq 0) {
        Write-Log "No log files found to analyze" "WARNING"
        if ($OutputFormat -eq 'Zabbix') {
            Write-Output "0"
        }
        exit 0
    }
    
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
