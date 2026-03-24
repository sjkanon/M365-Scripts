<#
.SYNOPSIS
    Test SAS WORK directory health and permissions

.DESCRIPTION
    Validates that G:\sas\work and U:\sas\userwork are accessible and functioning
    Performs basic I/O tests without interfering with running SAS jobs
    
.PARAMETER WorkDir
    SAS WORK directory to test (default: G:\sas\work)

.PARAMETER UserWorkDir
    SAS UserWork directory to test (default: U:\sas\userwork)

.PARAMETER Iterations
    Number of test iterations (default: 10, not 1000!)

.EXAMPLE
    .\Test-SASWorkDirectory.ps1

.EXAMPLE
    .\Test-SASWorkDirectory.ps1 -Iterations 100
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$WorkDir = "G:\sas\work",
    
    [Parameter(Mandatory=$false)]
    [string]$UserWorkDir = "U:\sas\userwork",
    
    [Parameter(Mandatory=$false)]
    [int]$Iterations = 10,
    
    [Parameter(Mandatory=$false)]
    [int]$EventLogHours = 24
)

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "ERROR" { "Red" }
        "WARNING" { "Yellow" }
        "SUCCESS" { "Green" }
        default { "White" }
    }
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Test-DirectoryAccess {
    param([string]$Path)
    
    Write-Log "Testing directory: $Path"
    
    # Check if directory exists
    if (-not (Test-Path $Path)) {
        Write-Log "Directory does not exist: $Path" "ERROR"
        return $false
    }
    
    # Test write permissions
    $testFile = Join-Path $Path "test_$(Get-Date -Format 'yyyyMMddHHmmss')_$([guid]::NewGuid()).tmp"
    
    try {
        # Write test
        "Test content" | Out-File -FilePath $testFile -ErrorAction Stop
        Write-Log "  ✓ Write test passed" "SUCCESS"
        
        # Read test
        $content = Get-Content -Path $testFile -ErrorAction Stop
        if ($content -eq "Test content") {
            Write-Log "  ✓ Read test passed" "SUCCESS"
        } else {
            Write-Log "  ✗ Read test failed - content mismatch" "ERROR"
            return $false
        }
        
        # Delete test
        Remove-Item -Path $testFile -ErrorAction Stop
        Write-Log "  ✓ Delete test passed" "SUCCESS"
        
        return $true
        
    } catch {
        Write-Log "  ✗ I/O test failed: $_" "ERROR"
        
        # Cleanup if file exists
        if (Test-Path $testFile) {
            Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
        }
        
        return $false
    }
}

function Test-DirectoryPerformance {
    param(
        [string]$Path,
        [int]$Iterations
    )
    
    Write-Log "Performance test: $Iterations iterations on $Path"
    
    $results = @{
        Successful = 0
        Failed = 0
        TotalTime = 0
        MinTime = [double]::MaxValue
        MaxTime = 0
    }
    
    for ($i = 1; $i -le $Iterations; $i++) {
        $testFile = Join-Path $Path "perftest_$i.tmp"
        
        try {
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            
            # Write operation
            "Test data for iteration $i" | Out-File -FilePath $testFile
            
            # Append operation
            "Additional line" | Out-File -FilePath $testFile -Append
            
            # Read operation
            $null = Get-Content -Path $testFile
            
            # Delete operation
            Remove-Item -Path $testFile
            
            $sw.Stop()
            $elapsed = $sw.Elapsed.TotalMilliseconds
            
            $results.Successful++
            $results.TotalTime += $elapsed
            
            if ($elapsed -lt $results.MinTime) { $results.MinTime = $elapsed }
            if ($elapsed -gt $results.MaxTime) { $results.MaxTime = $elapsed }
            
            if ($i % 10 -eq 0) {
                Write-Log "  Progress: $i/$Iterations completed" "INFO"
            }
            
        } catch {
            $results.Failed++
            Write-Log "  Iteration $i failed: $_" "WARNING"
            
            # Cleanup
            if (Test-Path $testFile) {
                Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
            }
        }
    }
    
    # Calculate statistics
    if ($results.Successful -gt 0) {
        $avgTime = $results.TotalTime / $results.Successful
        
        Write-Log "Performance Results:" "SUCCESS"
        Write-Log "  Successful: $($results.Successful)/$Iterations"
        Write-Log "  Failed: $($results.Failed)"
        Write-Log "  Avg time: $([math]::Round($avgTime, 2)) ms"
        Write-Log "  Min time: $([math]::Round($results.MinTime, 2)) ms"
        Write-Log "  Max time: $([math]::Round($results.MaxTime, 2)) ms"
        
        # Performance thresholds
        if ($avgTime -gt 100) {
            Write-Log "  WARNING: Average I/O time is high (>100ms)" "WARNING"
        }
        
        if ($results.Failed -gt 0) {
            Write-Log "  WARNING: Some operations failed" "WARNING"
        }
    }
    
    return $results
}

function Get-DiskSpace {
    param([string]$Path)
    
    $drive = (Get-Item $Path).PSDrive.Name
    $disk = Get-PSDrive -Name $drive
    
    $freeGB = [math]::Round($disk.Free / 1GB, 2)
    $usedGB = [math]::Round($disk.Used / 1GB, 2)
    $totalGB = [math]::Round(($disk.Free + $disk.Used) / 1GB, 2)
    $pctFree = [math]::Round(($disk.Free / ($disk.Free + $disk.Used)) * 100, 2)
    
    Write-Log "Disk Space - $drive`:"
    Write-Log "  Total: $totalGB GB"
    Write-Log "  Used: $usedGB GB"
    Write-Log "  Free: $freeGB GB ($pctFree%)"
    
    if ($pctFree -lt 10) {
        Write-Log "  WARNING: Low disk space (<10%)" "WARNING"
        return $false
    }
    
    return $true
}

function Get-EventLogErrors {
    param(
        [int]$Hours = 24,
        [string]$DriveLetter = $null
    )
    
    Write-Log "Checking Event Viewer logs (last $Hours hours)..." "INFO"
    
    $startTime = (Get-Date).AddHours(-$Hours)
    $foundIssues = $false
    
    # System Event Log - Disk errors
    Write-Log "Scanning System log for disk errors..." "INFO"
    
    $diskErrors = Get-WinEvent -FilterHashtable @{
        LogName = 'System'
        StartTime = $startTime
        Level = 1,2,3  # Critical, Error, Warning
    } -ErrorAction SilentlyContinue | Where-Object {
        $_.ProviderName -match 'disk|ntfs|storage|volsnap|filesystem' -or
        $_.Message -match 'disk|volume|file system|I/O|read|write'
    }
    
    if ($diskErrors) {
        $diskErrorCount = ($diskErrors | Measure-Object).Count
        Write-Log "  Found $diskErrorCount disk-related events:" "WARNING"
        
        $diskErrors | Select-Object -First 5 | ForEach-Object {
            Write-Log "  [$($_.TimeCreated)] $($_.ProviderName) - $($_.LevelDisplayName)" "WARNING"
            $firstLine = ($_.Message -split "`n")[0]
            Write-Log "    $firstLine" "WARNING"
        }
        
        if ($diskErrorCount -gt 5) {
            Write-Log "  ... and $($diskErrorCount - 5) more disk errors" "WARNING"
        }
        
        $foundIssues = $true
    } else {
        Write-Log "  ✓ No disk errors found" "SUCCESS"
    }
    
    # Application Event Log - SAS errors
    Write-Log "Scanning Application log for SAS errors..." "INFO"
    
    $sasErrors = Get-WinEvent -FilterHashtable @{
        LogName = 'Application'
        StartTime = $startTime
        Level = 1,2  # Critical, Error only
    } -ErrorAction SilentlyContinue | Where-Object {
        $_.ProviderName -match 'SAS|Stargate' -or
        $_.Message -match 'sas\.bat|sas\.exe|WORK library|authorization level'
    }
    
    if ($sasErrors) {
        $sasErrorCount = ($sasErrors | Measure-Object).Count
        Write-Log "  Found $sasErrorCount SAS-related events:" "WARNING"
        
        $sasErrors | Select-Object -First 5 | ForEach-Object {
            Write-Log "  [$($_.TimeCreated)] $($_.ProviderName) - $($_.LevelDisplayName)" "WARNING"
            $firstLine = ($_.Message -split "`n")[0]
            Write-Log "    $firstLine" "WARNING"
        }
        
        if ($sasErrorCount -gt 5) {
            Write-Log "  ... and $($sasErrorCount - 5) more SAS errors" "WARNING"
        }
        
        $foundIssues = $true
    } else {
        Write-Log "  ✓ No SAS errors found" "SUCCESS"
    }
    
    # Security Event Log - Permission/Access Denied
    Write-Log "Scanning Security log for access denied events..." "INFO"
    
    try {
        $accessDenied = Get-WinEvent -FilterHashtable @{
            LogName = 'Security'
            StartTime = $startTime
            Id = 4656,4663  # File/Object access failures
        } -ErrorAction SilentlyContinue | Where-Object {
            $_.Message -match 'sas\\work|sas\\userwork' -or
            $_.Message -match 'Access Denied'
        }
        
        if ($accessDenied) {
            $accessDeniedCount = ($accessDenied | Measure-Object).Count
            Write-Log "  Found $accessDeniedCount access denied events:" "WARNING"
            
            $accessDenied | Select-Object -First 3 | ForEach-Object {
                Write-Log "  [$($_.TimeCreated)] Access denied event" "WARNING"
            }
            
            $foundIssues = $true
        } else {
            Write-Log "  ✓ No access denied events found" "SUCCESS"
        }
    } catch {
        Write-Log "  (Security log check requires admin privileges)" "INFO"
    }
    
    # Check for specific drive errors if drive letter provided
    if ($DriveLetter) {
        Write-Log "Checking for $DriveLetter`: specific errors..." "INFO"
        
        $driveErrors = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            StartTime = $startTime
        } -ErrorAction SilentlyContinue | Where-Object {
            $_.Message -match "\\Device\\Harddisk.*$DriveLetter" -or
            $_.Message -match "volume $DriveLetter"
        }
        
        if ($driveErrors) {
            Write-Log "  Found drive-specific errors for $DriveLetter`:" "WARNING"
            $driveErrors | Select-Object -First 3 | ForEach-Object {
                Write-Log "  [$($_.TimeCreated)] $($_.ProviderName)" "WARNING"
            }
            $foundIssues = $true
        } else {
            Write-Log "  ✓ No drive-specific errors for $DriveLetter`:" "SUCCESS"
        }
    }
    
    return -not $foundIssues
}

# Main execution
Write-Log "=== SAS Work Directory Health Check ===" "INFO"
Write-Log ""

$allTestsPassed = $true

# Check Event Viewer first
Write-Log "--- Checking Windows Event Logs ---" "INFO"
if (-not (Get-EventLogErrors -Hours $EventLogHours)) {
    Write-Log "Event log check found issues - see details above" "WARNING"
    $allTestsPassed = $false
} else {
    Write-Log "Event logs look clean" "SUCCESS"
}

Write-Log ""

# Test G:\sas\work
Write-Log "--- Testing WORK directory ---" "INFO"
if (-not (Test-DirectoryAccess -Path $WorkDir)) {
    $allTestsPassed = $false
} else {
    if (-not (Get-DiskSpace -Path $WorkDir)) {
        $allTestsPassed = $false
    }
    
    Write-Log ""
    $workResults = Test-DirectoryPerformance -Path $WorkDir -Iterations $Iterations
    
    if ($workResults.Failed -gt 0) {
        $allTestsPassed = $false
    }
}

Write-Log ""

# Test U:\sas\userwork
Write-Log "--- Testing USERWORK directory ---" "INFO"
if (-not (Test-DirectoryAccess -Path $UserWorkDir)) {
    $allTestsPassed = $false
} else {
    if (-not (Get-DiskSpace -Path $UserWorkDir)) {
        $allTestsPassed = $false
    }
    
    Write-Log ""
    $userworkResults = Test-DirectoryPerformance -Path $UserWorkDir -Iterations $Iterations
    
    if ($userworkResults.Failed -gt 0) {
        $allTestsPassed = $false
    }
}

Write-Log ""
Write-Log "=== Test Complete ===" "INFO"

if ($allTestsPassed) {
    Write-Log "All tests PASSED" "SUCCESS"
    exit 0
} else {
    Write-Log "Some tests FAILED" "ERROR"
    exit 1
}
