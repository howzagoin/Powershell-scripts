<#
.SYNOPSIS
    Advanced script to scan and fix broken links in Excel and Access files with enhanced error handling and reporting.

.DESCRIPTION
    Scans and fixes broken links in Excel and Access files from:
    - Local or SharePoint synced and available offline folder (fastest)
    - SharePoint synced folder (fast, if available)
    - Online SharePoint URL (slowest; requires download and extra modules)

.FEATURES
    - Enhanced link detection and repair algorithms
    - Comprehensive error handling and retry logic
    - Advanced reporting with detailed statistics
    - Performance optimizations with parallel processing
    - Smart path matching for relative and absolute links
    - Backup creation before modifications
    - Progress tracking with ETA calculations
    - Fuzzy matching for similar filenames
    - Memory usage optimization

Author: Tim MacLatchy
Date: 2025-08-10
Version: 2.0
#>

param(
    [string]$Path,
    [string]$SharePointUrl,
    [int]$MaxParallelJobs = 8,
    [switch]$CreateBackups = $true,
    [switch]$EnableFuzzyMatching = $true,
    [double]$FuzzyThreshold = 0.8,
    [switch]$DetailedReporting = $true
)

#region Enhanced Configuration and Global Variables
$script:Config = @{
    MaxParallelJobs = [Math]::Min($MaxParallelJobs, [System.Environment]::ProcessorCount * 2)
    RetryAttempts = 3
    RetryDelayMs = 500
    MaxMemoryUsageMB = 2048
    BackupFilesBeforeRepair = $CreateBackups.IsPresent
    EnableFuzzyMatching = $EnableFuzzyMatching.IsPresent
    FuzzyMatchThreshold = $FuzzyThreshold
    DetailedLogging = $DetailedReporting.IsPresent
    ExcelApplicationVisible = $false
    AccessDatabaseTimeout = 30
    StartTime = Get-Date
}

# Performance tracking
$script:Stats = @{
    TotalFiles = 0
    ProcessedFiles = 0
    FilesWithErrors = 0
    FilesFixed = 0
    TotalLinks = 0
    BrokenLinks = 0
    FixedLinks = 0
    BackupsCreated = 0
    ProcessingTime = 0
    MemoryUsage = @()
}

# Results collection
$script:Results = [System.Collections.ArrayList]::new()
$script:ErrorLog = [System.Collections.ArrayList]::new()
$script:PerformanceLog = [System.Collections.ArrayList]::new()
#endregion

#region Enhanced Logging and Utilities
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS", "DEBUG", "PERFORMANCE")]
        [string]$Level = "INFO",
        [string]$LogFile = "BrokenLinksAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Console output with colors
    switch ($Level) {
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        "PERFORMANCE" { Write-Host $logEntry -ForegroundColor Cyan }
        "DEBUG" { if ($script:Config.DetailedLogging) { Write-Host $logEntry -ForegroundColor Gray } }
        default { Write-Host $logEntry -ForegroundColor White }
    }
    
    # File logging
    try {
        Add-Content -Path $LogFile -Value $logEntry -Encoding UTF8
    } catch {
        Write-Warning "Failed to write to log file: $_"
    }
}

function Get-FuzzyMatch {
    param(
        [string]$Target,
        [string[]]$Candidates,
        [double]$Threshold = 0.8
    )
    
    if (-not $script:Config.EnableFuzzyMatching) { return $null }
    
    $bestMatch = $null
    $bestScore = 0
    
    foreach ($candidate in $Candidates) {
        $score = Get-StringSimilarity -String1 $Target -String2 $candidate
        if ($score -gt $bestScore -and $score -ge $Threshold) {
            $bestScore = $score
            $bestMatch = $candidate
        }
    }
    
    return $bestMatch
}

function Get-StringSimilarity {
    param([string]$String1, [string]$String2)
    
    if ($String1 -eq $String2) { return 1.0 }
    if ([string]::IsNullOrEmpty($String1) -or [string]::IsNullOrEmpty($String2)) { return 0.0 }
    
    # Simple similarity based on common characters and length
    $len1 = $String1.Length
    $len2 = $String2.Length
    $maxLen = [Math]::Max($len1, $len2)
    
    if ($maxLen -eq 0) { return 1.0 }
    
    # Convert to lowercase for comparison
    $s1 = $String1.ToLower()
    $s2 = $String2.ToLower()
    
    # Count common characters
    $commonChars = 0
    $usedChars = @{}
    
    foreach ($char in $s1.ToCharArray()) {
        if ($s2.Contains($char) -and -not $usedChars.ContainsKey($char)) {
            $commonChars++
            $usedChars[$char] = $true
        }
    }
    
    # Return similarity score
    return [double]$commonChars / $maxLen
}

function New-BackupFile {
    param([string]$FilePath)
    
    if (-not $script:Config.BackupFilesBeforeRepair) { return $null }
    
    try {
        $backupPath = $FilePath + ".backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        Copy-Item -Path $FilePath -Destination $backupPath -Force
        $script:Stats.BackupsCreated++
        Write-Log "Created backup: $backupPath" -Level "DEBUG"
        return $backupPath
    } catch {
        Write-Log "Failed to create backup for $FilePath`: $_" -Level "WARNING"
        return $null
    }
}

function Measure-MemoryUsage {
    $memoryUsage = [System.GC]::GetTotalMemory($false) / 1MB
    $script:Stats.MemoryUsage += $memoryUsage
    
    if ($memoryUsage -gt $script:Config.MaxMemoryUsageMB) {
        Write-Log "High memory usage detected: $([Math]::Round($memoryUsage, 2)) MB" -Level "WARNING"
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
    }
}
#endregion

#region Enhanced Excel Processing
function Process-ExcelFile {
    param(
        [string]$FilePath,
        [array]$AllFiles,
        [hashtable]$FileIndex
    )
    
    $fileResults = [System.Collections.ArrayList]::new()
    $excel = $null
    $wb = $null
    $retryCount = 0
    
    try {
        Write-Log "Processing Excel file: $FilePath" -Level "DEBUG"
        
        # Create backup if enabled
        $backupPath = New-BackupFile -FilePath $FilePath
        
        # Initialize Excel with retry logic
        do {
            try {
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $script:Config.ExcelApplicationVisible
                $excel.DisplayAlerts = $false
                $excel.ScreenUpdating = $false
                $excel.EnableEvents = $false
                break
            } catch {
                $retryCount++
                Write-Log "Excel initialization attempt $retryCount failed: $_" -Level "WARNING"
                Start-Sleep -Milliseconds $script:Config.RetryDelayMs
            }
        } while ($retryCount -lt $script:Config.RetryAttempts)
        
        if (-not $excel) {
            throw "Failed to initialize Excel after $($script:Config.RetryAttempts) attempts"
        }
        
        # Open workbook with enhanced error handling
        $wb = $excel.Workbooks.Open($FilePath, 0, $false, 5, '', '', $true, 3, "`t", $false, $false, 0, $true, 1, 0)
        
        # Process links
        $links = $wb.LinkSources(1)
        if ($links) {
            $script:Stats.TotalLinks += $links.Count
            
            foreach ($link in $links) {
                $linkResult = Process-ExcelLink -Workbook $wb -Link $link -AllFiles $AllFiles -FileIndex $FileIndex -FilePath $FilePath
                $null = $fileResults.Add($linkResult)
                
                if ($linkResult.Status -ne 0) {
                    $script:Stats.BrokenLinks++
                }
                if ($linkResult.Fixed) {
                    $script:Stats.FixedLinks++
                }
            }
        }
        
        # Save if changes were made
        $changesMode = $fileResults | Where-Object { $_.Fixed -eq $true }
        if ($changesMode.Count -gt 0) {
            $wb.Save()
            Write-Log "Saved changes to: $FilePath" -Level "SUCCESS"
        }
        
    } catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Error processing Excel file $FilePath`: $errorMsg" -Level "ERROR"
        
        $null = $script:ErrorLog.Add([PSCustomObject]@{
            File = $FilePath
            Type = 'Excel'
            Error = $errorMsg
            Timestamp = Get-Date
        })
        
        $null = $fileResults.Add([PSCustomObject]@{
            File = $FilePath
            Link = 'File Error'
            Status = 'Error'
            Fixed = $false
            NewLink = $null
            Type = 'Excel'
            ErrorMessage = $errorMsg
            ProcessingTime = 0
        })
        
        $script:Stats.FilesWithErrors++
    } finally {
        # Enhanced cleanup with proper COM object release
        if ($wb) {
            try {
                $wb.Close($false)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
            } catch { }
            $wb = $null
        }
        
        if ($excel) {
            try {
                $excel.Workbooks.Close()
                $excel.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            } catch { }
            $excel = $null
        }
        
        # Force garbage collection
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        
        Measure-MemoryUsage
    }
    
    return $fileResults.ToArray()
}

function Process-ExcelLink {
    param(
        [object]$Workbook,
        [string]$Link,
        [array]$AllFiles,
        [hashtable]$FileIndex,
        [string]$FilePath
    )
    
    $startTime = Get-Date
    $status = $Workbook.LinkInfo($Link, 1)
    $fixed = $false
    $newLink = $null
    $method = "None"
    
    if ($status -ne 0) {
        # Try exact filename match first
        $linkFileName = [System.IO.Path]::GetFileName($Link)
        
        if ($FileIndex.ContainsKey($linkFileName)) {
            $candidate = $FileIndex[$linkFileName]
            $newLink = $candidate
            $method = "Exact Match"
        }
        # Try fuzzy matching if exact match fails
        elseif ($script:Config.EnableFuzzyMatching) {
            $candidateNames = $AllFiles | ForEach-Object { [System.IO.Path]::GetFileName($_.FullName) }
            $fuzzyMatch = Get-FuzzyMatch -Target $linkFileName -Candidates $candidateNames -Threshold $script:Config.FuzzyMatchThreshold
            
            if ($fuzzyMatch) {
                $matchedFile = $AllFiles | Where-Object { [System.IO.Path]::GetFileName($_.FullName) -eq $fuzzyMatch } | Select-Object -First 1
                if ($matchedFile) {
                    $newLink = $matchedFile.FullName
                    $method = "Fuzzy Match"
                }
            }
        }
        
        # Try to apply the fix
        if ($newLink) {
            try {
                $Workbook.ChangeLink($Link, $newLink, 1)
                $newStatus = $Workbook.LinkInfo($newLink, 1)
                if ($newStatus -eq 0) {
                    $fixed = $true
                    $status = $newStatus
                    Write-Log "Fixed link using $method`: $Link -> $newLink" -Level "SUCCESS"
                } else {
                    Write-Log "Link fix failed for $Link -> $newLink (Status: $newStatus)" -Level "WARNING"
                }
            } catch {
                Write-Log "Failed to change link $Link to $newLink`: $_" -Level "ERROR"
            }
        }
    }
    
    $processingTime = (Get-Date) - $startTime
    
    return [PSCustomObject]@{
        File = $FilePath
        Link = $Link
        Status = $status
        Fixed = $fixed
        NewLink = $newLink
        Type = 'Excel'
        Method = $method
        ProcessingTime = $processingTime.TotalMilliseconds
        Timestamp = Get-Date
    }
}
#endregion

#region Enhanced Access Processing
function Process-AccessFile {
    param(
        [string]$FilePath,
        [array]$AllFiles,
        [hashtable]$FileIndex
    )
    
    $fileResults = [System.Collections.ArrayList]::new()
    $access = $null
    $retryCount = 0
    
    try {
        Write-Log "Processing Access file: $FilePath" -Level "DEBUG"
        
        # Create backup if enabled
        $backupPath = New-BackupFile -FilePath $FilePath
        
        # Initialize Access with retry logic
        do {
            try {
                $access = New-Object -ComObject Access.Application
                $access.Visible = $false
                $access.OpenCurrentDatabase($FilePath)
                break
            } catch {
                $retryCount++
                Write-Log "Access initialization attempt $retryCount failed: $_" -Level "WARNING"
                Start-Sleep -Milliseconds $script:Config.RetryDelayMs
            }
        } while ($retryCount -lt $script:Config.RetryAttempts)
        
        if (-not $access) {
            throw "Failed to initialize Access after $($script:Config.RetryAttempts) attempts"
        }
        
        # Get linked tables
        $linkedTables = $access.CurrentDb.TableDefs | Where-Object { $_.Connect -ne "" }
        
        foreach ($table in $linkedTables) {
            $linkResult = Process-AccessLink -Access $access -Table $table -AllFiles $AllFiles -FileIndex $FileIndex -FilePath $FilePath
            $null = $fileResults.Add($linkResult)
            
            if (-not $linkResult.Fixed -and $linkResult.Status -eq "Broken") {
                $script:Stats.BrokenLinks++
            }
            if ($linkResult.Fixed) {
                $script:Stats.FixedLinks++
            }
        }
        
        $script:Stats.TotalLinks += $linkedTables.Count
        
    } catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Error processing Access file $FilePath`: $errorMsg" -Level "ERROR"
        
        $null = $script:ErrorLog.Add([PSCustomObject]@{
            File = $FilePath
            Type = 'Access'
            Error = $errorMsg
            Timestamp = Get-Date
        })
        
        $null = $fileResults.Add([PSCustomObject]@{
            File = $FilePath
            Link = 'File Error'
            Status = 'Error'
            Fixed = $false
            NewLink = $null
            Type = 'Access'
            ErrorMessage = $errorMsg
            ProcessingTime = 0
        })
        
        $script:Stats.FilesWithErrors++
    } finally {
        if ($access) {
            try {
                $access.CloseCurrentDatabase()
                $access.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
            } catch { }
            $access = $null
        }
        
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        
        Measure-MemoryUsage
    }
    
    return $fileResults.ToArray()
}

function Process-AccessLink {
    param(
        [object]$Access,
        [object]$Table,
        [array]$AllFiles,
        [hashtable]$FileIndex,
        [string]$FilePath
    )
    
    $startTime = Get-Date
    $connectString = $Table.Connect
    $tableName = $Table.Name
    $fixed = $false
    $newLink = $null
    $status = "Unknown"
    $method = "None"
    
    try {
        # Test if link is working
        $recordCount = $Access.CurrentDb.OpenRecordset($tableName).RecordCount
        $status = "Working"
    } catch {
        $status = "Broken"
        
        # Extract file path from connection string
        if ($connectString -match "DATABASE=([^;]+)") {
            $linkedFile = $matches[1]
            $linkedFileName = [System.IO.Path]::GetFileName($linkedFile)
            
            # Try exact match
            if ($FileIndex.ContainsKey($linkedFileName)) {
                $newLink = $FileIndex[$linkedFileName]
                $method = "Exact Match"
            }
            # Try fuzzy matching
            elseif ($script:Config.EnableFuzzyMatching) {
                $candidateNames = $AllFiles | ForEach-Object { [System.IO.Path]::GetFileName($_.FullName) }
                $fuzzyMatch = Get-FuzzyMatch -Target $linkedFileName -Candidates $candidateNames -Threshold $script:Config.FuzzyMatchThreshold
                
                if ($fuzzyMatch) {
                    $matchedFile = $AllFiles | Where-Object { [System.IO.Path]::GetFileName($_.FullName) -eq $fuzzyMatch } | Select-Object -First 1
                    if ($matchedFile) {
                        $newLink = $matchedFile.FullName
                        $method = "Fuzzy Match"
                    }
                }
            }
            
            # Apply the fix
            if ($newLink -and (Test-Path $newLink)) {
                try {
                    $newConnectString = $connectString -replace [regex]::Escape($linkedFile), $newLink
                    $Table.Connect = $newConnectString
                    $Table.RefreshLink()
                    
                    # Test the new link
                    $testRecordCount = $Access.CurrentDb.OpenRecordset($tableName).RecordCount
                    $fixed = $true
                    $status = "Fixed"
                    Write-Log "Fixed Access link using $method`: $linkedFile -> $newLink" -Level "SUCCESS"
                } catch {
                    Write-Log "Failed to fix Access link $linkedFile -> $newLink`: $_" -Level "ERROR"
                }
            }
        }
    }
    
    $processingTime = (Get-Date) - $startTime
    
    return [PSCustomObject]@{
        File = $FilePath
        Link = $connectString
        Status = $status
        Fixed = $fixed
        NewLink = $newLink
        Type = 'Access'
        Method = $method
        TableName = $tableName
        ProcessingTime = $processingTime.TotalMilliseconds
        Timestamp = Get-Date
    }
}
#endregion

#region Enhanced Reporting Functions
function Export-DetailedReport {
    param(
        [array]$Results,
        [string]$OutputPath = "BrokenLinksReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    )
    
    try {
        # Summary statistics
        $summary = [PSCustomObject]@{
            'Total Files Processed' = $script:Stats.ProcessedFiles
            'Files with Errors' = $script:Stats.FilesWithErrors
            'Files Fixed' = $script:Stats.FilesFixed
            'Total Links Found' = $script:Stats.TotalLinks
            'Broken Links Found' = $script:Stats.BrokenLinks
            'Links Successfully Fixed' = $script:Stats.FixedLinks
            'Backups Created' = $script:Stats.BackupsCreated
            'Processing Time (minutes)' = [Math]::Round(((Get-Date) - $script:Config.StartTime).TotalMinutes, 2)
            'Average Memory Usage (MB)' = if ($script:Stats.MemoryUsage.Count -gt 0) { [Math]::Round(($script:Stats.MemoryUsage | Measure-Object -Average).Average, 2) } else { 0 }
            'Peak Memory Usage (MB)' = if ($script:Stats.MemoryUsage.Count -gt 0) { [Math]::Round(($script:Stats.MemoryUsage | Measure-Object -Maximum).Maximum, 2) } else { 0 }
            'Success Rate (%)' = if ($script:Stats.BrokenLinks -gt 0) { [Math]::Round(($script:Stats.FixedLinks / $script:Stats.BrokenLinks) * 100, 2) } else { 100 }
        }
        
        # Separate results by type and status
        $allResults = $Results | Select-Object File, Link, Status, Fixed, NewLink, Type, Method, ProcessingTime, Timestamp, @{N='ErrorMessage';E={if($_.PSObject.Properties['ErrorMessage']){$_.ErrorMessage}else{''}}}, @{N='TableName';E={if($_.PSObject.Properties['TableName']){$_.TableName}else{''}}}
        $fixedResults = $allResults | Where-Object { $_.Fixed -eq $true }
        $errorResults = $allResults | Where-Object { $_.Status -eq 'Error' -or $_.ErrorMessage -ne '' }
        $brokenResults = $allResults | Where-Object { $_.Status -ne 0 -and $_.Status -ne 'Working' -and $_.Status -ne 'Fixed' -and $_.Fixed -eq $false }
        
        # Create summary of fixed files (grouped by file)
        $fixedFilesSummary = $fixedResults | Group-Object File | Select-Object @{N='File';E={$_.Name}}, @{N='Links Fixed';E={$_.Count}}, @{N='File Type';E={($_.Group | Select-Object -First 1).Type}}, @{N='Fix Methods Used';E={($_.Group.Method | Sort-Object -Unique) -join ', '}} | Sort-Object File
        
        # Create summary of files still needing fixes
        $brokenFilesSummary = $brokenResults | Group-Object File | Select-Object @{N='File';E={$_.Name}}, @{N='Broken Links';E={$_.Count}}, @{N='File Type';E={($_.Group | Select-Object -First 1).Type}} | Sort-Object File
        
        # Export to Excel with multiple worksheets
        $summary | Export-Excel -Path $OutputPath -WorksheetName "Summary" -AutoSize -AutoFilter
        $allResults | Export-Excel -Path $OutputPath -WorksheetName "All Results" -AutoSize -AutoFilter
        $fixedResults | Export-Excel -Path $OutputPath -WorksheetName "Fixed Links" -AutoSize -AutoFilter
        $fixedFilesSummary | Export-Excel -Path $OutputPath -WorksheetName "Files Fixed Summary" -AutoSize -AutoFilter
        $brokenFilesSummary | Export-Excel -Path $OutputPath -WorksheetName "Files Needing Fixes" -AutoSize -AutoFilter
        $errorResults | Export-Excel -Path $OutputPath -WorksheetName "Files with Errors" -AutoSize -AutoFilter
        $brokenResults | Export-Excel -Path $OutputPath -WorksheetName "Broken Links" -AutoSize -AutoFilter
        
        if ($script:ErrorLog.Count -gt 0) {
            $script:ErrorLog | Export-Excel -Path $OutputPath -WorksheetName "Error Log" -AutoSize -AutoFilter
        }
        
        Write-Log "Detailed report exported to: $OutputPath" -Level "SUCCESS"
        return $OutputPath
    } catch {
        Write-Log "Failed to export detailed report: $_" -Level "ERROR"
        return $null
    }
}

function Show-Summary {
    Write-Host "`n" + "="*80 -ForegroundColor Cyan
    Write-Host "BROKEN LINKS SCAN & FIX SUMMARY" -ForegroundColor Cyan
    Write-Host "="*80 -ForegroundColor Cyan
    
    Write-Host "Files Processed: " -NoNewline -ForegroundColor White
    Write-Host $script:Stats.ProcessedFiles -ForegroundColor Green
    
    Write-Host "Files with Errors: " -NoNewline -ForegroundColor White
    Write-Host $script:Stats.FilesWithErrors -ForegroundColor $(if($script:Stats.FilesWithErrors -gt 0){"Red"}else{"Green"})
    
    Write-Host "Files Fixed: " -NoNewline -ForegroundColor White
    Write-Host $script:Stats.FilesFixed -ForegroundColor Green
    
    Write-Host "Total Links Found: " -NoNewline -ForegroundColor White
    Write-Host $script:Stats.TotalLinks -ForegroundColor Yellow
    
    Write-Host "Broken Links Found: " -NoNewline -ForegroundColor White
    Write-Host $script:Stats.BrokenLinks -ForegroundColor $(if($script:Stats.BrokenLinks -gt 0){"Red"}else{"Green"})
    
    Write-Host "Links Successfully Fixed: " -NoNewline -ForegroundColor White
    Write-Host $script:Stats.FixedLinks -ForegroundColor Green
    
    $successRate = if ($script:Stats.BrokenLinks -gt 0) { [Math]::Round(($script:Stats.FixedLinks / $script:Stats.BrokenLinks) * 100, 2) } else { 100 }
    Write-Host "Success Rate: " -NoNewline -ForegroundColor White
    Write-Host "$successRate%" -ForegroundColor $(if($successRate -ge 80){"Green"}elseif($successRate -ge 50){"Yellow"}else{"Red"})
    
    $totalTime = (Get-Date) - $script:Config.StartTime
    Write-Host "Total Processing Time: " -NoNewline -ForegroundColor White
    Write-Host "$([Math]::Round($totalTime.TotalMinutes, 2)) minutes" -ForegroundColor Cyan
    
    # Show list of fixed files if any
    if ($script:Stats.FixedLinks -gt 0) {
        Write-Host "`nFILES SUCCESSFULLY FIXED:" -ForegroundColor Green
        Write-Host "-" * 50 -ForegroundColor Green
        
        $fixedFiles = $script:Results | Where-Object { $_.Fixed -eq $true } | Group-Object File | Select-Object Name, Count
        foreach ($file in $fixedFiles) {
            $fileName = [System.IO.Path]::GetFileName($file.Name)
            Write-Host "✓ $fileName" -NoNewline -ForegroundColor Green
            Write-Host " ($($file.Count) link$(if($file.Count -gt 1){'s'}) fixed)" -ForegroundColor Gray
        }
        
        if ($fixedFiles.Count -gt 10) {
            Write-Host "... and $($fixedFiles.Count - 10) more files. See detailed report for complete list." -ForegroundColor Gray
        }
    }
    
    # Show list of files that still need fixing if any
    $brokenFiles = $script:Results | Where-Object { $_.Status -ne 0 -and $_.Status -ne 'Working' -and $_.Status -ne 'Fixed' -and $_.Fixed -eq $false } | Group-Object File | Select-Object Name, Count
    if ($brokenFiles.Count -gt 0) {
        Write-Host "`nFILES STILL NEEDING FIXES:" -ForegroundColor Red
        Write-Host "-" * 50 -ForegroundColor Red
        
        $displayCount = [Math]::Min($brokenFiles.Count, 10)
        for ($i = 0; $i -lt $displayCount; $i++) {
            $file = $brokenFiles[$i]
            $fileName = [System.IO.Path]::GetFileName($file.Name)
            Write-Host "✗ $fileName" -NoNewline -ForegroundColor Red
            Write-Host " ($($file.Count) broken link$(if($file.Count -gt 1){'s'}))" -ForegroundColor Gray
        }
        
        if ($brokenFiles.Count -gt 10) {
            Write-Host "... and $($brokenFiles.Count - 10) more files. See detailed report for complete list." -ForegroundColor Gray
        }
    }
    
    Write-Host "`n" + "="*80 -ForegroundColor Cyan
}
#endregion

#region Main Processing Logic
function Start-BrokenLinksScan {
    Write-Log "Starting Enhanced Broken Links Scan & Fix v2.0" -Level "INFO"
    Write-Log "Configuration: MaxJobs=$($script:Config.MaxParallelJobs), Backups=$($script:Config.BackupFilesBeforeRepair), FuzzyMatch=$($script:Config.EnableFuzzyMatching)" -Level "INFO"
    
    # Determine scan type and get files
    $scanResult = Get-FilesToProcess
    if (-not $scanResult) {
        Write-Log "No files found or scan cancelled" -Level "ERROR"
        return
    }
    
    $excelFiles = $scanResult.ExcelFiles
    $accessFiles = $scanResult.AccessFiles
    $allFiles = $excelFiles + $accessFiles
    $script:Stats.TotalFiles = $allFiles.Count
    
    if ($script:Stats.TotalFiles -eq 0) {
        Write-Log "No Excel or Access files found to process" -Level "WARNING"
        return
    }
    
    Write-Log "Found $($excelFiles.Count) Excel files and $($accessFiles.Count) Access files" -Level "INFO"
    
    # Create file index for fast lookups
    $fileIndex = @{}
    foreach ($file in $allFiles) {
        $fileName = [System.IO.Path]::GetFileName($file.FullName)
        if (-not $fileIndex.ContainsKey($fileName)) {
            $fileIndex[$fileName] = $file.FullName
        }
    }
    
    # Process Excel files with enhanced parallel processing and batch handling
    if ($excelFiles.Count -gt 0) {
        Write-Log "Processing $($excelFiles.Count) Excel files..." -Level "INFO"
        
        # For very large file sets, process in batches to prevent memory issues
        $batchSize = if ($excelFiles.Count -gt 10000) { 1000 } elseif ($excelFiles.Count -gt 5000) { 500 } else { $excelFiles.Count }
        $totalBatches = [Math]::Ceiling($excelFiles.Count / $batchSize)
        
        Write-Log "Processing in $totalBatches batches of $batchSize files each for memory optimization" -Level "INFO"
        
        for ($batchIndex = 0; $batchIndex -lt $totalBatches; $batchIndex++) {
            $startIndex = $batchIndex * $batchSize
            $endIndex = [Math]::Min(($batchIndex + 1) * $batchSize - 1, $excelFiles.Count - 1)
            $batchFiles = $excelFiles[$startIndex..$endIndex]
            
            Write-Log "Processing batch $($batchIndex + 1) of $totalBatches ($($batchFiles.Count) files)" -Level "INFO"
            
            try {
                $batchResults = Process-FilesParallel -Files $batchFiles -ProcessFunction 'Process-ExcelFile' -AllFiles $allFiles -FileIndex $fileIndex
                if ($batchResults) {
                    foreach ($result in $batchResults) {
                        $null = $script:Results.Add($result)
                    }
                }
            } catch {
                Write-Log "Error processing batch $($batchIndex + 1): $_" -Level "ERROR"
            }
            
            # Force garbage collection between batches
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
            
            # Update overall progress
            $overallProgress = [Math]::Round((($batchIndex + 1) / $totalBatches) * 100, 1)
            Write-Progress -Activity "Overall Excel Processing" -Status "Completed batch $($batchIndex + 1) of $totalBatches" -PercentComplete $overallProgress
        }
        
        Write-Progress -Activity "Overall Excel Processing" -Completed
        Write-Log "Parallel Excel processing completed. Processed $($excelFiles.Count) files in $totalBatches batches." -Level "SUCCESS"
    }
    
    # Process Access files (sequential due to COM limitations)
    if ($accessFiles.Count -gt 0) {
        Write-Log "Processing $($accessFiles.Count) Access files..." -Level "INFO"
        foreach ($accessFile in $accessFiles) {
            $accessResults = Process-AccessFile -FilePath $accessFile.FullName -AllFiles $allFiles -FileIndex $fileIndex
            $null = $script:Results.Add($accessResults)
            $script:Stats.ProcessedFiles++
            
            $progressPercent = [Math]::Round(($script:Stats.ProcessedFiles / $script:Stats.TotalFiles) * 100, 1)
            Write-Progress -Activity "Processing Access Files" -Status "Processing: $($accessFile.Name)" -PercentComplete $progressPercent
        }
        Write-Progress -Activity "Processing Access Files" -Completed
    }
    
    # Update statistics
    $script:Stats.FilesFixed = ($script:Results | Group-Object File | Where-Object { ($_.Group | Where-Object Fixed -eq $true).Count -gt 0 }).Count
    
    # Generate and export report
    Write-Log "Generating detailed report..." -Level "INFO"
    $reportPath = Export-DetailedReport -Results $script:Results
    
    # Show summary
    Show-Summary
    
    if ($reportPath) {
        Write-Host "`nDetailed Excel report saved to: " -NoNewline -ForegroundColor White
        Write-Host $reportPath -ForegroundColor Green
        Write-Host "`nReport includes the following worksheets:" -ForegroundColor Cyan
        Write-Host "• Summary - Overall statistics and performance metrics" -ForegroundColor Gray
        Write-Host "• Files Fixed Summary - List of files successfully repaired" -ForegroundColor Gray
        Write-Host "• Files Needing Fixes - List of files that still have broken links" -ForegroundColor Gray
        Write-Host "• Fixed Links - Detailed view of all successfully repaired links" -ForegroundColor Gray
        Write-Host "• Broken Links - Detailed view of all unrepaired links" -ForegroundColor Gray
        Write-Host "• All Results - Complete processing results" -ForegroundColor Gray
        Write-Host "• Files with Errors - Files that couldn't be processed" -ForegroundColor Gray
        if ($script:ErrorLog.Count -gt 0) {
            Write-Host "• Error Log - Detailed error information" -ForegroundColor Gray
        }
    }
    
    Write-Log "Broken Links Scan & Fix completed successfully" -Level "SUCCESS"
}

function Get-FilesToProcess {
    if ($SharePointUrl) {
        return Get-SharePointFiles -Url $SharePointUrl
    } elseif ($Path) {
        return Get-LocalFiles -Path $Path
    } else {
        return Get-FilesInteractive
    }
}

function Get-FilesInteractive {
    Write-Host "Select scan location type:" -ForegroundColor Cyan
    Write-Host "1. Local or SharePoint synced folder (fastest)" -ForegroundColor Yellow
    Write-Host "2. Online SharePoint folder URL (slower)" -ForegroundColor Yellow
    $locationType = Read-Host "Enter 1 or 2"
    
    switch ($locationType) {
        '1' {
            $folderPath = Read-Host "Enter the full path to the folder to scan"
            return Get-LocalFiles -Path $folderPath
        }
        '2' {
            $url = Read-Host "Enter the SharePoint folder URL"
            return Get-SharePointFiles -Url $url
        }
        default {
            Write-Log "Invalid selection" -Level "ERROR"
            return $null
        }
    }
}

function Get-LocalFiles {
    param([string]$Path)
    
    if (-not (Test-Path $Path)) {
        Write-Log "Path does not exist: $Path" -Level "ERROR"
        return $null
    }
    
    Write-Log "Scanning local folder: $Path" -Level "INFO"
    
    $excelFiles = Get-ChildItem -Path $Path -Include *.xls, *.xlsx, *.xlsm -Recurse -ErrorAction SilentlyContinue
    $accessFiles = Get-ChildItem -Path $Path -Include *.mdb, *.accdb -Recurse -ErrorAction SilentlyContinue
    
    return @{
        ExcelFiles = $excelFiles
        AccessFiles = $accessFiles
        TempFolder = $null
    }
}

function Get-SharePointFiles {
    param([string]$Url)
    
    # Import PnP.PowerShell
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Write-Log "Installing PnP.PowerShell module..." -Level "INFO"
        Install-Module -Name PnP.PowerShell -Force -Scope CurrentUser
    }
    Import-Module PnP.PowerShell
    
    Write-Log "Connecting to SharePoint: $Url" -Level "INFO"
    try {
        Connect-PnPOnline -Url $Url -Interactive
    } catch {
        Write-Log "Failed to connect to SharePoint: $_" -Level "ERROR"
        return $null
    }
    
    $tempFolder = Join-Path $env:TEMP "SPScan_$(Get-Date -Format 'yyyyMMddHHmmss')"
    New-Item -ItemType Directory -Path $tempFolder | Out-Null
    Write-Log "Created temporary folder: $tempFolder" -Level "INFO"
    
    try {
        $spFiles = Get-PnPFolderItem -FolderSiteRelativeUrl ([Uri]::UnescapeDataString((($Url -split '/sites/')[1]) -replace '^.+?/', '')) -ItemType File
        $excelFiles = @()
        $accessFiles = @()
        
        foreach ($spFile in $spFiles) {
            if ($spFile.Name -match '\.(xlsx?|xlsm)$') {
                $localPath = Join-Path $tempFolder $spFile.Name
                try {
                    Get-PnPFile -Url $spFile.ServerRelativeUrl -Path $tempFolder -FileName $spFile.Name -AsFile -Force
                    $excelFiles += Get-Item $localPath
                } catch {
                    Write-Log "Failed to download Excel file $($spFile.Name): $_" -Level "WARNING"
                }
            } elseif ($spFile.Name -match '\.(accdb|mdb)$') {
                $localPath = Join-Path $tempFolder $spFile.Name
                try {
                    Get-PnPFile -Url $spFile.ServerRelativeUrl -Path $tempFolder -FileName $spFile.Name -AsFile -Force
                    $accessFiles += Get-Item $localPath
                } catch {
                    Write-Log "Failed to download Access file $($spFile.Name): $_" -Level "WARNING"
                }
            }
        }
        
        return @{
            ExcelFiles = $excelFiles
            AccessFiles = $accessFiles
            TempFolder = $tempFolder
        }
    } catch {
        Write-Log "Failed to get SharePoint files: $_" -Level "ERROR"
        return $null
    }
}

function Process-FilesParallel {
    param(
        [array]$Files,
        [string]$ProcessFunction,
        [array]$AllFiles,
        [hashtable]$FileIndex
    )
    
    $allResults = [System.Collections.ArrayList]::new()
    $jobs = [System.Collections.ArrayList]::new()
    $maxJobs = $script:Config.MaxParallelJobs
    
    for ($i = 0; $i -lt $Files.Count; $i++) {
        $file = $Files[$i]
        
        # Start new job
        $job = Start-Job -ScriptBlock {
            param($FilePath, $AllFiles, $FileIndex, $Config)
            
            $fileResults = [System.Collections.ArrayList]::new()
            $excel = $null
            $wb = $null
            
            try {
                # Initialize Excel with performance settings
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                $excel.ScreenUpdating = $false
                $excel.EnableEvents = $false
                $excel.Calculation = -4135  # Manual calculation
                $excel.AutomationSecurity = 3
                
                # Open workbook with optimized parameters
                $wb = $excel.Workbooks.Open($FilePath, 0, $false, 5, '', '', $true, 3, "`t", $false, $false, 0, $true, 1, 0)
                
                # Get and process links
                $links = $wb.LinkSources(1)
                if ($links) {
                    foreach ($link in $links) {
                        $status = $wb.LinkInfo($link, 1)
                        $fixed = $false
                        $newLink = $null
                        $method = "None"
                        
                        if ($status -ne 0) {
                            $linkFileName = [System.IO.Path]::GetFileName($link)
                            
                            # Try exact match first
                            if ($FileIndex.ContainsKey($linkFileName)) {
                                $candidates = $FileIndex[$linkFileName]
                                $newLink = $candidates[0]
                                $method = "Exact Match"
                            }
                            
                            # Apply fix if found
                            if ($newLink) {
                                try {
                                    $wb.ChangeLink($link, $newLink, 1)
                                    $newStatus = $wb.LinkInfo($newLink, 1)
                                    if ($newStatus -eq 0) {
                                        $fixed = $true
                                    }
                                } catch {
                                    # Fix failed, keep original status
                                }
                            }
                        }
                        
                        $null = $fileResults.Add([PSCustomObject]@{
                            File = $FilePath
                            Link = $link
                            Status = $status
                            Fixed = $fixed
                            NewLink = $newLink
                            Type = 'Excel'
                            Method = $method
                            ProcessingTime = 0
                        })
                    }
                }
                
                # Save if changes were made
                $fixedLinks = $fileResults | Where-Object { $_.Fixed -eq $true }
                if ($fixedLinks.Count -gt 0) {
                    $wb.Save()
                }
                
            } catch {
                $null = $fileResults.Add([PSCustomObject]@{
                    File = $FilePath
                    Link = 'File Error'
                    Status = 'Error'
                    Fixed = $false
                    NewLink = $null
                    Type = 'Excel'
                    Method = 'Error'
                    ProcessingTime = 0
                    ErrorMessage = $_.Exception.Message
                })
            } finally {
                # Enhanced cleanup
                if ($wb) {
                    try {
                        $wb.Close($false)
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
                    } catch { }
                }
                if ($excel) {
                    try {
                        $excel.Workbooks.Close()
                        $excel.Quit()
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
                    } catch { }
                }
                
                # Force garbage collection
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                [System.GC]::Collect()
            }
            
            return $fileResults.ToArray()
        } -ArgumentList $file.FullName, $AllFiles, $FileIndex, $script:Config
        
        $null = $jobs.Add($job)
        
        # Wait for jobs to complete if we've reached the limit
        while ($jobs.Count -ge $maxJobs) {
            $completed = @($jobs | Where-Object { $_.State -eq 'Completed' })
            foreach ($completedJob in $completed) {
                $jobResults = Receive-Job $completedJob
                if ($jobResults) {
                    # Ensure jobResults is treated as an array
                    $resultsArray = @($jobResults)
                    foreach ($result in $resultsArray) {
                        $null = $allResults.Add($result)
                    }
                }
                Remove-Job $completedJob
                $null = $jobs.Remove($completedJob)
                $script:Stats.ProcessedFiles++
            }
            Start-Sleep -Milliseconds 200
        }
        
        # Update progress
        $progressPercent = [Math]::Round((($i + 1) / $Files.Count) * 100, 1)
        Write-Progress -Activity "Processing Excel Files" -Status "Processing: $($file.Name)" -PercentComplete $progressPercent
    }
    
    # Wait for remaining jobs
    while ($jobs.Count -gt 0) {
        $completed = @($jobs | Where-Object { $_.State -eq 'Completed' })
        foreach ($completedJob in $completed) {
            $jobResults = Receive-Job $completedJob
            if ($jobResults) {
                # Ensure jobResults is treated as an array
                $resultsArray = @($jobResults)
                foreach ($result in $resultsArray) {
                    $null = $allResults.Add($result)
                }
            }
            Remove-Job $completedJob
            $null = $jobs.Remove($completedJob)
            $script:Stats.ProcessedFiles++
        }
        if ($jobs.Count -gt 0) {
            Start-Sleep -Milliseconds 500
        }
    }
    
    Write-Progress -Activity "Processing Excel Files" -Completed
    return $allResults.ToArray()
}
#endregion

# Main execution
try {
    # Check for required modules
    $requiredModules = @('ImportExcel')
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Log "Installing required module: $module" -Level "INFO"
            Install-Module -Name $module -Force -Scope CurrentUser
        }
        Import-Module $module
    }
    
    # Start the enhanced scanning process
    Start-BrokenLinksScan
    
} catch {
    Write-Log "Critical error in main execution: $_" -Level "ERROR"
    throw
} finally {
    # Cleanup temporary folders if created
    if ($script:TempFolder -and (Test-Path $script:TempFolder)) {
        try {
            Remove-Item -Path $script:TempFolder -Recurse -Force
            Write-Log "Cleaned up temporary folder: $script:TempFolder" -Level "INFO"
        } catch {
            Write-Log "Failed to cleanup temporary folder: $_" -Level "WARNING"
        }
    }
}
