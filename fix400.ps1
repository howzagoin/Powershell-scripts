<#
.SYNOPSIS
    SharePoint and Local File Path Length Scanner

.DESCRIPTION
    This script identifies files with paths exceeding URL length limitations in both local folders 
    and SharePoint sites. It generates comprehensive Excel reports with detailed analysis and 
    recommendations for fixing long path issues. The script supports app-only authentication for 
    SharePoint scanning and includes batch recovery functionality to resume interrupted scans.

.PARAMETER TargetPathOrUrl
    The local folder path or SharePoint URL to scan. If not provided, the script will prompt for input.

.PARAMETER ClientId
    Azure AD application client ID for SharePoint authentication. Default value provided.

.PARAMETER TenantId
    Azure AD tenant ID for SharePoint authentication. Default value provided.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only authentication. Default value provided.

.PARAMETER OutputPath
    Path where the Excel report should be saved. If not provided, a save dialog will be shown.

.PARAMETER OfficeLimit
    Character limit for Office compatibility (default: 218).

.PARAMETER SharePointLimit
    Character limit for SharePoint (default: 400).

.PARAMETER ExcelFilesOnly
    Switch to scan only Excel files (.xlsx, .xls, .xlsm, .xlsb).

.PARAMETER PageSize
    Number of items to retrieve in each batch for SharePoint queries (default: 2000).

.EXAMPLE
    .\LongPathScanner.ps1
    Interactive mode - prompts for scan type and location.

.EXAMPLE
    .\LongPathScanner.ps1 -TargetPathOrUrl "C:\MyFolder"
    Scans the specified local folder.

.EXAMPLE
    .\LongPathScanner.ps1 -TargetPathOrUrl "https://tenant.sharepoint.com/sites/siteName"
    Scans the specified SharePoint site using default authentication.

.EXAMPLE
    .\LongPathScanner.ps1 -TargetPathOrUrl "C:\MyFolder" -ExcelFilesOnly -OutputPath "C:\Reports\LongPaths.xlsx"
    Scans only Excel files in the specified folder and saves to a custom location.

.NOTES
    File Name: LongPathScanner.ps1
    Author: Your Name
    Version: 2.0
    Requires: PowerShell 5.1 or later
    Modules: PnP.PowerShell, ImportExcel
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$TargetPathOrUrl,
    
    [Parameter(Mandatory=$false)]
    [string]$ClientId = '278b9af9-888d-4344-93bb-769bdd739249',
    
    [Parameter(Mandatory=$false)]
    [string]$TenantId = 'ca0711e2-e703-4f4e-9099-17d97863211c',
    
    [Parameter(Mandatory=$false)]
    [string]$CertificateThumbprint = '2E2502BB1EDB8F36CF9DE50936B283BDD22D5BAD',
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$false)]
    [int]$OfficeLimit = 218,
    
    [Parameter(Mandatory=$false)]
    [int]$SharePointLimit = 400,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExcelFilesOnly = $false,
    
    [Parameter(Mandatory=$false)]
    [int]$PageSize = 2000
)

# Initialize logging
$logFile = Join-Path $PSScriptRoot "LongPathScanner.log"
$scriptStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Define all functions first
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Parameter(Mandatory=$false)]
        [ValidateSet("Info", "Warning", "Error", "Success", "Debug")]
        [string]$Level = "Info",
        [Parameter(Mandatory=$false)]
        [string]$ForegroundColor
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    $logMessage | Out-File -FilePath $logFile -Append -Encoding utf8
    
    # Write to console
    switch ($Level) {
        "Info"    { $color = if ($ForegroundColor) { $ForegroundColor } else { "Cyan" } }
        "Warning" { $color = if ($ForegroundColor) { $ForegroundColor } else { "Yellow" } }
        "Error"   { $color = if ($ForegroundColor) { $ForegroundColor } else { "Red" } }
        "Success" { $color = if ($ForegroundColor) { $ForegroundColor } else { "Green" } }
        "Debug"   { $color = if ($ForegroundColor) { $ForegroundColor } else { "Gray" } }
        default   { $color = "White" }
    }
    
    Write-Host $logMessage -ForegroundColor $color
}

function Show-Progress {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete,
        [string]$CurrentOperation,
        [int]$Id = 1
    )
    
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete -CurrentOperation $CurrentOperation -Id $Id
}

function Stop-Progress {
    param(
        [string]$Activity,
        [int]$Id = 1
    )
    
    Write-Progress -Activity $Activity -Completed -Id $Id
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory=$true)]
        [scriptblock]$ScriptBlock,
        
        [int]$MaxRetries = 5,
        [int]$DelaySeconds = 2,
        
        [string]$Activity = "Retrying Operation"
    )
    
    $attempt = 0
    $lastError = $null
    
    while ($attempt -lt $MaxRetries) {
        try {
            return & $ScriptBlock
        } 
        catch {
            $lastError = $_
            $attempt++
            
            # Handle different types of errors
            if ($_.Exception.Message -like "*throttled*" -or $_.Exception.Message -like "*429*") {
                $wait = $DelaySeconds * $attempt * 2  # Exponential backoff for throttling
                Write-Log "Throttling detected. Retrying in $wait seconds... (Attempt $attempt/$MaxRetries)" -Level Warning
                Start-Sleep -Seconds $wait
            }
            elseif ($_.Exception.Message -like "*timeout*" -or $_.Exception.Message -like "*503*") {
                $wait = $DelaySeconds * $attempt
                Write-Log "Service unavailable. Retrying in $wait seconds... (Attempt $attempt/$MaxRetries)" -Level Warning
                Start-Sleep -Seconds $wait
            }
            elseif ($attempt -ge $MaxRetries) {
                Write-Log "Max retries exceeded for operation: $Activity" -Level Error
                throw $lastError
            }
            else {
                Write-Log "Operation failed (Attempt $attempt/$MaxRetries): $($_.Exception.Message)" -Level Warning
                Start-Sleep -Seconds $DelaySeconds
            }
        }
    }
    
    throw $lastError
}

function Install-RequiredModules {
    Write-Log "Checking and installing required modules..." -Level Info
    
    $modules = @(
        @{ Name = "PnP.PowerShell"; MinVersion = "2.3.0" },
        @{ Name = "ImportExcel"; MinVersion = "7.8.0" }
    )
    
    foreach ($module in $modules) {
        try {
            $installedModule = Get-Module -Name $module.Name -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
            
            if (-not $installedModule -or $installedModule.Version -lt $module.MinVersion) {
                Write-Log "Installing/updating module: $($module.Name)" -Level Info
                Install-Module -Name $module.Name -MinimumVersion $module.MinVersion -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
                Write-Log "Successfully installed $($module.Name)" -Level Success
            } else {
                Write-Log "$($module.Name) is already installed (Version: $($installedModule.Version))" -Level Info
            }
        }
        catch {
            Write-Log "Failed to install module $($module.Name): $_" -Level Error
            throw
        }
    }
    
    # Import modules
    try {
        Import-Module PnP.PowerShell -Force -ErrorAction Stop
        Import-Module ImportExcel -Force -ErrorAction Stop
        Write-Log "All required modules loaded successfully" -Level Success
    }
    catch {
        Write-Log "Failed to import required modules: $_" -Level Error
        throw
    }
}

function Get-SaveFileDialog {
    param(
        [string]$InitialDirectory = [Environment]::GetFolderPath('Desktop'),
        [string]$Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
        [string]$DefaultFileName = "LongPathReport.xlsx",
        [string]$Title = "Save Long Path Report"
    )
    
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.InitialDirectory = $InitialDirectory
        $SaveFileDialog.Filter = $Filter
        $SaveFileDialog.FileName = $DefaultFileName
        $SaveFileDialog.Title = $Title
        $SaveFileDialog.DefaultExt = "xlsx"
        $SaveFileDialog.AddExtension = $true
        
        $result = $SaveFileDialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $SaveFileDialog.FileName
        } else {
            return $null
        }
    }
    catch {
        Write-Log "Could not show save dialog. Using default filename in current directory." -Level Warning
        return $DefaultFileName
    }
}

function Get-FolderBrowserDialog {
    param(
        [string]$Description = "Select folder to scan for long paths",
        [string]$InitialDirectory = [Environment]::GetFolderPath('Desktop')
    )
    
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $FolderBrowserDialog.Description = $Description
        $FolderBrowserDialog.ShowNewFolderButton = $false
        $FolderBrowserDialog.SelectedPath = $InitialDirectory
        
        $result = $FolderBrowserDialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $FolderBrowserDialog.SelectedPath
        } else {
            return $null
        }
    }
    catch {
        Write-Log "Could not show folder browser dialog. Please enter path manually." -Level Warning
        return $null
    }
}

function Test-SharePointSyncedFolder {
    param(
        [string]$Path
    )
    
    # Check for SharePoint sync markers
    $syncMarkers = @(".sharepoint", "sync_log.txt", "sync_root.log")
    
    foreach ($marker in $syncMarkers) {
        $markerPath = Join-Path $Path $marker
        if (Test-Path $markerPath -PathType Leaf) {
            return $true
        }
    }
    
    return $false
}

function Get-SharePointUrlFromLocalPath {
    param(
        [string]$LocalPath
    )
    
    try {
        # Find the sync root folder
        $currentPath = $LocalPath
        $syncRoot = $null
        
        while ($currentPath -and -not $syncRoot) {
            if (Test-SharePointSyncedFolder -Path $currentPath) {
                $syncRoot = $currentPath
                break
            }
            $currentPath = Split-Path $currentPath -Parent
        }
        
        if (-not $syncRoot) {
            return ""
        }
        
        # Get the relative path from sync root
        $relativePath = $LocalPath.Substring($syncRoot.Length).TrimStart('\')
        
        # Replace backslashes with forward slashes
        $relativePath = $relativePath -replace '\\', '/'
        
        # Get SharePoint URL from sync settings (simplified approach)
        $siteName = Split-Path (Split-Path $syncRoot -Parent) -Leaf
        $libraryName = Split-Path $syncRoot -Leaf
        
        # Remove common sync suffixes
        $libraryName = $libraryName -replace ' - .*$', ''
        
        # Construct SharePoint URL
        $sharePointUrl = "https://$($env:USERDNSDOMAIN).sharepoint.com/sites/$siteName/$libraryName/$relativePath"
        
        return $sharePointUrl
    }
    catch {
        return ""
    }
}

function Display-ResultsSummary {
    param(
        [array]$Results,
        [int]$OfficeLimit,
        [int]$SharePointLimit
    )
    
    $officeProblems = $Results | Where-Object { $_.ExceedsOfficeLimit }
    $sharePointOnlyProblems = $Results | Where-Object { $_.ExceedsSharePointLimit -and -not $_.ExceedsOfficeLimit }
    
    Write-Host "`nLong Path Files Analysis Results:" -ForegroundColor Cyan
    Write-Host "=================================" -ForegroundColor Cyan
    Write-Host "Total files with long paths: $($Results.Count)" -ForegroundColor Yellow
    Write-Host "Files exceeding Office limit ($OfficeLimit chars): $($officeProblems.Count)" -ForegroundColor Red
    Write-Host "Files exceeding SharePoint limit ($SharePointLimit chars): $($sharePointOnlyProblems.Count)" -ForegroundColor Yellow
    
    if ($Results.Count -gt 0) {
        $avgPathLength = [math]::Round(($Results | Measure-Object -Property PathLength -Average).Average, 0)
        $maxPathLength = ($Results | Measure-Object -Property PathLength -Maximum).Maximum
        Write-Host "Average path length: $avgPathLength characters" -ForegroundColor White
        Write-Host "Maximum path length: $maxPathLength characters" -ForegroundColor White
    }
    
    # Show top 10 longest paths
    Write-Host "`nTop 10 Longest Paths:" -ForegroundColor Cyan
    $topLongPaths = $Results | Sort-Object PathLength -Descending | Select-Object -First 10
    foreach ($file in $topLongPaths) {
        $priorityColor = if ($file.ExceedsOfficeLimit) { "Red" } else { "Yellow" }
        $displayPath = if ($file.FullPath.Length -gt 100) { $file.FullPath.Substring(0, 97) + "..." } else { $file.FullPath }
        Write-Host "$($file.PathLength) chars: $displayPath" -ForegroundColor $priorityColor
    }
    
    # Display priority recommendations
    Write-Host "`n*** PRIORITY RECOMMENDATIONS ***" -ForegroundColor Cyan
    Write-Host "1. HIGH PRIORITY: Fix $($officeProblems.Count) files exceeding Office limit ($OfficeLimit chars)" -ForegroundColor Red
    Write-Host "2. MEDIUM PRIORITY: Fix $($sharePointOnlyProblems.Count) files exceeding SharePoint limit ($SharePointLimit chars)" -ForegroundColor Yellow
    Write-Host "3. BEST PRACTICE: Keep all new files under $OfficeLimit characters for maximum compatibility" -ForegroundColor Green
    Write-Host "4. ACTIONS: Shorten folder names, move files to shorter paths, or abbreviate filenames" -ForegroundColor White
}

function Get-LocalLongPaths {
    param(
        [string]$RootPath
    )
    
    $longPathFiles = @()
    $processedFiles = 0
    
    Write-Log "Scanning local files..." -Level Info
    
    # Check if it's a SharePoint synced folder
    $isSharePointSynced = Test-SharePointSyncedFolder -Path $RootPath
    
    # Get all files
    $allFiles = Get-ChildItem -Path $RootPath -File -Recurse -ErrorAction SilentlyContinue
    
    if (-not $allFiles -or $allFiles.Count -eq 0) {
        Write-Log "No files found in the specified path." -Level Warning
        return @()
    }
    
    Write-Log "Found $($allFiles.Count) files to analyze" -Level Info
    
    foreach ($file in $allFiles) {
        $processedFiles++
        
        # Update progress
        if ($processedFiles % 100 -eq 0) {
            $percentComplete = [math]::Round(($processedFiles / $allFiles.Count) * 100, 1)
            Show-Progress -Activity "Scanning Local Files" -Status "Processed: $processedFiles/$($allFiles.Count)" -PercentComplete $percentComplete -CurrentOperation "Current: $($file.Name)"
        }
        
        # Filter for Excel files if requested
        if ($ExcelFilesOnly) {
            $extension = $file.Extension.ToLower()
            if ($extension -notin @('.xlsx', '.xls', '.xlsm', '.xlsb')) {
                continue
            }
        }
        
        $fullPath = $file.FullName
        $pathLength = $fullPath.Length
        $fileNameLength = $file.Name.Length
        $relativePath = $fullPath.Replace($RootPath, "").TrimStart('\')
        
        # Calculate SharePoint URL length if it's a synced folder
        $sharePointUrlLength = 0
        $sharePointUrl = ""
        
        if ($isSharePointSynced) {
            $sharePointUrl = Get-SharePointUrlFromLocalPath -LocalPath $fullPath
            if ($sharePointUrl) {
                $sharePointUrlLength = $sharePointUrl.Length
            }
        }
        
        # Check against limits
        $exceedsOfficeLimit = $false
        $exceedsSharePointLimit = $false
        
        if ($sharePointUrlLength -gt 0) {
            $exceedsOfficeLimit = $sharePointUrlLength -gt $OfficeLimit
            $exceedsSharePointLimit = $sharePointUrlLength -gt $SharePointLimit
        } else {
            $exceedsOfficeLimit = $pathLength -gt $OfficeLimit
            $exceedsSharePointLimit = $pathLength -gt $SharePointLimit
        }
        
        # Add to results if exceeds any limit
        if ($exceedsOfficeLimit -or $exceedsSharePointLimit) {
            # Calculate excess values properly
            $officeExcess = 0
            $sharePointExcess = 0
            
            if ($sharePointUrlLength -gt 0) {
                $officeExcess = [Math]::Max(0, $sharePointUrlLength - $OfficeLimit)
                $sharePointExcess = [Math]::Max(0, $sharePointUrlLength - $SharePointLimit)
            } else {
                $officeExcess = [Math]::Max(0, $pathLength - $OfficeLimit)
                $sharePointExcess = [Math]::Max(0, $pathLength - $SharePointLimit)
            }
            
            $longPathFiles += [PSCustomObject]@{
                FolderName = $file.Directory.Name
                FolderPath = $file.Directory.FullName
                FileName = $file.Name
                FileExtension = $file.Extension
                FullPath = $fullPath
                PathLength = $pathLength
                FileNameLength = $fileNameLength
                RelativePath = $relativePath
                FileSize = $file.Length
                FileSizeMB = [math]::Round($file.Length / 1MB, 2)
                LastModified = $file.LastWriteTime
                CreationTime = $file.CreationTime
                SharePointUrl = $sharePointUrl
                SharePointUrlLength = $sharePointUrlLength
                ExceedsOfficeLimit = $exceedsOfficeLimit
                ExceedsSharePointLimit = $exceedsSharePointLimit
                Source = "Local"
                OfficeExcess = $officeExcess
                SharePointExcess = $sharePointExcess
            }
        }
    }
    
    Stop-Progress -Activity "Scanning Local Files"
    Write-Log "Completed local scan. Found $($longPathFiles.Count) files with long paths" -Level Success
    return $longPathFiles
}

function Get-SharePointLongPaths {
    param(
        [string]$SharePointUrl,
        [string]$ClientId,
        [string]$TenantId,
        [string]$CertificateThumbprint
    )
    
    $longPathFiles = @()
    $success = $false
    
    # Create recovery files
    $siteKey = $SharePointUrl -replace '[^a-zA-Z0-9]', '_'
    $recoveryFile = Join-Path $PSScriptRoot "Recovery_$siteKey.json"
    $batchLogFile = Join-Path $PSScriptRoot "BatchLog_$siteKey.csv"
    
    try {
        # Load existing recovery data if available
        if (Test-Path $recoveryFile) {
            $recoveryData = Get-Content $recoveryFile | ConvertFrom-Json
            $longPathFiles = $recoveryData.LongPathFiles
            $processedLibraries = $recoveryData.ProcessedLibraries
            Write-Log "Loaded recovery data: $($longPathFiles.Count) files found, $($processedLibraries.Count) libraries processed" -Level Info
        } else {
            $processedLibraries = @()
        }
        
        # Initialize batch log
        if (-not (Test-Path $batchLogFile)) {
            "Timestamp,Library,Status,ItemCount,Message" | Out-File -FilePath $batchLogFile -Encoding utf8
        }
        
        # Derive tenant URL from SharePoint URL
        $tenantUrl = ($SharePointUrl -split '/')[0..2] -join '/'
        Write-Log "Derived tenant URL: $tenantUrl" -Level Info
        
        # Connect to SharePoint using app-only certificate authentication
        Write-Log "Connecting to SharePoint using app-only authentication..." -Level Info
        
        Invoke-WithRetry -ScriptBlock {
            Connect-PnPOnline -Url $SharePointUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $CertificateThumbprint
        } -Activity "Connecting to SharePoint"
        
        # Get all document libraries
        $libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }
        
        Write-Log "Found $($libraries.Count) document libraries" -Level Info
        
        foreach ($library in $libraries) {
            # Skip if already processed
            if ($library.Title -in $processedLibraries) {
                Write-Log "Skipping already processed library: $($library.Title)" -Level Info
                continue
            }
            
            Write-Log "Processing library: $($library.Title)" -Level Info
            
            # Get all files in the library with pagination
            $allItems = @()
            $position = $null
            $batchCount = 0
            
            do {
                $batchCount++
                Write-Log "Processing batch $batchCount for library: $($library.Title)" -Level Info
                
                try {
                    $items = Invoke-WithRetry -ScriptBlock {
                        Get-PnPListItem -List $library -Fields "FileRef", "File_x0020_Type" -PageSize $PageSize -Position $position
                    } -Activity "Getting items from $($library.Title)"
                    
                    $allItems += $items
                    $position = $items | Select-Object -Last 1 | ForEach-Object { $_["FileRef"] }
                    
                    # Log successful batch
                    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    "$timestamp,$($library.Title),Success,$($items.Count),Batch completed successfully" | Out-File -FilePath $batchLogFile -Append -Encoding utf8
                    
                    # Save recovery data after each batch
                    $recoveryData = @{
                        LongPathFiles = $longPathFiles
                        ProcessedLibraries = $processedLibraries
                        CurrentLibrary = $library.Title
                        LastBatch = $batchCount
                        LastPosition = $position
                    }
                    $recoveryData | ConvertTo-Json -Depth 10 | Set-Content $recoveryFile
                    
                    Write-Log "Batch $batchCount completed. Processed $($items.Count) items. Total items: $($allItems.Count)" -Level Success
                }
                catch {
                    # Log failed batch
                    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    "$timestamp,$($library.Title),Failed,0,$($_.Exception.Message)" | Out-File -FilePath $batchLogFile -Append -Encoding utf8
                    
                    Write-Log "Batch $batchCount failed for library $($library.Title): $_" -Level Error
                    Write-Log "Script will resume from this batch on next run" -Level Warning
                    
                    # Re-throw to stop processing this library
                    throw
                }
            } while ($items.Count -eq $PageSize)
            
            # Process files in the library
            $files = $allItems | Where-Object { $_.FileSystemObjectType -eq "File" }
            Write-Log "Found $($files.Count) files in $($library.Title)" -Level Info
            
            foreach ($file in $files) {
                $serverRelativeUrl = $file["FileRef"]
                $fileName = [System.IO.Path]::GetFileName($serverRelativeUrl)
                $folderPath = [System.IO.Path]::GetDirectoryName($serverRelativeUrl)
                
                # Construct full SharePoint URL
                $fullUrl = "$tenantUrl$serverRelativeUrl"
                $urlLength = $fullUrl.Length
                $fileNameLength = $fileName.Length
                $relativePath = $serverRelativeUrl.TrimStart('/')
                
                # Get file properties
                $fileItem = Invoke-WithRetry -ScriptBlock {
                    Get-PnPFile -Url $serverRelativeUrl -AsListItem
                } -Activity "Getting file properties for $fileName"
                
                $fileSize = $fileItem["File_x0020_Size"]
                $lastModified = $fileItem["Modified"]
                $creationTime = $fileItem["Created"]
                
                # Check against limits
                $exceedsOfficeLimit = $urlLength -gt $OfficeLimit
                $exceedsSharePointLimit = $urlLength -gt $SharePointLimit
                
                # Add to results if exceeds any limit
                if ($exceedsOfficeLimit -or $exceedsSharePointLimit) {
                    $longPathFiles += [PSCustomObject]@{
                        FolderName = [System.IO.Path]::GetFileName($folderPath)
                        FolderPath = $folderPath
                        FileName = $fileName
                        FileExtension = [System.IO.Path]::GetExtension($fileName)
                        FullPath = $fullUrl
                        PathLength = $urlLength
                        FileNameLength = $fileNameLength
                        RelativePath = $relativePath
                        FileSize = $fileSize
                        FileSizeMB = [math]::Round($fileSize / 1MB, 2)
                        LastModified = $lastModified
                        CreationTime = $creationTime
                        SharePointUrl = $fullUrl
                        SharePointUrlLength = $urlLength
                        ExceedsOfficeLimit = $exceedsOfficeLimit
                        ExceedsSharePointLimit = $exceedsSharePointLimit
                        Source = "SharePoint"
                        OfficeExcess = [Math]::Max(0, $urlLength - $OfficeLimit)
                        SharePointExcess = [Math]::Max(0, $urlLength - $SharePointLimit)
                    }
                }
            }
            
            # Mark library as processed
            $processedLibraries += $library.Title
            
            # Update recovery data
            $recoveryData = @{
                LongPathFiles = $longPathFiles
                ProcessedLibraries = $processedLibraries
                CurrentLibrary = $null
                LastBatch = 0
                LastPosition = $null
            }
            $recoveryData | ConvertTo-Json -Depth 10 | Set-Content $recoveryFile
            
            Write-Log "Completed processing library: $($library.Title)" -Level Success
        }
        
        Disconnect-PnPOnline
        $success = $true
    }
    catch {
        Write-Log "Error scanning SharePoint: $_" -Level Error
        Write-Log "Recovery data saved. Script will resume from last successful batch on next run." -Level Warning
        throw
    }
    finally {
        # Clean up recovery files if successful
        if ($success) {
            if (Test-Path $recoveryFile) {
                Remove-Item $recoveryFile -Force
                Write-Log "Removed recovery file: $recoveryFile" -Level Info
            }
            if (Test-Path $batchLogFile) {
                Remove-Item $batchLogFile -Force
                Write-Log "Removed batch log file: $batchLogFile" -Level Info
            }
        }
    }
    
    return $longPathFiles
}

function Export-ToExcel {
    param(
        [string]$FilePath,
        [array]$Data,
        [int]$OfficeLimit,
        [int]$SharePointLimit
    )
    
    try {
        Write-Log "Generating Excel report..." -Level Info
        
        # Remove existing file
        if (Test-Path $FilePath) {
            Remove-Item $FilePath -Force -ErrorAction SilentlyContinue
        }
        
        # Separate files by limit type
        $officeProblems = $Data | Where-Object { $_.ExceedsOfficeLimit }
        $sharePointProblems = $Data | Where-Object { $_.ExceedsSharePointLimit -and -not $_.ExceedsOfficeLimit }
        
        # Calculate summary statistics
        $totalLongPathFiles = $Data.Count
        $totalOfficeIssues = $officeProblems.Count
        $totalSharePointIssues = $sharePointProblems.Count
        $avgPathLength = if ($Data.Count -gt 0) { [math]::Round(($Data | Measure-Object -Property PathLength -Average).Average, 0) } else { 0 }
        $maxPathLength = if ($Data.Count -gt 0) { ($Data | Measure-Object -Property PathLength -Maximum).Maximum } else { 0 }
        
        # Create summary data
        $summaryData = @(
            [PSCustomObject]@{ 'Metric' = 'Analysis Date'; 'Value' = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
            [PSCustomObject]@{ 'Metric' = 'Scanned Path'; 'Value' = if ($Data[0].Source -eq "Local") { $Data[0].FullPath.Split('\')[0..2] -join '\' } else { $Data[0].SharePointUrl.Split('/')[0..2] -join '/' } }
            [PSCustomObject]@{ 'Metric' = 'Office Compatibility Limit'; 'Value' = "$OfficeLimit characters" }
            [PSCustomObject]@{ 'Metric' = 'SharePoint Online Limit'; 'Value' = "$SharePointLimit characters" }
            [PSCustomObject]@{ 'Metric' = ''; 'Value' = '' }  # Spacer
            [PSCustomObject]@{ 'Metric' = 'FINDINGS SUMMARY'; 'Value' = '' }
            [PSCustomObject]@{ 'Metric' = 'Total Files with Long Paths'; 'Value' = $totalLongPathFiles }
            [PSCustomObject]@{ 'Metric' = 'Files Exceeding Office Limit'; 'Value' = $totalOfficeIssues }
            [PSCustomObject]@{ 'Metric' = 'Files Exceeding SharePoint Limit'; 'Value' = $totalSharePointIssues }
            [PSCustomObject]@{ 'Metric' = 'Average Path Length'; 'Value' = $avgPathLength }
            [PSCustomObject]@{ 'Metric' = 'Maximum Path Length Found'; 'Value' = $maxPathLength }
            [PSCustomObject]@{ 'Metric' = ''; 'Value' = '' }  # Spacer
            [PSCustomObject]@{ 'Metric' = 'RECOMMENDATIONS'; 'Value' = '' }
            [PSCustomObject]@{ 'Metric' = 'Priority 1: Office Compatibility'; 'Value' = "Fix $totalOfficeIssues files exceeding $OfficeLimit chars" }
            [PSCustomObject]@{ 'Metric' = 'Priority 2: SharePoint Limits'; 'Value' = "Fix $totalSharePointIssues files exceeding $SharePointLimit chars" }
            [PSCustomObject]@{ 'Metric' = 'Best Practice'; 'Value' = "Keep file paths under $OfficeLimit characters for maximum compatibility" }
        )
        
        # Export Summary worksheet first
        Write-Log "Creating Summary worksheet..." -Level Info
        $summaryData | Export-Excel -Path $FilePath -WorksheetName "Summary" -AutoSize -TableStyle "Medium6" -Title "Long Path Analysis Summary & Recommendations"
        
        # Worksheet 1: Problem Files (Office Compatibility)
        if ($officeProblems -and $officeProblems.Count -gt 0) {
            Write-Log "Creating Office Problems worksheet with $($officeProblems.Count) entries..." -Level Info
            $officeData = $officeProblems | Select-Object @{N='Folder Name';E={$_.FolderName}}, 
                @{N='File Name';E={$_.FileName}}, 
                @{N='Extension';E={$_.FileExtension}}, 
                @{N='Full Path';E={$_.FullPath}}, 
                @{N='Path Length';E={$_.PathLength}}, 
                @{N='Excess Characters';E={$_.OfficeExcess}}, 
                @{N='File Size (MB)';E={$_.FileSizeMB}}, 
                @{N='Last Modified';E={$_.LastModified}},
                @{N='Priority';E={"HIGH - Office Incompatible"}},
                @{N='Relative Path';E={$_.RelativePath}}
                
            $officeData | Export-Excel -Path $FilePath -WorksheetName "Office Problems ($OfficeLimit+)" -AutoSize -TableStyle "Dark1" -Title "Files Exceeding Office Compatibility Limit ($OfficeLimit characters) - HIGH PRIORITY"
        }
        
        # Worksheet 2: Problem Files (SharePoint Limit)
        if ($sharePointProblems -and $sharePointProblems.Count -gt 0) {
            Write-Log "Creating SharePoint Problems worksheet with $($sharePointProblems.Count) entries..." -Level Info
            $sharePointData = $sharePointProblems | Select-Object @{N='Folder Name';E={$_.FolderName}}, 
                @{N='File Name';E={$_.FileName}}, 
                @{N='Extension';E={$_.FileExtension}}, 
                @{N='Full Path';E={$_.FullPath}}, 
                @{N='Path Length';E={$_.PathLength}}, 
                @{N='Excess Characters';E={$_.SharePointExcess}}, 
                @{N='File Size (MB)';E={$_.FileSizeMB}}, 
                @{N='Last Modified';E={$_.LastModified}},
                @{N='Priority';E={"MEDIUM - SharePoint Limit"}},
                @{N='Relative Path';E={$_.RelativePath}}
                
            $sharePointData | Export-Excel -Path $FilePath -WorksheetName "SharePoint Problems ($SharePointLimit+)" -AutoSize -TableStyle "Medium2" -Title "Files Exceeding SharePoint Limit ($SharePointLimit characters) - MEDIUM PRIORITY"
        }
        
        # Worksheet 3: All Problem Files (Combined view)
        if ($Data -and $Data.Count -gt 0) {
            Write-Log "Creating All Problems worksheet with $($Data.Count) entries..." -Level Info
            $allData = $Data | Select-Object @{N='Source';E={$_.Source}},
                @{N='Folder Name';E={$_.FolderName}}, 
                @{N='Folder Path';E={$_.FolderPath}}, 
                @{N='File Name';E={$_.FileName}}, 
                @{N='Extension';E={$_.FileExtension}}, 
                @{N='Full Path';E={$_.FullPath}}, 
                @{N='Path Length';E={$_.PathLength}}, 
                @{N='File Name Length';E={$_.FileNameLength}},
                @{N='Exceeds Office Limit';E={if($_.ExceedsOfficeLimit){"Yes"}else{"No"}}}, 
                @{N='Exceeds SharePoint Limit';E={if($_.ExceedsSharePointLimit){"Yes"}else{"No"}}}, 
                @{N='Office Excess';E={$_.OfficeExcess}}, 
                @{N='SharePoint Excess';E={$_.SharePointExcess}}, 
                @{N='File Size (MB)';E={$_.FileSizeMB}}, 
                @{N='Last Modified';E={$_.LastModified}}, 
                @{N='Creation Time';E={$_.CreationTime}}, 
                @{N='Relative Path';E={$_.RelativePath}}
                
            $allData | Export-Excel -Path $FilePath -WorksheetName "All Problems (Raw Data)" -AutoSize -TableStyle "Light1" -Title "All Files with Long Path Issues - Complete Dataset"
        }
        
        Write-Log "Excel report created successfully: $FilePath" -Level Success
    }
    catch {
        Write-Log "Failed to create Excel report: $_" -Level Error
        throw
    }
}

function Main {
    try {
        Write-Log "Script started at $scriptStartTime" -Level Info
        Write-Log "Long Path Scanner - SharePoint and Local File System Support" -Level Success
        Write-Log "============================================================" -Level Success
        
        # Get target path or URL
        if (-not $TargetPathOrUrl) {
            $choice = Read-Host "Scan type: (1) Local folder or (2) SharePoint URL? [1/2]"
            
            if ($choice -eq "2") {
                $TargetPathOrUrl = Read-Host "Enter SharePoint URL (e.g., https://tenant.sharepoint.com/sites/siteName)"
            }
            else {
                $TargetPathOrUrl = Get-FolderBrowserDialog
                if (-not $TargetPathOrUrl) {
                    Write-Host "No folder selected. Please enter the folder path manually (e.g., C:\MyFolder):" -ForegroundColor Yellow
                    $TargetPathOrUrl = Read-Host "Folder path"
                    if (-not $TargetPathOrUrl) {
                        Write-Log "No folder selected or entered. Exiting." -Level Info
                        return
                    }
                }
            }
        }
        
        Write-Log "Target for scan: $TargetPathOrUrl" -Level Info
        Write-Log "Office Compatibility Limit: $OfficeLimit characters" -Level Info
        Write-Log "SharePoint Online Limit: $SharePointLimit characters" -Level Info
        
        # Create output filename
        $dateStr = Get-Date -Format "yyyyMMdd_HHmm"
        $fileFilter = if ($ExcelFilesOnly) { "ExcelOnly" } else { "AllFiles" }
        $defaultFileName = "LongPathReport-$fileFilter-$dateStr.xlsx"
        
        # Get output path
        if ($OutputPath) {
            $excelFileName = $OutputPath
        } else {
            $excelFileName = Get-SaveFileDialog -DefaultFileName $defaultFileName -Title "Save Long Path Report"
            if (-not $excelFileName) {
                Write-Log "User cancelled the save dialog. Exiting." -Level Info
                return
            }
        }
        
        Write-Log "Report will be saved to: $excelFileName" -Level Info
        
        # Determine scan type and execute
        if ($TargetPathOrUrl -match '^https?://') {
            # SharePoint URL scanning
            Write-Log "Scanning SharePoint URL: $TargetPathOrUrl" -Level Info
            Install-RequiredModules
            $results = Get-SharePointLongPaths -SharePointUrl $TargetPathOrUrl -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
        }
        elseif (Test-Path $TargetPathOrUrl) {
            # Local folder scanning
            Write-Log "Scanning local folder: $TargetPathOrUrl" -Level Info
            $results = Get-LocalLongPaths -RootPath $TargetPathOrUrl
        }
        else {
            Write-Log "Invalid path or URL provided: $TargetPathOrUrl" -Level Error
            return
        }
        
        if ($results.Count -eq 0) {
            Write-Log "No files with long paths found." -Level Success
            return
        }
        
        # Sort results by folder path and filename
        $sortedResults = $results | Sort-Object -Property FolderPath, FileName
        
        # Export results
        Export-ToExcel -FilePath $excelFileName -Data $sortedResults -OfficeLimit $OfficeLimit -SharePointLimit $SharePointLimit
        
        Write-Log "Report completed successfully: $excelFileName" -Level Success
        
        # Display results summary
        Display-ResultsSummary -Results $sortedResults -OfficeLimit $OfficeLimit -SharePointLimit $SharePointLimit
        
        # Open report if possible
        try {
            Write-Log "Opening report..." -Level Info
            Invoke-Item $excelFileName
        }
        catch {
            Write-Log "Report saved but could not be opened automatically: $excelFileName" -Level Warning
        }
    }
    catch {
        Write-Log "Script execution failed: $_" -Level Error
        Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level Debug
    }
}

# Execute the main function
Main