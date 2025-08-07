<#
.SYNOPSIS
    SharePoint URL Length Scanner for Excel Files

.DESCRIPTION
    This script identifies SharePoint files with URLs exceeding length limitations for different file types.
    Based on SharePoint Diary article: https://www.sharepointdiary.com/2019/06/sharepoint-online-url-length-limitation-in-excel-files-218-characters.html
    The script calculates the URL-encoded length (accounting for special characters like spaces becoming %20) 
    and generates a simple Excel report with 5 columns: Path+Filename Length, Full Path Length, 
    Filename Length, Issue, and Full Path with Filename.

.PARAMETER SharePointUrl
    The SharePoint URL to scan (e.g., https://tenant.sharepoint.com/sites/siteName).

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

.PARAMETER OfficeAppLimit
    Character limit for Word/PowerPoint/Access (default: 259).

.PARAMETER SharePointLimit
    Character limit for SharePoint (default: 400).

.PARAMETER PageSize
    Number of items to retrieve in each batch for SharePoint queries (default: 2000).

.EXAMPLE
    .\SharePointUrlScanner.ps1 -SharePointUrl "https://tenant.sharepoint.com/sites/siteName"
    Scans the specified SharePoint site using default authentication.

.EXAMPLE
    .\SharePointUrlScanner.ps1 -SharePointUrl "https://tenant.sharepoint.com/sites/siteName" -OutputPath "C:\Reports\UrlLengths.xlsx"
    Scans the site and saves to a custom location.

.NOTES
    File Name: SharePointUrlScanner.ps1
    Author: Your Name
    Version: 1.0
    Requires: PowerShell 5.1 or later
    Modules: PnP.PowerShell, ImportExcel
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$SharePointUrl,
    
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
    [int]$OfficeAppLimit = 259,
    
    [Parameter(Mandatory=$false)]
    [int]$SharePointLimit = 400,
    
    [Parameter(Mandatory=$false)]
    [int]$PageSize = 2000
)

# Initialize logging
$logFile = Join-Path $PSScriptRoot "SharePointUrlScanner.log"
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
        [string]$DefaultFileName = "SharePointUrlLengths.xlsx",
        [string]$Title = "Save SharePoint URL Length Report"
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

function Get-UrlEncodedLength {
    param(
        [string]$Url
    )
    
    # Get the path part of the URL (after the domain)
    $uri = [System.Uri]$Url
    $path = $uri.LocalPath + $uri.Query
    
    # URL encode the path to get the actual length as it would be in a browser
    $encodedPath = [System.Web.HttpUtility]::UrlEncode($path)
    
    # Calculate the total URL length with the encoded path
    $baseUrlLength = $Url.Length - $path.Length
    $totalLength = $baseUrlLength + $encodedPath.Length
    
    return $totalLength
}

function Get-SharePointUrlLengths {
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
            $listItemCollectionPosition = $null
            $batchCount = 0
            
            do {
                $batchCount++
                Write-Log "Processing batch $batchCount for library: $($library.Title)" -Level Info
                
                try {
                    $items = Invoke-WithRetry -ScriptBlock {
                        # Try modern pagination approach first
                        if ($listItemCollectionPosition) {
                            try {
                                Get-PnPListItem -List $library -Fields "FileRef", "File_x0020_Type" -PageSize $PageSize -ListItemCollectionPosition $listItemCollectionPosition
                            }
                            catch {
                                # Fall back to simple pagination if ListItemCollectionPosition is not supported
                                Get-PnPListItem -List $library -Fields "FileRef", "File_x0020_Type" -PageSize $PageSize
                            }
                        } else {
                            Get-PnPListItem -List $library -Fields "FileRef", "File_x0020_Type" -PageSize $PageSize
                        }
                    } -Activity "Getting items from $($library.Title)"
                    
                    $allItems += $items
                    
                    # Try to get the next position for pagination (if supported)
                    try {
                        $listItemCollectionPosition = $items.ListItemCollectionPosition
                    }
                    catch {
                        # If ListItemCollectionPosition is not available, set to null to stop pagination
                        $listItemCollectionPosition = $null
                    }
                    
                    # Log successful batch
                    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    "$timestamp,$($library.Title),Success,$($items.Count),Batch completed successfully" | Out-File -FilePath $batchLogFile -Append -Encoding utf8
                    
                    # Save recovery data after each batch
                    $recoveryData = @{
                        LongPathFiles = $longPathFiles
                        ProcessedLibraries = $processedLibraries
                        CurrentLibrary = $library.Title
                        LastBatch = $batchCount
                        LastPosition = $listItemCollectionPosition
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
            } while ($items.Count -eq $PageSize -and ($listItemCollectionPosition -ne $null))
            
            # Process files in the library
            $files = $allItems | Where-Object { $_.FileSystemObjectType -eq "File" }
            Write-Log "Found $($files.Count) files in $($library.Title)" -Level Info
            
            foreach ($file in $files) {
                $serverRelativeUrl = $file["FileRef"]
                $fileName = [System.IO.Path]::GetFileName($serverRelativeUrl)
                $fileExtension = [System.IO.Path]::GetExtension($fileName).ToLower()
                
                # Get file name without extension
                $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
                $baseFileNameLength = $baseFileName.Length
                
                # Construct full SharePoint URL
                $fullUrl = "$tenantUrl$serverRelativeUrl"
                
                # Calculate URL-encoded length for the full URL (Path+Filename Length)
                $urlLength = Get-UrlEncodedLength -Url $fullUrl
                
                # Calculate the path length without the filename
                $uri = [System.Uri]$fullUrl
                $pathWithoutFile = $uri.AbsoluteUri.Substring(0, $uri.AbsoluteUri.LastIndexOf('/') + 1)
                $pathWithoutFileLength = Get-UrlEncodedLength -Url $pathWithoutFile
                
                # Determine the issue based on file type and path length
                $issue = ""
                $hasIssue = $false
                
                # Excel-specific limit
                if ($fileExtension -in @(".xls", ".xlsx", ".xlsm", ".xlsb") -and $urlLength -gt $OfficeLimit) {
                    $issue = "Excel warning"
                    $hasIssue = $true
                }
                # Word/PowerPoint/Access limit
                elseif ($fileExtension -in @(".doc", ".docx", ".ppt", ".pptx", ".accdb", ".mdb") -and $urlLength -gt $OfficeAppLimit) {
                    $issue = "Word/PowerPoint/Access warning"
                    $hasIssue = $true
                }
                # SharePoint hard limit
                elseif ($urlLength -gt $SharePointLimit) {
                    $issue = "SharePoint limit exceeded"
                    $hasIssue = $true
                }
                
                # Only add to results if there's an issue
                if ($hasIssue) {
                    $longPathFiles += [PSCustomObject]@{
                        PathPlusFilenameLength = $urlLength
                        FullPathLength = $pathWithoutFileLength
                        FileNameLength = $baseFileNameLength
                        Issue = $issue
                        FullPath = $fullUrl
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
        [int]$OfficeAppLimit,
        [int]$SharePointLimit
    )
    
    try {
        Write-Log "Generating Excel report..." -Level Info
        
        # Remove existing file
        if (Test-Path $FilePath) {
            Remove-Item $FilePath -Force -ErrorAction SilentlyContinue
        }
        
        # Prepare data for export with the exact columns in the required order
        $exportData = $Data | Select-Object `
            @{Name="Path+Filename Length"; Expression={$_.PathPlusFilenameLength}}, `
            @{Name="Full Path Length"; Expression={$_.FullPathLength}}, `
            @{Name="Filename Length"; Expression={$_.FileNameLength}}, `
            @{Name="Issue"; Expression={$_.Issue}}, `
            @{Name="Full Path with Filename"; Expression={$_.FullPath}}
        
        # Export to Excel with searchable headers
        $exportData | Export-Excel -Path $FilePath -WorksheetName "URL Length Issues" -AutoSize -TableStyle "Medium1" -Title "SharePoint URL Length Issues"
        
        Write-Log "Excel report created successfully: $FilePath" -Level Success
    }
    catch {
        Write-Log "Failed to create Excel report: $_" -Level Error
        throw
    }
}

function Display-ResultsSummary {
    param(
        [array]$Results,
        [int]$OfficeLimit,
        [int]$OfficeAppLimit,
        [int]$SharePointLimit
    )
    
    $excelWarnings = $Results | Where-Object { $_.Issue -eq "Excel warning" }
    $officeAppWarnings = $Results | Where-Object { $_.Issue -eq "Word/PowerPoint/Access warning" }
    $sharePointLimitExceeded = $Results | Where-Object { $_.Issue -eq "SharePoint limit exceeded" }
    
    Write-Host "`nSharePoint URL Length Issues Found:" -ForegroundColor Cyan
    Write-Host "=================================" -ForegroundColor Cyan
    Write-Host "Total files with issues: $($Results.Count)" -ForegroundColor Yellow
    Write-Host "Excel warnings (>$OfficeLimit chars): $($excelWarnings.Count)" -ForegroundColor Red
    Write-Host "Office app warnings (>$OfficeAppLimit chars): $($officeAppWarnings.Count)" -ForegroundColor Yellow
    Write-Host "SharePoint limit exceeded (>$SharePointLimit chars): $($sharePointLimitExceeded.Count)" -ForegroundColor Red
    
    if ($Results.Count -gt 0) {
        $avgPathLength = [math]::Round(($Results | Measure-Object -Property PathPlusFilenameLength -Average).Average, 0)
        $maxPathLength = ($Results | Measure-Object -Property PathPlusFilenameLength -Maximum).Maximum
        Write-Host "Average URL length: $avgPathLength characters" -ForegroundColor White
        Write-Host "Maximum URL length: $maxPathLength characters" -ForegroundColor White
    }
}

function Main {
    try {
        Write-Log "Script started at $scriptStartTime" -Level Info
        Write-Log "SharePoint URL Length Scanner for Excel Files" -Level Success
        Write-Log "=========================================" -Level Success
        
        # Get SharePoint URL if not provided
        if (-not $SharePointUrl) {
            $SharePointUrl = Read-Host "Enter SharePoint URL (e.g., https://tenant.sharepoint.com/sites/siteName)"
            if (-not $SharePointUrl) {
                Write-Log "No SharePoint URL provided. Exiting." -Level Info
                return
            }
        }
        
        Write-Log "SharePoint URL: $SharePointUrl" -Level Info
        Write-Log "Excel limit: $OfficeLimit characters" -Level Info
        Write-Log "Office app limit: $OfficeAppLimit characters" -Level Info
        Write-Log "SharePoint limit: $SharePointLimit characters" -Level Info
        
        # Create output filename
        $dateStr = Get-Date -Format "yyyyMMdd_HHmm"
        $defaultFileName = "SharePointUrlLengthIssues-$dateStr.xlsx"
        
        # Get output path
        if ($OutputPath) {
            $excelFileName = $OutputPath
        } else {
            $excelFileName = Get-SaveFileDialog -DefaultFileName $defaultFileName -Title "Save SharePoint URL Length Issues Report"
            if (-not $excelFileName) {
                Write-Log "User cancelled the save dialog. Exiting." -Level Info
                return
            }
        }
        
        Write-Log "Report will be saved to: $excelFileName" -Level Info
        
        # Install required modules and scan SharePoint
        Install-RequiredModules
        $results = Get-SharePointUrlLengths -SharePointUrl $SharePointUrl -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
        
        if ($results.Count -eq 0) {
            Write-Log "No files with URL length issues found in SharePoint site." -Level Success
            return
        }
        
        # Sort results by URL length (descending)
        $sortedResults = $results | Sort-Object -Property PathPlusFilenameLength -Descending
        
        # Export results
        Export-ToExcel -FilePath $excelFileName -Data $sortedResults -OfficeLimit $OfficeLimit -OfficeAppLimit $OfficeAppLimit -SharePointLimit $SharePointLimit
        
        Write-Log "Report completed successfully: $excelFileName" -Level Success
        
        # Display results summary
        Display-ResultsSummary -Results $sortedResults -OfficeLimit $OfficeLimit -OfficeAppLimit $OfficeAppLimit -SharePointLimit $SharePointLimit
        
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