<#
SharePointPathShortener.ps1

Description:
Enhanced SharePoint Long Path Finder and Shortener based on SharePoint Diary best practices.
This script identifies SharePoint files with paths exceeding URL length limitations and attempts 
to shorten them by removing spaces and special characters from filenames.

Key Improvements from SharePoint Diary approach:
- Uses PnP PowerShell for better performance and reliability
- Handles different URL length limits (218 for Excel/Office compatibility, 400 for SharePoint)
- Improved path calculation methodology
- Better error handling and throttling
- Enhanced reporting with multiple worksheets
- Support for both simulation and live renaming modes

Features:
- Scans all SharePoint sites and document libraries using PnP PowerShell
- Identifies files exceeding 218 characters (Office compatibility) or 400 characters (SharePoint limit)
- Attempts to shorten filenames by removing spaces, special characters, and truncating
- Generates comprehensive Excel reports with before/after analysis
            try {
                $result = Rename-FileWithOptimizedName -LongUrlFile $longUrlFile -WhatIf:$SimulationMode -TargetLimit $targetLimit
                $renameResults += $result
                if ($result.Success) {
                    $limitStatus = if ($result.WouldMeetOfficeLimit) { "Office✓" } else { "Office✗" }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
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
#endregion

#region Module Management and Authentication
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

function Connect-ToPnPOnline {
    param(
        [string]$TenantUrl,
        [string]$ClientId,
        [string]$TenantId,
        [string]$CertificateThumbprint,
        [switch]$UseInteractiveAuth
    )
    
    Write-Log "Connecting to SharePoint Online..." -Level Info
    
    try {
        # Disconnect any existing connections
        try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }
        
        if ($UseInteractiveAuth) {
            Write-Log "Using interactive authentication..." -Level Info
            Connect-PnPOnline -Url $TenantUrl -Interactive -WarningAction SilentlyContinue
        }
        else {
            Write-Log "Using certificate-based app-only authentication..." -Level Info
            
            # Get certificate
            $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction Stop
            if (-not $cert) { 
                throw "Certificate with thumbprint $CertificateThumbprint not found in CurrentUser\My store." 
            }
            
            # Connect with certificate
            Connect-PnPOnline -Url $TenantUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $CertificateThumbprint -WarningAction SilentlyContinue
        }
        
        # Verify connection
        $context = Get-PnPContext
        if (-not $context) {
            throw "PnP connection failed - no context available"
        }
        
        Write-Log "Successfully connected to SharePoint Online" -Level Success
        Write-Log "Connected to: $($context.Web.Url)" -Level Info
        
        return $true
    }
    catch {
        Write-Log "Authentication failed: $_" -Level Error
        throw
    }
}

function Get-TenantInformation {
    try {
        $tenantUrl = (Get-PnPContext).Site.Url
        $tenantName = ([Uri]$tenantUrl).Host -replace '\.sharepoint\.com$', '' -replace '\-admin$', ''
        
        Write-Log "Tenant URL: $tenantUrl" -Level Info
        Write-Log "Tenant Name: $tenantName" -Level Info
        
        return @{
            TenantUrl = $tenantUrl
            TenantName = $tenantName
        }
    }
    catch {
        Write-Log "Error getting tenant information: $_" -Level Warning
        return @{
            TenantUrl = "Unknown"
            TenantName = "Tenant"
        }
    }
}
#endregion

#region Path Processing Functions - Enhanced from SharePoint Diary approach
function Get-FullFileUrl {
    param(
        [Parameter(Mandatory=$true)]
        $File,
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl
    )
    
    try {
        # Build absolute URL using site URL + server relative URL
        $siteUri = [Uri]$SiteUrl
        $baseUrl = "$($siteUri.Scheme)://$($siteUri.Host)"
        
        if ($File.ServerRelativeUrl) {
            $fullUrl = $baseUrl + $File.ServerRelativeUrl
        }
        elseif ($File.FileRef) {
            $fullUrl = $baseUrl + $File.FileRef
        }
        else {
            # Fallback method
            $fullUrl = $SiteUrl.TrimEnd('/') + "/" + $File.Name
        }
        
        return $fullUrl
    }
    catch {
        Write-Log "Error building full URL for file $($File.Name): $_" -Level Debug
        return $SiteUrl + "/" + $File.Name
    }
}

function Get-OptimizedFileName {
    param(
        [Parameter(Mandatory=$true)]
        [string]$OriginalFileName,
        [Parameter(Mandatory=$true)]
        [int]$TargetReduction,
        [int]$MinNameLength = 5
    )
    
    $extension = [System.IO.Path]::GetExtension($OriginalFileName)
    $nameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($OriginalFileName)
    
    # If the name is already very short, don't modify it
    if ($nameWithoutExtension.Length -le $MinNameLength) {
        return $OriginalFileName
    }
    
    $shortened = $nameWithoutExtension
    
    # Strategy 1: Remove all whitespace (spaces, tabs, etc.)
    $shortened = $shortened -replace '\s+', ''
    
    # Strategy 2: Remove common redundant characters
    $shortened = $shortened -replace '[-_]{2,}', '-'  # Multiple dashes/underscores to single dash
    $shortened = $shortened -replace '[(){}[\]]', ''  # Remove brackets
    
    # Strategy 3: Remove special characters but keep alphanumeric, dash, underscore, period
    $shortened = $shortened -replace '[^\w\-_.]', ''
    
    # Strategy 4: Remove repeated characters (3 or more in a row)
    $shortened = $shortened -replace '(.)\1{2,}', '$1'
    
    # Strategy 5: If still not short enough, truncate intelligently
    $currentReduction = $nameWithoutExtension.Length - $shortened.Length
    if ($currentReduction -lt $TargetReduction -and $shortened.Length -gt $MinNameLength) {
        $additionalReduction = $TargetReduction - $currentReduction
        $targetLength = [Math]::Max($MinNameLength, $shortened.Length - $additionalReduction)
        
        if ($targetLength -lt $shortened.Length) {
            # Try to truncate at word boundaries if possible
            if ($shortened.Contains('_') -or $shortened.Contains('-')) {
                $words = $shortened -split '[-_]'
                $truncated = ""
                foreach ($word in $words) {
                    if (($truncated + $word).Length -le $targetLength) {
                        $truncated += $word
                    } else {
                        break
                    }
                }
                if ($truncated.Length -ge $MinNameLength) {
                    $shortened = $truncated
                }
            }
            
            # Final truncation if needed
            if ($shortened.Length -gt $targetLength) {
                $shortened = $shortened.Substring(0, $targetLength)
            }
        }
    }
    
    # Ensure minimum length
    if ($shortened.Length -lt $MinNameLength) {
        $shortened = $nameWithoutExtension.Substring(0, [Math]::Min($MinNameLength, $nameWithoutExtension.Length))
        $shortened = $shortened -replace '\s+', ''
    }
    
    return $shortened + $extension
}

function Test-UrlLength {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FullUrl,
        [int]$OfficeLimit = 218,
        [int]$SharePointLimit = 400
    )
    
    $urlLength = $FullUrl.Length
    
    return [PSCustomObject]@{
        Length = $urlLength
        ExceedsOfficeLimit = $urlLength -gt $OfficeLimit
        ExceedsSharePointLimit = $urlLength -gt $SharePointLimit
        OfficeExcess = [Math]::Max(0, $urlLength - $OfficeLimit)
        SharePointExcess = [Math]::Max(0, $urlLength - $SharePointLimit)
    }
}
#endregion

#region Site and File Discovery - Enhanced PnP PowerShell approach
function Get-AllSharePointSites {
    Write-Log "Enumerating all SharePoint sites using PnP PowerShell..." -Level Info
    
    try {
        $sites = @()
        
        # Method 1: Get all site collections (most comprehensive)
        try {
            Write-Log "Getting site collections from tenant..." -Level Info
            $siteCollections = Invoke-WithRetry -ScriptBlock {
                Get-PnPTenantSite -IncludeOneDriveSites:$false -Detailed -ErrorAction Stop
            } -Activity "Get tenant sites"
            
            foreach ($site in $siteCollections) {
                $sites += [PSCustomObject]@{
                    Title = $site.Title
                    Url = $site.Url
                    Template = $site.Template
                    Owner = $site.Owner
                    StorageUsed = $site.StorageUsageCurrent
                    LastActivity = $site.LastContentModifiedDate
                    IsOneDrive = $site.Template -like "*SPSPERS*"
                }
            }
            
            Write-Log "Found $($sites.Count) site collections" -Level Success
        }
        catch {
            Write-Log "Failed to get tenant sites, trying alternative method: $_" -Level Warning
            
            # Method 2: Fallback - get hub sites and search
            try {
                $hubSites = Get-PnPHubSite -ErrorAction SilentlyContinue
                foreach ($hubSite in $hubSites) {
                    $sites += [PSCustomObject]@{
                        Title = $hubSite.Title
                        Url = $hubSite.SiteUrl
                        Template = "Hub Site"
                        Owner = ""
                        StorageUsed = 0
                        LastActivity = $null
                        IsOneDrive = $false
                    }
                }
            }
            catch {
                Write-Log "Hub site method also failed: $_" -Level Warning
            }
        }
        
        # Filter out OneDrive personal sites for this use case (unless specifically requested)
        $filteredSites = $sites | Where-Object { -not $_.IsOneDrive } | Sort-Object Url -Unique
        
        if ($filteredSites.Count -eq 0) {
            Write-Log "No SharePoint sites found! This might indicate permission issues." -Level Warning
            return @()
        }
        
        Write-Log "Found $($filteredSites.Count) SharePoint sites after filtering" -Level Success
        return $filteredSites
    }
    catch {
        Write-Log "Failed to enumerate SharePoint sites: $_" -Level Error
        return @()
    }
}

function Get-LongUrlFiles {
    param(
        [Parameter(Mandatory=$true)]
        $Site,
        [int]$OfficeLimit = 218,
        [int]$SharePointLimit = 400,
        [int]$PageSize = 2000,
        [switch]$ExcelFilesOnly
    )
    
    Write-Log "Scanning site for long URL files: $($Site.Title)" -Level Info
    
    $longUrlFiles = @()
    $processedFiles = 0
    
    try {
        # Connect to the specific site
        Invoke-WithRetry -ScriptBlock {
            Connect-PnPOnline -Url $Site.Url -ClientId $script:ClientId -Tenant $script:TenantId -Thumbprint $script:CertificateThumbprint -WarningAction SilentlyContinue
        } -Activity "Connect to site"
        
        # Get all document libraries
        $lists = Invoke-WithRetry -ScriptBlock {
            Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false }  # Document Libraries only
        } -Activity "Get document libraries"
        
        if (-not $lists -or $lists.Count -eq 0) {
            Write-Log "No document libraries found in site: $($Site.Title)" -Level Warning
            return @()
        }
        
        Write-Log "Found $($lists.Count) document libraries in site: $($Site.Title)" -Level Info
        
        $libraryIndex = 0
        foreach ($list in $lists) {
            $libraryIndex++
            $percentComplete = [math]::Round(($libraryIndex / $lists.Count) * 100, 1)
            
            Show-Progress -Activity "Scanning Document Libraries for Long URLs" -Status "Processing: $($list.Title) | Files processed: $processedFiles ($libraryIndex/$($lists.Count))" -PercentComplete $percentComplete
            
            try {
                # Get all files from the library using CAML query for better performance
                $query = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq></Where></Query><RowLimit>$PageSize</RowLimit></View>"
                
                $listItems = @()
                $position = $null
                
                do {
                    $batch = Invoke-WithRetry -ScriptBlock {
                        if ($position) {
                            Get-PnPListItem -List $list -Query $query -PageSize $PageSize -ListItemCollectionPosition $position
                        } else {
                            Get-PnPListItem -List $list -Query $query -PageSize $PageSize
                        }
                    } -Activity "Get list items batch"
                    
                    if ($batch -and $batch.Count -gt 0) {
                        $listItems += $batch
                        $position = $batch.ListItemCollectionPosition
                    } else {
                        $position = $null
                    }
                    
                } while ($position)
                
                foreach ($item in $listItems) {
                    try {
                        $fileName = $item["FileLeafRef"]
                        if (-not $fileName) { continue }
                        
                        # Filter for Excel files if requested
                        if ($ExcelFilesOnly) {
                            $extension = [System.IO.Path]::GetExtension($fileName).ToLower()
                            if ($extension -notin @('.xlsx', '.xls', '.xlsm', '.xlsb')) {
                                continue
                            }
                        }
                        
                        $processedFiles++
                        
                        # Build full URL
                        $fullUrl = Get-FullFileUrl -File $item -SiteUrl $Site.Url
                        
                        # Test URL length against both limits
                        $urlTest = Test-UrlLength -FullUrl $fullUrl -OfficeLimit $OfficeLimit -SharePointLimit $SharePointLimit
                        
                        # If file exceeds either limit, add it to results
                        if ($urlTest.ExceedsOfficeLimit -or $urlTest.ExceedsSharePointLimit) {
                            $fileSize = 0
                            $lastModified = $null
                            
                            try {
                                $fileSize = [int]$item["File_x0020_Size"]
                                $lastModified = $item["Modified"]
                            } catch { }
                            
                            $longUrlFiles += [PSCustomObject]@{
                                SiteName = $Site.Title
                                SiteUrl = $Site.Url
                                LibraryName = $list.Title
                                FileName = $fileName
                                FileExtension = [System.IO.Path]::GetExtension($fileName)
                                FilePath = $item["FileDirRef"]
                                FullUrl = $fullUrl
                                UrlLength = $urlTest.Length
                                ExceedsOfficeLimit = $urlTest.ExceedsOfficeLimit
                                ExceedsSharePointLimit = $urlTest.ExceedsSharePointLimit
                                OfficeExcess = $urlTest.OfficeExcess
                                SharePointExcess = $urlTest.SharePointExcess
                                FileSize = $fileSize
                                FileSizeMB = [math]::Round($fileSize / 1MB, 2)
                                LastModified = $lastModified
                                ItemId = $item.Id
                                ListId = $list.Id
                                ServerRelativeUrl = $item["FileRef"]
                            }
                        }
                        
                        # Update progress periodically
                        if ($processedFiles % 100 -eq 0) {
                            $currentFileName = if ($fileName.Length -gt 50) { 
                                $fileName.Substring(0, 47) + "..." 
                            } else { 
                                $fileName 
                            }
                            Show-Progress -Activity "Scanning Document Libraries for Long URLs" -Status "Processing: $($list.Title) | Long URLs found: $($longUrlFiles.Count) | Files processed: $processedFiles" -PercentComplete $percentComplete -CurrentOperation "Current: $currentFileName"
                        }
                        
                    }
                    catch {
                        Write-Log "Error processing file item: $_" -Level Debug
                        continue
                    }
                }
                
                Write-Log "Processed library: $($list.Title) - Found $($longUrlFiles.Count) long URL files so far" -Level Info
            }
            catch {
                Write-Log "Failed to process library $($list.Title): $_" -Level Error
            }
        }
        
        Write-Log "Completed site: $($Site.Title) - Found $($longUrlFiles.Count) files with long URLs" -Level Success
        return $longUrlFiles
    }
    catch {
        Write-Log "Failed to scan site $($Site.Title) for long URL files: $_" -Level Error
        return @()
    }
}

function Rename-FileWithOptimizedName {
    param(
        [Parameter(Mandatory=$true)]
        $LongUrlFile,
        [switch]$WhatIf = $true,
        [int]$TargetLimit = 218
    )
    
    try {
        $originalFileName = $LongUrlFile.FileName
        $currentExcess = if ($TargetLimit -eq 218) { $LongUrlFile.OfficeExcess } else { $LongUrlFile.SharePointExcess }
        
        # Calculate how much we need to reduce the filename
        $targetReduction = $currentExcess + 10  # Add buffer for safety
        
        $optimizedFileName = Get-OptimizedFileName -OriginalFileName $originalFileName -TargetReduction $targetReduction
        
        # Calculate new URL length
        $newFullUrl = $LongUrlFile.FullUrl.Replace($originalFileName, $optimizedFileName)
        $newUrlLength = $newFullUrl.Length
        
        $result = [PSCustomObject]@{
            Success = $false
            OriginalFileName = $originalFileName
            OptimizedFileName = $optimizedFileName
            OriginalUrlLength = $LongUrlFile.UrlLength
            NewUrlLength = $newUrlLength
            LengthReduction = $LongUrlFile.UrlLength - $newUrlLength
            NewFullUrl = $newFullUrl
            ErrorMessage = ""
            Action = if ($WhatIf) { "Simulation" } else { "Attempted Rename" }
            WouldMeetOfficeLimit = $newUrlLength -le 218
            WouldMeetSharePointLimit = $newUrlLength -le 400
        }
        
        # Check if optimization would actually help
        if ($newUrlLength -ge $LongUrlFile.UrlLength -or $result.LengthReduction -lt 5) {
            $result.ErrorMessage = "Filename optimization would not provide sufficient URL length reduction"
            return $result
        }
        
        if ($WhatIf) {
            # Simulation mode - just return what would happen
            $result.Success = $true
            return $result
        }
        
        # Actual rename operation
        try {
            # Connect to the site containing the file
            Invoke-WithRetry -ScriptBlock {
                Connect-PnPOnline -Url $LongUrlFile.SiteUrl -ClientId $script:ClientId -Tenant $script:TenantId -Thumbprint $script:CertificateThumbprint -WarningAction SilentlyContinue
            } -Activity "Connect to site for rename"
            
            # Get the file and rename it
            $file = Invoke-WithRetry -ScriptBlock {
                Get-PnPFile -Url $LongUrlFile.ServerRelativeUrl -AsListItem -ErrorAction Stop
            } -Activity "Get file for rename"
            
            if ($file) {
                # Perform the rename
                Invoke-WithRetry -ScriptBlock {
                    Set-PnPListItem -List $LongUrlFile.ListId -Identity $file.Id -Values @{"FileLeafRef" = $optimizedFileName} -ErrorAction Stop
                } -Activity "Rename file"
                
                $result.Success = $true
                Write-Log "Successfully renamed: $originalFileName -> $optimizedFileName (saved $($result.LengthReduction) chars)" -Level Success
            }
            else {
                $result.ErrorMessage = "Could not retrieve file for renaming"
            }
        }
        catch {
            $result.ErrorMessage = "Rename operation failed: $($_.Exception.Message)"
            Write-Log "Failed to rename file $originalFileName : $($_.Exception.Message)" -Level Error
        }
        
        return $result
    }
    catch {
        Write-Log "Error in rename operation: $_" -Level Error
        return [PSCustomObject]@{
            Success = $false
            OriginalFileName = $LongUrlFile.FileName
            OptimizedFileName = ""
            OriginalUrlLength = $LongUrlFile.UrlLength
            NewUrlLength = 0
            LengthReduction = 0
            NewFullUrl = ""
            ErrorMessage = "Unexpected error: $($_.Exception.Message)"
            Action = "Failed"
            WouldMeetOfficeLimit = $false
            WouldMeetSharePointLimit = $false
        }
    }
}
#endregion

#region Excel Report Generation - Enhanced with multiple worksheets
function Get-SaveFileDialog {
    param(
        [string]$InitialDirectory = [Environment]::GetFolderPath('Desktop'),
        [string]$Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
        [string]$DefaultFileName = "SharePointLongUrlReport.xlsx",
        [string]$Title = "Save SharePoint Long URL Report"
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

function Export-EnhancedLongUrlReport {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ExcelFileName,
        
        [Parameter(Mandatory=$true)]
        [array]$AllLongUrlFiles,
        
        [Parameter(Mandatory=$true)]
        [array]$RenameResults,
        
        [Parameter(Mandatory=$true)]
        [array]$SiteSummaries,
        
        [string]$TenantName = "Tenant",
        [int]$OfficeLimit = 218,
        [int]$SharePointLimit = 400
    )
    
    Write-Log "Generating comprehensive Excel report..." -Level Info
    
    # Remove existing file
    if (Test-Path $ExcelFileName) {
        Remove-Item $ExcelFileName -Force -ErrorAction SilentlyContinue
    }
    
    try {
        # Separate files by limit type
        $officeProblems = $AllLongUrlFiles | Where-Object { $_.ExceedsOfficeLimit }
        $sharePointProblems = $AllLongUrlFiles | Where-Object { $_.ExceedsSharePointLimit -and -not $_.ExceedsOfficeLimit }
        
        # Prepare rename results
        $successfulRenames = $RenameResults | Where-Object { $_.Success }
        $failedRenames = $RenameResults | Where-Object { -not $_.Success }
        
        # Calculate summary statistics
        $totalLongUrlFiles = $AllLongUrlFiles.Count
        $totalOfficeIssues = $officeProblems.Count
        $totalSharePointIssues = $sharePointProblems.Count
        $totalAttempted = $RenameResults.Count
        $totalSuccessful = $successfulRenames.Count
        $totalFailed = $failedRenames.Count
        $avgUrlLength = if ($AllLongUrlFiles.Count -gt 0) { [math]::Round(($AllLongUrlFiles | Measure-Object -Property UrlLength -Average).Average, 0) } else { 0 }
        $maxUrlLength = if ($AllLongUrlFiles.Count -gt 0) { ($AllLongUrlFiles | Measure-Object -Property UrlLength -Maximum).Maximum } else { 0 }
        $totalCharactersSaved = if ($successfulRenames.Count -gt 0) { ($successfulRenames | Measure-Object -Property LengthReduction -Sum).Sum } else { 0 }
        
        # Create summary data
        $summaryData = @(
            [PSCustomObject]@{ 'Metric' = 'Analysis Date'; 'Value' = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
            [PSCustomObject]@{ 'Metric' = 'Tenant Name'; 'Value' = $TenantName }
            [PSCustomObject]@{ 'Metric' = 'Office Compatibility Limit'; 'Value' = "$OfficeLimit characters" }
            [PSCustomObject]@{ 'Metric' = 'SharePoint Online Limit'; 'Value' = "$SharePointLimit characters" }
            [PSCustomObject]@{ 'Metric' = ''; 'Value' = '' }  # Spacer
            [PSCustomObject]@{ 'Metric' = 'FINDINGS SUMMARY'; 'Value' = '' }
            [PSCustomObject]@{ 'Metric' = 'Total Files with Long URLs'; 'Value' = $totalLongUrlFiles }
            [PSCustomObject]@{ 'Metric' = 'Files Exceeding Office Limit (218)'; 'Value' = $totalOfficeIssues }
            [PSCustomObject]@{ 'Metric' = 'Files Exceeding SharePoint Limit (400)'; 'Value' = $totalSharePointIssues }
            [PSCustomObject]@{ 'Metric' = 'Average URL Length'; 'Value' = $avgUrlLength }
            [PSCustomObject]@{ 'Metric' = 'Maximum URL Length Found'; 'Value' = $maxUrlLength }
            [PSCustomObject]@{ 'Metric' = ''; 'Value' = '' }  # Spacer
            [PSCustomObject]@{ 'Metric' = 'REMEDIATION RESULTS'; 'Value' = '' }
            [PSCustomObject]@{ 'Metric' = 'Total Rename Attempts'; 'Value' = $totalAttempted }
            [PSCustomObject]@{ 'Metric' = 'Successfully Optimized'; 'Value' = $totalSuccessful }
            [PSCustomObject]@{ 'Metric' = 'Failed to Optimize'; 'Value' = $totalFailed }
            [PSCustomObject]@{ 'Metric' = 'Success Rate (%)'; 'Value' = if ($totalAttempted -gt 0) { [math]::Round(($totalSuccessful / $totalAttempted) * 100, 1) } else { 0 } }
            [PSCustomObject]@{ 'Metric' = 'Total Characters Saved'; 'Value' = $totalCharactersSaved }
            [PSCustomObject]@{ 'Metric' = ''; 'Value' = '' }  # Spacer
            [PSCustomObject]@{ 'Metric' = 'RECOMMENDATIONS'; 'Value' = '' }
            [PSCustomObject]@{ 'Metric' = 'Priority 1: Office Compatibility'; 'Value' = "Fix $totalOfficeIssues files exceeding 218 chars" }
            [PSCustomObject]@{ 'Metric' = 'Priority 2: SharePoint Limits'; 'Value' = "Fix $totalSharePointIssues files exceeding 400 chars" }
            [PSCustomObject]@{ 'Metric' = 'Best Practice'; 'Value' = "Keep file paths under 218 characters for maximum compatibility" }
        )
        
        # Export Summary worksheet first
        Write-Log "Creating Summary worksheet..." -Level Info
        $summaryData | Export-Excel -Path $ExcelFileName -WorksheetName "Summary" -AutoSize -TableStyle "Medium6" -Title "SharePoint Long URL Analysis Summary & Recommendations"
        
        # Worksheet 1: Problem Files (Office Compatibility - 218+ chars)
        if ($officeProblems -and $officeProblems.Count -gt 0) {
            Write-Log "Creating Office Problems worksheet with $($officeProblems.Count) entries..." -Level Info
            $officeData = $officeProblems | Select-Object @{N='Site Name';E={$_.SiteName}}, 
                @{N='Library';E={$_.LibraryName}}, 
                @{N='File Name';E={$_.FileName}}, 
                @{N='Extension';E={$_.FileExtension}}, 
                @{N='Full URL';E={$_.FullUrl}}, 
                @{N='URL Length';E={$_.UrlLength}}, 
                @{N='Excess Characters';E={$_.OfficeExcess}}, 
                @{N='File Size (MB)';E={$_.FileSizeMB}}, 
                @{N='Last Modified';E={$_.LastModified}},
                @{N='Priority';E={"HIGH - Office Incompatible"}}
                
            $officeData | Export-Excel -Path $ExcelFileName -WorksheetName "Office Problems (218+)" -AutoSize -TableStyle "Bad" -Title "Files Exceeding Office Compatibility Limit (218 characters) - HIGH PRIORITY"
        }
        
        # Worksheet 2: Problem Files (SharePoint Limit - 400+ chars)
        if ($sharePointProblems -and $sharePointProblems.Count -gt 0) {
            Write-Log "Creating SharePoint Problems worksheet with $($sharePointProblems.Count) entries..." -Level Info
            $sharePointData = $sharePointProblems | Select-Object @{N='Site Name';E={$_.SiteName}}, 
                @{N='Library';E={$_.LibraryName}}, 
                @{N='File Name';E={$_.FileName}}, 
                @{N='Extension';E={$_.FileExtension}}, 
                @{N='Full URL';E={$_.FullUrl}}, 
                @{N='URL Length';E={$_.UrlLength}}, 
                @{N='Excess Characters';E={$_.SharePointExcess}}, 
                @{N='File Size (MB)';E={$_.FileSizeMB}}, 
                @{N='Last Modified';E={$_.LastModified}},
                @{N='Priority';E={"MEDIUM - SharePoint Limit"}}
                
            $sharePointData | Export-Excel -Path $ExcelFileName -WorksheetName "SharePoint Problems (400+)" -AutoSize -TableStyle "Accent1" -Title "Files Exceeding SharePoint Limit (400 characters) - MEDIUM PRIORITY"
        }
        
        # Worksheet 3: Successfully Optimized Files
        if ($successfulRenames -and $successfulRenames.Count -gt 0) {
            Write-Log "Creating Successfully Optimized worksheet with $($successfulRenames.Count) entries..." -Level Info
            $successData = $successfulRenames | Select-Object @{N='Original File Name';E={$_.OriginalFileName}}, 
                @{N='Optimized File Name';E={$_.OptimizedFileName}}, 
                @{N='Original URL Length';E={$_.OriginalUrlLength}}, 
                @{N='New URL Length';E={$_.NewUrlLength}}, 
                @{N='Characters Saved';E={$_.LengthReduction}}, 
                @{N='Meets Office Limit';E={if($_.WouldMeetOfficeLimit){"✓ Yes"}else{"✗ No"}}}, 
                @{N='Meets SharePoint Limit';E={if($_.WouldMeetSharePointLimit){"✓ Yes"}else{"✗ No"}}}, 
                @{N='New Full URL';E={$_.NewFullUrl}}, 
                @{N='Action Status';E={$_.Action}}
                
            $successData | Export-Excel -Path $ExcelFileName -WorksheetName "Successfully Optimized" -AutoSize -TableStyle "Good" -Title "Files Successfully Optimized with Shortened Names"
        } else {
            # Create placeholder worksheet
            $placeholderSuccess = @([PSCustomObject]@{
                'Original File Name' = 'No files were successfully optimized'
                'Optimized File Name' = 'Run script in live mode to perform actual optimization'
                'Original URL Length' = ''
                'New URL Length' = ''
                'Characters Saved' = ''
                'Meets Office Limit' = ''
                'Meets SharePoint Limit' = ''
                'New Full URL' = ''
                'Action Status' = 'Simulation Mode'
            })
            $placeholderSuccess | Export-Excel -Path $ExcelFileName -WorksheetName "Successfully Optimized" -AutoSize -TableStyle "Good" -Title "Files Successfully Optimized with Shortened Names"
        }
        
        # Worksheet 4: Cannot Optimize
        if ($failedRenames -and $failedRenames.Count -gt 0) {
            Write-Log "Creating Cannot Optimize worksheet with $($failedRenames.Count) entries..." -Level Info
            $failData = $failedRenames | Select-Object @{N='Original File Name';E={$_.OriginalFileName}}, 
                @{N='Attempted New Name';E={$_.OptimizedFileName}}, 
                @{N='Original URL Length';E={$_.OriginalUrlLength}}, 
                @{N='Potential New Length';E={$_.NewUrlLength}}, 
                @{N='Potential Reduction';E={$_.LengthReduction}}, 
                @{N='Would Meet Office Limit';E={if($_.WouldMeetOfficeLimit){"✓ Yes"}else{"✗ No"}}}, 
                @{N='Would Meet SharePoint Limit';E={if($_.WouldMeetSharePointLimit){"✓ Yes"}else{"✗ No"}}}, 
                @{N='Error/Reason';E={$_.ErrorMessage}}, 
                @{N='Recommended Action';E={"Manual review required - consider moving to shorter path"}}
                
            $failData | Export-Excel -Path $ExcelFileName -WorksheetName "Cannot Optimize" -AutoSize -TableStyle "Bad" -Title "Files That Cannot Be Sufficiently Optimized - Manual Intervention Required"
        } else {
            # Create placeholder worksheet
            $placeholderFailed = @([PSCustomObject]@{
                'Original File Name' = if ($totalAttempted -eq 0) { 'No optimization attempts were made' } else { 'All attempted optimizations were successful' }
                'Attempted New Name' = ''
                'Original URL Length' = ''
                'Potential New Length' = ''
                'Potential Reduction' = ''
                'Would Meet Office Limit' = ''
                'Would Meet SharePoint Limit' = ''
                'Error/Reason' = if ($totalAttempted -eq 0) { 'Run in live mode to attempt optimization' } else { 'No failures occurred' }
                'Recommended Action' = ''
            })
            $placeholderFailed | Export-Excel -Path $ExcelFileName -WorksheetName "Cannot Optimize" -AutoSize -TableStyle "Bad" -Title "Files That Cannot Be Sufficiently Optimized"
        }
        
        # Worksheet 5: Site Summary
        if ($SiteSummaries -and $SiteSummaries.Count -gt 0) {
            Write-Log "Creating Site Summary worksheet..." -Level Info
            $SiteSummaries | Export-Excel -Path $ExcelFileName -WorksheetName "Site Summary" -AutoSize -TableStyle "Medium2" -Title "Long URL Issues by Site - Prioritization Guide"
        }
        
        # Worksheet 6: All Problem Files (Combined view for export/filtering)
        if ($AllLongUrlFiles -and $AllLongUrlFiles.Count -gt 0) {
            Write-Log "Creating All Problems worksheet with $($AllLongUrlFiles.Count) entries..." -Level Info
            $allData = $AllLongUrlFiles | Select-Object @{N='Site Name';E={$_.SiteName}}, 
                @{N='Site URL';E={$_.SiteUrl}}, 
                @{N='Library';E={$_.LibraryName}}, 
                @{N='File Name';E={$_.FileName}}, 
                @{N='Extension';E={$_.FileExtension}}, 
                @{N='Full URL';E={$_.FullUrl}}, 
                @{N='URL Length';E={$_.UrlLength}}, 
                @{N='Exceeds Office Limit';E={if($_.ExceedsOfficeLimit){"Yes"}else{"No"}}}, 
                @{N='Exceeds SharePoint Limit';E={if($_.ExceedsSharePointLimit){"Yes"}else{"No"}}}, 
                @{N='Office Excess';E={$_.OfficeExcess}}, 
                @{N='SharePoint Excess';E={$_.SharePointExcess}}, 
                @{N='File Size (MB)';E={$_.FileSizeMB}}, 
                @{N='Last Modified';E={$_.LastModified}}, 
                @{N='Server Relative URL';E={$_.ServerRelativeUrl}}
                
            $allData | Export-Excel -Path $ExcelFileName -WorksheetName "All Problems (Raw Data)" -AutoSize -TableStyle "Light1" -Title "All Files with Long URL Issues - Complete Dataset"
        }
        
        Write-Log "Excel report created successfully: $ExcelFileName" -Level Success
        return $true
    }
    catch {
        Write-Log "Failed to create Excel report: $_" -Level Error
        throw
    }
}
#endregion

#region Main Processing Function
function Main {
    try {
        Write-Host "[DEBUG] Entered Main function" -ForegroundColor Magenta
        # Prompt for SharePoint site/folder URL or local path at the very start
        if (-not $TargetPathOrUrl) {
            Write-Host "[DEBUG] Prompting for TargetPathOrUrl" -ForegroundColor Magenta
            $TargetPathOrUrl = Read-Host "Enter SharePoint site/folder URL or local path to scan (e.g. https://tenant.sharepoint.com/sites/Site/Shared%20Documents or C:\folder\subfolder)"
        }
        Write-Host "[DEBUG] TargetPathOrUrl: $TargetPathOrUrl" -ForegroundColor Magenta
        Write-Log "Target for scan: $TargetPathOrUrl" -Level Info

        Write-Log "SharePoint Long URL Finder and Optimizer (Enhanced)" -Level Success
        Write-Log "================================================" -Level Success
        Write-Log "Based on SharePoint Diary best practices with PnP PowerShell" -Level Info
        Write-Log "Mode: $(if ($SimulationMode) { 'SIMULATION (No actual renaming)' } else { 'LIVE (Will attempt to rename files)' })" -Level $(if ($SimulationMode) { 'Info' } else { 'Warning' })
        Write-Log "Office Compatibility Limit: $OfficeCompatibilityLimit characters" -Level Info
        Write-Log "SharePoint Online Limit: $SharePointLimit characters" -Level Info

        # Store parameters in script scope for use in other functions
        $script:ClientId = $ClientId
        $script:TenantId = $TenantId
        $script:CertificateThumbprint = $CertificateThumbprint

        # Install and import required modules
        Install-RequiredModules

        # Determine tenant URL if not provided
        if (-not $TenantUrl) {
            if ($TenantId) {
                # Try to construct tenant URL from tenant ID
                # This is a basic approach - in practice you might need to discover this differently
                $TenantUrl = "https://$($TenantId.Split('-')[0]).sharepoint.com"
                Write-Log "Auto-detected tenant URL: $TenantUrl" -Level Info
            } else {
                throw "TenantUrl parameter is required when TenantId is not provided or cannot be used to auto-detect"
            }
        }

        # Connect to SharePoint Online
        Connect-ToPnPOnline -TenantUrl $TenantUrl -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -UseInteractiveAuth:$UseInteractiveAuth

        # Get tenant information
        $tenantInfo = Get-TenantInformation
        $script:tenantName = $tenantInfo.TenantName

        # Create filename
        $script:dateStr = Get-Date -Format "yyyyMMdd_HHmm"
        $modePrefix = if ($SimulationMode) { "SIMULATION" } else { "LIVE" }
        $fileFilter = if ($ExcelFilesOnly) { "ExcelOnly" } else { "AllFiles" }
        $defaultFileName = "SharePointLongURL-$modePrefix-$fileFilter-$($script:tenantName)-$($script:dateStr).xlsx"

        # Get output path
        if ($OutputPath) {
            $script:excelFileName = $OutputPath
        } else {
            $script:excelFileName = Get-SaveFileDialog -DefaultFileName $defaultFileName -Title "Save SharePoint Long URL Report"
            if (-not $script:excelFileName) {
                Write-Log "User cancelled the save dialog. Exiting." -Level Info
                return
            }
        }

        Write-Log "Report will be saved to: $($script:excelFileName)" -Level Info

        # If the user entered a SharePoint URL, scan just that site/folder
        $allLongUrlFiles = @()
        $siteSummaries = @()
        if ($TargetPathOrUrl -match '^https?://') {
            # It's a SharePoint URL
            Write-Log "Scanning single SharePoint site/folder: $TargetPathOrUrl" -Level Info
            $siteObj = [PSCustomObject]@{ Title = $TargetPathOrUrl; Url = $TargetPathOrUrl; Template = ''; StorageUsed = 0; LastActivity = '' }
            $longUrlFiles = Get-LongUrlFiles -Site $siteObj -OfficeLimit $OfficeCompatibilityLimit -SharePointLimit $SharePointLimit -PageSize $PageSize -ExcelFilesOnly:$ExcelFilesOnly
            $allLongUrlFiles += $longUrlFiles
            # Create site summary
            $officeIssues = ($longUrlFiles | Where-Object { $_.ExceedsOfficeLimit }).Count
            $sharePointIssues = ($longUrlFiles | Where-Object { $_.ExceedsSharePointLimit }).Count
            $totalIssues = $longUrlFiles.Count
            $siteSummaries += [PSCustomObject]@{
                'Site Name' = $TargetPathOrUrl
                'Site URL' = $TargetPathOrUrl
                'Total Long URL Files' = $totalIssues
                'Office Issues (218+)' = $officeIssues
                'SharePoint Issues (400+)' = $sharePointIssues
                'Priority' = if ($officeIssues -gt 0) { "HIGH" } elseif ($sharePointIssues -gt 0) { "MEDIUM" } else { "LOW" }
                'Recommended Action' = if ($totalIssues -eq 0) { "No action needed" } 
                                     elseif ($officeIssues -gt 10) { "Immediate attention required" } 
                                     elseif ($officeIssues -gt 0) { "Address office compatibility issues" } 
                                     else { "Monitor SharePoint limits" }
                'Template' = ''
                'Storage Used (MB)' = ''
                'Last Activity' = ''
            }
        } elseif (Test-Path $TargetPathOrUrl) {
            Write-Log "Local folder scan is not implemented in this script version." -Level Warning
        } else {
            Write-Log "Invalid path or URL provided: $TargetPathOrUrl" -Level Error
            return
        }

        Write-Log "Scan complete! Found $($allLongUrlFiles.Count) files with long URLs" -Level Success

        # ...existing code...

        $officeProblems = $allLongUrlFiles | Where-Object { $_.ExceedsOfficeLimit }
        $sharePointOnlyProblems = $allLongUrlFiles | Where-Object { $_.ExceedsSharePointLimit -and -not $_.ExceedsOfficeLimit }
        
        Write-Host "`nLong URL Files Analysis Results:" -ForegroundColor Cyan
        Write-Host "================================" -ForegroundColor Cyan
        Write-Host "Total files with long URLs: $($allLongUrlFiles.Count)" -ForegroundColor Yellow
        Write-Host "Files exceeding Office limit (218 chars): $($officeProblems.Count)" -ForegroundColor Red
        Write-Host "Files exceeding SharePoint limit (400 chars): $($sharePointOnlyProblems.Count)" -ForegroundColor Yellow
        Write-Host "Average URL length: $([math]::Round(($allLongUrlFiles | Measure-Object -Property UrlLength -Average).Average, 0)) characters" -ForegroundColor White
        Write-Host "Maximum URL length: $(($allLongUrlFiles | Measure-Object -Property UrlLength -Maximum).Maximum) characters" -ForegroundColor White
        
        if ($allLongUrlFiles.Count -eq 0) {
            Write-Log "No files with long URLs found. Creating summary report..." -Level Info
            
            # Create empty results for report
            $renameResults = @()
            Export-EnhancedLongUrlReport -ExcelFileName $script:excelFileName -AllLongUrlFiles @() -RenameResults @() -SiteSummaries $siteSummaries -TenantName $script:tenantName -OfficeLimit $OfficeCompatibilityLimit -SharePointLimit $SharePointLimit
            
            Write-Log "Report completed successfully: $($script:excelFileName)" -Level Success
            return
        }
        
        # Show top 10 longest URLs
        Write-Host "`nTop 10 Longest URLs:" -ForegroundColor Cyan
        $topLongUrls = $allLongUrlFiles | Sort-Object UrlLength -Descending | Select-Object -First 10
        foreach ($file in $topLongUrls) {
            $priorityColor = if ($file.ExceedsOfficeLimit) { "Red" } else { "Yellow" }
            Write-Host "$($file.UrlLength) chars: $($file.FileName)" -ForegroundColor $priorityColor
        }
        
        # Process files for optimization (simulation or actual)
        Write-Log "Processing files for optimization..." -Level Info
        $renameResults = @()
        $processedCount = 0

        # Prioritize Office compatibility issues first
        $filesToProcess = $allLongUrlFiles | Sort-Object @{Expression={$_.ExceedsOfficeLimit}; Descending=$true}, UrlLength -Descending

        foreach ($longUrlFile in $filesToProcess) {
            $processedCount++
            $percentComplete = [math]::Round(($processedCount / $allLongUrlFiles.Count) * 100, 1)

            $targetLimit = if ($longUrlFile.ExceedsOfficeLimit) { $OfficeCompatibilityLimit } else { $SharePointLimit }
            $priority = if ($longUrlFile.ExceedsOfficeLimit) { "HIGH" } else { "MEDIUM" }

            Show-Progress -Activity "Processing Files for Optimization" -Status "$(if ($SimulationMode) { 'Simulating' } else { 'Attempting' }) optimization ($priority priority)" -PercentComplete $percentComplete -CurrentOperation "$processedCount of $($allLongUrlFiles.Count) files"

            try {
                $result = Rename-FileWithOptimizedName -LongUrlFile $longUrlFile -WhatIf:$SimulationMode -TargetLimit $targetLimit
                $renameResults += $result
            }
            catch {
                Write-Log "Error optimizing file: $_" -Level Error
            }
        }

        # Display priority recommendations
        Write-Host "`n*** PRIORITY RECOMMENDATIONS ***" -ForegroundColor Cyan
        Write-Host "1. HIGH PRIORITY: Fix $($officeProblems.Count) files exceeding Office limit (218 chars)" -ForegroundColor Red
        Write-Host "2. MEDIUM PRIORITY: Fix $($sharePointOnlyProblems.Count) files exceeding SharePoint limit (400 chars)" -ForegroundColor Yellow
        Write-Host "3. BEST PRACTICE: Keep all new files under 218 characters for maximum compatibility" -ForegroundColor Green
        Write-Host "4. Review the 'Site Summary' worksheet to prioritize which sites need immediate attention" -ForegroundColor White
    }
    }
    catch {
        Write-Log "Script execution failed: $_" -Level Error
        Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level Debug
    }
    finally {
        # Always disconnect from PnP Online
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            Write-Log "Disconnected from SharePoint Online" -Level Info
        }
        catch {
            # Silently handle disconnect errors
        }
    }
}
#endregion

#region Script Execution
# Execute the main function
Write-Host "[DEBUG] Script file loaded, about to call Main" -ForegroundColor Cyan
Main
#endregion

<#
USAGE EXAMPLES:

1. SIMULATION MODE (Safe - Recommended First Run):
   .\SharePointPathShortener.ps1

2. LIVE MODE (Actually rename files - USE WITH CAUTION):
   .\SharePointPathShortener.ps1 -SimulationMode:$false

3. FOCUS ON EXCEL FILES ONLY:
   .\SharePointPathShortener.ps1 -ExcelFilesOnly

4. CUSTOM LIMITS AND OUTPUT:
   .\SharePointPathShortener.ps1 -OfficeCompatibilityLimit 200 -SharePointLimit 350 -OutputPath "C:\Reports\MyReport.xlsx"

5. INTERACTIVE AUTHENTICATION (if certificate auth fails):
   .\SharePointPathShortener.ps1 -UseInteractiveAuth

6. CUSTOM TENANT URL:
   .\SharePointPathShortener.ps1 -TenantUrl "https://contoso.sharepoint.com"

IMPROVEMENTS OVER BASIC SHAREPOINT DIARY APPROACH:
- Uses PnP PowerShell for better performance and reliability
- Handles both Office (218) and SharePoint (400) character limits
- Enhanced error handling with intelligent retry logic
- Comprehensive Excel reporting with multiple worksheets
- Site prioritization and summary analysis  
- Simulation mode for safety
- Parallel processing capabilities
- Better path calculation methodology
- Certificate-based app-only authentication
- Support for filtering (Excel files only)
- Detailed logging and progress tracking

SAFETY FEATURES:
- Defaults to simulation mode
- Comprehensive logging
- Retry logic for throttling
- Error handling for failed sites
- Backup recommendations in output

EXCEL REPORT STRUCTURE:
- Summary: Overall statistics and recommendations
- Office Problems (218+): High priority files
- SharePoint Problems (400+): Medium priority files  
- Successfully Optimized: Files that were renamed
- Cannot Optimize: Files requiring manual intervention
- Site Summary: Per-site prioritization guide
- All Problems (Raw Data): Complete dataset for analysis

KEY DIFFERENCES FROM SHAREPOINT DIARY BASIC SCRIPTS:
1. Uses modern PnP PowerShell instead of older CSOM
2. Handles multiple character limits (Office vs SharePoint)
3. Includes comprehensive error handling and retry logic
4. Provides detailed Excel reporting with multiple worksheets
5. Implements simulation mode for safety
6. Supports both certificate and interactive authentication
7. Includes site prioritization and summary analysis
8. Better filename optimization strategies
9. Handles throttling and service issues gracefully
10. Provides actionable recommendations in output

CERTIFICATE AUTHENTICATION SETUP:
1. Create app registration in Azure AD
2. Generate certificate and upload to app registration
3. Grant appropriate SharePoint permissions
4. Use certificate thumbprint in script parameters

REQUIRED PERMISSIONS:
- Sites.Read.All (to read site structure)
- Sites.ReadWrite.All (to rename files in live mode)
- User.Read.All (for user context)

#>