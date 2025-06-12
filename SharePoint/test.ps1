<#
.SYNOPSIS
  Generates a focused SharePoint storage and access report with performance optimizations

.DESCRIPTION
  Creates an Excel report showing storage usage with:
  - Faster scanning through parallel processing
  - Progress tracking for all operations
  - Optimized API calls
#>

# Set strict error handling
$ErrorActionPreference = "Stop"
$WarningPreference = "SilentlyContinue"

#--- Configuration ---
$clientId = '278b9af9-888d-4344-93bb-769bdd739249'
$tenantId = 'ca0711e2-e703-4f4e-9099-17d97863211c'
$certificateThumbprint = 'B0AF0EF7659EA83D3140844F4BF89CCBB9413DBA'
$siteUrl = 'https://fbaint.sharepoint.com/sites/Marketing'

#--- Performance Optimizations ---
$ProgressPreference = 'Continue' # Enable progress display
$BatchSize = 200 # Items per API call
$MaxThreads = 5 # Parallel processing threads
$ThrottleDelay = 200 # Milliseconds between API calls

#--- Required Modules ---
$requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Sites', 
    'Microsoft.Graph.Files',
    'ImportExcel',
    'ThreadJob' # For parallel processing
)

# Install and import required modules
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing $module..." -ForegroundColor Yellow
        Install-Module -Name $module -Force -AllowClobber -SkipPublisherCheck
    }
    Import-Module -Name $module -Force
}

#--- Progress Tracking Functions ---
function Show-Progress {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete,
        [int]$SecondsRemaining,
        [int]$Id = 1
    )
    
    Write-Progress -Activity $Activity -Status $Status `
        -PercentComplete $PercentComplete `
        -SecondsRemaining $SecondsRemaining `
        -Id $Id
}

#--- Optimized Authentication ---
function Connect-ToGraph {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    
    # Clear existing connections
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    
    # Get certificate
    $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$certificateThumbprint" -ErrorAction Stop
    
    # Connect with app-only authentication
    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -Certificate $cert -NoWelcome
    
    # Verify app-only authentication
    $context = Get-MgContext
    if ($context.AuthType -ne 'AppOnly') {
        throw "App-only authentication required. Current: $($context.AuthType)"
    }
    
    Write-Host "Successfully connected with app-only authentication" -ForegroundColor Green
}

#--- Optimized Site Resolution ---
function Get-TargetSite {
    param([string]$Url)
    
    Write-Host "Resolving site information..." -ForegroundColor Cyan
    
    # Extract site ID from URL
    $uri = [Uri]$Url
    $sitePath = $uri.AbsolutePath.TrimEnd('/')
    $siteId = "$($uri.Host):$sitePath"
    
    $site = Get-MgSite -SiteId $siteId -ErrorAction Stop
    Write-Host "Found site: $($site.DisplayName)" -ForegroundColor Green
    
    return $site
}

#--- Parallel Processing for Faster Scanning ---
function Get-TotalItemCount {
    param(
        [string]$DriveId,
        [string]$Path = "root"
    )
    
    $count = 0
    try {
        $children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $Path -PageSize $BatchSize -All -ErrorAction SilentlyContinue
        if ($children) {
            $count += $children.Count
            # Process folders in parallel
            $folderJobs = @()
            $folders = $children | Where-Object { $_.Folder }
            
            foreach ($folder in $folders) {
                $folderJobs += Start-ThreadJob -ScriptBlock {
                    param($driveId, $folderId)
                    (Get-MgDriveItemChild -DriveId $driveId -DriveItemId $folderId -PageSize $BatchSize -All -ErrorAction SilentlyContinue).Count
                } -ArgumentList $DriveId, $folder.Id -ThrottleLimit $MaxThreads
            }
            
            # Wait for jobs and collect results
            $folderJobs | Wait-Job | ForEach-Object {
                $count += $_.Output
                Remove-Job $_
            }
        }
    }
    catch {
        Write-Host "Warning: Error counting items in $Path - $_" -ForegroundColor Yellow
    }
    return $count
}

#--- Optimized File Collection with Progress ---
function Get-FileData {
    param($Site)
    
    Write-Host "`n[+] Analyzing site structure..." -ForegroundColor Cyan
    
    # Get all drives for the site
    $drives = Get-MgSiteDrive -SiteId $Site.Id -All -ErrorAction Stop
    
    # Calculate total items with progress tracking
    Write-Host "Calculating total items for progress tracking..." -ForegroundColor Cyan
    $totalItems = 0
    $driveCount = 0
    
    foreach ($drive in $drives) {
        $driveCount++
        Show-Progress -Activity "Scanning Drives" -Status "Drive $driveCount of $($drives.Count)" `
            -PercentComplete ($driveCount/$drives.Count*100) -SecondsRemaining 30
        
        $driveItems = Get-TotalItemCount -DriveId $drive.Id
        $totalItems += $driveItems
        Write-Host "  Drive $($drive.Name) contains ~$driveItems items" -ForegroundColor Gray
    }
    
    Write-Host "Found approximately $totalItems items to process" -ForegroundColor Green
    
    # Initialize collections
    $allFiles = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $folderSizes = [System.Collections.Concurrent.ConcurrentDictionary[string,long]]::new()
    
    # Process drives in parallel
    $driveJobs = @()
    $processedItems = 0
    $startTime = Get-Date
    
    foreach ($drive in $drives) {
        $driveJobs += Start-ThreadJob -ScriptBlock {
            param($driveId, $siteId, $totalItemsRef, $processedItemsRef)
            
            try {
                $stack = [System.Collections.Stack]::new()
                $stack.Push(@{Id = "root"; Path = ""; Depth = 0})
                
                while ($stack.Count -gt 0) {
                    $current = $stack.Pop()
                    $children = $null
                    $retryCount = 0
                    
                    # Get children with retry logic
                    while ($retryCount -lt 3) {
                        try {
                            $children = Get-MgDriveItemChild -DriveId $driveId -DriveItemId $current.Id -PageSize $BatchSize -All -ErrorAction Stop
                            break
                        }
                        catch {
                            $retryCount++
                            if ($retryCount -ge 3) { throw }
                            Start-Sleep -Milliseconds (500 * $retryCount)
                        }
                    }
                    
                    foreach ($item in $children) {
                        $null = $processedItemsRef.Value++
                        $percent = [Math]::Min(100, [int](($processedItemsRef.Value/$totalItemsRef.Value)*100))
                        $elapsed = (Get-Date) - $startTime
                        $remaining = if ($percent -gt 0) { ($elapsed.TotalSeconds * (100-$percent)/$percent) } else { 0 }
                        
                        Write-Progress -Activity "Processing Items" -Status "$percent% Complete" `
                            -PercentComplete $percent -SecondsRemaining $remaining -Id 2
                        
                        if ($item.File) {
                            # Filter out system files
                            if (-not ($item.Name -match '^~|^\.|^_vti_|^appdata|^Forms$|^Thumbs\.db$|^\.DS_Store$' -or
                                      $item.WebUrl -match '/_layouts/|/_catalogs/|/_vti_bin/')) {
                                $fileObj = [PSCustomObject]@{
                                    Name = $item.Name
                                    Size = [long]$item.Size
                                    Path = $item.ParentReference.Path
                                    Drive = $driveId
                                    Extension = [System.IO.Path]::GetExtension($item.Name).ToLower()
                                }
                                $allFiles.Add($fileObj)
                                
                                # Track folder sizes
                                $folderPath = $item.ParentReference.Path
                                $folderSizes.AddOrUpdate($folderPath, $item.Size, { param($key, $value) $value + $item.Size })
                            }
                        }
                        elseif ($item.Folder -and $current.Depth -lt 10) {
                            $stack.Push(@{
                                Id = $item.Id
                                Path = "$($current.Path)/$($item.Name)"
                                Depth = $current.Depth + 1
                            })
                        }
                        
                        # Throttle requests
                        Start-Sleep -Milliseconds $ThrottleDelay
                    }
                }
            }
            catch {
                Write-Host "Error processing drive $driveId : $_" -ForegroundColor Red
            }
        } -ArgumentList $drive.Id, $Site.Id, ([ref]$totalItems), ([ref]$processedItems) -ThrottleLimit $MaxThreads
    }
    
    # Wait for all jobs to complete
    $driveJobs | Wait-Job | Out-Null
    $driveJobs | Remove-Job
    
    Write-Progress -Activity "Processing Items" -Completed -Id 2
    Write-Host "Site analysis complete - Found $($allFiles.Count) files across $($drives.Count) drives" -ForegroundColor Green
    
    return @{
        Files = $allFiles | Sort-Object Size -Descending
        FolderSizes = $folderSizes
        TotalFiles = $allFiles.Count
        TotalSizeGB = [math]::Round(($allFiles | Measure-Object -Property Size -Sum).Sum / 1GB, 2)
    }
}

#--- Main Execution ---
function Main {
    try {
        Connect-ToGraph
        $site = Get-TargetSite -Url $siteUrl
        
        # Get file data with progress tracking
        $fileData = Get-FileData -Site $site
        
        # Display summary
        Write-Host "`n[=] Audit Results Summary:" -ForegroundColor Green
        Write-Host " - Total files found: $($fileData.TotalFiles)" -ForegroundColor White
        Write-Host " - Total size: $($fileData.TotalSizeGB) GB" -ForegroundColor White
        
        # Export to Excel if requested
        $choice = Read-Host "`nExport results to Excel? (Y/N)"
        if ($choice -match '^[yY]') {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $filename = "SharePoint_Audit_$timestamp.xlsx"
            
            $fileData.Files | Select-Object Name, Size, Path, Drive, Extension |
                Export-Excel -Path $filename -AutoSize -TableStyle Medium2
                
            Write-Host "Report saved to: $filename" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "`n[!] Script failed:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Yellow
        Write-Host $_.ScriptStackTrace -ForegroundColor DarkYellow
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Progress -Activity "*" -Completed -Id 1
        Write-Progress -Activity "*" -Completed -Id 2
    }
}

# Execute the script
Main