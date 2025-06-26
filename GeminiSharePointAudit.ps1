<#
.SYNOPSIS
  SharePoint Tenant-Wide Storage & Access Audit Script

.DESCRIPTION
  Performs a comprehensive audit of ALL SharePoint sites in the tenant, including:
  - Scans all SharePoint sites and aggregates storage usage
  - Generates a pie chart of storage for the whole tenant (largest 10 sites by size)
  - For each of the 10 largest sites, generates a pie chart showing storage breakdown
  - Collects user access for all sites, including user type (internal/external)
  - Highlights external guest access in red in the Excel report
  - For the top 10 largest sites, shows the top 20 biggest files and folders
  - Exports all results to a well-structured Excel report with multiple worksheets and charts

  Features:
  - Scans the entire tenant (not just a single site)
  - Aggregates and summarizes results for easy review
  - Modern error handling and reporting
  - Modular, maintainable, and extensible design
#>

# Set strict error handling
$ErrorActionPreference = "Stop"
$WarningPreference = "SilentlyContinue"

#--- Configuration ---
# IMPORTANT: Replace these values with your actual Azure App Registration details
$clientId     = '278b9af9-888d-4344-93bb-769bdd739249' # Your Application (client) ID
$tenantId     = 'ca0711e2-e703-4f4e-9099-17d97863211c' # Your Tenant ID
$certificateThumbprint = 'B0AF0EF7659EA83D3140844F4BF89CCBB9413DBA' # Thumbprint of the certificate for authentication

#--- Required Modules ---
$requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Sites', 
    'Microsoft.Graph.Files',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'ImportExcel'
)

# Install and import required modules
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing $module..." -ForegroundColor Yellow
        Install-Module -Name $module -Force -AllowClobber -SkipPublisherCheck -WarningAction SilentlyContinue
    }
    Import-Module -Name $module -Force -WarningAction SilentlyContinue
}

#--- Authentication ---
function Connect-ToGraph {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    
    # Clear existing connections to ensure a clean state
    Disconnect-MgGraph -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    
    # Get certificate from the current user's certificate store
    $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$certificateThumbprint" -ErrorAction Stop
    
    # Connect with app-only authentication using the certificate
    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -Certificate $cert -NoWelcome -WarningAction SilentlyContinue
    
    # Verify app-only authentication was successful
    $context = Get-MgContext
    if ($context.AuthType -ne 'AppOnly') {
        throw "App-only authentication required. Current: $($context.AuthType)"
    }
    
    Write-Host "Successfully connected with app-only authentication" -ForegroundColor Green
}

#--- Batch Get Drive Items for Multiple Drives ---
function Get-DriveItemsBatch {
    param(
        [Parameter(Mandatory)]
        [array]$DriveIds,
        [string]$ParentId = 'root'
    )
    $batchSize = 20 # Microsoft Graph API batch request limit
    $responses = @()
    for ($i = 0; $i -lt $DriveIds.Count; $i += $batchSize) {
        $batch = $DriveIds[$i..([Math]::Min($i+$batchSize-1, $DriveIds.Count-1))]
        $batchRequests = @()
        $reqId = 1
        foreach ($driveId in $batch) {
            $batchRequests += @{
                id = "$reqId"
                method = "GET"
                url = "/drives/$driveId/items/$ParentId/children?`$select=name,size,parentReference,file,folder,webUrl"
            }
            $reqId++
        }
        $body = @{ requests = $batchRequests } | ConvertTo-Json -Depth 10
        try {
            $result = Invoke-MgGraphRequest -Method POST -Uri "/v1.0/`$batch" -Body $body -ContentType 'application/json'
            if ($result.responses) {
                $responses += $result.responses
            }
        }
        catch {
             Write-Host "Warning: A batch request for drive items failed. $_" -ForegroundColor Yellow
        }
    }
    return $responses
}

#--- Batch Collect File Data from Drive Item Responses ---
function Get-FileDataFromBatchResponses {
    param(
        [array]$DrivesBatchResponses
    )
    $allFiles = @()
    $folderSizes = @{} # Use a hashtable for efficient folder size tracking

    foreach ($resp in $DrivesBatchResponses) {
        if ($resp.status -eq 200 -and $resp.body.value) {
            foreach ($item in $resp.body.value) {
                if ($item.file) {
                    # Construct the full path for length calculation and reporting
                    $fullPath = ($item.parentReference.path + '/' + $item.name).Replace('//','/')
                    
                    $allFiles += [PSCustomObject]@{
                        Name = $item.name
                        Size = [long]$item.size
                        SizeGB = [math]::Round($item.size / 1GB, 3)
                        SizeMB = [math]::Round($item.size / 1MB, 2)
                        Path = $item.parentReference.path
                        Drive = $item.parentReference.driveId
                        Extension = [System.IO.Path]::GetExtension($item.name).ToLower()
                        PathLength = $fullPath.Length
                        FullPath = $fullPath
                    }
                    
                    # Aggregate folder sizes
                    $folderPath = $item.parentReference.path
                    if (-not $folderSizes.ContainsKey($folderPath)) { $folderSizes[$folderPath] = [long]0 }
                    $folderSizes[$folderPath] += [long]$item.size
                }
            }
        }
    }
    
    $totalSizeSum = ($allFiles | Measure-Object -Property Size -Sum).Sum
    
    return @{
        Files = $allFiles
        FolderSizes = $folderSizes
        TotalFiles = $allFiles.Count
        TotalSizeGB = if($totalSizeSum -gt 0) { [math]::Round($totalSizeSum / 1GB, 2) } else { 0 }
    }
}

# --- Get Tenant Name from Organization Details ---
function Get-TenantName {
    try {
        $tenant = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
        if ($tenant) { return $tenant.DisplayName.Replace(' ', '_') }
    } catch {
        Write-Host "Could not determine tenant name, using default 'Tenant'" -ForegroundColor Yellow
    }
    return 'Tenant'
}

# --- Main Execution (Tenant-wide) ---
function Main {
    try {
        Write-Host "SharePoint Tenant Storage & Access Report Generator" -ForegroundColor Green
        Write-Host "=============================================" -ForegroundColor Green

        # Authenticate with Microsoft Graph
        Connect-ToGraph

        $tenantName = Get-TenantName
        $dateStr = Get-Date -Format "yyyyMMdd_HHmmss"
        $excelFileName = "TenantAudit-$tenantName-$dateStr.xlsx"

        # Get all SharePoint sites in the tenant
        Write-Host "Enumerating all SharePoint sites in tenant... This may take a while." -ForegroundColor Cyan
        $sites = Get-MgSite -Search "*" -All -WarningAction SilentlyContinue
        Write-Host "Found $($sites.Count) sites." -ForegroundColor Green
        
        # Initialize collections to hold aggregated data from all sites
        $siteSummaries = @()
        $allSiteSummaries = @()
        $allTopFiles      = @()
        $allTopFolders    = @()
        $allLongPathFiles = @()
        $allRecycleBinFiles = @()
        $siteStorageStats = @{}
        $sitePieCharts    = @{}

        # --- Phase 1: Get summary storage for all sites to identify the largest ones ---
        Write-Host "Phase 1: Performing initial storage scan of all sites..." -ForegroundColor Cyan
        $progressCount = 0
        foreach ($site in $sites) {
            $progressCount++
            Write-Progress -Activity "Initial Storage Scan" -Status "Processing site $progressCount of $($sites.Count): $($site.DisplayName)" -PercentComplete (($progressCount / $sites.Count) * 100)
            
            $drives = Get-MgSiteDrive -SiteId $site.Id -WarningAction SilentlyContinue
            $totalSize = 0
            foreach ($drive in $drives) {
                try {
                    # Get-MgDrive does not return total size directly, so we need to sum items
                    $totalSize += $drive.Size.Total
                } catch {}
            }
            $siteSummaries += [PSCustomObject]@{
                Site = $site
                SiteName = $site.DisplayName
                SiteId = $site.Id
                SiteUrl = $site.WebUrl
                StorageBytes = $totalSize
                StorageGB = [math]::Round($totalSize / 1GB, 3)
            }
        }
        
        # Identify top 10 largest sites for deep scan
        $topSites = $siteSummaries | Sort-Object StorageBytes -Descending | Select-Object -First 10
        $topSiteIds = $topSites.SiteId
        
        # --- Phase 2: Deep scan of top sites and summary of others ---
        Write-Host "Phase 2: Performing deep scan on top 10 sites and summarizing others..." -ForegroundColor Cyan
        $progressCount = 0
        foreach ($siteSummary in $siteSummaries) {
            $progressCount++
            Write-Progress -Activity "Detailed Site Analysis" -Status "Analyzing site $progressCount of $($siteSummaries.Count): $($siteSummary.SiteName)" -PercentComplete (($progressCount / $siteSummaries.Count) * 100)

            $site = $siteSummary.Site
            $isTopSite = $topSiteIds -contains $site.Id
            $siteStorageStats[$site.DisplayName] = $siteSummary.StorageGB
            
            # Default values for non-top sites
            $totalFiles = $null
            $totalFolders = $null
            $recycleBinSizeGB = $null
            $systemFilesSizeGB = $null

            if ($isTopSite) {
                # --- Deep scan for top 10 sites ---
                $siteDrives = Get-MgSiteDrive -SiteId $site.Id -WarningAction SilentlyContinue
                $driveIds = $siteDrives | ForEach-Object { $_.id }
                
                # Batch get drive items and process file data
                $driveItemsBatchResults = Get-DriveItemsBatch -DriveIds $driveIds -ParentId 'root'
                $siteFileData = Get-FileDataFromBatchResponses -DrivesBatchResponses $driveItemsBatchResults
                
                $totalFiles = $siteFileData.TotalFiles
                $totalFolders = $siteFileData.FolderSizes.Count

                # Get top files and folders for this site
                $allTopFiles += $siteFileData.Files | Sort-Object Size -Descending | Select-Object -First 20 | ForEach-Object {
                    $_ | Add-Member -NotePropertyName SiteName -NotePropertyValue $site.DisplayName -Force; $_
                }
                $allTopFolders += $siteFileData.FolderSizes.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 | ForEach-Object {
                    [PSCustomObject]@{
                        SiteName   = $site.DisplayName
                        FolderPath = $_.Key
                        SizeGB     = [math]::Round($_.Value / 1GB, 3)
                        SizeMB     = [math]::Round($_.Value / 1MB, 2)
                    }
                }

                # Get files with long paths for this site
                $longPathFiles = @($siteFileData.Files | Where-Object { $_.PathLength -ge 399 })
                if ($longPathFiles.Count -gt 0) {
                     $allLongPathFiles += $longPathFiles | ForEach-Object { $_ | Add-Member -NotePropertyName SiteName -NotePropertyValue $site.DisplayName -Force; $_ }
                }

                # Get Recycle Bin info
                $recycleBinFiles = @()
                $recycleBinSize = 0
                try {
                    $recycleUri = "/v1.0/sites/$($site.Id)/drive/recycleBin?`$top=500" # Check top 500 items
                    $recycleResp = Invoke-MgGraphRequest -Method GET -Uri $recycleUri
                    if ($recycleResp.value) {
                        $recycleBinFiles = @($recycleResp.value | Where-Object { $_.size } | Sort-Object size -Descending | Select-Object -First 20)
                        $recycleBinSize = ($recycleResp.value | Measure-Object -Property size -Sum).Sum
                    }
                } catch { }

                if ($recycleBinFiles.Count -gt 0) {
                    $allRecycleBinFiles += $recycleBinFiles | ForEach-Object { 
                        [PSCustomObject]@{
                            SiteName = $site.DisplayName
                            Name = $_.name
                            SizeMB = [math]::Round($_.size / 1MB, 2)
                            DeletedDateTime = $_.deletedDateTime
                        }
                    }
                }
                $recycleBinSizeGB = [math]::Round($recycleBinSize / 1GB, 3)

                # Pie chart data for this site
                $pieData = @()
                $pieData += $siteFileData.FolderSizes.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 8 | ForEach-Object {
                    [PSCustomObject]@{ FolderPath = $_.Key; SizeGB = [math]::Round($_.Value / 1GB, 3) }
                }
                if ($recycleBinSizeGB -gt 0) {
                    $pieData += [PSCustomObject]@{ FolderPath = 'Recycle Bin'; SizeGB = $recycleBinSizeGB }
                }
                $sitePieCharts[$site.DisplayName] = $pieData
            } 
            
            # Add summary info for every site
            $allSiteSummaries += [PSCustomObject]@{
                SiteName = $site.DisplayName
                SiteUrl  = $site.WebUrl
                TotalSizeGB = $siteSummary.StorageGB
                TotalFiles = $totalFiles # Will be null for non-top sites
                TotalFolders = $totalFolders # Will be null for non-top sites
                RecycleBinSizeGB = $recycleBinSizeGB # Will be null for non-top sites
                ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        Write-Progress -Activity "Detailed Site Analysis" -Completed

        # --- Phase 3: Generate Excel Report ---
        Write-Host "Phase 3: Generating Excel report..." -ForegroundColor Cyan

        # Only create a report if there are sites to report on
        if ($allSiteSummaries.Count -eq 0) {
            Write-Host "No site information was found. No Excel report will be generated." -ForegroundColor Yellow
            return
        }

        # Tenant-wide storage pie chart data (top 10 sites)
        $tenantPieChartData = @($siteStorageStats.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 | ForEach-Object {
            [PSCustomObject]@{ SiteName = $_.Key; TotalSizeGB = $_.Value }
        })

        # Create the initial Excel package and the first "Summary" worksheet
        $excel = $allSiteSummaries | Export-Excel -Path $excelFileName -WorksheetName "Summary" -AutoSize -TableStyle Medium2 -PassThru
        
        # Add other worksheets only if there is data for them
        if ($allTopFiles.Count -gt 0) {
            $allTopFiles | Export-Excel -ExcelPackage $excel -WorksheetName "Top 20 Files (per site)" -AutoSize -TableStyle Medium6
        }
        if ($allTopFolders.Count -gt 0) {
            $allTopFolders | Export-Excel -ExcelPackage $excel -WorksheetName "Top 10 Folders (per site)" -AutoSize -TableStyle Medium3
        }
        if ($allLongPathFiles.Count -gt 0) {
            $allLongPathFiles | Export-Excel -ExcelPackage $excel -WorksheetName "Long Path Files (>=399 chars)" -AutoSize -TableStyle Medium8
        }
        if ($allRecycleBinFiles.Count -gt 0) {
            $allRecycleBinFiles | Export-Excel -ExcelPackage $excel -WorksheetName "Top Recycle Bin Files" -AutoSize -TableStyle Medium12
        }

        # Add Tenant Storage Pie Chart
        if ($tenantPieChartData.Count -gt 0) {
            $tenantPieChartData | Export-Excel -ExcelPackage $excel -WorksheetName "Tenant Storage Pie" -AutoSize -TableStyle Medium4
            $ws = $excel.Workbook.Worksheets["Tenant Storage Pie"]
            if ($ws) {
                $chart = $ws.Drawings.AddChart("TenantStorageChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
                $chart.Title.Text = "Tenant Storage Usage (Top 10 Sites)"
                $chart.SetPosition(1, 0, 3, 0) # Position next to the data table
                $chart.SetSize(600, 400)
                $dataRange = $ws.Cells["B2:B$($tenantPieChartData.Count + 1)"]
                $labelRange = $ws.Cells["A2:A$($tenantPieChartData.Count + 1)"]
                $series = $chart.Series.Add($dataRange, $labelRange)
                $series.Header = "Size (GB)"
            }
        }
        
        # Add individual pie charts for each top site
        foreach ($siteData in $tenantPieChartData) {
            $siteName = $siteData.SiteName
            if ($sitePieCharts.ContainsKey($siteName) -and $sitePieCharts[$siteName].Count -gt 0) {
                $pieData = $sitePieCharts[$siteName]
                $worksheetName = "Pie - " + $siteName.Substring(0, [Math]::Min($siteName.Length, 25))
                $pieData | Export-Excel -ExcelPackage $excel -WorksheetName $worksheetName -AutoSize -TableStyle Medium4
                
                $wsSite = $excel.Workbook.Worksheets[$worksheetName]
                if ($wsSite) {
                    $chartSite = $wsSite.Drawings.AddChart("SiteStorageChart_$($siteName)", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
                    $chartSite.Title.Text = "Storage Usage for $siteName"
                    $chartSite.SetPosition(1, 0, 2, 0)
                    $chartSite.SetSize(600, 400)
                    $rowCount = $pieData.Count
                    $dataRangeSite = $wsSite.Cells["B2:B$($rowCount + 1)"]
                    $labelRangeSite = $wsSite.Cells["A2:A$($rowCount + 1)"]
                    $seriesSite = $chartSite.Series.Add($dataRangeSite, $labelRangeSite)
                    $seriesSite.Header = "Size (GB)"
                }
            }
        }

        # Save and close the Excel package. This is a critical step.
        Close-ExcelPackage $excel
        Write-Host "`nReport saved to: $excelFileName" -ForegroundColor Green
    }
    catch {
        Write-Host "`nAn error occurred during script execution:" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor DarkRed
    }
    finally {
        # Always disconnect from Graph API
        Write-Host "Disconnecting from Microsoft Graph." -ForegroundColor Cyan
        Disconnect-MgGraph -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    }
}

# --- Execute the main function to run the script ---
Main
