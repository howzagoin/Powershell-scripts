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
$clientId     = '278b9af9-888d-4344-93bb-769bdd739249'
$tenantId     = 'ca0711e2-e703-4f4e-9099-17d97863211c'
$certificateThumbprint = 'B0AF0EF7659EA83D3140844F4BF89CCBB9413DBA'

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
    
    # Clear existing connections
    Disconnect-MgGraph -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    
    # Connect with app-only authentication using the certificate thumbprint directly
    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $certificateThumbprint -NoWelcome -WarningAction SilentlyContinue
    
    # Verify app-only authentication
    $context = Get-MgContext
    if ($context.AuthType -ne 'AppOnly') {
        throw "App-only authentication required. Current: $($context.AuthType)"
    }
    
    Write-Host "Successfully connected with app-only authentication" -ForegroundColor Green
}

# --- Get All SharePoint Sites in Tenant ---
function Get-AllSharePointSites {
    Write-Host "Enumerating all SharePoint sites in tenant..." -ForegroundColor Cyan
    $sites = Get-MgSite -All -Property "id,displayName,webUrl,siteCollection" -WarningAction SilentlyContinue
    Write-Host "Found $($sites.Count) sites." -ForegroundColor Green
    return $sites
}

# --- Get Tenant Name ---
function Get-TenantName {
    $tenant = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($tenant) { return $tenant.DisplayName.Replace(' ', '_') }
    return 'Tenant'
}

# --- Batch Get Drive Items for Multiple Drives ---
function Get-DriveItemsBatch {
    param(
        [Parameter(Mandatory)]
        [array]$DriveIds,
        [string]$ParentId = 'root'
    )
    $batchSize = 20
    $responses = @()
    for ($i = 0; $i -lt $DriveIds.Count; $i += $batchSize) {
        $batch = $DriveIds[$i..([Math]::Min($i+$batchSize-1, $DriveIds.Count-1))]
        $batchRequests = @()
        $reqId = 1
        foreach ($driveId in $batch) {
            $batchRequests += @{
                id = "$reqId"
                method = "GET"
                url = "/drives/$driveId/items/$ParentId/children"
            }
            $reqId++
        }
        $body = @{ requests = $batchRequests } | ConvertTo-Json -Depth 10
        $result = Invoke-MgGraphRequest -Method POST -Uri "/v1.0/`$batch" -Body $body -ContentType 'application/json'
        if ($result.responses) {
            $responses += $result.responses
        }
    }
    return $responses
}

# --- Batch Collect File Data for Drives ---
function Get-FileDataBatch {
    param(
        [array]$DrivesBatchResponses
    )
    $allFiles = @()
    $folderSizes = @{}
    
    # Helper function to recursively process items
    function Process-Items {
        param($items)
        foreach ($item in $items) {
            if ($item.file) {
                # Calculate full path+name length (Path + '/' + Name)
                $fullPath = ($item.parentReference.path + '/' + $item.name).Replace('//','/')
                $pathLength = $fullPath.Length
                $allFiles += [PSCustomObject]@{ 
                    Name = $item.name
                    Size = [long]$item.size
                    SizeGB = [math]::Round($item.size / 1GB, 3)
                    SizeMB = [math]::Round($item.size / 1MB, 2)
                    Path = $item.parentReference.path
                    Drive = $item.parentReference.driveId
                    Extension = [System.IO.Path]::GetExtension($item.name).ToLower()
                    PathLength = $pathLength
                    FullPath = $fullPath
                }
                $folderPath = $item.parentReference.path
                if (-not $folderSizes.ContainsKey($folderPath)) { $folderSizes[$folderPath] = 0 }
                $folderSizes[$folderPath] += [long]$item.size
            }
            # If the item is a folder and has children, process them
            if ($item.folder -and $item.children) {
                Process-Items -items $item.children
            }
        }
    }

    foreach ($resp in $DrivesBatchResponses) {
        if ($resp.status -eq 200 -and $resp.body.value) {
            Process-Items -items $resp.body.value
        }
    }
    return @{
        Files = $allFiles
        FolderSizes = $folderSizes
        TotalFiles = $allFiles.Count
        TotalSizeGB = [math]::Round(($allFiles | Measure-Object -Property Size -Sum).Sum / 1GB, 2)
    }
}

# --- Main Execution (Tenant-wide) ---
function Main {
    try {
        Write-Host "SharePoint Tenant Storage & Access Report Generator" -ForegroundColor Green
        Write-Host "=============================================" -ForegroundColor Green

        # Connect to Microsoft Graph
        Connect-ToGraph

        $tenantName = Get-TenantName
        $dateStr = Get-Date -Format yyyyMMdd_HHmmss
        $outputDir = "C:\Users\Howza Goin\GitHubRepositories\Powershell-scripts\output"
        if (-not (Test-Path $outputDir)) {
            New-Item -Path $outputDir -ItemType Directory | Out-Null
        }
        $excelFileName = Join-Path -Path $outputDir -ChildPath "TenantAudit-$tenantName-$dateStr.xlsx"

        # Get all SharePoint sites in the tenant
        $sites = Get-AllSharePointSites

        # First, get summary storage for all sites accurately
        Write-Host "Calculating storage for all $($sites.Count) sites... (This may take a while)" -ForegroundColor Cyan
        $siteSummaries = @()
        $progressCount = 0
        foreach ($site in $sites) {
            $progressCount++
            Write-Progress -Activity "Gathering Site Storage" -Status "Processing site $progressCount of $($sites.Count): $($site.DisplayName)" -PercentComplete (($progressCount / $sites.Count) * 100)
            $drives = Get-MgSiteDrive -SiteId $site.Id -WarningAction SilentlyContinue
            $totalUsedStorage = ($drives | Measure-Object -Property { $_.Quota.Used } -Sum).Sum
            
            $siteSummaries += [PSCustomObject]@{ 
                Site = $site
                SiteName = $site.DisplayName
                SiteId = $site.Id
                SiteUrl = $site.WebUrl
                StorageBytes = $totalUsedStorage
                StorageGB = [math]::Round($totalUsedStorage / 1GB, 3)
            }
        }
        Write-Progress -Activity "Gathering Site Storage" -Completed

        # Identify top 10 largest sites
        $topSites = $siteSummaries | Sort-Object StorageBytes -Descending | Select-Object -First 10
        $topSiteIds = $topSites.SiteId
        
        $allSiteSummaries = @()
        $allTopFiles      = @()
        $allTopFolders    = @()
        $allLongPathFiles = @()
        $allRecycleBinFiles = @()
        $siteStorageStats = @{}
        $sitePieCharts    = @{}

        Write-Host "Performing deep scan on the 10 largest sites..." -ForegroundColor Cyan
        $progressCount = 0
        foreach ($siteSummary in $topSites) {
            $progressCount++
            $site = $siteSummary.Site
            Write-Progress -Activity "Deep Scanning Top 10 Sites" -Status "Scanning $($site.DisplayName) ($progressCount of 10)" -PercentComplete ($progressCount / 10 * 100)
            
            $siteStorageStats[$site.DisplayName] = $siteSummary.StorageGB

            # --- Deep scan for top 10 sites ---
            $siteDrives = Get-MgSiteDrive -SiteId $site.Id -WarningAction SilentlyContinue
            $driveIds = $siteDrives | ForEach-Object { $_.id }
            
            $driveItemsBatchResults = Get-DriveItemsBatch -DriveIds $driveIds -ParentId 'root'
            $siteFileData = Get-FileDataBatch -DrivesBatchResponses $driveItemsBatchResults

            # --- Recycle Bin: Get top 10 largest files and total size ---
            $recycleBinSize = 0
            try {
                $recycleUri = "/v1.0/sites/$($site.Id)/drive/recycleBin/items?`$top=500"
                $recycleResp = Invoke-MgGraphRequest -Method GET -Uri $recycleUri
                if ($recycleResp.value) {
                    $recycledItems = $recycleResp.value | Where-Object { $_.size }
                    $allRecycleBinFiles += $recycledItems | Sort-Object size -Descending | Select-Object -First 10 |
                        ForEach-Object {
                            [PSCustomObject]@{ 
                                SiteName = $site.DisplayName
                                Name = $_.name
                                SizeMB = [math]::Round($_.size / 1MB, 2)
                                SizeGB = [math]::Round($_.size / 1GB, 3)
                                DeletedDateTime = $_.deletedDateTime
                            }
                        }
                    $recycleBinSize = ($recycledItems | Measure-Object -Property size -Sum).Sum
                }
            } catch {
                Write-Warning "Could not retrieve recycle bin for site $($site.DisplayName). Error: $($_.Exception.Message)"
                $recycleBinSize = 0
            }
            
            # System files (filtered out in main logic, but count for pie chart)
            $systemFiles = $siteFileData.Files | Where-Object {
                $_.Name -match '^~' -or $_.Name -match '^\.' -or $_.Name -match '^Forms$' -or 
                $_.Name -match '^_vti_' -or $_.Name -match '^appdata' -or $_.Name -match '^\.DS_Store$' -or 
                $_.Name -match '^Thumbs\.db$'
            }
            $systemFilesSize = ($systemFiles | Measure-Object -Property Size -Sum).Sum
            
            $allSiteSummaries += [PSCustomObject]@{ 
                SiteName = $site.DisplayName
                SiteUrl  = $site.WebUrl
                TotalFiles = $siteFileData.TotalFiles
                TotalSizeGB = $siteFileData.TotalSizeGB
                TotalFolders = $siteFileData.FolderSizes.Count
                RecycleBinSizeGB = [math]::Round($recycleBinSize / 1GB, 3)
                SystemFilesSizeGB = [math]::Round($systemFilesSize / 1GB, 3)
                ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            
            # Add SiteName to each file object for later grouping
            $siteFileData.Files | ForEach-Object { $_ | Add-Member -NotePropertyName SiteName -NotePropertyValue $site.DisplayName -Force }
            $allTopFiles += $siteFileData.Files | Sort-Object Size -Descending | Select-Object -First 20
            
            $allLongPathFiles += $siteFileData.Files | Where-Object { $_.PathLength -ge 399 }

            $allTopFolders += $siteFileData.FolderSizes.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 | ForEach-Object {
                [PSCustomObject]@{ 
                    SiteName   = $site.DisplayName
                    FolderPath = $_.Key
                    SizeGB     = [math]::Round($_.Value / 1GB, 3)
                    SizeMB     = [math]::Round($_.Value / 1MB, 2)
                }
            }
            
            # Pie chart breakdown for this site
            $pieData = @()
            $pieData += $siteFileData.FolderSizes.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 8 | ForEach-Object {
                [PSCustomObject]@{ 
                    FolderPath = $_.Key
                    SizeGB = [math]::Round($_.Value / 1GB, 3)
                    Category = 'Folder'
                }
            }
            if ($systemFilesSize -gt 0) {
                $pieData += [PSCustomObject]@{ FolderPath = 'System Files'; SizeGB = [math]::Round($systemFilesSize / 1GB, 3); Category = 'System' }
            }
            if ($recycleBinSize -gt 0) {
                $pieData += [PSCustomObject]@{ FolderPath = 'Recycle Bin'; SizeGB = [math]::Round($recycleBinSize / 1GB, 3); Category = 'RecycleBin' }
            }
            $sitePieCharts[$site.DisplayName] = $pieData
        }
        Write-Progress -Activity "Deep Scanning Top 10 Sites" -Completed

        # Add non-top sites to the summary list
        $otherSites = $siteSummaries | Where-Object { $topSiteIds -notcontains $_.SiteId }
        foreach($siteSummary in $otherSites){
             $allSiteSummaries += [PSCustomObject]@{ 
                SiteName = $siteSummary.SiteName
                SiteUrl  = $siteSummary.SiteUrl
                TotalFiles = "N/A (Not a top 10 site)"
                TotalSizeGB = $siteSummary.StorageGB
                TotalFolders = "N/A"
                RecycleBinSizeGB = "N/A"
                SystemFilesSizeGB = "N/A"
                ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }

        # Tenant-wide storage pie chart (top 10 sites)
        $tenantPieChartData = $topSites | Select-Object SiteName, StorageGB

        # --- Create Excel Report ---
        Write-Host "Creating Excel report: $excelFileName" -ForegroundColor Cyan
        
        $excelParams = @{
            Path = $excelFileName
            WorksheetName = "Summary"
            AutoSize = $true
            TableStyle = "Medium2"
            PassThru = $true
        }
        $excel = $allSiteSummaries | Sort-Object StorageGB -Descending | Export-Excel @excelParams
        
        $allTopFiles | Select-Object SiteName, Name, SizeMB, Path, Extension | Export-Excel -ExcelPackage $excel -WorksheetName "Top 20 Files (per site)" -AutoSize -TableStyle Medium6
        
        if ($allTopFolders.Count -gt 0) {
            $allTopFolders | Export-Excel -ExcelPackage $excel -WorksheetName "Top 10 Folders (per site)" -AutoSize -TableStyle Medium3
        }

        if ($allLongPathFiles.Count -gt 0) {
            $allLongPathFiles | Select-Object SiteName, FullPath, PathLength | Export-Excel -ExcelPackage $excel -WorksheetName "Long Path Files (>=399)" -AutoSize -TableStyle Medium8
        }

        if ($allRecycleBinFiles.Count -gt 0) {
            $allRecycleBinFiles | Export-Excel -ExcelPackage $excel -WorksheetName "Recycle Bin (Top 10 Largest)" -AutoSize -TableStyle Medium12
        }

        # Add tenant-wide pie chart
        $wsTenantPie = $tenantPieChartData | Export-Excel -ExcelPackage $excel -WorksheetName "Tenant Storage Pie" -AutoSize -TableStyle Medium4 -PassThru
        $chartTenant = $wsTenantPie.Drawings.AddChart("TenantStorageChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
        $chartTenant.Title.Text = "Tenant Storage Usage (Top 10 Sites)"
        $chartTenant.SetPosition(1, 0, 3, 0) # row, rowoffset, col, coloffset
        $chartTenant.SetSize(600, 400)
        $seriesTenant = $chartTenant.Series.Add($wsTenantPie.Cells["B2:B$($tenantPieChartData.Count + 1)"], $wsTenantPie.Cells["A2:A$($tenantPieChartData.Count + 1)"])
        $seriesTenant.Header = "Size (GB)"

        # Add site-specific pie charts
        foreach ($site in $topSites) {
            $siteName = $site.SiteName
            if ($sitePieCharts.ContainsKey($siteName) -and $sitePieCharts[$siteName]) {
                $pieData = $sitePieCharts[$siteName]
                if($pieData.Count -gt 0){
                    $wsSitePie = $pieData | Export-Excel -ExcelPackage $excel -WorksheetName ("Pie - " + $siteName.Substring(0, [Math]::Min($siteName.Length, 25))) -AutoSize -TableStyle Medium4 -PassThru
                    $chartSite = $wsSitePie.Drawings.AddChart("SiteStorageChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
                    $chartSite.Title.Text = "Storage Usage for $($siteName)"
                    $chartSite.SetPosition(1, 0, 4, 0)
                    $chartSite.SetSize(500, 400)
                    $rowCount = $pieData.Count
                    if ($rowCount -gt 0) {
                        $chartSite.Series.Add($wsSitePie.Cells["B2:B$($rowCount + 1)"], $wsSitePie.Cells["A2:A$($rowCount + 1)"])
                    }
                }
            }
        }
        
        Close-ExcelPackage $excel
        Write-Host "`nReport saved to: $excelFileName" -ForegroundColor Green

    }
    catch {
        Write-Host "`nError: $_" -ForegroundColor Red
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor DarkRed
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    }
}

# Execute the script
Main
