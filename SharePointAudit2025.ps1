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
# $siteUrl      = 'https://fbaint.sharepoint.com/sites/Marketing'  # No longer needed, script scans all sites
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
    
    # Get certificate
    $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$certificateThumbprint" -ErrorAction Stop
    
    # Connect with app-only authentication
    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -Certificate $cert -NoWelcome -WarningAction SilentlyContinue
    
    # Verify app-only authentication
    $context = Get-MgContext
    if ($context.AuthType -ne 'AppOnly') {
        throw "App-only authentication required. Current: $($context.AuthType)"
    }
    
    Write-Host "Successfully connected with app-only authentication" -ForegroundColor Green
}

#--- Get Site Information ---
function Get-SiteInfo {
    param([string]$SiteUrl)
    
    Write-Host "Getting site information..." -ForegroundColor Cyan
    
    # Extract site ID from URL
    $uri = [Uri]$SiteUrl
    $sitePath = $uri.AbsolutePath
    $siteId = "$($uri.Host):$sitePath"
    
    $site = Get-MgSite -SiteId $siteId
    Write-Host "Found site: $($site.DisplayName)" -ForegroundColor Green
    
    return $site
}

#--- Get Total Item Count (for progress bar) ---
function Get-TotalItemCount {
    param(
        [string]$DriveId,
        [string]$Path = "root"
    )
    
    $count = 0
    try {
        $children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $Path -All -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        if ($children) {
            $count += $children.Count
            foreach ($child in $children) {
                if ($child.Folder) {
                    $count += Get-TotalItemCount -DriveId $DriveId -Path $child.Id
                }
            }
        }
    }
    catch {
        # Silently handle errors
    }
    return $count
}

#--- Get Drive Items Recursively ---
function Get-DriveItems {
    param(
        [string]$DriveId,
        [string]$Path,
        [int]$Depth = 0,
        [ref]$GlobalItemIndex,
        [int]$TotalItems = 1
    )
    $items = @()
    try {
        $children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $Path -All -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        
        if ($children -and $children.Count -gt 0) {
            foreach ($child in $children) {
                if ($null -eq $child -or -not $child.Id) { continue }
                
                $items += $child
                $GlobalItemIndex.Value++
                
                # Update progress bar
                $percent = if ($TotalItems -gt 0) { [Math]::Min(100, [int](($GlobalItemIndex.Value/$TotalItems)*100)) } else { 100 }
                $progressBar = ('█' * ($percent / 2)) + ('░' * (50 - ($percent / 2)))
                Write-Progress -Activity "Scanning SharePoint Site Content" -Status "[$progressBar] $percent% - Processing: $($child.Name)" -PercentComplete $percent -Id 1
                
                # Recursively get folder contents
                if ($child.Folder -and $Depth -lt 10) {
                    try {
                        $items += Get-DriveItems -DriveId $DriveId -Path $child.Id -Depth ($Depth + 1) -GlobalItemIndex $GlobalItemIndex -TotalItems $TotalItems
                    } catch {
                        # Silently skip folders with access issues
                    }
                }
            }
        }
    }
    catch {
        # Silently handle errors
    }
    return $items
}

#--- Collect File Data ---
function Get-FileData {
    param($Site)
    
    Write-Host "Analyzing site structure..." -ForegroundColor Cyan
    
    $allFiles = @()
    $folderSizes = @{}
    
    # Get all drives for the site
    $drives = Get-MgSiteDrive -SiteId $Site.Id -WarningAction SilentlyContinue
    
    # Calculate total items for progress bar
    Write-Host "Calculating total items for progress tracking..." -ForegroundColor Cyan
    $totalItems = 0
    foreach ($drive in $drives) {
        $totalItems += Get-TotalItemCount -DriveId $drive.Id
    }
    
    Write-Host "Found approximately $totalItems items to process" -ForegroundColor Green
    
    $globalItemIndex = 0
    $globalItemIndexRef = [ref]$globalItemIndex
    $items = @()
    
    # Process each drive
    foreach ($drive in $drives) {
        try {
            $children = Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId "root" -All -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            
            foreach ($child in $children) {
                if ($null -eq $child) { continue }
                
                $items += $child
                $globalItemIndexRef.Value++
                
                # Update progress bar
                $percent = if ($totalItems -gt 0) { [Math]::Min(100, [int](($globalItemIndexRef.Value/$totalItems)*100)) } else { 100 }
                $progressBar = ('█' * ($percent / 2)) + ('░' * (50 - ($percent / 2)))
                Write-Progress -Activity "Scanning SharePoint Site Content" -Status "[$progressBar] $percent% - Processing: $($child.Name)" -PercentComplete $percent -Id 1
                
                if ($child.Folder) {
                    $items += Get-DriveItems -DriveId $drive.Id -Path $child.Id -Depth 1 -GlobalItemIndex $globalItemIndexRef -TotalItems $totalItems
                }
            }
        }
        catch {
            # Silently handle drive access errors
        }
    }
    
    Write-Progress -Activity "Scanning SharePoint Site Content" -Completed -Id 1
    Write-Host "Processing collected items..." -ForegroundColor Cyan
    
    foreach ($item in $items) {
        if ($item.File) {
            # Filter out system files
            $isSystem = $false
            if ($item.Name -match '^~' -or $item.Name -match '^\.' -or $item.Name -match '^Forms$' -or 
                $item.Name -match '^_vti_' -or $item.Name -match '^appdata' -or $item.Name -match '^.DS_Store$' -or 
                $item.Name -match '^Thumbs.db$') {
                $isSystem = $true
            }
            if ($item.File.MimeType -eq 'application/vnd.microsoft.sharepoint.system' -or 
                $item.File.MimeType -eq 'application/vnd.ms-sharepoint.folder') {
                $isSystem = $true
            }
            if ($item.WebUrl -match '/_layouts/' -or $item.WebUrl -match '/_catalogs/' -or 
                $item.WebUrl -match '/_vti_bin/') {
                $isSystem = $true
            }
            if ($isSystem) { continue }
            
            $allFiles += [PSCustomObject]@{
                Name = $item.Name
                Size = [long]$item.Size
                SizeGB = [math]::Round($item.Size / 1GB, 3)
                SizeMB = [math]::Round($item.Size / 1MB, 2)
                Path = $item.ParentReference.Path
                Drive = $item.ParentReference.DriveId
                Extension = [System.IO.Path]::GetExtension($item.Name).ToLower()
            }
            
            # Track folder sizes
            $folderPath = $item.ParentReference.Path
            if (-not $folderSizes.ContainsKey($folderPath)) {
                $folderSizes[$folderPath] = 0
            }
            $folderSizes[$folderPath] += $item.Size
        }
    }
    
    Write-Host "Site analysis complete - Found $($allFiles.Count) files across $($drives.Count) drives" -ForegroundColor Green
    
    return @{
        Files = $allFiles
        FolderSizes = $folderSizes
        TotalFiles = $allFiles.Count
        TotalSizeGB = [math]::Round(($allFiles | Measure-Object -Property Size -Sum).Sum / 1GB, 2)
    }
}

#--- Get Parent Folder Access Information ---
function Get-ParentFolderAccess {
    param($Site)
    
    Write-Host "Retrieving parent folder access information..." -ForegroundColor Cyan
    
    $folderAccess = @()
    $processedFolders = @{}
    
    try {
        # Get all drives for the site
        $drives = Get-MgSiteDrive -SiteId $Site.Id -WarningAction SilentlyContinue
        
        foreach ($drive in $drives) {
            try {
                # Get root folders only (first level)
                $rootFolders = Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId "root" -All -ErrorAction Stop | 
                              Where-Object { $_.Folder }
                
                foreach ($folder in $rootFolders) {
                    if ($processedFolders.ContainsKey($folder.Id)) { continue }
                    $processedFolders[$folder.Id] = $true
                    
                    try {
                        $permissions = Get-MgDriveItemPermission -DriveId $drive.Id -DriveItemId $folder.Id -All -ErrorAction Stop
                        
                        foreach ($perm in $permissions) {
                            $roles = ($perm.Roles | Where-Object { $_ }) -join ', '
                            
                            if ($perm.GrantedToIdentitiesV2) {
                                foreach ($identity in $perm.GrantedToIdentitiesV2) {
                                    if ($identity.User.DisplayName) {
                                        $folderAccess += [PSCustomObject]@{
                                            FolderName = $folder.Name
                                            FolderPath = $folder.ParentReference.Path + '/' + $folder.Name
                                            UserName = $identity.User.DisplayName
                                            UserEmail = $identity.User.Email
                                            PermissionLevel = $roles
                                            AccessType = if ($roles -match 'owner|write') { 'Full/Edit' } 
                                                        elseif ($roles -match 'read') { 'Read Only' } 
                                                        else { 'Other' }
                                        }
                                    }
                                }
                            }
                            
                            if ($perm.GrantedTo -and $perm.GrantedTo.User.DisplayName) {
                                $folderAccess += [PSCustomObject]@{
                                    FolderName = $folder.Name
                                    FolderPath = $folder.ParentReference.Path + '/' + $folder.Name
                                    UserName = $perm.GrantedTo.User.DisplayName
                                    UserEmail = $perm.GrantedTo.User.Email
                                    PermissionLevel = $roles
                                    AccessType = if ($roles -match 'owner|write') { 'Full/Edit' } 
                                                elseif ($roles -match 'read') { 'Read Only' } 
                                                else { 'Other' }
                                }
                            }
                        }
                    }
                    catch {
                        Write-Host "Warning: Could not retrieve permissions for folder $($folder.Name) - $_" -ForegroundColor Yellow
                    }
                }
            }
            catch {
                Write-Host "Warning: Could not access drive $($drive.Id) - $_" -ForegroundColor Yellow
            }
        }
        
        # Remove duplicates (same user with same access to same folder)
        $folderAccess = $folderAccess | Sort-Object FolderName, UserName, PermissionLevel -Unique
        
        Write-Host "Found access data for $($folderAccess.Count) parent folder permissions" -ForegroundColor Green
        
    }
    catch {
        Write-Host "Error retrieving folder access data: $_" -ForegroundColor Red
        $folderAccess = @([PSCustomObject]@{
            FolderName = "Permission Error"
            FolderPath = "Check permissions"
            UserName = "Unable to retrieve data"
            UserEmail = "Requires additional permissions"
            PermissionLevel = "N/A"
            AccessType = "Error"
        })
    }
    
    return $folderAccess
}

#--- Create Excel Report ---
function New-ExcelReport {
    param(
        $FileData,
        $FolderAccess,
        $Site,
        $FileName
    )
    
    Write-Host "Creating Excel report: $FileName" -ForegroundColor Cyan
    
    # Prepare data for different sheets
    $top20Files = $FileData.Files | Sort-Object Size -Descending | Select-Object -First 20 |
        Select-Object Name, SizeMB, Path, Drive, Extension
    
    $top10Folders = $FileData.FolderSizes.GetEnumerator() | 
        Sort-Object Value -Descending | Select-Object -First 10 |
        ForEach-Object { 
            [PSCustomObject]@{
                FolderPath = $_.Key
                SizeGB = [math]::Round($_.Value / 1GB, 3)
                SizeMB = [math]::Round($_.Value / 1MB, 2)
            }
        }
    
    # Storage breakdown by location for pie chart
    $storageBreakdown = $FileData.FolderSizes.GetEnumerator() | 
        Sort-Object Value -Descending | Select-Object -First 15 |
        ForEach-Object {
            $folderName = if ($_.Key -match '/([^/]+)/?$') { $matches[1] } else { "Root" }
            [PSCustomObject]@{
                Location = $folderName
                Path = $_.Key
                SizeGB = [math]::Round($_.Value / 1GB, 3)
                SizeMB = [math]::Round($_.Value / 1MB, 2)
                Percentage = [math]::Round(($_.Value / ($FileData.Files | Measure-Object Size -Sum).Sum) * 100, 1)
            }
        }
    
    # Parent folder access summary
    $accessSummary = $FolderAccess | Group-Object PermissionLevel | 
        ForEach-Object {
            [PSCustomObject]@{
                PermissionLevel = $_.Name
                UserCount = $_.Count
                Users = ($_.Group.UserName | Sort-Object -Unique) -join '; '
            }
        }
    
    # Site summary
    $siteSummary = @([PSCustomObject]@{
        SiteName = $Site.DisplayName
        SiteUrl = $Site.WebUrl
        TotalFiles = $FileData.TotalFiles
        TotalSizeGB = $FileData.TotalSizeGB
        TotalFolders = $FileData.FolderSizes.Count
        UniquePermissionLevels = ($FolderAccess.PermissionLevel | Sort-Object -Unique).Count
        ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    })
    
    # Create Excel file with multiple worksheets
    $excel = $siteSummary | Export-Excel -Path $FileName -WorksheetName "Summary" -AutoSize -TableStyle Medium2 -PassThru
    $top20Files | Export-Excel -ExcelPackage $excel -WorksheetName "Top 20 Files" -AutoSize -TableStyle Medium6
    $top10Folders | Export-Excel -ExcelPackage $excel -WorksheetName "Top 10 Folders" -AutoSize -TableStyle Medium3
    $storageBreakdown | Export-Excel -ExcelPackage $excel -WorksheetName "Storage Breakdown" -AutoSize -TableStyle Medium4
    $FolderAccess | Export-Excel -ExcelPackage $excel -WorksheetName "Folder Access" -AutoSize -TableStyle Medium5
    $accessSummary | Export-Excel -ExcelPackage $excel -WorksheetName "Access Summary" -AutoSize -TableStyle Medium1
    
    # Add charts to the storage breakdown worksheet
    $ws = $excel.Workbook.Worksheets["Storage Breakdown"]
    
    # Create pie chart for storage distribution by location
    $chart = $ws.Drawings.AddChart("StorageChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
    $chart.Title.Text = "Storage Usage by Location"
    $chart.SetPosition(1, 0, 7, 0)
    $chart.SetSize(500, 400)
    
    $series = $chart.Series.Add($ws.Cells["C2:C$($storageBreakdown.Count + 1)"], $ws.Cells["A2:A$($storageBreakdown.Count + 1)"])
    $series.Header = "Size (GB)"
    
    Close-ExcelPackage $excel
    
    Write-Host "Excel report created successfully!" -ForegroundColor Green
    Write-Host "`nReport Contents:" -ForegroundColor Cyan
    Write-Host "- Summary: Overall site statistics" -ForegroundColor White
    Write-Host "- Top 20 Files: Largest files by size" -ForegroundColor White  
    Write-Host "- Top 10 Folders: Largest folders by size" -ForegroundColor White
    Write-Host "- Storage Breakdown: Space usage by location with pie chart" -ForegroundColor White
    Write-Host "- Folder Access: Parent folder permissions" -ForegroundColor White
    Write-Host "- Access Summary: Users grouped by permission level" -ForegroundColor White
}

# --- Get All SharePoint Sites in Tenant ---
function Get-AllSharePointSites {
    Write-Host "Enumerating all SharePoint sites in tenant..." -ForegroundColor Cyan
    $sites = Get-MgSite -Search "*" -All -WarningAction SilentlyContinue
    return $sites
}

# --- Get Site Storage and User Access Info ---
function Get-SiteStorageAndAccess {
    param($Site)
    $siteInfo = @{}
    $siteInfo.SiteName = $Site.DisplayName
    $siteInfo.SiteUrl = $Site.WebUrl
    $siteInfo.SiteId = $Site.Id
    $siteInfo.StorageGB = 0
    $siteInfo.Users = @()
    $siteInfo.ExternalGuests = @()
    $siteInfo.TopFiles = @()
    $siteInfo.TopFolders = @()

    # Get drives and storage
    $drives = Get-MgSiteDrive -SiteId $Site.Id -WarningAction SilentlyContinue
    $allFiles = @()
    $folderSizes = @{
    }
    foreach ($drive in $drives) {
        try {
            $items = Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId "root" -All -ErrorAction SilentlyContinue
            foreach ($item in $items) {
                if ($item.File) {
                    $allFiles += $item
                    $folderPath = $item.ParentReference.Path
                    if (-not $folderSizes.ContainsKey($folderPath)) { $folderSizes[$folderPath] = 0 }
                    $folderSizes[$folderPath] += $item.Size
                }
            }
        } catch {
        }
    }
    $siteInfo.StorageGB = [math]::Round(($allFiles | Measure-Object -Property Size -Sum).Sum / 1GB, 2)
    $siteInfo.TopFiles = $allFiles | Sort-Object Size -Descending | Select-Object -First 20 | ForEach-Object {
        [PSCustomObject]@{
            Name = $_.Name
            SizeMB = [math]::Round($_.Size / 1MB, 2)
            Path = $_.ParentReference.Path
            Extension = [System.IO.Path]::GetExtension($_.Name).ToLower()
        }
    }
    $siteInfo.TopFolders = $folderSizes.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 20 | ForEach-Object {
        [PSCustomObject]@{
            FolderPath = $_.Key
            SizeGB = [math]::Round($_.Value / 1GB, 3)
            SizeMB = [math]::Round($_.Value / 1MB, 2)
        }
    }

    # Get user access (site permissions)
    $siteUsers = @()
    $externalGuests = @()
    try {
        $permissions = Get-MgSitePermission -SiteId $Site.Id -All -ErrorAction SilentlyContinue
        foreach ($perm in $permissions) {
            if ($perm.Invitation) {
                $userType = 'External Guest'
                $externalGuests += [PSCustomObject]@{
                    UserName = $perm.Invitation.InvitedUserDisplayName
                    UserEmail = $perm.Invitation.InvitedUserEmailAddress
                    AccessType = $perm.Roles -join ', '
                }
            } elseif ($perm.GrantedToIdentitiesV2) {
                foreach ($identity in $perm.GrantedToIdentitiesV2) {
                    $userType = if ($identity.User.UserType -eq 'Guest') { 'External Guest' } elseif ($identity.User.UserType -eq 'Member') { 'Internal' } else { $identity.User.UserType }
                    $userObj = [PSCustomObject]@{
                        UserName = $identity.User.DisplayName
                        UserEmail = $identity.User.Email
                        UserType = $userType
                        AccessType = $perm.Roles -join ', '
                    }
                    $siteUsers += $userObj
                    if ($userType -eq 'External Guest') { $externalGuests += $userObj }
                }
            }
        }
    } catch {
    }
    $siteInfo.Users = $siteUsers
    $siteInfo.ExternalGuests = $externalGuests
    return $siteInfo
}

# --- Get Tenant Name ---
function Get-TenantName {
    $tenant = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($tenant) { return $tenant.DisplayName.Replace(' ', '_') }
    return 'Tenant'
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
        $excelFileName = "TenantAudit-$tenantName-$dateStr.xlsx"

        # Get all SharePoint sites in the tenant
        $sites = Get-AllSharePointSites
        $siteIds = $sites | ForEach-Object { $_.Id }

        # First, get summary storage for all sites
        $siteSummaries = @()
        foreach ($site in $sites) {
            $drives = Get-MgSiteDrive -SiteId $site.Id -WarningAction SilentlyContinue
            $totalSize = 0
            foreach ($drive in $drives) {
                try {
                    $items = Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId "root" -All -ErrorAction SilentlyContinue
                    $totalSize += ($items | Measure-Object -Property Size -Sum).Sum
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
        # Identify top 10 largest sites
        $topSites = $siteSummaries | Sort-Object StorageBytes -Descending | Select-Object -First 10
        $topSiteIds = $topSites.SiteId
        $allSiteSummaries = @()
        $allTopFiles      = @()
        $allTopFolders    = @()
        $siteStorageStats = @{
        }
        $sitePieCharts    = @{
        }
        $detailedSites = @()
        foreach ($siteSummary in $siteSummaries) {
            $site = $siteSummary.Site
            $isTop = $topSiteIds -contains $site.Id
            $siteStorageStats[$site.DisplayName] = $siteSummary.StorageGB
            if ($isTop) {
                # --- Deep scan for top 10 sites ---
                $siteDrives = Get-MgSiteDrive -SiteId $site.Id -WarningAction SilentlyContinue
                $driveIds = $siteDrives | ForEach-Object { $_.id }
                $driveItemsBatchResults = Get-DriveItemsBatch -DriveIds $driveIds -ParentId 'root'
                $siteFileData = Get-FileDataBatch -DrivesBatchResponses $driveItemsBatchResults
                # --- Recycle Bin: Get top 10 largest files and total size ---
                $recycleBinFiles = @()
                $recycleBinSize = 0
                try {
                    $recycleUri = "/v1.0/sites/$($site.Id)/drive/recycleBin?\$top=500"
                    $recycleResp = Invoke-MgGraphRequest -Method GET -Uri $recycleUri
                    if ($recycleResp.value) {
                        $recycleBinFiles = $recycleResp.value | Where-Object { $_.size } | Sort-Object size -Descending | Select-Object -First 10 |
                            ForEach-Object {
                                [PSCustomObject]@{
                                    Name = $_.name
                                    SizeMB = [math]::Round($_.size / 1MB, 2)
                                    SizeGB = [math]::Round($_.size / 1GB, 3)
                                    DeletedDateTime = $_.deletedDateTime
                                }
                            }
                        $recycleBinSize = ($recycleResp.value | Measure-Object -Property size -Sum).Sum
                    }
                } catch {
                    $recycleBinFiles = @()
                    $recycleBinSize = 0
                }
                # System files (filtered out in main logic, but count for pie chart)
                $systemFiles = $siteFileData.Files | Where-Object {
                    $_.Name -match '^~' -or $_.Name -match '^\.' -or $_.Name -match '^Forms$' -or 
                    $_.Name -match '^_vti_' -or $_.Name -match '^appdata' -or $_.Name -match '^.DS_Store$' -or 
                    $_.Name -match '^Thumbs.db$'
                }
                $systemFilesSize = ($systemFiles | Measure-Object -Property Size -Sum).Sum
                $allSiteSummaries += [PSCustomObject]@{
                    SiteName = $site.DisplayName
                    SiteUrl  = $site.WebUrl
                    TotalFiles = $siteFileData.TotalFiles
                    TotalSizeGB = $siteFileData.TotalSizeGB
                    TotalFolders = $siteFileData.FolderSizes.Count
                    UniquePermissionLevels = 0
                    RecycleBinSizeGB = [math]::Round($recycleBinSize / 1GB, 3)
                    SystemFilesSizeGB = [math]::Round($systemFilesSize / 1GB, 3)
                    ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
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
                # Pie chart breakdown: add system files and recycle bin as categories
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
                # Add worksheet for recycle bin files
                if ($recycleBinFiles.Count -gt 0) {
                    $recycleBinFiles | Export-Excel -Path $excelFileName -WorksheetName ("RecycleBin - " + $site.DisplayName.Substring(0, [Math]::Min($site.DisplayName.Length, 25))) -AutoSize -TableStyle Medium12
                }
                $detailedSites += $site.Id
            } else {
                # --- Only summary for other sites ---
                $allSiteSummaries += [PSCustomObject]@{
                    SiteName = $site.DisplayName
                    SiteUrl  = $site.WebUrl
                    TotalFiles = $null
                    TotalSizeGB = $siteSummary.StorageGB
                    TotalFolders = $null
                    UniquePermissionLevels = $null
                    RecycleBinSizeGB = $null
                    SystemFilesSizeGB = $null
                    ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
            }
        }

        # Tenant-wide storage pie chart (top 10 sites)
        $tenantPieChart = $siteStorageStats.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 | ForEach-Object {
            [PSCustomObject]@{
                SiteName = $_.Key
                TotalSizeGB = $_.Value
            }
        }

        # Save Excel report
        $excel = $allSiteSummaries | Export-Excel -Path $excelFileName -WorksheetName "Summary" -AutoSize -TableStyle Medium2 -PassThru
        $allTopFiles      | Export-Excel -ExcelPackage $excel -WorksheetName "Top 20 Files"    -AutoSize -TableStyle Medium6
        if ($allTopFolders.Count -gt 0) {
            $allTopFolders | Export-Excel -ExcelPackage $excel -WorksheetName "Top 10 Folders"  -AutoSize -TableStyle Medium3
        }
        if ($tenantPieChart.Count -gt 0) {
            $tenantPieChart | Export-Excel -ExcelPackage $excel -WorksheetName "Tenant Storage Pie" -AutoSize -TableStyle Medium4
        }
        # Add worksheet for files with long path+name (>=399 chars)
        $longPathFiles = $fileData.Files | Where-Object { $_.PathLength -ge 399 }
        if ($longPathFiles.Count -gt 0) {
            $longPathFiles | Export-Excel -ExcelPackage $excel -WorksheetName "Long Path Files" -AutoSize -TableStyle Medium8
        }
        if ($tenantPieChart -and $tenantPieChart.Count -gt 0) {
            foreach ($site in $tenantPieChart) {
                $siteName = $site.SiteName
                if ($sitePieCharts.ContainsKey($siteName) -and $sitePieCharts[$siteName]) {
                    $sitePieCharts[$siteName] | Export-Excel -ExcelPackage $excel -WorksheetName ("Pie - " + $siteName.Substring(0, [Math]::Min($siteName.Length, 25))) -AutoSize -TableStyle Medium4
                }
            }
            $ws = $excel.Workbook.Worksheets["Tenant Storage Pie"]
            if ($ws) {
                $chart = $ws.Drawings.AddChart("TenantStorageChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
                $chart.Title.Text = "Tenant Storage Usage (Top 10 Sites)"
                $chart.SetPosition(1, 0, 7, 0)
                $chart.SetSize(500, 400)
                $series = $chart.Series.Add($ws.Cells["B2:B$($tenantPieChart.Count + 1)"], $ws.Cells["A2:A$($tenantPieChart.Count + 1)"])
                $series.Header = "Size (GB)"
            }
            foreach ($site in $tenantPieChart) {
                $siteName = $site.SiteName
                $wsSite = $excel.Workbook.Worksheets["Pie - " + $siteName.Substring(0, [Math]::Min($siteName.Length, 25))]
                if ($wsSite -and $sitePieCharts.ContainsKey($siteName) -and $sitePieCharts[$siteName]) {
                    $chartSite = $wsSite.Drawings.AddChart("SiteStorageChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
                    $chartSite.Title.Text = "Storage Usage by Folder (Top 10)"
                    $chartSite.SetPosition(1, 0, 7, 0)
                    $chartSite.SetSize(500, 400)
                    $rowCount = $sitePieCharts[$siteName].Count
                    if ($rowCount -gt 0) {
                        $chartSite.Series.Add($wsSite.Cells["B2:B$($rowCount + 1)"], $wsSite.Cells["A2:A$($rowCount + 1)"])
                    }
                }
            }
        }
        # Only close and save the Excel package if at least one worksheet exists
        if ($excel -and $excel.Workbook.Worksheets.Count -gt 0) {
            Close-ExcelPackage $excel
            Write-Host "`nReport saved to: $excelFileName" -ForegroundColor Green
        } else {
            Write-Host "No data was found to export. No Excel report was generated." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "`nError: $_" -ForegroundColor Red
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor DarkRed
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    }
}

# --- Batch Get Drives for Multiple Sites ---
function Get-DrivesBatch {
    param(
        [Parameter(Mandatory)]
        [array]$SiteIds
    )
    $batchRequests = @()
    $responses = @()
    $batchSize = 20  # Microsoft Graph batch limit is 20 per request
    for ($i = 0; $i -lt $SiteIds.Count; $i += $batchSize) {
        $batch = $SiteIds[$i..([Math]::Min($i+$batchSize-1, $SiteIds.Count-1))]
        $batchRequests = @()
        $reqId = 1
        foreach ($siteId in $batch) {
            $batchRequests += @{
                id = "$reqId"
                method = "GET"
                url = "/sites/$siteId/drives"
            }
            $reqId++
        }
        $body = @{ requests = $batchRequests } | ConvertTo-Json -Depth 6
        $result = Invoke-MgGraphRequest -Method POST -Uri "/v1.0/$batch" -Body $body -ContentType 'application/json'
        $responses += $result.responses
    }
    return $responses
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
        $body = @{ requests = $batchRequests } | ConvertTo-Json -Depth 6
        $result = Invoke-MgGraphRequest -Method POST -Uri "/v1.0/\$batch" -Body $body -ContentType 'application/json'
        $responses += $result.responses
    }
    return $responses
}

# --- Batch Collect File Data for Drives ---
function Get-FileDataBatch {
    param(
        [array]$DrivesBatchResponses
    )
    $allFiles = @()
    $folderSizes = @{
    }
    foreach ($resp in $DrivesBatchResponses) {
        if ($resp.status -eq 200 -and $resp.body.value) {
            foreach ($item in $resp.body.value) {
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
                    $folderSizes[$folderPath] += $item.size
                }
            }
        }
    }
    return @{
        Files = $allFiles
        FolderSizes = $folderSizes
        TotalFiles = $allFiles.Count
        TotalSizeGB = [math]::Round(($allFiles | Measure-Object -Property Size -Sum).Sum / 1GB, 2)
    }
}

# Execute the script
Main