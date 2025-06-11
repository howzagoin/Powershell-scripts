<#
.SYNOPSIS
  Generates a focused SharePoint storage and access report with Excel charts

.DESCRIPTION
  Creates an Excel report showing:
  - Top 20 largest files
  - Top 10 largest folders
  - Storage breakdown pie chart
  - User access summary showing only parent folder permissions
#>

# Set strict error handling
$ErrorActionPreference = "Stop"
$WarningPreference = "SilentlyContinue"

#--- Configuration ---
$clientId     = '278b9af9-888d-4344-93bb-769bdd739249'
$tenantId     = 'ca0711e2-e703-4f4e-9099-17d97863211c'
$siteUrl      = 'https://fbaint.sharepoint.com/sites/Marketing'
$certificateThumbprint = 'B0AF0EF7659EA83D3140844F4BF89CCBB9413DBA'

#--- Required Modules ---
$requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Sites', 
    'Microsoft.Graph.Files',
    'Microsoft.Graph.Users',
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

#--- Main Execution ---
function Main {
    try {
        Write-Host "SharePoint Storage & Access Report Generator" -ForegroundColor Green
        Write-Host "=============================================" -ForegroundColor Green
        
        # Connect to Microsoft Graph
        Connect-ToGraph
        
        # Get site information
        $site = Get-SiteInfo -SiteUrl $siteUrl
        
        # Collect file data (with progress bar)
        $fileData = Get-FileData -Site $site
        
        # Get parent folder access data
        $folderAccess = Get-ParentFolderAccess -Site $site
        
        # Output summary to console
        Write-Host "`n" + "="*50 -ForegroundColor Green
        Write-Host "SCAN COMPLETE - SITE ANALYSIS SUMMARY" -ForegroundColor Green
        Write-Host "="*50 -ForegroundColor Green
        Write-Host "Site: $($site.DisplayName)" -ForegroundColor Cyan
        Write-Host "Total Files: $($fileData.TotalFiles)" -ForegroundColor White
        Write-Host "Total Size: $($fileData.TotalSizeGB) GB" -ForegroundColor White
        Write-Host "Parent Folders with Access Data: $($folderAccess.Count)" -ForegroundColor White
        Write-Host "="*50 -ForegroundColor Green
        
        # Ask if user wants Excel report and where to save it
        $response = Read-Host "`nWould you like to generate an Excel report? (Y/N)"
        if ($response -match '^[Yy]') {
            Add-Type -AssemblyName PresentationFramework
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Title = "Save SharePoint Excel Report"
            $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $saveDialog.FileName = "SharePoint_Storage_Report_$(Get-Date -Format yyyyMMdd_HHmmss).xlsx"
            $dialogResult = $saveDialog.ShowDialog()
            
            if ($dialogResult -eq $true) {
                New-ExcelReport -FileData $fileData -FolderAccess $folderAccess -Site $site -FileName $saveDialog.FileName
                Write-Host "`nReport saved to: $($saveDialog.FileName)" -ForegroundColor Green
            } else {
                Write-Host "Excel report was not saved." -ForegroundColor Yellow
            }
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

# Execute the script
Main