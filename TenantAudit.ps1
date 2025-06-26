<#
.SYNOPSIS
  Unified Microsoft 365 Tenant Audit Script: Users, Groups, Mailboxes, and Teams

.DESCRIPTION
  Performs a comprehensive audit across the entire Microsoft 365 tenant, EXCLUDING SharePoint site/folder/file access or storage scanning. This script includes:
  - All users: MFA status, license status, group memberships, mailbox details, mailbox rules, delegates, external forwarding, calendar permissions, and Teams policies
  - All groups: membership, access, and permissions
  - All mailboxes: size, delegates, rules, external forwarding, and distribution list membership
  - Teams: user meeting, permission, and setup policies
  - Aggregates and exports all results to a well-structured Excel report with multiple worksheets and charts for each data type

  Features:
  - Scans the entire tenant (not just a single site)
  - Excludes SharePoint site/folder/file access and storage scanning (use SharePointAudit2025.ps1 for SharePoint audits)
  - Aggregates and summarizes results for easy review
  - Modern error handling and reporting
  - Modular, maintainable, and extensible design
#>

# Set strict error handling
$ErrorActionPreference = "Stop"
$WarningPreference     = "SilentlyContinue"

#--- Configuration ---
$clientId              = '278b9af9-888d-4344-93bb-769bdd739249'
$tenantId              = 'ca0711e2-e703-4f4e-9099-17d97863211c'
$certificateThumbprint = 'B0AF0EF7659EA83D3140844F4BF89CCBB9413DBA'

#--- Required Modules ---
$requiredModules = @(
    'Microsoft.Graph.Authentication',
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
    } catch {
        # Silently handle errors
    }
    return $count
}

function Show-Spinner {
    param([int]$Step)
    $spinners = @('|', '/', '-', '\')
    return $spinners[$Step % $spinners.Length]
}

#--- Retry Wrapper for Graph API Calls ---
function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 5,
        [int]$DelaySeconds = 2
    )
    $attempt = 0
    while ($true) {
        try {
            return & $ScriptBlock
        } catch {
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode -eq 429) {
                $attempt++
                if ($attempt -ge $MaxRetries) { throw }
                $wait = $DelaySeconds * $attempt
                Write-Host "Throttled (429). Retrying in $wait seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds $wait
            } else {
                throw
            }
        }
    }
}

#--- Collect File Data (with Batch API for children, with pagination, hashtable-safe) ---
function Get-FileData {
    param(
        $Site,
        [switch]$Incremental,
        [int]$MaxDepth        = 5,
        [string[]]$IncludeFolders = @(),
        [string[]]$ExcludeFolders = @()
    )
    Write-Host "Analyzing site structure (SharePoint List API)..." -ForegroundColor Cyan

    $scanCache = if ($Incremental) { Load-ScanCache } else { @{} }

    $result       = Get-SharePointLibraryFilesViaListApi -Site $Site -Incremental:$Incremental -scanCache $scanCache
    $allFiles     = $result.Files
    $folderSizes  = $result.FolderSizes

    Write-Host "Site analysis complete - Found $($allFiles.Count) files via List API" -ForegroundColor Green

    # Save scan cache for incremental runs
    if ($Incremental) {
        $newCache = @{}
        foreach ($item in $allFiles) {
            $newCache[$item.Name] = @{ lastModifiedDateTime = $item.LastModifiedDateTime }
        }
        Save-ScanCache $newCache
    }

    return @{
        Files        = $allFiles
        FolderSizes  = $folderSizes
        TotalFiles   = $allFiles.Count
        TotalSizeGB  = [math]::Round(($allFiles | Measure-Object -Property Size -Sum).Sum / 1GB, 2)
    }
}

#--- Get All Files in SharePoint Library via List API ---
function Get-SharePointLibraryFilesViaListApi {
    param(
        $Site,
        [switch]$Incremental,
        $scanCache = @{}
    )
    $allFiles    = @()
    $allFolders  = @()
    $folderSizes = @{}

    $lists = Invoke-WithRetry { Get-MgSiteList -SiteId $Site.Id -WarningAction SilentlyContinue }
    $totalLists = ($lists | Where-Object { $_.List -and $_.List.Template -eq 'documentLibrary' }).Count
    $listIndex = 0
    $totalFilesSoFar = 0
    foreach ($list in $lists) {
        if ($list.List -and $list.List.Template -eq 'documentLibrary') {
            $listIndex++
            # Reduce $top to 200 for smaller payloads
            $uri      = "/v1.0/sites/$($Site.Id)/lists/$($list.Id)/items?expand=fields,driveItem&`$top=200"
            $more     = $true
            $nextLink = $null
            $filesInThisList = 0
            $foldersInThisList = 0
            while ($more) {
                try {
                    $resp = Invoke-WithRetry {
                        if ($nextLink) {
                            Invoke-MgGraphRequest -Method GET -Uri $nextLink
                        } else {
                            Invoke-MgGraphRequest -Method GET -Uri $uri
                        }
                    }
                } catch {
                    Write-Host "Error during library scan: $($list.DisplayName)" -ForegroundColor Red
                    Write-Host "Request URI: $($nextLink ? $nextLink : $uri)" -ForegroundColor Yellow
                    Write-Host "Exception: $_" -ForegroundColor Red
                    throw
                }
                $batchCount = 0
                foreach ($item in $resp.value) {
                    if ($item.driveItem) {
                        if ($item.driveItem.file) {
                            $allFiles += [PSCustomObject]@{
                                Name                   = $item.driveItem.name
                                Size                   = [long]$item.driveItem.size
                                SizeGB                 = [math]::Round($item.driveItem.size / 1GB, 3)
                                SizeMB                 = [math]::Round($item.driveItem.size / 1MB, 2)
                                Path                   = $item.driveItem.parentReference ? $item.driveItem.parentReference.path : ''
                                Drive                  = $item.driveItem.parentReference ? $item.driveItem.parentReference.driveId : ''
                                Extension              = [System.IO.Path]::GetExtension($item.driveItem.name).ToLower()
                                LastModifiedDateTime   = $item.driveItem.lastModifiedDateTime
                            }
                            $folderPath = $item.driveItem.parentReference ? $item.driveItem.parentReference.path : ''
                            if (-not $folderSizes.ContainsKey($folderPath)) { $folderSizes[$folderPath] = 0 }
                            $folderSizes[$folderPath] += $item.driveItem.size
                            $filesInThisList++
                            $totalFilesSoFar++
                        } elseif ($item.driveItem.folder) {
                            $allFolders += [PSCustomObject]@{
                                Name = $item.driveItem.name
                                Path = $item.driveItem.parentReference ? $item.driveItem.parentReference.path : ''
                                Drive = $item.driveItem.parentReference ? $item.driveItem.parentReference.driveId : ''
                                ChildCount = $item.driveItem.folder.childCount
                            }
                            $foldersInThisList++
                        }
                        $batchCount++
                        # Show progress with current folder as encountered (no sorting)
                        $currentFolder = $item.driveItem.parentReference ? $item.driveItem.parentReference.path : 'Root'
                        Write-Progress -Activity "Scanning SharePoint Document Libraries" -Status "Library $listIndex/${totalLists}: $($list.DisplayName) | Files: $totalFilesSoFar | Folders: $($allFolders.Count) | Current Folder: $currentFolder" -PercentComplete ([math]::Min(100, ($listIndex-1)/$totalLists*100 + ($filesInThisList/1000)))
                    }
                }
                if ($resp.'@odata.nextLink') {
                    $nextLink = $resp.'@odata.nextLink'
                } else {
                    $more = $false
                }
                # Optional: Add a small randomized delay to avoid throttling
                Start-Sleep -Milliseconds (Get-Random -Minimum 100 -Maximum 400)
            }
        }
    }
    Write-Progress -Activity "Scanning SharePoint Document Libraries" -Completed
    return @{ Files = $allFiles; Folders = $allFolders; FolderSizes = $folderSizes }
}

#--- Get Parent Folder Access Information ---
function Get-ParentFolderAccess {
    param($Site)

    Write-Host "Retrieving parent folder access information..." -ForegroundColor Cyan

    $folderAccess     = @()
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
                                            FolderName      = $folder.Name
                                            FolderPath      = $folder.ParentReference.Path + '/' + $folder.Name
                                            UserName        = $identity.User.DisplayName
                                            UserEmail       = $identity.User.Email
                                            PermissionLevel = $roles
                                            AccessType      = if ($roles -match 'owner|write') { 'Full/Edit' } 
                                                              elseif ($roles -match 'read') { 'Read Only' } 
                                                              else { 'Other' }
                                        }
                                    }
                                }
                            }

                            if ($perm.GrantedTo -and $perm.GrantedTo.User.DisplayName) {
                                $folderAccess += [PSCustomObject]@{
                                    FolderName      = $folder.Name
                                    FolderPath      = $folder.ParentReference.Path + '/' + $folder.Name
                                    UserName        = $perm.GrantedTo.User.DisplayName
                                    UserEmail       = $perm.GrantedTo.User.Email
                                    PermissionLevel = $roles
                                    AccessType      = if ($roles -match 'owner|write') { 'Full/Edit' } 
                                                      elseif ($roles -match 'read') { 'Read Only' } 
                                                      else { 'Other' }
                                }
                            }
                        }
                    } catch {
                        Write-Host "Warning: Could not retrieve permissions for folder $($folder.Name) - $_" -ForegroundColor Yellow
                    }
                }
            } catch {
                Write-Host "Warning: Could not access drive $($drive.Id) - $_" -ForegroundColor Yellow
            }
        }

        # Remove duplicates (same user with same access to same folder)
        $folderAccess = $folderAccess | Sort-Object FolderName, UserName, PermissionLevel -Unique

        Write-Host "Found access data for $($folderAccess.Count) parent folder permissions" -ForegroundColor Green

    } catch {
        Write-Host "Error retrieving folder access data: $_" -ForegroundColor Red
        $folderAccess = @(
            [PSCustomObject]@{
                FolderName      = "Permission Error"
                FolderPath      = "Check permissions"
                UserName        = "Unable to retrieve data"
                UserEmail       = "Requires additional permissions"
                PermissionLevel = "N/A"
                AccessType      = "Error"
            }
        )
    }

    return $folderAccess
}

#--- Get Site Users and Groups (Owners, Members, Guests, Externals) ---
function Get-SiteUserAccessSummary {
    param($Site, $FolderAccess)

    $owners   = @()
    $members  = @()
    $guests   = @()
    $externals= @()

    try {
        # Get site groups
        $groups = Invoke-WithRetry { Get-MgSiteGroup -SiteId $Site.Id -WarningAction SilentlyContinue }
        $ownersGroup  = $groups | Where-Object { $_.DisplayName -match 'Owner' }
        $membersGroup = $groups | Where-Object { $_.DisplayName -match 'Member' }
        $visitorsGroup= $groups | Where-Object { $_.DisplayName -match 'Visitor' }

        # Get group members
        $getGroupUsers = { param($group) if ($group) { Get-MgGroupMember -GroupId $group.Id -All -WarningAction SilentlyContinue } else { @() } }
        $ownerUsers    = & $getGroupUsers $ownersGroup
        $memberUsers   = & $getGroupUsers $membersGroup
        $visitorUsers  = & $getGroupUsers $visitorsGroup

        # Helper to get user email
        function Get-UserEmail($user) {
            if ($user.UserPrincipalName) { return $user.UserPrincipalName }
            if ($user.Mail) { return $user.Mail }
            return $null
        }

        # Classify users
        $allAccess = $FolderAccess | Group-Object UserEmail
        foreach ($userGroup in $allAccess) {
            $userEmail = $userGroup.Name
            $userAccess = $userGroup.Group
            $topFolders = $userAccess | Sort-Object FolderPath | Select-Object -First 1 -ExpandProperty FolderPath
            $userObj = [PSCustomObject]@{
                UserName  = ($userAccess | Select-Object -First 1).UserName
                UserEmail = $userEmail
                TopFolder = $topFolders
                PermissionLevel = ($userAccess | Select-Object -First 1).PermissionLevel
            }
            if ($ownerUsers | Where-Object { Get-UserEmail $_ -eq $userEmail }) {
                $owners += $userObj
            } elseif ($memberUsers | Where-Object { Get-UserEmail $_ -eq $userEmail }) {
                $members += $userObj
            } elseif ($visitorUsers | Where-Object { Get-UserEmail $_ -eq $userEmail }) {
                $guests += $userObj
            } elseif ($userEmail -match '@' -and $userEmail -notmatch $Site.WebUrl) {
                $externals += $userObj
            }
        }
    } catch {
        Write-Host "Error retrieving site user/group access: $_" -ForegroundColor Red
    }
    return @{ Owners = $owners; Members = $members; Guests = $guests; Externals = $externals }
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

    # Top 10 folders by size
    $top10Folders = $FileData.FolderSizes.GetEnumerator() |
                    Sort-Object Value -Descending | Select-Object -First 10 |
                    ForEach-Object {
                        [PSCustomObject]@{
                            FolderPath = $_.Key
                            SizeGB     = [math]::Round($_.Value / 1GB, 3)
                            SizeMB     = [math]::Round($_.Value / 1MB, 2)
                        }
                    }

    # Storage breakdown by location for pie chart
    $storageBreakdown = $FileData.FolderSizes.GetEnumerator() |
                        Sort-Object Value -Descending | Select-Object -First 15 |
                        ForEach-Object {
                            $folderName = if ($_.Key -match '/([^/]+)/?$') { $matches[1] } else { "Root" }
                            [PSCustomObject]@{
                                Location    = $folderName
                                Path        = $_.Key
                                SizeGB      = [math]::Round($_.Value / 1GB, 3)
                                SizeMB      = [math]::Round($_.Value / 1MB, 2)
                                Percentage  = [math]::Round(($_.Value / ($FileData.Files | Measure-Object Size -Sum).Sum) * 100, 1)
                            }
                        }

    # Parent folder access summary
    $accessSummary = $FolderAccess | Group-Object PermissionLevel |
                     ForEach-Object {
                         [PSCustomObject]@{
                             PermissionLevel = $_.Name
                             UserCount       = $_.Count
                             Users           = ($_.Group.UserName | Sort-Object -Unique) -join '; '
                         }
                     }

    # Site summary
    $siteSummary = @(
        [PSCustomObject]@{
            SiteName               = $Site.DisplayName
            SiteUrl                = $Site.WebUrl
            TotalFiles             = $FileData.TotalFiles
            TotalFolders           = $FileData.Folders.Count
            TotalItems             = $FileData.TotalFiles + $FileData.Folders.Count
            TotalSizeGB            = $FileData.TotalSizeGB
            TotalSiteStorageGB     = $null  # Will be set below
            UniquePermissionLevels = ($FolderAccess.PermissionLevel | Sort-Object -Unique).Count
            ReportDate             = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
        }
    )
    # Try to get total site storage from Graph API (site quota usage summary)
    try {
        $siteUsage = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/sites/$($Site.Id)/drive"
        if ($siteUsage.quota -and $siteUsage.quota.used) {
            $siteSummary[0].TotalSiteStorageGB = [math]::Round($siteUsage.quota.used / 1GB, 2)
        }
    } catch {}

    # --- Recycle Bin: Get top 10 largest files and total size ---
    $recycleBinFiles = @()
    $recycleBinSize = 0
    try {
        $recycleUri = "/v1.0/sites/$($Site.Id)/drive/recycleBin?\$top=500"
        $recycleResp = Invoke-WithRetry { Invoke-MgGraphRequest -Method GET -Uri $recycleUri }
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

    # Create Excel file with multiple worksheets (do NOT close until all sheets and charts are added)
    $excel = $siteSummary | Export-Excel -Path $FileName -WorksheetName "Summary" -AutoSize -TableStyle Medium2 -PassThru
    $top20Files        | Export-Excel -ExcelPackage $excel -WorksheetName "Top 20 Files"     -AutoSize -TableStyle Medium6
    $top10Folders      | Export-Excel -ExcelPackage $excel -WorksheetName "Top 10 Folders"   -AutoSize -TableStyle Medium3
    $storageBreakdown  | Export-Excel -ExcelPackage $excel -WorksheetName "Storage Breakdown" -AutoSize -TableStyle Medium4
    $FolderAccess      | Export-Excel -ExcelPackage $excel -WorksheetName "Folder Access"    -AutoSize -TableStyle Medium5
    $accessSummary     | Export-Excel -ExcelPackage $excel -WorksheetName "Access Summary"   -AutoSize -TableStyle Medium1
    if ($recycleBinFiles.Count -gt 0) {
        $recycleBinFiles | Export-Excel -ExcelPackage $excel -WorksheetName "Recycle Bin Top 10" -AutoSize -TableStyle Medium12
    }
    if ($recycleBinSize -gt 0) {
        $recycleBinSummary = [PSCustomObject]@{
            TotalRecycleBinSizeGB = [math]::Round($recycleBinSize / 1GB, 2)
            TotalRecycleBinSizeMB = [math]::Round($recycleBinSize / 1MB, 2)
            FileCount = $recycleBinFiles.Count
        }
        $recycleBinSummary | Export-Excel -ExcelPackage $excel -WorksheetName "Recycle Bin Summary" -AutoSize -TableStyle Medium13
    }

    # Add user/group access summary sheets
    $userAccess = Get-SiteUserAccessSummary -Site $Site -FolderAccess $FolderAccess
    if ($userAccess.Owners.Count -gt 0) {
        $userAccess.Owners | Export-Excel -ExcelPackage $excel -WorksheetName "Owners Access" -AutoSize -TableStyle Medium7
    }
    if ($userAccess.Members.Count -gt 0) {
        $userAccess.Members | Export-Excel -ExcelPackage $excel -WorksheetName "Members Access" -AutoSize -TableStyle Medium8
    }
    if ($userAccess.Guests.Count -gt 0) {
        $userAccess.Guests | Export-Excel -ExcelPackage $excel -WorksheetName "Guests Access" -AutoSize -TableStyle Medium9
    }
    if ($userAccess.Externals.Count -gt 0) {
        $userAccess.Externals | Export-Excel -ExcelPackage $excel -WorksheetName "External Access" -AutoSize -TableStyle Medium10
    }

    # Find files with long path+name (>399 chars)
    $longPathFiles = $FileData.Files | Where-Object {
        $fullPath = (($_.Path -replace '^/drive/root:', '') + '/' + $_.Name).Trim('/').Replace('//','/')
        $fullPath.Length -gt 399
    } | ForEach-Object {
        $fullPath = (($_.Path -replace '^/drive/root:', '') + '/' + $_.Name).Trim('/').Replace('//','/')
        [PSCustomObject]@{
            Name     = $_.Name
            Path     = $_.Path
            FullPath = $fullPath
            Length   = $fullPath.Length
            SizeMB   = $_.SizeMB
        }
    }
    if ($longPathFiles.Count -gt 0) {
        $longPathFiles | Export-Excel -ExcelPackage $excel -WorksheetName "Long Paths (>399 chars)" -AutoSize -TableStyle Medium11
    }

    # Add charts to the storage breakdown worksheet
    $ws = $excel.Workbook.Worksheets["Storage Breakdown"]
    $chart           = $ws.Drawings.AddChart("StorageChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
    $chart.Title.Text= "Storage Usage by Location"
    $chart.SetPosition(1, 0, 7, 0)
    $chart.SetSize(500, 400)
    $series          = $chart.Series.Add($ws.Cells["C2:C$($storageBreakdown.Count + 1)"], $ws.Cells["A2:A$($storageBreakdown.Count + 1)"])
    $series.Header   = "Size (GB)"

    # Add additional summary info
    $summaryStats = [PSCustomObject]@{
        Owners_Count    = $userAccess.Owners.Count
        Members_Count   = $userAccess.Members.Count
        Guests_Count    = $userAccess.Guests.Count
        Externals_Count = $userAccess.Externals.Count
        OtherUsers_Count= ($FolderAccess | Group-Object UserEmail | Where-Object { $_.Name -and ($userAccess.Owners + $userAccess.Members + $userAccess.Guests + $userAccess.Externals | ForEach-Object { $_.UserEmail }) -notcontains $_.Name }).Count
        LongPathFiles   = $longPathFiles.Count
        LargestFile     = ($top20Files | Select-Object -First 1).Name
        LargestFileSize = ($top20Files | Select-Object -First 1).SizeMB
        LargestFolder   = ($top10Folders | Select-Object -First 1).FolderPath
        LargestFolderSize = ($top10Folders | Select-Object -First 1).SizeGB
        RecycleBinSizeGB = [math]::Round($recycleBinSize / 1GB, 2)
        RecycleBinTopFile = ($recycleBinFiles | Select-Object -First 1).Name
        RecycleBinTopFileSize = ($recycleBinFiles | Select-Object -First 1).SizeMB
    }
    $summaryStats | Export-Excel -ExcelPackage $excel -WorksheetName "Summary" -StartRow ($siteSummary.Count + 3) -AutoSize -TableStyle Medium13

    # Add pie chart for top 10 folders to the Summary worksheet
    $wsSummary = $excel.Workbook.Worksheets["Summary"]
    $wsFolders = $excel.Workbook.Worksheets["Top 10 Folders"]
    if ($wsFolders -and $wsSummary) {
        $chart2 = $wsSummary.Drawings.AddChart("FoldersPieChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
        $chart2.Title.Text = "Top 10 Folders by Size (GB)"
        $chart2.SetPosition($siteSummary.Count + 10, 0, 0, 0)
        $chart2.SetSize(500, 400)
        $series2 = $chart2.Series.Add($wsFolders.Cells["B2:B11"], $wsFolders.Cells["A2:A11"])
        $series2.Header = "Size (GB)"
    }

    Close-ExcelPackage $excel

    Write-Host "Excel report created successfully!" -ForegroundColor Green
    Write-Host "`nReport Contents:" -ForegroundColor Cyan
    Write-Host "- Summary: Overall site statistics" -ForegroundColor White
    Write-Host "- Top 20 Files: Largest files by size" -ForegroundColor White  
    Write-Host "- Top 10 Folders: Largest folders by size" -ForegroundColor White
    Write-Host "- Storage Breakdown: Space usage by location with pie chart" -ForegroundColor White
    Write-Host "- Folder Access: Parent folder permissions" -ForegroundColor White
    Write-Host "- Access Summary: Users grouped by permission level" -ForegroundColor White
    Write-Host "- Recycle Bin Top 10: Largest deleted files" -ForegroundColor White
    Write-Host "- Recycle Bin Summary: Total deleted size" -ForegroundColor White
    Write-Host "- Owners/Members/Guests/Externals: User access details" -ForegroundColor White
    Write-Host "- Long Paths: Files with long path+name (>399 chars)" -ForegroundColor White
}

# --- Unified User, Group, and SharePoint Audit Script ---
# Combines user/MFA/license audit and group/SharePoint access audit with unified Excel export

# --- Import user/MFA/license audit functions ---
# .\Microsoft365Audit_UserAudit.ps1
# --- Import group/SharePoint access audit functions ---
# .\Microsoft365Audit_SharePointAudit.ps1

# --- Main Script ---
try {
    # Connect to Microsoft Graph
    Connect-ToGraph

    # --- User, Group, and License Audit ---
    $allUsers = Invoke-WithRetry { Get-MgUser -All -Select Id,DisplayName,UserPrincipalName,Mail,UserType,AssignedLicenses,AccountEnabled,LastPasswordChangeDateTime -Top 999 }
    $userAuditResults = @()
    $userIndex = 0
    $totalUsers = $allUsers.Count
    foreach ($user in $allUsers) {
        $userIndex++
        $spinner = '|/-\'[$userIndex % 4]
        Write-Progress -Activity "Auditing Users" -Status "$spinner Processing user $userIndex of ${totalUsers}: $($user.DisplayName)" -PercentComplete ($userIndex / $totalUsers * 100)
        try {
            # MFA Status
            $mfaStatus = "Unknown"
            try {
                $authMethods = Get-MgUserAuthenticationMethod -UserId $user.Id -All -WarningAction SilentlyContinue
                $mfaStatus = if ($authMethods | Where-Object { $_.OdataType -like '*microsoftAuthenticator*' }) { "Enabled" } else { "Disabled" }
            } catch {}

            # License Status
            $licenseStatus = if ($user.AssignedLicenses -and $user.AssignedLicenses.Count -gt 0) { "Licensed" } else { "Unlicensed" }

            # Active/Inactive (based on AccountEnabled)
            $activeStatus = if ($user.AccountEnabled) { "Active" } else { "Inactive" }

            # Last Password Change
            $lastPasswordChange = $user.LastPasswordChangeDateTime

            # User Timezone
            $timezone = $null
            try {
                $mailboxSettings = Get-MgUserMailboxSetting -UserId $user.Id -ErrorAction SilentlyContinue
                $timezone = $mailboxSettings.TimeZone
            } catch {}

            # Mailbox Size & Archive Size
            $mailboxSize = $null; $archiveSize = $null
            try {
                $mbxStats = Get-MailboxStatistics -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
                if ($mbxStats) { $mailboxSize = $mbxStats.TotalItemSize.ToString() }
                $archiveStats = Get-MailboxStatistics -Identity $user.UserPrincipalName -Archive -ErrorAction SilentlyContinue
                if ($archiveStats) { $archiveSize = $archiveStats.TotalItemSize.ToString() }
            } catch {}

            # Safe Sender List Count
            $safeSenderCount = $null
            try {
                $safeSenders = Get-MailboxJunkEmailConfiguration -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
                if ($safeSenders) { $safeSenderCount = $safeSenders.TrustedSendersAndDomains.Count }
            } catch {}

            # Calendar Permissions (granted to others)
            $calendarPermissions = @()
            try {
                $calendarPermissions = Get-MailboxFolderPermission -Identity ("$($user.UserPrincipalName):\Calendar") -ErrorAction SilentlyContinue | Where-Object { $_.User -ne "Default" -and $_.User -ne "Anonymous" }
            } catch {}

            # Delegated Mailboxes and Delegation Type
            $delegates = @()
            try {
                $delegates = Get-MailboxPermission -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains "FullAccess" -and $_.IsInherited -eq $false }
            } catch {}

            # Mailbox Rules (highlight external redirects)
            $mailboxRules = @()
            $externalRedirects = @()
            try {
                $rules = Get-InboxRule -Mailbox $user.UserPrincipalName -ErrorAction SilentlyContinue
                foreach ($rule in $rules) {
                    $mailboxRules += $rule
                    if ($rule.RedirectTo -and ($rule.RedirectTo | Where-Object { $_ -notlike "*@yourcompany.com" })) {
                        $externalRedirects += $rule
                    }
                }
            } catch {}

            # Teams Policies
            $teamsPolicy = $null
            try {
                $teamsPolicy = Get-CsOnlineUser -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue | Select-Object Teams*Policy*
            } catch {}

            # Licenses (detailed)
            $licenses = $user.AssignedLicenses | ForEach-Object { $_.SkuId }

            $userAuditResults += [PSCustomObject]@{
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                UserGuid = $user.Id
                Mail = $user.Mail
                UserType = $user.UserType
                MFAStatus = $mfaStatus
                LicenseStatus = $licenseStatus
                Licenses = ($licenses -join ", ")
                ActiveStatus = $activeStatus
                LastPasswordChange = $lastPasswordChange
                TimeZone = $timezone
                MailboxSize = $mailboxSize
                ArchiveSize = $archiveSize
                SafeSenderCount = $safeSenderCount
                CalendarPermissions = (@($calendarPermissions | ForEach-Object { $_.User + ':' + $_.AccessRights }) -join "; ")
                Delegates = (@($delegates | ForEach-Object { $_.User + ':' + ($_.AccessRights -join ",") }) -join "; ")
                MailboxRules = (@($mailboxRules | ForEach-Object { $_.Name }) -join "; ")
                ExternalRedirectRules = (@($externalRedirects | ForEach-Object { $_.Name }) -join "; ")
                TeamsPolicies = ($teamsPolicy | ConvertTo-Json -Compress)
            }
        } catch {
            Write-Host "Error processing user $($user.DisplayName): $_" -ForegroundColor Red
        }
    }

    # Export to Excel
    $excelFile = Join-Path $PSScriptRoot "Microsoft365_UserAudit_Detailed_$($tenantId.Substring(0,8))_$((Get-Date).ToString('yyyyMMdd_HHmmss')).xlsx"
    try {
        $userAuditResults | Export-Excel -Path $excelFile -WorksheetName "User Audit" -AutoSize -TableStyle Medium2
        Write-Host "User audit exported to Excel: $excelFile" -ForegroundColor Green
    } catch {
        Write-Host "Error exporting user audit to Excel: $_" -ForegroundColor Red
    }

    # --- Group and SharePoint Access Audit ---
    $allGroups = Invoke-WithRetry { Get-MgGroup -All -Select Id,DisplayName,MailEnabled,SecurityEnabled,GroupTypes -Top 999 }
    $groupAuditResults = @()
    $groupIndex = 0
    $totalGroups = $allGroups.Count
    foreach ($group in $allGroups) {
        $groupIndex++
        $spinner = '|/-\'[$groupIndex % 4]
        Write-Progress -Activity "Auditing Groups" -Status "$spinner Processing group $groupIndex of ${totalGroups}: $($group.DisplayName)" -PercentComplete ($groupIndex / $totalGroups * 100)
        try {
            # Group Type (Unified, Security, etc.)
            $groupType = if ($group.GroupTypes -and $group.GroupTypes.Count -gt 0) { $group.GroupTypes -join ", " } else { "N/A" }

            # Mail-enabled (for distribution lists)
            $mailEnabled = $group.MailEnabled

            # Security-enabled (for security groups)
            $securityEnabled = $group.SecurityEnabled

            # Owners and Members
            $owners = @()
            $members = @()
            try {
                $ownersGroup = $null
                if ($groupType -like "*unified*") {
                    # For Microsoft 365 Groups (Unified groups)
                    $ownersGroup = $group.Id
                } else {
                    # For regular security groups
                    $ownersGroup = ($group | Get-MgGroupOwner -All -WarningAction SilentlyContinue)
                }
                $owners = $ownersGroup | Where-Object { $_.UserPrincipalName } | ForEach-Object { $_.UserPrincipalName }

                # Members (all types)
                $members = $group.Id | Get-MgGroupMember -All -WarningAction SilentlyContinue | Where-Object { $_.UserPrincipalName } | ForEach-Object { $_.UserPrincipalName }
            } catch {}

            # Group Email (for mail-enabled groups)
            $groupEmail = $null
            if ($mailEnabled) {
                try {
                    $groupEmail = ($group | Get-MgGroupEmail -ErrorAction SilentlyContinue).Mail
                } catch {}
            }

            $groupAuditResults += [PSCustomObject]@{
                GroupName         = $group.DisplayName
                GroupId           = $group.Id
                GroupType         = $groupType
                MailEnabled       = $mailEnabled
                SecurityEnabled   = $securityEnabled
                Owners             = ($owners -join "; ")
                Members            = ($members -join "; ")
                GroupEmail        = $groupEmail
            }
        } catch {
            Write-Host "Error processing group $($group.DisplayName): $_" -ForegroundColor Red
        }
    }

    # Export group audit to Excel
    $excelFileGroups = Join-Path $PSScriptRoot "Microsoft365_GroupAudit_$($tenantId.Substring(0,8))_$((Get-Date).ToString('yyyyMMdd_HHmmss')).xlsx"
    try {
        $groupAuditResults | Export-Excel -Path $excelFileGroups -WorksheetName "Group Audit" -AutoSize -TableStyle Medium2
        Write-Host "Group audit exported to Excel: $excelFileGroups" -ForegroundColor Green
    } catch {
        Write-Host "Error exporting group audit to Excel: $_" -ForegroundColor Red
    }

    # --- SharePoint Site and Content Audit ---
    $allSites = Invoke-WithRetry { Get-MgSite -All -Select Id,DisplayName,WebUrl,SiteCollection,CreatedDateTime,LastModifiedDateTime -Top 999 }
    $siteAuditResults = @()
    $siteIndex = 0
    $totalSites = $allSites.Count
    foreach ($site in $allSites) {
        $siteIndex++
        $spinner = '|/-\'[$siteIndex % 4]
        Write-Progress -Activity "Auditing Sites" -Status "$spinner Processing site $siteIndex of ${totalSites}: $($site.DisplayName)" -PercentComplete ($siteIndex / $totalSites * 100)
        try {
            # Site URL and Collection
            $siteUrl = $site.WebUrl
            $siteCollection = $site.SiteCollection

            # Created and Modified Dates
            $createdDate = $site.CreatedDateTime
            $lastModifiedDate = $site.LastModifiedDateTime

            # Owner and Member count
            $owners = @()
            $members = @()
            try {
                $siteGroups = Invoke-WithRetry { Get-MgSiteGroup -SiteId $site.Id -WarningAction SilentlyContinue }
                $ownersGroup  = $siteGroups | Where-Object { $_.DisplayName -match 'Owner' }
                $membersGroup = $siteGroups | Where-Object { $_.DisplayName -match 'Member' }

                $owners = & $getGroupUsers $ownersGroup
                $members = & $getGroupUsers $membersGroup
            } catch {}

            # Storage Metrics
            $storageUsedGB = $null
            $storageQuotaGB = $null
            try {
                $siteDrive = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/sites/$($site.Id)/drive"
                if ($siteDrive.quota -and $siteDrive.quota.used) {
                    $storageUsedGB = [math]::Round($siteDrive.quota.used / 1GB, 2)
                    $storageQuotaGB = [math]::Round($siteDrive.quota.total / 1GB, 2)
                }
            } catch {}

            $siteAuditResults += [PSCustomObject]@{
                SiteName         = $site.DisplayName
                SiteId           = $site.Id
                SiteUrl          = $siteUrl
                SiteCollection   = $siteCollection
                CreatedDate      = $createdDate
                LastModifiedDate = $lastModifiedDate
                OwnerCount       = $owners.Count
                MemberCount      = $members.Count
                StorageUsedGB    = $storageUsedGB
                StorageQuotaGB   = $storageQuotaGB
            }
        } catch {
            Write-Host "Error processing site $($site.DisplayName): $_" -ForegroundColor Red
        }
    }

    # Export site audit to Excel
    $excelFileSites = Join-Path $PSScriptRoot "Microsoft365_SiteAudit_$($tenantId.Substring(0,8))_$((Get-Date).ToString('yyyyMMdd_HHmmss')).xlsx"
    try {
        $siteAuditResults | Export-Excel -Path $excelFileSites -WorksheetName "Site Audit" -AutoSize -TableStyle Medium2
        Write-Host "Site audit exported to Excel: $excelFileSites" -ForegroundColor Green
    } catch {
        Write-Host "Error exporting site audit to Excel: $_" -ForegroundColor Red
    }

    # --- Mailbox and Teams Audit ---
    function Get-MailboxAudit {
        $mailboxes = Get-MgUser -Filter "mail ne null" -All -ErrorAction SilentlyContinue
        $mailboxAudit = @()
        foreach ($mb in $mailboxes) {
            $rules = @()
            try {
                $rules = Get-MgUserMailFolderMessageRule -UserId $mb.Id -MailFolderId 'inbox' -All -ErrorAction SilentlyContinue
            } catch {}
            $externalRedirect = $false
            foreach ($rule in $rules) {
                if ($rule.Actions.ForwardTo -or $rule.Actions.RedirectTo) {
                    foreach ($fwd in ($rule.Actions.ForwardTo + $rule.Actions.RedirectTo)) {
                        if ($fwd.EmailAddress -and $fwd.EmailAddress -notlike "*@yourdomain.com") {
                            $externalRedirect = $true
                        }
                    }
                }
            }
            $mailboxAudit += [PSCustomObject]@{
                DisplayName = $mb.DisplayName
                UserPrincipalName = $mb.UserPrincipalName
                Mail = $mb.Mail
                MailboxSizeMB = $null # Placeholder, requires Exchange Online PowerShell for accurate size
                ExternalRedirect = $externalRedirect
            }
        }
        return $mailboxAudit
    }

    function Get-TeamsAudit {
        $teams = Get-MgTeam -All -ErrorAction SilentlyContinue
        $teamsAudit = @()
        foreach ($team in $teams) {
            $members = Get-MgTeamMember -TeamId $team.Id -All -ErrorAction SilentlyContinue
            $teamsAudit += [PSCustomObject]@{
                TeamName = $team.DisplayName
                TeamId = $team.Id
                MemberCount = $members.Count
                Owners = ($members | Where-Object { $_.Roles -contains 'owner' } | ForEach-Object { $_.DisplayName }) -join '; '
            }
        }
        return $teamsAudit
    }

    $mailboxAudit = Get-MailboxAudit
    $teamsAudit = Get-TeamsAudit

    # Final Excel export (summary + mailbox + teams)
    $excelFileName = Join-Path $PSScriptRoot "Microsoft365_Audit_Summary_$($tenantId.Substring(0,8))_$((Get-Date).ToString('yyyyMMdd_HHmmss')).xlsx"
    $summary = [PSCustomObject]@{
        TenantId              = $tenantId
        ReportGeneratedDate   = Get-Date
        TotalUsers            = $allUsers.Count
        TotalGroups           = $allGroups.Count
        TotalSites            = $allSites.Count
        TotalFiles            = $null  # To be calculated
        TotalFolders          = $null  # To be calculated
        TotalMailboxAccounts  = $null  # To be calculated
        TotalTeams            = $null  # To be calculated
    }

    # Calculate totals for files, folders, mailbox accounts, and teams
    try {
        $fileDataAllSites = @()
        foreach ($site in $allSites) {
            $fileData = Get-FileData -Site $site -Incremental -MaxDepth 1
            $fileDataAllSites += $fileData.Files
        }
        $summary.TotalFiles = $fileDataAllSites.Count
        $summary.TotalFolders = ($fileDataAllSites | Where-Object { $_.Folder }).Count
    } catch {}

    try {
        $mailboxAuditData = Get-MailboxAudit
        $summary.TotalMailboxAccounts = $mailboxAuditData.Count
    } catch {}

    try {
        $teamsAuditData = Get-TeamsAudit
        $summary.TotalTeams = $teamsAuditData.Count
    } catch {}

    # Create final summary Excel file
    $excel = $summary | Export-Excel -Path $excelFileName -WorksheetName "Summary" -AutoSize -TableStyle Medium2 -PassThru
    $top20Files        | Export-Excel -ExcelPackage $excel -WorksheetName "Top 20 Files"     -AutoSize -TableStyle Medium6
    $top10Folders      | Export-Excel -ExcelPackage $excel -WorksheetName "Top 10 Folders"   -AutoSize -TableStyle Medium3
    $storageBreakdown  | Export-Excel -ExcelPackage $excel -WorksheetName "Storage Breakdown" -AutoSize -TableStyle Medium4
    $FolderAccess      | Export-Excel -ExcelPackage $excel -WorksheetName "Folder Access"    -AutoSize -TableStyle Medium5
    $accessSummary     | Export-Excel -ExcelPackage $excel -WorksheetName "Access Summary"   -AutoSize -TableStyle Medium1
    if ($recycleBinFiles.Count -gt 0) {
        $recycleBinFiles | Export-Excel -ExcelPackage $excel -WorksheetName "Recycle Bin Top 10" -AutoSize -TableStyle Medium12
    }
    if ($recycleBinSize -gt 0) {
        $recycleBinSummary = [PSCustomObject]@{
            TotalRecycleBinSizeGB = [math]::Round($recycleBinSize / 1GB, 2)
            TotalRecycleBinSizeMB = [math]::Round($recycleBinSize / 1MB, 2)
            FileCount = $recycleBinFiles.Count
        }
        $recycleBinSummary | Export-Excel -ExcelPackage $excel -WorksheetName "Recycle Bin Summary" -AutoSize -TableStyle Medium13
    }

    # Add user/group access summary sheets
    $userAccess = Get-SiteUserAccessSummary -Site $Site -FolderAccess $FolderAccess
    if ($userAccess.Owners.Count -gt 0) {
        $userAccess.Owners | Export-Excel -ExcelPackage $excel -WorksheetName "Owners Access" -AutoSize -TableStyle Medium7
    }
    if ($userAccess.Members.Count -gt 0) {
        $userAccess.Members | Export-Excel -ExcelPackage $excel -WorksheetName "Members Access" -AutoSize -TableStyle Medium8
    }
    if ($userAccess.Guests.Count -gt 0) {
        $userAccess.Guests | Export-Excel -ExcelPackage $excel -WorksheetName "Guests Access" -AutoSize -TableStyle Medium9
    }
    if ($userAccess.Externals.Count -gt 0) {
        $userAccess.Externals | Export-Excel -ExcelPackage $excel -WorksheetName "External Access" -AutoSize -TableStyle Medium10
    }

    # Mailbox and Teams audit sheets
    $mailboxAudit     | Export-Excel -ExcelPackage $excel -WorksheetName "MailboxAudit" -AutoSize -TableStyle Medium8
    $teamsAudit       | Export-Excel -ExcelPackage $excel -WorksheetName "TeamsAudit" -AutoSize -TableStyle Medium9

    # Find files with long path+name (>399 chars)
    $longPathFiles = $FileData.Files | Where-Object {
        $fullPath = (($_.Path -replace '^/drive/root:', '') + '/' + $_.Name).Trim('/').Replace('//','/')
        $fullPath.Length -gt 399
    } | ForEach-Object {
        $fullPath = (($_.Path -replace '^/drive/root:', '') + '/' + $_.Name).Trim('/').Replace('//','/')
        [PSCustomObject]@{
            Name     = $_.Name
            Path     = $_.Path
            FullPath = $fullPath
            Length   = $fullPath.Length
            SizeMB   = $_.SizeMB
        }
    }
    if ($longPathFiles.Count -gt 0) {
        $longPathFiles | Export-Excel -ExcelPackage $excel -WorksheetName "Long Paths (>399 chars)" -AutoSize -TableStyle Medium11
    }

    # Add charts to the storage breakdown worksheet
    $ws = $excel.Workbook.Worksheets["Storage Breakdown"]
    $chart           = $ws.Drawings.AddChart("StorageChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
    $chart.Title.Text= "Storage Usage by Location"
    $chart.SetPosition(1, 0, 7, 0)
    $chart.SetSize(500, 400)
    $series          = $chart.Series.Add($ws.Cells["C2:C$($storageBreakdown.Count + 1)"], $ws.Cells["A2:A$($storageBreakdown.Count + 1)"])
    $series.Header   = "Size (GB)"

    # Add additional summary info
    $summaryStats = [PSCustomObject]@{
        Owners_Count    = $userAccess.Owners.Count
        Members_Count   = $userAccess.Members.Count
        Guests_Count    = $userAccess.Guests.Count
        Externals_Count = $userAccess.Externals.Count
        OtherUsers_Count= ($FolderAccess | Group-Object UserEmail | Where-Object { $_.Name -and ($userAccess.Owners + $userAccess.Members + $userAccess.Guests + $userAccess.Externals | ForEach-Object { $_.UserEmail }) -notcontains $_.Name }).Count
        LongPathFiles   = $longPathFiles.Count
        LargestFile     = ($top20Files | Select-Object -First 1).Name
        LargestFileSize = ($top20Files | Select-Object -First 1).SizeMB
        LargestFolder   = ($top10Folders | Select-Object -First 1).FolderPath
        LargestFolderSize = ($top10Folders | Select-Object -First 1).SizeGB
        RecycleBinSizeGB = [math]::Round($recycleBinSize / 1GB, 2)
        RecycleBinTopFile = ($recycleBinFiles | Select-Object -First 1).Name
        RecycleBinTopFileSize = ($recycleBinFiles | Select-Object -First 1).SizeMB
    }
    $summaryStats | Export-Excel -ExcelPackage $excel -WorksheetName "Summary" -StartRow ($siteSummary.Count + 3) -AutoSize -TableStyle Medium13

    # Add pie chart for top 10 folders to the Summary worksheet
    $wsSummary = $excel.Workbook.Worksheets["Summary"]
    $wsFolders = $excel.Workbook.Worksheets["Top 10 Folders"]
    if ($wsFolders -and $wsSummary) {
        $chart2 = $wsSummary.Drawings.AddChart("FoldersPieChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
        $chart2.Title.Text = "Top 10 Folders by Size (GB)"
        $chart2.SetPosition($siteSummary.Count + 10, 0, 0, 0)
        $chart2.SetSize(500, 400)
        $series2 = $chart2.Series.Add($wsFolders.Cells["B2:B11"], $wsFolders.Cells["A2:A11"])
        $series2.Header = "Size (GB)"
    }

    Close-ExcelPackage $excel

    Write-Host "Audit complete! Review the exported Excel files for details." -ForegroundColor Green
} catch {
    Write-Host "Error in the unified audit script: $_" -ForegroundColor Red
}
