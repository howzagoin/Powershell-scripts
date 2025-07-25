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
$WarningPreference     = "SilentlyContinue"

#--- Configuration ---
$clientId              = '278b9af9-888d-4344-93bb-769bdd739249'
$tenantId              = 'ca0711e2-e703-4f4e-9099-17d97863211c'
$siteUrl               = 'https://fbaint.sharepoint.com/sites/Marketing'
$certificateThumbprint = '2E2502BB1EDB8F36CF9DE50936B283BDD22D5BAD'

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
    $uri      = [Uri]$SiteUrl
    $sitePath = $uri.AbsolutePath
    $siteId   = "$($uri.Host):$sitePath"

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

#--- Get Drive Items Recursively ---
function Get-DriveItems {
    param(
        [string]$DriveId,
        [string]$Path,
        [int]$Depth = 0,
        [ref]$GlobalItemIndex,
        [ref]$SpinnerStep
    )
    $items = @()
    try {
        $children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $Path -All -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        if ($children -and $children.Count -gt 0) {
            foreach ($child in $children) {
                if ($null -eq $child -or -not $child.Id) { continue }
                $items += $child
                $GlobalItemIndex.Value++
                $SpinnerStep.Value++
                $spinner = '|/-\'[$SpinnerStep.Value % 4]
                Write-Progress -Activity "Scanning SharePoint Site Content" `
                    -Status "$spinner Scanned: $($GlobalItemIndex.Value) items. Processing: $($child.Name)" `
                    -PercentComplete 0 -Id 1
                if ($child.Folder -and $Depth -lt 10) {
                    try {
                        $items += Get-DriveItems -DriveId $DriveId -Path $child.Id -Depth ($Depth + 1) -GlobalItemIndex $GlobalItemIndex -SpinnerStep $SpinnerStep
                    } catch {}
                }
            }
        }
    } catch {}
    return $items
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

#--- Load/Save Incremental Scan Cache ---
$cacheFile = Join-Path $PSScriptRoot 'SharePointScanCache.json'
function Load-ScanCache {
    if (Test-Path $cacheFile) {
        return Get-Content $cacheFile | ConvertFrom-Json
    }
    return @{}
}
function Save-ScanCache($cache) {
    $cache | ConvertTo-Json -Depth 10 | Set-Content $cacheFile
}

#--- Graph Batch API Helper (with nextLink support) ---
function Invoke-GraphBatchRequest {
    param(
        [string[]]$DriveIds,
        [string[]]$ItemIds,
        [string[]]$NextLinks = @()
    )
    $batchRequests = @()
    for ($i = 0; $i -lt $DriveIds.Count; $i++) {
        if ($NextLinks -and $NextLinks[$i]) {
            $batchRequests += @{ id = "$i"; method = "GET"; url = $NextLinks[$i] }
        } else {
            $batchRequests += @{ id = "$i"; method = "GET"; url = "/drives/$($DriveIds[$i])/items/$($ItemIds[$i])/children" }
        }
    }
    $body     = @{ requests = $batchRequests } | ConvertTo-Json -Depth 5
    $response = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/`$batch" -Body $body -ContentType 'application/json'
    $results  = @{}
    $nextLinks= @{}
    foreach ($resp in $response.responses) {
        if ($resp.status -eq 200 -and $resp.body.value) {
            $results[$resp.id]  = $resp.body.value
            if ($resp.body.'@odata.nextLink') {
                $nextLinks[$resp.id] = $resp.body.'@odata.nextLink'.Replace('https://graph.microsoft.com/v1.0', '')
            }
        } else {
            $results[$resp.id] = @()
        }
    }
    return @{ results = $results; nextLinks = $nextLinks }
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
        try {
            $chart2 = $wsSummary.Drawings.AddChart("FoldersPieChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
            $chart2.Title.Text = "Top 10 Folders by Size (GB)"
            $chart2.SetPosition($siteSummary.Count + 10, 0, 0, 0)
            $chart2.SetSize(500, 400)
            $series2 = $chart2.Series.Add($wsFolders.Cells["B2:B11"], $wsFolders.Cells["A2:A11"])
            $series2.Header = "Size (GB)"
        } catch {
            Write-Host "Warning: Could not add pie chart for Top 10 Folders - $_" -ForegroundColor Yellow
        }
    } else {
        Write-Host "Warning: Could not get worksheet Top 10 Folders. Pie chart not added to Summary." -ForegroundColor Yellow
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

# --- Import user/MFA/license audit functions from old_mfaversion4.ps1 ---
# (Imported functions: Write-Log, Ensure-Modules, Test-RequiredRoles, Test-GraphPermissions, Get-UserInfo, Get-LicensedInactiveUsers, Get-ActiveUsers, Fetch-UserLicenseStatus, Get-MFADisabledUsersFromCSV, Write-CustomLog, Get-UserInfoDetailed, Install-ModuleIfMissing, Connect-MicrosoftGraphWithMFA, Get-UserMFAStatus)
# ...imported functions from old_mfaversion4.ps1 here...
# (For brevity, only function headers shown. Full function bodies should be inserted.)

# --- Unified Main Menu and Workflow ---
function Unified-Main {
    try {
        Write-Host "Unified User, Group, and SharePoint Audit" -ForegroundColor Green
        Write-Host "====================================================" -ForegroundColor Green
        Write-Host "1. User/MFA/License Audit" -ForegroundColor Cyan
        Write-Host "2. Group/SharePoint Access Audit" -ForegroundColor Cyan
        Write-Host "3. Both Audits" -ForegroundColor Cyan
        $choice = Read-Host "Select an option (1/2/3)"

        # Ensure required modules
        $requiredModules = @(
            'Microsoft.Graph.Authentication',
            'Microsoft.Graph.Sites',
            'Microsoft.Graph.Files',
            'Microsoft.Graph.Users',
            'ImportExcel'
        )
        foreach ($module in $requiredModules) {
            if (-not (Get-Module -ListAvailable -Name $module)) {
                Write-Host "Installing $module..." -ForegroundColor Yellow
                Install-Module -Name $module -Force -AllowClobber -SkipPublisherCheck -WarningAction SilentlyContinue
            }
            Import-Module -Name $module -Force -WarningAction SilentlyContinue
        }

        # Connect to Microsoft Graph (only once)
        Connect-ToGraph
        $tenantName = (Get-MgOrganization | Select-Object -ExpandProperty DisplayName) -replace '[^a-zA-Z0-9_-]', '_'
        $dateStr = Get-Date -Format yyyyMMdd_HHmmss
        $defaultFileName = "${tenantName}_User&GroupAudit_${dateStr}.xlsx"
        Add-Type -AssemblyName PresentationFramework
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Title = "Save Unified Audit Excel File"
        $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $saveDialog.FileName = $defaultFileName
        $saveDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
        if ($saveDialog.ShowDialog() -ne $true) {
            Write-Host "Save operation cancelled." -ForegroundColor Yellow
            return
        }
        $filePath = $saveDialog.FileName

        $excel = $null
        if ($choice -eq '1' -or $choice -eq '3') {
            # --- Run User/MFA/License Audit ---
            Write-Host "Running User/MFA/License Audit..." -ForegroundColor Green
            # (Call user audit logic from old_mfaversion4.ps1, collect results in $userAuditResults)
            $userAuditResults = @() # Replace with actual function call to get user audit results
            # Example: $userAuditResults = Get-UserInfo -AllUsers
            if ($userAuditResults.Count -gt 0) {
                $excel = $userAuditResults | Export-Excel -Path $filePath -WorksheetName "User_MFA_License_Audit" -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter -TableStyle Medium6 -Title "User/MFA/License Audit" -PassThru
            }
        }
        if ($choice -eq '2' -or $choice -eq '3') {
            # --- Run Group/SharePoint Access Audit ---
            Write-Host "Running Group/SharePoint Access Audit..." -ForegroundColor Green
            $site = Get-SiteInfo -SiteUrl $siteUrl
            $fileData = Get-FileData -Site $site -Incremental:$false
            $folderAccess = Get-ParentFolderAccess -Site $site
            # Use improved Excel export logic, add as new worksheets to same file
            if ($excel) {
                New-ExcelReport -FileData $fileData -FolderAccess $folderAccess -Site $site -FileName $filePath -ExcelPackage $excel
            } else {
                New-ExcelReport -FileData $fileData -FolderAccess $folderAccess -Site $site -FileName $filePath
            }
        }
        if ($excel) {
            Close-ExcelPackage $excel
            Start-Process -FilePath $filePath
            Write-Host "Unified Excel report created at $filePath" -ForegroundColor Green
        }
    } catch {
        Write-Host "Error: $_" -ForegroundColor Red
    } finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Green
    }
}

# --- Replace old Main with Unified-Main ---
Unified-Main

#--- End of Script ---

# --- Calendar: Check Calendar Permissions ---
function Get-CalendarPermissionsForUser {
    param(
        [string]$CalendarEmail
    )
    $calendarPath = "${CalendarEmail}:\Calendar"
    try {
        $permissions = Get-MailboxFolderPermission -Identity $calendarPath
        return $permissions
    } catch {
        Write-Host "Failed to retrieve calendar permissions for ${CalendarEmail}: $_" -ForegroundColor Red
        return $null
    }
}

# --- Mailbox: Check All Mailbox Rules for External Redirects ---
function Get-MailboxesWithExternalRedirectRules {
    param(
        [string]$SafeDomain = "safecompanydomain.com"
    )
    $results = @()
    $mailboxes = Get-Mailbox -ResultSize Unlimited
    foreach ($mbx in $mailboxes) {
        $rules = Get-InboxRule -Mailbox $mbx.PrimarySmtpAddress
        $externalRules = $rules | Where-Object { $_.RedirectTo -and $_.RedirectTo -notlike "*@${SafeDomain}" }
        foreach ($rule in $externalRules) {
            $results += [PSCustomObject]@{
                Mailbox = $mbx.PrimarySmtpAddress
                RuleName = $rule.Name
                RedirectTo = $rule.RedirectTo
                Enabled = $rule.Enabled
            }
        }
    }
    return $results
}

# --- Mailbox: Get User Groups, Mailboxes, SharePoint Access ---
function Get-UserGroupsMailboxesSharePoint {
    param(
        [string]$UserPrincipalName
    )
    $user = Get-MgUser -UserId $UserPrincipalName
    $mailboxType = if ($user.MailboxSettings) { "Regular" } else { "No mailbox found" }
    $delegatedPermissions = Get-MgUserMailFolderPermission -UserId $UserPrincipalName -MailFolderId 'inbox'
    $groups = Get-MgUserMemberOf -UserId $UserPrincipalName
    $sites = Get-MgSite -Filter "owners/any(o: o/email eq '$UserPrincipalName')"
    return [PSCustomObject]@{
        User = $user.DisplayName
        MailboxType = $mailboxType
        DelegatedMailboxes = $delegatedPermissions
        Groups = $groups
        SharePointSites = $sites
    }
}

# --- Mailbox: Get All Mailbox Rules for a User ---
function Get-MailboxRulesForUser {
    param(
        [string]$EmailAddress
    )
    $rules = Get-InboxRule -Mailbox $EmailAddress | Sort-Object Date -Descending
    $redirectRule = Get-Mailbox -Identity $EmailAddress | Select-Object -ExpandProperty ForwardingAddress
    $redirectRuleEnabled = Get-Mailbox -Identity $EmailAddress | Select-Object -ExpandProperty DeliverToMailboxAndForward
    return [PSCustomObject]@{
        Rules = $rules
        RedirectRule = $redirectRule
        RedirectRuleEnabled = $redirectRuleEnabled
    }
}

# --- Mailbox: Export List of Mailboxes and Delegates ---
function Get-MailboxesAndDelegates {
    $results = @()
    $mailboxes = Get-Mailbox -ResultSize Unlimited
    foreach ($mailbox in $mailboxes) {
        $delegates = Get-MailboxPermission -Identity $mailbox.UserPrincipalName | Where-Object { $_.AccessRights -eq "FullAccess" -and $_.IsInherited -eq $false }
        $sendAs = Get-RecipientPermission -Identity $mailbox.UserPrincipalName | Where-Object { $_.AccessRights -contains "SendAs" }
        foreach ($delegate in $delegates) {
            $results += [PSCustomObject]@{
                Mailbox = $mailbox.UserPrincipalName
                DelegateUser = $delegate.User
                AccessType = "Delegate"
            }
        }
        foreach ($sendAsUser in $sendAs) {
            $results += [PSCustomObject]@{
                Mailbox = $mailbox.UserPrincipalName
                DelegateUser = $sendAsUser.Trustee
                AccessType = "SendAs"
            }
        }
    }
    return $results
}

# --- Mailbox: Get Mailboxes and Distribution Lists for a User ---
function Get-MailboxesAndDLsForUser {
    param(
        [string]$UserEmail
    )
    $mailboxes = Get-Mailbox | Get-MailboxPermission -User $UserEmail
    $dls = Get-DistributionGroup | Where-Object { (Get-DistributionGroupMember $_.Name | ForEach-Object { $_.PrimarySmtpAddress }) -contains $UserEmail }
    return [PSCustomObject]@{
        Mailboxes = $mailboxes
        DistributionLists = $dls
    }
}

# --- Teams: Get a User's Teams Meeting Policy Details ---
function Get-TeamsMeetingPolicyDetails {
    param(
        [string]$UserPrincipalName
    )
    $user = Get-MgUser -UserId $UserPrincipalName
    $meetingPolicies = Get-MgTeamsMeetingPolicy
    $userMeetingPolicy = Get-CsTeamsMeetingPolicy -Identity $UserPrincipalName
    $permissionPolicies = Get-MgTeamsAppPermissionPolicy
    $userPermissionPolicy = Get-CsTeamsAppPermissionPolicy -Identity $UserPrincipalName
    $setupPolicies = Get-MgTeamsAppSetupPolicy
    $userSetupPolicy = Get-CsTeamsAppSetupPolicy -Identity $UserPrincipalName
    return [PSCustomObject]@{
        User = $user
        UserMeetingPolicy = $userMeetingPolicy
        MeetingPolicies = $meetingPolicies
        UserPermissionPolicy = $userPermissionPolicy
        PermissionPolicies = $permissionPolicies
        UserSetupPolicy = $userSetupPolicy
        SetupPolicies = $setupPolicies
    }
}
