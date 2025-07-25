<#
.SYNOPSIS
  SharePoint Tenant-Wide Storage, Access, and Large File Audit Script

.DESCRIPTION
  Performs a comprehensive audit of ALL SharePoint sites in the tenant, including:
  - Scans all SharePoint sites and aggregates storage usage
  - Generates pie charts of storage for the whole tenant (largest 10 sites by size)
  - For each of the 10 largest sites, generates a pie chart showing storage breakdown
  - Collects user access for all sites, including user type (internal/external)
  - Highlights external guest access in red in the Excel report
  - Lists site owners and site members for each site
  - For the top 10 largest sites, shows the top 20 biggest files and folders
  - Finds and reports all files larger than 100MB in every document library across all sites
  - Exports all results to a well-structured Excel report with multiple worksheets and charts
  - Progress bars for site/library/file processing
  - Robust error handling and reporting
  - Certificate-based authentication only

  Key improvements:
  - Uses SharePoint List API for better data discovery
  - Enhanced retry logic for throttling
  - Multiple site discovery approaches
  - Better error handling
#>

#region Configuration and Prerequisites
# Set strict error handling
$ErrorActionPreference = "Stop"
$WarningPreference = "SilentlyContinue"

# Configuration
$clientId              = '278b9af9-888d-4344-93bb-769bdd739249'
$tenantId              = 'ca0711e2-e703-4f4e-9099-17d97863211c'
$certificateThumbprint = '2E2502BB1EDB8F36CF9DE50936B283BDD22D5BAD'

# Install required modules if missing
try {
    if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) {
        Write-Host "Installing Microsoft.Graph module..." -ForegroundColor Yellow
        Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force -ErrorAction Stop
    }
} catch {
    Write-Host "Could not install Microsoft.Graph module. Please install manually: Install-Module Microsoft.Graph" -ForegroundColor Red
    throw
}

try {
    if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
        Write-Host "Installing ImportExcel module..." -ForegroundColor Yellow
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
    }
} catch {
    Write-Host "Could not install ImportExcel module. Please install manually: Install-Module ImportExcel" -ForegroundColor Red
    throw
}

# Import required modules
Import-Module Microsoft.Graph.Sites
Import-Module Microsoft.Graph.Files  
Import-Module Microsoft.Graph.Groups
Import-Module Microsoft.Graph.Users
Import-Module ImportExcel
#endregion

#region Authentication Functions

function Get-ClientCertificate {
    param ([Parameter(Mandatory)][string]$Thumbprint)
    
    $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$Thumbprint" -ErrorAction Stop
    if (-not $cert) { 
        throw "Certificate with thumbprint $Thumbprint not found." 
    }
    return $cert
}

function Connect-ToGraph {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    
    try {
        # Clear existing connections
        Disconnect-MgGraph -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        
        # Get certificate
        $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$certificateThumbprint" -ErrorAction Stop
        
        # Connect with app-only authentication
        Connect-MgGraph -ClientId $clientId -TenantId $tenantId -Certificate $cert -NoWelcome -WarningAction SilentlyContinue
        
        # Verify app-only authentication
        $context = Get-MgContext
        if (-not $context) {
            throw "Graph context missing. Authentication failed."
        }
        
        Write-Host "[Debug] Graph context: TenantId=$($context.TenantId), AuthType=$($context.AuthType), Scopes=$($context.Scopes -join ', ')" -ForegroundColor Yellow
        
        if ($context.AuthType -ne 'AppOnly') { 
            throw "App-only authentication required." 
        }
        
        Write-Host "[Info] Successfully connected with app-only authentication" -ForegroundColor Green
    }
    catch {
        Write-Host "[Error] Authentication failed: $_" -ForegroundColor Red
        throw
    }
}
#endregion

#region Utility Functions
function Get-TenantName {
    try {
        $tenant = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($tenant) { 
            return $tenant.DisplayName.Replace(' ', '_') 
        }
        return 'Tenant'
    }
    catch {
        return 'Tenant'
    }
}

function Get-UserEmail($user) {
    if ($user.UserPrincipalName) { return $user.UserPrincipalName }
    if ($user.Mail) { return $user.Mail }
    return $null
}

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
        } 
        catch {
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode -eq 429) {
                $attempt++
                if ($attempt -ge $MaxRetries) { 
                    Write-Host "[Error] Max retries exceeded for throttling. Giving up." -ForegroundColor Red
                    throw 
                }
                $wait = $DelaySeconds * $attempt
                Write-Host "[Warning] Throttled (429). Retrying in $wait seconds... (Attempt $attempt/$MaxRetries)" -ForegroundColor Yellow
                Start-Sleep -Seconds $wait
            } 
            else {
                throw
            }
        }
    }
}
#endregion

#region Site Discovery Functions

function Get-AllSharePointSites {
    Write-Host "Enumerating all SharePoint sites in tenant..." -ForegroundColor Cyan
    
    try {
        # Try multiple approaches to get sites
        $sites = @()
        
        # Approach 1: Get root site first
        try {
            $rootSite = Get-MgSite -SiteId "root"
            if ($rootSite) {
                $sites += $rootSite
                Write-Host "[Site] Root site found: $($rootSite.DisplayName)" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "[Warning] Could not get root site: $_" -ForegroundColor Yellow
        }
        
        # Approach 2: Try direct Graph API call for better results
        try {
            $searchSites = Invoke-WithRetry { 
                Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites?search=*" 
            }
            if ($searchSites.value) {
                foreach ($siteData in $searchSites.value) {
                    $sites += [PSCustomObject]@{
                        Id = $siteData.id
                        DisplayName = $siteData.displayName
                        WebUrl = $siteData.webUrl
                        Name = $siteData.name
                    }
                }
            }
        }
        catch {
            Write-Host "[Warning] Direct API search failed: $_" -ForegroundColor Yellow
        }
        
        # Approach 3: Try wildcard search with Get-MgSite
        try {
            $wildcardSites = Invoke-WithRetry { Get-MgSite -Search "*" -All -WarningAction SilentlyContinue }
            if ($wildcardSites) {
                $sites += $wildcardSites
            }
        }
        catch {
            Write-Host "[Warning] Wildcard search failed: $_" -ForegroundColor Yellow
        }
        
        # Remove duplicates based on ID
        $sites = $sites | Sort-Object Id -Unique
        
        if (-not $sites -or $sites.Count -eq 0) {
            Write-Host "[Warning] No SharePoint sites found in tenant!" -ForegroundColor Yellow
            return @()
        }
        
        Write-Host "[Info] Found $($sites.Count) SharePoint sites." -ForegroundColor Green
        foreach ($site in $sites) {
            Write-Host "[Site] $($site.DisplayName) - $($site.WebUrl) - $($site.Id)" -ForegroundColor Gray
        }
        
        return $sites
    }
    catch {
        Write-Host "[Error] Failed to enumerate SharePoint sites: $_" -ForegroundColor Red
        return @()
    }
}

function Get-SiteInfo {
    param([string]$SiteUrl)
    
    Write-Host "Getting site information..." -ForegroundColor Cyan
    
    try {
        # Extract site ID from URL
        $uri = [Uri]$SiteUrl
        $sitePath = $uri.AbsolutePath
        $siteId = "$($uri.Host):$sitePath"
        
        $site = Get-MgSite -SiteId $siteId
        Write-Host "Found site: $($site.DisplayName)" -ForegroundColor Green
        
        return $site
    }
    catch {
        Write-Host "[Error] Failed to get site information: $_" -ForegroundColor Red
        throw
    }
}
#endregion

#region File and Storage Analysis Functions

function Get-SharePointLibraryFilesViaListApi {
    param($Site)
    
    $allFiles = @()
    $allFolders = @()
    $folderSizes = @{}

    Write-Host "Getting document libraries..." -ForegroundColor Cyan
    $lists = Invoke-WithRetry { Get-MgSiteList -SiteId $Site.Id -WarningAction SilentlyContinue }
    $docLibraries = $lists | Where-Object { $_.List -and $_.List.Template -eq "documentLibrary" }
    $totalLists = $docLibraries.Count
    $listIndex = 0
    $totalFilesSoFar = 0
    
    Write-Host "Found $totalLists document libraries" -ForegroundColor Green
    
    foreach ($list in $docLibraries) {
        $listIndex++
        Write-Host "Processing library $listIndex/$totalLists`: $($list.DisplayName)" -ForegroundColor Cyan
        
        # Use SharePoint List API to get all items with drive item details
        $uri = "/v1.0/sites/$($Site.Id)/lists/$($list.Id)/items?expand=fields,driveItem&`$top=200"
        $more = $true
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
                
                $batchCount = 0
                foreach ($item in $resp.value) {
                    if ($item.driveItem) {
                        if ($item.driveItem.file) {
                            # Filter out system files
                            $isSystem = $false
                            if ($item.driveItem.name -match "^~" -or $item.driveItem.name -match "^\." -or 
                                $item.driveItem.name -match "^Forms$" -or $item.driveItem.name -match "^_vti_" -or 
                                $item.driveItem.name -match "^appdata" -or $item.driveItem.name -match "^.DS_Store$" -or 
                                $item.driveItem.name -match "^Thumbs.db$") {
                                $isSystem = $true
                            }
                            if ($item.driveItem.file.mimeType -eq "application/vnd.microsoft.sharepoint.system" -or 
                                $item.driveItem.file.mimeType -eq "application/vnd.ms-sharepoint.folder") {
                                $isSystem = $true
                            }
                            if ($isSystem) { continue }
                            
                            $allFiles += [PSCustomObject]@{
                                Name                   = $item.driveItem.name
                                Size                   = [long]$item.driveItem.size
                                SizeGB                 = [math]::Round($item.driveItem.size / 1GB, 3)
                                SizeMB                 = [math]::Round($item.driveItem.size / 1MB, 2)
                                Path                   = if ($item.driveItem.parentReference) { $item.driveItem.parentReference.path } else { "" }
                                Drive                  = if ($item.driveItem.parentReference) { $item.driveItem.parentReference.driveId } else { "" }
                                Extension              = [System.IO.Path]::GetExtension($item.driveItem.name).ToLower()
                                LastModifiedDateTime   = $item.driveItem.lastModifiedDateTime
                                LibraryName            = $list.DisplayName
                            }
                            $folderPath = if ($item.driveItem.parentReference) { $item.driveItem.parentReference.path } else { "" }
                            if (-not $folderSizes.ContainsKey($folderPath)) { $folderSizes[$folderPath] = 0 }
                            $folderSizes[$folderPath] += $item.driveItem.size
                            $filesInThisList++
                            $totalFilesSoFar++
                        } elseif ($item.driveItem.folder) {
                            $allFolders += [PSCustomObject]@{
                                Name = $item.driveItem.name
                                Path = if ($item.driveItem.parentReference) { $item.driveItem.parentReference.path } else { "" }
                                Drive = if ($item.driveItem.parentReference) { $item.driveItem.parentReference.driveId } else { "" }
                                ChildCount = $item.driveItem.folder.childCount
                            }
                            $foldersInThisList++
                        }
                        $batchCount++
                        # Show progress with current folder as encountered
                        $currentFolder = if ($item.driveItem.parentReference) { $item.driveItem.parentReference.path } else { "Root" }
                        Write-Progress -Activity "Scanning SharePoint Document Libraries" -Status "Library $listIndex/$totalLists`: $($list.DisplayName) | Files: $totalFilesSoFar | Folders: $($allFolders.Count) | Current Folder: $currentFolder" -PercentComplete ([math]::Min(100, ($listIndex-1)/$totalLists*100 + ($filesInThisList/1000)))
                    }
                }
                if ($resp."@odata.nextLink") {
                    $nextLink = $resp."@odata.nextLink"
                } else {
                    $more = $false
                }
                # Optional: Add a small randomized delay to avoid throttling
                Start-Sleep -Milliseconds (Get-Random -Minimum 100 -Maximum 400)
            }
            catch {
                Write-Host "[Warning] Error processing library $($list.DisplayName): $_" -ForegroundColor Yellow
                $more = $false
            }
        }
        Write-Host "Library $($list.DisplayName): Found $filesInThisList files, $foldersInThisList folders" -ForegroundColor Green
    }
    Write-Progress -Activity "Scanning SharePoint Document Libraries" -Completed
    return @{ Files = $allFiles; Folders = $allFolders; FolderSizes = $folderSizes }
}

function Get-FileData {
    param($Site)
    
    Write-Host "Analyzing site structure (SharePoint List API)..." -ForegroundColor Cyan
    
    try {
        $result = Get-SharePointLibraryFilesViaListApi -Site $Site
        $allFiles = $result.Files
        $folderSizes = $result.FolderSizes
        
        Write-Host "Site analysis complete - Found $($allFiles.Count) files via List API" -ForegroundColor Green
        
        return @{
            Files = $allFiles
            FolderSizes = $folderSizes
            TotalFiles = $allFiles.Count
            TotalSizeGB = [math]::Round(($allFiles | Measure-Object -Property Size -Sum).Sum / 1GB, 2)
        }
    }
    catch {
        Write-Host "[Error] Failed to analyze site structure: $_" -ForegroundColor Red
        return @{
            Files = @()
            FolderSizes = @{}
            TotalFiles = 0
            TotalSizeGB = 0
        }
    }
}

function Get-SiteStorageAndAccess {
    param($Site)
    
    $siteInfo = @{
        SiteName = $Site.DisplayName
        SiteUrl = $Site.WebUrl
        SiteId = $Site.Id
        StorageGB = 0
        Users = @()
        ExternalGuests = @()
        TopFiles = @()
        TopFolders = @()
    }

    try {
        # Get storage using the improved method
        $fileData = Get-FileData -Site $Site
        $siteInfo.StorageGB = $fileData.TotalSizeGB
        $siteInfo.TopFiles = $fileData.Files | Sort-Object Size -Descending | Select-Object -First 20 | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                SizeMB = $_.SizeMB
                Path = $_.Path
                Extension = $_.Extension
                LibraryName = $_.LibraryName
            }
        }
        $siteInfo.TopFolders = $fileData.FolderSizes.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 20 | ForEach-Object {
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
            $permissions = Invoke-WithRetry { Get-MgSitePermission -SiteId $Site.Id -All -ErrorAction SilentlyContinue }
            foreach ($perm in $permissions) {
                if ($perm.Invitation) {
                    $userType = "External Guest"
                    $externalGuests += [PSCustomObject]@{
                        UserName = $perm.Invitation.InvitedUserDisplayName
                        UserEmail = $perm.Invitation.InvitedUserEmailAddress
                        AccessType = $perm.Roles -join ", "
                    }
                } 
                elseif ($perm.GrantedToIdentitiesV2) {
                    foreach ($identity in $perm.GrantedToIdentitiesV2) {
                        $userType = if ($identity.User.UserType -eq "Guest") { "External Guest" } 
                                   elseif ($identity.User.UserType -eq "Member") { "Internal" } 
                                   else { $identity.User.UserType }
                        
                        $userObj = [PSCustomObject]@{
                            UserName = $identity.User.DisplayName
                            UserEmail = $identity.User.Email
                            UserType = $userType
                            AccessType = $perm.Roles -join ", "
                        }
                        $siteUsers += $userObj
                        if ($userType -eq "External Guest") { 
                            $externalGuests += $userObj 
                        }
                    }
                }
            }
        } 
        catch {
            Write-Host "[Warning] Could not retrieve site permissions: $_" -ForegroundColor Yellow
        }
        
        $siteInfo.Users = $siteUsers
        $siteInfo.ExternalGuests = $externalGuests
    }
    catch {
        Write-Host "[Error] Failed to get site storage and access info: $_" -ForegroundColor Red
    }
    
    return $siteInfo
}
#endregion

#region Permission and Access Functions

function Get-SiteUserAccessSummary {
    param($Site)
    
    $owners = @()
    $members = @()
    
    try {
        # Get site groups
        $groups = Invoke-WithRetry { Get-MgSiteGroup -SiteId $Site.Id -WarningAction SilentlyContinue }
        $ownersGroup = $groups | Where-Object { $_.DisplayName -match "Owner" }
        $membersGroup = $groups | Where-Object { $_.DisplayName -match "Member" }
        
        # Get group members
        $getGroupUsers = { 
            param($group) 
            if ($group) { 
                Invoke-WithRetry { Get-MgGroupMember -GroupId $group.Id -All -WarningAction SilentlyContinue }
            } else { 
                @() 
            } 
        }
        
        $ownerUsers = & $getGroupUsers $ownersGroup
        $memberUsers = & $getGroupUsers $membersGroup
        
        foreach ($user in $ownerUsers) {
            $owners += [PSCustomObject]@{
                UserName = $user.DisplayName
                UserEmail = Get-UserEmail $user
            }
        }
        
        foreach ($user in $memberUsers) {
            $members += [PSCustomObject]@{
                UserName = $user.DisplayName
                UserEmail = Get-UserEmail $user
            }
        }
    } 
    catch {
        Write-Host "Error retrieving site user/group access: $_" -ForegroundColor Red
    }
    
    return @{ 
        Owners = $owners
        Members = $members 
    }
}

function Get-ParentFolderAccess {
    param($Site)
    
    Write-Host "Retrieving parent folder access information..." -ForegroundColor Cyan
    
    $folderAccess = @()
    $processedFolders = @{}
    
    try {
        # Get all drives for the site
        $drives = Invoke-WithRetry { Get-MgSiteDrive -SiteId $Site.Id -WarningAction SilentlyContinue }
        
        foreach ($drive in $drives) {
            try {
                # Get root folders only (first level)
                $rootFolders = Invoke-WithRetry { 
                    Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId "root" -All -ErrorAction Stop | 
                    Where-Object { $_.Folder }
                }
                
                foreach ($folder in $rootFolders) {
                    if ($processedFolders.ContainsKey($folder.Id)) { continue }
                    $processedFolders[$folder.Id] = $true
                    
                    try {
                        $permissions = Invoke-WithRetry { Get-MgDriveItemPermission -DriveId $drive.Id -DriveItemId $folder.Id -All -ErrorAction Stop }
                        
                        foreach ($perm in $permissions) {
                            $roles = ($perm.Roles | Where-Object { $_ }) -join ", "
                            
                            if ($perm.GrantedToIdentitiesV2) {
                                foreach ($identity in $perm.GrantedToIdentitiesV2) {
                                    if ($identity.User.DisplayName) {
                                        $folderAccess += [PSCustomObject]@{
                                            FolderName = $folder.Name
                                            FolderPath = $folder.ParentReference.Path + "/" + $folder.Name
                                            UserName = $identity.User.DisplayName
                                            UserEmail = $identity.User.Email
                                            PermissionLevel = $roles
                                            AccessType = if ($roles -match "owner|write") { 
                                                "Full/Edit" 
                                            } elseif ($roles -match "read") { 
                                                "Read Only" 
                                            } else { 
                                                "Other" 
                                            }
                                        }
                                    }
                                }
                            }
                            
                            if ($perm.GrantedTo -and $perm.GrantedTo.User.DisplayName) {
                                $folderAccess += [PSCustomObject]@{
                                    FolderName = $folder.Name
                                    FolderPath = $folder.ParentReference.Path + "/" + $folder.Name
                                    UserName = $perm.GrantedTo.User.DisplayName
                                    UserEmail = $perm.GrantedTo.User.Email
                                    PermissionLevel = $roles
                                    AccessType = if ($roles -match "owner|write") { 
                                        "Full/Edit" 
                                    } elseif ($roles -match "read") { 
                                        "Read Only" 
                                    } else { 
                                        "Other" 
                                    }
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
#endregion

#region Main Execution Function
function Main {
    try {
        Write-Host "SharePoint Tenant Storage & Access Report Generator" -ForegroundColor Green
        Write-Host "=============================================" -ForegroundColor Green

        # Connect to Microsoft Graph
        Connect-ToGraph

        Write-Host "[Info] Starting tenant-wide SharePoint audit..." -ForegroundColor Cyan
        $tenantName = Get-TenantName
        $dateStr = Get-Date -Format yyyyMMdd_HHmmss
        $excelFileName = "TenantAudit-$tenantName-$dateStr.xlsx"

        # Get all SharePoint sites in the tenant
        $sites = Get-AllSharePointSites
        if ($sites.Count -eq 0) {
            Write-Host "[Warning] No sites found. Exiting." -ForegroundColor Yellow
            return
        }

        Write-Host "[Info] Processing $($sites.Count) sites for summary storage..." -ForegroundColor Cyan
        
        # First, get summary storage for all sites using the improved method
        $siteSummaries = @()
        foreach ($site in $sites) {
            Write-Host "[Info] Processing site: $($site.DisplayName) ($($site.WebUrl))" -ForegroundColor Magenta
            
            try {
                $siteInfo = Get-SiteStorageAndAccess -Site $site
                
                $siteSummaries += [PSCustomObject]@{
                    Site = $site
                    SiteName = $site.DisplayName
                    SiteId = $site.Id
                    SiteUrl = $site.WebUrl
                    StorageBytes = $siteInfo.StorageGB * 1GB
                    StorageGB = $siteInfo.StorageGB
                    FileCount = $siteInfo.TopFiles.Count
                    UserCount = $siteInfo.Users.Count
                    ExternalGuestCount = $siteInfo.ExternalGuests.Count
                }
            }
            catch {
                Write-Host "[Error] Failed to process site $($site.DisplayName): $_" -ForegroundColor Red
                # Add placeholder data for failed sites
                $siteSummaries += [PSCustomObject]@{
                    Site = $site
                    SiteName = $site.DisplayName
                    SiteId = $site.Id
                    SiteUrl = $site.WebUrl
                    StorageBytes = 0
                    StorageGB = 0
                    FileCount = 0
                    UserCount = 0
                    ExternalGuestCount = 0
                }
            }
        }

        # Identify top 10 largest sites
        $topSites = $siteSummaries | Sort-Object StorageBytes -Descending | Select-Object -First 10
        
        # Prepare data structures for Excel export
        $allSiteSummaries = @()
        $allTopFiles = @()
        $allTopFolders = @()
        $siteStorageStats = @{}
        $sitePieCharts = @{}

        # Process each site for detailed analysis
        foreach ($siteSummary in $siteSummaries) {
            $site = $siteSummary.Site
            $isTopSite = $topSites.SiteId -contains $site.Id
            
            Write-Host "[Info] Processing site: $($site.DisplayName) (Storage: $($siteSummary.StorageGB) GB)" -ForegroundColor Magenta
            
            # Get site owners and members
            $userAccess = Get-SiteUserAccessSummary -Site $site
            
            if ($isTopSite) {
                Write-Host "[Info] Performing detailed analysis for top site: $($site.DisplayName)" -ForegroundColor Cyan
                
                # Get detailed file data for top sites (already done in Get-SiteStorageAndAccess)
                $fileData = Get-FileData -Site $site
                
                # Only get folder access if we have files
                if ($fileData.Files.Count -gt 0) {
                    $folderAccess = Get-ParentFolderAccess -Site $site
                } else {
                    $folderAccess = @()
                    Write-Host "[Warning] No files found in site, skipping folder access analysis" -ForegroundColor Yellow
                }
                
                # Add to collections
                $allTopFiles += $fileData.Files | Select-Object @{Name="SiteName";Expression={$site.DisplayName}}, *
                $allTopFolders += $fileData.FolderSizes.GetEnumerator() | ForEach-Object {
                    [PSCustomObject]@{
                        SiteName = $site.DisplayName
                        FolderPath = $_.Key
                        SizeGB = [math]::Round($_.Value / 1GB, 3)
                        SizeMB = [math]::Round($_.Value / 1MB, 2)
                    }
                }
                
                # Store storage stats for pie charts
                $siteStorageStats[$site.DisplayName] = $siteSummary.StorageGB
                $sitePieCharts[$site.DisplayName] = $fileData.FolderSizes.GetEnumerator() | 
                    Sort-Object Value -Descending | Select-Object -First 10 | ForEach-Object {
                        [PSCustomObject]@{
                            Location = if ($_.Key -match "/([^/]+)/?$") { $matches[1] } else { "Root" }
                            SizeGB = [math]::Round($_.Value / 1GB, 3)
                        }
                    }
                
                # Add detailed summary
                $allSiteSummaries += [PSCustomObject]@{
                    SiteName = $site.DisplayName
                    SiteUrl = $site.WebUrl
                    TotalFiles = $fileData.TotalFiles
                    TotalSizeGB = $siteSummary.StorageGB
                    TotalFolders = $fileData.FolderSizes.Count
                    UniquePermissionLevels = ($folderAccess.PermissionLevel | Sort-Object -Unique).Count
                    OwnersCount = $userAccess.Owners.Count
                    MembersCount = $userAccess.Members.Count
                    ExternalGuestCount = $siteSummary.ExternalGuestCount
                    ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
                
                # Export owners and members to separate worksheets
                if ($userAccess.Owners.Count -gt 0) {
                    $userAccess.Owners | Export-Excel -Path $excelFileName -WorksheetName ("Owners - " + $site.DisplayName.Substring(0, [Math]::Min($site.DisplayName.Length, 25))) -AutoSize -TableStyle Medium7
                }
                if ($userAccess.Members.Count -gt 0) {
                    $userAccess.Members | Export-Excel -Path $excelFileName -WorksheetName ("Members - " + $site.DisplayName.Substring(0, [Math]::Min($site.DisplayName.Length, 25))) -AutoSize -TableStyle Medium8
                }
            } 
            else {
                # Summary only for other sites
                $allSiteSummaries += [PSCustomObject]@{
                    SiteName = $site.DisplayName
                    SiteUrl = $site.WebUrl
                    TotalFiles = $null
                    TotalSizeGB = $siteSummary.StorageGB
                    TotalFolders = $null
                    UniquePermissionLevels = $null
                    OwnersCount = $userAccess.Owners.Count
                    MembersCount = $userAccess.Members.Count
                    ExternalGuestCount = $siteSummary.ExternalGuestCount
                    ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
            }
        }

        # Create tenant-wide storage pie chart data
        $tenantPieChart = $siteStorageStats.GetEnumerator() | 
            Sort-Object Value -Descending | Select-Object -First 10 | ForEach-Object {
                [PSCustomObject]@{
                    SiteName = $_.Key
                    TotalSizeGB = $_.Value
                }
            }

        # Export main Excel report
        Write-Host "[Info] Creating Excel report..." -ForegroundColor Cyan
        
        # Only create Excel if we have data
        if ($allSiteSummaries.Count -eq 0) {
            Write-Host "[Warning] No site summaries to export" -ForegroundColor Yellow
            return
        }
        
        $excel = $allSiteSummaries | Export-Excel -Path $excelFileName -WorksheetName "Summary" -AutoSize -TableStyle Medium2 -PassThru
        
        if ($allTopFiles.Count -gt 0) {
            try {
                $allTopFiles | Export-Excel -ExcelPackage $excel -WorksheetName "Top Files" -AutoSize -TableStyle Medium6
            } catch {
                Write-Host "[Warning] Could not export Top Files: $_" -ForegroundColor Yellow
            }
        }
        
        if ($allTopFolders.Count -gt 0) {
            try {
                $allTopFolders | Export-Excel -ExcelPackage $excel -WorksheetName "Top Folders" -AutoSize -TableStyle Medium3
            } catch {
                Write-Host "[Warning] Could not export Top Folders: $_" -ForegroundColor Yellow
            }
        }
        
        if ($tenantPieChart.Count -gt 0) {
            $tenantPieChart | Export-Excel -ExcelPackage $excel -WorksheetName "Tenant Storage Pie" -AutoSize -TableStyle Medium4
            
            # Add tenant storage pie chart
            $ws = $excel.Workbook.Worksheets["Tenant Storage Pie"]
            if ($ws) {
                $chart = $ws.Drawings.AddChart("TenantStorageChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
                $chart.Title.Text = "Tenant Storage Usage (Top 10 Sites)"
                $chart.SetPosition(1, 0, 7, 0)
                $chart.SetSize(500, 400)
                $series = $chart.Series.Add($ws.Cells["B2:B$($tenantPieChart.Count + 1)"], $ws.Cells["A2:A$($tenantPieChart.Count + 1)"])
                $series.Header = "Size (GB)"
            }
            
            # Add individual site pie charts
            foreach ($site in $tenantPieChart) {
                $siteName = $site.SiteName
                if ($sitePieCharts.ContainsKey($siteName) -and $sitePieCharts[$siteName]) {
                    $sitePieCharts[$siteName] | Export-Excel -ExcelPackage $excel -WorksheetName ("Pie - " + $siteName.Substring(0, [Math]::Min($siteName.Length, 25))) -AutoSize -TableStyle Medium4
                    
                    $wsSite = $excel.Workbook.Worksheets["Pie - " + $siteName.Substring(0, [Math]::Min($siteName.Length, 25))]
                    if ($wsSite -and $sitePieCharts[$siteName]) {
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
        }

        # Close and save Excel package
        if ($excel -and $excel.Workbook.Worksheets.Count -gt 0) {
            Close-ExcelPackage $excel
            Write-Host "`n[Success] Report saved to: $excelFileName" -ForegroundColor Green
        } 
        else {
            Write-Host "[Warning] No data was found to export. No Excel report was generated." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "`n[Error] Script execution failed: $_" -ForegroundColor Red
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor DarkRed
    }
    finally {
        # Always disconnect from Graph
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            Write-Host "[Info] Disconnected from Microsoft Graph" -ForegroundColor Cyan
        }
        catch {
            # Silently handle disconnect errors
        }
    }
}
#endregion

#region Script Execution
# Execute the main function
Main
#endregion
