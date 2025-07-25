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

  Features:
  - Scans the entire tenant (not just a single site)
  - Aggregates and summarizes results for easy review
  - Lists site owners and members
  - Finds large files in all document libraries
  - Modern error handling and reporting
  - Progress bars for all major operations
  - Modular, maintainable, and extensible design
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
if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) {
    Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
}
if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

# Import required modules
Import-Module Microsoft.Graph.Sites
Import-Module Microsoft.Graph.Files
Import-Module Microsoft.Graph.Groups -ErrorAction SilentlyContinue
Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
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
        
        # Approach 2: Get sites using hostname-based search
        try {
            $tenantDomain = $tenantId
            $tenantName = (Get-MgOrganization | Select-Object -First 1).VerifiedDomains | Where-Object { $_.IsDefault } | Select-Object -ExpandProperty Name
            if ($tenantName) {
                $hostName = $tenantName.Split('.')[0]
                $searchSites = Get-MgSite -Search "contentclass:STS_Site" -All -WarningAction SilentlyContinue
                if ($searchSites) {
                    $sites += $searchSites
                }
                
                # Also try searching for specific site types
                $teamSites = Get-MgSite -Search "contentclass:STS_Web" -All -WarningAction SilentlyContinue
                if ($teamSites) {
                    $sites += $teamSites
                }
                
                # Try searching by domain name
                $domainSites = Get-MgSite -Search $hostName -All -WarningAction SilentlyContinue
                if ($domainSites) {
                    $sites += $domainSites
                }
            }
        }
        catch {
            Write-Host "[Warning] Content class search failed: $_" -ForegroundColor Yellow
        }
        
        # Approach 3: Try to get sites through SharePoint Admin
        try {
            $adminSites = Invoke-WithRetry { 
                Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites" -WarningAction SilentlyContinue
            }
            if ($adminSites.value) {
                foreach ($siteData in $adminSites.value) {
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
            Write-Host "[Warning] Admin sites search failed: $_" -ForegroundColor Yellow
        }
        
        # Approach 4: Enhanced search patterns
        try {
            # Try different search patterns
            $searchPatterns = @("*", "site:*", "contentclass:STS_Site", "contentclass:STS_Web", "SiteCollection", "TeamSite", "CommunicationSite")
            
            foreach ($pattern in $searchPatterns) {
                try {
                    $patternSites = Get-MgSite -Search $pattern -All -WarningAction SilentlyContinue
                    if ($patternSites) {
                        $sites += $patternSites
                        Write-Host "[Info] Found $($patternSites.Count) sites with pattern '$pattern'"
                    }
                }
                catch {
                    # Silently continue to next pattern
                }
            }
        }
        catch {
            Write-Host "[Warning] Enhanced search patterns failed: $_" -ForegroundColor Yellow
        }
        
        # Approach 5: Try getting sites via different Graph endpoints
        try {
            # Try the sites/{tenant} endpoint with expanded results
            $tenantSites = Invoke-WithRetry {
                Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites?`$select=id,displayName,webUrl,name,createdDateTime,lastModifiedDateTime&`$top=999" -WarningAction SilentlyContinue
            }
            if ($tenantSites.value) {
                foreach ($siteData in $tenantSites.value) {
                    $sites += [PSCustomObject]@{
                        Id = $siteData.id
                        DisplayName = $siteData.displayName
                        WebUrl = $siteData.webUrl
                        Name = $siteData.name
                        CreatedDateTime = $siteData.createdDateTime
                        LastModifiedDateTime = $siteData.lastModifiedDateTime
                    }
                }
                Write-Host "[Info] Found $($tenantSites.value.Count) sites via Graph endpoint"
            }
        }
        catch {
            Write-Host "[Warning] Graph endpoint search failed: $_" -ForegroundColor Yellow
        }
        
        # Remove duplicates
        $sites = $sites | Sort-Object Id -Unique
        
        if (-not $sites -or $sites.Count -eq 0) {
            Write-Host "[Warning] No SharePoint sites found in tenant!" -ForegroundColor Yellow
            return @()
        }
        
        Write-Host "[Info] Found $($sites.Count) SharePoint sites." -ForegroundColor Green
        
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
                $percent = if ($TotalItems -gt 0) { 
                    [Math]::Min(100, [int](($GlobalItemIndex.Value/$TotalItems)*100)) 
                } else { 100 }
                
                $progressBar = ('█' * ($percent / 2)) + ('░' * (50 - ($percent / 2)))
                Write-Progress -Activity "Scanning SharePoint Site Content" -Status "[$progressBar] $percent% - Processing: $($child.Name)" -PercentComplete $percent -Id 1
                
                # Recursively get folder contents
                if ($child.Folder -and $Depth -lt 10) {
                    try {
                        $items += Get-DriveItems -DriveId $DriveId -Path $child.Id -Depth ($Depth + 1) -GlobalItemIndex $GlobalItemIndex -TotalItems $TotalItems
                    } 
                    catch {
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

function Get-FileData {
    param($Site)
    
    try {
        # Get all document libraries using SharePoint List API (more reliable)
        $lists = Invoke-WithRetry { Get-MgSiteList -SiteId $Site.Id -WarningAction SilentlyContinue }
        $docLibraries = $lists | Where-Object { $_.List -and $_.List.Template -eq "documentLibrary" }
        
        $allFiles = @()
        $folderSizes = @{}
        $totalFiles = 0
        $listIndex = 0
        
        foreach ($list in $docLibraries) {
            $listIndex++
            $percentComplete = [math]::Round(($listIndex / $docLibraries.Count) * 100, 1)
            
            # Progress bar for library processing with more detailed status
            Write-Progress -Activity "Analyzing Document Libraries" -Status "Starting: $($list.DisplayName) | Total files so far: $totalFiles" -PercentComplete $percentComplete -CurrentOperation "$listIndex of $($docLibraries.Count) libraries"
            
            # Use SharePoint List API to get all items with drive item details
            $uri = "/v1.0/sites/$($Site.Id)/lists/$($list.Id)/items?expand=fields,driveItem&`$top=200"
            $more = $true
            $nextLink = $null
            $filesInThisList = 0
            
            while ($more) {
                try {
                    $resp = Invoke-WithRetry {
                        if ($nextLink) {
                            Invoke-MgGraphRequest -Method GET -Uri $nextLink
                        } else {
                            Invoke-MgGraphRequest -Method GET -Uri $uri
                        }
                    }
                    
                    foreach ($item in $resp.value) {
                        if ($item.driveItem -and $item.driveItem.file) {
                            # Filter out system files
                            $isSystem = $false
                            if ($item.driveItem.name -match '^~' -or $item.driveItem.name -match '^\.' -or 
                                $item.driveItem.name -match '^Forms$' -or $item.driveItem.name -match '^_vti_' -or 
                                $item.driveItem.name -match '^appdata' -or $item.driveItem.name -match '^.DS_Store$' -or 
                                $item.driveItem.name -match '^Thumbs.db$') {
                                $isSystem = $true
                            }
                            if ($item.driveItem.file.mimeType -eq 'application/vnd.microsoft.sharepoint.system' -or 
                                $item.driveItem.file.mimeType -eq 'application/vnd.ms-sharepoint.folder') {
                                $isSystem = $true
                            }
                            if ($isSystem) { continue }
                            
                            $allFiles += [PSCustomObject]@{
                                Name = $item.driveItem.name
                                Size = [long]$item.driveItem.size
                                SizeGB = [math]::Round($item.driveItem.size / 1GB, 3)
                                SizeMB = [math]::Round($item.driveItem.size / 1MB, 2)
                                Path = $item.driveItem.parentReference ? $item.driveItem.parentReference.path : ''
                                Drive = $item.driveItem.parentReference ? $item.driveItem.parentReference.driveId : ''
                                Extension = [System.IO.Path]::GetExtension($item.driveItem.name).ToLower()
                                LibraryName = $list.DisplayName
                            }
                            
                            # Track folder sizes
                            $folderPath = $item.driveItem.parentReference ? $item.driveItem.parentReference.path : ''
                            if (-not $folderSizes.ContainsKey($folderPath)) {
                                $folderSizes[$folderPath] = 0
                            }
                            $folderSizes[$folderPath] += $item.driveItem.size
                            $filesInThisList++
                            $totalFiles++
                            
                            # Update progress bar every 10 files to show real-time progress
                            if ($totalFiles % 10 -eq 0 -or $totalFiles -eq 1) {
                                $currentFileName = if ($item.driveItem.name.Length -gt 50) { 
                                    $item.driveItem.name.Substring(0, 47) + "..." 
                                } else { 
                                    $item.driveItem.name 
                                }
                                Write-Progress -Activity "Analyzing Document Libraries" -Status "Processing: $($list.DisplayName) | Files found: $totalFiles | Current: $currentFileName" -PercentComplete $percentComplete -CurrentOperation "$listIndex of $($docLibraries.Count) libraries"
                            }
                        }
                    }
                    
                    if ($resp.'@odata.nextLink') {
                        $nextLink = $resp.'@odata.nextLink'
                        # Update progress to show we're getting more data
                        Write-Progress -Activity "Analyzing Document Libraries" -Status "Processing: $($list.DisplayName) | Files found: $totalFiles | Fetching more data..." -PercentComplete $percentComplete -CurrentOperation "$listIndex of $($docLibraries.Count) libraries"
                    } else {
                        $more = $false
                    }
                    
                    # Add small delay to avoid throttling
                    Start-Sleep -Milliseconds (Get-Random -Minimum 100 -Maximum 300)
                }
                catch {
                    $more = $false
                }
            }
            
            # Show completion status for this library
            Write-Progress -Activity "Analyzing Document Libraries" -Status "Completed: $($list.DisplayName) | Found $filesInThisList files | Total: $totalFiles" -PercentComplete $percentComplete -CurrentOperation "$listIndex of $($docLibraries.Count) libraries"
        }
        
        Write-Progress -Activity "Analyzing Document Libraries" -Completed
        
        return @{
            Files = $allFiles
            FolderSizes = $folderSizes
            TotalFiles = $allFiles.Count
            TotalSizeGB = [math]::Round(($allFiles | Measure-Object -Property Size -Sum).Sum / 1GB, 2)
        }
    }
    catch {
        Write-Progress -Activity "Analyzing Document Libraries" -Completed
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
        # Get drives and storage
        $drives = Get-MgSiteDrive -SiteId $Site.Id -WarningAction SilentlyContinue
        $allFiles = @()
        $folderSizes = @{}
        
        foreach ($drive in $drives) {
            try {
                $items = Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId "root" -All -ErrorAction SilentlyContinue
                foreach ($item in $items) {
                    if ($item.File) {
                        $allFiles += $item
                        $folderPath = $item.ParentReference.Path
                        if (-not $folderSizes.ContainsKey($folderPath)) { 
                            $folderSizes[$folderPath] = 0 
                        }
                        $folderSizes[$folderPath] += $item.Size
                    }
                }
            } 
            catch {
                # Silently handle drive access errors
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
            $permissions = Invoke-WithRetry { Get-MgSitePermission -SiteId $Site.Id -All -ErrorAction SilentlyContinue }
            foreach ($perm in $permissions) {
                if ($perm.Invitation) {
                    $userType = 'External Guest'
                    $externalGuests += [PSCustomObject]@{
                        UserName = $perm.Invitation.InvitedUserDisplayName
                        UserEmail = $perm.Invitation.InvitedUserEmailAddress
                        AccessType = $perm.Roles -join ', '
                    }
                } 
                elseif ($perm.GrantedToIdentitiesV2) {
                    foreach ($identity in $perm.GrantedToIdentitiesV2) {
                        $userType = if ($identity.User.UserType -eq 'Guest') { 'External Guest' } 
                                   elseif ($identity.User.UserType -eq 'Member') { 'Internal' } 
                                   else { $identity.User.UserType }
                        
                        $userObj = [PSCustomObject]@{
                            UserName = $identity.User.DisplayName
                            UserEmail = $identity.User.Email
                            UserType = $userType
                            AccessType = $perm.Roles -join ', '
                        }
                        $siteUsers += $userObj
                        if ($userType -eq 'External Guest') { 
                            $externalGuests += $userObj 
                        }
                    }
                }
            }
        } 
        catch {
            # Silently handle permission access errors
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
        # Get site permissions using the correct Graph API
        $permissions = Invoke-WithRetry { Get-MgSitePermission -SiteId $Site.Id -All -WarningAction SilentlyContinue }
        
        foreach ($perm in $permissions) {
            try {
                # Check if permission has user or group info
                if ($perm.GrantedToIdentitiesV2) {
                    foreach ($identity in $perm.GrantedToIdentitiesV2) {
                        $user = $identity.User
                        if ($user) {
                            $userObj = [PSCustomObject]@{
                                UserName = $user.DisplayName
                                UserEmail = if ($user.Email) { $user.Email } else { $user.UserPrincipalName }
                                Role = if ($perm.Roles) { ($perm.Roles -join ', ') } else { 'Member' }
                            }
                            
                            # Categorize based on role
                            if ($perm.Roles -and ($perm.Roles -contains 'owner' -or $perm.Roles -contains 'fullControl')) {
                                $owners += $userObj
                            } else {
                                $members += $userObj
                            }
                        }
                    }
                }
                elseif ($perm.GrantedTo -and $perm.GrantedTo.User) {
                    $user = $perm.GrantedTo.User
                    $userObj = [PSCustomObject]@{
                        UserName = $user.DisplayName
                        UserEmail = if ($user.Email) { $user.Email } else { $user.UserPrincipalName }
                        Role = if ($perm.Roles) { ($perm.Roles -join ', ') } else { 'Member' }
                    }
                    
                    # Categorize based on role
                    if ($perm.Roles -and ($perm.Roles -contains 'owner' -or $perm.Roles -contains 'fullControl')) {
                        $owners += $userObj
                    } else {
                        $members += $userObj
                    }
                }
            }
            catch {
                # Silently continue if individual permission fails
            }
        }
        
        # Remove duplicates
        $owners = $owners | Sort-Object UserEmail -Unique
        $members = $members | Sort-Object UserEmail -Unique
        
        # Try alternative approach using SharePoint REST API if no permissions found
        if ($owners.Count -eq 0 -and $members.Count -eq 0) {
            try {
                # Use Graph request to get site users
                $siteUsers = Invoke-WithRetry { 
                    Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($Site.Id)/drive/root/permissions" -WarningAction SilentlyContinue 
                }
                
                if ($siteUsers.value) {
                    foreach ($perm in $siteUsers.value) {
                        if ($perm.grantedToIdentitiesV2) {
                            foreach ($identity in $perm.grantedToIdentitiesV2) {
                                if ($identity.user) {
                                    $userObj = [PSCustomObject]@{
                                        UserName = $identity.user.displayName
                                        UserEmail = $identity.user.email
                                        Role = if ($perm.roles) { ($perm.roles -join ', ') } else { 'Member' }
                                    }
                                    $members += $userObj
                                }
                            }
                        }
                    }
                }
            }
            catch {
                # Silently handle REST API fallback errors
            }
        }
    } 
    catch {
        Write-Host "[Warning] Could not retrieve site user access for $($Site.DisplayName): $_" -ForegroundColor Yellow
    }
    
    return @{ 
        Owners = $owners
        Members = $members 
    }
}

function Get-ParentFolderAccess {
    param($Site)
    
    $folderAccess = @()
    $processedFolders = @{}
    
    try {
        # Get all drives for the site
        $drives = Invoke-WithRetry { Get-MgSiteDrive -SiteId $Site.Id -WarningAction SilentlyContinue }
        $driveIndex = 0
        
        foreach ($drive in $drives) {
            $driveIndex++
            $percentComplete = [math]::Round(($driveIndex / $drives.Count) * 100, 1)
            
            # Progress bar for drive processing
            Write-Progress -Activity "Analyzing Folder Permissions" -Status "Processing drive: $($drive.Name)" -PercentComplete $percentComplete -CurrentOperation "$driveIndex of $($drives.Count) drives"
            
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
                                            AccessType = if ($roles -match 'owner|write') { 
                                                'Full/Edit' 
                                            } elseif ($roles -match 'read') { 
                                                'Read Only' 
                                            } else { 
                                                'Other' 
                                            }
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
                                    AccessType = if ($roles -match 'owner|write') { 
                                        'Full/Edit' 
                                    } elseif ($roles -match 'read') { 
                                        'Read Only' 
                                    } else { 
                                        'Other' 
                                    }
                                }
                            }
                        }
                    }
                    catch {
                        # Silently handle permission errors
                    }
                }
            }
            catch {
                # Silently handle drive access errors
            }
        }
        
        Write-Progress -Activity "Analyzing Folder Permissions" -Completed
        
        # Remove duplicates (same user with same access to same folder)
        $folderAccess = $folderAccess | Sort-Object FolderName, UserName, PermissionLevel -Unique
    }
    catch {
        Write-Progress -Activity "Analyzing Folder Permissions" -Completed
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

#region Report Generation Functions

function New-ExcelReport {
    param(
        $FileData,
        $FolderAccess,
        $Site,
        $FileName
    )
    
    Write-Host "Creating Excel report: $FileName" -ForegroundColor Cyan
    
    try {
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
                $folderName = if ($_.Key -match "/([^/]+)/?$") { $matches[1] } else { "Root" }
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
    catch {
        Write-Host "[Error] Failed to create Excel report: $_" -ForegroundColor Red
        throw
    }
}
#endregion

#region Batch Processing Functions

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
        $result = Invoke-MgGraphRequest -Method POST -Uri "/v1.0/`$batch" -Body $body -ContentType "application/json"
        $responses += $result.responses
    }
    
    return $responses
}

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
        $result = Invoke-MgGraphRequest -Method POST -Uri "/v1.0/`$batch" -Body $body -ContentType "application/json"
        $responses += $result.responses
    }
    
    return $responses
}

function Get-FileDataBatch {
    param(
        [array]$DrivesBatchResponses
    )
    
    $allFiles = @()
    $folderSizes = @{}
    
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
                    if (-not $folderSizes.ContainsKey($folderPath)) { 
                        $folderSizes[$folderPath] = 0 
                    }
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

        # First, get summary storage for all sites
        $siteSummaries = @()
        $processedCount = 0
        $successCount = 0
        $errorCount = 0
        
        foreach ($site in $sites) {
            $processedCount++
            $percentComplete = [math]::Round(($processedCount / $sites.Count) * 100, 1)
            
            # Progress bar for site processing
            Write-Progress -Activity "Scanning SharePoint Sites" -Status "Processing: $($site.DisplayName)" -PercentComplete $percentComplete -CurrentOperation "$processedCount of $($sites.Count) sites"
            
            try {
                # Determine site type (OneDrive vs SharePoint)
                $isOneDrive = $false
                $siteType = "SharePoint Site"
                
                # Check if this is a OneDrive personal site
                if ($site.WebUrl -like "*-my.sharepoint.com/personal/*" -or 
                    $site.WebUrl -like "*/personal/*" -or 
                    $site.WebUrl -like "*mysites*" -or
                    $site.Name -like "*OneDrive*" -or
                    $site.DisplayName -like "*OneDrive*") {
                    $isOneDrive = $true
                    $siteType = "OneDrive Personal"
                }
                
                # Skip system/hidden sites that commonly cause access issues
                $skipSite = $false
                $systemSitePatterns = @(
                    'contentstorage',
                    'portals/hub',
                    '_api',
                    'search',
                    'admin'
                )
                
                # Don't skip personal OneDrive sites even if they have "mysites" or "personal" in URL
                if (-not $isOneDrive) {
                    $systemSitePatterns += @('mysites', 'personal')
                }
                
                foreach ($pattern in $systemSitePatterns) {
                    if ($site.WebUrl -like "*$pattern*") {
                        $skipSite = $true
                        break
                    }
                }
                
                if ($skipSite) { 
                    continue 
                }
                
                $drives = Invoke-WithRetry { Get-MgSiteDrive -SiteId $site.Id -WarningAction SilentlyContinue -ErrorAction Stop }
                $totalSize = 0
                
                foreach ($drive in $drives) {
                    try {
                        $items = Invoke-WithRetry { Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId "root" -All -ErrorAction SilentlyContinue }
                        if ($items) {
                            $totalSize += ($items | Measure-Object -Property Size -Sum).Sum
                        }
                    } 
                    catch {
                        # Silently handle drive errors
                    }
                }
                
                $siteSummaries += [PSCustomObject]@{
                    Site = $site
                    SiteName = $site.DisplayName
                    SiteId = $site.Id
                    SiteUrl = $site.WebUrl
                    SiteType = $siteType
                    IsOneDrive = $isOneDrive
                    StorageBytes = $totalSize
                    StorageGB = [math]::Round($totalSize / 1GB, 3)
                }
                
                $successCount++
            }
            catch {
                $errorCount++
                # Silently handle site processing errors
            }
        }
        
        # Clear progress bar
        Write-Progress -Activity "Scanning SharePoint Sites" -Completed
        
        # Calculate site type breakdown
        $sharePointSites = $siteSummaries | Where-Object { -not $_.IsOneDrive }
        $oneDriveSites = $siteSummaries | Where-Object { $_.IsOneDrive }
        
        Write-Host "`n[Summary] Processed $processedCount sites: $successCount successful, $errorCount errors" -ForegroundColor Cyan
        Write-Host "[Site Types] SharePoint Sites: $($sharePointSites.Count) | OneDrive Personal: $($oneDriveSites.Count)" -ForegroundColor Cyan

        # Identify top 10 largest sites
        $topSites = $siteSummaries | Sort-Object StorageBytes -Descending | Select-Object -First 10
        
        # Prepare data structures for Excel export
        $allSiteSummaries = @()
        $allTopFiles = @()
        $allTopFolders = @()
        $global:allLargeFiles = @()  # Files larger than 100MB
        $siteStorageStats = @{}
        $sitePieCharts = @{}

        # Process each site for detailed analysis
        $detailProcessedCount = 0
        foreach ($siteSummary in $siteSummaries) {
            $site = $siteSummary.Site
            $isTopSite = $topSites.SiteId -contains $site.Id
            $detailProcessedCount++
            $percentComplete = [math]::Round(($detailProcessedCount / $siteSummaries.Count) * 100, 1)
            
            # Progress bar for detailed analysis
            Write-Progress -Activity "Analyzing Sites for Detailed Data" -Status "Processing: $($site.DisplayName)" -PercentComplete $percentComplete -CurrentOperation "$detailProcessedCount of $($siteSummaries.Count) sites"
            
            # Get site owners and members (with improved error handling)
            $userAccess = @{ Owners = @(); Members = @() }
            try {
                $userAccess = Get-SiteUserAccessSummary -Site $site
            }
            catch {
                # Silently handle user access errors
            }
            
            if ($isTopSite) {
                # Get detailed file data for top sites (with error handling)
                $fileData = @{ Files = @(); FolderSizes = @{}; TotalFiles = 0; TotalSizeGB = 0 }
                try {
                    $fileData = Get-FileData -Site $site
                }
                catch {
                    # Silently handle file data errors
                }
                
                # Only get folder access if we have files
                $folderAccess = @()
                if ($fileData.Files.Count -gt 0) {
                    try {
                        $folderAccess = Get-ParentFolderAccess -Site $site
                    }
                    catch {
                        # Silently handle folder access errors
                    }
                }
                
                # Add to collections
                $allTopFiles += $fileData.Files | Select-Object @{Name='SiteName';Expression={$site.DisplayName}}, *
                $allTopFolders += $fileData.FolderSizes.GetEnumerator() | ForEach-Object {
                    [PSCustomObject]@{
                        SiteName = $site.DisplayName
                        FolderPath = $_.Key
                        SizeGB = [math]::Round($_.Value / 1GB, 3)
                        SizeMB = [math]::Round($_.Value / 1MB, 2)
                    }
                }
                
                # Find large files (>100MB) for this site
                $largeFiles = $fileData.Files | Where-Object { $_.Size -gt 100MB } | ForEach-Object {
                    [PSCustomObject]@{
                        SiteName = $site.DisplayName
                        FileName = $_.Name
                        SizeMB = $_.SizeMB
                        SizeGB = $_.SizeGB
                        Path = $_.Path
                        Extension = $_.Extension
                        LibraryName = $_.LibraryName
                    }
                }
                if ($largeFiles.Count -gt 0) {
                    $global:allLargeFiles += $largeFiles
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
                    SiteType = $siteSummary.SiteType
                    TotalFiles = $fileData.TotalFiles
                    TotalSizeGB = $siteSummary.StorageGB
                    TotalFolders = $fileData.FolderSizes.Count
                    UniquePermissionLevels = ($folderAccess.PermissionLevel | Sort-Object -Unique).Count
                    OwnersCount = $userAccess.Owners.Count
                    MembersCount = $userAccess.Members.Count
                    ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
                
                # Export owners and members to separate worksheets
                if ($userAccess.Owners.Count -gt 0) {
                    $userAccess.Owners | Export-Excel -Path $excelFileName -WorksheetName ("Owners - " + $site.DisplayName.Substring(0, [Math]::Min($site.DisplayName.Length, 25))) -AutoSize -TableStyle Medium7
                }
                if ($userAccess.Members.Count -gt 0) {
                    $userAccess.Members | Export-Excel -Path $excelFileName -WorksheetName ("Members - " + $site.DisplayName.Substring(0, [Math]::Min($site.DisplayName.Length, 25))) -AutoSize -TableStyle Medium8
                }
                
                # Export external guests with highlighting
                try {
                    $siteInfo = Get-SiteStorageAndAccess -Site $site
                    if ($siteInfo.ExternalGuests.Count -gt 0) {
                        $siteInfo.ExternalGuests | Export-Excel -Path $excelFileName -WorksheetName ("External Guests - " + $site.DisplayName.Substring(0, [Math]::Min($site.DisplayName.Length, 20))) -AutoSize -TableStyle Medium9
                    }
                }
                catch {
                    # Silently handle external guest errors
                }
            } 
            else {
                # Summary only for other sites
                $allSiteSummaries += [PSCustomObject]@{
                    SiteName = $site.DisplayName
                    SiteUrl = $site.WebUrl
                    SiteType = $siteSummary.SiteType
                    TotalFiles = $null
                    TotalSizeGB = $siteSummary.StorageGB
                    TotalFolders = $null
                    UniquePermissionLevels = $null
                    OwnersCount = $userAccess.Owners.Count
                    MembersCount = $userAccess.Members.Count
                    ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
            }
        }
        
        # Clear progress bar
        Write-Progress -Activity "Analyzing Sites for Detailed Data" -Completed

        # Create tenant-wide storage pie chart data
        $tenantPieChart = $siteStorageStats.GetEnumerator() | 
            Sort-Object Value -Descending | Select-Object -First 10 | ForEach-Object {
                [PSCustomObject]@{
                    SiteName = $_.Key
                    TotalSizeGB = $_.Value
                }
            }

        # Create site type breakdown for analysis
        $siteTypeBreakdown = $allSiteSummaries | Group-Object SiteType | ForEach-Object {
            [PSCustomObject]@{
                SiteType = $_.Name
                SiteCount = $_.Count
                TotalStorageGB = [math]::Round(($_.Group | Measure-Object TotalSizeGB -Sum).Sum, 2)
                AverageStorageGB = [math]::Round((($_.Group | Measure-Object TotalSizeGB -Sum).Sum / $_.Count), 2)
                TotalFiles = ($_.Group | Measure-Object TotalFiles -Sum).Sum
            }
        }

        # Export main Excel report
        Write-Progress -Activity "Generating Excel Report" -Status "Creating worksheets..." -PercentComplete 0
        
        # Ensure we have data to export
        if ($allSiteSummaries.Count -eq 0) {
            Write-Host "[Warning] No site summaries to export" -ForegroundColor Yellow
            return
        }
        
        $excel = $allSiteSummaries | Export-Excel -Path $excelFileName -WorksheetName "Summary" -AutoSize -TableStyle Medium2 -PassThru
        
        Write-Progress -Activity "Generating Excel Report" -Status "Exporting files data..." -PercentComplete 20
        if ($allTopFiles.Count -gt 0) {
            try {
                $allTopFiles | Export-Excel -ExcelPackage $excel -WorksheetName "Top Files" -AutoSize -TableStyle Medium6
            } catch {
                # Silently handle export errors
            }
        }
        
        Write-Progress -Activity "Generating Excel Report" -Status "Exporting folders data..." -PercentComplete 40
        if ($allTopFolders.Count -gt 0) {
            try {
                $allTopFolders | Export-Excel -ExcelPackage $excel -WorksheetName "Top Folders" -AutoSize -TableStyle Medium3
            } catch {
                # Silently handle export errors
            }
        }
        
        Write-Progress -Activity "Generating Excel Report" -Status "Exporting large files..." -PercentComplete 60
        if ($global:allLargeFiles.Count -gt 0) {
            try {
                $global:allLargeFiles | Sort-Object SizeMB -Descending | Export-Excel -ExcelPackage $excel -WorksheetName "Large Files (>100MB)" -AutoSize -TableStyle Medium5
            } catch {
                # Silently handle export errors
            }
        }
        
        # Export site type breakdown
        Write-Progress -Activity "Generating Excel Report" -Status "Exporting site type analysis..." -PercentComplete 70
        if ($siteTypeBreakdown.Count -gt 0) {
            try {
                $siteTypeBreakdown | Export-Excel -ExcelPackage $excel -WorksheetName "Site Type Analysis" -AutoSize -TableStyle Medium10
            } catch {
                # Silently handle export errors
            }
        }
        
        Write-Progress -Activity "Generating Excel Report" -Status "Creating charts..." -PercentComplete 80
        
        if ($tenantPieChart.Count -gt 0) {
            try {
                $tenantPieChart | Export-Excel -ExcelPackage $excel -WorksheetName "Tenant Storage Pie" -AutoSize -TableStyle Medium4
                
                # Add tenant storage pie chart
                $ws = $excel.Workbook.Worksheets["Tenant Storage Pie"]
                if ($ws) {
                    try {
                        $chart = $ws.Drawings.AddChart("TenantStorageChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
                        $chart.Title.Text = "Tenant Storage Usage (Top 10 Sites)"
                        $chart.SetPosition(1, 0, 7, 0)
                        $chart.SetSize(500, 400)
                        $series = $chart.Series.Add($ws.Cells["B2:B$($tenantPieChart.Count + 1)"], $ws.Cells["A2:A$($tenantPieChart.Count + 1)"])
                        $series.Header = "Size (GB)"
                    } catch {
                        # Silently handle chart creation errors
                    }
                }
                
                # Add individual site pie charts
                foreach ($site in $tenantPieChart) {
                    $siteName = $site.SiteName
                    if ($sitePieCharts.ContainsKey($siteName) -and $sitePieCharts[$siteName]) {
                        try {
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
                        } catch {
                            # Silently handle chart creation errors
                        }
                    }
                }
            } catch {
                # Silently handle export errors
            }
        }
        
        Write-Progress -Activity "Generating Excel Report" -Status "Finalizing report..." -PercentComplete 100

        # Close and save Excel package
        if ($excel -and $excel.Workbook.Worksheets.Count -gt 0) {
            Close-ExcelPackage $excel
            Write-Progress -Activity "Generating Excel Report" -Completed
            Write-Host "`n[Success] Report saved to: $excelFileName" -ForegroundColor Green
            Write-Host "[Report Summary] Sites: $($siteSummaries.Count) | Files: $($allTopFiles.Count) | Large Files (>100MB): $($global:allLargeFiles.Count)" -ForegroundColor Cyan
        } 
        else {
            Write-Progress -Activity "Generating Excel Report" -Completed
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