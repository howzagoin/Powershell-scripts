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
#>
#region Configuration and Prerequisites
# Set strict error handling
$ErrorActionPreference = "Continue"
$WarningPreference = "Continue"
# Configuration
$clientId              = '278b9af9-888d-4344-93bb-769bdd739249'
$tenantId              = 'ca0711e2-e703-4f4e-9099-17d97863211c'
$certificateThumbprint = '2E2502BB1EDB8F36CF9DE50936B283BDD22D5BAD'

# Function to check and install/update specific modules
function Install-OrUpdateModule {
    param(
        [string]$ModuleName,
        [string]$MinimumVersion,
        [switch]$Force
    )
    
    $installedModule = Get-Module -Name $ModuleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    
    if (-not $installedModule) {
        Write-Host "[Info] Installing module: $ModuleName (minimum version: $MinimumVersion)" -ForegroundColor Cyan
        Install-Module -Name $ModuleName -MinimumVersion $MinimumVersion -Scope CurrentUser -Force -AllowClobber
        return $true
    }
    elseif ($Force -or [version]$installedModule.Version -lt [version]$MinimumVersion) {
        Write-Host "[Info] Updating module: $ModuleName from $($installedModule.Version) to $MinimumVersion" -ForegroundColor Cyan
        # First, try to uninstall all versions
        try {
            Get-Module -Name $ModuleName -All | Remove-Module -Force
            Get-InstalledModule -Name $ModuleName -AllVersions | Uninstall-Module -Force -ErrorAction SilentlyContinue
        } catch {
            Write-Host "[Warning] Could not uninstall all versions of $ModuleName. Trying to install anyway." -ForegroundColor Yellow
        }
        
        # Then install the new version
        Install-Module -Name $ModuleName -MinimumVersion $MinimumVersion -Scope CurrentUser -Force -AllowClobber
        return $true
    }
    
    return $false
}

# Function to remove conflicting modules
function Remove-ConflictingModules {
    param(
        [string[]]$ModulePatterns
    )
    
    foreach ($pattern in $ModulePatterns) {
        $loadedModules = Get-Module | Where-Object { $_.Name -like $pattern }
        foreach ($module in $loadedModules) {
            Write-Host "[Info] Removing conflicting module: $($module.Name) v$($module.Version)" -ForegroundColor Yellow
            Remove-Module -Name $module.Name -Force -ErrorAction SilentlyContinue
        }
    }
}

# Function to check if a restart is needed
function Test-RestartNeeded {
    $loadedModules = Get-Module | Where-Object { $_.Name -like "Microsoft.Graph*" }
    $moduleVersions = @{}
    
    foreach ($module in $loadedModules) {
        if (-not $moduleVersions.ContainsKey($module.Name)) {
            $moduleVersions[$module.Name] = @()
        }
        $moduleVersions[$module.Name] += $module.Version
    }
    
    # Check if any module has multiple versions loaded
    foreach ($moduleName in $moduleVersions.Keys) {
        if ($moduleVersions[$moduleName].Count -gt 1) {
            Write-Host "[Warning] Multiple versions of $moduleName are loaded: $($moduleVersions[$moduleName] -join ', ')" -ForegroundColor Yellow
            return $true
        }
    }
    
    return $false
}

# Check for and install only required modules
$modulesToInstall = @(
    @{ Name = "Microsoft.Graph.Authentication"; MinimumVersion = "2.0.0" },
    @{ Name = "Microsoft.Graph.Sites"; MinimumVersion = "2.0.0" },
    @{ Name = "Microsoft.Graph.Files"; MinimumVersion = "2.0.0" },
    @{ Name = "Microsoft.Graph.Users"; MinimumVersion = "2.0.0" },
    @{ Name = "ImportExcel"; MinimumVersion = "7.0.0" }
)

$moduleUpdated = $false
foreach ($module in $modulesToInstall) {
    if (Install-OrUpdateModule -ModuleName $module.Name -MinimumVersion $module.MinimumVersion -Force:$false) {
        $moduleUpdated = $true
    }
}

# If modules were updated, we need to restart the PowerShell session
if ($moduleUpdated) {
    Write-Host "[Warning] Modules were updated. Please restart PowerShell and run the script again." -ForegroundColor Yellow
    Write-Host "[Info] This ensures the new module versions are properly loaded." -ForegroundColor Cyan
    return
}

# Remove any conflicting loaded modules
Remove-ConflictingModules -ModulePatterns @("Microsoft.Graph*")

# Check if restart is needed due to version conflicts
if (Test-RestartNeeded) {
    Write-Host "[Warning] Multiple versions of Microsoft Graph modules are loaded. Please restart PowerShell and run the script again." -ForegroundColor Yellow
    return
}

# Import required modules with specific error handling
$importErrorList = @()

try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Write-Host "[Success] Loaded Microsoft.Graph.Authentication module" -ForegroundColor Green
} catch {
    $importErrorList += "Microsoft.Graph.Authentication: $_"
}

try {
    Import-Module Microsoft.Graph.Sites -ErrorAction Stop
    Write-Host "[Success] Loaded Microsoft.Graph.Sites module" -ForegroundColor Green
} catch {
    $importErrorList += "Microsoft.Graph.Sites: $_"
}

try {
    Import-Module Microsoft.Graph.Files -ErrorAction Stop
    Write-Host "[Success] Loaded Microsoft.Graph.Files module" -ForegroundColor Green
} catch {
    $importErrorList += "Microsoft.Graph.Files: $_"
}

try {
    Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
    Write-Host "[Success] Loaded Microsoft.Graph.Users module" -ForegroundColor Green
} catch {
    Write-Host "[Warning] Microsoft.Graph.Users module not available: $_" -ForegroundColor Yellow
    $importErrorList += "Microsoft.Graph.Users: $_"
}

try {
    Import-Module ImportExcel -ErrorAction Stop
    Write-Host "[Success] Loaded ImportExcel module" -ForegroundColor Green
} catch {
    $importErrorList += "ImportExcel: $_"
}

# If there were import errors, try to force update the problematic modules
if ($importErrorList.Count -gt 0) {
    Write-Host "[Warning] Some modules failed to load. Attempting to force update..." -ForegroundColor Yellow
    foreach ($moduleError in $importErrorList) {
        $moduleName = $moduleError.Split(':')[0]
        Write-Host "[Info] Force updating module: $moduleName" -ForegroundColor Cyan
        Install-OrUpdateModule -ModuleName $moduleName -MinimumVersion "2.0.0" -Force
    }
    
    Write-Host "[Warning] Modules were force updated. Please restart PowerShell and run the script again." -ForegroundColor Yellow
    return
}
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
        $sites = @()
        
        # Approach 1: Get root site
        try {
            $rootSite = Get-MgSite -SiteId "root" -ErrorAction SilentlyContinue
            if ($rootSite) {
                $sites += $rootSite
                Write-Host "[Site] Root site found: $($rootSite.DisplayName)" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "[Warning] Could not get root site: $_" -ForegroundColor Yellow
        }
        
        # Approach 2: Get all sites using Graph API
        try {
            $allSites = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites?`$top=500" -ErrorAction SilentlyContinue
            if ($allSites.value) {
                $sites += $allSites.value
                Write-Host "[Info] Found $($allSites.value.Count) sites via Graph API"
            }
        }
        catch {
            Write-Host "[Warning] Graph API site search failed: $_" -ForegroundColor Yellow
        }
        
        # Approach 3: Search for site collections
        try {
            $searchSites = Get-MgSite -Search "*" -All -ErrorAction SilentlyContinue
            if ($searchSites) {
                $sites += $searchSites
                Write-Host "[Info] Found $($searchSites.Count) sites via search"
            }
        }
        catch {
            Write-Host "[Warning] Site search failed: $_" -ForegroundColor Yellow
        }
        
        # Remove duplicates and ensure we have valid sites
        $sites = $sites | Where-Object { $_ -and $_.Id -and $_.DisplayName } | Sort-Object Id -Unique
        
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
        
        $site = Get-MgSite -SiteId $siteId -ErrorAction Stop
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
        Write-Host "[Info] Processing site: $($Site.DisplayName)" -ForegroundColor Cyan
        
        # Get all document libraries
        $lists = Invoke-WithRetry { Get-MgSiteList -SiteId $Site.Id -WarningAction SilentlyContinue }
        $docLibraries = $lists | Where-Object { $_.List -and $_.List.Template -eq "documentLibrary" }
        
        if (-not $docLibraries -or $docLibraries.Count -eq 0) {
            Write-Host "[Warning] No document libraries found for site: $($Site.DisplayName)" -ForegroundColor Yellow
            return @{
                Files = @()
                FolderSizes = @{}
                TotalFiles = 0
                TotalSizeGB = 0
            }
        }
        
        Write-Host "[Info] Found $($docLibraries.Count) document libraries in site: $($Site.DisplayName)" -ForegroundColor Green
        
        $allFiles = @()
        $folderSizes = @{}
        $totalFiles = 0
        $listIndex = 0
        
        foreach ($list in $docLibraries) {
            $listIndex++
            $percentComplete = [math]::Round(($listIndex / $docLibraries.Count) * 100, 1)
            
            # Progress bar for library processing
            Write-Progress -Activity "Analyzing Document Libraries" -Status "Processing: $($list.DisplayName) | Files found: $totalFiles" -PercentComplete $percentComplete -CurrentOperation "$listIndex of $($docLibraries.Count) libraries"
            
            try {
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
                        
                        if (-not $resp.value -or $resp.value.Count -eq 0) {
                            $more = $false
                            continue
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
                                
                                # Update progress bar
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
                        } else {
                            $more = $false
                        }
                        
                        # Add small delay to avoid throttling
                        Start-Sleep -Milliseconds (Get-Random -Minimum 100 -Maximum 300)
                    }
                    catch {
                        Write-Host "[Error] Failed to process list items: $_" -ForegroundColor Red
                        $more = $false
                    }
                }
                
                Write-Host "[Info] Processed library: $($list.DisplayName) - Found $filesInThisList files" -ForegroundColor Green
            }
            catch {
                Write-Host "[Error] Failed to process library $($list.DisplayName): $_" -ForegroundColor Red
            }
        }
        
        Write-Progress -Activity "Analyzing Document Libraries" -Completed
        
        $result = @{
            Files = $allFiles
            FolderSizes = $folderSizes
            TotalFiles = $allFiles.Count
            TotalSizeGB = [math]::Round(($allFiles | Measure-Object -Property Size -Sum).Sum / 1GB, 2)
        }
        
        Write-Host "[Info] Completed site: $($Site.DisplayName) - Files: $($result.TotalFiles), Size: $($result.TotalSizeGB)GB" -ForegroundColor Green
        
        return $result
    }
    catch {
        Write-Host "[Error] Failed to get file data for site $($Site.DisplayName): $_" -ForegroundColor Red
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
        Write-Host "[Info] Getting storage and access for site: $($Site.DisplayName)" -ForegroundColor Cyan
        
        # Get drives and storage
        $drives = Get-MgSiteDrive -SiteId $Site.Id -WarningAction SilentlyContinue
        $allFiles = @()
        $folderSizes = @{}
        
        foreach ($drive in $drives) {
            try {
                $items = Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId "root" -All -ErrorAction SilentlyContinue
                if ($items) {
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
            } 
            catch {
                Write-Host "[Warning] Could not access drive $($drive.Name): $_" -ForegroundColor Yellow
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
                try {
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
                catch {
                    Write-Host "[Warning] Could not process permission: $_" -ForegroundColor Yellow
                }
            }
        } 
        catch {
            Write-Host "[Warning] Could not get permissions for site $($Site.DisplayName): $_" -ForegroundColor Yellow
        }
        
        $siteInfo.Users = $siteUsers
        $siteInfo.ExternalGuests = $externalGuests
        
        Write-Host "[Info] Completed storage and access for site: $($Site.DisplayName) - Users: $($siteUsers.Count), Guests: $($externalGuests.Count)" -ForegroundColor Green
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
        Write-Host "[Info] Getting user access summary for site: $($Site.DisplayName)" -ForegroundColor Cyan
        
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
                Write-Host "[Warning] Could not process permission entry: $_" -ForegroundColor Yellow
            }
        }
        
        # Remove duplicates
        $owners = $owners | Sort-Object UserEmail -Unique
        $members = $members | Sort-Object UserEmail -Unique
        
        Write-Host "[Info] Found $($owners.Count) owners and $($members.Count) members for site: $($Site.DisplayName)" -ForegroundColor Green
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
        Write-Host "[Info] Getting folder access for site: $($Site.DisplayName)" -ForegroundColor Cyan
        
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
                        Write-Host "[Warning] Could not get permissions for folder $($folder.Name): $_" -ForegroundColor Yellow
                    }
                }
            }
            catch {
                Write-Host "[Warning] Could not access drive $($drive.Name): $_" -ForegroundColor Yellow
            }
        }
        
        Write-Progress -Activity "Analyzing Folder Permissions" -Completed
        
        # Remove duplicates (same user with same access to same folder)
        $folderAccess = $folderAccess | Sort-Object FolderName, UserName, PermissionLevel -Unique
        
        Write-Host "[Info] Found $($folderAccess.Count) folder access entries for site: $($Site.DisplayName)" -ForegroundColor Green
    }
    catch {
        Write-Progress -Activity "Analyzing Folder Permissions" -Completed
        Write-Host "[Error] Failed to get folder access for site $($Site.DisplayName): $_" -ForegroundColor Red
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
                    Percentage = if ($FileData.Files.Count -gt 0) { 
                        [math]::Round(($_.Value / ($FileData.Files | Measure-Object Size -Sum).Sum) * 100, 1) 
                    } else { 0 }
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
        
        # Only export sheets if they have data
        if ($top20Files.Count -gt 0) {
            $top20Files | Export-Excel -ExcelPackage $excel -WorksheetName "Top 20 Files" -AutoSize -TableStyle Medium6
        } else {
            # Create empty sheet with message
            $emptyData = @([PSCustomObject]@{Message = "No files found in this site"})
            $emptyData | Export-Excel -ExcelPackage $excel -WorksheetName "Top 20 Files" -AutoSize
        }
        
        if ($top10Folders.Count -gt 0) {
            $top10Folders | Export-Excel -ExcelPackage $excel -WorksheetName "Top 10 Folders" -AutoSize -TableStyle Medium3
        } else {
            $emptyData = @([PSCustomObject]@{Message = "No folders found in this site"})
            $emptyData | Export-Excel -ExcelPackage $excel -WorksheetName "Top 10 Folders" -AutoSize
        }
        
        if ($storageBreakdown.Count -gt 0) {
            $storageBreakdown | Export-Excel -ExcelPackage $excel -WorksheetName "Storage Breakdown" -AutoSize -TableStyle Medium4
        } else {
            $emptyData = @([PSCustomObject]@{Message = "No storage data available"})
            $emptyData | Export-Excel -ExcelPackage $excel -WorksheetName "Storage Breakdown" -AutoSize
        }
        
        if ($FolderAccess.Count -gt 0) {
            $FolderAccess | Export-Excel -ExcelPackage $excel -WorksheetName "Folder Access" -AutoSize -TableStyle Medium5
        } else {
            $emptyData = @([PSCustomObject]@{Message = "No folder access data available"})
            $emptyData | Export-Excel -ExcelPackage $excel -WorksheetName "Folder Access" -AutoSize
        }
        
        if ($accessSummary.Count -gt 0) {
            $accessSummary | Export-Excel -ExcelPackage $excel -WorksheetName "Access Summary" -AutoSize -TableStyle Medium1
        } else {
            $emptyData = @([PSCustomObject]@{Message = "No access summary data available"})
            $emptyData | Export-Excel -ExcelPackage $excel -WorksheetName "Access Summary" -AutoSize
        }
        
        # Add charts to the storage breakdown worksheet if data exists
        if ($storageBreakdown.Count -gt 0) {
            $ws = $excel.Workbook.Worksheets["Storage Breakdown"]
            
            # Create pie chart for storage distribution by location
            $chart = $ws.Drawings.AddChart("StorageChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
            $chart.Title.Text = "Storage Usage by Location"
            $chart.SetPosition(1, 0, 7, 0)
            $chart.SetSize(500, 400)
            
            $series = $chart.Series.Add($ws.Cells["D2:D$($storageBreakdown.Count + 1)"], $ws.Cells["A2:A$($storageBreakdown.Count + 1)"])
            $series.Header = "Size (GB)"
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
        # Only process the first site for testing
        $sites = @($sites | Select-Object -First 1)
        Write-Host "[Test Mode] Only scanning site: $($sites[0].DisplayName) ($($sites[0].WebUrl))" -ForegroundColor Yellow

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
                    Write-Host "[Info] Skipping system site: $($site.DisplayName)" -ForegroundColor Yellow
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
                        Write-Host "[Warning] Could not access drive $($drive.Name): $_" -ForegroundColor Yellow
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
                Write-Host "[Error] Failed to process site $($site.DisplayName): $_" -ForegroundColor Red
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
                Write-Host "[Error] Failed to get user access for site $($site.DisplayName): $_" -ForegroundColor Red
            }
            
            if ($isTopSite) {
                # Get detailed file data for top sites (with error handling)
                $fileData = @{ Files = @(); FolderSizes = @{}; TotalFiles = 0; TotalSizeGB = 0 }
                try {
                    $fileData = Get-FileData -Site $site
                }
                catch {
                    Write-Host "[Error] Failed to get file data for site $($site.DisplayName): $_" -ForegroundColor Red
                }
                
                # Only get folder access if we have files
                $folderAccess = @()
                if ($fileData.Files.Count -gt 0) {
                    try {
                        $folderAccess = Get-ParentFolderAccess -Site $site
                    }
                    catch {
                        Write-Host "[Error] Failed to get folder access for site $($site.DisplayName): $_" -ForegroundColor Red
                    }
                }
                
                # Add to collections
                if ($fileData.Files.Count -gt 0) {
                    $allTopFiles += $fileData.Files | Select-Object @{Name='SiteName';Expression={$site.DisplayName}}, *
                }
                
                if ($fileData.FolderSizes.Count -gt 0) {
                    $allTopFolders += $fileData.FolderSizes.GetEnumerator() | ForEach-Object {
                        [PSCustomObject]@{
                            SiteName = $site.DisplayName
                            FolderPath = $_.Key
                            SizeGB = [math]::Round($_.Value / 1GB, 3)
                            SizeMB = [math]::Round($_.Value / 1MB, 2)
                        }
                    }
                }
                
                # Find large files (>100MB) for this site
                if ($fileData.Files.Count -gt 0) {
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
                }
                
                # Store storage stats for pie charts
                $siteStorageStats[$site.DisplayName] = $siteSummary.StorageGB
                if ($fileData.FolderSizes.Count -gt 0) {
                    $sitePieCharts[$site.DisplayName] = $fileData.FolderSizes.GetEnumerator() | 
                        Sort-Object Value -Descending | Select-Object -First 10 | ForEach-Object {
                            [PSCustomObject]@{
                                Location = if ($_.Key -match "/([^/]+)/?$") { $matches[1] } else { "Root" }
                                SizeGB = [math]::Round($_.Value / 1GB, 3)
                            }
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
                    UniquePermissionLevels = if ($folderAccess.Count -gt 0) { ($folderAccess.PermissionLevel | Sort-Object -Unique).Count } else { 0 }
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
                    Write-Host "[Error] Failed to get external guests for site $($site.DisplayName): $_" -ForegroundColor Red
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
                AverageStorageGB = if ($_.Count -gt 0) { [math]::Round(($_.Group | Measure-Object TotalSizeGB -Sum).Sum / $_.Count, 2) } else { 0 }
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
                Write-Host "[Error] Failed to export top files: $_" -ForegroundColor Red
            }
        } else {
            # Create empty sheet with message
            $emptyData = @([PSCustomObject]@{Message = "No files found in top sites"})
            $emptyData | Export-Excel -ExcelPackage $excel -WorksheetName "Top Files" -AutoSize
        }
        
        Write-Progress -Activity "Generating Excel Report" -Status "Exporting folders data..." -PercentComplete 40
        if ($allTopFolders.Count -gt 0) {
            try {
                $allTopFolders | Export-Excel -ExcelPackage $excel -WorksheetName "Top Folders" -AutoSize -TableStyle Medium3
            } catch {
                Write-Host "[Error] Failed to export top folders: $_" -ForegroundColor Red
            }
        } else {
            # Create empty sheet with message
            $emptyData = @([PSCustomObject]@{Message = "No folders found in top sites"})
            $emptyData | Export-Excel -ExcelPackage $excel -WorksheetName "Top Folders" -AutoSize
        }
        
        Write-Progress -Activity "Generating Excel Report" -Status "Exporting large files..." -PercentComplete 60
        if ($global:allLargeFiles.Count -gt 0) {
            try {
                $global:allLargeFiles | Sort-Object SizeMB -Descending | Export-Excel -ExcelPackage $excel -WorksheetName "Large Files (>100MB)" -AutoSize -TableStyle Medium5
            } catch {
                Write-Host "[Error] Failed to export large files: $_" -ForegroundColor Red
            }
        } else {
            # Create empty sheet with message
            $emptyData = @([PSCustomObject]@{Message = "No large files (>100MB) found"})
            $emptyData | Export-Excel -ExcelPackage $excel -WorksheetName "Large Files (>100MB)" -AutoSize
        }
        
        # Export site type breakdown
        Write-Progress -Activity "Generating Excel Report" -Status "Exporting site type analysis..." -PercentComplete 70
        if ($siteTypeBreakdown.Count -gt 0) {
            try {
                $siteTypeBreakdown | Export-Excel -ExcelPackage $excel -WorksheetName "Site Type Analysis" -AutoSize -TableStyle Medium10
            } catch {
                Write-Host "[Error] Failed to export site type analysis: $_" -ForegroundColor Red
            }
        } else {
            # Create empty sheet with message
            $emptyData = @([PSCustomObject]@{Message = "No site type data available"})
            $emptyData | Export-Excel -ExcelPackage $excel -WorksheetName "Site Type Analysis" -AutoSize
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
                        Write-Host "[Error] Failed to create tenant storage chart: $_" -ForegroundColor Red
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
                            Write-Host "[Error] Failed to create site storage chart for $siteName`: $_" -ForegroundColor Red
                        }
                    }
                }
            } catch {
                Write-Host "[Error] Failed to export tenant storage data: $_" -ForegroundColor Red
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