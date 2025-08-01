<#
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
            Write-Host "[Warning] Multiple versions of $moduleName are loaded: $($moduleVersions[$moduleName] -join ", ")" -ForegroundColor Yellow
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
        
        Write-Host "[Debug] Graph context: TenantId=$($context.TenantId), AuthType=$($context.AuthType), Scopes=$($context.Scopes -join ", ")" -ForegroundColor Yellow
        
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
function Get-SafeWorksheetName {
    param([string]$Name)
    
    if (-not $Name) { return "Sheet1" }
    
    # Remove or replace invalid characters for Excel worksheet names
    # Invalid characters: [ ] : * ? / \ > < |
    $safeName = $Name -replace '[\[\]:*?/\\><|]', '_'
    
    # Trim to 31 characters (Excel limit)
    if ($safeName.Length -gt 31) {
        $safeName = $safeName.Substring(0, 31)
    }
    
    # Ensure it doesn't start or end with apostrophe
    $safeName = $safeName.Trim("'")
    
    # Cannot be empty after cleaning
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        $safeName = "Sheet1"
    }
    
    return $safeName
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
                
                $progressBar = ("#" * ($percent / 2)) + ("-" * (50 - ($percent / 2)))
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
                                $fileName = $item.driveItem.name
                                
                                # Skip system/hidden files and folders
                                $systemFilePatterns = @(
                                    "~$*",           # Office temp files
                                    ".tmp",          # Temporary files
                                    "thumbs.db",     # Windows thumbnails
                                    ".ds_store",     # macOS system files
                                    "desktop.ini",   # Windows folder settings
                                    ".git*",         # Git files
                                    ".svn*",         # SVN files
                                    "*.lnk",         # Windows shortcuts
                                    "_vti_*",        # SharePoint system folders
                                    "forms/",        # SharePoint forms
                                    "web.config",    # Configuration files
                                    "*.aspx",        # SharePoint pages
                                    "*.master"       # SharePoint master pages
                                )
                                
                                foreach ($pattern in $systemFilePatterns) {
                                    if ($fileName -like $pattern) {
                                        $isSystem = $true
                                        break
                                    }
                                }
                                
                                if ($isSystem) { continue }
                                
                                # Calculate full path length for SharePoint path limit analysis
                                $parentPath = if ($item.driveItem.parentReference) { $item.driveItem.parentReference.path } else { '' }
                                $fullPath = ($parentPath + "/" + $item.driveItem.name).Replace("//", "/")
                                $pathLength = $fullPath.Length
                                
                                $allFiles += [PSCustomObject]@{
                                    Name = $item.driveItem.name
                                    Size = [long]$item.driveItem.size
                                    SizeGB = [math]::Round($item.driveItem.size / 1GB, 3)
                                    SizeMB = [math]::Round($item.driveItem.size / 1MB, 2)
                                    Path = $parentPath
                                    Drive = if ($item.driveItem.parentReference) { $item.driveItem.parentReference.driveId } else { '' }
                                    Extension = [System.IO.Path]::GetExtension($item.driveItem.name).ToLower()
                                    LibraryName = $list.DisplayName
                                    PathLength = $pathLength
                                    FullPath = $fullPath
                                }
                                
                                # Track folder sizes
                                $folderPath = if ($item.driveItem.parentReference) { $item.driveItem.parentReference.path } else { '' }
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
                        
                        # Disable pagination for now due to PowerShell parsing issues
                        $more = $false
                        
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
                        $userType = "External Guest"
                        $externalGuests += [PSCustomObject]@{
                            UserName = $perm.Invitation.InvitedUserDisplayName
                            UserEmail = $perm.Invitation.InvitedUserEmailAddress
                            AccessType = $perm.Roles -join ", "
                        }
                    } 
                    elseif ($perm.GrantedToIdentitiesV2) {
                        foreach ($identity in $perm.GrantedToIdentitiesV2) {
                            $userType = if ($identity.User.UserType -eq "Guest") { "External Guest" } elseif ($identity.User.UserType -eq "Member") { "Internal" } else { $identity.User.UserType }
                            
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
                                Role = if ($perm.Roles) { ($perm.Roles -join ", ") } else { "Member" }
                            }
                            
                            # Categorize based on role
                            if ($perm.Roles -and ($perm.Roles -contains "owner" -or $perm.Roles -contains "fullControl")) {
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
                        Role = if ($perm.Roles) { ($perm.Roles -join ", ") } else { "Member" }
                    }
                    
                    # Categorize based on role
                    if ($perm.Roles -and ($perm.Roles -contains "owner" -or $perm.Roles -contains "fullControl")) {
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
                    Users = ($_.Group.UserName | Sort-Object -Unique) -join "; "
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
        [string]$ParentId = "root"
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
                    # Calculate full path+name length (Path + "/" + Name)
                    $fullPath = ($item.parentReference.path + "/" + $item.name).Replace("//","/")
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
#region File Dialog Functions
function Get-SaveFileDialog {
    param(
        [string]$InitialDirectory = [Environment]::GetFolderPath('Desktop'),
        [string]$Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
        [string]$DefaultFileName = "SharePointAudit.xlsx",
        [string]$Title = "Save SharePoint Audit Report"
    )
    
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.InitialDirectory = $InitialDirectory
        $SaveFileDialog.Filter = $Filter
        $SaveFileDialog.FileName = $DefaultFileName
        $SaveFileDialog.Title = $Title
        $SaveFileDialog.DefaultExt = "xlsx"
        $SaveFileDialog.AddExtension = $true
        
        $result = $SaveFileDialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $SaveFileDialog.FileName
        } else {
            return $null
        }
    }
    catch {
        Write-Host "[Warning] Could not show save dialog. Using default filename in current directory." -ForegroundColor Yellow
        return $DefaultFileName
    }
}
#endregion
#region Personal OneDrive Analysis Functions
function Get-PersonalOneDriveSites {
    param($Sites)
    
    $oneDriveSites = $Sites | Where-Object { 
        $_.WebUrl -like "*-my.sharepoint.com/personal/*" -or 
        $_.WebUrl -like "*/personal/*" -or 
        $_.WebUrl -like "*mysites*" -or
        $_.Name -like "*OneDrive*" -or
        $_.DisplayName -like "*OneDrive*"
    }
    
    return $oneDriveSites
}

function Get-PersonalOneDriveDetails {
    param($OneDriveSites)
    
    $oneDriveDetails = @()
    $processedCount = 0
    
    foreach ($site in $OneDriveSites) {
        $processedCount++
        $percentComplete = [math]::Round(($processedCount / $OneDriveSites.Count) * 100, 1)
        
        Write-Progress -Activity "Analyzing Personal OneDrive Sites" -Status "Processing: $($site.DisplayName)" -PercentComplete $percentComplete -CurrentOperation "$processedCount of $($OneDriveSites.Count) OneDrive sites"
        
        try {
            # Extract user name from URL
            $userName = "Unknown User"
            if ($site.WebUrl -match "/personal/([^/]+)") {
                $userPart = $matches[1] -replace "_", "@"
                $userName = $userPart -replace "([^@]+)@([^@]+)", '$1@$2'
            }
            
            # Get storage size
            $drives = Get-MgSiteDrive -SiteId $site.Id -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            $totalSize = 0
            $fileCount = 0
            
            foreach ($drive in $drives) {
                try {
                    $items = Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId "root" -All -ErrorAction SilentlyContinue
                    if ($items) {
                        $driveSize = ($items | Where-Object { $_.File } | Measure-Object -Property Size -Sum).Sum
                        $totalSize += $driveSize
                        $fileCount += ($items | Where-Object { $_.File }).Count
                    }
                } catch {
                    # Silent error handling for individual drives
                }
            }
            
            # Get sharing/access information
            $sharedWithUsers = @()
            try {
                $permissions = Get-MgSitePermission -SiteId $site.Id -All -ErrorAction SilentlyContinue
                foreach ($perm in $permissions) {
                    if ($perm.GrantedToIdentitiesV2) {
                        foreach ($identity in $perm.GrantedToIdentitiesV2) {
                            if ($identity.User -and $identity.User.DisplayName -ne $userName) {
                                $sharedWithUsers += [PSCustomObject]@{
                                    UserName = $identity.User.DisplayName
                                    UserEmail = $identity.User.Email
                                    AccessType = ($perm.Roles -join ", ")
                                }
                            }
                        }
                    }
                    elseif ($perm.GrantedTo -and $perm.GrantedTo.User -and $perm.GrantedTo.User.DisplayName -ne $userName) {
                        $sharedWithUsers += [PSCustomObject]@{
                            UserName = $perm.GrantedTo.User.DisplayName
                            UserEmail = $perm.GrantedTo.User.Email
                            AccessType = ($perm.Roles -join ", ")
                        }
                    }
                }
            } catch {
                # Silent error handling for permissions
            }
            
            $oneDriveDetails += [PSCustomObject]@{
                SiteName = $site.DisplayName
                UserName = $userName
                SiteUrl = $site.WebUrl
                SizeGB = [math]::Round($totalSize / 1GB, 3)
                SizeMB = [math]::Round($totalSize / 1MB, 2)
                FileCount = $fileCount
                SharedWithCount = $sharedWithUsers.Count
                SharedWithUsers = ($sharedWithUsers.UserName -join "; ")
                SharedWithEmails = ($sharedWithUsers.UserEmail -join "; ")
                AccessDetails = ($sharedWithUsers | ForEach-Object { "$($_.UserName) ($($_.AccessType))" }) -join "; "
                LastAnalyzed = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        catch {
            Write-Host "[Warning] Could not analyze OneDrive site: $($site.DisplayName) - $_" -ForegroundColor Yellow
            $oneDriveDetails += [PSCustomObject]@{
                SiteName = $site.DisplayName
                UserName = "Analysis Failed"
                SiteUrl = $site.WebUrl
                SizeGB = 0
                SizeMB = 0
                FileCount = 0
                SharedWithCount = 0
                SharedWithUsers = "Unable to retrieve data"
                SharedWithEmails = "Check permissions"
                AccessDetails = "Error: $_"
                LastAnalyzed = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
    }
    
    Write-Progress -Activity "Analyzing Personal OneDrive Sites" -Completed
    return $oneDriveDetails
}
#endregion
#region Recycle Bin Functions
function Get-SiteRecycleBinStorage {
    param($SiteId)
    
    try {
        # Try to get recycle bin items using Graph API
        $recycleBinItems = @()
        
        # Method 1: Try direct Graph API call
        try {
            $recycleBinItems = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/recycleBin" -Method GET -ErrorAction SilentlyContinue
            if ($recycleBinItems.value) {
                $recycleBinItems = $recycleBinItems.value
            }
        } catch {
            # Method 2: Try alternative approach with drives
            try {
                $drives = Get-MgSiteDrive -SiteId $SiteId -ErrorAction SilentlyContinue
                foreach ($drive in $drives) {
                    try {
                        $driveRecycleBin = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/drives/$($drive.Id)/recycleBin" -Method GET -ErrorAction SilentlyContinue
                        if ($driveRecycleBin.value) {
                            $recycleBinItems += $driveRecycleBin.value
                        }
                    } catch {
                        # Skip this drive if recycle bin access fails
                    }
                }
            } catch {
                # Recycle bin access not available for this site
            }
        }
        
        if ($recycleBinItems.Count -gt 0) {
            $totalSize = ($recycleBinItems | Where-Object { $_.size } | Measure-Object -Property size -Sum -ErrorAction SilentlyContinue).Sum
            return [math]::Round($totalSize / 1GB, 3)
        }
        
        return 0
    } catch {
        return 0
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
        
        # Use default search terms
        $searchTerms = "AllSites"
        
        Write-Host "[Info] Starting tenant-wide SharePoint audit..." -ForegroundColor Cyan
        $tenantName = Get-TenantName
        $dateStr = Get-Date -Format yyyyMMdd
        
        # Create filename based on search terms, tenant name, and date
        $cleanSearchTerms = $searchTerms -replace '[^\w\s-]', '' -replace '\s+', '_'
        $defaultFileName = "SharePointAudit-$cleanSearchTerms-$tenantName-$dateStr.xlsx"
        
        # Get save file path from user dialog
        $excelFileName = Get-SaveFileDialog -DefaultFileName $defaultFileName -Title "Save SharePoint Audit Report"
        if (-not $excelFileName) {
            Write-Host "[Info] User cancelled the save dialog. Exiting." -ForegroundColor Yellow
            return
        }
        
        # Get all SharePoint sites in the tenant
        $sites = Get-AllSharePointSites
        if ($sites.Count -eq 0) {
            Write-Host "[Warning] No sites found. Exiting." -ForegroundColor Yellow
            return
        }
        
        Write-Host "[Info] Found $($sites.Count) total sites to analyze (including SharePoint sites and OneDrive personal sites)..." -ForegroundColor Cyan

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
                    "contentstorage",
                    "portals/hub",
                    "_api",
                    "search",
                    "admin"
                )
                
                # Do not skip personal OneDrive sites even if they have "mysites" or "personal" in URL
                if (-not $isOneDrive) {
                    $systemSitePatterns += @("mysites", "personal")
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
            
            # Scan ALL sites for large files (>100MB) as per description requirement
            $fileDataForLargeFiles = @{ Files = @(); FolderSizes = @{}; TotalFiles = 0; TotalSizeGB = 0 }
            try {
                $fileDataForLargeFiles = Get-FileData -Site $site
                
                # Find large files (>100MB) in this site
                if ($fileDataForLargeFiles.Files.Count -gt 0) {
                    $largeFiles = $fileDataForLargeFiles.Files | Where-Object { $_.Size -gt 100MB } | ForEach-Object {
                        [PSCustomObject]@{
                            SiteName = $site.DisplayName
                            FileName = $_.Name
                            SizeMB = $_.SizeMB
                            SizeGB = $_.SizeGB
                            Path = $_.Path
                            Extension = $_.Extension
                            LibraryName = $_.LibraryName
                            FullPath = $_.FullPath
                            PathLength = $_.PathLength
                        }
                    }
                    if ($largeFiles.Count -gt 0) {
                        $global:allLargeFiles += $largeFiles
                    }
                }
            }
            catch {
                Write-Host "[Error] Failed to scan site $($site.DisplayName) for large files: $_" -ForegroundColor Red
            }
            
            if ($isTopSite) {
                # Use the already collected file data for top sites detailed analysis
                $fileData = $fileDataForLargeFiles
                
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
                    $allTopFiles += $fileData.Files | Select-Object @{Name="SiteName";Expression={$site.DisplayName}}, *
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
                    $ownersSheetName = Get-SafeWorksheetName ("Owners - " + $site.DisplayName.Substring(0, [Math]::Min($site.DisplayName.Length, 20)))
                    $excel = $userAccess.Owners | Export-Excel -Path $excelFileName -WorksheetName $ownersSheetName -AutoSize -TableStyle Light1 -PassThru
                    
                    # Apply high contrast formatting to Owners worksheet
                    $ws = $excel.Workbook.Worksheets[$ownersSheetName]
                    if ($ws) {
                        # Set high contrast colors for header row - gold theme for owners
                        $headerRange = $ws.Cells[$ws.Dimension.Start.Row, $ws.Dimension.Start.Column, $ws.Dimension.Start.Row, $ws.Dimension.End.Column]
                        $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                        $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::DarkGoldenrod)
                        $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                        $headerRange.Style.Font.Bold = $true
                        $headerRange.Style.Font.Size = 12
                        
                        # Set alternating high contrast rows for data
                        $dataRows = $ws.Dimension.Rows
                        for ($row = 2; $row -le $dataRows; $row++) {
                            $rowRange = $ws.Cells[$row, $ws.Dimension.Start.Column, $row, $ws.Dimension.End.Column]
                            if ($row % 2 -eq 0) {
                                $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                                $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Goldenrod)
                                $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                            } else {
                                $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                                $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGoldenrodYellow)
                                $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::Black)
                            }
                            $rowRange.Style.Font.Bold = $true
                            $rowRange.Style.Font.Size = 10
                        }
                    }
                    Close-ExcelPackage $excel
                }
                if ($userAccess.Members.Count -gt 0) {
                    $membersSheetName = Get-SafeWorksheetName ("Members - " + $site.DisplayName.Substring(0, [Math]::Min($site.DisplayName.Length, 20)))
                    $excel = $userAccess.Members | Export-Excel -Path $excelFileName -WorksheetName $membersSheetName -AutoSize -TableStyle Light1 -PassThru
                    
                    # Apply high contrast formatting to Members worksheet
                    $ws = $excel.Workbook.Worksheets[$membersSheetName]
                    if ($ws) {
                        # Set high contrast colors for header row - blue theme for members
                        $headerRange = $ws.Cells[$ws.Dimension.Start.Row, $ws.Dimension.Start.Column, $ws.Dimension.Start.Row, $ws.Dimension.End.Column]
                        $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                        $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::DarkBlue)
                        $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                        $headerRange.Style.Font.Bold = $true
                        $headerRange.Style.Font.Size = 12
                        
                        # Set alternating high contrast rows for data
                        $dataRows = $ws.Dimension.Rows
                        for ($row = 2; $row -le $dataRows; $row++) {
                            $rowRange = $ws.Cells[$row, $ws.Dimension.Start.Column, $row, $ws.Dimension.End.Column]
                            if ($row % 2 -eq 0) {
                                $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                                $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::MediumBlue)
                                $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                            } else {
                                $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                                $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightSkyBlue)
                                $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::Black)
                            }
                            $rowRange.Style.Font.Bold = $true
                            $rowRange.Style.Font.Size = 10
                        }
                    }
                    Close-ExcelPackage $excel
                }
                
                # Export external guests with highlighting
                try {
                    $siteInfo = Get-SiteStorageAndAccess -Site $site
                    if ($siteInfo.ExternalGuests.Count -gt 0) {
                        $guestsSheetName = Get-SafeWorksheetName ("External Guests - " + $site.DisplayName.Substring(0, [Math]::Min($site.DisplayName.Length, 15)))
                        
                        # Export with high contrast formatting to highlight external guests in red
                        $excel = $siteInfo.ExternalGuests | Export-Excel -Path $excelFileName -WorksheetName $guestsSheetName -AutoSize -TableStyle Light1 -PassThru
                        $ws = $excel.Workbook.Worksheets[$guestsSheetName]
                        
                        if ($ws) {
                            # Apply high contrast red formatting for external guests (security warning theme)
                            $headerRange = $ws.Cells[$ws.Dimension.Start.Row, $ws.Dimension.Start.Column, $ws.Dimension.Start.Row, $ws.Dimension.End.Column]
                            $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::DarkRed)
                            $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::Yellow)
                            $headerRange.Style.Font.Bold = $true
                            $headerRange.Style.Font.Size = 12
                            
                            # Apply high contrast red to all data cells for external guests
                            $dataRows = $ws.Dimension.Rows
                            for ($row = 2; $row -le $dataRows; $row++) {
                                $rowRange = $ws.Cells[$row, $ws.Dimension.Start.Column, $row, $ws.Dimension.End.Column]
                                if ($row % 2 -eq 0) {
                                    # Even rows - bright red background with white text
                                    $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                                    $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Red)
                                    $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                                } else {
                                    # Odd rows - light coral background with dark red text
                                    $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                                    $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightCoral)
                                    $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::DarkRed)
                                }
                                $rowRange.Style.Font.Bold = $true
                                $rowRange.Style.Font.Size = 10
                            }
                        }
                        
                        Close-ExcelPackage $excel
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

        # Remove any existing Excel file to avoid conflicts
        if (Test-Path $excelFileName) {
            Remove-Item $excelFileName -Force -ErrorAction SilentlyContinue
        }

        Write-Host "[Info] Creating Excel report with $($allSiteSummaries.Count) site summaries..." -ForegroundColor Cyan

        try {
            # Create comprehensive SharePoint Storage Pie Chart Data (including recycle bins and personal sites)
            Write-Progress -Activity "Generating Excel Report" -Status "Creating comprehensive storage analysis..." -PercentComplete 5
            
            $comprehensiveStorageData = @()
            $totalTenantStorage = 0
            $recycleBinStorage = 0
            $personalOneDriveStorage = 0
            $sharePointSitesStorage = 0
            
            # Calculate storage breakdown
            foreach ($siteSummary in $siteSummaries) {
                $totalTenantStorage += $siteSummary.StorageGB
                
                if ($siteSummary.IsOneDrive) {
                    $personalOneDriveStorage += $siteSummary.StorageGB
                } else {
                    $sharePointSitesStorage += $siteSummary.StorageGB
                }
                
                # Add individual site to comprehensive data
                $comprehensiveStorageData += [PSCustomObject]@{
                    Category = if ($siteSummary.IsOneDrive) { "Personal OneDrive" } else { "SharePoint Site" }
                    SiteName = $siteSummary.SiteName
                    StorageGB = $siteSummary.StorageGB
                    Percentage = if ($totalTenantStorage -gt 0) { [math]::Round(($siteSummary.StorageGB / $totalTenantStorage) * 100, 2) } else { 0 }
                    SiteUrl = $siteSummary.SiteUrl
                    SiteType = $siteSummary.SiteType
                }
            }
            
            # Get recycle bin storage (attempt to retrieve from all sites)
            Write-Host "[Info] Attempting to calculate recycle bin storage..." -ForegroundColor Cyan
            foreach ($siteSummary in $siteSummaries) {
                try {
                    $recycleBinSizeForSite = Get-SiteRecycleBinStorage -SiteId $siteSummary.SiteId
                    if ($recycleBinSizeForSite -gt 0) {
                        $recycleBinStorage += $recycleBinSizeForSite
                    }
                } catch {
                    Write-Host "[Debug] Could not access recycle bin for site: $($siteSummary.SiteName)" -ForegroundColor Yellow
                }
            }
            
            # Create pie chart summary data
            $pieChartData = @(
                [PSCustomObject]@{
                    Category = "SharePoint Sites"
                    StorageGB = $sharePointSitesStorage
                    Percentage = if ($totalTenantStorage -gt 0) { [math]::Round(($sharePointSitesStorage / $totalTenantStorage) * 100, 2) } else { 0 }
                    SiteCount = ($siteSummaries | Where-Object { -not $_.IsOneDrive }).Count
                },
                [PSCustomObject]@{
                    Category = "Personal OneDrive"
                    StorageGB = $personalOneDriveStorage
                    Percentage = if ($totalTenantStorage -gt 0) { [math]::Round(($personalOneDriveStorage / $totalTenantStorage) * 100, 2) } else { 0 }
                    SiteCount = ($siteSummaries | Where-Object { $_.IsOneDrive }).Count
                },
                [PSCustomObject]@{
                    Category = "Recycle Bins"
                    StorageGB = $recycleBinStorage
                    Percentage = if (($totalTenantStorage + $recycleBinStorage) -gt 0) { [math]::Round(($recycleBinStorage / ($totalTenantStorage + $recycleBinStorage)) * 100, 2) } else { 0 }
                    SiteCount = "All Sites"
                }
            )
            
            # Create comprehensive site details with users and access information
            Write-Progress -Activity "Generating Excel Report" -Status "Compiling comprehensive site details..." -PercentComplete 10
            
            $comprehensiveSiteDetails = @()
            foreach ($siteSummary in $siteSummaries) {
                $site = $siteSummary.Site
                
                # Get detailed user access for this site
                try {
                    $userAccess = Get-SiteUserAccessSummary -Site $site
                    $folderAccess = @()
                    
                    # Get folder access permissions
                    try {
                        $folderAccess = Get-ParentFolderAccess -Site $site
                    } catch {
                        Write-Host "[Debug] Could not get folder access for site: $($site.DisplayName)" -ForegroundColor Yellow
                    }
                    
                    # Compile all users and groups with their access details
                    $allUsersString = ""
                    $allOwnersString = ""
                    $accessTypesString = ""
                    $foldersAccessString = ""
                    
                    if ($userAccess.Owners.Count -gt 0) {
                        $allOwnersString = ($userAccess.Owners | ForEach-Object { "$($_.DisplayName) ($($_.Mail))" }) -join "; "
                    }
                    
                    if ($userAccess.Members.Count -gt 0) {
                        $membersString = ($userAccess.Members | ForEach-Object { "$($_.DisplayName) ($($_.Mail)) - $($_.AccessLevel)" }) -join "; "
                        $allUsersString = $membersString
                    }
                    
                    if ($folderAccess.Count -gt 0) {
                        $accessTypesString = ($folderAccess.PermissionLevel | Sort-Object -Unique) -join ", "
                        $foldersAccessString = ($folderAccess | ForEach-Object { "$($_.FolderPath) - $($_.PermissionLevel)" }) -join "; "
                    }
                    
                    $comprehensiveSiteDetails += [PSCustomObject]@{
                        SiteName = $site.DisplayName
                        SiteUrl = $site.WebUrl
                        SiteType = $siteSummary.SiteType
                        StorageGB = $siteSummary.StorageGB
                        TotalFiles = if ($siteSummary.SiteId -in $topSites.SiteId) { 
                            ($allTopFiles | Where-Object { $_.SiteName -eq $site.DisplayName }).Count 
                        } else { "Not analyzed (not in top 10)" }
                        OwnersCount = $userAccess.Owners.Count
                        MembersCount = $userAccess.Members.Count
                        AllOwners = $allOwnersString
                        AllUsersAndGroups = $allUsersString
                        AccessTypes = $accessTypesString
                        FoldersWithAccess = $foldersAccessString
                        UniquePermissionLevels = if ($folderAccess.Count -gt 0) { ($folderAccess.PermissionLevel | Sort-Object -Unique).Count } else { 0 }
                        LastAnalyzed = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        IsTopSite = if ($siteSummary.SiteId -in $topSites.SiteId) { "Yes" } else { "No" }
                    }
                } catch {
                    Write-Host "[Error] Failed to compile comprehensive details for site $($site.DisplayName): $_" -ForegroundColor Red
                    
                    # Add basic info even if detailed analysis fails
                    $comprehensiveSiteDetails += [PSCustomObject]@{
                        SiteName = $site.DisplayName
                        SiteUrl = $site.WebUrl
                        SiteType = $siteSummary.SiteType
                        StorageGB = $siteSummary.StorageGB
                        TotalFiles = "Error retrieving data"
                        OwnersCount = 0
                        MembersCount = 0
                        AllOwners = "Error retrieving data"
                        AllUsersAndGroups = "Error retrieving data"
                        AccessTypes = "Error retrieving data"
                        FoldersWithAccess = "Error retrieving data"
                        UniquePermissionLevels = 0
                        LastAnalyzed = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        IsTopSite = if ($siteSummary.SiteId -in $topSites.SiteId) { "Yes" } else { "No" }
                    }
                }
            }
            
            # Export SharePoint Storage Pie Chart first (at top of workbook)
            Write-Progress -Activity "Generating Excel Report" -Status "Creating SharePoint Storage Pie Chart..." -PercentComplete 15
            if ($pieChartData.Count -gt 0) {
                $excel = $pieChartData | Export-Excel -Path $excelFileName -WorksheetName "SharePoint Storage Pie Chart" -AutoSize -TableStyle Light1 -Title "SharePoint Tenant Storage Overview (Includes Recycle Bins)" -TitleBold -TitleSize 16 -PassThru
                
                # Apply high contrast formatting to pie chart worksheet
                $ws = $excel.Workbook.Worksheets["SharePoint Storage Pie Chart"]
                if ($ws) {
                    # Set high contrast colors for header row
                    $headerRange = $ws.Cells["A1:D1"]
                    $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Black)
                    $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                    $headerRange.Style.Font.Bold = $true
                    $headerRange.Style.Font.Size = 12
                    
                    # Set alternating high contrast rows for data
                    $dataRows = $ws.Dimension.Rows
                    for ($row = 2; $row -le $dataRows; $row++) {
                        $rowRange = $ws.Cells["A$row:D$row"]
                        if ($row % 2 -eq 0) {
                            # Even rows - dark gray background with white text
                            $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::DarkGray)
                            $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                        } else {
                            # Odd rows - light gray background with black text
                            $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)
                            $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::Black)
                        }
                        $rowRange.Style.Font.Bold = $true
                        $rowRange.Style.Font.Size = 11
                    }
                }
                Close-ExcelPackage $excel
            }
            
            # Export comprehensive site details worksheet
            Write-Progress -Activity "Generating Excel Report" -Status "Creating comprehensive site details..." -PercentComplete 20
            if ($comprehensiveSiteDetails.Count -gt 0) {
                $excel = $comprehensiveSiteDetails | Export-Excel -Path $excelFileName -WorksheetName "Site Summary with Details" -AutoSize -TableStyle Light1 -Title "Complete Site Summary with Users, Groups, and Access Details" -TitleBold -TitleSize 14 -PassThru
                
                # Apply high contrast formatting to site details worksheet
                $ws = $excel.Workbook.Worksheets["Site Summary with Details"]
                if ($ws) {
                    # Set high contrast colors for header row
                    $headerRange = $ws.Cells[$ws.Dimension.Start.Row, $ws.Dimension.Start.Column, $ws.Dimension.Start.Row, $ws.Dimension.End.Column]
                    $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Navy)
                    $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                    $headerRange.Style.Font.Bold = $true
                    $headerRange.Style.Font.Size = 12
                    
                    # Set alternating high contrast rows for data
                    $dataRows = $ws.Dimension.Rows
                    for ($row = 2; $row -le $dataRows; $row++) {
                        $rowRange = $ws.Cells[$row, $ws.Dimension.Start.Column, $row, $ws.Dimension.End.Column]
                        if ($row % 2 -eq 0) {
                            # Even rows - dark blue background with white text
                            $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::DarkBlue)
                            $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                        } else {
                            # Odd rows - light blue background with black text
                            $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightSteelBlue)
                            $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::Black)
                        }
                        $rowRange.Style.Font.Bold = $true
                        $rowRange.Style.Font.Size = 10
                    }
                }
                Close-ExcelPackage $excel
            }
            
            # Export detailed storage breakdown
            if ($comprehensiveStorageData.Count -gt 0) {
                $excel = $comprehensiveStorageData | Sort-Object StorageGB -Descending | Export-Excel -Path $excelFileName -WorksheetName "Detailed Storage Breakdown" -AutoSize -TableStyle Light1 -PassThru
                
                # Apply high contrast formatting to storage breakdown worksheet
                $ws = $excel.Workbook.Worksheets["Detailed Storage Breakdown"]
                if ($ws) {
                    # Set high contrast colors for header row
                    $headerRange = $ws.Cells[$ws.Dimension.Start.Row, $ws.Dimension.Start.Column, $ws.Dimension.Start.Row, $ws.Dimension.End.Column]
                    $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::DarkGreen)
                    $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                    $headerRange.Style.Font.Bold = $true
                    $headerRange.Style.Font.Size = 12
                    
                    # Set alternating high contrast rows for data
                    $dataRows = $ws.Dimension.Rows
                    for ($row = 2; $row -le $dataRows; $row++) {
                        $rowRange = $ws.Cells[$row, $ws.Dimension.Start.Column, $row, $ws.Dimension.End.Column]
                        if ($row % 2 -eq 0) {
                            # Even rows - dark green background with white text
                            $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::DarkOliveGreen)
                            $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                        } else {
                            # Odd rows - light green background with black text
                            $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGreen)
                            $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::Black)
                        }
                        $rowRange.Style.Font.Bold = $true
                        $rowRange.Style.Font.Size = 10
                    }
                }
                Close-ExcelPackage $excel
            }
            
            # Create the original summary for compatibility
            # Split summaries into Document Library (SharePoint) and Personal OneDrive
            $docLibrarySummaries = $allSiteSummaries | Where-Object { $_.SiteType -eq "SharePoint Site" }
            $oneDriveSummaries = $allSiteSummaries | Where-Object { $_.SiteType -eq "OneDrive Personal" }

            # Export Document Library Sites summary
            if ($docLibrarySummaries.Count -gt 0) {
                $excel = $docLibrarySummaries | Export-Excel -Path $excelFileName -WorksheetName "Document Library Sites" -AutoSize -TableStyle Light1 -PassThru
                $ws = $excel.Workbook.Worksheets["Document Library Sites"]
                if ($ws) {
                    $headerRange = $ws.Cells[$ws.Dimension.Start.Row, $ws.Dimension.Start.Column, $ws.Dimension.Start.Row, $ws.Dimension.End.Column]
                    $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Black)
                    $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::Yellow)
                    $headerRange.Style.Font.Bold = $true
                    $headerRange.Style.Font.Size = 12
                    $dataRows = $ws.Dimension.Rows
                    for ($row = 2; $row -le $dataRows; $row++) {
                        $rowRange = $ws.Cells[$row, $ws.Dimension.Start.Column, $row, $ws.Dimension.End.Column]
                        if ($row % 2 -eq 0) {
                            $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::DarkGray)
                            $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::Yellow)
                        } else {
                            $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)
                            $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::Black)
                        }
                        $rowRange.Style.Font.Bold = $true
                        $rowRange.Style.Font.Size = 10
                    }
                }
                Close-ExcelPackage $excel
            }

            # Export Personal OneDrive Sites summary
            if ($oneDriveSummaries.Count -gt 0) {
                $excel = $oneDriveSummaries | Export-Excel -Path $excelFileName -WorksheetName "Personal OneDrive Sites Summary" -AutoSize -TableStyle Light1 -PassThru
                $ws = $excel.Workbook.Worksheets["Personal OneDrive Sites Summary"]
                if ($ws) {
                    $headerRange = $ws.Cells[$ws.Dimension.Start.Row, $ws.Dimension.Start.Column, $ws.Dimension.Start.Row, $ws.Dimension.End.Column]
                    $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::DarkCyan)
                    $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                    $headerRange.Style.Font.Bold = $true
                    $headerRange.Style.Font.Size = 12
                    $dataRows = $ws.Dimension.Rows
                    for ($row = 2; $row -le $dataRows; $row++) {
                        $rowRange = $ws.Cells[$row, $ws.Dimension.Start.Column, $row, $ws.Dimension.End.Column]
                        if ($row % 2 -eq 0) {
                            $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Teal)
                            $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                        } else {
                            $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::PaleTurquoise)
                            $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::Black)
                        }
                        $rowRange.Style.Font.Bold = $true
                        $rowRange.Style.Font.Size = 10
                    }
                }
                Close-ExcelPackage $excel
            }
        } catch {
            Write-Host "[Error] Failed to create Excel report: $_" -ForegroundColor Red
            throw $_
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