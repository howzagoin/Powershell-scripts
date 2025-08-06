<#
SharePointAudit.ps1

Description:
This script audits all SharePoint and OneDrive sites in a Microsoft 365 tenant, gathering storage, file/folder, and user access details. It authenticates using app-only certificate, queries Microsoft Graph, and generates a comprehensive Excel report and CSV export for direct comparison with SharePoint admin exports.

Excel Output:
- Worksheet: "SharePoint Storage Overview"
  Columns:
    - Site name
    - URL
    - Teams
    - Channel sites
    - IBMode
    - Storage used (GB)
    - Recycle bin (GB)
    - Total storage (GB)
    - Primary admin
    - Hub
    - Template
    - Last activity (UTC)
    - Date created
    - Created by
    - Storage limit (GB)
    - Storage used (%)
    - Microsoft 365 group
    - Files viewed or edited
    - Page views
    - Page visits
    - Files
    - Sensitivity
    - External sharing
    - OwnersCount
    - MembersCount
    - ReportDate

- Worksheet: "Tenant Storage Breakdown" (Pie Chart Data)
  Columns:
    - Category (SharePoint Sites, Personal OneDrive, Recycle Bins)
    - StorageGB
    - Percentage
    - SiteCount

- Worksheet: "User & Group Access Overview"
  Columns:
    - User/Group
    - Email
    - Type (Internal, External Guest)
    - Role
    - Site
    - Site URL
    - Access
    - Object
    - Highlight (for external guests)

- Worksheet: "OneDrive Personal Sites"
  Columns:
    - User Name
    - Site Name
    - URL
    - Storage Used (GB)
    - Storage Used (%)
    - Storage Limit (GB)
    - Last Activity (UTC)
    - Date Created
    - Files Count
    - External Sharing
    - Site Type
    - Report Date

- Worksheet: [site-specific] "$siteName Storage" (for top 10 sites)
  Columns:
    - Location (folder breakdown pie chart data)
    - SizeGB
    - Top 20 files (Name, Size, Path, Drive, Extension, etc.)
    - Top 20 folders (FolderPath, SizeGB, SizeMB)

Key Features:
- Enhanced storage calculations with multiple approaches for accuracy
- Recycle bin storage analysis across all sites
- Comprehensive user access reporting including external guests
- OneDrive personal site detection and aggregation
- Parallel processing for improved performance
- Detailed file and folder analysis for largest sites
- Pie chart data for tenant-level and site-level storage breakdowns

CSV Output:
- Contains the same columns as "SharePoint Storage Overview" worksheet for direct comparison with admin export.
#>

## All function definitions moved to top of file

## All function definitions moved to top of file
#region Logging, Progress, and Utility Functions
function Get-SiteUserAccessSummary {
    param(
        [Parameter(Mandatory=$true)]
        $Site
    )
    $owners = @()
    $members = @()
    try {
        $permissions = Get-MgSitePermission -SiteId $Site.Id -All -ErrorAction SilentlyContinue
        foreach ($perm in $permissions) {
            if ($perm.GrantedToV2 -and $perm.GrantedToV2.User) {
                $userType = if ($perm.GrantedToV2.User.UserType -eq 'Guest') { "External Guest" } else { "Internal" }
                $userObj = [PSCustomObject]@{
                    DisplayName = $perm.GrantedToV2.User.DisplayName
                    UserEmail = $perm.GrantedToV2.User.Email
                    UserType = $userType
                    Role = ($perm.Roles -join ', ')
                }
                if ($userObj.Role -match 'Owner|Admin') {
                    $owners += $userObj
                } else {
                    $members += $userObj
                }
            }
        }
    } catch {
        Write-Log "Failed to get user access for site $($Site.DisplayName): $_" -Level Error
    }
    return @{ Owners = $owners; Members = $members }
}
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [ValidateSet("Info", "Warning", "Error", "Success", "Debug")]
        [string]$Level = "Info",
        
        [string]$ForegroundColor
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "Info"    { $color = if ($ForegroundColor) { $ForegroundColor } else { "Cyan" } }
        "Warning" { $color = if ($ForegroundColor) { $ForegroundColor } else { "Yellow" } }
        "Error"   { $color = if ($ForegroundColor) { $ForegroundColor } else { "Red" } }
        "Success" { $color = if ($ForegroundColor) { $ForegroundColor } else { "Green" } }
        "Debug"   { $color = if ($ForegroundColor) { $ForegroundColor } else { "Gray" } }
        default   { $color = "White" }
    }
    
    Write-Host $logMessage -ForegroundColor $color
}

function Show-Progress {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete,
        [string]$CurrentOperation,
        [int]$Id = 1
    )
    
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete -CurrentOperation $CurrentOperation -Id $Id
}

function Stop-Progress {
    param(
        [string]$Activity,
        [int]$Id = 1
    )
    
    Write-Progress -Activity $Activity -Completed -Id $Id
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory=$true)]
        [scriptblock]$ScriptBlock,
        
        [int]$MaxRetries = 5,
        [int]$DelaySeconds = 2,
        
        [string]$Activity = "Retrying Operation"
    )
    
    $attempt = 0
    $lastError = $null
    
    while ($attempt -lt $MaxRetries) {
        try {
            return & $ScriptBlock
        } 
        catch {
            $lastError = $_
            $attempt++
            
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode -eq 429) {
                $wait = $DelaySeconds * $attempt
                Write-Log "Throttled (429). Retrying in $wait seconds... (Attempt $attempt/$MaxRetries)" -Level Warning
                Start-Sleep -Seconds $wait
            } 
            elseif ($attempt -ge $MaxRetries) {
                Write-Log "Max retries exceeded for operation: $Activity" -Level Error
                throw $lastError
            }
            else {
                Write-Log "Operation failed (Attempt $attempt/$MaxRetries): $($_.Exception.Message)" -Level Warning
                Start-Sleep -Seconds $DelaySeconds
            }
        }
    }
    
    throw $lastError
}

function Format-WorksheetName {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name
    )
    
    if ([string]::IsNullOrWhiteSpace($Name)) { 
        return "Sheet1" 
    }
    
    # Remove or replace invalid characters for Excel worksheet names
    # Invalid characters: [ ] : * ? / \ > < |
    # Remove invalid Excel worksheet characters and non-printable characters
    $safeName = $Name -replace '[\[\]:*?/\\><|]', '_'
    $safeName = $safeName -replace '[\x00-\x1F]', ''
    
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
#endregion

#region Module Management
function Install-OrUpdateModule {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ModuleName,
        
        [Parameter(Mandatory=$true)]
        [string]$MinimumVersion,
        
        [switch]$Force
    )
    
    $installedModule = Get-Module -Name $ModuleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    
    if (-not $installedModule) {
        Write-Log "Installing module: $ModuleName (minimum version: $MinimumVersion)" -Level Info
        Install-Module -Name $ModuleName -MinimumVersion $MinimumVersion -Scope CurrentUser -Force -AllowClobber
        return $true
    }
    elseif ($Force -or [version]$installedModule.Version -lt [version]$MinimumVersion) {
        Write-Log "Updating module: $ModuleName from $($installedModule.Version) to $MinimumVersion" -Level Info
        # First, try to uninstall all versions
        try {
            Get-Module -Name $ModuleName -All | Remove-Module -Force
            Get-InstalledModule -Name $ModuleName -AllVersions | Uninstall-Module -Force -ErrorAction SilentlyContinue
        } catch {
            Write-Log "Could not uninstall all versions of $ModuleName. Trying to install anyway." -Level Warning
        }
        
        # Then install the new version
        Install-Module -Name $ModuleName -MinimumVersion $MinimumVersion -Scope CurrentUser -Force -AllowClobber
        return $true
    }
    
    return $false
}
#endregion

#region Authentication
function Connect-ToGraph {
    param(
        [string]$ClientId,
        [string]$TenantId,
        [string]$CertificateThumbprint
    )
    
    Write-Log "Connecting to Microsoft Graph..." -Level Info
    
    try {
        # Clear existing connections
        Disconnect-MgGraph -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        
        # Get certificate
        $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction Stop
        if (-not $cert) { 
            throw "Certificate with thumbprint $CertificateThumbprint not found." 
        }
        
        # Connect with app-only authentication
        Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Certificate $cert -NoWelcome -WarningAction SilentlyContinue
        
        # Verify app-only authentication
        $context = Get-MgContext
        if (-not $context) {
            throw "Graph context missing. Authentication failed."
        }
        
        Write-Log "Graph context: TenantId=$($context.TenantId), AuthType=$($context.AuthType), Scopes=$($context.Scopes -join ", ")" -Level Debug
        
        if ($context.AuthType -ne 'AppOnly') { 
            throw "App-only authentication required." 
        }
        
        Write-Log "Successfully connected with app-only authentication" -Level Success
        return $true
    }
    catch {
        Write-Log "Authentication failed: $_" -Level Error
        throw
    }
}

function Get-TenantName {
    try {
        $tenant = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($tenant) { 
            return $tenant.DisplayName.Replace(' ', '_') 
        }
        return 'Tenant'
    }
    catch {
        Write-Log "Error getting tenant name: $_" -Level Warning
        return 'Tenant'
    }
}
#endregion

#region Site Discovery
function Get-AllSharePointSites {
    Write-Log "Enumerating all SharePoint sites in tenant..." -Level Info
    
    try {
        $sites = @()
        
        # Approach 1: Get root site
        try {
            $rootSite = Get-MgSite -SiteId "root" -ErrorAction SilentlyContinue
            if ($rootSite) {
                $sites += $rootSite
                Write-Log "Root site found: $($rootSite.DisplayName)" -Level Success
            }
        }
        catch {
            Write-Log "Could not get root site: $_" -Level Warning
        }
        
        # Approach 2: Get all sites using Graph API
        try {
            $allSites = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites?`$top=500" -ErrorAction SilentlyContinue
            if ($allSites.value) {
                $sites += $allSites.value
                Write-Log "Found $($allSites.value.Count) sites via Graph API" -Level Info
            }
        }
        catch {
            Write-Log "Graph API site search failed: $_" -Level Warning
        }
        
        # Approach 3: Search for site collections
        try {
            $searchSites = Get-MgSite -Search "*" -All -ErrorAction SilentlyContinue
            if ($searchSites) {
                $sites += $searchSites
                Write-Log "Found $($searchSites.Count) sites via search" -Level Info
            }
        }
        catch {
            Write-Log "Site search failed: $_" -Level Warning
        }
        
        # Approach 4: Try to get OneDrive personal sites via users endpoint
        try {
            Write-Log "Attempting to discover OneDrive personal sites..." -Level Info
            $users = Get-MgUser -All -Property "Id,UserPrincipalName,DisplayName" -Filter "UserType eq 'Member'" -Top 100 -ErrorAction SilentlyContinue
            foreach ($user in $users) {
                try {
                    $userDrive = Get-MgUserDrive -UserId $user.Id -ErrorAction SilentlyContinue
                    if ($userDrive -and $userDrive.WebUrl) {
                        # Try to get the site information for the OneDrive
                        $oneDriveSite = Get-MgSite -SiteId "root" -ErrorAction SilentlyContinue | Where-Object { $_.WebUrl -like "*$($user.UserPrincipalName.Split('@')[0])*" }
                        if ($oneDriveSite) {
                            $sites += $oneDriveSite
                        }
                    }
                } catch {
                    # Skip users without accessible OneDrive
                    continue
                }
            }
            Write-Log "Completed OneDrive personal site discovery" -Level Info
        }
        catch {
            Write-Log "OneDrive personal site discovery failed: $_" -Level Warning
        }
        
        # Remove duplicates and ensure we have valid sites
        $sites = $sites | Where-Object { $_ -and $_.Id -and $_.DisplayName } | Sort-Object Id -Unique
        
        if (-not $sites -or $sites.Count -eq 0) {
            Write-Log "No SharePoint sites found in tenant!" -Level Warning
            return @()
        }
        
        Write-Log "Found $($sites.Count) SharePoint sites." -Level Success
        
        return $sites
    }
    catch {
        Write-Log "Failed to enumerate SharePoint sites: $_" -Level Error
        return @()
    }
}

function Get-SiteInfo {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl
    )
    
    Write-Log "Getting site information for: $SiteUrl" -Level Info
    
    try {
        # Extract site ID from URL
        $uri = [Uri]$SiteUrl
        $sitePath = $uri.AbsolutePath
        $siteId = "$($uri.Host):$sitePath"
        
        $site = Get-MgSite -SiteId $siteId -ErrorAction Stop
        Write-Log "Found site: $($site.DisplayName)" -Level Success
        
        return $site
    }
    catch {
        Write-Log "Failed to get site information: $_" -Level Error
        throw
    }
}
#endregion

#region File and Storage Analysis
function Get-TotalItemCount {
    param(
        [Parameter(Mandatory=$true)]
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
        [Parameter(Mandatory=$true)]
        [string]$DriveId,
        
        [Parameter(Mandatory=$true)]
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
                # Improved progress bar
                $percent = if ($TotalItems -gt 0) { [Math]::Min(100, [int](($GlobalItemIndex.Value/$TotalItems)*100)) } else { 100 }
                $progressBar = ("#" * ($percent / 2)) + ("-" * (50 - ($percent / 2)))
                Show-Progress -Activity "Scanning SharePoint Site Content" -Status "[$progressBar] $percent% - Processing: $($child.Name) ($($GlobalItemIndex.Value)/$TotalItems)" -PercentComplete $percent -Id 1
                
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
    param(
        [Parameter(Mandatory=$true)]
        $Site
    )
    
    try {
        Write-Log "Processing site: $($Site.DisplayName)" -Level Info
        
        # Get all document libraries
        $lists = Invoke-WithRetry -ScriptBlock { Get-MgSiteList -SiteId $Site.Id -WarningAction SilentlyContinue } -Activity "Get site lists"
        $docLibraries = $lists | Where-Object { $_.List -and $_.List.Template -eq "documentLibrary" }
        
        if (-not $docLibraries -or $docLibraries.Count -eq 0) {
            Write-Log "No document libraries found for site: $($Site.DisplayName)" -Level Warning
            return @{
                Files = @()
                FolderSizes = @{}
                TotalFiles = 0
                TotalSizeGB = 0
            }
        }
        
        Write-Log "Found $($docLibraries.Count) document libraries in site: $($Site.DisplayName)" -Level Success
        
        $allFiles = [System.Collections.Generic.List[psobject]]::new()
        $systemFiles = [System.Collections.Generic.List[psobject]]::new()
        $folderSizes = @{}
        $totalFiles = 0
        $listIndex = 0
        
        foreach ($list in $docLibraries) {
            $listIndex++
            $percentComplete = [math]::Round(($listIndex / $docLibraries.Count) * 100, 1)
            # Improved progress bar for library processing
            Show-Progress -Activity "Analyzing Document Libraries" -Status "Processing: $($list.DisplayName) | Files found: $totalFiles ($listIndex/$($docLibraries.Count))" -PercentComplete $percentComplete -CurrentOperation "$listIndex of $($docLibraries.Count) libraries"
            
            try {
                # Use SharePoint List API to get all items with drive item details
                $uri = "/v1.0/sites/$($Site.Id)/lists/$($list.Id)/items?expand=fields,driveItem&`$top=200"
                $nextLink = $uri
                $filesInThisList = 0
                while ($nextLink) {
                    try {
                        $resp = Invoke-WithRetry -ScriptBlock {
                            Invoke-MgGraphRequest -Method GET -Uri $nextLink
                        } -Activity "Get list items"
                        if (-not $resp.value -or $resp.value.Count -eq 0) {
                            break
                        }
                        foreach ($item in $resp.value) {
                            if ($item.driveItem -and $item.driveItem.file) {
                                $isSystem = $false
                                $fileName = $item.driveItem.name
                                $systemFilePatterns = @(
                                    "~$*", ".tmp", "thumbs.db", ".ds_store", "desktop.ini", ".git*", ".svn*", "*.lnk", "_vti_*", "forms/", "web.config", "*.aspx", "*.master"
                                )
                                foreach ($pattern in $systemFilePatterns) {
                                    if ($fileName -like $pattern) {
                                        $isSystem = $true
                                        break
                                    }
                                }
                                $parentPath = if ($item.driveItem.parentReference) { $item.driveItem.parentReference.path } else { '' }
                                $fullPath = ($parentPath + "/" + $item.driveItem.name).Replace("//", "/")
                                $pathLength = $fullPath.Length
                                $fileObj = [PSCustomObject]@{
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
                                if ($isSystem) {
                                    $systemFiles.Add($fileObj) | Out-Null
                                    continue
                                }
                                $allFiles.Add($fileObj) | Out-Null
                                $folderPath = $parentPath
                                if (-not $folderSizes.ContainsKey($folderPath)) {
                                    $folderSizes[$folderPath] = 0
                                }
                                $folderSizes[$folderPath] += $item.driveItem.size
                                $filesInThisList++
                                $totalFiles++
                                $currentFileName = if ($item.driveItem.name.Length -gt 50) { $item.driveItem.name.Substring(0, 47) + "..." } else { $item.driveItem.name }
                                Show-Progress -Activity "Analyzing Document Libraries" -Status "Processing: $($list.DisplayName) | Files found: $totalFiles | Current: $currentFileName ($filesInThisList)" -PercentComplete $percentComplete -CurrentOperation "$listIndex of $($docLibraries.Count) libraries"
                            }
                        }
                        # Check for next page link
                        $nextLink = $resp.'@odata.nextLink'
                        if ($nextLink) {
                            Start-Sleep -Milliseconds (Get-Random -Minimum 100 -Maximum 300)
                        }
                    }
                    catch {
                        Write-Log "Failed to process page of list items: $_" -Level Error
                        $nextLink = $null
                    }
                }
                
                Write-Log "Processed library: $($list.DisplayName) - Found $filesInThisList files" -Level Success
            }
            catch {
                Write-Log "Failed to process library $($list.DisplayName): $_" -Level Error
            }
        }
        
        Stop-Progress -Activity "Analyzing Document Libraries"
        
        $recycleBinSizeGB = 0
        try {
            # Try multiple approaches to get recycle bin storage
            $recycleBinSize = 0
            
            # Approach 1: Try via Graph API with different endpoints
            try {
                $recycleBinItems = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/sites/$($Site.Id)/drive/special/recycle" -ErrorAction SilentlyContinue
                if ($recycleBinItems -and $recycleBinItems.quota -and $recycleBinItems.quota.used) {
                    $recycleBinSize = $recycleBinItems.quota.used
                }
            } catch { }
            
            # Approach 2: Try alternative recycle bin endpoint
            if ($recycleBinSize -eq 0) {
                try {
                    $recycleBinItems = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/sites/$($Site.Id)/recyclebin" -ErrorAction SilentlyContinue
                    if ($recycleBinItems -and $recycleBinItems.value) {
                        $recycleBinSize = ($recycleBinItems.value | Measure-Object -Property size -Sum).Sum
                    }
                } catch { }
            }
            
            # Approach 3: Try drive-specific recycle bin
            if ($recycleBinSize -eq 0) {
                try {
                    $drives = Get-MgSiteDrive -SiteId $Site.Id -ErrorAction SilentlyContinue
                    foreach ($drive in $drives) {
                        try {
                            $recycleBinItems = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/drives/$($drive.Id)/special/recycle" -ErrorAction SilentlyContinue
                            if ($recycleBinItems -and $recycleBinItems.quota -and $recycleBinItems.quota.used) {
                                $recycleBinSize += $recycleBinItems.quota.used
                            }
                        } catch { }
                    }
                } catch { }
            }
            
            $recycleBinSizeGB = [math]::Round($recycleBinSize / 1GB, 2)
            
        } catch {
            Write-Log "Could not access recycle bin for site: $($Site.DisplayName)" -Level Debug
        }
        $result = @{
            Files = $allFiles
            SystemFiles = $systemFiles
            FolderSizes = $folderSizes
            TotalFiles = $allFiles.Count
            TotalSizeGB = [math]::Round(($allFiles | Measure-Object -Property Size -Sum).Sum / 1GB, 2)
            SystemSizeGB = [math]::Round(($systemFiles | Measure-Object -Property Size -Sum).Sum / 1GB, 2)
            RecycleBinSizeGB = $recycleBinSizeGB
            TotalWithRecycleBinGB = [math]::Round((($allFiles | Measure-Object -Property Size -Sum).Sum + ($systemFiles | Measure-Object -Property Size -Sum).Sum + ($recycleBinSizeGB * 1GB)) / 1GB, 2)
        }
        
        Write-Log "Completed site: $($Site.DisplayName) - Files: $($result.TotalFiles), Size: $($result.TotalSizeGB)GB" -Level Success
        
        # Get user access (site permissions)
        $siteUsers = @()
        $externalGuests = @()
        try {
            $permissions = Invoke-WithRetry -ScriptBlock { Get-MgSitePermission -SiteId $Site.Id -All -ErrorAction SilentlyContinue } -Activity "Get site permissions"
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
                    Write-Log "Could not process permission: $_" -Level Warning
                }
            }
        }
        catch {
            Write-Log "Could not get permissions for site $($Site.DisplayName): $_" -Level Warning
        }
        
        $result.Users = $siteUsers
        $result.ExternalGuests = $externalGuests
        
        Write-Log "Completed storage and access for site: $($Site.DisplayName) - Users: $($siteUsers.Count), Guests: $($externalGuests.Count)" -Level Success
        
        return $result
    }
    catch {
        Write-Log "Failed to get site storage and access info: $_" -Level Error
        return @{
            Files = @()
            FolderSizes = @{}
            TotalFiles = 0
            TotalSizeGB = 0
            Users = @()
            ExternalGuests = @()
        }
    }
}

function Get-SiteStorageBatch {
    param($sites)
    $details = @()
    $siteIds = $sites | ForEach-Object { $_.Id }
    $batchSize = 20
    for ($i = 0; $i -lt $siteIds.Count; $i += $batchSize) {
        $batch = $siteIds[$i..([Math]::Min($i+$batchSize-1, $siteIds.Count-1))]
        $batchRequests = @()
        $reqId = 1
        foreach ($siteId in $batch) {
            $batchRequests += @{ id = "$reqId"; method = "GET"; url = "/sites/$siteId/storage" }
            $reqId++
        }
        $body = @{ requests = $batchRequests } | ConvertTo-Json -Depth 6
        $result = Invoke-MgGraphRequest -Method POST -Uri "/v1.0/`$batch" -Body $body -ContentType "application/json"
        foreach ($resp in $result.responses) {
            if ($resp.status -eq 200 -and $resp.body) {
                $originalSiteId = $batch[$([int]$resp.id - 1)]
                $details += [PSCustomObject]@{
                    SiteId       = $originalSiteId
                    QuotaGB      = [math]::Round($resp.body.quota.total / 1GB, 2)
                    UsedGB       = [math]::Round($resp.body.quota.used / 1GB, 2)
                    RecycleBinGB = [math]::Round($resp.body.storage.recycleBin.total / 1GB, 2)
                }
            }
        }
    }
    return $details
}
#endregion

#region Excel Report Generation
function Test-ExcelFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    try {
        Import-Module ImportExcel -ErrorAction SilentlyContinue
        if (-not (Test-Path $FilePath)) {
            Write-Log "Excel file not found: $FilePath" -Level Error
            return $false
        }
        
        $sheets = Get-ExcelSheetInfo -Path $FilePath
        Write-Log "Excel file validation: $FilePath" -Level Info
        Write-Log "Worksheets found: $($sheets.Name -join ', ')" -Level Info
        
        $valid = $true
        foreach ($sheet in $sheets) {
            $data = Import-Excel -Path $FilePath -WorksheetName $sheet.Name
            $rowCount = $data.Count
            Write-Log "Worksheet '$($sheet.Name)': $rowCount rows" -Level Info
            
            $blankCells = ($data | ForEach-Object { $_.PSObject.Properties | Where-Object { -not $_.Value } }).Count
            if ($blankCells -gt 0) {
                Write-Log "Worksheet '$($sheet.Name)' has $blankCells blank cells" -Level Warning
            }
            
            # Check for suspicious strings (e.g., XML errors)
            $badStrings = ($data | ForEach-Object { $_.PSObject.Properties | Where-Object { $_.Value -match '[\x00-\x08\x0B\x0C\x0E-\x1F]' } }).Count
            if ($badStrings -gt 0) {
                Write-Log "Worksheet '$($sheet.Name)' has $badStrings cells with invalid characters" -Level Error
                $valid = $false
            }
        }
        
        if ($valid) {
            Write-Log "Excel file passed validation." -Level Success
        } else {
            Write-Log "Excel file failed validation. See above for details." -Level Error
        }
        
        return $valid
    } catch {
        Write-Log "Excel file validation error: $_" -Level Error
        return $false
    }
}

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
        Write-Log "Could not show save dialog. Using default filename in current directory." -Level Warning
        return $DefaultFileName
    }
}

function Export-ExcelWorksheet {
    param(
        [Parameter(Mandatory=$true)]
        [object]$Data,
        
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [Parameter(Mandatory=$true)]
        [string]$WorksheetName,
        
        [string]$Title,
        
        [string]$TableStyle = "Medium2",
        
        [switch]$AutoSize,
        
        [switch]$BoldTopRow,
        
        [switch]$FreezeTopRow,
        
        [string]$HeaderColor = "Black",
        
        [string]$HeaderTextColor = "Yellow",
        
        [string]$ChartType,
        
        [string]$ChartColumn,
        
        [string]$ChartTitle
    )
    if (-not $Data -or $Data.Count -eq 0) {
        Write-Log "No data to export for worksheet '$WorksheetName'. Skipping export." -Level Warning
        return
    }
    try {
        $params = @{
            Path = $Path
            WorksheetName = $WorksheetName
            TableStyle = $TableStyle
            AutoSize = $AutoSize
            BoldTopRow = $BoldTopRow
            FreezeTopRow = $FreezeTopRow
            HeaderColor = $HeaderColor
            HeaderTextColor = $HeaderTextColor
        }
        if ($Title) { $params.Title = $Title }
        Import-Module ImportExcel -ErrorAction SilentlyContinue
        $Data | Export-Excel @params
        if ($ChartType -and $ChartColumn -and $ChartTitle) {
            if ($Data.Count -gt 0) {
                # Add chart logic here if needed
            }
        }
    } catch {
        Write-Log "Failed to export worksheet '$WorksheetName': $_" -Level Error
    }
}
#endregion

#region Main Processing Functions
function Get-SiteSummaries {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Sites,
        
        [int]$ParallelLimit = 4
    )
    
    Write-Log "Getting site summaries for $($Sites.Count) sites..." -Level Info
    
    # Filter sites to include both SharePoint and OneDrive personal sites with valid Ids
    $filteredSites = $Sites | Where-Object {
        $_.Id -and ($_.Id -ne "")
    }
    
    # Separate SharePoint sites from OneDrive personal sites
    $sharePointSites = $filteredSites | Where-Object {
        $_.WebUrl -notlike "*-my.sharepoint.com/personal/*" -and
        $_.WebUrl -notlike "*/personal/*" -and
        $_.WebUrl -notlike "*mysites*" -and
        $_.Name -notlike "*OneDrive*" -and
        $_.DisplayName -notlike "*OneDrive*"
    }
    
    $oneDrivePersonalSites = $filteredSites | Where-Object {
        $_.WebUrl -like "*-my.sharepoint.com/personal/*" -or
        $_.WebUrl -like "*/personal/*" -or
        $_.WebUrl -like "*mysites*" -or
        $_.Name -like "*OneDrive*" -or
        $_.DisplayName -like "*OneDrive*"
    }
    
    Write-Log "Filtered SharePoint library sites to process: $($sharePointSites.Count)" -Level Info
    Write-Log "Filtered OneDrive personal sites to process: $($oneDrivePersonalSites.Count)" -Level Info
    Write-Log "Total sites to process: $($filteredSites.Count)" -Level Info
    
    $filteredSites | Select-Object -First 10 | ForEach-Object {
        Write-Log "Id: '$($_.Id)', DisplayName: '$($_.DisplayName)', WebUrl: '$($_.WebUrl)'" -Level Debug
    }
    
    $sitesToProcess = if ($TestMode) { $filteredSites | Select-Object -First 5 } else { $filteredSites }
    
    # Start timer for parallel scan
    $stepTimer = [System.Diagnostics.Stopwatch]::StartNew()
    
    # Process sites in parallel
    $siteSummaries = $sitesToProcess | ForEach-Object -Parallel {
        $site = $_
        $siteId = $site.Id
        $displayName = $site.DisplayName
        $webUrl = $site.WebUrl
        
        if (-not $siteId) {
            Write-Log "Parallel block received null or invalid site Id" -Level Error
            return $null
        }
        
        $siteType = if ($webUrl -like "*-my.sharepoint.com/personal/*" -or $webUrl -like "*/personal/*" -or $webUrl -like "*mysites*" -or $displayName -like "*OneDrive*") { 
            "OneDrive Personal" 
        } else { 
            "SharePoint Site" 
        }
        
        $storageGB = 0
        $recycleBinGB = 0
        
        try {
            # Approach 1: Try to get site storage usage via Graph API
            try {
                $siteStorageInfo = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/sites/$siteId/drive" -ErrorAction SilentlyContinue
                if ($siteStorageInfo -and $siteStorageInfo.quota -and $siteStorageInfo.quota.used) {
                    $storageGB = [math]::Round($siteStorageInfo.quota.used / 1GB, 2)
                }
            } catch { }
            
            # Approach 2: If no storage from site drive, try all drives for the site
            if ($storageGB -eq 0) {
                $drives = Get-MgSiteDrive -SiteId $siteId -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                
                foreach ($drive in $drives) {
                    try {
                        $driveInfo = Get-MgDrive -DriveId $drive.Id -ErrorAction SilentlyContinue
                        if ($driveInfo -and $driveInfo.Quota -and $driveInfo.Quota.Used) {
                            $storageGB += [math]::Round($driveInfo.Quota.Used / 1GB, 2)
                        }
                    } catch {
                        # If quota not available, try to calculate from root folder size
                        try {
                            $rootFolder = Get-MgDriveItem -DriveId $drive.Id -DriveItemId "root" -ErrorAction SilentlyContinue
                            if ($rootFolder -and $rootFolder.Size) {
                                $storageGB += [math]::Round($rootFolder.Size / 1GB, 2)
                            }
                        } catch {
                            # Silently handle errors
                        }
                    }
                    
                    # Try to get recycle bin size for this drive
                    try {
                        $recycleBinItems = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/drives/$($drive.Id)/special/recycle" -ErrorAction SilentlyContinue
                        if ($recycleBinItems -and $recycleBinItems.size) {
                            $recycleBinGB += [math]::Round($recycleBinItems.size / 1GB, 2)
                        }
                    } catch { }
                }
            }
            
            # Approach 3: Try site-level recycle bin if drive-level failed
            if ($recycleBinGB -eq 0) {
                try {
                    $siteRecycleBin = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/sites/$siteId/drive/special/recycle" -ErrorAction SilentlyContinue
                    if ($siteRecycleBin -and $siteRecycleBin.size) {
                        $recycleBinGB = [math]::Round($siteRecycleBin.size / 1GB, 2)
                    }
                } catch { }
            }
            
        } catch {
            # Silently handle any unexpected errors
        }
        
        return [PSCustomObject]@{
            Site = $site
            SiteId = $siteId
            SiteName = $displayName
            SiteType = $siteType
            StorageGB = $storageGB
            RecycleBinGB = $recycleBinGB
            TotalStorageGB = ($storageGB + $recycleBinGB)
            StorageBytes = $storageGB * 1GB
            WebUrl = $webUrl
            IsOneDrive = ($siteType -eq "OneDrive Personal")
        }
    } -ThrottleLimit $ParallelLimit
    
    # Filter out null results
    $siteSummaries = $siteSummaries | Where-Object { $_ -ne $null }
    
    $stepTimer.Stop()
    $elapsed = $stepTimer.Elapsed
    Write-Log "Parallel Site Summary Scan completed in $($elapsed.TotalSeconds) seconds ($($elapsed.TotalMinutes) min)" -Level Info
    
    # Clear progress bar
    Stop-Progress -Activity "Scanning SharePoint Sites"
    
    # Output summary table to console matching SharePoint Admin Centre
    Write-Host "`nActive Sites Summary:" -ForegroundColor Cyan
    $header = "{0,-30} {1,-40} {2,12} {3,-20}" -f "Site name", "URL", "Storage (GB)", "Primary admin"
    Write-Host $header -ForegroundColor White
    Write-Host ("-" * 110) -ForegroundColor DarkGray
    foreach ($site in $siteSummaries) {
        $siteName = $site.DisplayName
        $url = $site.WebUrl
        $storage = if ($site.StorageGB) { [math]::Round($site.StorageGB,2) } else { "-" }
        $admin = if ($site.PrimaryAdmin) { $site.PrimaryAdmin } else { "Group owners" }
        $row = "{0,-30} {1,-40} {2,12} {3,-20}" -f $siteName, $url, $storage, $admin
        Write-Host $row -ForegroundColor Gray
    }
    
    return $siteSummaries
}

function Get-SiteDetails {
    param(
        [Parameter(Mandatory=$true)]
        [array]$SiteSummaries,
        
        [array]$TopSites,
        
        [int]$ParallelLimit = 4
    )
    
    Write-Log "Getting detailed information for sites..." -Level Info
    
    $allSiteSummaries = @()
    $allTopFiles = @()
    $allTopFolders = @()
    $siteStorageStats = @{}
    $sitePieCharts = @{}
    
    # Process each site for detailed analysis
    $detailProcessedCount = 0
    $totalSites = $SiteSummaries.Count
    
    foreach ($siteSummary in $SiteSummaries) {
        $site = $siteSummary.Site
        $isTopSite = $TopSites.SiteId -contains $site.Id
        $detailProcessedCount++
        $percentComplete = [math]::Round(($detailProcessedCount / $totalSites) * 100, 1)
        
        # Progress bar for detailed analysis
        Show-Progress -Activity "Analyzing Sites for Detailed Data" -Status "Processing: $($site.DisplayName)" -PercentComplete $percentComplete -CurrentOperation "$detailProcessedCount of $totalSites sites"
        
        # Get site owners and members (with improved error handling)
        $userAccess = @{ Owners = @(); Members = @() }
        try {
            $userAccess = Get-SiteUserAccessSummary -Site $site
        }
        catch {
            Write-Log "Failed to get user access for site $($site.DisplayName): $_" -Level Error
        }
        
        if ($isTopSite) {
            # For top sites, perform detailed file and folder analysis
            try {
                $fileData = Get-FileData -Site $site
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
                $allSiteSummaries += [PSCustomObject]@{
                    SiteName = $site.DisplayName
                    SiteUrl = $site.WebUrl
                    SiteType = $siteSummary.SiteType
                    TotalFiles = $fileData.TotalFiles
                    TotalSizeGB = $siteSummary.StorageGB
                    RecycleBinGB = $siteSummary.RecycleBinGB
                    TotalStorageGB = $siteSummary.TotalStorageGB
                    TotalFolders = $fileData.FolderSizes.Count
                    UniquePermissionLevels = $null
                    OwnersCount = $userAccess.Owners.Count
                    MembersCount = $userAccess.Members.Count
                    Owners = $userAccess.Owners
                    Members = $userAccess.Members
                    ExternalGuests = $fileData.ExternalGuests
                    ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
            }
            catch {
                Write-Log "Failed to analyze files and folders for site $($site.DisplayName): $_" -Level Error
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
                RecycleBinGB = $siteSummary.RecycleBinGB
                TotalStorageGB = $siteSummary.TotalStorageGB
                TotalFolders = $null
                UniquePermissionLevels = $null
                OwnersCount = $userAccess.Owners.Count
                MembersCount = $userAccess.Members.Count
                Owners = $userAccess.Owners
                Members = $userAccess.Members
                ExternalGuests = @()
                ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
    }
    
    # Clear progress bar
    Stop-Progress -Activity "Analyzing Sites for Detailed Data"
    
    return @{
        AllSiteSummaries = $allSiteSummaries
        AllTopFiles = $allTopFiles
        AllTopFolders = $allTopFolders
        SiteStorageStats = $siteStorageStats
        SitePieCharts = $sitePieCharts
    }
}

function Export-ComprehensiveExcelReport {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ExcelFileName,
        
        [Parameter(Mandatory=$true)]
        [array]$SiteSummaries,
        
        [Parameter(Mandatory=$true)]
        [array]$AllSiteSummaries,
        
        [Parameter(Mandatory=$true)]
        [array]$AllTopFiles,
        
        [Parameter(Mandatory=$true)]
        [array]$AllTopFolders,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$SiteStorageStats,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$SitePieCharts,
        
        [Parameter(Mandatory=$false)]
        [array]$ComprehensiveSiteDetails = @()
    )
    
    Write-Log "Generating Excel report..." -Level Info
    
    # Remove any existing Excel file to avoid conflicts
    if (Test-Path $ExcelFileName) {
        Remove-Item $ExcelFileName -Force -ErrorAction SilentlyContinue
    }
    
    Write-Log "Creating Excel report with $($AllSiteSummaries.Count) site summaries..." -Level Info
    
    try {
        # Create comprehensive SharePoint Storage Pie Chart Data (including recycle bins and personal sites)
        Show-Progress -Activity "Generating Excel Report" -Status "Creating comprehensive storage analysis..." -PercentComplete 5

        $comprehensiveStorageData = @()
        $totalTenantStorage = 0
        $recycleBinStorage = 0
        $personalOneDriveStorage = 0
        $sharePointSitesStorage = 0

        # Calculate storage breakdown using AllSiteSummaries which has the correct data
        foreach ($siteSummary in $AllSiteSummaries) {
            $storageGB = if ($siteSummary.TotalSizeGB) { $siteSummary.TotalSizeGB } else { 0 }
            $recycleBinGB = if ($siteSummary.RecycleBinGB) { $siteSummary.RecycleBinGB } else { 0 }
            $totalStorageGB = if ($siteSummary.TotalStorageGB) { $siteSummary.TotalStorageGB } else { $storageGB + $recycleBinGB }
            
            $totalTenantStorage += $storageGB
            $recycleBinStorage += $recycleBinGB
            
            if ($siteSummary.SiteType -eq "OneDrive Personal") {
                $personalOneDriveStorage += $storageGB
            } else {
                $sharePointSitesStorage += $storageGB
            }
            
            $comprehensiveStorageData += [PSCustomObject]@{
                Category = if ($siteSummary.SiteType -eq "OneDrive Personal") { "Personal OneDrive" } else { "SharePoint Site" }
                SiteName = $siteSummary.SiteName
                StorageGB = $storageGB
                RecycleBinGB = $recycleBinGB
                TotalStorageGB = $totalStorageGB
                Percentage = if (($totalTenantStorage + $recycleBinStorage) -gt 0) { [math]::Round(($totalStorageGB / ($totalTenantStorage + $recycleBinStorage)) * 100, 2) } else { 0 }
                SiteUrl = $siteSummary.SiteUrl
                SiteType = $siteSummary.SiteType
            }
        }



        # Create pie chart summary data
        $pieChartData = @(
            [PSCustomObject]@{
                Category = "SharePoint Sites"
                StorageGB = $sharePointSitesStorage
                Percentage = if ($totalTenantStorage -gt 0) { [math]::Round(($sharePointSitesStorage / $totalTenantStorage) * 100, 2) } else { 0 }
                SiteCount = ($AllSiteSummaries | Where-Object { $_.SiteType -ne "OneDrive Personal" }).Count
            },
            [PSCustomObject]@{
                Category = "Personal OneDrive"
                StorageGB = $personalOneDriveStorage
                Percentage = if ($totalTenantStorage -gt 0) { [math]::Round(($personalOneDriveStorage / $totalTenantStorage) * 100, 2) } else { 0 }
                SiteCount = ($AllSiteSummaries | Where-Object { $_.SiteType -eq "OneDrive Personal" }).Count
            },
            [PSCustomObject]@{
                Category = "Recycle Bins"
                StorageGB = $recycleBinStorage
                Percentage = if (($totalTenantStorage + $recycleBinStorage) -gt 0) { [math]::Round(($recycleBinStorage / ($totalTenantStorage + $recycleBinStorage)) * 100, 2) } else { 0 }
                SiteCount = "All Sites"
            }
        )
                StorageGB = $recycleBinStorage
                Percentage = if (($totalTenantStorage + $recycleBinStorage) -gt 0) { [math]::Round(($recycleBinStorage / ($totalTenantStorage + $recycleBinStorage)) * 100, 2) } else { 0 }
                SiteCount = "All Sites"
            }
        )

        # Use the comprehensive site details passed from main function, or build if not provided
        if ($ComprehensiveSiteDetails -and $ComprehensiveSiteDetails.Count -gt 0) {
            Write-Log "Using pre-built comprehensive site details with $($ComprehensiveSiteDetails.Count) sites" -Level Info
            # Update with user access counts from AllSiteSummaries
            foreach ($detail in $ComprehensiveSiteDetails) {
                $matchingSiteDetail = $AllSiteSummaries | Where-Object { $_.SiteName -eq $detail.'Site name' } | Select-Object -First 1
                if ($matchingSiteDetail) {
                    $detail.OwnersCount = $matchingSiteDetail.OwnersCount
                    $detail.MembersCount = $matchingSiteDetail.MembersCount
                }
            }
        } else {
            Write-Log "Building comprehensive site details from SiteSummaries..." -Level Warning
            $ComprehensiveSiteDetails = @()
            foreach ($siteSummary in $SiteSummaries) {
                $site = $siteSummary.Site
                $ComprehensiveSiteDetails += [PSCustomObject]@{
                    'Site name' = $site.DisplayName
                    'URL' = $site.WebUrl
                    'Teams' = $site.Teams
                    'Channel sites' = $site.ChannelSites
                    'IBMode' = $site.IBMode
                    'Storage used (GB)' = $siteSummary.StorageGB
                    'Recycle bin (GB)' = $siteSummary.RecycleBinGB
                    'Total storage (GB)' = $siteSummary.TotalStorageGB
                    'Primary admin' = $site.PrimaryAdmin
                    'Hub' = $site.Hub
                    'Template' = $site.Template
                    'Last activity (UTC)' = $site.LastActivityUTC
                    'Date created' = $site.CreatedDate
                    'Created by' = $site.CreatedBy
                    'Storage limit (GB)' = $site.StorageLimitGB
                    'Storage used (%)' = $site.StorageUsedPercent
                    'Microsoft 365 group' = $site.Microsoft365Group
                    'Files viewed or edited' = $site.FilesViewedOrEdited
                    'Page views' = $site.PageViews
                    'Page visits' = $site.PageVisits
                    'Files' = $site.Files
                    'Sensitivity' = $site.Sensitivity
                    'External sharing' = $site.ExternalSharing
                    'OwnersCount' = $siteSummary.OwnersCount
                    'MembersCount' = $siteSummary.MembersCount
                    'ReportDate' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
            }
        }

        # Add OneDrive Personal Sites Combined Summary as a single entry
        $oneDriveSites = $SiteSummaries | Where-Object { $_.IsOneDrive }
        if ($oneDriveSites.Count -gt 0) {
            $totalOneDriveStorage = ($oneDriveSites | Measure-Object -Property StorageGB -Sum).Sum
            $totalOneDriveSiteCount = $oneDriveSites.Count
            
            $ComprehensiveSiteDetails += [PSCustomObject]@{
                'Site name' = "All OneDrive Personal Sites (Combined)"
                'URL' = "Multiple OneDrive Personal Sites"
                'Teams' = ""
                'Channel sites' = ""
                'IBMode' = ""
                'Storage used (GB)' = $totalOneDriveStorage
                'Recycle bin (GB)' = ($oneDriveSites | Measure-Object -Property RecycleBinGB -Sum).Sum
                'Total storage (GB)' = ($oneDriveSites | Measure-Object -Property TotalStorageGB -Sum).Sum
                'Primary admin' = "Individual Users"
                'Hub' = ""
                'Template' = "OneDrive Personal"
                'Last activity (UTC)' = ""
                'Date created' = ""
                'Created by' = ""
                'Storage limit (GB)' = ""
                'Storage used (%)' = ""
                'Microsoft 365 group' = ""
                'Files viewed or edited' = ""
                'Page views' = ""
                'Page visits' = ""
                'Files' = ""
                'Sensitivity' = ""
                'External sharing' = ""
                'OwnersCount' = $totalOneDriveSiteCount
                'MembersCount' = 0
                'ReportDate' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            
            Write-Log "Added OneDrive Personal Sites combined entry: $totalOneDriveSiteCount sites with $totalOneDriveStorage GB total storage" -Level Info
        }

        # Create combined data for main worksheet: comprehensive site details + tenant pie chart summary
        $combinedMainWorksheetData = @()
        
        # Add the pie chart summary data at the top
        $combinedMainWorksheetData += [PSCustomObject]@{
            'Site name' = "=== TENANT STORAGE BREAKDOWN ==="
            'URL' = ""
            'Teams' = ""
            'Channel sites' = ""
            'IBMode' = ""
            'Storage used (GB)' = ""
            'Recycle bin (GB)' = ""
            'Total storage (GB)' = ""
            'Primary admin' = ""
            'Hub' = ""
            'Template' = ""
            'Last activity (UTC)' = ""
            'Date created' = ""
            'Created by' = ""
            'Storage limit (GB)' = ""
            'Storage used (%)' = ""
            'Microsoft 365 group' = ""
            'Files viewed or edited' = ""
            'Page views' = ""
            'Page visits' = ""
            'Files' = ""
            'Sensitivity' = ""
            'External sharing' = ""
            'OwnersCount' = ""
            'MembersCount' = ""
            'ReportDate' = ""
        }
        
        # Add pie chart data
        foreach ($chartData in $pieChartData) {
            $combinedMainWorksheetData += [PSCustomObject]@{
                'Site name' = "📊 $($chartData.Category)"
                'URL' = ""
                'Teams' = ""
                'Channel sites' = ""
                'IBMode' = ""
                'Storage used (GB)' = $chartData.StorageGB
                'Recycle bin (GB)' = ""
                'Total storage (GB)' = $chartData.StorageGB
                'Primary admin' = ""
                'Hub' = ""
                'Template' = ""
                'Last activity (UTC)' = ""
                'Date created' = ""
                'Created by' = ""
                'Storage limit (GB)' = ""
                'Storage used (%)' = "$($chartData.Percentage)%"
                'Microsoft 365 group' = ""
                'Files viewed or edited' = ""
                'Page views' = ""
                'Page visits' = ""
                'Files' = ""
                'Sensitivity' = ""
                'External sharing' = ""
                'OwnersCount' = if ($chartData.SiteCount -eq "All Sites") { "" } else { $chartData.SiteCount }
                'MembersCount' = ""
                'ReportDate' = ""
            }
        }
        
        # Add separator
        $combinedMainWorksheetData += [PSCustomObject]@{
            'Site name' = "=== INDIVIDUAL SITES ==="
            'URL' = ""
            'Teams' = ""
            'Channel sites' = ""
            'IBMode' = ""
            'Storage used (GB)' = ""
            'Recycle bin (GB)' = ""
            'Total storage (GB)' = ""
            'Primary admin' = ""
            'Hub' = ""
            'Template' = ""
            'Last activity (UTC)' = ""
            'Date created' = ""
            'Created by' = ""
            'Storage limit (GB)' = ""
            'Storage used (%)' = ""
            'Microsoft 365 group' = ""
            'Files viewed or edited' = ""
            'Page views' = ""
            'Page visits' = ""
            'Files' = ""
            'Sensitivity' = ""
            'External sharing' = ""
            'OwnersCount' = ""
            'MembersCount' = ""
            'ReportDate' = ""
        }
        
        # Add all site details
        $combinedMainWorksheetData += $ComprehensiveSiteDetails

        # Export merged worksheet with comprehensive site details and pie chart data
        $combinedMainWorksheetData | Export-Excel -Path $ExcelFileName -WorksheetName "SharePoint Storage Overview" -AutoSize -TableStyle "Light1" -Title "SharePoint Tenant Storage Overview (Includes Pie Chart & Recycle Bins)"
        
        Write-Log "SharePoint Storage Overview worksheet created with $($combinedMainWorksheetData.Count) entries (pie chart + sites)" -Level Info

        # Worksheet 2: User & Group Access Overview
        $userAccessRows = @()
        foreach ($siteDetail in $AllSiteSummaries) {
            $siteName = $siteDetail.SiteName
            $siteUrl = $siteDetail.SiteUrl
            # Owners
            if ($siteDetail.OwnersCount -gt 0 -and $siteDetail.Owners) {
                foreach ($owner in $siteDetail.Owners) {
                    $row = [PSCustomObject]@{
                        'User/Group' = $owner.DisplayName
                        'Email' = $owner.UserEmail
                        'Type' = $owner.UserType
                        'Role' = $owner.Role
                        'Site' = $siteName
                        'Site URL' = $siteUrl
                        'Access' = 'Owner/Admin'
                        'Object' = 'Site'
                    }
                    $userAccessRows += $row
                }
            }
            # Members
            if ($siteDetail.MembersCount -gt 0 -and $siteDetail.Members) {
                foreach ($member in $siteDetail.Members) {
                    $row = [PSCustomObject]@{
                        'User/Group' = $member.DisplayName
                        'Email' = $member.UserEmail
                        'Type' = $member.UserType
                        'Role' = $member.Role
                        'Site' = $siteName
                        'Site URL' = $siteUrl
                        'Access' = $member.Role
                        'Object' = 'Site'
                    }
                    $userAccessRows += $row
                }
            }
            # External Guests
            if ($siteDetail.ExternalGuests) {
                foreach ($guest in $siteDetail.ExternalGuests) {
                    $row = [PSCustomObject]@{
                        'User/Group' = $guest.UserName
                        'Email' = $guest.UserEmail
                        'Type' = 'External Guest'
                        'Role' = $guest.AccessType
                        'Site' = $siteName
                        'Site URL' = $siteUrl
                        'Access' = $guest.AccessType
                        'Object' = 'Site'
                        'Highlight' = 'Red'
                    }
                    $userAccessRows += $row
                }
            }
        }
        # Export user/group access worksheet
        if ($userAccessRows -and $userAccessRows.Count -gt 0) {
            $userAccessRows | Export-Excel -Path $ExcelFileName -WorksheetName "User & Group Access Overview" -AutoSize -TableStyle "Medium2" -Title "All Users, Groups, Members, and Guests with Access"
            Write-Log "User & Group Access Overview worksheet created with $($userAccessRows.Count) access entries" -Level Info
        } else {
            # Create empty worksheet with headers if no data
            $emptyUserAccess = @([PSCustomObject]@{
                'User/Group' = "No access data available"
                'Email' = ""
                'Type' = ""
                'Role' = ""
                'Site' = ""
                'Site URL' = ""
                'Access' = ""
                'Object' = ""
            })
            $emptyUserAccess | Export-Excel -Path $ExcelFileName -WorksheetName "User & Group Access Overview" -AutoSize -TableStyle "Medium2" -Title "All Users, Groups, Members, and Guests with Access"
            Write-Log "User & Group Access Overview worksheet created with placeholder data (no user access data found)" -Level Warning
        }

        # Worksheet 3: OneDrive Personal Sites Details
        $oneDrivePersonalRows = @()
        foreach ($siteSummary in $SiteSummaries | Where-Object { $_.IsOneDrive }) {
            $site = $siteSummary.Site
            $oneDrivePersonalRows += [PSCustomObject]@{
                'User Name' = if ($site.DisplayName -match "OneDrive - (.+)") { $matches[1] } else { $site.DisplayName }
                'Site Name' = $site.DisplayName
                'URL' = $site.WebUrl
                'Storage Used (GB)' = $siteSummary.StorageGB
                'Storage Used (%)' = if ($site.StorageLimitGB -and $site.StorageLimitGB -gt 0) { [math]::Round(($siteSummary.StorageGB / $site.StorageLimitGB) * 100, 2) } else { "Unknown" }
                'Storage Limit (GB)' = $site.StorageLimitGB
                'Last Activity (UTC)' = $site.LastActivityUTC
                'Date Created' = $site.CreatedDate
                'Files Count' = $site.Files
                'External Sharing' = $site.ExternalSharing
                'Site Type' = $siteSummary.SiteType
                'Report Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        
        if ($oneDrivePersonalRows.Count -gt 0) {
            $oneDrivePersonalRows | Export-Excel -Path $ExcelFileName -WorksheetName "OneDrive Personal Sites" -AutoSize -TableStyle "Light1" -Title "OneDrive Personal Sites Storage and Details"
            Write-Log "OneDrive Personal Sites worksheet created with $($oneDrivePersonalRows.Count) personal sites" -Level Info
        } else {
            Write-Log "No OneDrive Personal sites found to create worksheet" -Level Warning
        }

        # Export site-specific worksheets with pie charts and detailed analysis
        if ($SitePieCharts.Count -gt 0 -and $AllTopFiles -and $AllTopFolders) {
            foreach ($siteName in $SitePieCharts.Keys) {
                $safeName = Format-WorksheetName "$siteName Storage"
                $sitePieData = $SitePieCharts[$siteName]
                
                # Get top 20 files and folders for this site
                $topFiles = @()
                $topFolders = @()
                if ($AllTopFiles) {
                    $topFiles = $AllTopFiles | Where-Object { $_.SiteName -eq $siteName } | Sort-Object Size -Descending | Select-Object -First 20
                }
                if ($AllTopFolders) {
                    $topFolders = $AllTopFolders | Where-Object { $_.SiteName -eq $siteName } | Sort-Object SizeGB -Descending | Select-Object -First 20
                }
                
                # Create comprehensive site worksheet data with clear sections
                $siteWorksheetData = @()
                
                # Section 1: Site Storage Breakdown Pie Chart Data
                $siteWorksheetData += [PSCustomObject]@{
                    'Section' = "SITE STORAGE BREAKDOWN (PIE CHART DATA)"
                    'Category/Location' = "Location"
                    'Size (GB)' = "Size (GB)"
                    'Type' = "Chart Data"
                    'Details' = "Top 10 Folders by Size"
                }
                
                foreach ($folder in $sitePieData) {
                    $siteWorksheetData += [PSCustomObject]@{
                        'Section' = "Storage Breakdown"
                        'Category/Location' = $folder.Location
                        'Size (GB)' = $folder.SizeGB
                        'Type' = "Folder"
                        'Details' = "Folder Storage"
                    }
                }
                
                # Add spacer row
                $siteWorksheetData += [PSCustomObject]@{
                    'Section' = ""
                    'Category/Location' = ""
                    'Size (GB)' = ""
                    'Type' = ""
                    'Details' = ""
                }
                
                # Section 2: Top 20 Largest Files
                $siteWorksheetData += [PSCustomObject]@{
                    'Section' = "TOP 20 LARGEST FILES"
                    'Category/Location' = "File Name"
                    'Size (GB)' = "Size (GB)"
                    'Type' = "File Type"
                    'Details' = "File Path"
                }
                
                foreach ($file in $topFiles) {
                    $siteWorksheetData += [PSCustomObject]@{
                        'Section' = "Large Files"
                        'Category/Location' = $file.Name
                        'Size (GB)' = [math]::Round($file.Size / 1GB, 3)
                        'Type' = if ($file.Name -match '\.([^.]+)$') { $matches[1].ToUpper() } else { "Unknown" }
                        'Details' = $file.Path
                    }
                }
                
                # Add spacer row
                $siteWorksheetData += [PSCustomObject]@{
                    'Section' = ""
                    'Category/Location' = ""
                    'Size (GB)' = ""
                    'Type' = ""
                    'Details' = ""
                }
                
                # Section 3: Top 20 Largest Folders
                $siteWorksheetData += [PSCustomObject]@{
                    'Section' = "TOP 20 LARGEST FOLDERS"
                    'Category/Location' = "Folder Name"
                    'Size (GB)' = "Size (GB)"
                    'Type' = "Folder Type"
                    'Details' = "Folder Path"
                }
                
                foreach ($folder in $topFolders) {
                    $siteWorksheetData += [PSCustomObject]@{
                        'Section' = "Large Folders"
                        'Category/Location' = $folder.Name
                        'Size (GB)' = $folder.SizeGB
                        'Type' = "Folder"
                        'Details' = $folder.Path
                    }
                }
                
                # Export the comprehensive site worksheet
                if ($siteWorksheetData.Count -gt 0) {
                    $siteWorksheetData | Export-Excel -Path $ExcelFileName -WorksheetName $safeName -AutoSize -TableStyle "Light1" -Title "Comprehensive Storage Analysis for $siteName"
                }
            }
        }

        Write-Log "Excel report created successfully: $ExcelFileName" -Level Success
        return $true
    }
    catch {
        Write-Log "Failed to create Excel report: $_" -Level Error
        throw
    }
}
#endregion

function Initialize-Modules {
    $importErrorList = @()
    try { Import-Module Microsoft.Graph.Sites -ErrorAction Stop; Write-Log "Loaded Microsoft.Graph.Sites module" -Level Success } catch { $importErrorList += "Microsoft.Graph.Sites: $_" }
    try { Import-Module Microsoft.Graph.Files -ErrorAction Stop; Write-Log "Loaded Microsoft.Graph.Files module" -Level Success } catch { $importErrorList += "Microsoft.Graph.Files: $_" }
    try { Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue; Write-Log "Loaded Microsoft.Graph.Users module" -Level Success } catch { Write-Log "Microsoft.Graph.Users module not available: $_" -Level Warning; $importErrorList += "Microsoft.Graph.Users: $_" }
    try { Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop; Write-Log "Loaded Microsoft.Graph.Identity.DirectoryManagement module" -Level Success } catch { $importErrorList += "Microsoft.Graph.Identity.DirectoryManagement: $_" }
    try { Import-Module ImportExcel -ErrorAction Stop; Write-Log "Loaded ImportExcel module" -Level Success } catch { $importErrorList += "ImportExcel: $_" }
    if ($importErrorList.Count -gt 0) {
        Write-Log "Some modules failed to load. Attempting to force update..." -Level Warning
        foreach ($moduleError in $importErrorList) {
            $moduleName = $moduleError.Split(':')[0]
            Write-Log "Force updating module: $moduleName" -Level Info
            Install-OrUpdateModule -ModuleName $moduleName -MinimumVersion "2.0.0" -Force
        }
        Write-Log "Modules were force updated. Please restart PowerShell and run the script again." -Level Warning
        return $false
    }
    return $true
}

#region Main Function
function Main {
    try {
        Write-Log "SharePoint Tenant Storage & Access Report Generator" -Level Success
        Write-Log "=============================================" -Level Success
        
        # Initialize modules
        if (-not (Initialize-Modules)) {
            return
        }
        
        # Connect to Microsoft Graph
        if (-not $ClientId) { $ClientId = '278b9af9-888d-4344-93bb-769bdd739249' }
        if (-not $TenantId) { $TenantId = 'ca0711e2-e703-4f4e-9099-17d97863211c' }
        if (-not $CertificateThumbprint) { $CertificateThumbprint = '2E2502BB1EDB8F36CF9DE50936B283BDD22D5BAD' }
        Connect-ToGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
        
        # Get tenant name
        $script:tenantName = Get-TenantName
        
        # Create date string and filename
        $script:dateStr = Get-Date -Format "yyyyMMdd"
        $defaultFileName = "SharePointAudit-AllSites-$($script:tenantName)-$($script:dateStr).xlsx"
        
        # Get save file path from user dialog or use provided path
        if ($OutputPath) {
            $script:excelFileName = $OutputPath
        } else {
            $script:excelFileName = Get-SaveFileDialog -DefaultFileName $defaultFileName -Title "Save SharePoint Audit Report"
            if (-not $script:excelFileName) {
                Write-Log "User cancelled the save dialog. Exiting." -Level Info
                return
            }
        }
        
        # Get all SharePoint sites in the tenant
        $sites = Get-AllSharePointSites
        if ($sites.Count -eq 0) {
            Write-Log "No sites found. Exiting." -Level Warning
            return
        }
        
        Write-Log "Found $($sites.Count) total sites to analyze (including SharePoint sites and OneDrive personal sites)..." -Level Info
        
        # Get site summaries with storage information
        if (-not $ParallelLimit -or $ParallelLimit -lt 1) { $ParallelLimit = 1 }
        $siteSummaries = Get-SiteSummaries -Sites $sites -ParallelLimit $ParallelLimit
        if ($siteSummaries.Count -eq 0) {
            Write-Log "No valid site summaries were generated. Exiting." -Level Warning
            return
        }
        
        # Calculate site type breakdown
        $sharePointSites = $siteSummaries | Where-Object { -not $_.IsOneDrive }
        $oneDriveSites = $siteSummaries | Where-Object { $_.IsOneDrive }
        
        Write-Log "Processed $($siteSummaries.Count) sites: SharePoint Sites: $($sharePointSites.Count) | OneDrive Personal: $($oneDriveSites.Count)" -Level Info
        
        # Identify top 10 largest sites
        $topSites = $siteSummaries | Sort-Object StorageBytes -Descending | Select-Object -First 10
        
        # Build comprehensive site details with all admin CSV columns BEFORE getting detailed info
        Write-Log "Building comprehensive site details for Excel export..." -Level Info
        $comprehensiveSiteDetails = @()
        foreach ($siteSummary in $siteSummaries) {
            $site = $siteSummary.Site
            $comprehensiveSiteDetails += [PSCustomObject]@{
                'Site name' = $site.DisplayName
                'URL' = $site.WebUrl
                'Teams' = $site.Teams
                'Channel sites' = $site.ChannelSites
                'IBMode' = $site.IBMode
                'Storage used (GB)' = $siteSummary.StorageGB
                'Recycle bin (GB)' = $siteSummary.RecycleBinGB
                'Total storage (GB)' = $siteSummary.TotalStorageGB
                'Primary admin' = $site.PrimaryAdmin
                'Hub' = $site.Hub
                'Template' = $site.Template
                'Last activity (UTC)' = $site.LastActivityUTC
                'Date created' = $site.CreatedDate
                'Created by' = $site.CreatedBy
                'Storage limit (GB)' = $site.StorageLimitGB
                'Storage used (%)' = $site.StorageUsedPercent
                'Microsoft 365 group' = $site.Microsoft365Group
                'Files viewed or edited' = $site.FilesViewedOrEdited
                'Page views' = $site.PageViews
                'Page visits' = $site.PageVisits
                'Files' = $site.Files
                'Sensitivity' = $site.Sensitivity
                'External sharing' = $site.ExternalSharing
                'OwnersCount' = $null  # Will be filled in by Get-SiteDetails
                'MembersCount' = $null # Will be filled in by Get-SiteDetails
                'ReportDate' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        Write-Log "Built comprehensive site details for $($comprehensiveSiteDetails.Count) sites" -Level Info
        
        # Get detailed information for all sites
        $siteDetails = Get-SiteDetails -SiteSummaries $siteSummaries -TopSites $topSites -ParallelLimit $ParallelLimit
        
        # Generate Excel report
        $success = Export-ComprehensiveExcelReport -ExcelFileName $script:excelFileName -SiteSummaries $siteSummaries -AllSiteSummaries $siteDetails.AllSiteSummaries -AllTopFiles $siteDetails.AllTopFiles -AllTopFolders $siteDetails.AllTopFolders -SiteStorageStats $siteDetails.SiteStorageStats -SitePieCharts $siteDetails.SitePieCharts -ComprehensiveSiteDetails $comprehensiveSiteDetails

        # Automated export of site summary data to CSV for comparison
        $csvExportPath = [System.IO.Path]::ChangeExtension($script:excelFileName, ".csv")
        $siteDetails.AllSiteSummaries | Export-Csv -Path $csvExportPath -NoTypeInformation -Force
        Write-Log "Site summary data exported to CSV: $csvExportPath" -Level Info

        # Output 10 largest sites and their sizes to console (with recycle bin info)
        $topSites = $siteSummaries | Sort-Object TotalStorageGB -Descending | Select-Object -First 10
        Write-Host "`nTop 10 Largest Sites by Total Storage (Including Recycle Bin):" -ForegroundColor Cyan
        $header = "{0,-30} {1,-40} {2,12} {3,12} {4,12}" -f "Site name", "URL", "Storage (GB)", "Recycle (GB)", "Total (GB)"
        Write-Host $header -ForegroundColor White
        Write-Host ("-" * 120) -ForegroundColor DarkGray
        foreach ($site in $topSites) {
            $siteName = $site.SiteName
            $url = $site.WebUrl
            $storage = if ($site.StorageGB) { [math]::Round($site.StorageGB,2) } else { "-" }
            $recycle = if ($site.RecycleBinGB) { [math]::Round($site.RecycleBinGB,2) } else { "0" }
            $total = if ($site.TotalStorageGB) { [math]::Round($site.TotalStorageGB,2) } else { "-" }
            $row = "{0,-30} {1,-40} {2,12} {3,12} {4,12}" -f $siteName, $url, $storage, $recycle, $total
            Write-Host $row -ForegroundColor Gray
        }

        # Match admin screenshot: header wording and column alignment (with enhanced data)
        Write-Host "\nActive sites (Enhanced with Recycle Bin Data)" -ForegroundColor Cyan
        $headerText = "{0,-25} {1,-40} {2,15} {3,15} {4,15}" -f "Site name", "URL", "Storage (GB)", "Recycle (GB)", "Total (GB)"
        Write-Host $headerText -ForegroundColor White
        Write-Host ("-" * 110) -ForegroundColor DarkGray
        foreach ($site in $topSites) {
            $siteName = $site.SiteName
            $url = $site.WebUrl
            $storage = if ($site.StorageGB -is [double]) { [math]::Round($site.StorageGB,2).ToString("F2") } else { $site.StorageGB }
            $recycle = if ($site.RecycleBinGB -is [double]) { [math]::Round($site.RecycleBinGB,2).ToString("F2") } else { "0.00" }
            $total = if ($site.TotalStorageGB -is [double]) { [math]::Round($site.TotalStorageGB,2).ToString("F2") } else { $site.TotalStorageGB }
            $rowText = "{0,-25} {1,-40} {2,15} {3,15} {4,15}" -f $siteName, $url, $storage, $recycle, $total
            Write-Host $rowText -ForegroundColor Gray
        }

        if ($success) {
            Write-Log "Audit completed successfully! Report saved to: $($script:excelFileName)" -Level Success
        }
    }
    catch {
        Write-Log "Script execution failed: $_" -Level Error
        Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level Debug
    }
    finally {
        # Always disconnect from Graph
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            Write-Log "Disconnected from Microsoft Graph" -Level Info
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

function Initialize-Modules {
    $importErrorList = @()
    try { Import-Module Microsoft.Graph.Sites -ErrorAction Stop; Write-Log "Loaded Microsoft.Graph.Sites module" -Level Success } catch { $importErrorList += "Microsoft.Graph.Sites: $_" }
    try { Import-Module Microsoft.Graph.Files -ErrorAction Stop; Write-Log "Loaded Microsoft.Graph.Files module" -Level Success } catch { $importErrorList += "Microsoft.Graph.Files: $_" }
    try { Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue; Write-Log "Loaded Microsoft.Graph.Users module" -Level Success } catch { Write-Log "Microsoft.Graph.Users module not available: $_" -Level Warning; $importErrorList += "Microsoft.Graph.Users: $_" }
    try { Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop; Write-Log "Loaded Microsoft.Graph.Identity.DirectoryManagement module" -Level Success } catch { $importErrorList += "Microsoft.Graph.Identity.DirectoryManagement: $_" }
    try { Import-Module ImportExcel -ErrorAction Stop; Write-Log "Loaded ImportExcel module" -Level Success } catch { $importErrorList += "ImportExcel: $_" }
    if ($importErrorList.Count -gt 0) {
        Write-Log "Some modules failed to load. Attempting to force update..." -Level Warning
        foreach ($moduleError in $importErrorList) {
            $moduleName = $moduleError.Split(':')[0]
            Write-Log "Force updating module: $moduleName" -Level Info
            Install-OrUpdateModule -ModuleName $moduleName -MinimumVersion "2.0.0" -Force
        }
        Write-Log "Modules were force updated. Please restart PowerShell and run the script again." -Level Warning
        return $false
    }
    return $true
}

function Get-SiteUserAccessSummary {
    param(
        [Parameter(Mandatory=$true)]
        $Site
    )
    $owners = @()
    $members = @()
    try {
        $permissions = Get-MgSitePermission -SiteId $Site.Id -All -ErrorAction SilentlyContinue
        foreach ($perm in $permissions) {
            if ($perm.GrantedToV2 -and $perm.GrantedToV2.User) {
                $userType = if ($perm.GrantedToV2.User.UserType -eq 'Guest') { "External Guest" } else { "Internal" }
                $userObj = [PSCustomObject]@{
                    DisplayName = $perm.GrantedToV2.User.DisplayName
                    UserEmail = $perm.GrantedToV2.User.Email
                    UserType = $userType
                    Role = ($perm.Roles -join ', ')
                }
                if ($userObj.Role -match 'Owner|Admin') {
                    $owners += $userObj
                } else {
                    $members += $userObj
                }
            }
        }
    } catch {
        Write-Log "Failed to get user access for site $($Site.DisplayName): $_" -Level Error
    }
    return @{ Owners = $owners; Members = $members }
}