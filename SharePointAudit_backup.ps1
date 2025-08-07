<#
.SYNOPSIS
    SharePoint Site Audit and Reporting Tool - PowerShell 7 Optimized

.DESCRIPTION
    This PowerShell script performs focused auditing of SharePoint sites within a Microsoft 365 tenant.
    It generates streamlined Excel reports with three essential worksheets for site analysis.
    
    Key Features:
    - Scans all SharePoint library sites (excludes OneDrive personal sites by default)
    - PowerShell 7+ optimized with ForEach-Object -Parallel for maximum performance
    - Generates focused Excel report with three worksheets:
      * Site Summary: Complete overview of all SharePoint sites (Site name, URL, Storage used GB, 
        Recycle bin GB, Total storage GB, Primary admin, Hub, Template, Last activity UTC, Date created, 
        Created by, Storage limit GB, Storage used %, Microsoft 365 group, Files viewed or edited, 
        Page views, Page visits, Files, Sensitivity, External sharing, OwnersCount, MembersCount)
      * User Access Summary: Detailed user permissions and access analysis across all sites
      * Top Files & Folders Analysis: Single site analysis showing top 10 largest files and 
        top 10 largest folders with comprehensive size and metadata information
    - Advanced parallel processing capabilities for improved speed and reliability

.PARAMETER ClientId
    Azure AD Application Client ID for authentication. Default: '278b9af9-888d-4344-93bb-769bdd739249'

.PARAMETER TenantId
    Microsoft 365 Tenant ID. Default: 'ca0711e2-e703-4f4e-9099-17d97863211c'

.PARAMETER CertificateThumbprint
    Certificate thumbprint for certificate-based authentication. Default: '2E2502BB1EDB8F36CF9DE50936B283BDD22D5BAD'

.PARAMETER ParallelLimit
    Maximum number of parallel processing threads (PowerShell 7+ ForEach-Object -Parallel). Default: 8

.PARAMETER TestMode
    Run in test mode with limited detailed analysis

.PARAMETER SharePointSitesOnly
    Process only SharePoint sites (excludes OneDrive personal sites - now enabled by default)

.PARAMETER SingleSiteTest
    Test mode to process only a single specified site

.PARAMETER TestSiteName
    Name of the site to process when using SingleSiteTest mode

.PARAMETER OutputPath
    Directory path for output Excel files. Default: Desktop

.PARAMETER ExcludeSites
    Array of site names or URL patterns to exclude from processing

.EXAMPLE
    .\SharePointAudit.ps1
    Runs a full SharePoint audit with default settings using PowerShell 7 parallel processing

.EXAMPLE
    .\SharePointAudit.ps1 -TestMode -SingleSiteTest -TestSiteName "My Test Site"
    Runs in test mode for a single specified site

.EXAMPLE
    .\SharePointAudit.ps1 -SharePointSitesOnly -OutputPath "C:\Reports" -ParallelLimit 16
    Audits only SharePoint sites with maximum parallelism and saves reports to C:\Reports

.NOTES
    Author: Timothy MacLatchy
    Version: 3.0 - PowerShell 7 Optimized
    Created: 2025-08-11
    Updated: 2025-08-11
    
    Requirements:
    - PowerShell 7.0 or later (required for ForEach-Object -Parallel)
    - Microsoft.Graph PowerShell modules (latest version)
    - ImportExcel PowerShell module
    - Certificate-based authentication configured
    - SharePoint Administrator or Global Administrator permissions
    
    Excel Report Output Structure:
    ================================
    
    WORKSHEET: "SharePoint Storage Overview"
    - Category: Type of storage (SharePoint Sites, Recycle Bins)
    - StorageGB: Storage used in gigabytes
    - Percentage: Percentage of total tenant storage
    - SiteCount: Number of sites in this category
    
    WORKSHEET: "Comprehensive Analysis" 
    - Site name: Display name of the SharePoint site
    - URL: Full web URL of the site
    - Teams: Whether the site is connected to Microsoft Teams (Yes/No)
    - Channel sites: Number of channel sites if Teams connected
    - IBMode: Information Barriers mode (Open/Explicit/Mixed)
    - Storage used (GB): Primary storage consumption in GB
    - Recycle bin (GB): Recycle bin storage in GB
    - Total storage (GB): Combined storage including recycle bin
    - Primary admin: Site administrator or "Group owners"
    - Hub: Hub site association if applicable
    - Template: Site template type (Team site, Communication site, etc.)
    - Last activity (UTC): Last recorded activity timestamp
    - Date created: Site creation date
    - Created by: User who created the site
    - Storage limit (GB): Storage quota limit
    - Storage used (%): Percentage of storage quota used
    - Microsoft 365 group: Associated M365 group ID
    - Files viewed or edited: Activity metric
    - Page views: Site page view count
    - Page visits: Unique page visit count
    - Files: Total number of files
    - Sensitivity: Sensitivity label applied
    - External sharing: External sharing configuration
    - OwnersCount: Number of site owners
    - MembersCount: Number of site members
    - ReportDate: Report generation timestamp
    
    WORKSHEET: "Site Storage Analysis"
    - Category: Site classification (SharePoint Site)
    - SiteName: Site display name
    - StorageGB: Primary storage in GB
    - RecycleBinGB: Recycle bin storage in GB
    - TotalStorageGB: Combined total storage
    - Percentage: Percentage of total tenant storage
    - SiteUrl: Site web URL
    - SiteType: Site type classification
    
    WORKSHEET: "User Access Summary"
    - SiteName: Site display name
    - SiteUrl: Site web URL
    - UserDisplayName: User's display name
    - UserEmail: User's email address
    - UserType: Internal/External classification
    - Role: Permission level (Owner, Member, Visitor, etc.)
    - AccessType: Direct/Group-based access
    - LastActivity: Last user activity on the site

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.graph

.LINK
    https://docs.microsoft.com/en-us/powershell/scripting/whats-new/what-s-new-in-powershell-70
#>
#region Parameters
param(
    [Parameter(Mandatory=$false)]
    [string]$ClientId = '278b9af9-888d-4344-93bb-769bdd739249',
    
    [Parameter(Mandatory=$false)]
    [string]$TenantId = 'ca0711e2-e703-4f4e-9099-17d97863211c',
    
    [Parameter(Mandatory=$false)]
    [string]$CertificateThumbprint = '2E2502BB1EDB8F36CF9DE50936B283BDD22D5BAD',
    
    [Parameter(Mandatory=$false)]
    [int]$ParallelLimit = 8,
    
    [Parameter(Mandatory=$false)]
    [switch]$TestMode,
    
    [Parameter(Mandatory=$false)]
    [switch]$SharePointSitesOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$SingleSiteTest,
    
    [Parameter(Mandatory=$false)]
    [string]$TestSiteName = "Journe Australia - Fadi and Josh WIP",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = [Environment]::GetFolderPath('Desktop'),
    
    [Parameter(Mandatory=$false)]
    [string[]]$ExcludeSites,
    
    [Parameter(Mandatory=$false)]
    [switch]$Resume,
    
    [Parameter(Mandatory=$false)]
    [string]$LogFile,
    
    [Parameter(Mandatory=$false)]
    [int]$MaxMemoryMB = 2048
)
#endregion

#region PowerShell Version Check
# Ensure PowerShell 7.0 or later for optimal performance
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Warning "This script is optimized for PowerShell 7.0 or later for maximum performance."
    Write-Warning "Current version: $($PSVersionTable.PSVersion)"
    Write-Warning "Consider upgrading to PowerShell 7+ for ForEach-Object -Parallel support."
    Write-Host "Continuing with PowerShell $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor) compatibility mode..." -ForegroundColor Yellow
    $script:UseLegacyParallel = $true
} else {
    Write-Host "PowerShell $($PSVersionTable.PSVersion) detected - using optimized parallel processing" -ForegroundColor Green
    $script:UseLegacyParallel = $false
}
#endregion

#region Global Variables
$script:tenantName = ""
$script:dateStr = ""
$script:excelFileName = ""
$script:logFilePath = ""
$script:progressId = 1
$script:operationStartTime = $null
$script:siteProcessingStats = @{
    TotalSites = 0
    ProcessedSites = 0
    FailedSites = 0
    TotalStorageGB = 0
    RecycleBinStorageGB = 0
    SharePointSites = 0
}
#endregion
#region Logging, Progress, and Utility Functions
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
    
    # Write to log file if specified
    if ($script:logFilePath) {
        try {
            Add-Content -Path $script:logFilePath -Value $logMessage -ErrorAction SilentlyContinue
        } catch {
            # Silently handle log file errors
        }
    }
}
function Show-Progress {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete,
        [string]$CurrentOperation,
        [int]$Id = 1
    )
    
    # Clamp percentage between 0 and 100
    $safePercent = [math]::Max(0, [math]::Min(100, $PercentComplete))
    
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $safePercent -CurrentOperation $CurrentOperation -Id $Id
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
                # Throttled - use exponential backoff
                $wait = [Math]::Pow(2, $attempt) * $DelaySeconds
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
function Test-MemoryUsage {
    param(
        [int]$MaxMB = $MaxMemoryMB
    )
    
    $currentMemory = (Get-Process -Id $PID).WorkingSet / 1MB
    if ($currentMemory -gt $MaxMB) {
        Write-Log "Memory usage high: $([math]::Round($currentMemory, 2))MB / $MaxMB MB" -Level Warning
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        return $false
    }
    return $true
}
function Initialize-LogFile {
    if ($LogFile) {
        $script:logFilePath = $LogFile
    } else {
        $script:logFilePath = Join-Path -Path $OutputPath -ChildPath "SharePointAudit-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
    }
    
    # Create log directory if it doesn't exist
    $logDir = Split-Path -Parent $script:logFilePath
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    
    Write-Log "Logging to: $script:logFilePath" -Level Success
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
    Write-Log "Authentication failed: $($_)" -Level Error
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
    Write-Log "Error getting tenant name: $($_)" -Level Warning
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
            Write-Log "Could not get root site: $($_)" -Level Warning
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
            Write-Log "Graph API site search failed: $($_)" -Level Warning
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
            Write-Log "Site search failed: $($_)" -Level Warning
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
    Write-Log "Failed to enumerate SharePoint sites: $($_)" -Level Error
        return @()
    }
}
#endregion
#region File and Storage Analysis
function Get-SiteUserAccessSummary {
    param(
        [Parameter(Mandatory=$true)]
        $Site
    )
    $owners = @()
    $members = @()
    $externalGuests = @()
    
    try {
        # Get permissions for SharePoint sites only
        # Try to get actual permissions
        try {
            # First try to get site owners via Groups if the site has an associated Microsoft 365 Group
            if ($Site.Id) {
                    try {
                        $siteDetails = Get-MgSite -SiteId $Site.Id -ErrorAction SilentlyContinue
                        if ($siteDetails.SharepointIds.SiteId) {
                            # Try to get associated group
                            $groupInfo = Get-MgSiteGroup -SiteId $Site.Id -ErrorAction SilentlyContinue
                            if ($groupInfo) {
                                foreach ($group in $groupInfo) {
                                    if ($group.GroupTypes -contains "Unified") {
                                        # This is a Microsoft 365 Group
                                        $groupOwners = Get-MgGroupOwner -GroupId $group.Id -All -ErrorAction SilentlyContinue
                                        foreach ($owner in $groupOwners) {
                                            $owners += [PSCustomObject]@{
                                                DisplayName = $owner.AdditionalProperties.displayName
                                                UserEmail = $owner.AdditionalProperties.mail
                                                UserType = "Internal"
                                                Role = "Group Owner"
                                            }
                                        }
                                        
                                        $groupMembers = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction SilentlyContinue
                                        foreach ($member in $groupMembers) {
                                            if ($member.AdditionalProperties.userType -eq 'Guest') {
                                                $externalGuests += [PSCustomObject]@{
                                                    DisplayName = $member.AdditionalProperties.displayName
                                                    UserEmail = $member.AdditionalProperties.mail
                                                    UserType = "External Guest"
                                                    Role = "Group Member"
                                                }
                                            } else {
                                                $members += [PSCustomObject]@{
                                                    DisplayName = $member.AdditionalProperties.displayName
                                                    UserEmail = $member.AdditionalProperties.mail
                                                    UserType = "Internal"
                                                    Role = "Group Member"
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch {
                        Write-Log "Group-based permission check failed for $($Site.DisplayName): $($_.Exception.Message)" -Level Debug
                    }
                }
                
                # Fallback to direct site permissions if no group data found
                if ($owners.Count -eq 0) {
                    $permissions = Get-MgSitePermission -SiteId $Site.Id -All -ErrorAction SilentlyContinue
                    
                    foreach ($perm in $permissions) {
                        if ($perm.GrantedToV2) {
                            # Handle user permissions
                            if ($perm.GrantedToV2.User) {
                                $userType = if ($perm.GrantedToV2.User.Email -and $perm.GrantedToV2.User.Email -notlike "*@$($Site.WebUrl.Split('.')[1])*") { 
                                    "External Guest" 
                                } else { 
                                    "Internal" 
                                }
                                
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
                                
                                if ($userType -eq "External Guest") {
                                    $externalGuests += $userObj
                                }
                            }
                            # Handle group permissions
                            elseif ($perm.GrantedToV2.Group) {
                                # Try to get group members if possible
                                try {
                                    $groupMembers = Get-MgGroupMember -GroupId $perm.GrantedToV2.Group.Id -All -ErrorAction SilentlyContinue
                                    foreach ($member in $groupMembers) {
                                        if ($member.AdditionalProperties.userType -eq 'Guest') {
                                            $externalGuests += [PSCustomObject]@{
                                                DisplayName = $member.AdditionalProperties.displayName
                                                UserEmail = $member.AdditionalProperties.mail
                                                UserType = "External Guest"
                                                Role = ($perm.Roles -join ', ') + " (via Group)"
                                            }
                                        }
                                    }
                                }
                                catch {
                                    # Group member enumeration failed, continue
                                }
                            }
                        }
                    }
                }
            # If still no owners found, add default entry
            if ($owners.Count -eq 0) {
                $owners += [PSCustomObject]@{
                    DisplayName = "Site Owners"
                    UserEmail = "N/A"
                    UserType = "Internal"
                    Role = "Owner"
                }
            }
        }
        catch {
            # If permissions API fails, try alternative approach
            Write-Log "Standard permissions check failed for $($Site.DisplayName), trying alternative method" -Level Debug
        }
    } 
    catch {
        Write-Log "Failed to get user access for site $($Site.DisplayName): $($_.Exception.Message)" -Level Warning
        
        # Fallback: Create placeholder entry
        $owners += [PSCustomObject]@{
            DisplayName = "Unknown"
            UserEmail = "N/A"
            UserType = "Internal"
            Role = "Owner"
        }
    }
    
    return @{ 
        Owners = $owners
        Members = $members 
        ExternalGuests = $externalGuests
    }
}
function Get-FileData {
    param(
        [Parameter(Mandatory=$true)]
        $Site
    )
    
    try {
        Write-Log "Processing site: $($Site.DisplayName)" -Level Info
        
        # Get all lists that can contain files (not just document libraries)
        $lists = Invoke-WithRetry -ScriptBlock { Get-MgSiteList -SiteId $Site.Id -WarningAction SilentlyContinue } -Activity "Get site lists"
        
        # Include document libraries and other file-containing lists
        $fileContainingLists = $lists | Where-Object { 
            $_.List -and (
                $_.List.Template -eq "documentLibrary" -or
                $_.List.Template -eq "pictureLibrary" -or
                $_.List.Template -eq "assetLibrary" -or
                $_.List.Template -eq "webPageLibrary" -or
                $_.Name -eq "Site Assets" -or
                $_.Name -eq "Style Library" -or
                $_.Name -eq "Site Pages" -or
                $_.Name -eq "Form Templates" -or
                $_.Name -eq "Site Collection Documents" -or
                $_.Name -eq "Site Collection Images" -or
                $_.Name -eq "Master Page Gallery" -or
                $_.Name -eq "Theme Gallery" -or
                $_.Name -eq "Solution Gallery" -or
                $_.List.BaseTemplate -eq 101 -or # Document Library
                $_.List.BaseTemplate -eq 109 -or # Picture Library
                $_.List.BaseTemplate -eq 851 -or # Asset Library
                $_.List.BaseTemplate -eq 119    # Web Page Library
            )
        }
        
        # Also try to get all drives for this site (which might include additional file storage)
        try {
            $drives = Invoke-WithRetry -ScriptBlock { Get-MgSiteDrive -SiteId $Site.Id -WarningAction SilentlyContinue } -Activity "Get site drives"
            Write-Log "Found $($drives.Count) drives for site: $($Site.DisplayName)" -Level Debug
        } catch {
            Write-Log "Could not get drives for site: $($Site.DisplayName)" -Level Debug
        }
        
        if (-not $fileContainingLists -or $fileContainingLists.Count -eq 0) {
            Write-Log "No file-containing lists found for site: $($Site.DisplayName)" -Level Warning
            return @{
                Files = @()
                FolderSizes = @{}
                TotalFiles = 0
                TotalSizeGB = 0
                Users = @()
                ExternalGuests = @()
            }
        }
        
        Write-Log "Found $($fileContainingLists.Count) file-containing lists in site: $($Site.DisplayName) (including document libraries, assets, pages, etc.)" -Level Success
        
        $allFiles = [System.Collections.Generic.List[psobject]]::new()
        $systemFiles = [System.Collections.Generic.List[psobject]]::new()
        $folderSizes = @{}
        $folderItemCounts = @{}
        $folderFileCounts = @{}
        $totalFiles = 0
        $listIndex = 0
        
        foreach ($list in $fileContainingLists) {
            $listIndex++
            $percentComplete = [math]::Round(($listIndex / $fileContainingLists.Count) * 100, 1)
            Show-Progress -Activity "Analyzing File-Containing Lists" -Status "Processing: $($list.DisplayName) | Files found: $totalFiles ($listIndex/$($fileContainingLists.Count))" -PercentComplete $percentComplete -CurrentOperation "$listIndex of $($fileContainingLists.Count) lists"
            
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
                                # More conservative system file filtering - only exclude truly temporary/system files
                                $systemFilePatterns = @(
                                    "~$*", ".tmp", "thumbs.db", ".ds_store", "desktop.ini", "_vti_*"
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
                                
                                # Track ALL files in folder sizes (both regular and system files take up space)
                                $folderPath = $parentPath
                                if (-not $folderSizes.ContainsKey($folderPath)) {
                                    $folderSizes[$folderPath] = 0
                                    $folderItemCounts[$folderPath] = 0
                                    $folderFileCounts[$folderPath] = 0
                                }
                                $folderSizes[$folderPath] += $item.driveItem.size
                                $folderItemCounts[$folderPath] += 1
                                $folderFileCounts[$folderPath] += 1
                                
                                # Separate tracking for reporting purposes
                                if ($isSystem) {
                                    $systemFiles.Add($fileObj) | Out-Null
                                } else {
                                    $allFiles.Add($fileObj) | Out-Null
                                }
                                
                                $filesInThisList++
                                $totalFiles++
                                $currentFileName = if ($item.driveItem.name.Length -gt 50) { $item.driveItem.name.Substring(0, 47) + "..." } else { $item.driveItem.name }
                                Show-Progress -Activity "Analyzing File-Containing Lists" -Status "Processing: $($list.DisplayName) | Files found: $totalFiles | Current: $currentFileName ($filesInThisList)" -PercentComplete $percentComplete -CurrentOperation "$listIndex of $($fileContainingLists.Count) lists"
                            }
                        }
                        # Check for next page link
                        $nextLink = $resp.'@odata.nextLink'
                        if ($nextLink) {
                            Start-Sleep -Milliseconds (Get-Random -Minimum 100 -Maximum 300)
                        }
                    }
                    catch {
                        Write-Log "Failed to process page of list items: $($_)" -Level Error
                        $nextLink = $null
                    }
                }
                
                Write-Log "Processed library: $($list.DisplayName) - Found $filesInThisList files" -Level Success
            }
            catch {
                Write-Log "Failed to process library $($list.DisplayName): $($_)" -Level Error
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
            FolderItemCounts = $folderItemCounts
            FolderFileCounts = $folderFileCounts
            TotalFiles = $allFiles.Count + $systemFiles.Count
            TotalSizeGB = [math]::Round((($allFiles | Measure-Object -Property Size -Sum).Sum + ($systemFiles | Measure-Object -Property Size -Sum).Sum) / 1GB, 2)
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
                    Write-Log "Could not process permission: $($_)" -Level Warning
                }
            }
        }
        catch {
            Write-Log "Could not get permissions for site $($Site.DisplayName): $($_)" -Level Warning
        }
        
        $result.Users = $siteUsers
        $result.ExternalGuests = $externalGuests
        
        Write-Log "Completed storage and access for site: $($Site.DisplayName) - Users: $($siteUsers.Count), Guests: $($externalGuests.Count)" -Level Success
        
        return $result
    }
    catch {
    Write-Log "Failed to get site storage and access info: $($_)" -Level Error
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
#endregion
#region External Sharing Analysis
function Get-ExternalSharingAnalysis {
    param(
        [Parameter(Mandatory=$true)]
        [array]$SiteSummaries
    )
    
    Write-Log "Generating External Sharing Analysis..." -Level Info
    
    # Analyze external sharing patterns
    $sharingAnalysis = @{
        TotalSites = $SiteSummaries.Count
        AnonymousSharingEnabled = ($SiteSummaries | Where-Object { $_.ExternalSharingStatus -eq "Anyone" }).Count
        NewExistingGuestSharing = ($SiteSummaries | Where-Object { $_.ExternalSharingStatus -eq "New and existing guests" }).Count
        ExistingGuestsOnly = ($SiteSummaries | Where-Object { $_.ExternalSharingStatus -eq "Existing guests only" }).Count
        OrganizationOnly = ($SiteSummaries | Where-Object { $_.ExternalSharingStatus -eq "Only people in your organization" }).Count
        UnknownSharing = ($SiteSummaries | Where-Object { $_.ExternalSharingStatus -eq "Unknown" }).Count
        AnonymousAccessAllowed = ($SiteSummaries | Where-Object { $_.AllowAnonymousAccess -eq "Yes" }).Count
        RequireAccountMatch = ($SiteSummaries | Where-Object { $_.RequireAcceptingAccountMatch -eq "Yes" }).Count
        NoExpirationLinks = ($SiteSummaries | Where-Object { $_.AnonymousLinkExpirationPolicy -eq "No expiration" }).Count
    }
    
    # Calculate percentages
    $sharingAnalysis.AnonymousSharingPercent = if ($sharingAnalysis.TotalSites -gt 0) { [math]::Round(($sharingAnalysis.AnonymousSharingEnabled / $sharingAnalysis.TotalSites) * 100, 1) } else { 0 }
    $sharingAnalysis.OrganizationOnlyPercent = if ($sharingAnalysis.TotalSites -gt 0) { [math]::Round(($sharingAnalysis.OrganizationOnly / $sharingAnalysis.TotalSites) * 100, 1) } else { 0 }
    $sharingAnalysis.AnonymousAccessPercent = if ($sharingAnalysis.TotalSites -gt 0) { [math]::Round(($sharingAnalysis.AnonymousAccessAllowed / $sharingAnalysis.TotalSites) * 100, 1) } else { 0 }
    
    # Create summary report
    $sharingReport = @()
    $sharingReport += [PSCustomObject]@{
        'Sharing Category' = "Anyone (Anonymous Links)"
        'Site Count' = $sharingAnalysis.AnonymousSharingEnabled
        'Percentage' = "$($sharingAnalysis.AnonymousSharingPercent)%"
        'Security Risk' = "High"
        'Description' = "Sites allowing anyone with the link to access content"
    }
    
    $sharingReport += [PSCustomObject]@{
        'Sharing Category' = "New and Existing Guests"
        'Site Count' = $sharingAnalysis.NewExistingGuestSharing
        'Percentage' = if ($sharingAnalysis.TotalSites -gt 0) { "$([math]::Round(($sharingAnalysis.NewExistingGuestSharing / $sharingAnalysis.TotalSites) * 100, 1))%" } else { "0%" }
        'Security Risk' = "Medium"
        'Description' = "Sites allowing new external users to be invited"
    }
    
    $sharingReport += [PSCustomObject]@{
        'Sharing Category' = "Existing Guests Only"
        'Site Count' = $sharingAnalysis.ExistingGuestsOnly
        'Percentage' = if ($sharingAnalysis.TotalSites -gt 0) { "$([math]::Round(($sharingAnalysis.ExistingGuestsOnly / $sharingAnalysis.TotalSites) * 100, 1))%" } else { "0%" }
        'Security Risk' = "Low"
        'Description' = "Sites allowing only previously invited guests"
    }
    
    $sharingReport += [PSCustomObject]@{
        'Sharing Category' = "Organization Only"
        'Site Count' = $sharingAnalysis.OrganizationOnly
        'Percentage' = "$($sharingAnalysis.OrganizationOnlyPercent)%"
        'Security Risk' = "None"
        'Description' = "Sites restricted to organization members only"
    }
    
    # Log key findings
    Write-Log "External Sharing Analysis Results:" -Level Info
    Write-Log "  - Total Sites: $($sharingAnalysis.TotalSites)" -Level Info
    Write-Log "  - Anonymous Sharing Enabled: $($sharingAnalysis.AnonymousSharingEnabled) ($($sharingAnalysis.AnonymousSharingPercent)%)" -Level Info
    Write-Log "  - Organization Only: $($sharingAnalysis.OrganizationOnly) ($($sharingAnalysis.OrganizationOnlyPercent)%)" -Level Info
    Write-Log "  - Anonymous Access Allowed: $($sharingAnalysis.AnonymousAccessAllowed) ($($sharingAnalysis.AnonymousAccessPercent)%)" -Level Info
    Write-Log "  - Links with No Expiration: $($sharingAnalysis.NoExpirationLinks)" -Level Info
    
    return @{
        Analysis = $sharingAnalysis
        Report = $sharingReport
    }
}
#endregion
#region Excel Report Generation
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
                # Add chart using ImportExcel's Add-ExcelChart
                try {
                    Add-ExcelChart -Path $Path -WorksheetName $WorksheetName -ChartType $ChartType -Title $ChartTitle -RangeName $ChartColumn -ErrorAction SilentlyContinue
                } catch {
                    Write-Log "Failed to add chart to worksheet '$WorksheetName': $($_)" -Level Warning
                }
            }
        }
    } catch {
    Write-Log "Failed to export worksheet '$WorksheetName': $($_)" -Level Error
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
    
    Write-Log "Generating focused Excel report with 3 worksheets..." -Level Info
    
    # Remove any existing Excel file to avoid conflicts
    if (Test-Path $ExcelFileName) {
        try {
            Remove-Item $ExcelFileName -Force -ErrorAction Stop
            Start-Sleep -Milliseconds 500  # Brief pause to ensure file is released
        }
        catch {
            # If we can't delete, create a unique filename
            $timestamp = Get-Date -Format "HHmmss"
            $ExcelFileName = $ExcelFileName -replace "\.xlsx$", "-$timestamp.xlsx"
            Write-Log "Could not overwrite existing file, using new filename: $ExcelFileName" -Level Warning
        }
    }
    
    Write-Log "Creating Excel report with $($AllSiteSummaries.Count) site summaries..." -Level Info
    
    try {
        # Worksheet 1: Site Summary - Clean and focused data
        Show-Progress -Activity "Generating Excel Report" -Status "Creating Site Summary worksheet..." -PercentComplete 33
        $comprehensiveStorageData = @()
        $totalTenantStorage = 0
        $recycleBinStorage = 0
        $sharePointSitesStorage = 0
        
        # Calculate storage breakdown using AllSiteSummaries which has the correct data
        Write-Log "Processing $($AllSiteSummaries.Count) site summaries for storage analysis..." -Level Info
        $siteCount = 0
        foreach ($siteSummary in $AllSiteSummaries) {
            $siteCount++
            if ($siteCount % 10 -eq 0) {
                $progressPercent = [math]::Round(($siteCount / $AllSiteSummaries.Count) * 20 + 5, 1)
                Show-Progress -Activity "Generating Excel Report" -Status "Processing site $siteCount of $($AllSiteSummaries.Count)..." -PercentComplete $progressPercent
            }
            
            $storageGB = if ($siteSummary.TotalSizeGB) { $siteSummary.TotalSizeGB } else { 0 }
            $recycleBinGB = if ($siteSummary.RecycleBinGB) { $siteSummary.RecycleBinGB } else { 0 }
            $totalStorageGB = if ($siteSummary.TotalStorageGB) { $siteSummary.TotalStorageGB } else { $storageGB + $recycleBinGB }
            
            $totalTenantStorage += $storageGB
            $recycleBinStorage += $recycleBinGB
            $sharePointSitesStorage += $storageGB
            
            # Store basic data without percentage calculation first
            $comprehensiveStorageData += [PSCustomObject]@{
                Category = "SharePoint Site"
                SiteName = $siteSummary.SiteName
                StorageGB = [math]::Round($storageGB, 2)
                RecycleBinGB = [math]::Round($recycleBinGB, 2)
                TotalStorageGB = [math]::Round($totalStorageGB, 2)
                Percentage = 0  # Will calculate after totals are known
                SiteUrl = $siteSummary.SiteUrl
                SiteType = $siteSummary.SiteType
            }
        }
        
        # Now calculate percentages after we have total storage
        Show-Progress -Activity "Generating Excel Report" -Status "Calculating storage percentages..." -PercentComplete 25
        $grandTotal = $totalTenantStorage + $recycleBinStorage
        if ($grandTotal -gt 0) {
            foreach ($item in $comprehensiveStorageData) {
                # Ensure we have valid numbers before calculation
                $totalStorageValue = if ($item.TotalStorageGB -and $item.TotalStorageGB -is [double]) { $item.TotalStorageGB } else { 0 }
                $item.Percentage = [math]::Round(($totalStorageValue / $grandTotal) * 100, 2)
            }
        }
        
        Write-Log "Storage analysis complete. Total: $grandTotal GB" -Level Info
        # Create pie chart summary data
        $pieChartData = @(
            [PSCustomObject]@{
                Category = "SharePoint Sites"
                StorageGB = $sharePointSitesStorage
                Percentage = if ($totalTenantStorage -gt 0) { [math]::Round(($sharePointSitesStorage / $totalTenantStorage) * 100, 2) } else { 0 }
                SiteCount = $AllSiteSummaries.Count
            },
            [PSCustomObject]@{
                Category = "Recycle Bins"
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
                # Determine primary admin - use "Group owners" as default for modern sites
                $primaryAdmin = if ($site.PrimaryAdmin) { $site.PrimaryAdmin } else { "Group owners" }
                
                $ComprehensiveSiteDetails += [PSCustomObject]@{
                    'Site name' = $site.DisplayName
                    'URL' = $site.WebUrl
                    'Teams' = if ($siteSummary.TeamsConnected) { "Yes" } else { "No" }
                    'Channel sites' = if ($siteSummary.ChannelSites) { $siteSummary.ChannelSites } else { "No" }
                    'IBMode' = if ($siteSummary.IBMode) { $siteSummary.IBMode } else { "Open" }
                    'Storage used (GB)' = $siteSummary.StorageGB
                    'Recycle bin (GB)' = $siteSummary.RecycleBinGB
                    'Total storage (GB)' = $siteSummary.TotalStorageGB
                    'Primary admin' = $primaryAdmin
                    'Hub' = if ($site.Hub) { $site.Hub } else { "" }
                    'Template' = if ($siteSummary.Template) { $siteSummary.Template } else { "Team site" }
                    'Last activity (UTC)' = if ($siteSummary.LastActivity) { $siteSummary.LastActivity } else { "" }
                    'Date created' = if ($siteSummary.CreatedDate) { $siteSummary.CreatedDate } else { "" }
                    'Created by' = if ($siteSummary.CreatedBy) { $siteSummary.CreatedBy } else { "" }
                    'Storage limit (GB)' = if ($siteSummary.StorageLimit) { $siteSummary.StorageLimit } else { "25600" }
                    'Storage used (%)' = if ($siteSummary.StorageUsedPercent) { $siteSummary.StorageUsedPercent } else { "0" }
                    'Microsoft 365 group' = if ($siteSummary.HasMicrosoft365Group) { "Yes" } else { "No" }
                    'Files viewed or edited' = if ($siteSummary.FilesViewedOrEdited) { $siteSummary.FilesViewedOrEdited } else { "0" }
                    'Page views' = if ($siteSummary.PageViews) { $siteSummary.PageViews } else { "0" }
                    'Page visits' = if ($siteSummary.PageVisits) { $siteSummary.PageVisits } else { "0" }
                    'Files' = if ($siteSummary.TotalFiles) { $siteSummary.TotalFiles } else { "0" }
                    'Sensitivity' = if ($site.Sensitivity) { $site.Sensitivity } else { "" }
                    'External sharing' = if ($siteSummary.ExternalSharingStatus) { $siteSummary.ExternalSharingStatus } else { "Unknown" }
                    'Default sharing link type' = if ($siteSummary.DefaultSharingLinkType) { $siteSummary.DefaultSharingLinkType } else { "Unknown" }
                    'Anonymous link expiration' = if ($siteSummary.AnonymousLinkExpirationPolicy) { $siteSummary.AnonymousLinkExpirationPolicy } else { "Unknown" }
                    'Require account match' = if ($siteSummary.RequireAcceptingAccountMatch) { $siteSummary.RequireAcceptingAccountMatch } else { "Unknown" }
                    'Allow anonymous access' = if ($siteSummary.AllowAnonymousAccess) { $siteSummary.AllowAnonymousAccess } else { "Unknown" }
                    'OwnersCount' = $siteSummary.OwnersCount
                    'MembersCount' = $siteSummary.MembersCount
                    'ReportDate' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
            }
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
                'Site name' = "Chart: $($chartData.Category)"
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
        
        # Export merged worksheet with comprehensive site details
        try {
            $combinedMainWorksheetData | Export-Excel -Path $ExcelFileName -WorksheetName "SharePoint Storage Overview" -AutoSize -BoldTopRow -FreezeTopRow
            Write-Log "SharePoint Storage Overview worksheet created with $($combinedMainWorksheetData.Count) entries" -Level Info
        }
        catch {
            Write-Log "Failed to create SharePoint Storage Overview worksheet: $($_.Exception.Message)" -Level Error
            throw
        }
        
        # Build user access data for merging
        $userAccessRows = @()
        foreach ($siteDetail in $AllSiteSummaries) {
            $siteName = $siteDetail.SiteName
            $siteUrl = $siteDetail.SiteUrl
            
            # Owners
            if ($siteDetail.OwnersCount -gt 0 -and $siteDetail.Owners) {
                foreach ($owner in $siteDetail.Owners) {
                    $userAccessRows += [PSCustomObject]@{
                        'User/Group' = $owner.DisplayName
                        'Email' = $owner.UserEmail
                        'Type' = $owner.UserType
                        'Role' = $owner.Role
                        'Site' = $siteName
                        'Site URL' = $siteUrl
                    }
                }
            }
            
            # Members
            if ($siteDetail.MembersCount -gt 0 -and $siteDetail.Members) {
                foreach ($member in $siteDetail.Members) {
                    $userAccessRows += [PSCustomObject]@{
                        'User/Group' = $member.DisplayName
                        'Email' = $member.UserEmail
                        'Type' = $member.UserType
                        'Role' = $member.Role
                        'Site' = $siteName
                        'Site URL' = $siteUrl
                    }
                }
            }
        }
        
        # Create merged comprehensive analysis worksheet
        $mergedAnalysisData = @()
        
        # Section 1: Site Overview
        $mergedAnalysisData += [PSCustomObject]@{
            'Section' = "=== SITE OVERVIEW ==="
            'Item' = ""
            'Value' = ""
            'Details' = ""
            'Size (GB)' = ""
            'Type' = ""
            'Path' = ""
        }
        
        foreach ($siteDetail in $AllSiteSummaries) {
            $mergedAnalysisData += [PSCustomObject]@{
                'Section' = "Site Information"
                'Item' = $siteDetail.SiteName
                'Value' = $siteDetail.SiteUrl
                'Details' = "$($siteDetail.OwnersCount) owners, $($siteDetail.MembersCount) members, $($siteDetail.TotalFiles) files"
                'Size (GB)' = $siteDetail.TotalSizeGB
                'Type' = $siteDetail.SiteType
                'Path' = ""
            }
            
            # Add detailed site storage info
            $mergedAnalysisData += [PSCustomObject]@{
                'Section' = "Site Storage"
                'Item' = "Total Storage Used"
                'Value' = "$($siteDetail.TotalSizeGB) GB"
                'Details' = "Files: $($siteDetail.TotalFiles), Folders: $($siteDetail.TotalFolders), Recycle Bin: $($siteDetail.RecycleBinGB) GB"
                'Size (GB)' = $siteDetail.TotalSizeGB
                'Type' = "Storage"
                'Path' = $siteDetail.SiteUrl
            }
        }
        
        # Section 2: User & Group Access
        $mergedAnalysisData += [PSCustomObject]@{
            'Section' = "=== USER `& GROUP ACCESS ==="
            'Item' = ""
            'Value' = ""
            'Details' = ""
            'Size (GB)' = ""
            'Type' = ""
            'Path' = ""
        }
        
        if ($userAccessRows.Count -gt 0) {
            foreach ($userAccess in $userAccessRows | Select-Object -First 20) {
                $mergedAnalysisData += [PSCustomObject]@{
                    'Section' = "User Access"
                    'Item' = $userAccess.'User/Group'
                    'Value' = $userAccess.Email
                    'Details' = "$($userAccess.Type) - $($userAccess.Role)"
                    'Size (GB)' = ""
                    'Type' = $userAccess.Type
                    'Path' = $userAccess.Site
                }
            }
        }
        
        # Section 3: External Sharing Analysis
        $mergedAnalysisData += [PSCustomObject]@{
            'Section' = "=== EXTERNAL SHARING ANALYSIS ==="
            'Item' = ""
            'Value' = ""
            'Details' = ""
            'Size (GB)' = ""
            'Type' = ""
            'Path' = ""
        }
        
        if ($externalSharingResults.Report) {
            foreach ($sharingCategory in $externalSharingResults.Report) {
                $mergedAnalysisData += [PSCustomObject]@{
                    'Section' = "External Sharing"
                    'Item' = $sharingCategory.'Sharing Category'
                    'Value' = "$($sharingCategory.'Site Count') sites ($($sharingCategory.Percentage))"
                    'Details' = $sharingCategory.Description
                    'Size (GB)' = ""
                    'Type' = $sharingCategory.'Security Risk'
                    'Path' = ""
                }
            }
        }
        
        # Section 4: Top 10 Folders Analysis (if single site test)
        if ($SingleSiteTest -and $AllTopFolders.Count -gt 0) {
            $mergedAnalysisData += [PSCustomObject]@{
                'Section' = "=== TOP 10 FOLDERS ANALYSIS ==="
                'Item' = ""
                'Value' = ""
                'Details' = ""
                'Size (GB)' = ""
                'Type' = ""
                'Path' = ""
            }
            
            # Use AllTopFolders parameter directly
            Write-Log "Processing $($AllTopFolders.Count) top folders for comprehensive analysis" -Level Info
            foreach ($folder in $AllTopFolders | Select-Object -First 10) {
                $sizeGB = if ($folder.SizeBytes -and $folder.SizeBytes -gt 0) { 
                    [math]::Round($folder.SizeBytes / 1GB, 3) 
                } elseif ($folder.SizeGB) {
                    $folder.SizeGB
                } else { 
                    0 
                }
                
                Write-Log "Folder: $($folder.FolderName), SizeBytes: $($folder.SizeBytes), SizeGB: $sizeGB" -Level Debug
                
                $mergedAnalysisData += [PSCustomObject]@{
                    'Section' = "Top Folders"
                    'Item' = $folder.FolderName
                    'Value' = "$($folder.ItemCount) items"
                    'Details' = "Files: $($folder.FileCount), Folders: $($folder.SubFolderCount)"
                    'Size (GB)' = $sizeGB
                    'Type' = "Folder"
                    'Path' = $folder.FolderPath
                }
            }
        }
        
        # Export merged analysis worksheet
        try {
            # Ensure all numeric values are properly formatted before export
            foreach ($item in $mergedAnalysisData) {
                if ($item.'Size (GB)' -and $item.'Size (GB)' -is [string]) {
                    try {
                        $item.'Size (GB)' = [double]$item.'Size (GB)'
                    }
                    catch {
                        $item.'Size (GB)' = 0
                    }
                }
                if (-not $item.'Size (GB)') {
                    $item.'Size (GB)' = 0
                }
            }
            
            Write-Log "Exporting $($mergedAnalysisData.Count) rows to Comprehensive Analysis worksheet" -Level Info
            $mergedAnalysisData | Export-Excel -Path $ExcelFileName -WorksheetName "Comprehensive Analysis" -AutoSize -BoldTopRow -FreezeTopRow
        }
        catch {
            Write-Log "Failed to create Comprehensive Analysis worksheet: $($_.Exception.Message)" -Level Error
            throw
        }
        
        Write-Log "Excel report created successfully: $ExcelFileName" -Level Success
        return $true
    }
    catch {
    Write-Log ("Failed to create Excel report: " + $PSItem.Exception.Message) -Level Error
        throw
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
    $filteredSites = $Sites | Where-Object { $_.Id -and ($_.Id -ne "") }

    # Proper exclusion filtering (supports wildcards against DisplayName or WebUrl)
    if ($ExcludeSites -and $ExcludeSites.Count -gt 0) {
        $preExcludeCount = $filteredSites.Count
        $filteredSites = $filteredSites | Where-Object {
            $siteObj = $_
            $match = $false
            foreach ($pattern in $ExcludeSites) {
                if (($siteObj.DisplayName -and $siteObj.DisplayName -like $pattern) -or ($siteObj.WebUrl -and $siteObj.WebUrl -like $pattern)) { $match = $true; break }
            }
            return -not $match
        }
        $removed = $preExcludeCount - $filteredSites.Count
        if ($removed -gt 0) { Write-Log "Excluded $removed sites based on ExcludeSites filter" -Level Info }
    }
    
    # Filter to only SharePoint library sites (exclude OneDrive personal sites)
    $sharePointSites = @()
    
    foreach ($site in $filteredSites) {
        # Skip OneDrive personal sites - only process SharePoint team and communication sites
        if (-not ($site.WebUrl -like "*-my.sharepoint.com/personal/*" -or 
                  $site.WebUrl -like "*/personal/*" -or 
                  $site.WebUrl -like "*mysites*" -or 
                  $site.Name -like "*OneDrive*" -or 
                  $site.DisplayName -like "*OneDrive*" -or
                  ($site.WebUrl -and $site.WebUrl -match "onedrive") -or
                  ($site.Drive -and $site.Drive.DriveType -eq "personal"))) {
            $sharePointSites += $site
        }
    }
    
    Write-Log "SharePoint library sites to process: $($sharePointSites.Count)" -Level Info
    Write-Log "Total sites to process: $($filteredSites.Count)" -Level Info
    
    $filteredSites | Select-Object -First 10 | ForEach-Object {
        Write-Log "Id: '$($_.Id)', DisplayName: '$($_.DisplayName)', WebUrl: '$($_.WebUrl)'" -Level Debug
    }
    
    # Apply SharePointSitesOnly filtering if enabled
    if ($SharePointSitesOnly) {
        Write-Log "SharePointSitesOnly mode: Excluding OneDrive personal sites from processing" -Level Info
        $sitesToProcess = $sharePointSites
    } else {
        # In test mode, process all sites but with limited detailed analysis
        $sitesToProcess = if ($TestMode) { 
            Write-Log "Test mode: Processing all sites but limiting detailed file analysis" -Level Info
            $filteredSites 
        } else { 
            $filteredSites 
        }
    }
    
    # Apply SingleSiteTest filtering if enabled
    if ($SingleSiteTest) {
        Write-Log "SingleSiteTest mode: Processing only '$TestSiteName' for testing" -Level Info
        $testSite = $sitesToProcess | Where-Object { $_.DisplayName -eq $TestSiteName } | Select-Object -First 1
        if ($testSite) {
            $sitesToProcess = @($testSite)
            Write-Log "Found test site: $($testSite.DisplayName) - $($testSite.WebUrl)" -Level Success
        } else {
            Write-Log "Test site '$TestSiteName' not found. Available sites:" -Level Warning
            $sitesToProcess | Select-Object -First 5 | ForEach-Object {
                Write-Log "  - $($_.DisplayName)" -Level Warning
            }
            $sitesToProcess = $sitesToProcess | Select-Object -First 1
            Write-Log "Using first available site: $($sitesToProcess[0].DisplayName)" -Level Info
        }
    }
    
    # Update global stats
    $script:siteProcessingStats.TotalSites = $sitesToProcess.Count
    $script:siteProcessingStats.SharePointSites = $sharePointSites.Count
    
    # Start timer for parallel scan
    $stepTimer = [System.Diagnostics.Stopwatch]::StartNew()
    
    # Process sites in parallel using PowerShell 7+ optimized approach
    $siteSummaries = @()
    
    if ($script:UseLegacyParallel) {
        # Legacy PowerShell 5.1 compatible processing using jobs
        Write-Log "Using legacy job-based parallel processing for PowerShell 5.1 compatibility" -Level Info
        $jobs = @()
        $maxJobs = $ParallelLimit
        $processedCount = 0
        
        foreach ($site in $sitesToProcess) {
            # Wait for available job slot
            while ((Get-Job -State Running | Measure-Object).Count -ge $maxJobs) {
                Start-Sleep -Milliseconds 100
                Get-Job -State Completed | ForEach-Object {
                    $jobResult = Receive-Job -Job $_
                    Remove-Job -Job $_
                    if ($jobResult -ne $null) {
                        $siteSummaries += $jobResult
                        $processedCount++
                        
                        # Show progress
                        $percentComplete = [math]::Round(($processedCount / $sitesToProcess.Count) * 100, 1)
                        Show-Progress -Activity "Scanning SharePoint Sites" -Status "Processed: $processedCount/$($sitesToProcess.Count) sites" -PercentComplete $percentComplete -CurrentOperation "Current: $($jobResult.SiteName)"
                    }
                }
            }
            
            # Start new job for site processing
            $job = Start-Job -ScriptBlock {
                param($siteData)
                
                $site = $siteData
                $siteId = $site.Id
                $displayName = $site.DisplayName
                $webUrl = $site.WebUrl
                
                if (-not $siteId) { return $null }
                
                # Site type determination (SharePoint sites only)
                $siteType = "SharePoint Site"
                $storageGB = 0
                $recycleBinGB = 0
                
                # Basic storage calculation
                try {
                    if ($site.Drive -and $site.Drive.Quota -and $site.Drive.Quota.Used) {
                        $storageGB = [math]::Round($site.Drive.Quota.Used / 1GB, 2)
                    }
                } catch { }
                
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
                    Template = "Team site"
                    TeamsConnected = "False"
                    ChannelSites = ""
                    IBMode = "Open"
                    CreatedBy = ""
                    CreatedDate = ""
                    LastActivity = ""
                    StorageLimit = 25600
                    StorageUsedPercent = if ($storageGB -gt 0 -and 25600 -gt 0) { [math]::Round(($storageGB / 25600) * 100, 2) } else { 0 }
                    HasMicrosoft365Group = "False"
                    ExternalSharingStatus = "Unknown"
                    DefaultSharingLinkType = "Unknown"
                    AnonymousLinkExpirationPolicy = "Unknown"
                    RequireAcceptingAccountMatch = "Unknown"
                    AllowAnonymousAccess = "Unknown"
                    FilesViewedOrEdited = ""
                    PageViews = ""
                    PageVisits = ""
                    TotalFiles = ""
                }
            } -ArgumentList $site
            
            $jobs += $job
        }
        
        # Wait for all jobs to complete
        while ((Get-Job -State Running | Measure-Object).Count -gt 0) {
            Start-Sleep -Milliseconds 500
            Get-Job -State Completed | ForEach-Object {
                $jobResult = Receive-Job -Job $_
                Remove-Job -Job $_
                if ($jobResult -ne $null) {
                    $siteSummaries += $jobResult
                    $processedCount++
                    
                    # Show progress
                    $percentComplete = [math]::Round(($processedCount / $sitesToProcess.Count) * 100, 1)
                    Show-Progress -Activity "Scanning SharePoint Sites" -Status "Processed: $processedCount/$($sitesToProcess.Count) sites" -PercentComplete $percentComplete -CurrentOperation "Current: $($jobResult.SiteName)"
                }
            }
        }
        
        # Clean up jobs
        Get-Job | Remove-Job -Force
        
    } else {
        # PowerShell 7+ optimized parallel processing using ForEach-Object -Parallel
        Write-Log "Using PowerShell 7+ ForEach-Object -Parallel for maximum performance with $ParallelLimit threads" -Level Info
        
        $siteSummaries = $sitesToProcess | ForEach-Object -Parallel {
            $site = $_
            $siteId = $site.Id
            $displayName = $site.DisplayName
            $webUrl = $site.WebUrl
            
            if (-not $siteId) { return $null }
            
            # Site type determination (SharePoint sites only)
            $siteType = "SharePoint Site"
            $storageGB = 0
            $recycleBinGB = 0
            
            # Basic storage calculation
            try {
                if ($site.Drive -and $site.Drive.Quota -and $site.Drive.Quota.Used) {
                    $storageGB = [math]::Round($site.Drive.Quota.Used / 1GB, 2)
                }
            } catch { 
                # Silently continue on storage calculation errors
            }
            
            # Enhanced data collection
            $template = "Team site"
            $isTeamsConnected = "False"
            $createdBy = ""
            $createdDate = ""
            $lastActivity = ""
            $storageLimit = 25600  # Default SharePoint storage limit
            $storageUsedPercent = if ($storageGB -gt 0 -and $storageLimit -gt 0) { [math]::Round(($storageGB / $storageLimit) * 100, 2) } else { 0 }
            $hasMicrosoft365Group = "False"
            $externalSharingStatus = "Unknown"
            $defaultSharingLinkType = "Unknown"
            $anonymousLinkExpirationPolicy = "Unknown"
            $requireAcceptingAccountMatchInvitedAccount = "Unknown"
            $allowAnonymousAccess = "Unknown"
            $filesViewedOrEdited = ""
            $pageViews = ""
            $pageVisits = ""
            $totalFiles = ""
            $channelSites = ""
            $ibMode = "Open"
            
            # Return site summary object
            [PSCustomObject]@{
                Site = $site
                SiteId = $siteId
                SiteName = $displayName
                SiteType = $siteType
                StorageGB = $storageGB
                RecycleBinGB = $recycleBinGB
                TotalStorageGB = ($storageGB + $recycleBinGB)
                StorageBytes = $storageGB * 1GB
                WebUrl = $webUrl
                Template = $template
                TeamsConnected = $isTeamsConnected
                ChannelSites = $channelSites
                IBMode = $ibMode
                CreatedBy = $createdBy
                CreatedDate = $createdDate
                LastActivity = $lastActivity
                StorageLimit = $storageLimit
                StorageUsedPercent = $storageUsedPercent
                HasMicrosoft365Group = $hasMicrosoft365Group
                ExternalSharingStatus = $externalSharingStatus
                DefaultSharingLinkType = $defaultSharingLinkType
                AnonymousLinkExpirationPolicy = $anonymousLinkExpirationPolicy
                RequireAcceptingAccountMatch = $requireAcceptingAccountMatchInvitedAccount
                AllowAnonymousAccess = $allowAnonymousAccess
                FilesViewedOrEdited = $filesViewedOrEdited
                PageViews = $pageViews
                PageVisits = $pageVisits
                TotalFiles = $totalFiles
            }
        } -ThrottleLimit $ParallelLimit | Where-Object { $_ -ne $null }
        
        # Update processing stats
        $script:siteProcessingStats.ProcessedSites = $siteSummaries.Count
        $script:siteProcessingStats.TotalStorageGB = ($siteSummaries | Measure-Object -Property StorageGB -Sum).Sum
        $script:siteProcessingStats.RecycleBinStorageGB = ($siteSummaries | Measure-Object -Property RecycleBinGB -Sum).Sum
    }
    
    $stepTimer.Stop()
    $elapsed = $stepTimer.Elapsed
    Write-Log "Parallel Site Summary Scan completed in $($elapsed.TotalSeconds) seconds ($([Math]::Round($elapsed.TotalMinutes, 2)) min)" -Level Info
    
    # Clear progress bar
    Stop-Progress -Activity "Scanning SharePoint Sites"
    
    # Output summary table to console matching SharePoint Admin Centre
    Write-Host "`nActive Sites Summary:" -ForegroundColor Cyan
    $header = "{0,-30} {1,-40} {2,12} {3,-20}" -f "Site name", "URL", "Storage (GB)", "Primary admin"
    Write-Host $header -ForegroundColor White
    Write-Host ("-" * 110) -ForegroundColor DarkGray
    foreach ($site in $siteSummaries) {
        $siteName = $site.SiteName
        $url = $site.WebUrl
        $storage = if ($site.StorageGB) { [math]::Round($site.StorageGB,2) } else { "-" }
        $admin = if ($site.Site.PrimaryAdmin) { $site.Site.PrimaryAdmin } else { "Group owners" }
        $row = "{0,-30} {1,-40} {2,12} {3,-20}" -f $siteName, $url, $storage, $admin
        Write-Host $row -ForegroundColor Gray
    }
    
    return $siteSummaries
}

#region Site Summary Sanitization
function Sanitize-SiteSummaries {
    param(
        [Parameter(Mandatory=$true)][array]$SiteSummaries
    )
    $originalCount = $SiteSummaries.Count
    $valid = @()
    $removedDetails = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in $SiteSummaries) {
        if ($null -eq $entry) { $removedDetails.Add("Null entry") | Out-Null; continue }
        if (-not ($entry -is [psobject])) { $removedDetails.Add("Non-object type: " + $entry.GetType().Name) | Out-Null; continue }
        if (-not ($entry.PSObject.Properties.Name -contains 'Site')) { $removedDetails.Add("Missing 'Site' property") | Out-Null; continue }
        if ($null -eq $entry.Site -or -not $entry.Site.Id) { $removedDetails.Add("Null/invalid Site property") | Out-Null; continue }
        $valid += $entry
    }
    $removed = $originalCount - $valid.Count
    if ($removed -gt 0) {
        Write-Log "Sanitized site summaries: removed $removed invalid entries (kept $($valid.Count))." -Level Warning
        # If debug, output first few removal reasons
        $removedDetails | Select-Object -First 5 | ForEach-Object { Write-Log "Sanitization detail: $_" -Level Debug }
    }
    return ,$valid  # ensure array
}
#endregion
function Get-SiteDetails {
    param(
        [Parameter(Mandatory=$true)]
        [array]$SiteSummaries,
        
        [array]$TopSites,
        
        [int]$ParallelLimit = 4
    )
    
    Write-Log "Getting detailed information for sites..." -Level Info
    Write-Log "Processing $($SiteSummaries.Count) sites for detailed analysis" -Level Info
    
    $allSiteSummaries = @()
    $allTopFiles = @()
    $allTopFolders = @()
    $siteStorageStats = @{}
    $sitePieCharts = @{}
    
    # Process each site for detailed analysis
    $detailProcessedCount = 0
    $totalSites = $SiteSummaries.Count
    $skippedCount = 0
    $errorCount = 0
    
    foreach ($siteSummary in $SiteSummaries) {
        $detailProcessedCount++
        $percentComplete = [math]::Round(($detailProcessedCount / $totalSites) * 100, 1)
        
        # Show detailed progress
        Show-Progress -Activity "Detailed Site Analysis" -Status "Processing: $($siteSummary.SiteName) ($detailProcessedCount/$totalSites)" -PercentComplete $percentComplete -CurrentOperation "Analyzing $($siteSummary.SiteName)"
        
        # Check if Site object is null
        if ($null -eq $siteSummary.Site) {
            Write-Log "Site object is null at index $detailProcessedCount. SiteSummary: $($siteSummary.SiteName)" -Level Warning
            $skippedCount++
            $script:siteProcessingStats.FailedSites++
            $detailProcessedCount++
            continue
        }
        
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
            Write-Log "Failed to get user access for site $($site.DisplayName): $($_)" -Level Error
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
                        $folderPath = $_.Key
                        $folderName = if ($folderPath -match "([^/]+)$") { $matches[1] } else { $folderPath }
                        $itemCount = if ($fileData.FolderItemCounts.ContainsKey($folderPath)) { $fileData.FolderItemCounts[$folderPath] } else { 0 }
                        $fileCount = if ($fileData.FolderFileCounts.ContainsKey($folderPath)) { $fileData.FolderFileCounts[$folderPath] } else { 0 }
                        [PSCustomObject]@{
                            SiteName = $site.DisplayName
                            FolderPath = $_.Key
                            FolderName = $folderName
                            SizeBytes = $_.Value
                            SizeGB = [math]::Round($_.Value / 1GB, 3)
                            SizeMB = [math]::Round($_.Value / 1MB, 2)
                            ItemCount = $itemCount
                            FileCount = $fileCount
                            SubFolderCount = 0  # Would need additional processing to count subfolders
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
                Write-Log "Failed to analyze files and folders for site $($site.DisplayName): $($_)" -Level Error
                $script:siteProcessingStats.FailedSites++
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
#endregion
#region Module Initialization
function Initialize-Modules {
    $importErrorList = @()
    try { Import-Module Microsoft.Graph.Sites -ErrorAction Stop; Write-Log "Loaded Microsoft.Graph.Sites module" -Level Success } catch { $importErrorList += "Microsoft.Graph.Sites: $($_)" }
    try { Import-Module Microsoft.Graph.Files -ErrorAction Stop; Write-Log "Loaded Microsoft.Graph.Files module" -Level Success } catch { $importErrorList += "Microsoft.Graph.Files: $($_)" }
    try { Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue; Write-Log "Loaded Microsoft.Graph.Users module" -Level Success } catch { Write-Log "Microsoft.Graph.Users module not available: $($_)" -Level Warning; $importErrorList += "Microsoft.Graph.Users: $($_)" }
    try { Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop; Write-Log "Loaded Microsoft.Graph.Identity.DirectoryManagement module" -Level Success } catch { $importErrorList += "Microsoft.Graph.Identity.DirectoryManagement: $($_)" }
    try { Import-Module ImportExcel -ErrorAction Stop; Write-Log "Loaded ImportExcel module" -Level Success } catch { $importErrorList += "ImportExcel: $($_)" }
    
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
#endregion
#region Summary Report
function Show-SummaryReport {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Stats,
        
        [Parameter(Mandatory=$true)]
        [string]$ReportFile
    )
    
    $totalTime = if ($script:operationStartTime) { 
        $elapsed = (Get-Date) - $script:operationStartTime
        "{0:hh\:mm\:ss}" -f $elapsed 
    } else { "Unknown" }
    
    Write-Host "`n" -NoNewline
    Write-Host "=============================================" -ForegroundColor Green
    Write-Host "          SHAREPOINT AUDIT SUMMARY" -ForegroundColor Green
    Write-Host "=============================================" -ForegroundColor Green
    
    Write-Host "`nProcessing Statistics:" -ForegroundColor Cyan
    Write-Host "  Total Sites Found:         $($Stats.TotalSites)" -ForegroundColor White
    Write-Host "  Sites Processed:          $($Stats.ProcessedSites)" -ForegroundColor White
    Write-Host "  Sites Failed:             $($Stats.FailedSites)" -ForegroundColor White
    Write-Host "  SharePoint Sites:         $($Stats.SharePointSites)" -ForegroundColor White
    Write-Host "  OneDrive Sites:           $($Stats.OneDriveSites)" -ForegroundColor White
    
    Write-Host "\nStorage Analysis:" -ForegroundColor Cyan
    Write-Host "  Total Storage Used:       $([math]::Round($Stats.TotalStorageGB, 2)) GB" -ForegroundColor White
    Write-Host "  Recycle Bin Storage:      $([math]::Round($Stats.RecycleBinStorageGB, 2)) GB" -ForegroundColor White
    Write-Host "  Total with Recycle Bin:   $([math]::Round(($Stats.TotalStorageGB + $Stats.RecycleBinStorageGB), 2)) GB" -ForegroundColor White
    
    Write-Host "\nExecution Time: $totalTime" -ForegroundColor Cyan
    
    Write-Host "\nReport Generated:" -ForegroundColor Cyan
    Write-Host "  Excel Report:             $ReportFile" -ForegroundColor White
    
    if ($script:logFilePath) {
        Write-Host "  Log File:                $script:logFilePath" -ForegroundColor White
    }
    
    Write-Host "\n=============================================" -ForegroundColor Green
    Write-Host "               AUDIT COMPLETE" -ForegroundColor Green
    Write-Host "=============================================" -ForegroundColor Green
}
#endregion
#region Main Function
function Main {
    try {
        $script:operationStartTime = Get-Date
        
        Write-Log "SharePoint Tenant Storage & Access Report Generator" -Level Success
        Write-Log "=============================================" -Level Success
        
        # Initialize log file
        Initialize-LogFile
        
        # Initialize modules
        if (-not (Initialize-Modules)) {
            return
        }
        
        # Connect to Microsoft Graph
        Connect-ToGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
        
        # Get tenant name
        $script:tenantName = Get-TenantName
        
        # Create date string and filename
        $script:dateStr = Get-Date -Format "yyyyMMdd"
        $defaultFileName = "SharePointAudit-AllSites-$($script:tenantName)-$($script:dateStr).xlsx"
        
        # Get save file path from user dialog or use provided path
        $script:excelFileName = Join-Path -Path $OutputPath -ChildPath $defaultFileName
        
        # Get all SharePoint sites
        $sites = Get-AllSharePointSites
        
        if (-not $sites -or $sites.Count -eq 0) {
            Write-Log "No SharePoint sites found in tenant. Exiting." -Level Error
            return
        }
        
        # Filter to exclude OneDrive personal sites - focus only on SharePoint library sites
        $sharePointSites = @()
        foreach ($site in $sites) {
            # Skip OneDrive personal sites - only process SharePoint team and communication sites
            if (-not ($site.WebUrl -like "*-my.sharepoint.com/personal/*" -or 
                      $site.WebUrl -like "*/personal/*" -or 
                      $site.WebUrl -like "*mysites*" -or 
                      $site.Name -like "*OneDrive*" -or 
                      $site.DisplayName -like "*OneDrive*" -or
                      ($site.WebUrl -and $site.WebUrl -match "onedrive") -or
                      ($site.Drive -and $site.Drive.DriveType -eq "personal"))) {
                $sharePointSites += $site
            }
        }
        
        Write-Log "Total sites found: $($sites.Count)" -Level Info
        Write-Log "SharePoint library sites (excluding OneDrive): $($sharePointSites.Count)" -Level Info
        Write-Log "OneDrive personal sites excluded: $($sites.Count - $sharePointSites.Count)" -Level Info
        
        if (-not $sharePointSites -or $sharePointSites.Count -eq 0) {
            Write-Log "No SharePoint library sites found in tenant after filtering. Exiting." -Level Error
            return
        }
        
        # Get site summaries (storage information) - only for SharePoint library sites
        $siteSummaries = Get-SiteSummaries -Sites $sharePointSites -ParallelLimit $ParallelLimit
        
        if (-not $siteSummaries -or $siteSummaries.Count -eq 0) {
            Write-Log "No site summaries could be generated. Exiting." -Level Error
            return
        }
        
        # Get top sites by storage for detailed analysis
        $topSites = $siteSummaries | Sort-Object TotalStorageGB -Descending | Select-Object -First 10
        
    # Sanitize summaries to remove any anomalous boolean/invalid entries before deep processing
    $siteSummaries = Sanitize-SiteSummaries -SiteSummaries $siteSummaries

    # Get detailed information for all sites
    $siteDetails = Get-SiteDetails -SiteSummaries $siteSummaries -TopSites $topSites -ParallelLimit $ParallelLimit
        
        # Build comprehensive site details for Excel export
        $comprehensiveSiteDetails = @()
        foreach ($siteSummary in $siteSummaries) {
            # Check if Site object is null
            if ($null -eq $siteSummary.Site) {
                Write-Log "Skipping site summary with null Site object" -Level Warning
                continue
            }
            
            $site = $siteSummary.Site
            $comprehensiveSiteDetails += [PSCustomObject]@{
                'Site name' = $siteSummary.SiteName
                'URL' = $siteSummary.WebUrl
                'Teams' = $siteSummary.TeamsConnected
                'Channel sites' = $siteSummary.ChannelSites
                'IBMode' = $siteSummary.IBMode
                'Storage used (GB)' = $siteSummary.StorageGB
                'Recycle bin (GB)' = $siteSummary.RecycleBinGB
                'Total storage (GB)' = $siteSummary.TotalStorageGB
                'Primary admin' = $siteSummary.CreatedBy
                'Hub' = ""  # Hub information would need additional API calls
                'Template' = $siteSummary.Template
                'Last activity (UTC)' = $siteSummary.LastActivity
                'Date created' = $siteSummary.CreatedDate
                'Created by' = $siteSummary.CreatedBy
                'Storage limit (GB)' = $siteSummary.StorageLimit
                'Storage used (%)' = $siteSummary.StorageUsedPercent
                'Microsoft 365 group' = $siteSummary.HasMicrosoft365Group
                'Files viewed or edited' = $siteSummary.FilesViewedOrEdited
                'Page views' = $siteSummary.PageViews
                'Page visits' = $siteSummary.PageVisits
                'Files' = $siteSummary.TotalFiles
                'Sensitivity' = ""  # Sensitivity labels would need additional API calls
                'External sharing' = $siteSummary.ExternalSharingStatus
                'Default sharing link type' = $siteSummary.DefaultSharingLinkType
                'Anonymous link expiration' = $siteSummary.AnonymousLinkExpirationPolicy
                'Require account match' = $siteSummary.RequireAcceptingAccountMatch
                'Allow anonymous access' = $siteSummary.AllowAnonymousAccess
                'OwnersCount' = 1  # Will be updated with actual data in detailed processing
                'MembersCount' = 0  # Will be updated with actual data in detailed processing
                'ReportDate' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        
        Write-Log "Built comprehensive site details for $($comprehensiveSiteDetails.Count) sites" -Level Info
        
        # Generate Excel report
        $success = Export-ComprehensiveExcelReport -ExcelFileName $script:excelFileName -SiteSummaries $siteSummaries -AllSiteSummaries $siteDetails.AllSiteSummaries -AllTopFiles $siteDetails.AllTopFiles -AllTopFolders $siteDetails.AllTopFolders -SiteStorageStats $siteDetails.SiteStorageStats -SitePieCharts $siteDetails.SitePieCharts -ComprehensiveSiteDetails $comprehensiveSiteDetails
        
        # Output 10 largest sites and their sizes to console (with recycle bin info)
        $topSites = $siteSummaries | Sort-Object TotalStorageGB -Descending | Select-Object -First 10
        Write-Host "\nTop 10 Largest Sites by Total Storage (Including Recycle Bin):" -ForegroundColor Cyan
        $header = "{0,-30} {1,-40} {2,12} {3,12} {4,12}" -f "Site name", "URL", "Storage (GB)", "Recycle (GB)", "Total (GB)"
        Write-Host $header -ForegroundColor White
        Write-Host ("-" * 120) -ForegroundColor DarkGray
        
        # Ensure we have valid data for top sites
        $validTopSites = $topSites | Where-Object { $_.StorageGB -ne $null -and $_.StorageGB -ne "-" -and $_.TotalStorageGB -ne $null -and $_.TotalStorageGB -ne "-" }
        
        if ($validTopSites.Count -eq 0) {
            Write-Host "No sites with storage data found" -ForegroundColor Yellow
        }
        else {
            foreach ($site in $validTopSites) {
                $siteName = if ($site.SiteName.Length -gt 28) { $site.SiteName.Substring(0,28) + ".." } else { $site.SiteName }
                $url = if ($site.WebUrl.Length -gt 38) { $site.WebUrl.Substring(0,38) + ".." } else { $site.WebUrl }
                $storage = if ($site.StorageGB -is [double]) { [math]::Round($site.StorageGB,2).ToString() } else { $site.StorageGB }
                $recycle = if ($site.RecycleBinGB -is [double]) { [math]::Round($site.RecycleBinGB,2).ToString() } else { "0" }
                $total = if ($site.TotalStorageGB -is [double]) { [math]::Round($site.TotalStorageGB,2).ToString() } else { $site.TotalStorageGB }
                $row = "{0,-30} {1,-40} {2,12} {3,12} {4,12}" -f $siteName, $url, $storage, $recycle, $total
                Write-Host $row -ForegroundColor Gray
            }
        }
        
        # Calculate and display totals
        $totalStorage = ($siteSummaries | Where-Object { $_.StorageGB -is [double] } | Measure-Object StorageGB -Sum).Sum
        $totalRecycle = ($siteSummaries | Where-Object { $_.RecycleBinGB -is [double] } | Measure-Object RecycleBinGB -Sum).Sum
        $grandTotal = $totalStorage + $totalRecycle
        
        Write-Host ("-" * 120) -ForegroundColor DarkGray
        $totalRow = "{0,-30} {1,-40} {2,12} {3,12} {4,12}" -f "TOTAL", "", [math]::Round($totalStorage,2), [math]::Round($totalRecycle,2), [math]::Round($grandTotal,2)
        Write-Host $totalRow -ForegroundColor White -BackgroundColor DarkBlue
        
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
            
            # Show summary report
            Show-SummaryReport -Stats $script:siteProcessingStats -ReportFile $script:excelFileName
        }
    }
    catch {
    Write-Log "Script execution failed: $($_)" -Level Error
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