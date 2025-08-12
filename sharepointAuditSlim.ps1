<#
.DESCRIPTION
    Optimized SharePoint audit script using Microsoft Graph API for maximum performance.
    Generates comprehensive Excel reports with minimal API calls and memory footprint.
    Includes a resume capability to continue after an interruption and uses delta queries for efficiency.
    Generates 3 worksheets:
    1. Site Summary: Overview of all SharePoint sites with columns: Site name, Storage used GB, Recycle bin GB, Total storage GB, Storage limit GB, Storage used % of tenant storage, File Count, any linked Microsoft 365 groups, OwnersCount, MembersCount, URL
    2. User Access Summary: All users and all sites they can access
    3. Top Files & Folders Analysis: For the largest site: 10 largest files, 10 largest folders, all files over 1gb, all folders over 1gb, with file count and storage used % of tenant storage
    REQUIRED PERMISSIONS (App Registration):
    - Sites.Read.All
    - Sites.ReadWrite.All
    - Group.Read.All
    - Directory.Read.All
    - Reports.Read.All
    - User.Read.All
    - Files.Read.All
    - Files.ReadWrite.All
    (Grant admin consent for all above permissions)
#>
param(
    [string]$ClientId = '278b9af9-888d-4344-93bb-769bdd739249',
    [string]$TenantId = 'ca0711e2-e703-4f4e-9099-17d97863211c',
    [string]$CertificateThumbprint = '2E2502BB1EDB8F36CF9DE50936B283BDD22D5BAD',
    [int]$ParallelLimit = 8,
    [string]$OutputPath = [Environment]::GetFolderPath('Desktop'),
    [string[]]$ExcludeSites,
    [switch]$Resume,
    [string]$LogFile,
    [switch]$EnhancedFileCount,
    [int]$MaxMemoryMB = 2048
)

# Global variable to store tenant total including OneDrive
$script:tenantTotalStorageGB = 0

# PowerShell Version Check - Require PowerShell 7+
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Error "This script requires PowerShell 7.0 or later for optimal performance. Current version: $($PSVersionTable.PSVersion)"
    Write-Host "Please upgrade to PowerShell 7+ and run again." -ForegroundColor Red
    exit 1
}
Write-Host "PowerShell $($PSVersionTable.PSVersion) detected." -ForegroundColor Green
# Initialize
$ErrorActionPreference = "Stop"
$ProgressPreference = "Continue"
$script:operationStartTime = Get-Date
$script:siteCache = @{}
$script:permissionCache = @{}
$script:groupCache = @{}
$script:groupMemberCache = @{} # ✅ ADDED: Cache for group members
$script:progressId = 1
$script:performanceCounters = @{
    ApiCalls = 0
    BatchedCalls = 0
    CacheHits = 0
    ProcessingTime = @{}
    ThrottleRetries = 0
}
# --- Restored Functions ---
# Retry logic with throttling detection (Restored as requested)
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
            # Note: The main Invoke-GraphApi function has its own, more advanced retry logic for API calls.
            # This generic function is restored for potential use in non-Graph API operations.
            return & $ScriptBlock
        } 
        catch {
            $lastError = $_
            $attempt++
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode -eq 429) {
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
# Delta query helper (Restored and Integrated)
function Get-GraphDelta {
    param(
        [Parameter(Mandatory)]
        [string]$Uri
    )
    $all = @()
    $nextLink = $Uri # Start with the initial (delta) URI
    $deltaLink = $null
    do {
        try {
            $response = Invoke-GraphApi -Uri $nextLink
            if ($response.value) { $all += $response.value }
            $nextLink = $response.'@odata.nextLink'
            # The deltaLink is present on the final page of a delta query response
            if ($response.'@odata.deltaLink') {
                $deltaLink = $response.'@odata.deltaLink'
            }
        } catch {
            Write-Log "Error during delta query paging for URI '$nextLink': $($_)" -Level Warning
            break
        }
    } while ($nextLink)
    return @{ Items = $all; DeltaLink = $deltaLink }
}
# --- End of Restored Functions ---
# Logging and Progress Functions
function Write-Log($Message, $Level = "Info", $ForegroundColor = "Cyan") {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    switch ($Level) {
        "Info"    { $color = if ($ForegroundColor) { $ForegroundColor } else { "Cyan" } }
        "Warning" { $color = "Yellow" }
        "Error"   { $color = "Red" }
        "Success" { $color = "Green" }
        "Debug"   { $color = "Gray" }
        default   { $color = "White" }
    }
    Write-Host $logMessage -ForegroundColor $color
    if ($script:logFilePath) { 
        try { 
            Add-Content -Path $script:logFilePath -Value $logMessage -ErrorAction SilentlyContinue
        } catch {
            # Ignore logging errors to prevent script termination
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
# Memory Management
function Test-MemoryUsage {
    param([int]$MaxMB = $MaxMemoryMB)
    $currentMemory = (Get-Process -Id $PID).WorkingSet / 1MB
    if ($currentMemory -gt $MaxMB) {
        Write-Log "Memory usage high: $([math]::Round($currentMemory, 2))MB / $MaxMB MB. Forcing garbage collection." -Level Warning
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        return $false
    }
    return $true
}
# Module Management
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
        Install-Module -Name $ModuleName -MinimumVersion $MinimumVersion -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        return $true
    }
    elseif ($Force -or [version]$installedModule.Version -lt [version]$MinimumVersion) {
        Write-Log "Updating module: $ModuleName from $($installedModule.Version) to $MinimumVersion" -Level Info
        try {
            # Attempt to uninstall gracefully first
            Get-InstalledModule -Name $ModuleName -AllVersions | Uninstall-Module -Force -ErrorAction SilentlyContinue
        } catch {
            Write-Log "Could not uninstall all versions of $ModuleName. Trying to install anyway." -Level Warning
        }
        Install-Module -Name $ModuleName -MinimumVersion $MinimumVersion -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        return $true
    }
    return $false
}
# Initialize log file
if ($LogFile) { 
    $script:logFilePath = $LogFile 
} else { 
    $script:logFilePath = Join-Path $OutputPath "SharePointAudit-$(Get-Date -Format 'yyyyMMdd-HHmmss').log" 
}
# Ensure output directory exists and define state file for resume functionality
$outputDir = Split-Path $script:logFilePath -Parent
$stateFilePath = Join-Path $outputDir "SharePointAudit-State.json"
if (-not (Test-Path $outputDir)) { 
    try {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    } catch {
        Write-Error "Cannot create output directory: $outputDir"
        exit 1
    }
}
Write-Log "Logging to: $script:logFilePath" -Level Success
if($Resume) {
    Write-Log "Resume mode enabled. State will be loaded from and saved to: $stateFilePath" -Level Info
}
# Enhanced module initialization with version checking
$requiredModules = @(
    @{Name="Microsoft.Graph.Sites"; Version="2.0.0"},
    @{Name="Microsoft.Graph.Files"; Version="2.0.0"},
    @{Name="Microsoft.Graph.Reports"; Version="2.0.0"},
    @{Name="Microsoft.Graph.Groups"; Version="2.0.0"},
    @{Name="ImportExcel"; Version="7.8.0"}
)
foreach ($moduleInfo in $requiredModules) {
    try {
        if (-not (Get-Module $moduleInfo.Name -ListAvailable)) {
            Write-Log "Installing module: $($moduleInfo.Name)" -Level Info
            Install-OrUpdateModule -ModuleName $moduleInfo.Name -MinimumVersion $moduleInfo.Version
        }
        Import-Module $moduleInfo.Name -ErrorAction Stop
        Write-Log "Loaded $($moduleInfo.Name) module" -Level Success
    }
    catch {
        Write-Log "Failed to load module $($moduleInfo.Name): $($_)" -Level Error
        Write-Log "Attempting to install/update module..." -Level Info
        try {
            Install-OrUpdateModule -ModuleName $moduleInfo.Name -MinimumVersion $moduleInfo.Version -Force
            Import-Module $moduleInfo.Name -ErrorAction Stop
            Write-Log "Successfully loaded $($moduleInfo.Name) module after installation." -Level Success
        } catch {
            Write-Log "Critical error: Cannot install/load required module $($moduleInfo.Name). Please install it manually and re-run the script." -Level Error
            throw "Module installation failed: $($moduleInfo.Name)"
        }
    }
}
# Authentication
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction Stop
    if (-not $cert) {
        throw "Certificate with thumbprint $CertificateThumbprint not found in CurrentUser\My certificate store"
    }
    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Certificate $cert -NoWelcome
    $context = Get-MgContext
    if (-not $context -or $context.AuthType -ne 'AppOnly') { 
        throw "App-only authentication required. Current auth type: $($context.AuthType)"
    }
    Write-Log "Connected to Microsoft Graph with App-Only authentication for tenant $($context.TenantId)" -Level Success
} catch {
    Write-Log "Authentication failed: $($_)" -Level Error
    throw
}

# Fixed Graph API function with proper batch handling
function Invoke-GraphApi {
    param(
        [Parameter(Mandatory)]
        [string]$Uri,
        [string]$Method = "GET",
        [object]$Body = $null,
        [hashtable]$Headers = $null,
        [switch]$MinimalResponse
    )
    
    $maxRetries = 5
    $attempt = 0
    
    if ($MinimalResponse) {
        if (-not $Headers) { $Headers = @{} }
        $Headers["Prefer"] = "return=minimal"
    }
    
    do {
        try {
            $script:performanceCounters.ApiCalls++
            
            $invokeParams = @{
                Method = $Method
                Uri = $Uri
            }
            
            if ($Headers) {
                $invokeParams.Headers = $Headers
            }
            
            if ($Body) {
                # ✅ Fixed: Proper JSON conversion for POST requests
                if ($Method -eq "POST" -and $Body -is [hashtable]) {
                    $invokeParams.Body = ($Body | ConvertTo-Json -Depth 10)
                    $invokeParams.ContentType = "application/json"
                } else {
                    $invokeParams.Body = $Body
                }
            }
            
            $result = Invoke-MgGraphRequest @invokeParams
            return $result
            
        } catch {
            $attempt++
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode -eq 429) {
                $script:performanceCounters.ThrottleRetries++
                $retryAfter = $_.Exception.Response.Headers["Retry-After"]
                $wait = if ($retryAfter) { [int]$retryAfter } else { [Math]::Pow(2, $attempt) }
                Write-Log "Throttled (429). Retrying in $wait seconds... (Attempt $attempt/$maxRetries)" -Level Warning
                Start-Sleep -Seconds $wait
            } elseif ($attempt -ge $maxRetries) {
                Write-Log "Max retries exceeded for $Uri" -Level Error
                throw $_
            } else {
                $wait = [Math]::Pow(2, $attempt)
                Write-Log "Error calling $Uri (Attempt $attempt/$maxRetries): $($_.Exception.Message). Retrying in $wait seconds." -Level Warning
                Start-Sleep -Seconds $wait
            }
        }
    } while ($attempt -lt $maxRetries)
}

# Paging helper
function Get-AllGraphPages {
    param(
        [Parameter(Mandatory)]
        [string]$InitialUri
    )
    $results = @()
    $uri = $InitialUri
    do {
        try {
            $response = Invoke-GraphApi -Uri $uri
            if ($response.value) { $results += $response.value }
            $uri = $response.'@odata.nextLink'
        } catch {
            Write-Log "Error during paging for URI starting with '$InitialUri': $($_)" -Level Warning
            break # Exit loop on error to avoid infinite loops
        }
    } while ($uri)
    return $results
}
# Batch helper (FIXED: Proper URL encoding and structure)
# Fixed Batch helper - Corrected URL construction and JSON handling
function Invoke-GraphBatch {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Requests
    )
    if ($Requests.Count -eq 0) { return @() }

    Write-Log "Preparing to batch $($Requests.Count) requests" -Level Debug
    $batches = for ($i = 0; $i -lt $Requests.Count; $i += 20) { 
        $Requests[$i..([Math]::Min($i+19, $Requests.Count-1))] 
    }
    $results = @()
    $batchNum = 0

    foreach ($batch in $batches) {
        $batchNum++
        Write-Log "Sending batch #$batchNum with $($batch.Count) requests" -Level Debug

        # ✅ Fixed: Properly construct batch requests
        $batchRequests = @()
        foreach ($req in $batch) {
            # Clean the URL properly - remove the full graph URL prefix
            $cleanUrl = $req.url -replace '^https://graph\.microsoft\.com/v1\.0', ''
            # Ensure URL starts with forward slash
            if (-not $cleanUrl.StartsWith('/')) { 
                $cleanUrl = '/' + $cleanUrl 
            }
            
            $batchRequests += @{
                id = $req.id
                method = "GET"
                url = $cleanUrl
            }
        }

        $batchRequest = @{
            requests = $batchRequests
        }

        try {
            $script:performanceCounters.BatchedCalls++
            # ✅ Fixed: Use the standard Invoke-GraphApi function with proper parameters
            $response = Invoke-GraphApi -Uri "https://graph.microsoft.com/v1.0/`$batch" -Method "POST" -Body $batchRequest
            
            if ($response.responses) {
                $results += $response.responses
                Write-Log "Batch #$batchNum succeeded: $($response.responses.Count) responses" -Level Debug
            }
        } catch {
            Write-Log "Batch #$batchNum failed: $($_.Exception.Message). Falling back to individual calls." -Level Warning
            # Fallback to individual API calls
            foreach ($req in $batch) {
                try {
                    $result = Invoke-GraphApi -Uri $req.url
                    $results += @{ id = $req.id; status = 200; body = $result }
                } catch {
                    Write-Log "Individual call failed for $($req.url): $($_.Exception.Message)" -Level Warning
                    $results += @{ id = $req.id; status = 500; body = @{ error = $_.Exception.Message } }
                }
            }
        }
    }
    return $results
}

# Refactored: Use unified Graph API helper for site retrieval
function Get-SiteWithCache($siteId) {
    if ($script:siteCache.ContainsKey($siteId)) { 
        $script:performanceCounters.CacheHits++
        return $script:siteCache[$siteId] 
    }
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId?`$select=id,displayName,webUrl,drive"
    $site = Invoke-GraphApi -Uri $uri -Method GET
    $script:siteCache[$siteId] = $site
    return $site
}
# Refactored: Use unified Graph API helper for permissions retrieval
function Get-PermissionsWithCache($siteId) {
    if ($script:permissionCache.ContainsKey($siteId)) { 
        $script:performanceCounters.CacheHits++
        return $script:permissionCache[$siteId] 
    }
    try {
        $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/permissions?`$top=999"
        $permissions = Get-AllGraphPages -InitialUri $uri
        $script:permissionCache[$siteId] = $permissions
        return $permissions
    } catch {
        Write-Log "Could not retrieve permissions for site $siteId : $($_)" -Level Warning
        $script:permissionCache[$siteId] = @() # Cache failure to avoid retries
        return @()
    }
}
# Refactored: Use unified Graph API helper for group retrieval
function Get-GroupWithCache($groupId) {
    if ($script:groupCache.ContainsKey($groupId)) { 
        $script:performanceCounters.CacheHits++
        return $script:groupCache[$groupId] 
    }
    try {
        $uri = "https://graph.microsoft.com/v1.0/groups/$groupId?`$select=id,displayName,groupTypes"
        $group = Invoke-GraphApi -Uri $uri -Method GET
        $script:groupCache[$groupId] = $group
        return $group
    } catch {
        Write-Log "Could not retrieve group $groupId : $($_)" -Level Warning
        $script:groupCache[$groupId] = $null # Cache failure to avoid retries
        return $null
    }
}

# ✅ ADDED: Function to get group members with caching
function Get-GroupMembersWithCache($groupId) {
    if ($script:groupMemberCache.ContainsKey($groupId)) {
        $script:performanceCounters.CacheHits++
        return $script:groupMemberCache[$groupId]
    }
    try {
        # Select only necessary properties to improve performance
        $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/members?`$select=id,displayName,userPrincipalName"
        $members = Get-AllGraphPages -InitialUri $uri
        $script:groupMemberCache[$groupId] = $members
        return $members
    } catch {
        Write-Log "Could not retrieve members for group $groupId : $($_)" -Level Warning
        $script:groupMemberCache[$groupId] = @() # Cache the failure to avoid repeated failed calls
        return @()
    }
}


# Fast SharePoint site discovery with direct API totals, now with Delta Query support
function Get-AllSharePointSites {
    param(
        [hashtable]$State
    )
    Write-Log "Discovering all SharePoint and OneDrive sites..." -Level Info
    $sitesFromState = @{}

    if ($Resume -and $State.sites) {
        Write-Log "Resume mode: Loading $($State.sites.Count) sites from state file." -Level Info
        $State.sites.GetEnumerator() | ForEach-Object { $sitesFromState[$_.Name] = $_.Value }
    }

    if ($Resume -and $State.siteDeltaLink) {
        Write-Log "Fetching changes using delta link..." -Level Info
        $deltaResult = Get-GraphDelta -Uri $State.siteDeltaLink
        foreach ($item in $deltaResult.Items) {
            if ($item.'@removed') {
                $sitesFromState.Remove($item.id)
            } else {
                $sitesFromState[$item.id] = $item
            }
        }
        $State.siteDeltaLink = $deltaResult.DeltaLink
    } else {
        Write-Log "Performing full scan of all sites..." -Level Info
        $allSites = Get-AllGraphPages -InitialUri "https://graph.microsoft.com/v1.0/sites"
        $deltaInit = Invoke-GraphApi -Uri "https://graph.microsoft.com/v1.0/sites/delta"
        $State.siteDeltaLink = $deltaInit.'@odata.deltaLink'
        $sitesFromState.Clear()
        $allSites | ForEach-Object { $sitesFromState[$_.id] = $_ }
    }

    # --- STEP 1: Get quota for ALL sites (including OneDrive) for tenant total ---
    Write-Log "Fetching quota data for $($sitesFromState.Values.Count) sites (including OneDrive) to calculate tenant totals..." -Level Debug
    $allBatchRequests = $sitesFromState.Values | ForEach-Object {
        @{ id = $_.id; url = "/sites/$($_.id)?`$select=id,displayName,webUrl,drive&`$expand=drive(`$select=quota)" }
    }
    $allBatchResponses = Invoke-GraphBatch -Requests $allBatchRequests

    $allSitesWithQuota = @()
    foreach ($resp in $allBatchResponses) {
        if ($resp.status -eq 200 -and $resp.body.drive.quota) {
            $allSitesWithQuota += $resp.body
        }
    }

    # ✅ Calculate tenant total including OneDrive
    $totalUsedBytes = ($allSitesWithQuota.drive.quota.used | Measure-Object -Sum).Sum
    $script:tenantTotalStorageGB = [math]::Round($totalUsedBytes / 1GB, 2)
    Write-Log "Tenant total storage (including OneDrive): $($script:tenantTotalStorageGB) GB from $($allSitesWithQuota.Count) sites" -Level Success

    # --- STEP 2: Filter to get ONLY SharePoint sites (exclude OneDrive) ---
    Write-Log "Filtering $($allSitesWithQuota.Count) sites to exclude OneDrive..." -Level Debug
    $filteredSites = $allSitesWithQuota | Where-Object {
        $_.webUrl -and 
        $_.webUrl -notlike "*-my.sharepoint.com*" -and
        $_.webUrl -notlike "*/personal/*" -and
        ($_.displayName -eq $null -or $_.displayName -notlike "*OneDrive*")
    } | Where-Object {
        $exclude = $false
        if ($ExcludeSites) {
            foreach ($pattern in $ExcludeSites) {
                if (($_.displayName -and $_.displayName -like $pattern) -or ($_.webUrl -and $_.webUrl -like $pattern)) { $exclude = $true; break }
            }
        }
        -not $exclude
    }

    # Debug: Log some sample URLs to understand the data
    if ($allSitesWithQuota.Count -gt 0) {
        Write-Log "Sample URLs from all sites:" -Level Debug
        $allSitesWithQuota | Select-Object -First 5 | ForEach-Object {
            $siteName = if ($_.displayName) { $_.displayName } else { "Unknown Site" }
            Write-Log "  - ${siteName}: $($_.webUrl)" -Level Debug
        }
    }

    Write-Log "Found $($filteredSites.Count) SharePoint sites (excluded OneDrive and filtered)" -Level Success
    return $filteredSites
}

# Fixed Get-SiteDetails function with comprehensive null checks
function Get-SiteDetails($site) {
    # Validate input
    if (-not $site) {
        Write-Log "Get-SiteDetails called with null site object" -Level Error
        return $null
    }
    
    if (-not $site.id) {
        Write-Log "Site object missing required 'id' property: $($site | ConvertTo-Json -Compress)" -Level Error
        return $null
    }

    $details = [ordered]@{
        Id = $site.id
        SiteName = if ($site.displayName) { $site.displayName } else { "Unknown Site" }
        SiteUrl = if ($site.webUrl) { $site.webUrl } else { "Unknown URL" }
        StorageGB = 0
        RecycleBinGB = 0
        StorageLimitGB = 0
        TotalFiles = 0
        HasMicrosoft365Group = "No"
        GroupName = ""
        Owners = @()
        Members = @()
        ExternalGuests = @()
        StorageUsedPercentOfTenant = 0
    }
    
    # Get storage info with null checks
    try {
        if ($site.drive -and $site.drive.quota) {
            $quota = $site.drive.quota
            if ($null -ne $quota.used) {
                $details.StorageGB = [math]::Round($quota.used / 1GB, 2)
            }
            if ($null -ne $quota.deleted) {
                $details.RecycleBinGB = [math]::Round($quota.deleted / 1GB, 2)
            }
            if ($null -ne $quota.total) {
                $details.StorageLimitGB = [math]::Round($quota.total / 1GB, 2)
            }
            if ($null -ne $quota.fileCount) {
                $details.TotalFiles = $quota.fileCount
            }
        } else {
            Write-Log "Site $($details.SiteName) has no drive or quota information" -Level Debug
        }
    } catch {
        Write-Log "Error processing quota for site $($details.SiteName): $($_)" -Level Warning
    }
    
    # Calculate storage used % of tenant (using total including OneDrive)
    try {
        if ($script:tenantTotalStorageGB -gt 0) {
            $siteTotalStorage = $details.StorageGB + $details.RecycleBinGB
            $details.StorageUsedPercentOfTenant = [math]::Round(($siteTotalStorage / $script:tenantTotalStorageGB) * 100, 4)
        }
    } catch {
        Write-Log "Error calculating storage percentage for site $($details.SiteName): $($_)" -Level Warning
    }
    
    # Enhanced file count with null checks
    if ($details.TotalFiles -eq 0 -and $EnhancedFileCount) {
        try {
            if ($details.SiteUrl -and $details.SiteUrl -ne "Unknown URL") {
                $encodedUrl = [System.Web.HttpUtility]::UrlEncode($details.SiteUrl)
                $uri = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D7')?`$filter=siteUrl eq '$($details.SiteUrl)'"
                $siteUsage = Invoke-GraphApi -Uri $uri -Method GET
                if ($siteUsage -and $siteUsage.value -and $siteUsage.value.Count -gt 0) {
                    $fileCount = $siteUsage.value[0].fileCount
                    if ($null -ne $fileCount) {
                        $details.TotalFiles = $fileCount
                    }
                }
            }
        } catch {
            Write-Log "Could not get file count from reports for $($details.SiteName): $($_)" -Level Debug
        }
    }
    
    # Check for Microsoft 365 Group with optimized error handling (no retries for expected errors)
    try {
        if ($site.id) {
            $uri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/group"
            try {
                # Use direct Invoke-MgGraphRequest to avoid unnecessary retries for expected 400 errors
                $group = Invoke-MgGraphRequest -Method GET -Uri $uri
                if ($group -and $group.id) {
                    $details.HasMicrosoft365Group = "Yes"
                    $details.GroupName = if ($group.displayName) { $group.displayName } else { "Unknown Group" }
                }
            } catch {
                # Check for specific "Resource not found for the segment 'group'" error (400) or Not Found (404)
                if ($_.Exception.Response -and 
                    ($_.Exception.Response.StatusCode -eq 400 -or $_.Exception.Response.StatusCode -eq 404)) {
                    Write-Log "Site $($details.SiteName) is not connected to a Microsoft 365 Group" -Level Debug
                } else {
                    # For other errors, use the retry-enabled function
                    Write-Log "Retrying group check for $($details.SiteName) due to non-expected error: $($_)" -Level Debug
                    $group = Invoke-GraphApi -Uri $uri
                    if ($group -and $group.id) {
                        $details.HasMicrosoft365Group = "Yes"
                        $details.GroupName = if ($group.displayName) { $group.displayName } else { "Unknown Group" }
                    }
                }
            }
        }
    } catch {
        Write-Log "Error in group check logic for $($details.SiteName): $($_)" -Level Warning
    }
    
    # ✅ FIXED: Get permissions, now handles both direct user access and group-based access
    try {
        if ($site.id) {
            $permissions = Get-PermissionsWithCache -SiteId $site.id
            if ($permissions -and $permissions.Count -gt 0) {
                foreach ($perm in $permissions) {
                    try {
                        # Skip if no identity or roles are assigned
                        if (-not ($perm -and $perm.grantedToV2 -and $perm.roles -and $perm.roles.Count -gt 0)) {
                            continue
                        }

                        $role = $perm.roles -join ', '

                        # Helper Action to add a user to the details object, avoids code duplication
                        $addUserAction = {
                            param($userObject, $assignedRole)
                            
                            # Skip if no user object or UPN
                            if (-not ($userObject -and $userObject.userPrincipalName)) { return }

                            $userPrincipalName = $userObject.userPrincipalName
                            $userType = if ($userPrincipalName -like "*#EXT#*") { "External Guest" } else { "Internal" }
                            
                            $userObjToAdd = [PSCustomObject]@{
                                DisplayName       = if ($userObject.displayName) { $userObject.displayName } else { "Unknown User" }
                                UserPrincipalName = $userPrincipalName
                                UserType          = $userType
                                Role              = $assignedRole
                                SiteName          = $details.SiteName
                                SiteUrl           = $details.SiteUrl
                            }

                            # Categorize user by role
                            if ($assignedRole -match 'owner|admin') { 
                                $details.Owners += $userObjToAdd
                            } else { 
                                $details.Members += $userObjToAdd
                            }
                            
                            if ($userType -eq "External Guest") { 
                                $details.ExternalGuests += $userObjToAdd
                            }
                        }

                        # Case 1: Permission is granted directly to a user
                        if ($perm.grantedToV2.user) {
                            & $addUserAction -userObject $perm.grantedToV2.user -assignedRole $role
                        }

                        # Case 2: Permission is granted to a group
                        elseif ($perm.grantedToV2.group) {
                            $groupId = $perm.grantedToV2.group.id
                            if ($groupId) {
                                Write-Log "Expanding group permission for group ID $groupId with role '$role' on site $($details.SiteName)" -Level Debug
                                $groupMembers = Get-GroupMembersWithCache -GroupId $groupId
                                foreach ($member in $groupMembers) {
                                    # The member object from /members is a user object, so pass it to the helper
                                    & $addUserAction -userObject $member -assignedRole $role
                                }
                            }
                        }

                    } catch {
                        Write-Log "Error processing a specific permission entry for site $($details.SiteName): $($_)" -Level Warning
                        continue
                    }
                }
            } else {
                Write-Log "No permissions found for site $($details.SiteName)" -Level Debug
            }
        }
    } catch {
        Write-Log "General error retrieving permissions for site $($details.SiteName): $($_)" -Level Warning
    }
    
    return $details
}
# Advanced Excel Export Function
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
        [switch]$AutoSize = $true,
        [switch]$BoldTopRow = $true,
        [switch]$FreezeTopRow = $true
    )
    if (-not $Data -or $Data.Count -eq 0) {
        Write-Log "No data to export for worksheet '$WorksheetName'. Creating placeholder worksheet." -Level Warning
        $Data = @([PSCustomObject]@{ Message = "No data available for this report." })
    }
    try {
        $params = @{
            Path = $Path
            WorksheetName = $WorksheetName
            InputObject = $Data
            TableName = "Table_$($WorksheetName.Replace(' ','_'))"
            TableStyle = $TableStyle
            AutoSize = $AutoSize
            BoldTopRow = $BoldTopRow
            FreezeTopRow = $FreezeTopRow
            ErrorAction = 'Stop'
        }
        if ($Title) { $params.Title = $Title }
        Export-Excel @params
        Write-Log "Successfully exported worksheet: $WorksheetName with $($Data.Count) rows" -Level Success
    } catch {
        Write-Log "Failed to export worksheet '$WorksheetName': $($_)" -Level Error
        # Create a basic CSV as fallback
        try {
            $csvPath = $Path -replace '\.xlsx$', "_$WorksheetName.csv"
            $Data | Export-Csv -Path $csvPath -NoTypeInformation -UseCulture
            Write-Log "Created fallback CSV: $csvPath" -Level Info
        } catch {
            Write-Log "Failed to create fallback CSV for $WorksheetName" -Level Error
        }
    }
}
# FIXED: Use efficient recursive search for file/folder retrieval
function Get-ComprehensiveFileFolderReport($site) {
    $siteName = if ($site.displayName) { $site.displayName } else { if ($site.name) { $site.name } else { $site.id } }
    Write-Log "Analyzing files and folders for largest site: ${siteName}" -Level Info
    $allItemsReport = @()
    $folderSizes = [System.Collections.Generic.Dictionary[string, int64]]::new()
    $files = @()
    try {
        # Use search to recursively get all items. This is far more efficient than manual recursion.
        $uri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/drive/root/search(q='')?`$top=999&`$select=name,file,folder,size,parentReference,lastModifiedDateTime,webUrl"
        $allItems = Get-AllGraphPages -InitialUri $uri
        if ($allItems -and $allItems.Count -gt 0) {
            Write-Log "Processing $($allItems.Count) total items found in ${siteName}..." -Level Info
            # First pass: process files and aggregate folder sizes
            foreach ($item in $allItems) {
                if ($item.file -and $item.size -gt 0) {
                    $filePath = ($item.parentReference.path -split 'root:')[-1]
                    if ([string]::IsNullOrEmpty($filePath)) { $filePath = "/" }
                    $fileObj = [PSCustomObject]@{
                        Name = $item.name
                        Size = [int64]$item.size
                        Path = $filePath
                        LastModified = if ($item.lastModifiedDateTime) { [DateTime]$item.lastModifiedDateTime } else { [DateTime]::MinValue }
                        WebUrl = $item.webUrl
                    }
                    $files += $fileObj
                    # Aggregate folder sizes
                    if (-not $folderSizes.ContainsKey($filePath)) {
                        $folderSizes[$filePath] = 0
                    }
                    $folderSizes[$filePath] += $fileObj.Size
                }
            }
            Write-Log "Found $($files.Count) files in $($folderSizes.Count) folders" -Level Info
            $siteTotalStorageBytes = ($files.Size | Measure-Object -Sum).Sum
            $siteTotalStorageGB = [math]::Round($siteTotalStorageBytes / 1GB, 2)
            # Get top 10 largest files
            $topFiles = $files | Sort-Object Size -Descending | Select-Object -First 10
            $rank = 1
            foreach ($file in $topFiles) {
                $allItemsReport += [PSCustomObject]@{
                    Type = "File"
                    Name = $file.Name
                    SizeGB = [math]::Round($file.Size / 1GB, 4)
                    Path = $file.Path
                    URL = $file.WebUrl
                    LastModified = if ($file.LastModified -ne [DateTime]::MinValue) { $file.LastModified.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
                    Category = "Top 10 Largest File"
                    Rank = $rank
                    StorageUsedPercentOfTenant = if ($script:tenantTotalStorageGB -gt 0) { [math]::Round(($file.Size / 1GB / $script:tenantTotalStorageGB) * 100, 4) } else { 0 }
                }
                $rank++
            }
            # Get all files over 1GB
            $largeFiles = $files | Where-Object { $_.Size -gt 1GB -and $_.Name -notin $topFiles.Name } | Sort-Object Size -Descending
            foreach ($file in $largeFiles) {
                $allItemsReport += [PSCustomObject]@{
                    Type = "File"; Name = $file.Name; SizeGB = [math]::Round($file.Size / 1GB, 4); Path = $file.Path; URL = $file.WebUrl; LastModified = if ($file.LastModified -ne [DateTime]::MinValue) { $file.LastModified.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }; Category = "Large File (>1GB)"; Rank = ""; StorageUsedPercentOfTenant = if ($script:tenantTotalStorageGB -gt 0) { [math]::Round(($file.Size / 1GB / $script:tenantTotalStorageGB) * 100, 4) } else { 0 }
                }
            }
            # Get top 10 largest folders
            $topFolders = $folderSizes.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 10
            $rank = 1
            foreach ($folder in $topFolders) {
                $folderName = Split-Path -Path $folder.Key -Leaf
                if ([string]::IsNullOrEmpty($folderName)) { $folderName = "Root" }
                $allItemsReport += [PSCustomObject]@{
                    Type = "Folder"; Name = $folderName; SizeGB = [math]::Round($folder.Value / 1GB, 4); Path = $folder.Key; URL = ""; LastModified = "N/A"; Category = "Top 10 Largest Folder"; Rank = $rank; StorageUsedPercentOfTenant = if ($script:tenantTotalStorageGB -gt 0) { [math]::Round(($folder.Value / 1GB / $script:tenantTotalStorageGB) * 100, 4) } else { 0 }
                }
                $rank++
            }
            # Get all folders over 1GB
            $largeFolders = $folderSizes.GetEnumerator() | Where-Object { $_.Value -gt 1GB -and $_.Key -notin $topFolders.Key } | Sort-Object Value -Descending
            foreach ($folder in $largeFolders) {
                $folderName = Split-Path -Path $folder.Key -Leaf
                if ([string]::IsNullOrEmpty($folderName)) { $folderName = "Root" }
                 $allItemsReport += [PSCustomObject]@{
                    Type = "Folder"; Name = $folderName; SizeGB = [math]::Round($folder.Value / 1GB, 4); Path = $folder.Key; URL = ""; LastModified = "N/A"; Category = "Large Folder (>1GB)"; Rank = ""; StorageUsedPercentOfTenant = if ($script:tenantTotalStorageGB -gt 0) { [math]::Round(($folder.Value / 1GB / $script:tenantTotalStorageGB) * 100, 4) } else { 0 }
                }
            }
        }
        Write-Log "Generated file/folder report with $($allItemsReport.Count) items." -Level Success
        return $allItemsReport
    } catch {
        Write-Log "File/folder analysis failed for ${siteName}: $($_)" -Level Error
        return @([PSCustomObject]@{ Message = "Detailed file/folder analysis failed. See log for details." })
    }
}
# Main processing
try {
    Write-Log "SharePoint Audit Started" -Level Success
    # Load state for Resume mode
    $script:State = @{}
    if ($Resume -and (Test-Path $stateFilePath)) {
        try {
            $script:State = Get-Content -Path $stateFilePath | ConvertFrom-Json -AsHashtable
            Write-Log "Successfully loaded previous state from $stateFilePath" -Level Info
        } catch {
            Write-Log "Could not read or parse state file at $stateFilePath. Performing a full run." -Level Warning
            $script:State = @{}
        }
    }
    # Get all SharePoint sites (excluding OneDrive from reporting but including in tenant totals)
    $sites = Get-AllSharePointSites -State $script:State

    # Fixed main processing loop with comprehensive null checks
    try {
        Write-Log "Processing $($sites.Count) SharePoint sites..." -Level Info
        
        # Validate sites array
        if (-not $sites -or $sites.Count -eq 0) { 
            throw "No SharePoint sites found. Please check App Registration permissions and authentication." 
        }
        
        # Debug: Check for sites with null IDs
        $sitesWithNullIds = @($sites | Where-Object { -not $_.id })
        Write-Log "Sites with null IDs: $($sitesWithNullIds.Count)" -Level Debug
        
        if ($sitesWithNullIds.Count -gt 0) {
            Write-Log "Found sites with null IDs. This may cause processing errors." -Level Warning
            # Log details of problematic sites
            foreach ($nullSite in $sitesWithNullIds) {
                Write-Log "Null ID site: $($nullSite | ConvertTo-Json -Compress)" -Level Debug
            }
            # Filter out sites with null IDs
            $sites = @($sites | Where-Object { $_.id })
            Write-Log "After filtering null IDs: $($sites.Count) sites remaining" -Level Info
        }
        
        # ✅ FIXED: Initialize processed site IDs with proper null checks
        $processedSiteIds = if ($Resume -and $script:State -and $script:State.processedSiteIds) { 
            try {
                # Ensure we get a proper list object
                $existingIds = $script:State.processedSiteIds
                if ($existingIds -is [array]) {
                    [System.Collections.Generic.List[string]]::new($existingIds)
                } elseif ($existingIds -is [System.Collections.Generic.List[string]]) {
                    $existingIds
                } else {
                    Write-Log "Invalid processedSiteIds type in state: $($existingIds.GetType().Name). Creating new list." -Level Warning
                    [System.Collections.Generic.List[string]]::new()
                }
            } catch {
                Write-Log "Error loading processedSiteIds from state: $($_). Creating new list." -Level Warning
                [System.Collections.Generic.List[string]]::new()
            }
        } else { 
            [System.Collections.Generic.List[string]]::new()
        }
        
        # Validate that processedSiteIds is not null and has the Contains method
        if (-not $processedSiteIds) {
            Write-Log "processedSiteIds is null, creating new list" -Level Warning
            $processedSiteIds = [System.Collections.Generic.List[string]]::new()
        }
        
        # Test the Contains method to ensure it works
        try {
            $testResult = $processedSiteIds.Contains("test-id")
            Write-Log "processedSiteIds.Contains() method test successful" -Level Debug
        } catch {
            Write-Log "processedSiteIds.Contains() method failed: $($_). Recreating list." -Level Warning
            $processedSiteIds = [System.Collections.Generic.List[string]]::new()
        }
        
        $newlyProcessedSiteIds = [System.Collections.Generic.List[string]]::new()
        $allSiteDetails = @()
        
        # ✅ FIXED: Filter sites to process with safe null checks
        Write-Log "Filtering sites to process..." -Level Debug
        $sitesToProcess = @()
        foreach ($site in $sites) {
            if (-not $site) {
                Write-Log "Skipping null site object" -Level Warning
                continue
            }
            if (-not $site.id) {
                $siteName = if ($site.displayName) { $site.displayName } else { "Unknown Site" }
                Write-Log "Skipping site without ID: $siteName" -Level Warning
                continue
            }
            
            # Safe check for already processed sites
            $isAlreadyProcessed = try {
                $processedSiteIds.Contains($site.id)
            } catch {
                Write-Log "Error checking if site $($site.id) is already processed: $($_)" -Level Warning
                $false  # Assume not processed if we can't check
            }
            
            if (-not $isAlreadyProcessed) {
                $sitesToProcess += $site
            }
        }
        
        $skippedCount = $sites.Count - $sitesToProcess.Count
        if ($skippedCount -gt 0) {
            Write-Log "Resuming: Skipping $skippedCount already processed sites." -Level Info
            # Load details for skipped sites with null checks
            if ($script:State -and $script:State.allSiteDetails) {
                try {
                    $existingSiteDetails = $script:State.allSiteDetails
                    if ($existingSiteDetails -and $processedSiteIds.Count -gt 0) {
                        foreach ($existingDetail in $existingSiteDetails) {
                            if ($existingDetail -and $existingDetail.Id -and $processedSiteIds.Contains($existingDetail.Id)) {
                                $allSiteDetails += $existingDetail
                            }
                        }
                    }
                } catch {
                    Write-Log "Error loading existing site details from state: $($_)" -Level Warning
                }
            }
        }
        
        Write-Log "Processing $($sitesToProcess.Count) new sites..." -Level Info
        
        if ($sitesToProcess.Count -eq 0 -and $allSiteDetails.Count -eq 0) {
            Write-Log "No sites to process and no existing site details found." -Level Warning
        }
        
        $siteCounter = 0
        foreach ($site in $sitesToProcess) {
            $siteCounter++
            
            # Additional validation
            if (-not $site) {
                Write-Log "Skipping null site at position $siteCounter" -Level Warning
                continue
            }
            
            if (-not $site.id) {
                $siteName = if ($site.displayName) { $site.displayName } else { if ($site.name) { $site.name } else { "Unknown Site" } }
                Write-Log "Skipping site without ID at position ${siteCounter}: ${siteName}" -Level Warning
                continue
            }
            
            # Get safe site name for display
            $siteName = if ($site.displayName) { 
                $site.displayName 
            } elseif ($site.name) { 
                $site.name 
            } else { 
                $site.id 
            }
            
            Show-Progress -Activity "Processing Sites" -Status "Site $siteCounter/$($sitesToProcess.Count): ${siteName}" -PercentComplete ([int](($siteCounter / $sitesToProcess.Count) * 100))
            
            try {
                $details = Get-SiteDetails -site $site
                if ($details) {
                    $allSiteDetails += $details
                    try {
                        $newlyProcessedSiteIds.Add($site.id)
                    } catch {
                        Write-Log "Error adding site ID to processed list: $($_)" -Level Warning
                    }
                } else {
                    Write-Log "Get-SiteDetails returned null for site: ${siteName}" -Level Warning
                    # Create error placeholder
                    $errorSite = [ordered]@{
                        Id = $site.id
                        SiteName = "${siteName} (NULL RESULT ERROR)"
                        SiteUrl = if ($site.webUrl) { $site.webUrl } else { "Unknown" }
                        StorageGB = 0; RecycleBinGB = 0; StorageLimitGB = 0; TotalFiles = 0
                        HasMicrosoft365Group = "Error"; GroupName = "Error"
                        Owners = @(); Members = @(); ExternalGuests = @()
                        StorageUsedPercentOfTenant = 0
                    }
                    $allSiteDetails += $errorSite
                }
            } catch {
                Write-Log "FATAL: Failed to process site ${siteName} due to error: $($_)" -Level Error
                Write-Log "Site object: $($site | ConvertTo-Json -Compress)" -Level Debug
                
                # Create error placeholder
                $errorSite = [ordered]@{
                    Id = if ($site.id) { $site.id } else { "unknown-$(Get-Date -Format 'yyyyMMddHHmmss')" }
                    SiteName = "${siteName} (PROCESSING ERROR)"
                    SiteUrl = if ($site.webUrl) { $site.webUrl } else { "Unknown" }
                    StorageGB = 0; RecycleBinGB = 0; StorageLimitGB = 0; TotalFiles = 0
                    HasMicrosoft365Group = "Error"; GroupName = "Error"
                    Owners = @(); Members = @(); ExternalGuests = @()
                    StorageUsedPercentOfTenant = 0
                }
                $allSiteDetails += $errorSite
            }
            
            # Memory check every 10 sites
            if ($siteCounter % 10 -eq 0) {
                Test-MemoryUsage
            }
        }
        
        Stop-Progress -Activity "Processing Sites"
        
        # Final validation of results
        if (-not $allSiteDetails) {
            $allSiteDetails = @()
        }
        
        $siteDetails = $allSiteDetails
        
        if ($siteDetails.Count -eq 0) {
            throw "No site details could be processed. Check permissions and site access."
        }
        
        Write-Log "Successfully processed $($siteDetails.Count) total sites" -Level Success
        
    } catch {
        Write-Log "Error in main processing loop: $($_)" -Level Error
        Write-Log "Error details: $($_.Exception.GetType().FullName)" -Level Error
        if ($_.ScriptStackTrace) {
            Write-Log "Script stack trace: $($_.ScriptStackTrace)" -Level Debug
        }
        throw
    }
    # Generate Excel report
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $excelFile = Join-Path $OutputPath "SharePointAudit-$timestamp.xlsx"
    Write-Log "Generating Excel report: $excelFile" -Level Info
    if (Test-Path $excelFile) { 
        try { Remove-Item $excelFile -Force } catch { Write-Log "Could not remove existing report file. It may be open." -Level Warning }
    }
    # Worksheet 1: Site Summary
    Show-Progress -Activity "Generating Excel Report" -Status "Creating Site Summary..." -PercentComplete 33
    $siteSummaryData = $siteDetails | Select-Object \
        @{N='Site Name';E={ $_.SiteName }}, \
        @{N='Storage Used GB';E={ [math]::Round([double]$_.StorageGB, 2) }}, \
        @{N='Recycle Bin GB';E={ [math]::Round([double]$_.RecycleBinGB, 2) }}, \
        @{N='Total Storage GB';E={ [math]::Round(([double]$_.StorageGB + [double]$_.RecycleBinGB), 2) }}, \
        @{N='Storage Limit GB';E={ if ($_.StorageLimitGB -ne $null -and $_.StorageLimitGB -gt 0) { [math]::Round([double]$_.StorageLimitGB, 2) } else { "N/A" } }}, \
        @{N='Storage Used % of Tenant';E={ if ($script:tenantTotalStorageGB -gt 0) { "{0:P4}" -f (([double]$_.StorageGB + [double]$_.RecycleBinGB) / $script:tenantTotalStorageGB) } else { "N/A" } }}, \
        @{N='File Count';E={ if ($_.TotalFiles -ne $null) { $_.TotalFiles } else { 0 } }}, \
        @{N='Linked M365 Group';E={ if ($_.GroupName) { $_.GroupName } else { "No" } }}, \
        @{N='Owners';E={ if ($_.Owners -is [System.Collections.IEnumerable]) { ($_.Owners | Where-Object { $_ -ne $null }).Count } else { 0 } }}, \
        @{N='Members+Visitors';E={ if ($_.Members -is [System.Collections.IEnumerable]) { ($_.Members | Where-Object { $_ -ne $null }).Count } else { 0 } }}, \
        @{N='URL';E={ $_.SiteUrl }}
    Export-ExcelWorksheet -Data $siteSummaryData -Path $excelFile -WorksheetName "Site_Summary" -Title "SharePoint Site Summary Report"
    
    # Worksheet 2: User Access Summary
    Show-Progress -Activity "Generating Excel Report" -Status "Creating User Access Summary..." -PercentComplete 66
    $userAccessData = [System.Collections.Generic.Dictionary[string,psobject]]::new()
    foreach ($site in $siteDetails) {
        $allUsers = @()
        if ($site.Owners) { $allUsers += $site.Owners }
        if ($site.Members) { $allUsers += $site.Members }
        if ($site.ExternalGuests) { $allUsers += $site.ExternalGuests }
        # Add any other permissioned users if present
        if ($site.AllUsers) { $allUsers += $site.AllUsers }
        foreach ($user in $allUsers) {
            if ($user -and $user.UserPrincipalName) {
                $key = $user.UserPrincipalName
                if (-not $userAccessData.ContainsKey($key)) {
                    $userAccessData[$key] = [PSCustomObject]@{
                        'User/Group' = $user.DisplayName
                        'User Principal Name' = $key
                        'Type' = $user.UserType
                        Sites = [System.Collections.Generic.List[string]]::new()
                        Roles = [System.Collections.Generic.List[string]]::new()
                    }
                }
                if (-not $userAccessData[$key].Sites.Contains($site.SiteName)) {
                    $userAccessData[$key].Sites.Add($site.SiteName)
                }
                if ($user.Role -and (-not $userAccessData[$key].Roles.Contains($user.Role))) {
                    $userAccessData[$key].Roles.Add($user.Role)
                }
            }
        }
    }
    
    # ✅ FIXED: Initialize as an empty array to prevent null argument error if no users are found.
    $userAccessExport = @()
    if ($userAccessData.Values.Count -gt 0) {
        $userAccessExport = $userAccessData.Values | ForEach-Object {
            [PSCustomObject]@{ 'User/Group' = $_.'User/Group'; 'User Principal Name' = $_.'User Principal Name'; Type = $_.Type; 'Site Count' = $_.Sites.Count; Sites = $_.Sites -join "; "; 'Roles (Unique)' = (($_.Roles | Sort-Object -Unique) -join ", ") }
        }
    }
    Export-ExcelWorksheet -Data $userAccessExport -Path $excelFile -WorksheetName "User_Access_Summary" -Title "User Access Summary Report"

    # Worksheet 3: Top Files & Folders Analysis for largest site
    Show-Progress -Activity "Generating Excel Report" -Status "Analyzing largest site for files/folders..." -PercentComplete 90
    $largestSiteDetails = $siteDetails | Where-Object { $_.SiteName -notlike "*ERROR*" } | Sort-Object -Property @{Expression={$_.StorageGB + $_.RecycleBinGB}} -Descending | Select-Object -First 1
    if ($largestSiteDetails -and $largestSiteDetails.StorageGB -gt 0) {
        $siteForAnalysis = $sites | Where-Object { $_.id -eq $largestSiteDetails.Id } | Select-Object -First 1
        if ($siteForAnalysis) {
            $fileFolderReport = Get-ComprehensiveFileFolderReport -site $siteForAnalysis
            Export-ExcelWorksheet -Data $fileFolderReport -Path $excelFile -WorksheetName "Largest_Site_Analysis" -Title "Top Files & Folders Analysis - $($largestSiteDetails.SiteName)"
        }
    } else {
        Write-Log "No suitable site found for file/folder analysis." -Level Warning
        Export-ExcelWorksheet -Data @() -Path $excelFile -WorksheetName "Largest_Site_Analysis" -Title "Top Files & Folders Analysis"
    }
    Stop-Progress -Activity "Generating Excel Report"
    # Final Summary
    $totalTime = (New-TimeSpan -Start $script:operationStartTime -End (Get-Date)).ToString("g")
    $totalStorage = ($siteDetails.StorageGB | Measure-Object -Sum).Sum
    $totalRecycle = ($siteDetails.RecycleBinGB | Measure-Object -Sum).Sum
    $totalFiles = ($siteDetails.TotalFiles | Measure-Object -Sum).Sum
    Write-Host "`n=============================================" -ForegroundColor Green
    Write-Host "          SHAREPOINT AUDIT SUMMARY" -ForegroundColor Green
    Write-Host "=============================================" -ForegroundColor Green
    Write-Host "`nProcessing Statistics:" -ForegroundColor Cyan
    Write-Host "  SharePoint Sites Found:    $($sites.Count)"
    Write-Host "  Sites Processed:           $($siteDetails.Count)"
    Write-Host "  Total Storage Used:       $([math]::Round($totalStorage, 2)) GB"
    Write-Host "  Recycle Bin Storage:      $([math]::Round($totalRecycle, 2)) GB"
    Write-Host "  Total with Recycle Bin:   $([math]::Round(($totalStorage + $totalRecycle), 2)) GB"
    Write-Host "  Total Files Found:        $totalFiles"
    Write-Host "  Unique Users Found:       $($userAccessData.Count)"
    Write-Host "`nPerformance Metrics:" -ForegroundColor Cyan
    Write-Host "  Total API Calls:           $($script:performanceCounters.ApiCalls)"
    Write-Host "  Batched API Calls:        $($script:performanceCounters.BatchedCalls)"
    Write-Host "  Cache Hits:               $($script:performanceCounters.CacheHits)"
    Write-Host "  Throttle Retries:         $($script:performanceCounters.ThrottleRetries)"
    Write-Host "  PowerShell Version:        $($PSVersionTable.PSVersion)"
    Write-Host "  Total Execution Time:      $totalTime"
    Write-Host "`nReport Generated:" -ForegroundColor Cyan
    Write-Host "  Excel Report:             $excelFile"
    Write-Host "  Log File:                $script:logFilePath"
    Write-Host "`n=============================================" -ForegroundColor Green
    Write-Host "               AUDIT COMPLETE" -ForegroundColor Green
    Write-Host "=============================================" -ForegroundColor Green
    Write-Log "Audit completed successfully! Report saved to: $excelFile" -Level Success
}
catch {
    Write-Log "Script failed with a terminating error: $($_)" -Level Error
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level Debug
    Write-Host "`nScript execution failed. Check the log file for details: $script:logFilePath" -ForegroundColor Red
    exit 1
}
finally {
    # Save state at the end of the script, regardless of success or failure
    if ($Resume) {
        Write-Log "Saving current state for resume capability..." -Level Info
        try {
            # Add newly processed sites to the master list
            $processedSiteIds.AddRange($newlyProcessedSiteIds)
            $script:State.processedSiteIds = $processedSiteIds | Sort-Object -Unique
            # Save all site details collected so far for faster report generation on resume
            $script:State.allSiteDetails = $siteDetails | Select-Object * -ExcludeProperty Owners, Members, ExternalGuests # Exclude large nested objects to keep state file smaller
            $script:State | ConvertTo-Json -Depth 5 | Set-Content -Path $stateFilePath
            Write-Log "State saved successfully to $stateFilePath" -Level Success
        } catch {
            Write-Log "Failed to save state file: $_" -Level Error
        }
    }
    try { 
        Disconnect-MgGraph -ErrorAction SilentlyContinue 
        Write-Log "Disconnected from Microsoft Graph." -Level Info
    } catch {
        Write-Log "Error disconnecting from Graph: $($_)" -Level Warning
    }
}