# Author: Tim MacLatchy
# Date: 05/11/2024
# License: MIT License
# Description: This script retrieves SharePoint sites and external users with their permissions,
# outputting Resource Path, Item Type, Permission, User Name, and User Email.

# Required Modules with specific versions
$moduleRequirements = @(
    @{Name = 'Microsoft.Graph'; MinimumVersion = '2.0.0'},
    @{Name = 'PnP.PowerShell'; MinimumVersion = '2.0.0'},
    @{Name = 'ImportExcel'; MinimumVersion = '7.0.0'}
)

# Function to ensure proper module installation and importing
function Initialize-RequiredModules {
    foreach ($moduleReq in $moduleRequirements) {
        $moduleName = $moduleReq.Name
        $moduleVersion = $moduleReq.MinimumVersion
        
        Write-Host "Checking module: $moduleName..."
        
        # Check if module is installed with required version
        $installedModule = Get-Module -ListAvailable -Name $moduleName | 
            Where-Object { $_.Version -ge $moduleVersion }
        
        if (-not $installedModule) {
            try {
                Write-Host "Installing $moduleName module version $moduleVersion or higher..."
                Install-Module -Name $moduleName -MinimumVersion $moduleVersion -Force -Scope CurrentUser -AllowClobber -ErrorAction Stop
                Write-Host "Successfully installed $moduleName" -ForegroundColor Green
            }
            catch {
                Write-Host "Error installing $moduleName module: $_" -ForegroundColor Red
                return $false
            }
        }
        
        try {
            # Force reimport module to ensure clean state
            Remove-Module -Name $moduleName -ErrorAction SilentlyContinue
            Import-Module -Name $moduleName -MinimumVersion $moduleVersion -Force -ErrorAction Stop
            Write-Host "Successfully imported $moduleName" -ForegroundColor Green
        }
        catch {
            Write-Host "Error importing $moduleName module: $_" -ForegroundColor Red
            return $false
        }
    }
    return $true
}

# Function to authenticate with web-based MFA
function Connect-Tenant {
    try {
        Write-Host "Connecting to Microsoft Graph..."
        Connect-MgGraph -Scopes "Directory.Read.All", "Sites.Read.All", "Sites.FullControl.All" -UseDeviceAuthentication
        Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
        
        Write-Host "Connecting to SharePoint Online..."
        Connect-PnPOnline -Interactive
        Write-Host "Connected to SharePoint Online." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to connect: $_" -ForegroundColor Red
        return $false
    }
    return $true
}

# Function to retrieve all SharePoint sites
function Get-AllSharePointSites {
    Write-Host "Retrieving all SharePoint sites..."
    try {
        $sites = Get-PnPTenantSite -ErrorAction Stop
        Write-Host "Successfully retrieved SharePoint sites." -ForegroundColor Green
        return $sites
    }
    catch {
        Write-Host "Error retrieving SharePoint sites: $_" -ForegroundColor Red
        return $null
    }
}

# Function to get permissions for a specific site/item
function Get-ItemPermissions {
    param(
        [string]$siteUrl,
        [string]$itemPath,
        [string]$itemType
    )
    try {
        $permissions = Get-PnPPermissions -List $itemPath -ErrorAction SilentlyContinue
        if (-not $permissions) {
            # If not a list/library, try getting site permissions
            $permissions = Get-PnPPermissions -ErrorAction SilentlyContinue
        }
        return $permissions
    }
    catch {
        Write-Host "Error getting permissions for $itemPath : $_" -ForegroundColor Yellow
        return $null
    }
}

# Function to get all items (sites, lists, folders, files) with external access
function Get-ExternalAccessItems {
    param(
        [string]$siteUrl
    )
    Write-Host "Processing site: $siteUrl"
    $results = @()
    
    try {
        # Connect to the specific site
        Connect-PnPOnline -Url $siteUrl -Interactive -ErrorAction Stop

        # Get site-level permissions
        $sitePermissions = Get-PnPPermissions -ErrorAction SilentlyContinue
        foreach ($perm in $sitePermissions) {
            if ($perm.IsExternal) {
                $results += [PSCustomObject]@{
                    'Resource Path' = $siteUrl
                    'Item Type' = 'Site'
                    'Permission' = $perm.PermissionLevels -join '; '
                    'User Name' = $perm.Principal.Title
                    'User Email' = $perm.Principal.Email
                }
            }
        }

        # Get all lists and libraries
        $lists = Get-PnPList -ErrorAction SilentlyContinue
        foreach ($list in $lists) {
            $listPermissions = Get-PnPPermissions -List $list.Title -ErrorAction SilentlyContinue
            
            foreach ($perm in $listPermissions) {
                if ($perm.IsExternal) {
                    $results += [PSCustomObject]@{
                        'Resource Path' = "$siteUrl/$($list.RootFolder.ServerRelativeUrl)"
                        'Item Type' = if ($list.BaseTemplate -eq 101) {'Document Library'} else {'List'}
                        'Permission' = $perm.PermissionLevels -join '; '
                        'User Name' = $perm.Principal.Title
                        'User Email' = $perm.Principal.Email
                    }
                }
            }

            # If it's a document library, check folders and files
            if ($list.BaseTemplate -eq 101) {
                $items = Get-PnPListItem -List $list.Title -PageSize 500
                foreach ($item in $items) {
                    $itemPermissions = Get-PnPPermissions -List $list.Title -ItemId $item.Id -ErrorAction SilentlyContinue
                    
                    foreach ($perm in $itemPermissions) {
                        if ($perm.IsExternal) {
                            $itemType = if ($item.FileSystemObjectType -eq 'Folder') {'Folder'} else {'File'}
                            $results += [PSCustomObject]@{
                                'Resource Path' = "$siteUrl/$($item.FieldValues.FileRef)"
                                'Item Type' = $itemType
                                'Permission' = $perm.PermissionLevels -join '; '
                                'User Name' = $perm.Principal.Title
                                'User Email' = $perm.Principal.Email
                            }
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-Host "Error processing site $siteUrl : $_" -ForegroundColor Red
    }
    
    return $results
}

# Function to save results to Excel
function Save-ResultsToExcel {
    param(
        [array]$results
    )
    if (-not $results -or $results.Count -eq 0) {
        Write-Host "No results to save." -ForegroundColor Yellow
        return
    }

    $fileName = "SharePoint_External_Access_Report_$((Get-Date).ToString('yyyy-MM-dd_HHmm')).xlsx"
    
    do {
        $filePath = Read-Host "Enter the path to save the Excel file (or press Enter for current directory)"
        if ([string]::IsNullOrWhiteSpace($filePath)) {
            $filePath = Get-Location
        }
    } while (-not (Test-Path $filePath -IsValid))

    $fullPath = Join-Path -Path $filePath -ChildPath $fileName

    try {
        Write-Host "Saving results to Excel..."
        $results | Export-Excel -Path $fullPath -WorksheetName "External Access" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow -ErrorAction Stop
        Write-Host "Results successfully saved to $fullPath" -ForegroundColor Green
    }
    catch {
        Write-Host "Error saving to Excel: $_" -ForegroundColor Red
    }
}

# Main Script Execution
function Execute-Script {
    Clear-Host
    Write-Host "SharePoint External Access Report" -ForegroundColor Cyan
    Write-Host "================================" -ForegroundColor Cyan

    # Initialize modules
    if (-not (Initialize-RequiredModules)) {
        Write-Host "Failed to initialize required modules. Exiting script." -ForegroundColor Red
        exit
    }

    # Connect to tenant
    if (-not (Connect-Tenant)) {
        Write-Host "Failed to connect to Microsoft Graph or SharePoint Online. Exiting script." -ForegroundColor Red
        exit
    }

    # Get all SharePoint sites
    $sites = Get-AllSharePointSites
    if (-not $sites) {
        Write-Host "No SharePoint sites found. Exiting script." -ForegroundColor Red
        exit
    }

    # Collect external user access data
    $allResults = @()
    $siteCount = $sites.Count
    $counter = 0
    foreach ($site in $sites) {
        $counter++
        $results = Get-ExternalAccessItems -siteUrl $site.Url
        $allResults += $results
        Write-Host "Processing site $counter of ${siteCount}: $($site.Url)"
    }

    # Display results
    if ($allResults.Count -gt 0) {
        $allResults | Format-Table -AutoSize
        $saveToExcel = Read-Host "Do you want to save the results to Excel? (y/n)"
        if ($saveToExcel -eq 'y') {
            Save-ResultsToExcel -results $allResults
        }
    }
    else {
        Write-Host "No external user access found." -ForegroundColor Yellow
    }
}

# Execute the script
Execute-Script
