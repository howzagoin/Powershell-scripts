# Set verbose preference for debugging output
$VerbosePreference = "Continue"

# Connect to SharePoint Online with MFA
function Connect-SharePointTenant {
    param (
        [string]$AdminUrl
    )
    try {
        Write-Verbose ("Connecting to SharePoint tenant at {0}" -f $AdminUrl)
        Connect-PnPOnline -Url $AdminUrl -Interactive
        Write-Verbose "Connection successful!"
        return $true
    } catch {
        Write-Error ("Failed to connect to SharePoint tenant: {0}" -f $_)
        return $false
    }
}

# Define function to retrieve all SharePoint sites
function Get-AllSites {
    try {
        Write-Verbose "Retrieving all SharePoint sites in the tenant"
        $sites = Get-PnPTenantSite -Detailed -WarningAction SilentlyContinue
        Write-Verbose ("Successfully retrieved {0} sites." -f $sites.Count)
        return $sites
    } catch {
        Write-Error ("Failed to retrieve SharePoint sites: {0}" -f $_)
        return $null
    }
}

# Define function to process each site
function Process-Site {
    param (
        [string]$SiteUrl
    )

    Write-Verbose ("Processing site: {0}" -f $SiteUrl)

    try {
        # Connect to individual site
        Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop

        Write-Verbose ("Connected to site: {0}" -f $SiteUrl)

        # Retrieve lists
        $lists = Get-PnPList
        Write-Verbose ("Successfully retrieved {0} lists from site: {1}" -f $lists.Count, $SiteUrl)

        # Retrieve largest 10 files
        $files = Get-PnPFolderItem -Folder "/Shared Documents" -ErrorAction Stop | Sort-Object Length -Descending | Select-Object -First 10
        Write-Verbose ("Top 10 files retrieved from {0}" -f $SiteUrl)
        
        # Retrieve largest 10 folders
        $folders = Get-PnPFolder -Folder "/Shared Documents" -ErrorAction Stop | Sort-Object Length -Descending | Select-Object -First 10
        Write-Verbose ("Top 10 folders retrieved from {0}" -f $SiteUrl)
        
        return [PSCustomObject]@{
            SiteUrl = $SiteUrl
            Files = $files
            Folders = $folders
        }
    } catch {
        Write-Error ("Failed to retrieve files or folders for site: {0}. Error: {1}" -f $SiteUrl, $_.Exception.Message)
    }
}

# Define function to export data to Excel
function Export-ToExcel {
    param (
        [string]$FilePath,
        [array]$Data
    )

    Write-Verbose ("Exporting data to Excel file: {0}" -f $FilePath)
    $Data | Export-Excel -Path $FilePath -AutoSize -TableName 'SharePointReport' -BoldTopRow
    Write-Verbose ("Data exported successfully to {0}" -f $FilePath)
}

# Define the Main function
function Main {
    Write-Verbose "Entering Main function"
    
    $AdminUrl = "https://compraraau-admin.sharepoint.com/"
    
    # Connect to SharePoint tenant
    if (-not (Connect-SharePointTenant -AdminUrl $AdminUrl)) {
        Write-Error "Could not connect to SharePoint tenant. Exiting script."
        return
    }

    # Retrieve all sites
    $sites = Get-AllSites
    if (-not $sites) {
        Write-Error "No sites retrieved. Exiting script."
        return
    }

    # Initialize progress bar
    $totalSites = $sites.Count
    $currentSite = 0
    $reportData = @()

    foreach ($site in $sites) {
        $currentSite++
        Write-Progress -PercentComplete (($currentSite / $totalSites) * 100) -Status "Processing SharePoint sites" -CurrentOperation ("Processing site " + $currentSite + " of " + $totalSites + ": " + $site.Url)
        
        Write-Verbose ("Processing site {0} of {1}: {2}" -f $currentSite, $totalSites, $site.Url)
        $siteData = Process-Site -SiteUrl $site.Url
        if ($siteData) {
            $reportData += $siteData
        }
    }

    # Export to Excel after processing
    $saveExcel = Read-Host "Would you like to save the results to an Excel file? (Y/N)"
    if ($saveExcel -eq 'Y') {
        $savePath = Read-Host "Enter the full path to save the Excel file (e.g., C:\Reports\SharePointReport.xlsx)"
        Export-ToExcel -FilePath $savePath -Data $reportData
    }

    Write-Verbose "Script execution completed!"
}

# Execute Main function
Write-Verbose "Starting script execution"
Main
# Set verbose preference for debugging output
$VerbosePreference = "Continue"

# Connect to SharePoint Online with MFA
function Connect-SharePointTenant {
    param (
        [string]$AdminUrl
    )
    try {
        Write-Verbose ("Connecting to SharePoint tenant at {0}" -f $AdminUrl)
        Connect-PnPOnline -Url $AdminUrl -Interactive
        Write-Verbose "Connection successful!"
        return $true
    } catch {
        Write-Error ("Failed to connect to SharePoint tenant: {0}" -f $_)
        return $false
    }
}

# Define function to retrieve all SharePoint sites
function Get-AllSites {
    try {
        Write-Verbose "Retrieving all SharePoint sites in the tenant"
        $sites = Get-PnPTenantSite -Detailed -WarningAction SilentlyContinue
        Write-Verbose ("Successfully retrieved {0} sites." -f $sites.Count)
        return $sites
    } catch {
        Write-Error ("Failed to retrieve SharePoint sites: {0}" -f $_)
        return $null
    }
}

# Define function to process each site
function Process-Site {
    param (
        [string]$SiteUrl
    )

    Write-Verbose ("Processing site: {0}" -f $SiteUrl)

    try {
        # Connect to individual site
        Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop

        Write-Verbose ("Connected to site: {0}" -f $SiteUrl)

        # Retrieve lists
        $lists = Get-PnPList
        Write-Verbose ("Successfully retrieved {0} lists from site: {1}" -f $lists.Count, $SiteUrl)

        # Retrieve largest 10 files
        $files = Get-PnPListItem -List "Documents" | Sort-Object -Property Length -Descending | Select-Object -First 10
        Write-Verbose ("Top 10 files retrieved from {0}" -f $SiteUrl)
        
        # Retrieve largest 10 folders
        $folders = Get-PnPListItem -List "Documents" -Folder "/" | Sort-Object -Property Length -Descending | Select-Object -First 10
        Write-Verbose ("Top 10 folders retrieved from {0}" -f $SiteUrl)
        
        return [PSCustomObject]@{
            SiteUrl = $SiteUrl
            Files = $files
            Folders = $folders
        }
    } catch {
        Write-Error ("Failed to retrieve files or folders for site: {0}. Error: {1}" -f $SiteUrl, $_.Exception.Message)
    }
}

# Define function to export data to Excel
function Export-ToExcel {
    param (
        [array]$Data
    )

    Add-Type -AssemblyName System.Windows.Forms
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    $saveFileDialog.Title = "Select location to save Excel file"
    
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        $FilePath = $saveFileDialog.FileName
        Write-Verbose ("Exporting data to Excel file: {0}" -f $FilePath)
        $Data | Export-Excel -Path $FilePath -AutoSize -TableName 'SharePointReport' -BoldTopRow
        Write-Verbose ("Data exported successfully to {0}" -f $FilePath)
    } else {
        Write-Verbose "Export to Excel cancelled by user."
    }
}

# Define the Main function
function Main {
    Write-Verbose "Entering Main function"
    
    $AdminUrl = "https://compraraau-admin.sharepoint.com/"
    
    # Connect to SharePoint tenant
    if (-not (Connect-SharePointTenant -AdminUrl $AdminUrl)) {
        Write-Error "Could not connect to SharePoint tenant. Exiting script."
        return
    }

    # Retrieve all sites
    $sites = Get-AllSites
    if (-not $sites) {
        Write-Error "No sites retrieved. Exiting script."
        return
    }

    # Initialize progress bar
    $totalSites = $sites.Count
    $currentSite = 0
    $reportData = @()

    foreach ($site in $sites) {
        $currentSite++
        Write-Progress -PercentComplete (($currentSite / $totalSites) * 100) -Status "Processing SharePoint sites" -CurrentOperation ("Processing site " + $currentSite + " of " + $totalSites + ": " + $site.Url)
        
        Write-Verbose ("Processing site {0} of {1}: {2}" -f $currentSite, $totalSites, $site.Url)
        $siteData = Process-Site -SiteUrl $site.Url
        if ($siteData) {
            $reportData += $siteData
        }
    }

    # Ask if user wants to save results to Excel
    $saveExcel = Read-Host "Would you like to save the results to an Excel file? (Y/N)"
    if ($saveExcel -eq 'Y') {
        Export-ToExcel -Data $reportData
    }

    Write-Verbose "Script execution completed!"
}

# Execute Main function
Write-Verbose "Starting script execution"
Main
