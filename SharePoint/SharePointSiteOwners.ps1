<#
Author: Tim MacLatchy
Date: 2024-09-12
License: MIT License
Description: This script lists all SharePoint Online sites and the owners of each site.
Modules: PnP.PowerShell, missing modules are installed automatically.
#>

# Check and install missing modules
$modules = @("PnP.PowerShell", "ImportExcel")
foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Module $module is missing. Installing..."
        Install-Module -Name $module -Force -AllowClobber
    }
}
Import-Module PnP.PowerShell

# Prompt for SharePoint Online admin URL
$adminUrl = Read-Host "Enter the SharePoint Online Admin URL (e.g., https://tenant-admin.sharepoint.com)"

# Validate URL format
if ($adminUrl -notmatch "^https://.*\.sharepoint\.com") {
    Write-Host "The URL provided is not valid. Please provide a valid SharePoint Online admin URL." -ForegroundColor Red
    Exit
}

# Use MFA login for SharePoint Online
try {
    Connect-PnPOnline -Url $adminUrl -UseWebLogin -ErrorAction Stop
    Write-Host "Successfully connected to SharePoint Online." -ForegroundColor Green
} catch {
    Write-Host "Error connecting to SharePoint Online: $_" -ForegroundColor Red
    Exit
}

# Graceful error handling
try {
    # Get all SharePoint sites
    $sites = Get-PnPTenantSite -Detailed -ErrorAction Stop
    Write-Host "Successfully retrieved all SharePoint sites." -ForegroundColor Green
} catch {
    Write-Host "Error retrieving SharePoint sites: $_" -ForegroundColor Red
    Exit
}

# Create an array to store site and owner information
$siteOwners = @()

# Progress bar for long-running task
$progress = 0
$totalSites = $sites.Count

foreach ($site in $sites) {
    $progress++
    Write-Progress -Activity "Fetching site owners" -Status "$progress out of $totalSites" -PercentComplete (($progress / $totalSites) * 100)

    try {
        # Get owners of the current site
        $owners = Get-PnPUser -Web $site.Url | Where-Object { $_.IsSiteAdmin -eq $true }
        foreach ($owner in $owners) {
            $siteOwners += [PSCustomObject]@{
                SiteUrl   = $site.Url
                SiteTitle = $site.Title
                Owner     = $owner.LoginName
            }
        }
    } catch {
        Write-Host "Error retrieving owners for site $($site.Url): $_" -ForegroundColor Red
    }
}

# Display results in the console
$siteOwners | Format-Table -Property SiteUrl, SiteTitle, Owner -AutoSize

# Prompt to export results to Excel or CSV
$exportChoice = Read-Host "Would you like to export the results to Excel or CSV? (Enter 'Excel', 'CSV', or 'No')"
if ($exportChoice -eq 'Excel' -or $exportChoice -eq 'CSV') {
    # Use a file dialog for saving the file
    Add-Type -AssemblyName System.Windows.Forms
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')

    # Generate filename based on tenant name, date, and operation
    $tenantName = $adminUrl.Split('/')[2].Split('.')[0]
    $date = (Get-Date).ToString("yyyy-MM-dd")
    $fileName = "${tenantName}_${date}_SiteOwners"

    if ($exportChoice -eq 'Excel') {
        $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $saveDialog.FileName = "${fileName}.xlsx"
    } else {
        $saveDialog.Filter = "CSV Files (*.csv)|*.csv"
        $saveDialog.FileName = "${fileName}.csv"
    }

    $saveDialog.Title = "Select file save location"
    $result = $saveDialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $filePath = $saveDialog.FileName

        if ($exportChoice -eq 'Excel') {
            # Export to Excel
            $siteOwners | Export-Excel -Path $filePath -AutoSize
            Write-Host "Results successfully exported to Excel at $filePath" -ForegroundColor Green
        } elseif ($exportChoice -eq 'CSV') {
            # Export to CSV
            $siteOwners | Export-Csv -Path $filePath -NoTypeInformation
            Write-Host "Results successfully exported to CSV at $filePath" -ForegroundColor Green
        }
    } else {
        Write-Host "No file was selected, export cancelled." -ForegroundColor Yellow
    }
} else {
    Write-Host "No file exported." -ForegroundColor Yellow
}

# Clean up
Disconnect-PnPOnline
