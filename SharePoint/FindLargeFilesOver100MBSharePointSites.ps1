# MIT License
# Date Created: 05/09/2024
# Author: Tim Maclatchy
# Script Summary: This script connects to SharePoint Online, retrieves the largest files over 100MB from document libraries, and exports the results to an Excel file with separate worksheets for each site. Each worksheet includes a pie chart showing total site size and a table of the largest libraries.

# Config Variables
$TenantAdminURL = "https://tenant-admin.sharepoint.com/"
$ExcelFilePath = "C:\Temp\LargeFiles.xlsx"

# Import the ImportExcel module
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}
Import-Module ImportExcel

# Connect to Admin Center using PnP Online
try {
    Connect-PnPOnline -Url $TenantAdminURL -Interactive
    Write-Host "Connected to SharePoint Online Tenant Admin Center at $TenantAdminURL" -ForegroundColor Cyan
} catch {
    Write-Host "Error connecting to SharePoint Online: $_" -ForegroundColor Red
    exit
}

# Prepare an array to hold all file data
$AllFileData = @()

# Get All Site collections - Exclude certain sites
$SiteCollections = Get-PnPTenantSite | Where-Object {
    $_.URL -like '*/sites*' -and $_.Template -NotIn ("SRCHCEN#0", "REDIRECTSITE#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1")
}

# Exclude certain lists from the search
$ExcludedLists = @("Form Templates", "Preservation Hold Library", "Site Assets", "Pages", "Site Pages", "Images",
                    "Site Collection Documents", "Site Collection Images", "Style Library")

$SiteCounter = 1
# Loop through each site collection
ForEach ($Site in $SiteCollections) {
    # Display a Progress bar
    Write-Progress -Id 1 -Activity "Processing Site Collections" -Status "Processing Site: $($Site.URL) ($SiteCounter of $($SiteCollections.Count))" -PercentComplete (($SiteCounter / $SiteCollections.Count) * 100)
    
    # Connect to the site
    try {
        Connect-PnPOnline -Url $Site.URL -Interactive
    } catch {
        Write-Host "Error connecting to site $($Site.URL): $_" -ForegroundColor Red
        continue
    }

    # Get all document libraries
    $DocumentLibraries = Get-PnPList | Where-Object {
        $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $False -and $_.Title -notin $ExcludedLists -and $_.ItemCount -gt 0
    }

    $ListCounter = 1
    # Iterate through document libraries
    ForEach ($List in $DocumentLibraries) {
        $FileData = @()

        # Display progress for document libraries
        Write-Progress -Id 2 -ParentId 1 -Activity "Processing Document Libraries" -Status "Processing Document Library: $($List.Title) ($ListCounter of $($DocumentLibraries.Count))" -PercentComplete (($ListCounter / $DocumentLibraries.Count) * 100)

        # Get all files in the library larger than 100MB
        try {
            $TotalItems = $List.ItemCount
            $Counter = 0

            $Files = Get-PnPListItem -List $List -Fields FileLeafRef, FileRef, SMTotalFileStreamSize -PageSize 500 -ScriptBlock {
                Param($items)
                $global:Counter += $items.Count
                $PercentComplete = [math]::Min([math]::Round(($global:Counter / $TotalItems) * 100, 2), 100)
                Write-Progress -Id 3 -ParentId 2 -PercentComplete $PercentComplete -Activity "Getting List Items of '$($List.Title)'" -Status "Processing Items $global:Counter to $TotalItems"
            } | Where-Object { ($_.FileSystemObjectType -eq "File") -and ($_.FieldValues.SMTotalFileStreamSize / 1MB -gt 100) }
        } catch {
            Write-Host "Error retrieving items from list $($List.Title): $_" -ForegroundColor Red
            continue
        }

        # Collect data from each file
        ForEach ($File in $Files) {
            $FileData += [PSCustomObject][ordered]@{
                Site       = $Site.URL
                Library    = $List.Title
                FileName   = $File.FieldValues.FileLeafRef
                URL        = $File.FieldValues.FileRef
                Size_MB    = [math]::Round(($File.FieldValues.SMTotalFileStreamSize / 1MB), 2)
            }
        }

        # Add collected file data to the global array
        $AllFileData += $FileData
        $ListCounter++
    }
    $SiteCounter++
}

# Export data to Excel file
if (Test-Path $ExcelFilePath) { Remove-Item $ExcelFilePath }

# Group data by Site
$GroupedData = $AllFileData | Group-Object -Property Site

# Create a new Excel package
$ExcelPackage = Open-ExcelPackage -Path $ExcelFilePath

# Add a worksheet for each site
foreach ($Group in $GroupedData) {
    $SiteURL = $Group.Name
    $SiteData = $Group.Group

    # Add a new worksheet for the site
    $Worksheet = $ExcelPackage.Workbook.Worksheets.Add($SiteURL)

    # Group data by Library and sort by overall size (largest to smallest)
    $GroupedLibraries = $SiteData | Group-Object Library | Sort-Object { ($_.Group | Measure-Object -Property Size_MB -Sum).Sum } -Descending

    $CurrentRow = 1
    foreach ($LibraryGroup in $GroupedLibraries) {
        $LibraryName = $LibraryGroup.Name
        $LibrarySize_MB = ($LibraryGroup.Group | Measure-Object -Property Size_MB -Sum).Sum

        # Add a header row for each library
        $Worksheet.Cells[$CurrentRow, 1].Value = "$LibraryName (Total Size: $([math]::Round($LibrarySize_MB, 2)) MB)"
        $Worksheet.Cells[$CurrentRow, 1].Style.Font.Bold = $true
        $CurrentRow++

        # Sort files by size within the library
        $SortedFiles = $LibraryGroup.Group | Sort-Object Size_MB -Descending
        $Worksheet.Cells[$CurrentRow, 1].LoadFromCollection($SortedFiles, $true)
        $CurrentRow += $SortedFiles.Count

        # Add a blank row for separation
        $CurrentRow++
    }

    # Create a pie chart for total site size
    $Chart = $Worksheet.Drawings.AddChart("SiteSizeChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
    $Chart.Title.Text = "Total Site Size for $SiteURL"
    $Chart.SetPosition(1, 0, 0, 0)
    $Chart.SetSize(600, 400)
    $Chart.Series.Add($Worksheet.Cells["E2:E$($SiteData.Count + 1)"], $Worksheet.Cells["A2:A$($SiteData.Count + 1)"])
    $Chart.Series[0].Header = "Size_MB"
    $Chart.Series[0].XSeries = $Worksheet.Cells["Library"]
    $Chart.Series[0].YSeries = $Worksheet.Cells["Size_MB"]

    # Adjust column widths for better readability
    $Worksheet.Cells.AutoFitColumns()
}

# Save the Excel package
Close-ExcelPackage $ExcelPackage

Write-Host "Export completed successfully. Excel saved at: $ExcelFilePath" -ForegroundColor Green
