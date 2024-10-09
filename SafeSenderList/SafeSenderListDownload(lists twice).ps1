<#
    .SYNOPSIS
        Script to retrieve a user's safe sender list and blocked sender list from Exchange Online and save it to an Excel file, sorted and formatted.
    .DESCRIPTION
        This script connects to Exchange Online using browser-based MFA login, retrieves the safe sender list and blocked sender list for a specified user, 
        and exports the data to an Excel file. The lists are sorted and formatted with headers in bold and larger font, and columns auto-width.
    .NOTES
        Author: Tim MacLatchy
        Date: 2024-09-17
        License: MIT License
        Module Check: The script checks for and installs any missing modules automatically at the start.
#>

# Ensure required modules are installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
}
Import-Module ExchangeOnlineManagement

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Force -Scope CurrentUser
}
Import-Module ImportExcel

# Connect to Exchange Online with browser-based MFA login
Write-Host "Connecting to Exchange Online with browser-based MFA..."
try {
    Connect-ExchangeOnline -ShowProgress:$true -ErrorAction Stop
    Write-Host "Successfully connected to Exchange Online."
} catch {
    Write-Host "Error connecting to Exchange Online: $_"
    exit
}

# Prompt for user email to search for safe sender list and blocked sender list
$UserToSearch = Read-Host "Enter the email address of the user to retrieve the lists"

# Retrieve the safe sender and blocked sender lists for the specified user
Write-Host "Retrieving safe sender list for user $UserToSearch..."
try {
    $SafeSenders = Get-MailboxJunkEmailConfiguration -Identity $UserToSearch -ErrorAction Stop
    $AllowedSenders = $SafeSenders.TrustedSendersAndDomains | Where-Object { $_ -ne "" }
    $BlockedSenders = $SafeSenders.BlockedSendersAndDomains | Where-Object { $_ -ne "" }

    # Prepare data for export
    $ExportData = $AllowedSenders | ForEach-Object {
        $domain = $_.Split('@')[-1]
        [PSCustomObject]@{
            EmailAddress = $_
            Domain       = $domain
        }
    }

    # Count occurrences of each domain
    $DomainCounts = $ExportData | Group-Object -Property Domain | Sort-Object Count -Descending
    $CommonDomains = $DomainCounts | Where-Object { $_.Count -gt 1 }

    # Prepare data for each column
    $ColumnAData = $ExportData | Where-Object { $CommonDomains.Name -contains $_.Domain } | Sort-Object Domain, EmailAddress
    $OtherDomainData = $ExportData | Where-Object { $CommonDomains.Name -notcontains $_.Domain } | Sort-Object EmailAddress
    $BlockedSendersData = $BlockedSenders | ForEach-Object { [PSCustomObject]@{ EmailAddress = $_ } } | Sort-Object EmailAddress

    # Prepare data for allowed and blocked domains
    $AllowedDomains = $SafeSenders.TrustedSendersAndDomains | Where-Object { $_ -ne "" } | ForEach-Object { $_.Split('@')[-1] } | Sort-Object -Unique
    $BlockedDomains = $SafeSenders.BlockedSendersAndDomains | Where-Object { $_ -ne "" } | ForEach-Object { $_.Split('@')[-1] } | Sort-Object -Unique

    # Combine data for export
    $CombinedData = @()
    foreach ($address in $ColumnAData) {
        $CombinedData += [PSCustomObject]@{ ColumnA = $address.EmailAddress; ColumnB = ""; ColumnC = ""; ColumnD = ""; ColumnE = "" }
    }
    foreach ($address in $OtherDomainData) {
        $CombinedData += [PSCustomObject]@{ ColumnA = ""; ColumnB = $address.EmailAddress; ColumnC = ""; ColumnD = ""; ColumnE = "" }
    }
    foreach ($address in $BlockedSendersData) {
        $CombinedData += [PSCustomObject]@{ ColumnA = ""; ColumnB = ""; ColumnC = $address.EmailAddress; ColumnD = ""; ColumnE = "" }
    }
    foreach ($domain in $AllowedDomains) {
        $CombinedData += [PSCustomObject]@{ ColumnA = ""; ColumnB = ""; ColumnC = ""; ColumnD = $domain; ColumnE = "" }
    }
    foreach ($domain in $BlockedDomains) {
        $CombinedData += [PSCustomObject]@{ ColumnA = ""; ColumnB = ""; ColumnC = ""; ColumnD = ""; ColumnE = $domain }
    }

    # Display results in console
    Write-Host "Safe sender and blocked sender lists retrieved and sorted successfully."

    # Count entries in each column
    $ColumnAEntries = ($CombinedData | Where-Object { $_.ColumnA -ne "" }).Count
    $ColumnBEntries = ($CombinedData | Where-Object { $_.ColumnB -ne "" }).Count
    $ColumnCEntries = ($CombinedData | Where-Object { $_.ColumnC -ne "" }).Count
    $ColumnDEntries = ($CombinedData | Where-Object { $_.ColumnD -ne "" }).Count
    $ColumnEEntries = ($CombinedData | Where-Object { $_.ColumnE -ne "" }).Count

    # Output the number of entries in each column to console
    Write-Host "Number of entries in each column:"
    Write-Host "Column A: $ColumnAEntries"
    Write-Host "Column B: $ColumnBEntries"
    Write-Host "Column C: $ColumnCEntries"
    Write-Host "Column D: $ColumnDEntries"
    Write-Host "Column E: $ColumnEEntries"

    # Prompt user to save results (Excel)
    Add-Type -AssemblyName System.Windows.Forms
    $saveDialog = New-Object -TypeName System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    $saveDialog.DefaultExt = "xlsx"
    $saveDialog.FileName = "${UserToSearch}_$(Get-Date -Format yyyyMMdd)_Lists"

    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedPath = $saveDialog.FileName

        # Export to Excel
        $CombinedData | Export-Excel -Path $selectedPath -WorksheetName 'Lists' -AutoSize -TableName 'Lists' -ClearSheet

        # Open the Excel package to apply formatting
        $excelPackage = Open-ExcelPackage -Path $selectedPath
        $worksheet = $excelPackage.Workbook.Worksheets['Lists']

        # Clear all existing styles and formatting
        $worksheet.Cells.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::None
        $worksheet.Cells.Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::None)
        $worksheet.Cells.Style.Font.Color.SetColor([System.Drawing.Color]::Black)
        $worksheet.Cells.Style.Font.Size = 11
        $worksheet.Cells.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
        $worksheet.Cells.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Bottom

        # Set main headers
        $worksheet.Cells["A1"].Value = "Safe Sender List addresses with common domains"
        $worksheet.Cells["B1"].Value = "Safe Sender addresses with no common domains"
        $worksheet.Cells["C1"].Value = "Blocked Sender addresses"
        $worksheet.Cells["D1"].Value = "Allowed Domains"
        $worksheet.Cells["E1"].Value = "Blocked Domains"

        # Make headers bold and larger, with word wrap enabled
        $worksheet.Cells["A1:E1"].Style.Font.Bold = $true
        $worksheet.Cells["A1:E1"].Style.Font.Size = 16
        $worksheet.Cells["A1:E1"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        $worksheet.Cells["A1:E1"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
        $worksheet.Cells["A1:E1"].Style.WrapText = $true

        # Write Column A data with domain headers and color coding
        $rowIndex = 2
        foreach ($domainGroup in $CommonDomains) {
            $domain = $domainGroup.Name
            $worksheet.Cells["A$rowIndex"].Value = $domain
            $worksheet.Cells["A$rowIndex"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $worksheet.Cells["A$rowIndex"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGreen)
            $rowIndex++

            foreach ($address in $ExportData | Where-Object { $_.Domain -eq $domain }) {
                $worksheet.Cells["A$rowIndex"].Value = $address.EmailAddress
                $rowIndex++
            }
        }

        # Write other domain and blocked sender data
        $rowIndex = 2  # Reset row index for other columns
        for ($col = 1; $col -le 5; $col++) {
            $usedRange = $worksheet.Cells[1, $col, $rowIndex, $col]
            $uniqueValues = @($usedRange | Where-Object { $_.Value -ne "" } | Select-Object -Unique)
            $worksheet.Cells[1, $col, $rowIndex, $col].Clear()  # Clear existing cells
            
            for ($i = 0; $i -lt $uniqueValues.Count; $i++) {
                $worksheet.Cells[$i + 2, $col].Value = $uniqueValues[$i].Value  # Adjusted index to start from row 2
            }
        }

        # Save and close the Excel package
        Close-ExcelPackage $excelPackage

        Write-Host "Data successfully exported to $selectedPath"
    } else {
        Write-Host "File save dialog was canceled."
    }

    # Output totals for console
    Write-Host "Total Trusted Senders and Domains (addresses): $($AllowedSenders.Count)"
    Write-Host "Total Trusted Senders and Domains (domains): $($AllowedDomains.Count)"
    Write-Host "Total Blocked Senders and Domains (addresses): $($BlockedSenders.Count)"
    Write-Host "Total Blocked Senders and Domains (domains): $($BlockedDomains.Count)"

} catch {
    Write-Host "Error retrieving sender lists: $_"
} finally {
    Disconnect-ExchangeOnline -Confirm:$false
}
