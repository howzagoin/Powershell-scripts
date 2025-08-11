<#
Script to scan and fix broken links in Excel and Access files from:
  - Local or SharePoint synced and available offline folder(fastest)
  - SharePoint synced folder (fast, if available)
  - Online SharePoint URL (slowest; requires download and extra modules)

NOTE:
  Scanning a local folder or SharePoint synced and available offline folder is much faster than scanning an online SharePoint URL. Online scanning may take significantly longer due to network latency, file download times, and SharePoint API throttling. For online SharePoint, files are downloaded to a temporary folder for processing.

Author: Tim MacLatchy
Date: 2025-08-01
#>

#region Configuration
Write-Host "Select scan location type:" -ForegroundColor Cyan
Write-Host "1. Local or SharePoint synced and available offline folder(fastest)" -ForegroundColor Yellow
Write-Host "2. SharePoint synced folder (fast, if available)" -ForegroundColor Yellow
Write-Host "3. Online SharePoint folder URL (slowest)" -ForegroundColor Yellow
$locationType = Read-Host "Enter 1, 2, or 3"

if ($locationType -eq '1' -or $locationType -eq '2') {
    $path = Read-Host "Enter the full path to the folder to scan"
    if (-not (Test-Path $path)) {
        Write-Host "Path does not exist. Exiting." -ForegroundColor Red
        exit
    }
    $excelFiles = Get-ChildItem -Path $path -Include *.xls, *.xlsx -Recurse
    $accessFiles = Get-ChildItem -Path $path -Include *.mdb, *.accdb -Recurse
} elseif ($locationType -eq '3') {
    # Online SharePoint scan using PnP.PowerShell
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Write-Host "PnP.PowerShell module not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name PnP.PowerShell -Force -Scope CurrentUser
    }
    Import-Module PnP.PowerShell
    $url = Read-Host "Enter the online SharePoint folder URL to scan (e.g., https://yourtenant.sharepoint.com/sites/yoursite/Shared%20Documents)"
    Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan
    try {
        Connect-PnPOnline -Url $url -Interactive
    } catch {
        Write-Host "Failed to connect to SharePoint: $_" -ForegroundColor Red
        exit
    }
    $tempFolder = Join-Path $env:TEMP "SPScan_$(Get-Date -Format 'yyyyMMddHHmmss')"
    New-Item -ItemType Directory -Path $tempFolder | Out-Null
    Write-Host "Downloading files from SharePoint to $tempFolder..." -ForegroundColor Cyan
    try {
        $spFiles = Get-PnPFolderItem -FolderSiteRelativeUrl ([Uri]::UnescapeDataString((($url -split '/sites/')[1]) -replace '^.+?/', '')) -ItemType File
    } catch {
        Write-Host "Failed to get SharePoint files: $_" -ForegroundColor Red
        exit
    }
    $excelFiles = @()
    $accessFiles = @()
    foreach ($spFile in $spFiles) {
        $fileName = $spFile.Name
        $localPath = Join-Path $tempFolder $fileName
        try {
            Get-PnPFile -Url $spFile.ServerRelativeUrl -Path $tempFolder -FileName $fileName -AsFile -Force
            if ($fileName -match '\.xlsx?$') {
                $excelFiles += Get-Item $localPath
            } elseif ($fileName -match '\.accdb$|\.mdb$') {
                $accessFiles += Get-Item $localPath
            }
        } catch {
            Write-Host "Failed to download file $fileName`: $_" -ForegroundColor Yellow
        }
    }
    $path = $tempFolder
} else {
    Write-Host "Invalid selection. Exiting." -ForegroundColor Red
    exit
}

# Initialize Excel only if we have Excel files
$excel = $null
if ($excelFiles.Count -gt 0) {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
    } catch {
        Write-Host "Failed to initialize Excel COM object: $_" -ForegroundColor Red
        $excel = $null
    }
}
#endregion

#region Results Collection
# Collect results for export
$results = @()
#endregion

#region Excel Scan & Fix
# Scan and fix Excel files (parallel for local/synced folders)
if ($excelFiles.Count -gt 0) {
    if ($locationType -eq '1' -or $locationType -eq '2') {
        Write-Host "Scanning and fixing Excel files..." -ForegroundColor Cyan
        $excelCount = $excelFiles.Count
        $excelResults = @()
        $excelJobs = @()
        $maxJobs = 4
        for ($i = 0; $i -lt $excelCount; $i++) {
            $file = $excelFiles[$i]
            $excelJobs += Start-Job -ScriptBlock {
                param($filePath, $excelFiles)
                $localResults = @()
                $excel = $null
                $wb = $null
                try {
                    $excel = New-Object -ComObject Excel.Application
                    $excel.Visible = $false
                    $excel.DisplayAlerts = $false
                    $excel.ScreenUpdating = $false
                    $wb = $excel.Workbooks.Open($filePath, 0, $true, $false, '', '', $false, $false, 1, $false, $false, $false, $false, $false, $false)
                    $links = $wb.LinkSources(1)
                    if ($links) {
                        foreach ($link in $links) {
                            $status = $wb.LinkInfo($link, 1)
                            $fixed = $false
                            $newLink = $null
                            if ($status -ne 0) {
                                $linkFileName = [System.IO.Path]::GetFileName($link)
                                $candidate = $excelFiles | Where-Object { $_.Name -eq $linkFileName }
                                if ($candidate.Count -gt 0) {
                                    $newLink = $candidate[0].FullName
                                    try {
                                        $wb.ChangeLink($link, $newLink, 1)
                                        $fixed = $true
                                        $status = $wb.LinkInfo($newLink, 1)
                                    } catch {
                                        $fixed = $false
                                    }
                                }
                            }
                            $localResults += [PSCustomObject]@{
                                File   = $filePath
                                Link   = $link
                                Status = $status
                                Fixed  = $fixed
                                NewLink = $newLink
                                Type   = 'Excel'
                            }
                        }
                    }
                } catch {
                    $localResults += [PSCustomObject]@{
                        File   = $filePath
                        Link   = 'Error opening workbook'
                        Status = 'Error'
                        Fixed  = $false
                        NewLink = $null
                        Type   = 'Excel'
                    }
                } finally {
                    # Critical: Proper cleanup in finally block
                    if ($wb) {
                        try {
                            $wb.Close($false)
                            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
                        } catch { }
                        $wb = $null
                    }
                    if ($excel) {
                        try {
                            $excel.Workbooks.Close()
                            $excel.Quit()
                            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
                        } catch { }
                        $excel = $null
                    }
                    # Force garbage collection in job
                    [GC]::Collect()
                    [GC]::WaitForPendingFinalizers()
                    [GC]::Collect()
                }
                return $localResults
            } -ArgumentList $file.FullName, $excelFiles
            
            # Wait for jobs to complete before starting more
            while ($excelJobs.Count -ge $maxJobs) {
                $finished = $excelJobs | Where-Object { $_.State -eq 'Completed' }
                foreach ($job in $finished) {
                    $excelResults += Receive-Job $job
                    Remove-Job $job
                }
                $excelJobs = $excelJobs | Where-Object { $_.State -ne 'Completed' }
                Start-Sleep -Milliseconds 500
            }
            $percentComplete = [math]::Round((($i+1) / $excelCount) * 100, 1)
            Write-Progress -Activity "Excel Scan & Fix" -Status "Processing: $($file.Name)" -PercentComplete $percentComplete -CurrentOperation "$($i+1) of $excelCount files"
        }
        
        # Collect remaining jobs
        foreach ($job in $excelJobs) {
            Wait-Job $job | Out-Null
            $excelResults += Receive-Job $job
            Remove-Job $job
        }
        Write-Progress -Activity "Excel Scan & Fix" -Completed
        $results += $excelResults
    } else {
        # Sequential processing for online SharePoint
        Write-Host "Scanning and fixing Excel files..." -ForegroundColor Cyan
        foreach ($file in $excelFiles) {
            try {
                $wb = $excel.Workbooks.Open($file.FullName, 0, $true)
                $links = $wb.LinkSources(1)
                if ($links) {
                    foreach ($link in $links) {
                        $status = $wb.LinkInfo($link, 1)
                        $fixed = $false
                        $newLink = $null
                        if ($status -ne 0) {
                            $linkFileName = [System.IO.Path]::GetFileName($link)
                            $candidate = $excelFiles | Where-Object { $_.Name -eq $linkFileName }
                            if ($candidate.Count -gt 0) {
                                $newLink = $candidate[0].FullName
                                try {
                                    $wb.ChangeLink($link, $newLink, 1)
                                    $fixed = $true
                                    $status = $wb.LinkInfo($newLink, 1)
                                } catch {
                                    $fixed = $false
                                }
                            }
                        }
                        $results += [PSCustomObject]@{
                            File   = $file.FullName
                            Link   = $link
                            Status = $status
                            Fixed  = $fixed
                            NewLink = $newLink
                            Type   = 'Excel'
                        }
                    }
                }
                # Proper cleanup sequence
                $wb.Close($true)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
                $wb = $null
            } catch {
                $results += [PSCustomObject]@{
                    File   = $file.FullName
                    Link   = 'Error opening workbook'
                    Status = 'Error'
                    Fixed  = $false
                    NewLink = $null
                    Type   = 'Excel'
                }
            }
        }
    }
}
#endregion

#region Access Scan & Fix
# Scan and fix Access files (parallel for local/synced folders)
if ($accessFiles.Count -gt 0) {
    if ($locationType -eq '1' -or $locationType -eq '2') {
        Write-Host "Scanning and fixing Access files..." -ForegroundColor Cyan
        $accessCount = $accessFiles.Count
        $accessResults = @()
        $accessJobs = @()
        $maxJobs = 4
        for ($i = 0; $i -lt $accessCount; $i++) {
            $file = $accessFiles[$i]
            $accessJobs += Start-Job -ScriptBlock {
                param($filePath, $accessFiles)
                $localResults = @()
                try {
                    $accessApp = New-Object -ComObject Access.Application
                    $accessApp.Visible = $false
                    $accessApp.UserControl = $false
                    $accessApp.OpenCurrentDatabase($filePath)
                    $db = $accessApp.CurrentDb()
                    foreach ($tableDef in $db.TableDefs) {
                        if ($tableDef.Attributes -band 32) {
                            $link = $tableDef.Connect
                            $fixed = $false
                            $newLink = $null
                            $linkFileName = $null
                            if ($link -match "DATABASE=(.+?);") {
                                $linkFileName = [System.IO.Path]::GetFileName($matches[1])
                            }
                            if ($linkFileName) {
                                $candidate = $accessFiles | Where-Object { $_.Name -eq $linkFileName }
                                if ($candidate.Count -gt 0) {
                                    $newLink = $candidate[0].FullName
                                    try {
                                        $tableDef.Connect = "DATABASE=$newLink;"
                                        $tableDef.RefreshLink()
                                        $fixed = $true
                                    } catch {
                                        $fixed = $false
                                    }
                                }
                            }
                            $localResults += [PSCustomObject]@{
                                File   = $filePath
                                Link   = $link
                                Status = if ($fixed) { 'Fixed' } else { 'Broken' }
                                Fixed  = $fixed
                                NewLink = $newLink
                                Type   = 'Access'
                            }
                        }
                    }
                    $accessApp.Quit()
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($accessApp) | Out-Null
                } catch {
                    $localResults += [PSCustomObject]@{
                        File   = $filePath
                        Link   = 'Error opening database'
                        Status = 'Error'
                        Fixed  = $false
                        NewLink = $null
                        Type   = 'Access'
                    }
                }
                return $localResults
            } -ArgumentList $file.FullName, $accessFiles
            
            # Wait for jobs to complete before starting more
            while ($accessJobs.Count -ge $maxJobs) {
                $finished = $accessJobs | Where-Object { $_.State -eq 'Completed' }
                foreach ($job in $finished) {
                    $accessResults += Receive-Job $job
                    Remove-Job $job
                }
                $accessJobs = $accessJobs | Where-Object { $_.State -ne 'Completed' }
                Start-Sleep -Milliseconds 500
            }
            $percentComplete = [math]::Round((($i+1) / $accessCount) * 100, 1)
            Write-Progress -Activity "Access Scan & Fix" -Status "Processing: $($file.Name)" -PercentComplete $percentComplete -CurrentOperation "$($i+1) of $accessCount files"
        }
        
        # Collect remaining jobs
        foreach ($job in $accessJobs) {
            Wait-Job $job | Out-Null
            $accessResults += Receive-Job $job
            Remove-Job $job
        }
        Write-Progress -Activity "Access Scan & Fix" -Completed
        $results += $accessResults
    } else {
        # Sequential processing for online SharePoint
        Write-Host "Scanning and fixing Access files..." -ForegroundColor Cyan
        foreach ($file in $accessFiles) {
            try {
                $accessApp = New-Object -ComObject Access.Application
                $accessApp.Visible = $false
                $accessApp.UserControl = $false
                $accessApp.OpenCurrentDatabase($file.FullName)
                $db = $accessApp.CurrentDb()
                foreach ($tableDef in $db.TableDefs) {
                    if ($tableDef.Attributes -band 32) {
                        $link = $tableDef.Connect
                        $fixed = $false
                        $newLink = $null
                        $linkFileName = $null
                        if ($link -match "DATABASE=(.+?);") {
                            $linkFileName = [System.IO.Path]::GetFileName($matches[1])
                        }
                        if ($linkFileName) {
                            $candidate = $accessFiles | Where-Object { $_.Name -eq $linkFileName }
                            if ($candidate.Count -gt 0) {
                                $newLink = $candidate[0].FullName
                                try {
                                    $tableDef.Connect = "DATABASE=$newLink;"
                                    $tableDef.RefreshLink()
                                    $fixed = $true
                                } catch {
                                    $fixed = $false
                                }
                            }
                        }
                        $results += [PSCustomObject]@{
                            File   = $file.FullName
                            Link   = $link
                            Status = if ($fixed) { 'Fixed' } else { 'Broken' }
                            Fixed  = $fixed
                            NewLink = $newLink
                            Type   = 'Access'
                        }
                    }
                }
                $accessApp.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($accessApp) | Out-Null
            } catch {
                $results += [PSCustomObject]@{
                    File   = $file.FullName
                    Link   = 'Error opening database'
                    Status = 'Error'
                    Fixed  = $false
                    NewLink = $null
                    Type   = 'Access'
                }
            }
        }
    }
}
#endregion

#region Cleanup
# Clean up Excel COM object
if ($excel) {
    try {
        # Close all workbooks
        $excel.Workbooks.Close()
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        $excel = $null
    } catch {
        # Ignore cleanup errors
    }
}

# Force cleanup of any remaining Excel processes
try {
    Get-Process "excel" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq "" } | Stop-Process -Force
} catch {
    # Ignore if no processes found
}

# Force garbage collection
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
[GC]::Collect()

# Clean up temp folder if created
if ($locationType -eq '3' -and (Test-Path $tempFolder)) {
    try {
        Remove-Item -Path $tempFolder -Recurse -Force
        Write-Host "Cleaned up temporary folder: $tempFolder" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not clean up temporary folder: $tempFolder" -ForegroundColor Yellow
    }
}
#endregion

#region Display Results Summary
Write-Host "`nScan Results Summary:" -ForegroundColor Cyan
Write-Host "Total files processed: $($excelFiles.Count + $accessFiles.Count)" -ForegroundColor White
Write-Host "Total links found: $($results.Count)" -ForegroundColor White
if ($results.Count -gt 0) {
    $brokenLinks = $results | Where-Object { $_.Status -ne 0 -and $_.Status -ne 'Fixed' -and $_.Status -ne 'Error' }
    $fixedLinks = $results | Where-Object { $_.Fixed -eq $true }
    $errorLinks = $results | Where-Object { $_.Status -eq 'Error' }
    Write-Host "Broken links: $($brokenLinks.Count)" -ForegroundColor Red
    Write-Host "Fixed links: $($fixedLinks.Count)" -ForegroundColor Green
    Write-Host "Error processing: $($errorLinks.Count)" -ForegroundColor Yellow
}
#endregion

#region Export Report
# Check if ImportExcel module is available
$importExcelAvailable = $false
try {
    Import-Module ImportExcel -ErrorAction Stop
    $importExcelAvailable = $true
} catch {
    Write-Host "ImportExcel module not available. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name ImportExcel -Force -Scope CurrentUser
        Import-Module ImportExcel
        $importExcelAvailable = $true
    } catch {
        Write-Host "Failed to install ImportExcel module. Exporting as CSV instead." -ForegroundColor Yellow
        $importExcelAvailable = $false
    }
}

if ($results.Count -gt 0) {
    Add-Type -AssemblyName System.Windows.Forms
    if ($importExcelAvailable) {
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv"
        $saveDialog.FileName = "ExcelBrokenLinksReport_$(Get-Date -Format 'yyyyMMdd').xlsx"
        $saveDialog.Title = "Save Broken Links Report"
        $result = $saveDialog.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $filePath = $saveDialog.FileName
            if ($filePath.EndsWith('.xlsx')) {
                try {
                    $excelPkg = $results | Export-Excel -Path $filePath -WorksheetName "Broken Links" -AutoSize -TableStyle Light1 -PassThru
                    $ws = $excelPkg.Workbook.Worksheets["Broken Links"]
                    if ($ws) {
                        # High contrast formatting: header row dark blue, white text, bold
                        $headerRange = $ws.Cells[$ws.Dimension.Start.Row, $ws.Dimension.Start.Column, $ws.Dimension.Start.Row, $ws.Dimension.End.Column]
                        $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                        $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::DarkBlue)
                        $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                        $headerRange.Style.Font.Bold = $true
                        $headerRange.Style.Font.Size = 12
                        # Alternating row colors
                        $dataRows = $ws.Dimension.Rows
                        for ($row = 2; $row -le $dataRows; $row++) {
                            $rowRange = $ws.Cells[$row, $ws.Dimension.Start.Column, $row, $ws.Dimension.End.Column]
                            if ($row % 2 -eq 0) {
                                $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                                $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightSteelBlue)
                                $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::Black)
                            } else {
                                $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                                $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::White)
                                $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::Black)
                            }
                            $rowRange.Style.Font.Bold = $true
                            $rowRange.Style.Font.Size = 10
                        }
                    }
                    Close-ExcelPackage $excelPkg
                    Write-Host "Report exported to $filePath" -ForegroundColor Green
                } catch {
                    Write-Host "Failed to export Excel file: $_" -ForegroundColor Red
                    Write-Host "Exporting as CSV instead..." -ForegroundColor Yellow
                    $csvPath = $filePath -replace '\.xlsx$', '.csv'
                    $results | Export-Csv -Path $csvPath -NoTypeInformation
                    Write-Host "Report exported to $csvPath" -ForegroundColor Green
                }
            } else {
                $results | Export-Csv -Path $filePath -NoTypeInformation
                Write-Host "Report exported to $filePath" -ForegroundColor Green
            }
        } else {
            Write-Host "Export cancelled." -ForegroundColor Yellow
        }
    } else {
        # Fallback to CSV export
        $csvPath = Join-Path ([Environment]::GetFolderPath('Desktop')) "ExcelBrokenLinksReport_$(Get-Date -Format 'yyyyMMdd').csv"
        $results | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Host "Report exported to $csvPath" -ForegroundColor Green
    }
} else {
    Write-Host "No broken links found." -ForegroundColor Yellow
}
#endregion