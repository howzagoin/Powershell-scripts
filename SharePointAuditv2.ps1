# ============================================================================
# SharePoint Audit Script
# Author      : Timothy MacLatchy
# Description : Comprehensive audit of SharePoint sites in a Microsoft 365 tenant
#               Certificate-based authentication only
# Prerequisite: PowerShell 7.1+, Microsoft Graph PowerShell SDK, ImportExcel module
# ============================================================================

# Remove all loaded Microsoft.Graph modules to prevent version conflicts
Get-Module -Name Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue

param (
  [string]$OutputPath,
  [switch]$TestMode,
  [int]$ParallelLimit = 4
)

if ([string]::IsNullOrWhiteSpace($ClientId) -or [string]::IsNullOrWhiteSpace($TenantId) -or [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
  Write-Host "ERROR: ClientId, TenantId, and CertificateThumbprint must not be empty. Please check your parameters." -ForegroundColor Red
  exit 1
 

<#
.DESCRIPTION
    This script performs a comprehensive audit of SharePoint sites in a Microsoft 365 tenant, including:
    - Scans all SharePoint sites and aggregates storage usage
    - Generates pie charts of storage for the whole tenant (largest 10 sites by size)
    - For each of the 10 largest sites, generates a pie chart showing storage breakdown
    - Collects user access for all sites, including user type (internal/external)
    - Highlights external guest access in red in the Excel report
    - Lists site owners and site members for each site
    - For the top 10 largest sites, shows the top 20 biggest files and folders
    - Exports all results to a well-structured Excel report with multiple worksheets and charts
    - Progress bars for site/library/file processing
    - Robust error handling and reporting
#>

## ============================================================================
## Logging, Progress, and Utility Functions
## ============================================================================

# ============================================================================
# Permission and Access Analysis
# ============================================================================
## ============================================================================
# Excel Report Generation
## ============================================================================

function Test-ExcelFile {
param (
    [Parameter(Mandatory=$true)]
    [string]$FilePath
)
        Import-Module ImportExcel -ErrorAction SilentlyContinue
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
    } catch {
        Write-Log "Could not show save dialog. Using default filename in current directory." -Level Warning
        return $DefaultFileName
    }
}

function Export-ExcelWorksheet {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Data,
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$WorksheetName,
        [string]$Title,
        [string]$TableStyle = "Light1",
        [hashtable]$ConditionalFormatting = @{},
        [switch]$AutoSize,
        [switch]$FreezeTopRow,
        [switch]$BoldTopRow,
        [string]$HeaderColor,
        [string]$HeaderTextColor,
        [array]$AlternateRowColors = @("LightGray", "White"),
        [switch]$PassThru
    )

    try {
        $params = @{
            Path = $Path
            WorksheetName = $WorksheetName
            AutoSize = $AutoSize
            PassThru = $PassThru
        }

        if ($Title) {
            $params.Add("Title", $Title)
            $params.Add("TitleBold", $true)
            $params.Add("TitleSize", 16)
        }

        if ($TableStyle) {
            $params.Add("TableStyle", $TableStyle)
        }

        $excel = $Data | Export-Excel @params

        # Apply conditional formatting if specified
        if ($ConditionalFormatting.Count -gt 0) {
            $ws = $excel.Workbook.Worksheets[$WorksheetName]
            if ($ws) {
                foreach ($cf in $ConditionalFormatting.GetEnumerator()) {
                    $range = $cf.Key
                    $format = $cf.Value

                    if ($format -is [hashtable]) {
                        $ruleType = $format.RuleType
                        $condition = $format.Condition
                        $color = $format.Color

                        if ($ruleType -eq "ContainsText" -and $condition -and $color) {
                            Add-ConditionalFormatting -Worksheet $ws -Range $range -RuleType ContainsText -Condition $condition -ForegroundColor $color
                        } elseif ($ruleType -eq "Expression" -and $condition -and $color) {
                            Add-ConditionalFormatting -Worksheet $ws -Range $range -RuleType Expression -Condition $condition -ForegroundColor $color
                        }
                    }
                }
            }
        }

        # Apply formatting options
        if ($HeaderColor -or $HeaderTextColor -or $BoldTopRow -or $AlternateRowColors -or $FreezeTopRow) {
            $ws = $excel.Workbook.Worksheets[$WorksheetName]
            if ($ws) {
                # Get the range of the header row
                $headerRow = 1
                $headerRange = $ws.Cells[$headerRow, 1, $headerRow, $ws.Dimension.End.Column]

                # Apply header formatting
                if ($HeaderColor) {
                    $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $headerRange.Style.Fill.BackgroundColor.SetColor($HeaderColor)
                }

                if ($HeaderTextColor) {
                    $headerRange.Style.Font.Color.SetColor($HeaderTextColor)
                }

                if ($BoldTopRow) {
                    $headerRange.Style.Font.Bold = $true
                    $headerRange.Style.Font.Size = 12
                }

                # Apply alternate row colors
                if ($AlternateRowColors.Count -ge 2) {
                    $dataRows = $ws.Dimension.Rows
                    for ($row = 2; $row -le $dataRows; $row++) {
                        $rowRange = $ws.Cells[$row, 1, $row, $ws.Dimension.End.Column]
                        $colorIndex = ($row - 2) % $AlternateRowColors.Count

                        try {
                            $color = $AlternateRowColors[$colorIndex]
                            if ($color -is [string]) {
                                $color = [System.Drawing.Color]::FromName($color)
                            }

                            $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $rowRange.Style.Fill.BackgroundColor.SetColor($color)

                            # Set contrasting text color
                            if ($color.GetBrightness() -lt 0.5) {
                                $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                            } else {
                                $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::Black)
                            }

                            $rowRange.Style.Font.Bold = $true
                            $rowRange.Style.Font.Size = 10
                        } catch {
                            # Skip if color is invalid
                        }
                    }
                }

                # Freeze top row if requested
                if ($FreezeTopRow) {
                    $ws.View.FreezeRows(1)
                }
            }
        }

        if ($PassThru) {
            $safeName = Sanitize-WorksheetName "Top 20 Files"
            $emptyData = @([PSCustomObject]@{ Message = "No files found in this site" })
            $emptyData | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        }

        if ($top10Folders.Count -gt 0) {
            $safeName = Sanitize-WorksheetName "Top 10 Folders"
            $top10Folders | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        } else {
            $safeName = Sanitize-WorksheetName "Top 10 Folders"
            $emptyData = @([PSCustomObject]@{ Message = "No folders found in this site" })
            $emptyData | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        }

        if ($storageBreakdown.Count -gt 0) {
            $safeName = Sanitize-WorksheetName "Storage Breakdown"
            $storageBreakdown | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        } else {
            $safeName = Sanitize-WorksheetName "Storage Breakdown"
            $emptyData = @([PSCustomObject]@{ Message = "No storage data available" })
            $emptyData | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        }

        if ($FolderAccess.Count -gt 0) {
            $safeName = Sanitize-WorksheetName "Folder Access"

            # Add conditional formatting to highlight external guests in red
            $conditionalFormatting = @{
                "D2:D$(($FolderAccess.Count) + 1)" = @{
                    RuleType = "ContainsText"
                    Condition = "External Guest"
                    Color = "Red"
                }
            }

            $FolderAccess | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ConditionalFormatting $conditionalFormatting -ExcelPackage $excel
        } else {
            $safeName = Sanitize-WorksheetName "Folder Access"
            $emptyData = @([PSCustomObject]@{ Message = "No folder access data available" })
            $emptyData | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        }

        if ($accessSummary.Count -gt 0) {
            $safeName = Sanitize-WorksheetName "Access Summary"
            $accessSummary | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        } else {
            $safeName = Sanitize-WorksheetName "Access Summary"
            $emptyData = @([PSCustomObject]@{ Message = "No access summary data available" })
            $emptyData | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        }

        Close-ExcelPackage $excel

        Write-Log "Excel report created successfully!" -Level Success
        Write-Log "`nReport Contents:" -Level Info
        Write-Log "- Summary: Overall site statistics" -Level Info
        Write-Log "- Top 20 Files: Largest files by size" -Level Info
        Write-Log "- Top 10 Folders: Largest folders by size" -Level Info
        Write-Log "- Storage Breakdown: Space usage by location" -Level Info
        Write-Log "- Folder Access: Parent folder permissions" -Level Info
        Write-Log "- Access Summary: Users grouped by permission level" -Level Info
    } catch {
        Write-Log "Failed to create Excel report: $_" -Level Error
        throw
    }
}
                        foreach ($item in $resp.value) {
                            if ($item.driveItem -and $item.driveItem.file) {
                                # Filter out system files
                                $isSystem = $false
                                $fileName = $item.driveItem.name
                                # Skip system/hidden files and folders
                                $systemFilePatterns = @(
                                    "~$*", ".tmp", "thumbs.db", ".ds_store", "desktop.ini", ".git*", ".svn*", "*.lnk", "_vti_*", "forms/", "web.config", "*.aspx", "*.master"
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
                                # Improved file progress bar
                                $currentFileName = if ($item.driveItem.name.Length -gt 50) { $item.driveItem.name.Substring(0, 47) + "..." } else { $item.driveItem.name }
                                Show-Progress -Activity "Analyzing Document Libraries" -Status "Processing: $($list.DisplayName) | Files found: $totalFiles | Current: $currentFileName ($filesInThisList)" -PercentComplete $percentComplete -CurrentOperation "$listIndex of $($docLibraries.Count) libraries"
                            }
                        }
                        # Add small delay to avoid throttling
                        Start-Sleep -Milliseconds (Get-Random -Minimum 100 -Maximum 300)
                    } catch {
                        Write-Log "Failed to process list items: $_" -Level Error
                        # Error handled, continue
                    }
                }
                Write-Log "Processed library: $($list.DisplayName) - Found $filesInThisList files" -Level Success
            } catch {
                Write-Log "Failed to process library $($list.DisplayName): $_" -Level Error
            }
        }
        Stop-Progress -Activity "Analyzing Document Libraries"
        $result = @{
            Files = $allFiles
            FolderSizes = $folderSizes
            TotalFiles = $allFiles.Count
            TotalSizeGB = [math]::Round(($allFiles | Measure-Object -Property Size -Sum).Sum / 1GB, 2)
        }
        Write-Log "Completed site: $($Site.DisplayName) - Files: $($result.TotalFiles), Size: $($result.TotalSizeGB)GB" -Level Success
        return $result
    } catch {
        Write-Log "Failed to get file data for site $($Site.DisplayName): $_" -Level Error
        return @{
            Files = @()
            FolderSizes = @{}
            TotalFiles = 0
            TotalSizeGB = 0
        }
    }
}

function Get-SiteStorageAndAccess {
    param(
        [Parameter(Mandatory=$true)]
        $Site
    )
    
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
        Write-Log "Getting storage and access for site: $($Site.DisplayName)" -Level Info
        
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
                Write-Log "Could not access drive $($drive.Name): $_" -Level Warning
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
        
        $siteInfo.Users = $siteUsers
        $siteInfo.ExternalGuests = $externalGuests
        
        Write-Log "Completed storage and access for site: $($Site.DisplayName) - Users: $($siteUsers.Count), Guests: $($externalGuests.Count)" -Level Success
    }
    catch {
        Write-Log "Failed to get site storage and access info: $_" -Level Error
    }
    
    return $siteInfo
}
##endregion

##region Permission and Access Analysis
function Get-SiteUserAccessSummary {
    param(
        [Parameter(Mandatory=$true)]
        $Site
    )
    
    $owners = @()
    $members = @()
    
    try {
        Write-Log "Getting user access summary for site: $($Site.DisplayName)" -Level Info
        
        # Get site permissions using the correct Graph API
        $permissions = Invoke-WithRetry -ScriptBlock { Get-MgSitePermission -SiteId $Site.Id -All -WarningAction SilentlyContinue } -Activity "Get site permissions"
        
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
                Write-Log "Could not process permission entry: $_" -Level Warning
            }
        }
        
        # Remove duplicates
        $owners = $owners | Sort-Object UserEmail -Unique
        $members = $members | Sort-Object UserEmail -Unique
        
        Write-Log "Found $($owners.Count) owners and $($members.Count) members for site: $($Site.DisplayName)" -Level Success
    } 
    catch {
        Write-Log "Could not retrieve site user access for $($Site.DisplayName): $_" -Level Warning
    }
    
    return @{ 
        Owners = $owners
        Members = $members 
    }
}

function Get-ParentFolderAccess {
    param(
        [Parameter(Mandatory=$true)]
        $Site
    )
    
    $folderAccess = @()
    $processedFolders = @{}
    
    try {
        Write-Log "Getting folder access for site: $($Site.DisplayName)" -Level Info
        
        # Get all drives for the site
        $drives = Invoke-WithRetry -ScriptBlock { Get-MgSiteDrive -SiteId $Site.Id -WarningAction SilentlyContinue } -Activity "Get site drives"
        $driveIndex = 0
        
        foreach ($drive in $drives) {
            $driveIndex++
            $percentComplete = [math]::Round(($driveIndex / $drives.Count) * 100, 1)
            
            # Progress bar for drive processing
            Show-Progress -Activity "Analyzing Folder Permissions" -Status "Processing drive: $($drive.Name)" -PercentComplete $percentComplete -CurrentOperation "$driveIndex of $($drives.Count) drives"
            
            try {
                # Get root folders only (first level)
                $rootFolders = Invoke-WithRetry -ScriptBlock { 
                    Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId "root" -All -ErrorAction Stop | 
                    Where-Object { $_.Folder }
                } -Activity "Get drive root folders"
                
                foreach ($folder in $rootFolders) {
                    if ($processedFolders.ContainsKey($folder.Id)) { continue }
                    $processedFolders[$folder.Id] = $true
                    
                    try {
                        $permissions = Invoke-WithRetry -ScriptBlock { Get-MgDriveItemPermission -DriveId $drive.Id -DriveItemId $folder.Id -All -ErrorAction Stop } -Activity "Get folder permissions"
                        
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
                        Write-Log "Could not get permissions for folder $($folder.Name): $_" -Level Warning
                    }
                }
            }
            catch {
                Write-Log "Could not access drive $($drive.Name): $_" -Level Warning
            }
        }
        
        Stop-Progress -Activity "Analyzing Folder Permissions"
        
        # Remove duplicates (same user with same access to same folder)
        $folderAccess = $folderAccess | Sort-Object FolderName, UserName, PermissionLevel -Unique
        
        Write-Log "Found $($folderAccess.Count) folder access entries for site: $($Site.DisplayName)" -Level Success
    }
    catch {
        Stop-Progress -Activity "Analyzing Folder Permissions"
        Write-Log "Failed to get folder access for site $($Site.DisplayName): $_" -Level Error
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
##endregion

##region Recycle Bin Analysis
## ============================================================================
# Recycle Bin Analysis
## ============================================================================

function Get-SiteRecycleBinStorage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteId
    )

    try {
        # Attempt to retrieve recycle bin items using Microsoft Graph API
        $recycleBinItems = @()

        # Method 1: Direct Graph API call for site recycle bin
        try {
            $recycleBinItems = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/recycleBin" -Method GET -ErrorAction SilentlyContinue
            if ($recycleBinItems.value) {
                $recycleBinItems = $recycleBinItems.value
            }
        } catch {
            # Method 2: Fallback to drive-level recycle bin if site-level fails
            try {
                $drives = Get-MgSiteDrive -SiteId $SiteId -ErrorAction SilentlyContinue
                foreach ($drive in $drives) {
                    try {
                        $driveRecycleBin = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/drives/$($drive.Id)/recycleBin" -Method GET -ErrorAction SilentlyContinue
                        if ($driveRecycleBin.value) {
                            $recycleBinItems += $driveRecycleBin.value
                        }
                    } catch {
                        # Skip drive if recycle bin access fails
                    }
                }
            } catch {
                # No recycle bin access available for this site
            }
        }

        # Calculate total size of recycle bin items (in GB)
        if ($recycleBinItems.Count -gt 0) {
            $totalSize = ($recycleBinItems | Where-Object { $_.size } | Measure-Object -Property size -Sum -ErrorAction SilentlyContinue).Sum
            return [math]::Round($totalSize / 1GB, 3)
        }

        return 0
    } catch {
        return 0
    }
}
##endregion

##region Personal OneDrive Analysis
## ============================================================================
# Personal OneDrive Analysis
## ============================================================================

function Get-PersonalOneDriveSites {
    param(
        [Parameter(Mandatory = $true)]
        $Sites
    )

    # Identify OneDrive sites by URL or name patterns
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
    param(
        [Parameter(Mandatory = $true)]
        $OneDriveSites
    )

    $oneDriveDetails = @()
    $processedCount = 0

    foreach ($site in $OneDriveSites) {
        $processedCount++
        $percentComplete = [math]::Round(($processedCount / $OneDriveSites.Count) * 100, 1)

        Show-Progress -Activity "Analyzing Personal OneDrive Sites" -Status "Processing: $($site.DisplayName)" -PercentComplete $percentComplete -CurrentOperation "$processedCount of $($OneDriveSites.Count) OneDrive sites"

        try {
            # Extract user name from OneDrive URL
            $userName = "Unknown User"
            if ($site.WebUrl -match "/personal/([^/]+)") {
                $userPart = $matches[1] -replace "_", "@"
                $userName = $userPart -replace "([^@]+)@([^@]+)", '$1@$2'
            }

            # Get storage size and file count for user's OneDrive
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

            # Get sharing/access information for user's OneDrive
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
        } catch {
            Write-Log "Could not analyze OneDrive site: $($site.DisplayName) - $_" -Level Warning
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

    Stop-Progress -Activity "Analyzing Personal OneDrive Sites"
    return $oneDriveDetails
}
##endregion

##region Batch Processing
## ============================================================================
# Batch Processing
## ============================================================================

function Get-DrivesBatch {
    param(
        [Parameter(Mandatory = $true)]
        [array]$SiteIds
    )

    # Microsoft Graph batch limit is 20 requests per batch
    $batchSize = 20
    $responses = @()

    for ($i = 0; $i -lt $SiteIds.Count; $i += $batchSize) {
        $batch = $SiteIds[$i..([Math]::Min($i + $batchSize - 1, $SiteIds.Count - 1))]
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
        [Parameter(Mandatory = $true)]
        [array]$DriveIds,
        [string]$ParentId = "root"
    )

    $batchSize = 20
    $responses = @()

    for ($i = 0; $i -lt $DriveIds.Count; $i += $batchSize) {
        $batch = $DriveIds[$i..([Math]::Min($i + $batchSize - 1, $DriveIds.Count - 1))]
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
        [Parameter(Mandatory = $true)]
        [array]$DrivesBatchResponses
    )

    $allFiles = @()
    $folderSizes = @{}

    foreach ($resp in $DrivesBatchResponses) {
        if ($resp.status -eq 200 -and $resp.body.value) {
            foreach ($item in $resp.body.value) {
                if ($item.file) {
                    # Calculate full path+name length (Path + "/" + Name)
                    $fullPath = ($item.parentReference.path + "/" + $item.name).Replace("//", "/")
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
##endregion

##region Excel Report Generation
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
        
        [string]$TableStyle = "Light1",
        
        [hashtable]$ConditionalFormatting = @{},
        
        [switch]$AutoSize,
        
        [switch]$FreezeTopRow,
        
        [switch]$BoldTopRow,
        
        [string]$HeaderColor,
        
        [string]$HeaderTextColor,
        
        [array]$AlternateRowColors = @("LightGray", "White"),
        
        [switch]$PassThru
    )
    
    try {
        $params = @{
            Path = $Path
            WorksheetName = $WorksheetName
            AutoSize = $AutoSize
            PassThru = $PassThru
        }
        
        if ($Title) {
            $params.Add("Title", $Title)
            $params.Add("TitleBold", $true)
            $params.Add("TitleSize", 16)
        }
        
        if ($TableStyle) {
            $params.Add("TableStyle", $TableStyle)
        }
        
        $excel = $Data | Export-Excel @params
        
        # Apply conditional formatting if specified
        if ($ConditionalFormatting.Count -gt 0) {
            $ws = $excel.Workbook.Worksheets[$WorksheetName]
            if ($ws) {
                foreach ($cf in $ConditionalFormatting.GetEnumerator()) {
                    $range = $cf.Key
                    $format = $cf.Value
                    
                    if ($format -is [hashtable]) {
                        $ruleType = $format.RuleType
                        $condition = $format.Condition
                        $color = $format.Color
                        
                        if ($ruleType -eq "ContainsText" -and $condition -and $color) {
                            Add-ConditionalFormatting -Worksheet $ws -Range $range -RuleType ContainsText -Condition $condition -ForegroundColor $color
                        }
                        elseif ($ruleType -eq "Expression" -and $condition -and $color) {
                            Add-ConditionalFormatting -Worksheet $ws -Range $range -RuleType Expression -Condition $condition -ForegroundColor $color
                        }
                    }
                }
            }
        }
        
        # Apply formatting options
        if ($HeaderColor -or $HeaderTextColor -or $BoldTopRow -or $AlternateRowColors -or $FreezeTopRow) {
            $ws = $excel.Workbook.Worksheets[$WorksheetName]
            if ($ws) {
                # Get the range of the header row
                $headerRow = 1
                $headerRange = $ws.Cells[$headerRow, 1, $headerRow, $ws.Dimension.End.Column]
                
                # Apply header formatting
                if ($HeaderColor) {
                    $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $headerRange.Style.Fill.BackgroundColor.SetColor($HeaderColor)
                }
                
                if ($HeaderTextColor) {
                    $headerRange.Style.Font.Color.SetColor($HeaderTextColor)
                }
                
                if ($BoldTopRow) {
                    $headerRange.Style.Font.Bold = $true
                    $headerRange.Style.Font.Size = 12
                }
                
                # Apply alternate row colors
                if ($AlternateRowColors.Count -ge 2) {
                    $dataRows = $ws.Dimension.Rows
                    for ($row = 2; $row -le $dataRows; $row++) {
                        $rowRange = $ws.Cells[$row, 1, $row, $ws.Dimension.End.Column]
                        $colorIndex = ($row - 2) % $AlternateRowColors.Count
                        
                        try {
                            $color = $AlternateRowColors[$colorIndex]
                            if ($color -is [string]) {
                                $color = [System.Drawing.Color]::FromName($color)
                            }
                            
                            $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $rowRange.Style.Fill.BackgroundColor.SetColor($color)
                            
                            # Set contrasting text color
                            if ($color.GetBrightness() -lt 0.5) {
                                $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                            } else {
                                $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::Black)
                            }
                            
                            $rowRange.Style.Font.Bold = $true
                            $rowRange.Style.Font.Size = 10
                        }
                        catch {
                            # Skip if color is invalid
                        }
                    }
                }
                
                # Freeze top row if requested
                if ($FreezeTopRow) {
                    $ws.View.FreezeRows(1)
                }
            }
        }
        
        if ($PassThru) {
            return $excel
        } else {
            Close-ExcelPackage $excel
            return $null
        }
    }
    catch {
        Write-Log "Failed to export Excel worksheet: $_" -Level Error
        throw
    }
}

function New-ExcelReport {
    param(
        [Parameter(Mandatory=$true)]
        $FileData,
        
        [Parameter(Mandatory=$true)]
        $FolderAccess,
        
        [Parameter(Mandatory=$true)]
        $Site,
        
        [Parameter(Mandatory=$true)]
        [string]$FileName
    )
    
    Write-Log "Creating Excel report: $FileName" -Level Info
    
    try {
        $top20Files = $FileData.Files | Sort-Object Size -Descending | Select-Object -First 20 |
            Select-Object Name, SizeMB, Path, Drive, Extension
        
        $top10Folders = $FileData.FolderSizes.GetEnumerator() | 
            Sort-Object Value -Descending | Select-Object -First 10 |
            ForEach-Object { $_ }
        
        # Storage breakdown by location for pie chart
        $storageBreakdown = $FileData.FolderSizes.GetEnumerator() | 
            Sort-Object Value -Descending | Select-Object -First 15 |
            ForEach-Object { $_ }
        
        # Parent folder access summary
        $accessSummary = $FolderAccess | Group-Object PermissionLevel | 
            ForEach-Object { $_ }
        
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
        $safeName = Sanitize-WorksheetName "Summary"
        $excel = $siteSummary | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -PassThru

        # Only export sheets if they have data
        if ($top20Files.Count -gt 0) {
            $safeName = Sanitize-WorksheetName "Top 20 Files"
            $top20Files | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        } else {
            $safeName = Sanitize-WorksheetName "Top 20 Files"
            $emptyData = @([PSCustomObject]@{Message = "No files found in this site"})
            $emptyData | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        }
        
        if ($top10Folders.Count -gt 0) {
            $safeName = Sanitize-WorksheetName "Top 10 Folders"
            $top10Folders | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        } else {
            $safeName = Sanitize-WorksheetName "Top 10 Folders"
            $emptyData = @([PSCustomObject]@{Message = "No folders found in this site"})
            $emptyData | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        }
        
        if ($storageBreakdown.Count -gt 0) {
            $safeName = Sanitize-WorksheetName "Storage Breakdown"
            $storageBreakdown | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        } else {
            $safeName = Sanitize-WorksheetName "Storage Breakdown"
            $emptyData = @([PSCustomObject]@{Message = "No storage data available"})
            $emptyData | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        }
        
        if ($FolderAccess.Count -gt 0) {
            $safeName = Sanitize-WorksheetName "Folder Access"
            
            # Add conditional formatting to highlight external guests in red
            $conditionalFormatting = @{
                "D2:D$(($FolderAccess.Count) + 1)" = @{
                    RuleType = "ContainsText"
                    Condition = "External Guest"
                    Color = "Red"
                }
            }
            
            $FolderAccess | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ConditionalFormatting $conditionalFormatting -ExcelPackage $excel
        } else {
            $safeName = Sanitize-WorksheetName "Folder Access"
            $emptyData = @([PSCustomObject]@{Message = "No folder access data available"})
            $emptyData | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        }
        
        if ($accessSummary.Count -gt 0) {
            $safeName = Sanitize-WorksheetName "Access Summary"
            $accessSummary | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        } else {
            $safeName = Sanitize-WorksheetName "Access Summary"
            $emptyData = @([PSCustomObject]@{Message = "No access summary data available"})
            $emptyData | Export-ExcelWorksheet -Path $FileName -WorksheetName $safeName -AutoSize -ExcelPackage $excel
        }
        
        Close-ExcelPackage $excel

        Write-Log "Excel report created successfully!" -Level Success
        Write-Log "`nReport Contents:" -Level Info
        Write-Log "- Summary: Overall site statistics" -Level Info
        Write-Log "- Top 20 Files: Largest files by size" -Level Info  
        Write-Log "- Top 10 Folders: Largest folders by size" -Level Info
        Write-Log "- Storage Breakdown: Space usage by location" -Level Info
        Write-Log "- Folder Access: Parent folder permissions" -Level Info
        Write-Log "- Access Summary: Users grouped by permission level" -Level Info
    }
    catch {
        Write-Log "Failed to create Excel report: $_" -Level Error
        throw
    }
}
##endregion

##region Microsoft Graph Connection
## ============================================================================

function Connect-ToGraph {
    param(
        [string]$ClientId,
        [string]$TenantId,
        [string]$CertificateThumbprint
    )
    Write-Log "Connecting to Microsoft Graph..." -Level Info
    try {
        # For certificate authentication, do NOT use -Scopes, only use required parameters
        Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop
        Write-Log "Connected to Microsoft Graph successfully." -Level Success
    } catch {
        Write-Log "Failed to connect to Microsoft Graph: $_" -Level Error
        throw
    }
}

function Get-TenantName {
    try {
        # Only import the module if not already loaded
        if (-not (Get-Module -Name Microsoft.Graph.Identity.DirectoryManagement)) {
            Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction SilentlyContinue
        }
        $org = Get-MgOrganization -ErrorAction Stop
        if ($org -and $org.DisplayName) {
            Write-Log "Tenant name found: $($org.DisplayName)" -Level Info
            return $org.DisplayName
        } else {
            Write-Log "Tenant name not found in organization object." -Level Warning
            return "UnknownTenant"
        }
    } catch {
        Write-Log "Failed to get tenant name: $_" -Level Error
        return "UnknownTenant"
    }
}
##endregion

##region Main Processing Functions
## ============================================================================

function Get-SiteSummaries {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Sites,
        [int]$ParallelLimit = 4
    )

    Write-Log "Getting site summaries for $($Sites.Count) sites..." -Level Info

    # Filter all SharePoint library sites with valid Ids
    $filteredSites = $Sites | Where-Object {
        $_.WebUrl -notlike "*-my.sharepoint.com/personal/*" -and
        $_.WebUrl -notlike "*/personal/*" -and
        $_.WebUrl -notlike "*mysites*" -and
        $_.Name -notlike "*OneDrive*" -and
        $_.DisplayName -notlike "*OneDrive*" -and
        $_.Id -and ($_.Id -ne "")
    }

    Write-Log "Filtered SharePoint library sites to process: $($filteredSites.Count)" -Level Info
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
            return $null
        }

        $siteType = if ($webUrl -like "*-my.sharepoint.com/personal/*" -or $webUrl -like "*/personal/*" -or $webUrl -like "*mysites*" -or $displayName -like "*OneDrive*") {
            "OneDrive Personal"
        } else {
            "SharePoint Site"
        }

        $storageGB = 0

        try {
            # Get drives for the site
            $drives = Get-MgSiteDrive -SiteId $siteId -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            
            # Calculate total storage
            foreach ($drive in $drives) {
                try {
                    $driveInfo = Get-MgDrive -DriveId $drive.Id -ErrorAction SilentlyContinue
                    if ($driveInfo) {
                        $storageGB += [math]::Round($driveInfo.Quota.Used / 1GB, 2)
                    }
                } catch {
                    # If quota not available, try to calculate from items
                    try {
                        $items = Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId "root" -All -ErrorAction SilentlyContinue
                        if ($items) {
                            $driveSize = ($items | Measure-Object -Property Size -Sum).Sum
                            $storageGB += [math]::Round($driveSize / 1GB, 2)
                        }
                    } catch {
                        # Silently handle errors
                    }
                }
            }
        } catch {
            Write-Log "Error getting storage for site $displayName`: $_" -Level Warning
        }
        
        return [PSCustomObject]@{
            SiteId = $siteId
            DisplayName = $displayName
            WebUrl = $webUrl
            SiteType = $siteType
            StorageGB = $storageGB
        }
    } -ThrottleLimit $ParallelLimit

    # Filter out null results
    $siteSummaries = $siteSummaries | Where-Object { $_ -ne $null }

    $stepTimer.Stop()
    $elapsed = $stepTimer.Elapsed
    Write-Log "Parallel Site Summary Scan completed in $($elapsed.TotalSeconds) seconds ($($elapsed.TotalMinutes) min)" -Level Info

    # Clear progress bar
    Stop-Progress -Activity "Scanning SharePoint Sites"

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
                    TotalFolders = $fileData.FolderSizes.Count
                    UniquePermissionLevels = $null  # Skipped
                    OwnersCount = $userAccess.Owners.Count
                    MembersCount = $userAccess.Members.Count
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
                TotalFolders = $null
                UniquePermissionLevels = $null
                OwnersCount = $userAccess.Owners.Count
                MembersCount = $userAccess.Members.Count
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
        [hashtable]$SitePieCharts
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
        
        # Calculate storage breakdown
        foreach ($siteSummary in $SiteSummaries) {
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
        Write-Log "Attempting to calculate recycle bin storage..." -Level Info
        foreach ($siteSummary in $SiteSummaries) {
            try {
                $recycleBinSizeForSite = Get-SiteRecycleBinStorage -SiteId $siteSummary.SiteId
                if ($recycleBinSizeForSite -gt 0) {
                    $recycleBinStorage += $recycleBinSizeForSite
                }
            } catch {
                Write-Log "Could not access recycle bin for site: $($siteSummary.SiteName)" -Level Debug
            }
        }
        
        # Create pie chart summary data
        $pieChartData = @(
            [PSCustomObject]@{
                Category = "SharePoint Sites"
                StorageGB = $sharePointSitesStorage
                Percentage = if ($totalTenantStorage -gt 0) { [math]::Round(($sharePointSitesStorage / $totalTenantStorage) * 100, 2) } else { 0 }
                SiteCount = ($SiteSummaries | Where-Object { -not $_.IsOneDrive }).Count
            },
            [PSCustomObject]@{
                Category = "Personal OneDrive"
                StorageGB = $personalOneDriveStorage
                Percentage = if ($totalTenantStorage -gt 0) { [math]::Round(($personalOneDriveStorage / $totalTenantStorage) * 100, 2) } else { 0 }
                SiteCount = ($SiteSummaries | Where-Object { $_.IsOneDrive }).Count
            },
            [PSCustomObject]@{
                Category = "Recycle Bins"
                StorageGB = $recycleBinStorage

                Percentage = if (($totalTenantStorage + $recycleBinStorage) -gt 0) { [math]::Round(($recycleBinStorage / ($totalTenantStorage + $recycleBinStorage)) * 100, 2) } else { 0 }
                SiteCount = "All Sites"
            }
        )
        
        # Create comprehensive site details with users and access information
        Show-Progress -Activity "Generating Excel Report" -Status "Compiling comprehensive site details..." -PercentComplete 10
        
        $comprehensiveSiteDetails = @()
        foreach ($siteSummary in $SiteSummaries) {
            $site = $siteSummary.Site
            
            # Get detailed user access for this site
            try {
                $userAccess = Get-SiteUserAccessSummary -Site $site
                $folderAccess = @()
                
                # Get folder access permissions
                try {
                    $folderAccess = Get-ParentFolderAccess -Site $site
                } catch {
                    Write-Log "Could not get folder access for site: $($site.DisplayName)" -Level Debug
                }
                
                # Compile all users and groups with their access details
                $allUsersString = ""
                $allOwnersString = ""
                $accessTypesString = ""
                $foldersAccessString = ""
                
                if ($userAccess.Owners.Count -gt 0) {
                    $allOwnersString = ($userAccess.Owners | ForEach-Object { "$($_.DisplayName) ($($_.UserEmail))" }) -join "; "
                }
                
                if ($userAccess.Members.Count -gt 0) {
                    $membersString = ($userAccess.Members | ForEach-Object { "$($_.DisplayName) ($($_.UserEmail)) - $($_.Role)" }) -join "; "
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
                    TotalFiles = if ($siteSummary.SiteId -in $SiteSummaries.SiteId) { 
                        ($AllTopFiles | Where-Object { $_.SiteName -eq $site.DisplayName }).Count 
                    } else { "Not analyzed (not in top 10)" }
                    OwnersCount = $userAccess.Owners.Count
                    MembersCount = $userAccess.Members.Count
                    AllOwners = $allOwnersString
                    AllUsersAndGroups = $allUsersString
                    AccessTypes = $accessTypesString
                    FoldersWithAccess = $foldersAccessString
                    UniquePermissionLevels = if ($folderAccess.Count -gt 0) { ($folderAccess.PermissionLevel | Sort-Object -Unique).Count } else { 0 }
                    LastAnalyzed = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    IsTopSite = if ($siteSummary.SiteId -in $SiteSummaries.SiteId) { "Yes" } else { "No" }
                }
            } catch {
                Write-Log "Failed to compile comprehensive details for site $($site.DisplayName): $_" -Level Error
                
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
                    IsTopSite = if ($siteSummary.SiteId -in $SiteSummaries.SiteId) { "Yes" } else { "No" }
                }
            }
        }
        
        # Export SharePoint Storage Pie Chart first (at top of workbook)
        Show-Progress -Activity "Generating Excel Report" -Status "Creating SharePoint Storage Pie Chart..." -PercentComplete 15
        if ($pieChartData.Count -gt 0) {
            $excel = Export-ExcelWorksheet -Data $pieChartData -Path $ExcelFileName -WorksheetName "SharePoint Storage Pie Chart" -AutoSize -TableStyle "Light1" -Title "SharePoint Tenant Storage Overview (Includes Recycle Bins)" -PassThru -HeaderColor "Black" -HeaderTextColor "Yellow" -AlternateRowColors @("DarkGray", "LightGray")
            Close-ExcelPackage $excel
        }
        
        # Export comprehensive site details worksheet
        Show-Progress -Activity "Generating Excel Report" -Status "Creating comprehensive site details..." -PercentComplete 20
        if ($comprehensiveSiteDetails.Count -gt 0) {
            $excel = Export-ExcelWorksheet -Data $comprehensiveSiteDetails -Path $ExcelFileName -WorksheetName "Site Summary with Details" -AutoSize -TableStyle "Light1" -Title "Complete Site Summary with Users, Groups, and Access Details" -PassThru -HeaderColor "Navy" -HeaderTextColor "White" -AlternateRowColors @("DarkBlue", "LightSteelBlue")
            Close-ExcelPackage $excel
        }
        
        # Export detailed storage breakdown
        if ($comprehensiveStorageData.Count -gt 0) {
            $excel = Export-ExcelWorksheet -Data ($comprehensiveStorageData | Sort-Object StorageGB -Descending) -Path $ExcelFileName -WorksheetName "Detailed Storage Breakdown" -AutoSize -TableStyle "Light1" -PassThru -HeaderColor "DarkGreen" -HeaderTextColor "White" -AlternateRowColors @("DarkOliveGreen", "LightGreen")
            Close-ExcelPackage $excel
        }
        
        # Create the original summary for compatibility
        # Split summaries into Document Library (SharePoint) and Personal OneDrive
        $docLibrarySummaries = $AllSiteSummaries | Where-Object { $_.SiteType -eq "SharePoint Site" }
        $oneDriveSummaries = $AllSiteSummaries | Where-Object { $_.SiteType -eq "OneDrive Personal" }
        
        # Export Document Library Sites summary
        if ($docLibrarySummaries.Count -gt 0) {
            $excel = Export-ExcelWorksheet -Data $docLibrarySummaries -Path $ExcelFileName -WorksheetName "Document Library Sites" -AutoSize -TableStyle "Light1" -PassThru -HeaderColor "Black" -HeaderTextColor "Yellow" -AlternateRowColors @("DarkGray", "LightGray")
            Close-ExcelPackage $excel
        }
        
        # Export Personal OneDrive Sites summary
        if ($oneDriveSummaries.Count -gt 0) {
            $excel = Export-ExcelWorksheet -Data $oneDriveSummaries -Path $ExcelFileName -WorksheetName "Personal OneDrive Sites Summary" -AutoSize -TableStyle "Light1" -PassThru -HeaderColor "DarkCyan" -HeaderTextColor "White" -AlternateRowColors @("Teal", "PaleTurquoise")
            Close-ExcelPackage $excel
        }
        
        # Export top 10 sites details if available
        if ($SitePieCharts.Count -gt 0) {
            foreach ($siteName in $SitePieCharts.Keys) {
                $safeName = Sanitize-WorksheetName "$siteName Storage"
                $excel = Export-ExcelWorksheet -Data $SitePieCharts[$siteName] -Path $ExcelFileName -WorksheetName $safeName -AutoSize -TableStyle "Light1" -Title "Storage Breakdown for $siteName" -PassThru
                Close-ExcelPackage $excel
            }
        }
        
        # Export top files and folders if available
        if ($AllTopFiles.Count -gt 0) {
            $safeName = Sanitize-WorksheetName "Top Files Across Sites"
            $excel = Export-ExcelWorksheet -Data ($AllTopFiles | Sort-Object Size -Descending | Select-Object -First 100) -Path $ExcelFileName -WorksheetName $safeName -AutoSize -TableStyle "Light1" -Title "Top 100 Largest Files Across All Sites" -PassThru
            Close-ExcelPackage $excel
        }
        
        if ($AllTopFolders.Count -gt 0) {
            $safeName = Sanitize-WorksheetName "Top Folders Across Sites"
            $excel = Export-ExcelWorksheet -Data ($AllTopFolders | Sort-Object SizeGB -Descending | Select-Object -First 100) -Path $ExcelFileName -WorksheetName $safeName -AutoSize -TableStyle "Light1" -Title "Top 100 Largest Folders Across All Sites" -PassThru
            Close-ExcelPackage $excel
        }
        
        Write-Log "Excel report created successfully: $ExcelFileName" -Level Success
        return $true
    }
    catch {
        Write-Log "Failed to create Excel report: $_" -Level Error
        throw
    }
}
##endregion

##region Main Function
function Main {
    try {
        Write-Log "SharePoint Tenant Storage & Access Report Generator" -Level Success
        Write-Log "=============================================" -Level Success
        
        # Initialize modules
        if (-not (Initialize-Modules)) {
            return
        }
        
        # Connect to Microsoft Graph
        Connect-ToGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
        
        # Get tenant name
        $script:tenantName = Get-TenantName
        
        # Create filename based on tenant name and date
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
        
        # Get detailed information for all sites
        $siteDetails = Get-SiteDetails -SiteSummaries $siteSummaries -TopSites $topSites -ParallelLimit $ParallelLimit
        
        # Generate Excel report
        $success = Export-ComprehensiveExcelReport -ExcelFileName $script:excelFileName -SiteSummaries $siteSummaries -AllSiteSummaries $siteDetails.AllSiteSummaries -AllTopFiles $siteDetails.AllTopFiles -AllTopFolders $siteDetails.AllTopFolders -SiteStorageStats $siteDetails.SiteStorageStats -SitePieCharts $siteDetails.SitePieCharts
        
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
##endregion

##region Script Execution
# Execute the main function
Main
##endregion
