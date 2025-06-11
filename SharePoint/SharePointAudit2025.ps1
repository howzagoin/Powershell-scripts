#
<#
.SYNOPSIS
  Audits a SharePoint site via Microsoft Graph (app-only auth)

.DESCRIPTION
  Fixed version with better error handling and troubleshooting for drive access issues.
#>

#--- Configuration ---
$clientId     = ''
$tenantId     = ''
$clientSecret = ''
$siteUrl      = ''

#--- Modules ---
function Install-ModuleIfMissing {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Install-Module -Name $Name -Scope CurrentUser -Force -ErrorAction Stop
    }
    Import-Module -Name $Name -Force -ErrorAction Stop
}
Install-ModuleIfMissing 'Microsoft.Graph'
Install-ModuleIfMissing 'ImportExcel'

#--- Improved Error Handling ---
function Invoke-LoggedCommand {
    param([string]$Description, [scriptblock]$Action)
    Write-Host "`n[+] $Description" -ForegroundColor Cyan
    try {
        $result = & $Action
        Write-Host "[√] Success" -ForegroundColor Green
        return $result
    } catch {
        Write-Host "[X] Error during '$Description'" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "StackTrace: $($_.ScriptStackTrace)" -ForegroundColor DarkYellow
        throw
    }
}

#--- Authentication ---
function Connect-Graph {
    Invoke-LoggedCommand 'Authenticating to Microsoft Graph' {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Remove-Item "$env:USERPROFILE\.mgcontext" -ErrorAction SilentlyContinue
        
        # Connect with retry logic
        $retryCount = 0
        $maxRetries = 3
        $connected = $false
        
        while (-not $connected -and $retryCount -lt $maxRetries) {
            try {
                Connect-MgGraph -ClientId $clientId -TenantId $tenantId -ClientSecret $clientSecret `
                    -Scopes @("https://graph.microsoft.com/.default") -ErrorAction Stop
                
                if ((Get-MgContext).AuthType -ne 'AppOnly') {
                    throw "Expected AppOnly auth, got '$((Get-MgContext).AuthType)'"
                }
                $connected = $true
            } catch {
                $retryCount++
                if ($retryCount -ge $maxRetries) {
                    throw
                }
                Start-Sleep -Seconds (5 * $retryCount)
            }
        }
    }
}

#--- Site Resolution ---
function Get-FilesFromDrive {
    param(
        [string]$DriveId,
        [string]$SiteId
    )
    
    Invoke-LoggedCommand "Processing drive $DriveId" {
        # Initialize variables
        $files = @()
        $processedItems = 0
        $foldersFound = 0
        $filesFound = 0
        $excludedFiles = 0

        # Use stack for DFS traversal
        $stack = New-Object System.Collections.Stack
        $stack.Push(@{Id = "root"; Depth = 1})

        # Set maximum items to process
        $maxItems = 5000
        $maxDepth = 20

        while ($stack.Count -gt 0 -and $processedItems -lt $maxItems) {
            $current = $stack.Pop()
            $currentItemId = $current.Id
            $currentDepth = $current.Depth

            Write-Host "Processing item $currentItemId (Depth: $currentDepth)" -ForegroundColor DarkGray

            try {
                # Get children with retry logic
                $children = $null
                $retryCount = 0
                $maxRetries = 3
                
                while ($retryCount -lt $maxRetries) {
                    try {
                        $children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $currentItemId -All -ErrorAction Stop
                        break
                    } catch {
                        $retryCount++
                        if ($retryCount -ge $maxRetries) {
                            throw
                        }
                        Start-Sleep -Seconds (2 * $retryCount)
                    }
                }

                # Process children in reverse order to maintain DFS order
                $reverseChildren = $children | Sort-Object { $_.Folder -ne $null } -Descending
                
                foreach ($item in $reverseChildren) {
                    $processedItems++
                    
                    # Progress reporting
                    if ($processedItems % 100 -eq 0) {
                        Write-Host "Processed $processedItems items..." -ForegroundColor Gray
                    }

                    # Debug output for first 20 items
                    if ($processedItems -le 20) {
                        $type = if ($item.Folder) { "Folder" } else { "File" }
                        Write-Host "  Found: $($item.Name) | Type: $type | Path: $($item.ParentReference.Path)" -ForegroundColor DarkCyan
                    }

                    if ($item.Folder) {
                        $foldersFound++
                        # Push folders to stack (with increased depth)
                        if ($currentDepth -lt $maxDepth) {
                            $stack.Push(@{Id = $item.Id; Depth = $currentDepth + 1})
                            if ($processedItems -le 20) {
                                Write-Host "  Added folder to stack: $($item.Name)" -ForegroundColor DarkGreen
                            }
                        }
                    }
                    elseif ($item.File) {
                        $filesFound++
                        $ext = ([IO.Path]::GetExtension($item.Name)).ToLower()
                        
                        # Very permissive filtering - only exclude true system files
                        if ($ext -notin @('.aspx','.master','.webpart')) {
                            $fileObj = [PSCustomObject]@{
                                Name    = $item.Name
                                Path    = $item.ParentReference.Path
                                Size    = $item.Size
                                DriveId = $DriveId
                                ItemId  = $item.Id
                                Extension = $ext
                                CreatedDateTime = $item.CreatedDateTime
                                LastModifiedDateTime = $item.LastModifiedDateTime
                                WebUrl  = $item.WebUrl
                            }
                            $files += $fileObj
                            if ($processedItems -le 20) {
                                Write-Host "  Added file: $($item.Name) ($ext)" -ForegroundColor Green
                            }
                        } else {
                            $excludedFiles++
                            if ($processedItems -le 20) {
                                Write-Host "  Excluded system file: $($item.Name)" -ForegroundColor Yellow
                            }
                        }
                    }
                }
            } catch {
                Write-Host "Warning: Could not process item $currentItemId - $($_.Exception.Message)" -ForegroundColor Yellow
                continue
            }
        }

        Write-Host "`nDrive Processing Summary:" -ForegroundColor Cyan
        Write-Host "  Total items processed: $processedItems" -ForegroundColor White
        Write-Host "  Folders found: $foldersFound" -ForegroundColor White
        Write-Host "  Files found: $filesFound" -ForegroundColor White
        Write-Host "  Files excluded: $excludedFiles" -ForegroundColor White
        Write-Host "  User files included: $($files.Count)" -ForegroundColor Green

        # Debug: Show sample of files if found
        if ($files.Count -gt 0) {
            Write-Host "`nSample files found:" -ForegroundColor Cyan
            $files | Select-Object -First 5 | Format-Table Name, Size, Extension, WebUrl
        } else {
            # If no files found, try a direct query to root
            Write-Host "`nNo files found in traversal. Trying direct root query..." -ForegroundColor Yellow
            try {
                $rootItems = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId "root" -All -ErrorAction Stop
                $rootFiles = $rootItems | Where-Object { $_.File }
                Write-Host "Found $($rootFiles.Count) files in root directory" -ForegroundColor Cyan
                $rootFiles | Select-Object -First 5 | Format-Table Name, Size, @{Name="Type";Expression={if($_.File){"File"}else{"Folder"}}}
            } catch {
                Write-Host "Direct root query failed: $_" -ForegroundColor Red
            }
        }

        return $files
    }
}

#--- Enhanced File Retrieval ---
function Get-FilesFromDrive {
    param(
        [string]$DriveId,
        [string]$SiteId
    )
    
    Invoke-LoggedCommand "Processing drive $DriveId" {
        $stack = @('root')
        $files = @()
        $processedItems = 0
        $foldersFound = 0
        $filesFound = 0
        $excludedFiles = 0
        $maxDepth = 20 # Prevent infinite recursion
        $currentDepth = 0

        while ($stack.Count -gt 0 -and $currentDepth -lt $maxDepth) {
            $currentDepth++
            $current = $stack[0]
            $stack = $stack[1..($stack.Count - 1)]
            
            Write-Host "Processing item $current (Depth: $currentDepth)" -ForegroundColor DarkGray
            
            try {
                $children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $current -All -ErrorAction Stop
                
                foreach ($item in $children) {
                    $processedItems++
                    if ($processedItems % 100 -eq 0) {
                        Write-Host "Processed $processedItems items..." -ForegroundColor Gray
                    }
                    
                    # Debug: Show what we're finding
                    if ($processedItems -le 20) {
                        Write-Host "  Found: $($item.Name) | Type: $(if($item.Folder){'Folder'}else{'File'}) | Path: $($item.ParentReference.Path)" -ForegroundColor DarkCyan
                    }
                    
                    $pathLower = if ($item.ParentReference.Path) { $item.ParentReference.Path.ToLower() } else { "" }
                    
                    if ($item.Folder) {
                        $foldersFound++
                        # More permissive folder filtering - only exclude system folders
                        if ($pathLower -notmatch '/forms$|/siteassets$|/site pages$|/_catalogs|/_layouts') {
                            $stack += $item.Id
                            if ($processedItems -le 20) {
                                Write-Host "  Added folder to queue: $($item.Name)" -ForegroundColor DarkGreen
                            }
                        } else {
                            if ($processedItems -le 20) {
                                Write-Host "  Excluded system folder: $($item.Name)" -ForegroundColor DarkYellow
                            }
                        }
                    }
                    elseif ($item.File) {
                        $filesFound++
                        $ext = ([IO.Path]::GetExtension($item.Name)).ToLower()
                        
                        # More permissive file filtering - include more file types
                        if ($ext -notin @('.aspx','.js','.css','.master','.html','.xml','.webpart')) {
                            $files += [PSCustomObject]@{
                                Name    = $item.Name
                                Path    = $item.ParentReference.Path
                                Size    = $item.Size
                                DriveId = $DriveId
                                ItemId  = $item.Id
                                Extension = $ext
                                CreatedDateTime = $item.CreatedDateTime
                                LastModifiedDateTime = $item.LastModifiedDateTime
                            }
                            if ($processedItems -le 20) {
                                Write-Host "  Added file: $($item.Name) ($ext)" -ForegroundColor Green
                            }
                        } else {
                            $excludedFiles++
                            if ($processedItems -le 20) {
                                Write-Host "  Excluded system file: $($item.Name) ($ext)" -ForegroundColor Yellow
                            }
                        }
                    }
                }
            } catch {
                Write-Host "Warning: Could not process item $current - $_" -ForegroundColor Yellow
                continue
            }
        }
        
        Write-Host "`nDrive Processing Summary:" -ForegroundColor Cyan
        Write-Host "  Total items processed: $processedItems" -ForegroundColor White
        Write-Host "  Folders found: $foldersFound" -ForegroundColor White
        Write-Host "  Files found: $filesFound" -ForegroundColor White
        Write-Host "  Files excluded: $excludedFiles" -ForegroundColor White
        Write-Host "  User files included: $($files.Count)" -ForegroundColor Green
        
        return $files
    }
}

#--- Main File Collection ---
function Get-AllFiles {
    param($Site)  # Removed specific type constraint to handle array/object flexibility
    
    Invoke-LoggedCommand 'Gathering all user-uploaded files' {
        try {
            $drives = Get-MgSiteDrive -SiteId $Site.Id -ErrorAction Stop
            Write-Host "Found $($drives.Count) drives in site" -ForegroundColor Cyan
            
            $results = @()
            foreach ($drive in $drives) {
                Write-Host "Processing drive: $($drive.Name) (ID: $($drive.Id))" -ForegroundColor Cyan
                try {
                    $driveFiles = Get-FilesFromDrive -DriveId $drive.Id -SiteId $Site.Id
                    $results += $driveFiles
                } catch {
                    Write-Host "Skipping drive $($drive.Id) due to error: $_" -ForegroundColor Yellow
                    continue
                }
            }
            
            return $results
        } catch {
            Write-Host "Error accessing drives. Ensure the app has 'Files.Read.All' permission." -ForegroundColor Red
            throw
        }
    }
}

#--- SharePoint Groups and Users ---
function Get-SharePointGroupsAndUsers {
    param($Site)
    
    Invoke-LoggedCommand 'Gathering SharePoint groups and users' {
        try {
            # Get site permissions
            $sitePermissions = Get-MgSitePermission -SiteId $Site.Id -ErrorAction SilentlyContinue
            Write-Host "Found $($sitePermissions.Count) site permissions" -ForegroundColor Cyan
            
            # Display permissions if found
            if ($sitePermissions.Count -gt 0) {
                Write-Host "Site Permissions:" -ForegroundColor Cyan
                foreach ($perm in $sitePermissions | Select-Object -First 5) {
                    Write-Host "  ID: $($perm.Id) | Roles: $($perm.Roles -join ', ')" -ForegroundColor White
                }
            }
            
            return $sitePermissions
        } catch {
            Write-Host "Warning: Could not retrieve site permissions - $_" -ForegroundColor Yellow
            return @()
        }
    }
}

#--- Excel Export ---
function Export-ExcelReport {
    param($Data)
    
    Invoke-LoggedCommand 'Exporting to Excel' {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $filename = "SharePoint_Audit_$timestamp.xlsx"
        
        try {
            $Data | Export-Excel -Path $filename -AutoSize -TableStyle Medium2 -WorksheetName "Files"
            Write-Host "Report exported to: $filename" -ForegroundColor Green
        } catch {
            Write-Host "Excel export failed. Exporting to CSV instead..." -ForegroundColor Yellow
            $csvFile = "SharePoint_Audit_$timestamp.csv"
            $Data | Export-Csv -Path $csvFile -NoTypeInformation
            Write-Host "Report exported to: $csvFile" -ForegroundColor Green
        }
    }
}

#--- Main Execution ---
function Main {
    try {
        Connect-Graph
        $site = Get-TargetSite -Url $siteUrl
        
        # Verify permissions
        Write-Host "`n[!] Verifying required permissions..." -ForegroundColor Cyan
        try {
            $test = Get-MgSiteDrive -SiteId $site.Id -Top 1 -ErrorAction Stop
            Write-Host "[√] Drive access verified" -ForegroundColor Green
        } catch {
            Write-Host "[X] Drive access failed. Required permissions:" -ForegroundColor Red
            Write-Host " - Sites.Read.All or Sites.ReadWrite.All" -ForegroundColor Yellow
            Write-Host " - Files.Read.All or Files.ReadWrite.All" -ForegroundColor Yellow
            throw "Insufficient permissions"
        }
        
        $allFiles = Get-AllFiles -Site $site
        Get-SharePointGroupsAndUsers -Site $site

        # Display results
        Write-Host "`n[=] Audit Results Summary:" -ForegroundColor Green
        Write-Host " - Total files found: $($allFiles.Count)"
        if ($allFiles.Count -gt 0) {
            Write-Host " - Total size: $([math]::Round(($allFiles | Measure-Object -Property Size -Sum).Sum / 1MB, 2)) MB"
            
            # Show file type breakdown
            $fileTypes = $allFiles | Group-Object Extension | Sort-Object Count -Descending
            Write-Host " - File types found:" -ForegroundColor Cyan
            foreach ($type in $fileTypes | Select-Object -First 10) {
                Write-Host "   $($type.Name): $($type.Count) files" -ForegroundColor White
            }
            
            # Show largest files
            $largestFiles = $allFiles | Sort-Object Size -Descending | Select-Object -First 5
            Write-Host " - Largest files:" -ForegroundColor Cyan
            foreach ($file in $largestFiles) {
                $sizeMB = [math]::Round($file.Size / 1MB, 2)
                Write-Host "   $($file.Name): $sizeMB MB" -ForegroundColor White
            }
        } else {
            Write-Host " - No user files found. This could mean:" -ForegroundColor Yellow
            Write-Host "   1. The site contains only SharePoint system files" -ForegroundColor White
            Write-Host "   2. All files are in excluded system folders" -ForegroundColor White
            Write-Host "   3. The drive might be empty or access restricted" -ForegroundColor White
        }
        
        $choice = Read-Host "`nDo you want to export results to Excel? (Y/N)"
        if ($choice -match '^(y|yes)$') {
            if ($allFiles.Count -gt 0) {
                Export-ExcelReport -Data $allFiles
            } else {
                Write-Host "No files to export." -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Host "`n[!] Script failed with error:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Yellow
        Write-Host "`nTroubleshooting tips:" -ForegroundColor Cyan
        Write-Host "1. Verify the app registration has these permissions:" -ForegroundColor White
        Write-Host "   - Sites.Read.All" -ForegroundColor White
        Write-Host "   - Files.Read.All" -ForegroundColor White
        Write-Host "   - Group.Read.All" -ForegroundColor White
        Write-Host "2. Check the client secret hasn't expired" -ForegroundColor White
        Write-Host "3. Verify the service principal has been created in the tenant" -ForegroundColor White
        exit 1
    }
}

Main