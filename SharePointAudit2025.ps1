# --- SharePoint Audit Full Final Script ---

$clientId = '278b9af9-888d-4344-93bb-769bdd739249'
$tenantId = 'ca0711e2-e703-4f4e-9099-17d97863211c'
$certificateThumbprint = '5DFBBD691A96E55416E5AA35E4B6505A45740324'

# Install required modules if missing
if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) {
    Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
}
if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Sites
Import-Module Microsoft.Graph.Files
Import-Module ImportExcel

function Get-ClientCertificate {
    param ([Parameter(Mandatory)][string]$Thumbprint)
    $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$Thumbprint" -ErrorAction Stop
    if (-not $cert) { throw "Certificate with thumbprint $Thumbprint not found." }
    return $cert
}

function Connect-ToGraph {
    param ([Parameter(Mandatory)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert)
    Write-Host "[Info] Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -Certificate $Cert -NoWelcome -WarningAction SilentlyContinue
    $context = Get-MgContext
    if ($context.AuthType -ne 'AppOnly') { throw "App-only authentication required." }
    Write-Host "[Info] Successfully connected with app-only authentication" -ForegroundColor Green
}

function Get-SharePointAccessToken {
    param ([Parameter(Mandatory)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert)
    try {
        $spToken = Get-MsalToken -ClientId $clientId -TenantId $tenantId -ClientCertificate $Cert -Scopes "https://fbaint.sharepoint.com/.default"
        return $spToken.AccessToken
    } catch {
        Write-Host "[Error] Failed to get SharePoint access token: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Test-GraphContext {
    $context = Get-MgContext
    if (-not $context) { throw "Microsoft Graph context not found." }
    if (-not ($context.Scopes -match 'Sites.Read.All')) { throw "Required scope 'Sites.Read.All' is missing." }
    Write-Host "[Success] Graph context and scopes validated." -ForegroundColor Green
}

function Get-AllSharePointSites {
    Write-Host "[Info] Enumerating SharePoint sites..." -ForegroundColor Cyan
    $sites = Get-MgSite -All | Where-Object { $_.SiteCollection -and $_.WebUrl -notmatch '-my\.sharepoint\.com' -and $_.WebUrl -notmatch '/personal/' }
    if (-not $sites) { throw "No SharePoint sites found." }
    Write-Host "[Info] Found $($sites.Count) sites." -ForegroundColor Green
    return $sites
}

function Get-SiteStorageSummary {
    param (
        [Parameter(Mandatory)]$site,
        [Parameter(Mandatory)]$spToken
    )

    $siteId = $site.Id
    $siteUrl = $site.WebUrl
    $drives = Get-MgSiteDrive -SiteId $siteId -WarningAction SilentlyContinue
    $totalSize = 0
    $totalFileCount = 0
    $allFiles = @()
    $folderSizes = @{}

    foreach ($drive in $drives) {
        $items = Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId 'root' -All -WarningAction SilentlyContinue
        foreach ($item in $items) {
            if ($item.File) {
                $size = $item.Size
                $path = $item.ParentReference.Path
                $allFiles += [PSCustomObject]@{ Name = $item.Name; Path = $path; Size = $size; SizeMB = [math]::Round($size / 1MB, 2) }
                if (-not $folderSizes.ContainsKey($path)) { $folderSizes[$path] = 0 }
                $folderSizes[$path] += $size
                $totalSize += $size
                $totalFileCount++
            }
        }
    }

    $recycleSize = 0
    try {
        $headers = @{ Authorization = "Bearer $spToken"; Accept = "application/json;odata=verbose" }
        $recycleUrl = "$siteUrl/_api/site/RecycleBin"
        $response = Invoke-RestMethod -Uri $recycleUrl -Headers $headers -Method Get -ErrorAction Stop

        if ($response.d.results) {
            foreach ($item in $response.d.results) {
                $recycleSize += $item.Size
            }
        }
    } catch {
        Write-Warning "Failed to get Recycle Bin size for $siteUrl. Setting to 0 GB."
        $recycleSize = 0
    }

    $topFiles = $allFiles | Sort-Object Size -Descending | Select-Object -First 10
    $topFolders = @()
    foreach ($folder in $folderSizes.Keys) {
        $topFolders += [PSCustomObject]@{ Path = $folder; Size = $folderSizes[$folder]; SizeMB = [math]::Round($folderSizes[$folder] / 1MB, 2) }
    }
    $topFolders = $topFolders | Sort-Object Size -Descending | Select-Object -First 10

    return [PSCustomObject]@{
        Name = $site.DisplayName
        Url = $siteUrl
        UserSizeGB = [math]::Round($totalSize / 1GB, 2)
        RecycleSizeGB = [math]::Round($recycleSize / 1GB, 2)
        TotalFileCount = $totalFileCount
        TopFiles = $topFiles
        TopFolders = $topFolders
    }
}

function Get-SiteUserAccess {
    param (
        [Parameter(Mandatory)]$site,
        [Parameter(Mandatory)]$spToken
    )
    $siteId = $site.Id
    $siteUrl = $site.WebUrl
    $accessList = @()
    try {
        $drives = Get-MgSiteDrive -SiteId $siteId -WarningAction SilentlyContinue
        foreach ($drive in $drives) {
            $items = Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId 'root' -All -WarningAction SilentlyContinue
            foreach ($item in $items) {
                if ($item.Folder) {
                    $permissions = Get-MgDriveItemPermission -DriveId $drive.Id -DriveItemId $item.Id -WarningAction SilentlyContinue
                    foreach ($perm in $permissions) {
                        # Handle both GrantedTo (single) and GrantedToIdentities (multiple)
                        $grantees = @()
                        if ($perm.GrantedTo) { $grantees += $perm.GrantedTo }
                        if ($perm.GrantedToIdentities) { $grantees += $perm.GrantedToIdentities }
                        foreach ($grantee in $grantees) {
                            $displayName = $grantee.User.DisplayName
                            $userPrincipalName = $grantee.User.Email
                            $role = if ($perm.Roles) { ($perm.Roles -join ', ') } else { 'User' }
                            if ($displayName -or $userPrincipalName) {
                                $accessList += [PSCustomObject]@{
                                    SiteName = $site.DisplayName
                                    SiteUrl = $siteUrl
                                    UserName = $displayName
                                    UserPrincipalName = $userPrincipalName
                                    Role = $role
                                    MainFolders = $item.Name
                                }
                            }
                        }
                    }
                }
            }
        }
    } catch {
        Write-Warning "Failed to get user access for $($site.WebUrl): $($_.Exception.Message)"
    }
    return $accessList
}

function Get-SaveFilePath {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    $dialog.FileName = "SharePointAudit_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    } else {
        throw "File save canceled."
    }
}

function Export-ResultsToExcel {
    param ([array]$sites, [string]$excelPath, [array]$siteAccess)

    $summaryData = $sites | Select-Object Name, Url, UserSizeGB, RecycleSizeGB, TotalFileCount
    $summaryData | Export-Excel -Path $excelPath -WorksheetName "Summary" -AutoSize -BoldTopRow -TableName "SummaryTable"

    # Add summary pie chart
    $ws = Open-ExcelPackage -Path $excelPath
    Add-ExcelChart -Worksheet $ws.Summary -RangeAddress "C2:C$($summaryData.Count+1)" -ChartType PieExploded3D -Title "Storage Distribution (GB)" -XRange "A2:A$($summaryData.Count+1)" -Row 2 -Column 7
    Close-ExcelPackage $ws

    # Add top 10 largest sites detail sheets
    $topSites = $sites | Sort-Object UserSizeGB -Descending | Select-Object -First 10
    foreach ($site in $topSites) {
        $siteName = ($site.Name -replace '[^\w]', '_').Substring(0, [Math]::Min(28, ($site.Name -replace '[^\w]', '_').Length))
        $folders = $site.TopFolders | Select-Object Path, SizeMB
        $files = $site.TopFiles | Select-Object Name, Path, SizeMB

        # Export folders and files to new sheet
        $folders | Export-Excel -Path $excelPath -WorksheetName "$siteName" -AutoSize -BoldTopRow -StartRow 1 -StartColumn 1 -TableName "${siteName}_Folders"
        $files | Export-Excel -Path $excelPath -WorksheetName "$siteName" -AutoSize -BoldTopRow -StartRow ($folders.Count + 5) -StartColumn 1 -TableName "${siteName}_Files" -Append

        # Add pie chart for folders
        $ws = Open-ExcelPackage -Path $excelPath
        Add-ExcelChart -Worksheet $ws.$siteName -RangeAddress "B2:B$($folders.Count+1)" -ChartType PieExploded3D -Title "Folder Size Breakdown" -XRange "A2:A$($folders.Count+1)" -Row 2 -Column 5
        Close-ExcelPackage $ws
    }

    # Add SiteAccess worksheet
    if ($siteAccess -and $siteAccess.Count -gt 0) {
        $siteAccess | Export-Excel -Path $excelPath -WorksheetName "SiteAccess" -AutoSize -BoldTopRow -TableName "SiteAccessTable" -Append
    }

    Write-Host "[Success] Excel report saved: $excelPath" -ForegroundColor Green
}

# Main execution
try {
    $cert = Get-ClientCertificate -Thumbprint $certificateThumbprint
    Connect-ToGraph -Cert $cert
    Test-GraphContext
    $sites = Get-AllSharePointSites
    $spToken = Get-SharePointAccessToken -Cert $cert
    if (-not $spToken) { throw "Failed to acquire SharePoint token!" }

    $siteResults = @()
    $siteAccess = @()
    $progress = 0
    $totalSites = $sites.Count

    foreach ($site in $sites) {
        $progress++
        Write-Progress -Activity "Auditing SharePoint Sites" -Status "Processing $progress of $totalSites" -PercentComplete (($progress / $totalSites) * 100)
        $summary = Get-SiteStorageSummary -site $site -spToken $spToken
        if ($summary) { $siteResults += $summary }
        $access = Get-SiteUserAccess -site $site -spToken $spToken
        if ($access) { $siteAccess += $access }
    }

    $excelPath = Get-SaveFilePath
    Export-ResultsToExcel -sites $siteResults -excelPath $excelPath -siteAccess $siteAccess

    Write-Host "[Success] Audit complete!" -ForegroundColor Green
    Write-Progress -Activity "Auditing SharePoint Sites" -Completed
}
catch {
    Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Write-Host "[Info] Finished and cleaned up." -ForegroundColor Cyan
}
# --- SharePoint Audit Full Final Script (Advanced Excel + Charts) ---

$clientId = '278b9af9-888d-4344-93bb-769bdd739249'
$tenantId = 'ca0711e2-e703-4f4e-9099-17d97863211c'
$certificateThumbprint = '5DFBBD691A96E55416E5AA35E4B6505A45740324'

# Install required modules if needed
if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) { Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force }
if (-not (Get-Module -Name ImportExcel -ListAvailable)) { Install-Module -Name ImportExcel -Scope CurrentUser -Force }

Import-Module Microsoft.Graph.Sites
Import-Module Microsoft.Graph.Files
Import-Module ImportExcel

function Get-ClientCertificate {
    param ([Parameter(Mandatory)][string]$Thumbprint)
    $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$Thumbprint" -ErrorAction Stop
    if (-not $cert) { throw "Certificate with thumbprint $Thumbprint not found." }
    return $cert
}

function Connect-ToGraph {
    param ([Parameter(Mandatory)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert)
    Write-Host "[Info] Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -Certificate $Cert -NoWelcome -WarningAction SilentlyContinue
    $context = Get-MgContext
    if ($context.AuthType -ne 'AppOnly') { throw "App-only authentication required." }
    Write-Host "[Info] Successfully connected with app-only authentication" -ForegroundColor Green
}

function Get-SharePointAccessToken {
    param ([Parameter(Mandatory)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert)
    try {
        $spToken = Get-MsalToken -ClientId $clientId -TenantId $tenantId -ClientCertificate $Cert -Scopes "https://fbaint.sharepoint.com/.default"
        return $spToken.AccessToken
    } catch {
        Write-Host "[Error] Failed to get SharePoint access token: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Test-GraphContext {
    $context = Get-MgContext
    if (-not $context) { throw "Microsoft Graph context not found." }
    if (-not ($context.Scopes -match 'Sites.Read.All')) { throw "Required scope 'Sites.Read.All' is missing." }
    Write-Host "[Success] Graph context and scopes validated." -ForegroundColor Green
}

function Get-AllSharePointSites {
    Write-Host "[Info] Enumerating SharePoint sites..." -ForegroundColor Cyan
    $sites = Get-MgSite -All | Where-Object { $_.SiteCollection -and $_.WebUrl -notmatch '-my\.sharepoint\.com' -and $_.WebUrl -notmatch '/personal/' }
    if (-not $sites) { throw "No SharePoint sites found." }
    Write-Host "[Info] Found $($sites.Count) sites." -ForegroundColor Green
    return $sites
}

function Get-SiteStorageSummary {
    param (
        [Parameter(Mandatory)]$site,
        [Parameter(Mandatory)]$spToken
    )

    $siteId = $site.Id
    $siteUrl = $site.WebUrl
    $drives = Get-MgSiteDrive -SiteId $siteId -WarningAction SilentlyContinue
    $totalSize = 0
    $totalFileCount = 0
    $allFiles = @()
    $folderSizes = @{}

    foreach ($drive in $drives) {
        $items = Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId 'root' -All -WarningAction SilentlyContinue
        foreach ($item in $items) {
            if ($item.File) {
                $size = $item.Size
                $path = $item.ParentReference.Path
                $allFiles += [PSCustomObject]@{ Name = $item.Name; Path = $path; Size = $size; SizeMB = [math]::Round($size / 1MB, 2) }
                if (-not $folderSizes.ContainsKey($path)) { $folderSizes[$path] = 0 }
                $folderSizes[$path] += $size
                $totalSize += $size
                $totalFileCount++
            }
        }
    }

    $recycleSize = 0
    try {
        $headers = @{ Authorization = "Bearer $spToken"; Accept = "application/json;odata=verbose" }
        $recycleUrl = "$siteUrl/_api/site/RecycleBin"
        $response = Invoke-RestMethod -Uri $recycleUrl -Headers $headers -Method Get -ErrorAction Stop

        if ($response.d.results) {
            foreach ($item in $response.d.results) {
                $recycleSize += $item.Size
            }
        }
    } catch {
        Write-Warning "Failed to get Recycle Bin size for $siteUrl. Setting to 0 GB."
        $recycleSize = 0
    }

    $topFiles = $allFiles | Sort-Object Size -Descending | Select-Object -First 10
    $topFolders = @()
    foreach ($folder in $folderSizes.Keys) {
        $topFolders += [PSCustomObject]@{ Path = $folder; Size = $folderSizes[$folder]; SizeMB = [math]::Round($folderSizes[$folder] / 1MB, 2) }
    }
    $topFolders = $topFolders | Sort-Object Size -Descending | Select-Object -First 10

    return [PSCustomObject]@{
        Name = $site.DisplayName
        Url = $siteUrl
        UserSizeGB = [math]::Round($totalSize / 1GB, 2)
        RecycleSizeGB = [math]::Round($recycleSize / 1GB, 2)
        TotalFileCount = $totalFileCount
        TopFiles = $topFiles
        TopFolders = $topFolders
    }
}

function Get-SiteUserAccess {
    param (
        [Parameter(Mandatory)]$site,
        [Parameter(Mandatory)]$spToken
    )
    $siteId = $site.Id
    $siteUrl = $site.WebUrl
    $accessList = @()
    try {
        $drives = Get-MgSiteDrive -SiteId $siteId -WarningAction SilentlyContinue
        foreach ($drive in $drives) {
            $items = Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId 'root' -All -WarningAction SilentlyContinue
            foreach ($item in $items) {
                if ($item.Folder) {
                    $permissions = Get-MgDriveItemPermission -DriveId $drive.Id -DriveItemId $item.Id -WarningAction SilentlyContinue
                    foreach ($perm in $permissions) {
                        # Handle both GrantedTo (single) and GrantedToIdentities (multiple)
                        $grantees = @()
                        if ($perm.GrantedTo) { $grantees += $perm.GrantedTo }
                        if ($perm.GrantedToIdentities) { $grantees += $perm.GrantedToIdentities }
                        foreach ($grantee in $grantees) {
                            $displayName = $grantee.User.DisplayName
                            $userPrincipalName = $grantee.User.Email
                            $role = if ($perm.Roles) { ($perm.Roles -join ', ') } else { 'User' }
                            if ($displayName -or $userPrincipalName) {
                                $accessList += [PSCustomObject]@{
                                    SiteName = $site.DisplayName
                                    SiteUrl = $siteUrl
                                    UserName = $displayName
                                    UserPrincipalName = $userPrincipalName
                                    Role = $role
                                    MainFolders = $item.Name
                                }
                            }
                        }
                    }
                }
            }
        }
    } catch {
        Write-Warning "Failed to get user access for $($site.WebUrl): $($_.Exception.Message)"
    }
    return $accessList
}

function Get-SaveFilePath {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    $dialog.FileName = "SharePointAudit_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    } else {
        throw "File save canceled."
    }
}

function Export-ResultsToExcel {
    param ([array]$sites, [string]$excelPath, [array]$siteAccess)

    $summaryData = $sites | Select-Object Name, Url, UserSizeGB, RecycleSizeGB, TotalFileCount
    $summaryData | Export-Excel -Path $excelPath -WorksheetName "Summary" -AutoSize -BoldTopRow -TableName "SummaryTable"

    # Add summary pie chart
    $ws = Open-ExcelPackage -Path $excelPath
    Add-ExcelChart -Worksheet $ws.Summary -RangeAddress "C2:C$($summaryData.Count+1)" -ChartType PieExploded3D -Title "Storage Distribution (GB)" -XRange "A2:A$($summaryData.Count+1)" -Row 2 -Column 7
    Close-ExcelPackage $ws

    # Add top 10 largest sites detail sheets
    $topSites = $sites | Sort-Object UserSizeGB -Descending | Select-Object -First 10
    foreach ($site in $topSites) {
        $siteName = ($site.Name -replace '[^\w]', '_').Substring(0, [Math]::Min(28, ($site.Name -replace '[^\w]', '_').Length))
        $folders = $site.TopFolders | Select-Object Path, SizeMB
        $files = $site.TopFiles | Select-Object Name, Path, SizeMB

        # Export folders and files to new sheet
        $folders | Export-Excel -Path $excelPath -WorksheetName "$siteName" -AutoSize -BoldTopRow -StartRow 1 -StartColumn 1 -TableName "${siteName}_Folders"
        $files | Export-Excel -Path $excelPath -WorksheetName "$siteName" -AutoSize -BoldTopRow -StartRow ($folders.Count + 5) -StartColumn 1 -TableName "${siteName}_Files" -Append

        # Add pie chart for folders
        $ws = Open-ExcelPackage -Path $excelPath
        Add-ExcelChart -Worksheet $ws.$siteName -RangeAddress "B2:B$($folders.Count+1)" -ChartType PieExploded3D -Title "Folder Size Breakdown" -XRange "A2:A$($folders.Count+1)" -Row 2 -Column 5
        Close-ExcelPackage $ws
    }

    # Add SiteAccess worksheet
    if ($siteAccess -and $siteAccess.Count -gt 0) {
        $siteAccess | Export-Excel -Path $excelPath -WorksheetName "SiteAccess" -AutoSize -BoldTopRow -TableName "SiteAccessTable" -Append
    }

    Write-Host "[Success] Excel report saved: $excelPath" -ForegroundColor Green
}

# Main execution
try {
    $cert = Get-ClientCertificate -Thumbprint $certificateThumbprint
    Connect-ToGraph -Cert $cert
    Test-GraphContext
    $sites = Get-AllSharePointSites
    $spToken = Get-SharePointAccessToken -Cert $cert
    if (-not $spToken) { throw "Failed to acquire SharePoint token!" }

    $siteResults = @()
    $siteAccess = @()
    $progress = 0
    $totalSites = $sites.Count

    foreach ($site in $sites) {
        $progress++
        Write-Progress -Activity "Auditing SharePoint Sites" -Status "Processing $progress of $totalSites" -PercentComplete (($progress / $totalSites) * 100)
        $summary = Get-SiteStorageSummary -site $site -spToken $spToken
        if ($summary) { $siteResults += $summary }
        $access = Get-SiteUserAccess -site $site -spToken $spToken
        if ($access) { $siteAccess += $access }
    }

    $excelPath = Get-SaveFilePath
    Export-ResultsToExcel -sites $siteResults -excelPath $excelPath -siteAccess $siteAccess

    Write-Host "[Success] Audit complete!" -ForegroundColor Green
    Write-Progress -Activity "Auditing SharePoint Sites" -Completed
}
catch {
    Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Write-Host "[Info] Finished and cleaned up." -ForegroundColor Cyan
}
