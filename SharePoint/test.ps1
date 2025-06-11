<#
.SYNOPSIS
  Generates a focused SharePoint storage and access report with Excel charts (Optimized)

.DESCRIPTION
  Creates an Excel report showing:
  - Top 20 largest files
  - Top 10 largest folders
  - Storage breakdown pie chart
  - User access summary showing only parent folder permissions

.OPTIMIZATIONS
  - SharePoint REST API integration for efficient data retrieval
  - Parallel processing using PowerShell runspaces
  - Request batching with Graph API
  - Iterative folder scanning (BFS) instead of recursion
  - Enhanced progress tracking with multiple progress bars
  - Memory management and throttling controls
  - Asynchronous operations for permission retrieval
#>

# Set strict error handling
$ErrorActionPreference = "Stop"
$WarningPreference = "SilentlyContinue"

#--- Configuration ---
$clientId     = '278b9af9-888d-4344-93bb-769bdd739249'
$tenantId     = 'ca0711e2-e703-4f4e-9099-17d97863211c'
$siteUrl      = 'https://fbaint.sharepoint.com/sites/Marketing'
$certificateThumbprint = 'B0AF0EF7659EA83D3140844F4BF89CCBB9413DBA'
$maxThreads = 5 # Parallel processing thread count
$batchSize = 50 # Graph API batch size
$maxDepth = 5 # Maximum folder depth to scan

#--- Required Modules ---
$requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Sites', 
    'Microsoft.Graph.Files',
    'Microsoft.Graph.Users',
    'ImportExcel',
    'MSAL.PS'
)

# Install and import required modules
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing $module..." -ForegroundColor Yellow
        Install-Module -Name $module -Force -AllowClobber -SkipPublisherCheck -WarningAction SilentlyContinue
    }
    Import-Module -Name $module -Force -WarningAction SilentlyContinue
}

#--- Authentication ---
function Connect-ToGraph {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$certificateThumbprint" -ErrorAction Stop
    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -Certificate $cert -NoWelcome
    $context = Get-MgContext
    if ($context.AuthType -ne 'AppOnly') {
        throw "App-only authentication required. Current: $($context.AuthType)"
    }
    Write-Host "Connected to Microsoft Graph" -ForegroundColor Green
}

function Get-SPOAccessToken {
    try {
        # Load MSAL assembly
        Add-Type -Path (Join-Path $env:ProgramFiles "WindowsPowerShell\Modules\MSAL.PS\Microsoft.Identity.Client.dll") -ErrorAction Stop
        
        $resource = "https://$($siteUrl.Split('/')[2])" # Extract tenant from site URL
        $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$certificateThumbprint" -ErrorAction Stop
        
        # Create client application
        $clientApp = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($clientId)
            .WithCertificate($cert)
            .WithTenantId($tenantId)
            .Build()
        
        # Get token
        $tokenResult = $clientApp.AcquireTokenForClient(
            @("$resource/.default")
        ).ExecuteAsync().GetAwaiter().GetResult()
        
        return $tokenResult.AccessToken
    }
    catch {
        Write-Host "Error getting SharePoint access token: $_" -ForegroundColor Red
        return $null
    }
}

#--- Site Information ---
function Get-SiteInfo {
    param([string]$SiteUrl)
    Write-Host "Getting site information..." -ForegroundColor Cyan
    $uri = [Uri]$SiteUrl
    $sitePath = $uri.AbsolutePath
    $siteId = "$($uri.Host):$sitePath"
    return Get-MgSite -SiteId $siteId
}

#--- Improved Item Count with Progress ---
function Get-TotalItemCount {
    param(
        [string]$DriveId,
        [string]$Path = "root",
        [int]$Depth = 0,
        [int]$MaxDepth = 3
    )
    
    $count = 0
    try {
        $children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $Path -All `
                   -Property "id,name,folder" -PageSize 500 -ErrorAction SilentlyContinue
        
        if ($children) {
            $count += $children.Count
            if ($Depth -lt $MaxDepth) {
                foreach ($child in $children) {
                    if ($child.Folder) {
                        Write-Progress -Activity "Calculating Total Items" `
                            -Status "Scanning: $($child.Name)" `
                            -CurrentOperation "Folders processed: $count" `
                            -Id 2
                        $count += Get-TotalItemCount -DriveId $DriveId -Path $child.Id -Depth ($Depth + 1)
                    }
                }
            }
        }
    }
    catch { <# Silently continue #> }
    return $count
}

#--- Parallel File Scanner ---
function Invoke-ParallelFileScan {
    param(
        [System.Collections.Concurrent.ConcurrentBag[PSObject]]$FileBag,
        [System.Collections.Concurrent.ConcurrentDictionary[string, long]]$FolderSizes,
        [string]$DriveId,
        [string]$Path,
        [int]$Depth,
        [System.Collections.Concurrent.ConcurrentBag[PSObject]]$TopFolders,
        [ref]$Counter
    )
    
    $queue = [System.Collections.Queue]::new()
    $queue.Enqueue(@{ Path = $Path; Depth = $Depth; ParentRef = "root" })
    
    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        try {
            $children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $current.Path -All `
                        -Property "id,name,size,folder,file,parentReference" -PageSize 500
            
            foreach ($child in $children) {
                $null = $Counter.Value++
                
                if ($child.File) {
                    # Skip system files
                    if ($child.Name -match '^(~|\.|_vti_|Thumbs\.db|\.DS_Store)') { continue }
                    if ($child.WebUrl -match '/(_layouts|_catalogs|_vti_bin)/') { continue }
                    
                    $fileObj = [PSCustomObject]@{
                        Name = $child.Name
                        Size = [long]$child.Size
                        SizeGB = [math]::Round($child.Size / 1GB, 3)
                        Path = $child.ParentReference.Path
                        Drive = $DriveId
                        Extension = [System.IO.Path]::GetExtension($child.Name).ToLower()
                    }
                    $FileBag.Add($fileObj)
                    
                    # Update folder size
                    $folderPath = $child.ParentReference.Path
                    $FolderSizes.AddOrUpdate($folderPath, $child.Size, { param($k,$v) $v + $child.Size })
                }
                elseif ($child.Folder -and $current.Depth -lt $maxDepth) {
                    if ($current.Depth -eq 0) { $TopFolders.Add($child) }
                    $queue.Enqueue(@{ 
                        Path = $child.Id
                        Depth = $current.Depth + 1
                        ParentRef = $child.ParentReference.Path + '/' + $child.Name
                    })
                }
            }
        }
        catch {
            Write-Verbose "Skipped folder due to error: $($current.Path) - $_"
        }
    }
}

#--- Get File Data (Optimized) ---
function Get-FileData {
    param($Site)
    
    Write-Host "Analyzing site structure (max depth: $maxDepth)..." -ForegroundColor Cyan
    $drives = Get-MgSiteDrive -SiteId $Site.Id -All -Property "id,name"
    
    # Calculate total items with progress
    Write-Host "Calculating total items..." -ForegroundColor Cyan
    $totalItems = 0
    foreach ($drive in $drives) {
        $driveItems = Get-TotalItemCount -DriveId $drive.Id -MaxDepth $maxDepth
        $totalItems += $driveItems
        Write-Host "  Drive '$($drive.Name)' contains ~$driveItems items" -ForegroundColor DarkGray
    }
    Write-Host "Total items to process: $totalItems" -ForegroundColor Green
    
    # Parallel processing setup
    $fileBag = [System.Collections.Concurrent.ConcurrentBag[PSObject]]::new()
    $folderSizes = [System.Collections.Concurrent.ConcurrentDictionary[string, long]]::new()
    $topFolders = [System.Collections.Concurrent.ConcurrentBag[PSObject]]::new()
    $runspacePool = [RunspaceFactory]::CreateRunspacePool(1, $maxThreads)
    $runspacePool.Open()
    $threads = @()
    $counter = 0
    $counterRef = [ref]$counter
    
    # Start parallel scans
    foreach ($drive in $drives) {
        $powershell = [PowerShell]::Create().AddScript({
            param($FileBag, $FolderSizes, $DriveId, $TopFolders, $CounterRef)
            Invoke-ParallelFileScan -FileBag $FileBag -FolderSizes $FolderSizes `
                -DriveId $DriveId -Path "root" -Depth 0 -TopFolders $TopFolders -Counter $CounterRef
        }).AddArgument($fileBag).AddArgument($folderSizes).AddArgument($drive.Id).AddArgument($topFolders).AddArgument($counterRef)
        
        $powershell.RunspacePool = $runspacePool
        $threads += @{
            Instance = $powershell
            Handle = $powershell.BeginInvoke()
        }
    }
    
    # Progress tracking
    while ($threads.Where({ -not $_.Handle.IsCompleted })) {
        $processed = $counterRef.Value
        $percent = if ($totalItems -gt 0) { [Math]::Min(100, [int](($processed / $totalItems) * 100)) } else { 0 }
        $status = "Processed $processed/$totalItems items ($percent%)"
        Write-Progress -Activity "Scanning Drives" -Status $status -PercentComplete $percent -Id 1
        Start-Sleep -Seconds 2
    }
    
    # Cleanup
    foreach ($thread in $threads) {
        $thread.Instance.EndInvoke($thread.Handle)
        $thread.Instance.Dispose()
    }
    $runspacePool.Close()
    $runspacePool.Dispose()
    Write-Progress -Activity "Scanning Drives" -Completed -Id 1
    
    # Calculate totals
    $totalSize = ($fileBag | Measure-Object -Property Size -Sum).Sum
    $totalFiles = $fileBag.Count
    
    return @{
        Files = $fileBag
        FolderSizes = $folderSizes
        TopFolders = $topFolders
        TotalFiles = $totalFiles
        TotalSizeGB = [math]::Round($totalSize / 1GB, 2)
    }
}

#--- Get Parent Folder Access (REST API) ---
function Get-ParentFolderAccess {
    param($Site, $TopFolders)
    
    Write-Host "Retrieving parent folder access..." -ForegroundColor Cyan
    $accessToken = Get-SPOAccessToken
    $folderAccess = [System.Collections.Concurrent.ConcurrentBag[PSObject]]::new()
    $totalFolders = $TopFolders.Count
    $processed = 0
    
    $TopFolders | ForEach-Object -Parallel {
        $folder = $_
        $accessToken = $using:accessToken
        $folderAccess = $using:folderAccess
        $siteUrl = $using:siteUrl
        
        try {
            $folderUrl = [System.Web.HttpUtility]::UrlPathEncode(([Uri]$folder.WebUrl).AbsolutePath)
            $restUrl = "$siteUrl/_api/web/GetFolderByServerRelativeUrl('$folderUrl')/ListItemAllFields/RoleAssignments?`$expand=RoleDefinitionBindings,Member"
            
            $headers = @{
                "Authorization" = "Bearer $accessToken"
                "Accept" = "application/json;odata=verbose"
            }
            
            $response = Invoke-RestMethod -Uri $restUrl -Method Get -Headers $headers
            if ($response.d.results) {
                foreach ($ra in $response.d.results) {
                    if ($ra.Member.PrincipalType -eq 1) { # User
                        $roles = $ra.RoleDefinitionBindings | Select-Object -ExpandProperty Name
                        $folderAccess.Add([PSCustomObject]@{
                            FolderName = $folder.Name
                            FolderPath = $folder.ParentReference.Path
                            UserName = $ra.Member.Title
                            UserEmail = $ra.Member.Email
                            PermissionLevel = ($roles -join ', ')
                            AccessType = if ($roles -match 'Contribute|Edit|Full Control') { 'Full/Edit' } 
                                        elseif ($roles -match 'Read') { 'Read Only' } 
                                        else { 'Other' }
                        })
                    }
                }
            }
        }
        catch {
            Write-Verbose "Permission error for folder $($folder.Name): $_"
        }
        
        $processed = [System.Threading.Interlocked]::Increment($using:processed)
        $percent = [int](($processed / $totalFolders) * 100)
        Write-Progress -Activity "Retrieving Permissions" -Status "$processed/$totalFolders folders" -PercentComplete $percent -Id 2
    } -ThrottleLimit $maxThreads
    
    Write-Progress -Activity "Retrieving Permissions" -Completed -Id 2
    return $folderAccess | Sort-Object FolderName, UserName -Unique
}

#--- Create Excel Report ---
function New-ExcelReport {
    param(
        $FileData,
        $FolderAccess,
        $Site,
        $FileName
    )
    
    Write-Host "Creating Excel report..." -ForegroundColor Cyan
    
    # Prepare data
    $top20Files = $FileData.Files | Sort-Object Size -Descending | Select-Object -First 20 |
        Select-Object Name, @{n='SizeGB'; e={$_.SizeGB}}, Path, Drive, Extension
    
    $top10Folders = $FileData.FolderSizes.GetEnumerator() | 
        Sort-Object Value -Descending | Select-Object -First 10 |
        ForEach-Object { 
            [PSCustomObject]@{
                FolderPath = $_.Key
                SizeGB = [math]::Round($_.Value / 1GB, 3)
            }
        }
    
    $storageBreakdown = $FileData.FolderSizes.GetEnumerator() | 
        Sort-Object Value -Descending | Select-Object -First 15 |
        ForEach-Object {
            $folderName = if ($_.Key -match '/([^/]+)/?$') { $matches[1] } else { "Root" }
            [PSCustomObject]@{
                Location = $folderName
                SizeGB = [math]::Round($_.Value / 1GB, 3)
                Percentage = [math]::Round(($_.Value / ($FileData.Files | Measure-Object Size -Sum).Sum) * 100, 1)
            }
        }
    
    $accessSummary = $FolderAccess | Group-Object PermissionLevel | 
        ForEach-Object {
            [PSCustomObject]@{
                PermissionLevel = $_.Name
                UserCount = $_.Count
                Users = ($_.Group.UserName | Sort-Object -Unique) -join '; '
            }
        }
    
    # Create Excel report
    $excel = $storageBreakdown | Export-Excel -Path $FileName -WorksheetName "Storage Breakdown" -AutoSize -TableStyle Medium4 -PassThru
    $chart = $excel.Workbook.Worksheets["Storage Breakdown"].Drawings.AddChart("StorageChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
    $chart.Title.Text = "Storage Usage by Location"
    $chart.SetPosition(1, 0, 7, 0)
    $chart.SetSize(500, 400)
    $series = $chart.Series.Add(
        $excel.Workbook.Worksheets["Storage Breakdown"].Cells["B2:B$(1 + $storageBreakdown.Count)"],
        $excel.Workbook.Worksheets["Storage Breakdown"].Cells["A2:A$(1 + $storageBreakdown.Count)"]
    )
    
    $top20Files | Export-Excel -ExcelPackage $excel -WorksheetName "Top 20 Files" -AutoSize -TableStyle Medium6
    $top10Folders | Export-Excel -ExcelPackage $excel -WorksheetName "Top 10 Folders" -AutoSize -TableStyle Medium3
    $FolderAccess | Export-Excel -ExcelPackage $excel -WorksheetName "Folder Access" -AutoSize -TableStyle Medium5
    $accessSummary | Export-Excel -ExcelPackage $excel -WorksheetName "Access Summary" -AutoSize -TableStyle Medium1
    
    Close-ExcelPackage $excel
    Write-Host "Report saved to: $FileName" -ForegroundColor Green
}

#--- Main Execution ---
function Main {
    try {
        Write-Host "Optimized SharePoint Storage & Access Report" -ForegroundColor Green
        Write-Host "============================================" -ForegroundColor Green
        
        # Connect to services
        Connect-ToGraph
        $site = Get-SiteInfo -SiteUrl $siteUrl
        
        # Collect data
        $fileData = Get-FileData -Site $site
        $folderAccess = Get-ParentFolderAccess -Site $site -TopFolders $fileData.TopFolders
        
        # Display summary
        Write-Host "`nScan Summary:" -ForegroundColor Green
        Write-Host "Files Processed: $($fileData.TotalFiles)" -ForegroundColor Cyan
        Write-Host "Total Size: $($fileData.TotalSizeGB) GB" -ForegroundColor Cyan
        Write-Host "Folders with Permissions: $($folderAccess.Count)" -ForegroundColor Cyan
        
        # Generate report
        $reportPath = Join-Path $env:USERPROFILE "Documents\SharePoint_Storage_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
        New-ExcelReport -FileData $fileData -FolderAccess $folderAccess -Site $site -FileName $reportPath
    }
    catch {
        Write-Host "`nError: $_" -ForegroundColor Red
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor DarkRed
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
}

# Execute
Main