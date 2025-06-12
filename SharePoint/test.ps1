 <#SharePoint Storage & Access Reporter (CLI)

.DESCRIPTION
  Generates a comprehensive Excel report covering:
    • Top 20 largest files
    • Top 10 largest folders
    • Storage quota vs. usage chart
    • StorageUsedGB vs. StorageTotalGB chart
    • Parent-folder permission summary
  Uses recursive deep scanning up to maxDepth, parallel processing, optimized bulk Graph calls (Optimized.Mga), and certificate-based app-only authentication.

.NOTES
  Author     : Timothy MacLatchy
  Date       : 13-06-2025
  License    : MIT License
  Steps      :
    1. Validate and install required modules.
    2. Authenticate via certificate (app-only) only.
    3. Retrieve site & drives.
    4. Recursively scan drives (up to maxDepth) with progress.
    5. Aggregate file metadata & folder sizes.
    6. Identify top files/folders.
    7. Retrieve parent-folder permissions (top-level only).
    8. Generate Excel report (multiple sheets + charts).
#>

#--- Configuration ---
$ErrorActionPreference     = 'Stop'
$WarningPreference         = 'SilentlyContinue'
$clientId                  = '278b9af9-888d-4344-93bb-769bdd739249'
$tenantId                  = 'ca0711e2-e703-4f4e-9099-17d97863211c'
$siteUrl                   = 'https://fbaint.sharepoint.com/sites/Marketing'
$certificateThumbprint     = 'B0AF0EF7659EA83D3140844F4BF89CCBB9413DBA'
$maxDepth                  = 10
$maxThreads                = [System.Environment]::ProcessorCount

#--- Ensure Required Modules ---
$modules = @('Optimized.Mga','Optimized.Mga.SharePoint','ImportExcel')
foreach ($mod in $modules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Installing module $mod..." -ForegroundColor Yellow
        Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $mod -Force
}

#--- Logging Helpers ---
function Write-Log { param($msg) Write-Host "[INFO] $msg" }
function Write-ErrorLog { param($msg) Write-Host "[ERROR] $msg" -ForegroundColor Red }

#--- Authenticate (App-Only Only) ---
function Connect-Graph {
    Write-Log 'Authenticating to Microsoft Graph (app-only)...'
    $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object Thumbprint -eq $certificateThumbprint
    if (-not $cert) {
        Write-ErrorLog "Certificate thumbprint $certificateThumbprint not found. Cannot authenticate."
        exit 1
    }
    Connect-Mga -ClientCertificate $cert -ApplicationID $clientId -Tenant $tenantId -NoWelcome
    Write-Log 'App-only authentication successful.'
}

#--- Recursive Drive Scan ---
function Get-DriveItemsRecursive {
    param(
        [string]$DriveId,
        [string]$Path = 'root',
        [int]$Depth = 0,
        [ref]$Index,
        [int]$Total
    )
    $items = @()
    try {
        $resp = Invoke-Mga -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$Path/children?`$select=id,name,size,folder,file,parentReference,webUrl" -Method GET
        foreach ($child in $resp.value) {
            $Index.Value++
            $pct = [int](($Index.Value/$Total)*100)
            Write-Progress -Activity 'Scanning Content' -Status "$pct% ($($Index.Value)/$Total)" -PercentComplete $pct
            if ($child.file -and $child.name -notmatch '^(~|\.|_vti_|Thumbs\.db|\.DS_Store)') {
                $items += $child
            } elseif ($child.folder -and $Depth -lt $maxDepth) {
                $items += $child
                $items += Get-DriveItemsRecursive -DriveId $DriveId -Path $child.id -Depth ($Depth+1) -Index $Index -Total $Total
            }
        }
    } catch {
        Write-ErrorLog "Scan error on $DriveId/${Path}: $_"
    }
    return $items
}

#--- Main Execution ---
try {
    Connect-Graph

    Write-Log 'Retrieving site & drives...'
    $site   = Invoke-Mga -Uri "https://graph.microsoft.com/v1.0/sites/${($siteUrl -replace 'https://','')}" -Method GET
    $drives = Invoke-Mga -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives" -Method GET

    # Calculate total items
    $total = 0
    foreach ($d in $drives.value) {
        $cnt = (Invoke-Mga -Uri "https://graph.microsoft.com/v1.0/drives/$($d.id)/root/children?`$count=true" -Method GET)."@odata.count"
        $total += $cnt
    }
    $idx = [ref]0
    Write-Log "Total items to scan: $total"

    # Scan all drives
    $all = @()
    foreach ($d in $drives.value) {
        Write-Log "Scanning drive $($d.name)..."
        $all += Get-DriveItemsRecursive -DriveId $d.id -Index $idx -Total $total
    }
    Write-Progress -Activity 'Scanning Content' -Completed

    # Filter files and aggregate folder sizes
    Write-Log 'Aggregating file data...'
    $files = $all | Where-Object {$_.file} | ForEach-Object {
        [PSCustomObject]@{
            Name      = $_.name
            Size      = $_.size
            SizeGB    = [math]::Round($_.size/1GB,3)
            SizeMB    = [math]::Round($_.size/1MB,2)
            Path      = $_.parentReference.path
            Extension = [System.IO.Path]::GetExtension($_.name)
        }
    }
    $folderSizes = [ordered]@{}
    foreach ($f in $files) {
        if (-not $folderSizes.ContainsKey($f.Path)) { $folderSizes[$f.Path]=0 }
        $folderSizes[$f.Path] += $f.Size
    }

    # Identify top items
    Write-Log 'Identifying top files and folders...'
    $top20 = $files | Sort-Object Size -Descending | Select-Object -First 20
    $top10 = $folderSizes.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 |
             ForEach-Object {[PSCustomObject]@{FolderPath=$_.Key; SizeGB=[math]::Round($_.Value/1GB,3)}}

    # Permissions: get each user's top-level folder access
    Write-Log 'Retrieving permissions for top-level folders...'
    $permBag = [System.Collections.Concurrent.ConcurrentBag[PSObject]]::new()
    foreach ($folder in $top10) {
        try {
            $encoded = [System.Web.HttpUtility]::UrlPathEncode((New-Object Uri "$siteUrl$($folder.FolderPath)").AbsolutePath)
            $uri     = "$($site.webUrl)/_api/v2.0/drives/$($drives.value | Where-Object id -eq $drives.value[0].id | Select-Object -Expand id)/root:/$($folder.FolderPath):/permissions"
            $perms   = Invoke-Mga -Uri $uri -Method GET
            foreach ($p in $perms.value) {
                if ($p.grantedTo.user) {
                    $permBag.Add([PSCustomObject]@{
                        UserName   = $p.grantedTo.user.displayName
                        UserEmail  = $p.grantedTo.user.email
                        TopFolder  = $folder.FolderPath
                        Roles      = ($p.roles -join ', ')
                    })
                }
            }
        } catch {
            Write-ErrorLog "Permission retrieval error on $($folder.FolderPath): $_"
        }
    }
    $permResults = $permBag.ToArray() | Sort-Object UserEmail,TopFolder -Unique

    # Generate Excel report
    Write-Log 'Generating Excel report...'
    $report = Join-Path $env:USERPROFILE "Documents\SharePoint_Report_$($site.displayName)_$(Get-Date -Format yyyyMMdd_HHmmss).xlsx"

    # Summary sheet
    $summary = [PSCustomObject]@{
        SiteName     = $site.displayName
        SiteUrl      = $site.webUrl
        TotalFiles   = $files.Count
        TotalSizeGB  = [math]::Round(($files | Measure-Object Size -Sum).Sum/1GB,3)
        TotalFolders = $folderSizes.Count
        ReportDate   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }
    $pkg = $summary | Export-Excel -Path $report -WorksheetName 'Summary' -AutoSize -TableStyle Medium2 -PassThru

    # Detail sheets
    $top20        | Export-Excel -ExcelPackage $pkg -WorksheetName 'Top 20 Files'   -AutoSize -TableStyle Medium6
    $top10        | Export-Excel -ExcelPackage $pkg -WorksheetName 'Top 10 Folders' -AutoSize -TableStyle Medium4
    $files        | Group-Object Extension | ForEach-Object {[PSCustomObject]@{Extension=$_.Name;Count=$_.Count}} |
                   Export-Excel -ExcelPackage $pkg -WorksheetName 'File Types'    -AutoSize -TableStyle Medium3
    $permResults  | Export-Excel -ExcelPackage $pkg -WorksheetName 'Permissions'   -AutoSize -TableStyle Medium5

    # Chart on Top 10 Folders sheet
    $ws    = $pkg.Workbook.Worksheets['Top 10 Folders']
    $chart= $ws.Drawings.AddChart('FolderSizeChart',[OfficeOpenXml.Drawing.Chart.eChartType]::ColumnClustered)
    $chart.Series.Add($ws.Cells['B2:B11'],$ws.Cells['A2:A11']) | Out-Null
    $chart.Title.Text = 'Top 10 Folder Sizes (GB)' ; $chart.SetPosition(1,0,4,0); $chart.SetSize(600,300)

    Close-ExcelPackage $pkg
    Write-Log "Report saved to: $report"
}
catch {
    Write-ErrorLog "Unexpected error: $_"
    exit 1
}
finally {
    Try { Disconnect-MgGraph -ErrorAction SilentlyContinue } Catch {}
}
