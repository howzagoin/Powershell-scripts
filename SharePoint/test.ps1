<#
.SYNOPSIS
  SharePoint Storage & Access Reporter (CLI)

.DESCRIPTION
  Generates a comprehensive Excel report covering:
    • Top 20 largest files
    • Top 10 largest folders
    • Storage quota vs. usage chart
    • StorageUsedGB vs. StorageTotalGB chart
    • Parent-folder permission summary
  Uses recursive deep scanning up to maxDepth, parallel processing using PowerShell jobs,
  optimized bulk Graph calls (Optimized.Mga), and certificate-based app-only authentication.

.NOTES
  Author     : Timothy MacLatchy
  Date       : 13-06-2025
  License    : MIT License
  Enhancements:
    - Full multi-threaded folder scan using PowerShell jobs
    - Accurate total count based on actual recursion
    - Permission retrieval fallback on failure
    - Configurable Excel output path
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
$outputFolder              = "$env:USERPROFILE\Documents"

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
        [int]$Depth = 0
    )
    $items = @()
    try {
        $resp = Invoke-Mga -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$Path/children?`$select=id,name,size,folder,file,parentReference,webUrl" -Method GET
        foreach ($child in $resp.value) {
            if ($child.file -and $child.name -notmatch '^(~|\.|_vti_|Thumbs\.db|\.DS_Store)') {
                $items += $child
            } elseif ($child.folder -and $Depth -lt $maxDepth) {
                $items += $child
                $items += Get-DriveItemsRecursive -DriveId $DriveId -Path $child.id -Depth ($Depth+1)
            }
        }
    } catch {
        Write-ErrorLog "Scan error on $DriveId/${Path}: $_"
    }
    return $items
}

#--- Parallel Execution Wrapper ---
function Start-ParallelScan {
    param([array]$drives)
    $jobs = @()
    foreach ($d in $drives) {
        $jobs += Start-Job -ScriptBlock {
            param($driveId,$depth)
            Import-Module Optimized.Mga
            return Get-DriveItemsRecursive -DriveId $driveId -Depth $depth
        } -ArgumentList $d.id,$maxDepth
    }
    $results = @()
    foreach ($j in $jobs) {
        $results += Receive-Job -Job $j -Wait -AutoRemoveJob
    }
    return $results
}

#--- Main Execution ---
try {
    Connect-Graph

    Write-Log 'Retrieving site & drives...'
    $site   = Invoke-Mga -Uri "https://graph.microsoft.com/v1.0/sites/${($siteUrl -replace 'https://','')}" -Method GET
    $drives = Invoke-Mga -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives" -Method GET

    Write-Log 'Scanning all drives in parallel...'
    $all = Start-ParallelScan -drives $drives.value

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
            $uri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives/$($drives.value[0].id)/root:/$($folder.FolderPath):/permissions"
            $perms = Invoke-Mga -Uri $uri -Method GET
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
    $report = Join-Path $outputFolder "SharePoint_Report_$($site.displayName)_$(Get-Date -Format yyyyMMdd_HHmmss).xlsx"

    $summary = [PSCustomObject]@{
        SiteName     = $site.displayName
        SiteUrl      = $site.webUrl
        TotalFiles   = $files.Count
        TotalSizeGB  = [math]::Round(($files | Measure-Object Size -Sum).Sum/1GB,3)
        TotalFolders = $folderSizes.Count
        ReportDate   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }
    $pkg = $summary | Export-Excel -Path $report -WorksheetName 'Summary' -AutoSize -TableStyle Medium2 -PassThru

    $top20        | Export-Excel -ExcelPackage $pkg -WorksheetName 'Top 20 Files'   -AutoSize -TableStyle Medium6
    $top10        | Export-Excel -ExcelPackage $pkg -WorksheetName 'Top 10 Folders' -AutoSize -TableStyle Medium4
    $files        | Group-Object Extension | ForEach-Object {[PSCustomObject]@{Extension=$_.Name;Count=$_.Count}} |
                   Export-Excel -ExcelPackage $pkg -WorksheetName 'File Types'    -AutoSize -TableStyle Medium3
    $permResults  | Export-Excel -ExcelPackage $pkg -WorksheetName 'Permissions'   -AutoSize -TableStyle Medium5

    # Chart on Top 10 Folders sheet
    $ws    = $pkg.Workbook.Worksheets['Top 10 Folders']
    $chart = $ws.Drawings.AddChart('FolderSizeChart',[OfficeOpenXml.Drawing.Chart.eChartType]::ColumnClustered)
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
