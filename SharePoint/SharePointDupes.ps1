<#
.SYNOPSIS
  Audits a SharePoint site via Microsoft Graph (app-only auth) for duplicate files,
  largest folders, and largest files. Exports optional Excel report.

.AUTHOR
  Tim MacLatchy

.DATE
  02-06-2025

.LICENSE
  MIT License

.DESCRIPTION
  Connects to a specified SharePoint site using Microsoft Graph app-only auth (client secret),
  recursively retrieves all user–uploaded files from each document library,
  identifies duplicate file names & sizes, top 10 largest folders, and top 20 largest files,
  displays results in the console, and optionally exports to Excel.
#>

#--- Configuration ---
$clientId     = '278b9af9-888d-4344-93bb-769bdd739249'
$tenantId     = 'ca0711e2-e703-4f4e-9099-17d97863211c'
$clientSecret = ''
$siteUrl      = ''

#--- Ensure modules installed & loaded ---
function Install-ModuleIfMissing {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name ${Name})) {
        Install-Module -Name ${Name} -Scope CurrentUser -Force -ErrorAction Stop
    }
    Import-Module -Name ${Name} -Force -ErrorAction Stop
}
Install-ModuleIfMissing 'Microsoft.Graph'
Install-ModuleIfMissing 'ImportExcel'

#--- Logging wrapper ---
function Invoke-LoggedCommand {
    param([string]$Description, [scriptblock]$Action)
    Write-Host "`n[+] ${Description}" -ForegroundColor Cyan
    try {
        & $Action
    } catch {
        Write-Host "[-] Error during '${Description}': $_" -ForegroundColor Red
        throw
    }
}

#--- Authenticate with app-only auth ---
function Connect-Graph {
    Invoke-LoggedCommand 'Authenticating to Microsoft Graph (app-only)' {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Remove-Item "$env:USERPROFILE\.mgcontext" -ErrorAction SilentlyContinue

        Connect-MgGraph `
            -ClientId     ${clientId} `
            -TenantId     ${tenantId} `
            -ClientSecret ${clientSecret} `
            -Scopes       @("https://graph.microsoft.com/.default") `
            -ErrorAction  Stop

        if ((Get-MgContext).AuthType -ne 'AppOnly') {
            throw "Expected AppOnly auth, got '$((Get-MgContext).AuthType)'"
        }
        Write-Host "[+] App-only authentication successful." -ForegroundColor Green
    }
}

#--- Resolve target site ---
function Get-TargetSite {
    param([string]$Url)
    Invoke-LoggedCommand "Resolving site ${Url}" {
        $u = [Uri]${Url}
        $siteId = "$($u.Host):$($u.AbsolutePath.TrimEnd('/'))"
        $site = Get-MgSite -SiteId ${siteId} -ErrorAction Stop
        if (-not $site.Id) { throw "Failed to retrieve Site.Id" }
        Write-Host "[+] Site: $($site.DisplayName) (ID: $($site.Id))" -ForegroundColor Green
        return $site
    }
}

#--- Recursively fetch files from a drive ---
function Get-FilesFromDrive {
    param([string]$DriveId)
    $stack = @('root')
    $files = @()

    Write-Host "`n    → Debug: Starting traversal on DriveId ${DriveId}" -ForegroundColor DarkYellow

    while ($stack.Count -gt 0) {
        $current = $stack[0]; $stack = $stack[1..($stack.Count - 1)]
        Write-Host "        • Debug: Fetching children of folder ID '${current}'" -ForegroundColor DarkGray
        $children = Get-MgDriveItemChild -DriveId ${DriveId} -DriveItemId ${current} -All -ErrorAction Stop

        Write-Host "        • Debug: Retrieved $($children.Count) children" -ForegroundColor DarkGray

        foreach ($item in $children) {
            $pathLower = $item.ParentReference.Path.ToLower()
            if ($item.Folder -and $pathLower -notmatch '/forms$|/siteassets$|/site pages$|/_catalogs|/_layouts') {
                Write-Host "            • Debug: Queuing folder '$($item.Name)' for traversal" -ForegroundColor DarkGray
                $stack += $item.Id
            }
            elseif ($item.File) {
                $ext = ([IO.Path]::GetExtension($item.Name)).ToLower()
                if ($ext -notin '.aspx','.js','.css','.master','.html','.xml') {
                    Write-Host "            • Debug: Adding file '$($item.Name)' (Size: $($item.Size) bytes)" -ForegroundColor DarkGray
                    $files += [PSCustomObject]@{
                        Name    = $item.Name
                        Path    = $item.ParentReference.Path
                        Size    = $item.Size
                        DriveId = $DriveId
                    }
                } else {
                    Write-Host "            • Debug: Skipping system or non-user file '$($item.Name)'" -ForegroundColor DarkGray
                }
            }
        }
    }

    Write-Host "    → Debug: Completed traversal on DriveId ${DriveId}, found $($files.Count) files" -ForegroundColor DarkYellow
    return $files
}

#--- Gather all files ---
function Get-AllFiles {
    param([Microsoft.Graph.PowerShell.Models.IMicrosoftGraphSite]$Site)
    Invoke-LoggedCommand 'Gathering all user-uploaded files' {
        if (-not $Site.Id) { throw "Site.Id is not set" }

        $drives = Get-MgSiteDrive -SiteId $Site.Id -ErrorAction Stop
        Write-Host "[+] Debug: Found $($drives.Count) drives for site '$($Site.DisplayName)'" -ForegroundColor DarkYellow

        $results = @()
        $i = 0; $total = $drives.Count

        foreach ($drive in $drives) {
            $i++
            Write-Host "`n    → Debug: Scanning Drive $i of ${total}: '$($drive.Name)' (ID: $($drive.Id))" -ForegroundColor DarkYellow
            Write-Progress -Activity 'Enumerating Drives' -Status "Drive $i of ${total}: $($drive.Name)" -PercentComplete ([math]::Round(($i / $total) * 100, 1))
            $results += Get-FilesFromDrive -DriveId $drive.Id
        }

        Write-Progress -Activity 'Enumerating Drives' -Completed
        Write-Host "[+] Debug: Total user-uploaded files gathered: $($results.Count)" -ForegroundColor DarkYellow
        return $results
    }
}

#--- Analysis functions ---
function Find-Duplicates {
    param([array]$Files)
    $Files | Group-Object Name, Size | Where-Object Count -gt 1 | ForEach-Object { $_.Group }
}

function Find-TopFolders {
    param([array]$Files)
    $Files | Group-Object Path | ForEach-Object {
        [PSCustomObject]@{
            Folder    = $_.Name
            FileCount = $_.Count
            TotalMB   = [math]::Round(( $_.Group | Measure-Object Size -Sum ).Sum / 1MB, 2)
        }
    } | Sort-Object TotalMB -Descending | Select-Object -First 10
}

function Find-TopFiles {
    param([array]$Files)
    $Files | Sort-Object Size -Descending | Select-Object -First 20 -Property Name, @{n='SizeMB';e={[math]::Round($_.Size / 1MB,2)}}, Path
}

#--- Export to Excel ---
function Export-ExcelReport {
    param([array]$Data)
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = [System.Windows.Forms.SaveFileDialog]::new()
    $dlg.Filter = 'Excel Workbook (*.xlsx)|*.xlsx'
    $dlg.FileName = "Marketing_Audit_$(Get-Date -Format 'dd-MM-yyyy').xlsx"
    if ($dlg.ShowDialog() -ne 'OK') { return }

    $dupes   = Find-Duplicates -Files $Data
    $folders = Find-TopFolders -Files $Data
    $tfiles  = Find-TopFiles -Files $Data

    $Data    | Export-Excel -Path $dlg.FileName -WorksheetName 'AllFiles'   -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WrapText
    $dupes   | Export-Excel -Path $dlg.FileName -WorksheetName 'Duplicates' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WrapText
    $folders | Export-Excel -Path $dlg.FileName -WorksheetName 'TopFolders' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WrapText
    $tfiles  | Export-Excel -Path $dlg.FileName -WorksheetName 'TopFiles'   -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WrapText

    Add-ExcelChart -ExcelPackage $dlg.FileName `
        -WorksheetName 'TopFolders' `
        -ChartType PieExploded3D `
        -Title "Top 10 Folders (MB)" `
        -XRange "A2:A$($folders.Count + 1)" `
        -YRange "C2:C$($folders.Count + 1)"

    Write-Host "`n[+] Excel report saved to: $($dlg.FileName)" -ForegroundColor Green
}

#--- Main ---
function Main {
    Connect-Graph
    $site = Get-TargetSite -Url $siteUrl
    $allFiles = Get-AllFiles -Site $site

    # Console output
    Write-Host "`n=== Summary for $($site.DisplayName) ===" -ForegroundColor Yellow
    Write-Host "Total files   : $($allFiles.Count)"
    Write-Host "Duplicates    : $((Find-Duplicates -Files $allFiles).Count)"
    Write-Host "`nTop 10 Folders:"; Find-TopFolders -Files $allFiles | Format-Table -AutoSize
    Write-Host "`nTop 20 Files:";   Find-TopFiles   -Files $allFiles | Format-Table -AutoSize

    if ((Read-Host "`nExport to Excel? (y/n)") -eq 'y') {
        Export-ExcelReport -Data $allFiles
    }

    Write-Host "`n[+] Audit complete." -ForegroundColor Magenta
}

#--- Run ---
Main
