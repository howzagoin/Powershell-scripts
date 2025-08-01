# SharePoint File Recovery Tool - Universal Search and Restore
# Configuration (update these with your tenant details)
$clientId              = '278b9af9-888d-4344-93bb-769bdd739249'
$tenantId              = 'ca0711e2-e703-4f4e-9099-17d97863211c'
$certificateThumbprint = '2E2502BB1EDB8F36CF9DE50936B283BDD22D5BAD'

# --- USER INPUT SECTION ---
Write-Host "=== SharePoint File Recovery Tool ===" -ForegroundColor Magenta
Write-Host "This tool will search for lost files/folders based on keywords you provide" -ForegroundColor Cyan
Write-Host ""

# Get SharePoint site URL from user
Write-Host "SharePoint Site Selection:" -ForegroundColor Cyan
Write-Host "Enter the SharePoint site URL you want to search" -ForegroundColor White
Write-Host "Example:" -ForegroundColor Gray
Write-Host "  https://company.sharepoint.com/sites/SiteName" -ForegroundColor Gray
Write-Host ""

do {
    $SiteURL = Read-Host "SharePoint Site URL"
    if (-not $SiteURL -or $SiteURL.Trim() -eq "") {
        Write-Host "Site URL is required. Please enter a valid SharePoint site URL." -ForegroundColor Red
    } elseif (-not $SiteURL.StartsWith("https://")) {
        Write-Host "Please enter a complete URL starting with https://" -ForegroundColor Red
        $SiteURL = ""
    }
} while (-not $SiteURL -or $SiteURL.Trim() -eq "")

$SiteURL = $SiteURL.Trim()
Write-Host "Will search site: $SiteURL" -ForegroundColor Green
Write-Host ""

# Get search keywords from user
Write-Host "Search Keywords:" -ForegroundColor Cyan
Write-Host "Enter keywords to search for (separate multiple keywords with commas)" -ForegroundColor White
Write-Host "Examples:" -ForegroundColor Gray
Write-Host "  Bunnings, Statement, BUNN001" -ForegroundColor Gray
Write-Host "  Mode 2.0, photography, ezy storage" -ForegroundColor Gray
Write-Host "  invoice, customer, July 2025" -ForegroundColor Gray
Write-Host ""

do {
    $keywordInput = Read-Host "Enter keywords (comma-separated)"
    if (-not $keywordInput -or $keywordInput.Trim() -eq "") {
        Write-Host "Keywords are required. Please enter at least one keyword." -ForegroundColor Red
    }
} while (-not $keywordInput -or $keywordInput.Trim() -eq "")

# Parse and clean keywords
$searchKeywords = $keywordInput.Split(',') | ForEach-Object { 
    $_.Trim() 
} | Where-Object { 
    $_ -ne "" 
}

Write-Host "Will search for files containing any of these keywords:" -ForegroundColor Green
$searchKeywords | ForEach-Object { Write-Host "  • $_" -ForegroundColor White }
Write-Host ""

# Get file extensions to search for
Write-Host ""
Write-Host "File type options:" -ForegroundColor Cyan
Write-Host "1. All file types"
Write-Host "2. Images only (.jpg, .png, .gif, etc.)"
Write-Host "3. Documents only (.docx, .pdf, .xlsx, etc.)"
Write-Host "4. Custom extensions"

$fileTypeChoice = Read-Host "Select option (1-4)"

switch ($fileTypeChoice) {
    "1" { 
        $fileExtensions = @() # Empty array means all files
        Write-Host "Will search all file types" -ForegroundColor Green
    }
    "2" { 
        $fileExtensions = @('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', '.heic', '.heif', '.webp', '.svg')
        Write-Host "Will search image files only" -ForegroundColor Green
    }
    "3" { 
        $fileExtensions = @('.docx', '.doc', '.pdf', '.xlsx', '.xls', '.pptx', '.ppt', '.txt', '.rtf')
        Write-Host "Will search document files only" -ForegroundColor Green
    }
    "4" { 
        $customExtensions = Read-Host "Enter file extensions separated by commas (e.g., .xlsx,.pdf,.docx)"
        $fileExtensions = $customExtensions.Split(',') | ForEach-Object { $_.Trim() }
        Write-Host "Will search for: $($fileExtensions -join ', ')" -ForegroundColor Green
    }
    default { 
        $fileExtensions = @()
        Write-Host "Invalid selection. Will search all file types" -ForegroundColor Yellow
    }
}

# --- Connection ---
Write-Host ""
Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteURL -ClientId $clientId -Tenant $tenantId -Thumbprint $certificateThumbprint

$web = Get-PnPWeb
if ($web) {
    Write-Host "Successfully connected to $($web.Title)" -ForegroundColor Green
} else {
    Write-Host "Failed to connect to SharePoint." -ForegroundColor Red
    exit
}

# --- Function to ensure folder exists ---
function Ensure-Folder {
    param([string]$FolderSiteRelativeUrl)
    
    try {
        $folder = Get-PnPFolder -Url $FolderSiteRelativeUrl -ErrorAction SilentlyContinue
        if (-not $folder) {
            Write-Host "Creating folder: $FolderSiteRelativeUrl" -ForegroundColor Yellow
            
            $pathParts = $FolderSiteRelativeUrl.Split('/')
            $currentPath = $pathParts[0]
            
            for ($i = 1; $i -lt $pathParts.Length; $i++) {
                $nextPath = "$currentPath/$($pathParts[$i])"
                $testFolder = Get-PnPFolder -Url $nextPath -ErrorAction SilentlyContinue
                
                if (-not $testFolder) {
                    Add-PnPFolder -Name $pathParts[$i] -Folder $currentPath -ErrorAction SilentlyContinue | Out-Null
                }
                $currentPath = $nextPath
            }
            
            $folder = Get-PnPFolder -Url $FolderSiteRelativeUrl -ErrorAction SilentlyContinue
        }
        
        if ($folder) {
            Write-Host "Folder ready: $FolderSiteRelativeUrl" -ForegroundColor Green
        } else {
            # Fallback to root Shared Documents
            $fallbackName = "File Recovery $(Get-Date -Format 'yyyy-MM-dd HH-mm')"
            $folder = Add-PnPFolder -Name $fallbackName -Folder "Shared Documents" -ErrorAction SilentlyContinue
            Write-Host "Created fallback folder: Shared Documents/$fallbackName" -ForegroundColor Yellow
        }
        return $folder
    }
    catch {
        Write-Host "Error with folder: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# --- Build search queries ---
function Build-SearchQuery {
    param([array]$Keywords, [array]$Extensions)
    
    # Build keyword part
    $keywordQuery = ($Keywords | ForEach-Object { "Title:*$_* OR Filename:*$_*" }) -join " OR "
    
    # Build extension part if specified
    if ($Extensions -and $Extensions.Count -gt 0) {
        $extQuery = ($Extensions | ForEach-Object { "FileExtension:$($_.TrimStart('.'))" }) -join " OR "
        return "($keywordQuery) AND ($extQuery) AND IsDocument:true"
    } else {
        return "($keywordQuery) AND IsDocument:true"
    }
}

# Create recovery folders
$timestamp = Get-Date -Format 'yyyy-MM-dd HH-mm'
$recycleBinFolder = "Shared Documents/Recovery - Recycle Bin - $timestamp"
$siteSearchFolder = "Shared Documents/Recovery - Site Search - $timestamp"

$recycleBinRecoveryFolder = Ensure-Folder -FolderSiteRelativeUrl $recycleBinFolder
$siteSearchRecoveryFolder = Ensure-Folder -FolderSiteRelativeUrl $siteSearchFolder

# --- STEP 1: Search existing files on the site ---
Write-Host ""
Write-Host "=== STEP 1: Searching existing files on the site ===" -ForegroundColor Magenta

$siteSearchQuery = Build-SearchQuery -Keywords $searchKeywords -Extensions $fileExtensions
Write-Host "Search query: $siteSearchQuery" -ForegroundColor Gray

$existingFiles = Submit-PnPSearchQuery -Query $siteSearchQuery -All -ErrorAction SilentlyContinue

if ($existingFiles -and $existingFiles.ResultRows.Count -gt 0) {
    Write-Host "Found $($existingFiles.ResultRows.Count) matching files on the site!" -ForegroundColor Green
    
    $copiedCount = 0
    foreach ($file in $existingFiles.ResultRows) {
        $fileName = $file.Title
        if (-not $fileName) { $fileName = Split-Path $file.Path -Leaf }
        $filePath = $file.Path
        
        Write-Host "---"
        Write-Host "Found: $fileName" -ForegroundColor White
        Write-Host "Location: $filePath" -ForegroundColor Gray
        
        try {
            $sourceRelativeUrl = $filePath -replace "https://[^/]+/sites/[^/]+/", ""
            $fileStream = Get-PnPFile -Url $sourceRelativeUrl -AsMemoryStream -ErrorAction SilentlyContinue
            
            if ($fileStream -and $siteSearchRecoveryFolder) {
                $newFile = Add-PnPFile -FileName $fileName -Folder $siteSearchFolder -Stream $fileStream -ErrorAction SilentlyContinue
                if ($newFile) {
                    Write-Host "✓ Copied to recovery folder" -ForegroundColor Green
                    $copiedCount++
                } else {
                    Write-Host "✗ Failed to copy" -ForegroundColor Red
                }
                $fileStream.Dispose()
            }
        }
        catch {
            Write-Host "✗ Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Write-Host ""
    Write-Host "Copied $copiedCount files to: $siteSearchFolder" -ForegroundColor Green
} else {
    Write-Host "No matching files found on the site" -ForegroundColor Yellow
}

# --- STEP 2: Search Recycle Bin ---
Write-Host ""
Write-Host "=== STEP 2: Searching Recycle Bin ===" -ForegroundColor Magenta

$allDeletedItems = Get-PnPRecycleBinItem
Write-Host "Found $($allDeletedItems.Count) items in recycle bin. Filtering..." -ForegroundColor Gray

# Filter deleted items based on keywords and extensions
$matchingDeletedItems = $allDeletedItems | Where-Object {
    $item = $_
    $matchesKeyword = $false
    $matchesExtension = $true
    
    # Check if filename matches any keyword
    foreach ($keyword in $searchKeywords) {
        if ($item.LeafName -like "*$keyword*") {
            $matchesKeyword = $true
            break
        }
    }
    
    # Check extension if specified
    if ($fileExtensions -and $fileExtensions.Count -gt 0) {
        $matchesExtension = $false
        foreach ($ext in $fileExtensions) {
            if ($item.LeafName -like "*$ext") {
                $matchesExtension = $true
                break
            }
        }
    }
    
    $matchesKeyword -and $matchesExtension
}

if ($matchingDeletedItems) {
    Write-Host "Found $($matchingDeletedItems.Count) matching deleted files!" -ForegroundColor Green
    
    $restoredCount = 0
    foreach ($deletedItem in $matchingDeletedItems) {
        $stage = if ($deletedItem.ItemState -eq "SecondStageRecycleBin") { "Second-Stage" } else { "First-Stage" }
        
        Write-Host "---"
        Write-Host "Name: $($deletedItem.LeafName)" -ForegroundColor White
        Write-Host "Original Path: $($deletedItem.DirName)" -ForegroundColor Gray
        Write-Host "Deleted By: $($deletedItem.DeletedByName)" -ForegroundColor Gray
        Write-Host "Deleted Date: $($deletedItem.DeletedDate.ToLocalTime())" -ForegroundColor Gray
        Write-Host "Stage: $stage" -ForegroundColor Gray
        
        try {
            # Restore from recycle bin
            Restore-PnPRecycleBinItem -Identity $deletedItem.Id -Force -ErrorAction Stop
            Write-Host "✓ Restored from recycle bin" -ForegroundColor Green
            
            Start-Sleep -Seconds 1
            
            # Move to recovery folder
            $originalRelativePath = $deletedItem.DirName -replace "^sites/[^/]+/", ""
            $fileName = $deletedItem.LeafName
            $sourceFileUrl = "$originalRelativePath/$fileName"
            
            $restoredFile = Get-PnPFile -Url $sourceFileUrl -AsMemoryStream -ErrorAction SilentlyContinue
            
            if ($restoredFile -and $recycleBinRecoveryFolder) {
                $newFile = Add-PnPFile -FileName $fileName -Folder $recycleBinFolder -Stream $restoredFile -ErrorAction SilentlyContinue
                if ($newFile) {
                    Write-Host "✓ Moved to recovery folder" -ForegroundColor Green
                    Remove-PnPFile -ServerRelativeUrl $sourceFileUrl -Force -ErrorAction SilentlyContinue
                    $restoredCount++
                } else {
                    Write-Host "✗ Failed to move to recovery folder" -ForegroundColor Red
                }
                $restoredFile.Dispose()
            }
        }
        catch {
            Write-Host "✗ Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Write-Host ""
    Write-Host "Restored $restoredCount files to: $recycleBinFolder" -ForegroundColor Green
} else {
    Write-Host "No matching files found in recycle bin" -ForegroundColor Yellow
}

# --- SUMMARY ---
Write-Host ""
Write-Host "=== RECOVERY SUMMARY ===" -ForegroundColor Magenta
Write-Host "Search keywords: $($searchKeywords -join ', ')" -ForegroundColor White
Write-Host "File types: $(if ($fileExtensions.Count -gt 0) { $fileExtensions -join ', ' } else { 'All types' })" -ForegroundColor White
Write-Host ""
Write-Host "Recovery folders created:" -ForegroundColor Cyan
Write-Host "• Files found on site copied to: $siteSearchFolder" -ForegroundColor White
Write-Host "• Files restored from recycle bin to: $recycleBinFolder" -ForegroundColor White
Write-Host ""
Write-Host "Navigate to Shared Documents to find your recovery folders" -ForegroundColor Yellow

Disconnect-PnPOnline
Write-Host "Disconnected from SharePoint." -ForegroundColor Cyan