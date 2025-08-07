# Debug script to check what sites are being found and why they're filtered
param()

# Import required modules
Import-Module Microsoft.Graph.Sites -Force

# Connect to Graph
Connect-MgGraph -AppId "278b9af9-888d-4344-93bb-769bdd739249" -CertificateThumbprint "2E2502BB1EDB8F36CF9DE50936B283BDD22D5BAD" -TenantId "ca0711e2-e703-4f4e-9099-17d97863211c" -NoWelcome

Write-Host "=== DEBUGGING SITE ENUMERATION ===" -ForegroundColor Yellow

$sites = @()

# Approach 1: Get root site
Write-Host "`n1. Root Site:" -ForegroundColor Cyan
try {
    $rootSite = Get-MgSite -SiteId "root" -ErrorAction SilentlyContinue
    if ($rootSite) {
        $sites += $rootSite
        Write-Host "   Found: $($rootSite.DisplayName) - $($rootSite.WebUrl)"
    }
} catch {
    Write-Host "   Error: $($_)" -ForegroundColor Red
}

# Approach 2: Graph API
Write-Host "`n2. Graph API Sites:" -ForegroundColor Cyan
try {
    $allSites = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites?`$top=500" -ErrorAction SilentlyContinue
    if ($allSites.value) {
        Write-Host "   Found $($allSites.value.Count) sites via Graph API"
        $sites += $allSites.value
        Write-Host "   First 5 sites from Graph API:"
        $allSites.value | Select-Object -First 5 | ForEach-Object {
            Write-Host "      $($_.displayName) - $($_.webUrl)"
        }
    }
} catch {
    Write-Host "   Error: $($_)" -ForegroundColor Red
}

# Approach 3: Search
Write-Host "`n3. Search Sites:" -ForegroundColor Cyan
try {
    $searchSites = Get-MgSite -Search "*" -All -ErrorAction SilentlyContinue
    if ($searchSites) {
        Write-Host "   Found $($searchSites.Count) sites via search"
        $sites += $searchSites
    }
} catch {
    Write-Host "   Error: $($_)" -ForegroundColor Red
}

Write-Host "`n=== BEFORE DEDUPLICATION ===" -ForegroundColor Yellow
Write-Host "Total sites collected: $($sites.Count)"

# Check first few sites from Graph API to see their properties
Write-Host "`nInvestigating first 3 Graph API sites:"
$sites | Where-Object { $_.webUrl -like "*sharepoint.com/sites/*" } | Select-Object -First 3 | ForEach-Object {
    Write-Host "  Site: $($_.displayName)"
    Write-Host "    ID: $($_.Id)"
    Write-Host "    DisplayName: $($_.DisplayName)"
    Write-Host "    displayName: $($_.displayName)"
    Write-Host "    WebUrl: $($_.WebUrl)"
    Write-Host "    webUrl: $($_.webUrl)"
    Write-Host "    HasId: $($_.Id -ne $null)"
    Write-Host "    HasDisplayName: $(($_.DisplayName -ne $null) -or ($_.displayName -ne $null))"
    Write-Host "    ---"
}

# Remove duplicates
$uniqueSites = $sites | Where-Object { 
    $_ -and $_.Id -and (
        $_.DisplayName -or $_.displayName
    ) 
} | Sort-Object Id -Unique

# Normalize property names to ensure consistency
$uniqueSites = $uniqueSites | ForEach-Object {
    if (-not $_.DisplayName -and $_.displayName) {
        $_ | Add-Member -NotePropertyName 'DisplayName' -NotePropertyValue $_.displayName -Force
    }
    if (-not $_.WebUrl -and $_.webUrl) {
        $_ | Add-Member -NotePropertyName 'WebUrl' -NotePropertyValue $_.webUrl -Force
    }
    $_
}

Write-Host "`n=== AFTER DEDUPLICATION ===" -ForegroundColor Yellow
Write-Host "Unique sites: $($uniqueSites.Count)"

# Show first 10 unique sites
Write-Host "`nFirst 10 unique sites:"
$uniqueSites | Select-Object -First 10 | ForEach-Object {
    Write-Host "   $($_.DisplayName) - $($_.WebUrl) - OneDrive: $($_.WebUrl -like '*-my.sharepoint.com*')"
}

# Filter out OneDrive sites
$sharePointSites = $uniqueSites | Where-Object {
    -not ($_.WebUrl -like "*-my.sharepoint.com/personal/*" -or 
          $_.WebUrl -like "*/personal/*" -or 
          $_.WebUrl -like "*mysites*" -or 
          $_.DisplayName -like "*OneDrive*" -or
          ($_.WebUrl -and $_.WebUrl -match "onedrive"))
}

Write-Host "`n=== AFTER ONEDRIVE FILTERING ===" -ForegroundColor Yellow
Write-Host "SharePoint sites: $($sharePointSites.Count)"

Write-Host "`nSharePoint sites:"
$sharePointSites | Select-Object -First 10 | ForEach-Object {
    Write-Host "   $($_.DisplayName) - $($_.WebUrl)"
}

Disconnect-MgGraph
