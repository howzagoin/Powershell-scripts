# Quick test script to check user access function
param()

# Import required modules
Import-Module Microsoft.Graph.Sites -Force
Import-Module Microsoft.Graph.Identity.DirectoryManagement -Force

# Connect to Graph
Connect-MgGraph -AppId "278b9af9-888d-4344-93bb-769bdd739249" -CertificateThumbprint "2E2502BB1EDB8F36CF9DE50936B283BDD22D5BAD" -TenantId "ca0711e2-e703-4f4e-9099-17d97863211c"

# Get the Communication site
$sites = Get-MgSite -All | Where-Object { $_.DisplayName -eq "Communication site" }
$site = $sites[0]

Write-Host "Testing User Access for: $($site.DisplayName)"
Write-Host "Site ID: $($site.Id)"
Write-Host "Site URL: $($site.WebUrl)"

# Test the Get-SiteUserAccessSummary function logic manually
$owners = @()
$members = @()
$externalGuests = @()

Write-Host "`nTesting fallback user creation..."

try {
    $siteInfo = Get-MgSite -SiteId $site.Id -ErrorAction SilentlyContinue
    Write-Host "Site Info Retrieved: $($siteInfo -ne $null)"
    
    if ($siteInfo -and $siteInfo.CreatedBy) {
        Write-Host "CreatedBy found: $($siteInfo.CreatedBy.User.DisplayName)"
        $owners += [PSCustomObject]@{
            DisplayName = if ($siteInfo.CreatedBy.User.DisplayName) { $siteInfo.CreatedBy.User.DisplayName } else { "Site Creator" }
            UserEmail = if ($siteInfo.CreatedBy.User.Email) { $siteInfo.CreatedBy.User.Email } else { "unknown@domain.com" }
            UserType = "Internal"
            Role = "Owner (Creator)"
        }
    } else {
        Write-Host "No CreatedBy info, using fallback"
        # Fallback: Create generic owner entry
        $owners += [PSCustomObject]@{
            DisplayName = "Site Owners"
            UserEmail = "site-owners@" + ($site.WebUrl -replace "https://", "" -replace "/.*", "")
            UserType = "Internal"
            Role = "Owner"
        }
    }
} catch {
    Write-Host "Error in fallback: $($_.Exception.Message)"
    # Final fallback
    $owners += [PSCustomObject]@{
        DisplayName = "Site Owners"
        UserEmail = "site-owners@" + ($site.WebUrl -replace "https://", "" -replace "/.*", "")
        UserType = "Internal"
        Role = "Owner"
    }
}

Write-Host "`nFinal Results:"
Write-Host "Owners Count: $($owners.Count)"
$owners | ForEach-Object { Write-Host "  - $($_.DisplayName) ($($_.UserEmail)) - $($_.Role)" }

# Disconnect
Disconnect-MgGraph
