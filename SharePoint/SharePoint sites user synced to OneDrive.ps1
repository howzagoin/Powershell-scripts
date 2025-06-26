# Import necessary modules
Import-Module Microsoft.Online.SharePoint.PowerShell
Import-Module PartnerCenter

# Prompt user for a username
$userEmail = Read-Host -Prompt "Enter the user's email address"

# Extract the domain from the email address
$domain = $userEmail.Split('@')[1]

# Connect to Partner Center or Microsoft Lighthouse to find the tenant
# (Assuming you have already authenticated to Partner Center or Lighthouse)
$customer = Get-PartnerCustomer -Domain $domain

if ($customer) {
    $tenantId = $customer.CustomerId
    $adminSiteUrl = "https://$($customer.Domain)-admin.sharepoint.com"

    # Connect to SharePoint Online admin center for the tenant
    Connect-SPOService -Url $adminSiteUrl

    # Retrieve the OneDrive URL for the user
    $oneDriveUrl = (Get-SPOSite -Identity "https://$($customer.Domain)-my.sharepoint.com/personal/$($userEmail.Replace('@', '_').Replace('.', '_'))").Url

    # Get the SharePoint sites synced to the user's OneDrive
    $oneDriveSyncClientPath = "$env:LocalAppData\Microsoft\OneDrive\settings\Business1"

    # Parse the .ini files in the OneDrive sync client directory
    $sharePointSites = Get-ChildItem -Path $oneDriveSyncClientPath -Filter "*.ini" | ForEach-Object {
        Select-String -Path $_.FullName -Pattern "LibraryPath" | ForEach-Object {
            $_ -replace ".*LibraryPath=|;.*", ""
        }
    }

    # Output the SharePoint sites
    Write-Output "SharePoint sites synced to $userEmail's OneDrive:"
    $sharePointSites | ForEach-Object {
        Write-Output $_
    }

    # Disconnect from SharePoint Online
    Disconnect-SPOService
} else {
    Write-Error "Tenant not found for the domain $domain"
}
