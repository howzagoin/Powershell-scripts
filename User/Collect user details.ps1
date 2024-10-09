# Import necessary modules
Import-Module Az
Import-Module AzureAD
Import-Module ExchangeOnlineManagement
Import-Module PnP.PowerShell

# Connect to Azure interactively
Connect-AzAccount

# After logging in, list the tenants
$tenants = Get-AzTenant

# Display the tenants
$tenants

# Select a tenant to work with (prompt for tenant ID or choose the first one)
$tenantId = Read-Host -Prompt "Enter Tenant ID (or leave blank to use the first tenant)"
if (-not $tenantId) {
    $tenantId = $tenants[0].TenantId
}

# Set the context to the selected tenant
Set-AzContext -TenantId $tenantId

# Connect to Azure AD
Connect-AzureAD -TenantId $tenantId

# Connect to Exchange Online interactively
Connect-ExchangeOnline -UserPrincipalName "<YourUPN>" -ShowProgress $true

# Connect to PnP Online interactively
Connect-PnPOnline -Url "https://<YourTenant>-admin.sharepoint.com" -Interactive

# Prompt for user first name to search
$firstName = Read-Host -Prompt
