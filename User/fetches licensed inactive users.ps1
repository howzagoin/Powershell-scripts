# Connect to Azure account
Connect-AzAccount

# Get all delegated tenants
$tenants = Get-AzManagedServicesAssignment | Select-Object -ExpandProperty TenantId -Unique

# Define the inactive period (e.g., 90 days)
$inactivePeriod = (Get-Date).AddDays(-90)

# Array to hold inactive users across all tenants
$allInactiveUsers = @()

foreach ($tenantId in $tenants) {
    # Set context to the tenant
    Set-AzContext -TenantId $tenantId

    # Connect to Microsoft Graph for the current tenant
    Connect-MgGraph -TenantId $tenantId -Scopes "User.Read.All"

    # Fetch all users
    $users = Get-MgUser -All

    # Array to hold inactive users for the current tenant
    $inactiveUsers = @()

    foreach ($user in $users) {
        # Get the user's last sign-in activity
        $signInActivity = Get-MgUserAuthenticationMethodSignInActivity -UserId $user.Id -ErrorAction SilentlyContinue

        # Check if the user is inactive
        if ($null -ne $signInActivity) {
            if ($signInActivity.LastSignInDateTime -lt $inactivePeriod) {
                # Get user's license status
                $licenseDetails = Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction SilentlyContinue

                $userDetails = [PSCustomObject]@{
                    TenantId          = $tenantId
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName       = $user.DisplayName
                    LastSignIn        = $signInActivity.LastSignInDateTime
                    Licenses          = $licenseDetails.SkuPartNumber -join ", "
                }

                # Add to the array
                $inactiveUsers += $userDetails
            }
        } else {
            # If no sign-in activity is found, consider the user as inactive
            $licenseDetails = Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction SilentlyContinue

            $userDetails = [PSCustomObject]@{
                TenantId          = $tenantId
                UserPrincipalName = $user.UserPrincipalName
                DisplayName       = $user.DisplayName
                LastSignIn        = "No sign-in activity found"
                Licenses          = $licenseDetails.SkuPartNumber -join ", "
            }

            # Add to the array
            $inactiveUsers += $userDetails
        }
    }

    # Add current tenant's inactive users to the global array
    $allInactiveUsers += $inactiveUsers

    # Disconnect from Microsoft Graph
    Disconnect-MgGraph
}

# Output all inactive users across tenants
$allInactiveUsers | Format-Table -AutoSize

# Optionally, export to CSV
$allInactiveUsers | Export-Csv -Path "AllTenantsInactiveUsers.csv" -NoTypeInformation
