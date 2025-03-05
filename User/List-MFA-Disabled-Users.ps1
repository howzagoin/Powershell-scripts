# Define the path to the CSV file
$csvFilePath = "GranularAdministerRelationship.csv"

# Import the CSV file
$tenants = Import-Csv -Path $csvFilePath

# Prepare a list to store users with MFA disabled across all tenants
$allMfaDisabledUsers = @()

# Connect to Partner Center
Connect-PartnerCenter

# Iterate through each tenant
foreach ($tenant in $tenants) {
    $tenantId = $tenant."Microsoft ID"

    # Get the tenant's context
    $customer = Get-PartnerCustomer -CustomerId $tenantId
    $context = New-PartnerCustomerContext -Customer $customer

    # Get all users from Azure AD
    $users = Get-AzureADUser -All $true -TenantId $tenantId -PartnerContext $context

    # Prepare a list to store users with MFA disabled for the current tenant
    $mfaDisabledUsers = @()

    foreach ($user in $users) {
        # Get MFA status for the user
        $mfaStatus = Get-MsolUser -UserPrincipalName $user.UserPrincipalName -TenantId $tenantId | Select-Object -ExpandProperty StrongAuthenticationRequirements

        # Check if MFA is disabled
        if ($mfaStatus.Count -eq 0) {
            $mfaDisabledUsers += [PSCustomObject]@{
                TenantID          = $tenantId
                UserPrincipalName = $user.UserPrincipalName
                DisplayName       = $user.DisplayName
            }
        }
    }

    # Add the users with MFA disabled from the current tenant to the overall list
    $allMfaDisabledUsers += $mfaDisabledUsers
}

# Display all users with MFA disabled
$allMfaDisabledUsers | Format-Table -AutoSize

# Optionally, export the list to a CSV file
$allMfaDisabledUsers | Export-Csv -Path "C:\All_Tenants_MFA_Disabled_Users.csv" -NoTypeInformation
