# Tenant Audit Script

# Assuming the necessary modules are imported and users are retrieved.

$users = Get-MsolUser

# 1. Gathering User Roles
foreach ($user in $users) {
    $roles = Get-MsolUserRole -UserPrincipalName $user.UserPrincipalName
    if ($roles -ne $null) {
        $user.Roles = $roles | ForEach-Object { $_.RoleName } -join '; '
    } else {
        $user.Roles = "None" # Default to "None" if no roles are found
    }
}

# 2. Gathering Delegates for Mailboxes and Calendars
foreach ($user in $users) {
    try {
        $mailDelegate = Get-MailboxPermission -Identity $user.UserPrincipalName | Where-Object { $_.AccessRights -contains 'FullAccess' }
        $calendarDelegate = Get-MailboxFolderPermission -Identity "$($user.UserPrincipalName):\Calendar" | Where-Object { $_.AccessRights -ne 'None' }

        $user.MailboxDelegates = $mailDelegate | ForEach-Object { $_.User } -join '; '
        $user.CalendarDelegates = $calendarDelegate | ForEach-Object { $_.User } -join '; '
    } catch {
        Write-Error "Error retrieving delegates for user: $($user.UserPrincipalName) - $_" # Log the error
    }
}

# 3. Gathering Enterprise Apps and their Members
$enterpriseApps = Get-AzureADServicePrincipal
foreach ($app in $enterpriseApps) {
    $members = Get-AzureADServiceAppRoleAssignment -ObjectId $app.ObjectId
    $app.Members = $members | ForEach-Object { $_.PrincipalDisplayName } -join '; '
}

# 4. Unhandled Errors for Users
foreach ($user in $users) {
    if ($user -eq $null) {
        Write-Error "User is null - Check user input data or prior commands."
        # Add logging to errors tab if necessary
    }
}