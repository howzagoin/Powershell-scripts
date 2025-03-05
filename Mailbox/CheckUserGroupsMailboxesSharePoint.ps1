# Author: Tim MacLatchy
# Date: 03 October 2024
# License: MIT License
# Description: This script retrieves mailbox type, delegated mailboxes, security group memberships,
# and SharePoint site access for a specified user.

# Check for required modules and install if missing
$requiredModules = @('Microsoft.Graph')

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module -Name $module -Force
    }
}

# Login to Microsoft Graph using web-based MFA
try {
    Write-Host "Signing into tenant with web MFA..."
    Connect-MgGraph -Scopes "User.Read.All", "Mail.Read.Shared", "GroupMember.Read.All"
} catch {
    Write-Error "Failed to sign in: $_"
    exit
}

Write-Host "Login successful. Continuing the script..."

# Prompt for user's email address
$UserPrincipalName = Read-Host "Enter the user's email address"

# Getting user information
try {
    $user = Get-MgUser -UserId $UserPrincipalName
    Write-Host "User information retrieved for: $($user.DisplayName)"
} catch {
    Write-Error "Failed to retrieve user information for $UserPrincipalName: $_"
    exit
}

# Getting mailbox type
$mailboxType = if ($user.MailboxSettings) { "Regular" } else { "No mailbox found" }
Write-Host "Mailbox Type: $mailboxType"

# Getting delegated mailboxes
try {
    $delegatedPermissions = Get-MgUserMailFolderPermission -UserId $UserPrincipalName -MailFolderId 'inbox'
    if ($delegatedPermissions) {
        Write-Host "Delegated Mailboxes:"
        foreach ($perm in $delegatedPermissions) {
            Write-Host "- $($perm.User.DisplayName) with permission: $($perm.Role)"
        }
    } else {
        Write-Warning "No delegated mailboxes found for $UserPrincipalName."
    }
} catch {
    Write-Error "Failed to retrieve delegated mailboxes for $UserPrincipalName: $_"
}

# Getting security group memberships
try {
    $groups = Get-MgUserMemberOf -UserId $UserPrincipalName
    if ($groups) {
        Write-Host "Security Groups:"
        foreach ($group in $groups) {
            Write-Host "- $($group.DisplayName)"
        }
    } else {
        Write-Warning "No group memberships found for $UserPrincipalName."
    }
} catch {
    Write-Error "Failed to retrieve group memberships for $UserPrincipalName: $_"
}

# Getting SharePoint sites access
try {
    $sites = Get-MgSite -Filter "owners/any(o: o/email eq '$UserPrincipalName')"
    if ($sites) {
        Write-Host "SharePoint Sites Access:"
        foreach ($site in $sites) {
            Write-Host "- $($site.Name)"
        }
    } else {
        Write-Warning "No SharePoint sites found for $UserPrincipalName."
    }
} catch {
    Write-Error "Failed to retrieve SharePoint sites access for $UserPrincipalName: $_"
}

# Getting Teams site access
try {
    # Assuming you have a method to retrieve Teams site access here
    Write-Host "Getting Teams site access..."
    # Placeholder for Teams site access logic
} catch {
    Write-Error "Failed to retrieve Teams site access for $UserPrincipalName: $_"
}

# Getting users who have delegated access to the user's mailbox
try {
    $delegatedAccess = Get-MgUserMailFolderPermission -UserId $UserPrincipalName -MailFolderId 'inbox'
    if ($delegatedAccess) {
        Write-Host "Users with delegated access to mailbox:"
        foreach ($access in $delegatedAccess) {
            Write-Host "- $($access.User.DisplayName) with permission: $($access.Role)"
        }
    } else {
        Write-Warning "No mailbox delegates found or error occurred for $UserPrincipalName."
    }
} catch {
    Write-Error "Failed to retrieve mailbox delegates for $UserPrincipalName: $_"
}

# Prompt for file location to save results as an Excel file
$excelFilePath = [System.Windows.Forms.MessageBox]::Show('Enter the path to save the Excel file (or leave blank to skip):', 'Save Excel File', [System.Windows.Forms.MessageBoxButtons]::OKCancel)

if ($excelFilePath -ne [System.Windows.Forms.DialogResult]::Cancel) {
    # Export results to Excel (you will need to implement this part)
    Write-Host "Exporting results to Excel at: $excelFilePath"
    # Placeholder for Excel export logic
}

Write-Host "Script execution completed."
