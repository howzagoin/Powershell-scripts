# Script to check calendar permissions in Exchange Online
# This script ensures required modules are installed, connects to Exchange Online with MFA,
# retrieves calendar permissions for a specified email address, and then disconnects from Exchange Online.
# Date: 2024-08-29
# Author: Tim MacLatchy

# Ensure required modules are installed
$modules = @("ExchangeOnlineManagement")
foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module -Name $module -Force -Scope CurrentUser
    }
}

# Request tenant admin username and calendar email
$tenantAdmin = Read-Host "Enter tenant admin username"
$calendarEmail = Read-Host "Enter the email address of the calendar to check"

# Connect to Exchange Online with MFA
Connect-ExchangeOnline -UserPrincipalName $tenantAdmin -ShowProgress $true

# Define the mailbox to check
$calendarPath = "${calendarEmail}:\Calendar"

# Get calendar permissions
Get-MailboxFolderPermission -Identity $calendarPath | Format-Table

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Script completed. Disconnected from Exchange Online."
