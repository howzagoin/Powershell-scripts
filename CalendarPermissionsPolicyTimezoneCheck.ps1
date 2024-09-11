# Author: Tim MacLatchy
# Date: 2024-09-11
# License: MIT License
# Description: This script logs into a tenant with interactive MFA, checks for the required modules, 
# and outputs the time zone, booking policy, and calendar permissions for a specified calendar.

# Check for required modules and install if not present
$modules = @("ExchangeOnlineManagement")
foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Module $module is not installed. Installing..."
        Install-Module $module -Force
    }
}

# Sign into Exchange Online with MFA
Write-Host "Connecting to Exchange Online with MFA..."
$adminUser = Read-Host "Enter your admin username"
Connect-ExchangeOnline -UserPrincipalName $adminUser -ShowProgress $true

# Prompt for resource email
$resourceEmail = Read-Host "Enter the calendar email you want to check"

# Fetch and display the resource mailbox's timezone
Write-Host "Fetching calendar mailbox details..."
$timeZone = Get-MailboxRegionalConfiguration -Identity $resourceEmail | Select-Object -ExpandProperty TimeZone
if (-not $timeZone) {
    $timeZone = "Not set"
}
Write-Host "Time Zone for ${resourceEmail}: $timeZone"

# Fetch and display the resource mailbox's booking policy
$bookingPolicy = Get-CalendarProcessing -Identity $resourceEmail
Write-Host "Booking Policy for ${resourceEmail}:"
Write-Host "  Booking Window (in days): $($bookingPolicy.MaximumBookingDuration)"
Write-Host "  Enforce Scheduling Horizon: $($bookingPolicy.EnforceSchedulingHorizon)"
Write-Host "  Allow Conflicts: $($bookingPolicy.AllowConflicts)"

# Fetch and display the resource mailbox's calendar permissions
Write-Host "Fetching calendar permissions for $resourceEmail..."
$calendarPermissions = Get-MailboxFolderPermission -Identity "${resourceEmail}:\Calendar" | Sort-Object -Property User

Write-Host "Calendar Permissions for ${resourceEmail}:"
foreach ($permission in $calendarPermissions) {
    Write-Host "  User: $($permission.User) - Access Rights: $($permission.AccessRights)"
}

# Disconnect from Exchange Online
Write-Host "Disconnecting from Exchange Online..."
Disconnect-ExchangeOnline -Confirm:$false

Write-Host "Script completed."
