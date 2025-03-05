# Import Microsoft Teams module
Import-Module MicrosoftTeams

# Connect to Microsoft Teams
Connect-MicrosoftTeams

# Prompt for user email
$userPrincipalName = Read-Host -Prompt "Enter the email address of the user (e.g., user@domain.com)"

# Retrieve user details
$user = Get-CsOnlineUser -Identity $userPrincipalName
Write-Output "User Details:"
$user | Format-List DisplayName, UserPrincipalName

# Retrieve Teams Meeting Policy details
Write-Output "`nTeams Meeting Policy Details:"
$meetingPolicies = Get-CsTeamsMeetingPolicy
$meetingPolicies | Format-Table -AutoSize

# Retrieve Teams App Permission Policy details
Write-Output "`nTeams App Permission Policy Details:"
$permissionPolicies = Get-CsTeamsAppPermissionPolicy
$permissionPolicies | Format-Table -AutoSize

# Retrieve Teams App Setup Policy details
Write-Output "`nTeams App Setup Policy Details:"
$setupPolicies = Get-CsTeamsAppSetupPolicy
$setupPolicies | Format-Table -AutoSize

# List available policies
Write-Output "`nAvailable Teams Meeting Policies:"
$meetingPolicies | Format-Table -AutoSize

Write-Output "`nAvailable Teams App Permission Policies:"
$permissionPolicies | Format-Table -AutoSize

Write-Output "`nAvailable Teams App Setup Policies:"
$setupPolicies | Format-Table -AutoSize
