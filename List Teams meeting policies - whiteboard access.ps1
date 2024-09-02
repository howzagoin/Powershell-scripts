# Connect to Teams PowerShell
Connect-MicrosoftTeams

# Get all users with the specific Teams meeting policy
$policy = "Tag:AllOff"
$users = Get-CsUser -PolicyAssignment @{ TeamsMeetingPolicy = $policy }

# Display the list of users
$users | Select-Object DisplayName, UserPrincipalName
