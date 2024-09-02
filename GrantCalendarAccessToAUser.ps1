# Define the list of possible permission levels
$permissionsList = @(
    "Owner",
    "PublishingEditor",
    "Editor",
    "PublishingAuthor",
    "Author",
    "NoneditingAuthor",
    "Reviewer",
    "Contributor",
    "AvailabilityOnly",
    "LimitedDetails"
)

# Function to prompt for user input
function Get-UserInput($prompt) {
    [Console]::Write("$prompt ")
    return [Console]::ReadLine()
}

# Prompt for the email addresses of the users
$user1Email = Read-Host -Prompt "Enter the email address of the user whose calendar is being shared:"
$user2Email = Read-Host -Prompt "Enter the email address of the user who will be given access:"

# Prompt for the permission level
Write-Host "Select the permission level to grant:"
for ($i = 0; $i -lt $permissionsList.Count; $i++) {
    Write-Host "$($i + 1). $($permissionsList[$i])"
}
$permissionIndex = [int](Read-Host -Prompt "Enter the number corresponding to the desired permission level:") - 1

if ($permissionIndex -lt 0 -or $permissionIndex -ge $permissionsList.Count) {
    Write-Host "Invalid selection. Exiting."
    exit
}

$selectedPermission = $permissionsList[$permissionIndex]

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..."
Connect-ExchangeOnline

# Grant the specified permissions
Write-Host "Granting $selectedPermission permission to $user2Email on $user1Email's calendar..."
Add-MailboxFolderPermission -Identity "${user1Email}:\Calendar" -User $user2Email -AccessRights $selectedPermission

# Clean up the session
Disconnect-ExchangeOnline

Write-Host "Permissions successfully granted."
