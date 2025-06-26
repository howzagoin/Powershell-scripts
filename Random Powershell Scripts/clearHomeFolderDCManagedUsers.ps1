Import-Module ActiveDirectory

# Define the target OU
$ou = "OU=Managed Users,DC=fba,DC=com,DC=au"

# Get all users in the specified OU
$users = Get-ADUser -Filter * -SearchBase $ou -Properties HomeDirectory, HomeDrive

foreach ($user in $users) {
    Set-ADUser -Identity $user.DistinguishedName `
               -HomeDirectory $null `
               -HomeDrive $null
    Write-Host "✅ Cleared home folder for: $($user.SamAccountName)"
}
