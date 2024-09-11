# Prompt for admin username
$adminUser = Read-Host "Enter admin username"

# Ask if the user wants to output each section
$outputResources = Read-Host "Do you want to list resource mailboxes and their time zones? (y/n)"
$outputUsers = Read-Host "Do you want to list user mailboxes and their time zones? (y/n)"
$outputGuestUsers = Read-Host "Do you want to list guest users and their time zones? (y/n)"

# Connect to Exchange Online with MFA
Write-Host "Connecting to Exchange Online with MFA..."
Connect-ExchangeOnline -UserPrincipalName $adminUser

# Function to get time zone or display "Not Set"
function Get-TimeZone {
    param ($Identity)
    $timeZoneConfig = Get-MailboxRegionalConfiguration -Identity $Identity
    if ($timeZoneConfig.TimeZone -ne $null) {
        return $timeZoneConfig.TimeZone
    } else {
        return "Not Set"
    }
}

# List resource mailboxes if the user selected 'y'
if ($outputResources -eq 'y') {
    Write-Host "`nListing all resource mailboxes and their time zones..."
    $resources = Get-Mailbox -RecipientTypeDetails RoomMailbox, EquipmentMailbox | Sort-Object DisplayName
    foreach ($resource in $resources) {
        $timeZone = Get-TimeZone -Identity $resource.Identity
        Write-Host "$($resource.DisplayName) ($($resource.PrimarySmtpAddress)) - Time Zone: $timeZone"
    }
}

# List user mailboxes if the user selected 'y'
if ($outputUsers -eq 'y') {
    Write-Host "`nListing all users and their time zones..."
    $users = Get-Mailbox -RecipientTypeDetails UserMailbox | Sort-Object DisplayName
    foreach ($user in $users) {
        $timeZone = Get-TimeZone -Identity $user.Identity
        Write-Host "$($user.DisplayName) ($($user.PrimarySmtpAddress)) - Time Zone: $timeZone"
    }
}

# List guest users if the user selected 'y'
if ($outputGuestUsers -eq 'y') {
    Write-Host "`nListing all guest users and their time zones..."
    $guestUsers = Get-MailUser | Sort-Object DisplayName
    foreach ($guestUser in $guestUsers) {
        $timeZone = Get-TimeZone -Identity $guestUser.Identity
        Write-Host "$($guestUser.DisplayName) ($($guestUser.PrimarySmtpAddress)) - Time Zone: $timeZone"
    }
}

# Disconnect from Exchange Online
Write-Host "Disconnecting from Exchange Online..."
Disconnect-ExchangeOnline -Confirm:$false

Write-Host "Script complete."
