# Ensure you have the Exchange Online module loaded
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
} catch {
    Write-Host "Error loading Exchange Online Management module: $_"
    exit
}

# Connect to Exchange Online as admin
$adminEmail = "admin@domain.com"  # Update this to your admin email
try {
    Connect-ExchangeOnline -UserPrincipalName $adminEmail -ShowProgress $true -ErrorAction Stop
} catch {
    Write-Host "Error connecting to Exchange Online: $_"
    exit
}

# Prompt for user email to check the safe sender list
$userEmail = Read-Host "Enter the user email to check the safe sender list"

# Define addresses to remove
$addressesToRemove = @(
    "user@domain.com"
)

# Get current trusted senders and domains for the specified user
try {
    $currentTrustedSenders = Get-MailboxJunkEmailConfiguration -Identity $userEmail -ErrorAction Stop
    $trustedSendersList = $currentTrustedSenders.TrustedSendersAndDomains

    # Ensure it's an array
    if (-not $trustedSendersList) {
        $trustedSendersList = @()
    } elseif (-not ($trustedSendersList -is [array])) {
        $trustedSendersList = @($trustedSendersList)
    }
} catch {
    Write-Host "Error getting mailbox junk email configuration for ${userEmail}: $_"
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# Initial count
$initialCount = $trustedSendersList.Count
Write-Host "Initial number of trusted senders and domains: $initialCount"

# Remove specified addresses from the list
$filteredSenders = $trustedSendersList | Where-Object { $addressesToRemove -notcontains $_ }

# Update the mailbox junk email configuration if there are changes
if ($filteredSenders.Count -ne $initialCount) {
    try {
        Set-MailboxJunkEmailConfiguration -Identity $userEmail -TrustedSendersAndDomains $filteredSenders -ErrorAction Stop
        Write-Host "Updated trusted senders and domains list successfully."
    } catch {
        Write-Host "Error updating mailbox junk email configuration: $($_.Exception.Message)"
    }
} else {
    Write-Host "No addresses were removed; the list is unchanged."
}

# Final output
$finalCount = $filteredSenders.Count
Write-Host "Final number of trusted senders and domains: $finalCount"
Write-Host "Number of entries removed: $($initialCount - $finalCount)"

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
