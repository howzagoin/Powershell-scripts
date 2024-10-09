# Ensure you have the Exchange Online module loaded
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
} catch {
    Write-Host "Error loading Exchange Online Management module: $_"
    exit
}

# Connect to Exchange Online as admin
$adminEmail = "welkin@firstfinancial.com.au"  # Update this to your admin email
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
    "katarina.coleman@crowehorwath.com",
    "kinsey.staver@crowehorwath.com",
    "laurie.marson@crowehorwath.com",
    "tina.tobin@crowehorwath.com",
    "bernita.pocock@anz.com",
    "brenda.wong@anz.com",
    "advisers@genlife.com.au",
    "cdong@genlife.com.au",
    "enquiry@genlife.com.au",
    "faraujo@genlife.com.au",
    "ghackett@genlife.com.au",
    "mcridland@genlife.com.au",
    "mfairbairn@genlife.com.au",
    "services@genlife.com.au",
    "servicesrequests@genlife.com.au",
    "franco.pistritto@anz.com",
    "jennifer.mcculloch@anz.com",
    "mark.curran@anz.com",
    "matthew.scarmozzino@anz.com",
    "peter.james@anz.com",
    "scott.brading@anz.com",
    "vicuw@anz.com",
    "accelerateservice@tal.com.au",
    "communications@tal.com.au",
    "customerservice@tal.com.au",
    "jessica.magarry@tal.com.au",
    "laura.sicari@tal.com.au",
    "max.warton@tal.com.au",
    "paul.bird@tal.com.au",
    "preassessvic@tal.com.au",
    "andrew.dobson@findex.com.au",
    "chris.hall@findex.com.au",
    "cloud.melbourne@findex.com.au",
    "james.potiphar@findex.com.au",
    "jenny.zanon@findex.com.au",
    "mieka.decker@findex.com.au",
    "naween.fernando@findex.com.au",
    "visal.kim@findex.com.au",
    "apeters@hallmarc.com.au",
    "icheah@hallmarc.com.au",
    "michael@hallmarc.com.au",
    "mkoraus@hallmarc.com.au",
    "mloccisano@hallmarc.com.au",
    "pdigiorgio@hallmarc.com.au",
    "srava@hallmarc.com.au",
    "francisco@performanceproperty.com.au",
    "jericha@performanceproperty.com.au",
    "justin@performanceproperty.com.au",
    "melinda@performanceproperty.com.au",
    "paul@performanceproperty.com.au",
    "phillip@performanceproperty.com.au",
    "william@performanceproperty.com.au",
    "anne.noakes@agedcaresteps.com.au",
    "info@agedcaresteps.com.au",
    "lara.hansen@agedcaresteps.com.au",
    "louise.biti@agedcaresteps.com.au",
    "natasha.panagis@agedcaresteps.com.au",
    "paraplanning@agedcaresteps.com.au",
    "training@agedcaresteps.com.au",
    "cassandra.mckay@slatergordon.com.au",
    "laura.harding@slatergordon.com.au",
    "rabia.javed-may@slatergordon.com.au",
    "rabia.javed@slatergordon.com.au",
    "rabia.javedmay@slatergordon.com.au",
    "rachael.maharaj@slatergordon.com.au",
    "sarah.murphy@agedcaresteps.com.au"
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
