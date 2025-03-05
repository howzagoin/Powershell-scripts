# Ensure the ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

# Define credentials and connection details
$AdminUsername = 'admin@domain.com'
$AdminPassword = 'Password123'
$UserToSearch = 'user@domain.com'

# Connect to Exchange Online
$SecurePassword = ConvertTo-SecureString $AdminPassword -AsPlainText -Force
$Credentials = New-Object System.Management.Automation.PSCredential ($AdminUsername, $SecurePassword)
Connect-ExchangeOnline -Credential $Credentials

Write-Host "Successfully connected to Exchange Online."

# Define the list of email addresses to add to the safe sender list
$SafeSenders = @(
    "example@domain.com"
)

# Get the current safe sender list and remove duplicates
$CurrentSafeSenders = (Get-MailboxJunkEmailConfiguration -Identity $UserToSearch).TrustedSendersAndDomains
$SafeSenders = $SafeSenders | Sort-Object -Unique

# Ensure that we only add new safe senders
$NewSafeSenders = ($CurrentSafeSenders + $SafeSenders) | Sort-Object -Unique

# Update the safe sender list
Set-MailboxJunkEmailConfiguration -Identity $UserToSearch -TrustedSendersAndDomains $NewSafeSenders

Write-Host "Safe sender list updated successfully for $UserToSearch."
