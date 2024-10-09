Connect-ExchangeOnline
$user = Read-Host -Prompt "Enter user's email address"

Write-Host "FINDING MAILBOXES..."
Write-Host "List of Mailboxes::"
Get-Mailbox | Get-MailboxPermission -User $user
Write-Host "FINDING DISTRIBUTION LISTS..."
Write-Host "List of Distribution Lists::"
Get-DistributionGroup | where { (Get-DistributionGroupMember $_.Name | foreach {$_.PrimarySmtpAddress}) -contains "$user"}
Disconnect-ExchangeOnline