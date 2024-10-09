Connect-ExchangeOnline
$user = Read-Host -Prompt "enter user's email address"
Enable-Mailbox $user -AutoExpandingArchive
Disconnect-ExchangeOnline