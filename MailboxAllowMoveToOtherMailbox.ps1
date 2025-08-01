# 1. Connect to Exchange Online
Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser
Connect-ExchangeOnline -UserPrincipalName timothy.maclatchy@journebrands.com

# 2. Create a new custom OWA mailbox policy
New-OwaMailboxPolicy -Name "AllowMoveBetweenAccounts" -Confirm:$false

# 3. Enable cross-account message move/copy
Set-OwaMailboxPolicy -Identity "AllowMoveBetweenAccounts" -ItemsToOtherAccountsEnabled $true

# 4. Assign the new policy to specific mailboxes
Set-CASMailbox -Identity "timothy.maclatchy@journebrands.com" -OwaMailboxPolicy "AllowMoveBetweenAccounts"
Set-CASMailbox -Identity "itsupport@journebrands.com"       -OwaMailboxPolicy "AllowMoveBetweenAccounts"

# 5. Verify assignments
Get-CASMailbox -Identity "timothy.maclatchy@journebrands.com" | Format-List Name,OwaMailboxPolicy
Get-CASMailbox -Identity "itsupport@journebrands.com"       | Format-List Name,OwaMailboxPolicy
