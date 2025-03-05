# Connect to Exchange Online
Connect-ExchangeOnline

$outputFile = "C:\Temp\mailboxPermissionsReport.csv"
# Initialize an array to store the results
$results = @()

# Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

foreach ($mailbox in $mailboxes) {
    $delegates = Get-MailboxPermission -Identity $mailbox.UserPrincipalName | Where-Object { $_.AccessRights -eq "FullAccess" -and $_.IsInherited -eq $false }
    $sendAs = Get-RecipientPermission -Identity $mailbox.UserPrincipalName | Where-Object { $_.AccessRights -contains "SendAs" }

    foreach ($delegate in $delegates) {
        $results += [PSCustomObject]@{
            Mailbox         = $mailbox.UserPrincipalName
            DelegateUser    = $delegate.User
            AccessType      = "Delegate"
        }
    }

    foreach ($sendAsUser in $sendAs) {
        $results += [PSCustomObject]@{
            Mailbox         = $mailbox.UserPrincipalName
            DelegateUser    = $sendAsUser.Trustee
            AccessType      = "SendAs"
        }
    }
}

# Export results to CSV
$results | Export-Csv -Path $outputFile -NoTypeInformation

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

Write-Output "Export completed. Results saved to $outputFile"
