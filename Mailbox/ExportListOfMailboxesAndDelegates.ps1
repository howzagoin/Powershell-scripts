# Connect to Exchange Online
Connect-ExchangeOnline

# Set output file path
$outputFile = "C:\Temp\mailboxPermissionsReport.csv"

# Ensure the output directory exists
if (-not (Test-Path -Path (Split-Path $outputFile))) {
    New-Item -Path (Split-Path $outputFile) -ItemType Directory -Force
}

# Initialize an array to store the results
$results = @()

# Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

foreach ($mailbox in $mailboxes) {
    try {
        # Retrieve Full Access delegates
        $delegates = Get-MailboxPermission -Identity $mailbox.UserPrincipalName | Where-Object { 
            $_.AccessRights -contains "FullAccess" -and $_.IsInherited -eq $false 
        }
        
        # Retrieve Send-As permissions
        $sendAs = Get-RecipientPermission -Identity $mailbox.UserPrincipalName | Where-Object { 
            $_.AccessRights -contains "SendAs" 
        }

        # Process Full Access delegates
        foreach ($delegate in $delegates) {
            $results += [PSCustomObject]@{
                Mailbox         = $mailbox.UserPrincipalName
                DelegateUser    = $delegate.User
                AccessType      = "Full Access"
            }
        }

        # Process Send-As permissions
        foreach ($sendAsUser in $sendAs) {
            $results += [PSCustomObject]@{
                Mailbox         = $mailbox.UserPrincipalName
                DelegateUser    = $sendAsUser.Trustee
                AccessType      = "Send-As"
            }
        }
    }
    catch {
        Write-Warning "Failed to retrieve permissions for mailbox: $($mailbox.UserPrincipalName). Error: $_"
    }
}

# Export results to CSV
if ($results.Count -gt 0) {
    $results | Sort-Object Mailbox, AccessType | Export-Csv -Path $outputFile -NoTypeInformation
    Write-Output "Export completed. Results saved to $outputFile"
} else {
    Write-Output "No permissions found to export."
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
