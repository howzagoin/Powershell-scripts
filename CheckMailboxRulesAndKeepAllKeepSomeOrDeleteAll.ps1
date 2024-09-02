# Import required modules
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
}
Import-Module ExchangeOnlineManagement

# Prompt for admin login
$adminAccount = Read-Host -Prompt "Enter the admin account (e.g., admin@domain.com)"
Connect-ExchangeOnline -UserPrincipalName $adminAccount -ShowProgress $true

# Function to inspect and manage inbox rules
function Inspect-InboxRules {
    param (
        [string]$mailbox
    )

    Write-Host "Inspecting mailbox: $mailbox"
    try {
        $rules = Get-InboxRule -Mailbox $mailbox
    } catch {
        Write-Host "Failed to retrieve inbox rules for mailbox: $mailbox"
        return
    }

    if ($rules.Count -eq 0) {
        Write-Host "No inbox rules found for mailbox: $mailbox"
        return
    }

    # Display and handle rules
    $rulesList = @()
    $index = 1
    foreach ($rule in $rules) {
        Write-Host "--------------------------------------"
        Write-Host "Rule ${index}:"
        Write-Host "Name       : $($rule.Name)"
        Write-Host "Description: $($rule.Description)"
        Write-Host "Enabled    : $($rule.Enabled)"
        Write-Host "Priority   : $($rule.Priority)"
        Write-Host "Conditions : $($rule.Conditions)"
        Write-Host "Actions    : $($rule.Actions)"
        Write-Host "RuleIdentity : $($rule.RuleIdentity)"
        Write-Host "--------------------------------------"
        $rulesList += [pscustomobject]@{
            Index        = $index
            Name         = $rule.Name
            Description  = $rule.Description
            Enabled      = $rule.Enabled
            Priority     = $rule.Priority
            Conditions   = $rule.Conditions
            Actions      = $rule.Actions
            RuleIdentity = $rule.RuleIdentity
        }
        $index++
    }

    # Ask user for action
    $action = Read-Host "Do you want to (D)elete all rules, (S)elect a specific rule to delete, or (L)eave them unchanged? (Enter D, S, or L)"
    switch ($action.ToUpper()) {
        'D' {
            foreach ($rule in $rulesList) {
                try {
                    Remove-InboxRule -Mailbox $mailbox -Identity $rule.RuleIdentity
                    Write-Host "Deleted rule: $($rule.Name)"
                } catch {
                    Write-Host "Failed to delete rule: $($rule.RuleIdentity)"
                }
            }
            Write-Host "All inbox rules have been deleted."
        }
        'S' {
            $ruleNumber = Read-Host "Enter the number of the rule you want to delete:"
            $ruleToDelete = $rulesList | Where-Object { $_.Index -eq [int]$ruleNumber }
            if ($ruleToDelete) {
                try {
                    Remove-InboxRule -Mailbox $mailbox -Identity $ruleToDelete.RuleIdentity
                    Write-Host "Deleted rule: $($ruleToDelete.Name)"
                } catch {
                    Write-Host "Failed to delete rule: $($ruleToDelete.RuleIdentity)"
                }
            } else {
                Write-Host "Invalid number. No rule deleted."
            }
        }
        'L' {
            Write-Host "No changes made to rules."
        }
        default {
            Write-Host "Invalid option. Exiting."
            exit
        }
    }
}

# Main script logic
do {
    $choice = Read-Host "Do you want to inspect (A)ll mailboxes, (S)hared mailboxes only, (O)ne specific mailbox, or (E)xit? (Enter A, S, O, or E)"
    switch ($choice.ToUpper()) {
        'A' {
            $mailboxes = Get-Mailbox -ResultSize Unlimited
            if ($mailboxes.Count -eq 0) {
                Write-Host "No mailboxes found."
            } else {
                Write-Host "Listing all mailboxes..."
                foreach ($mailbox in $mailboxes) {
                    Write-Host "--------------------------------------"
                    Write-Host "Mailbox: $($mailbox.PrimarySmtpAddress)"
                    Inspect-InboxRules -mailbox $mailbox.PrimarySmtpAddress
                }

                # Ask user if they want to save results to Excel
                $saveExcel = Read-Host "Do you want to save the mailbox data to an Excel file? (Y/N)"
                if ($saveExcel.ToUpper() -eq 'Y') {
                    # Export to Excel
                    $allMailboxesInfo = @()
                    foreach ($mailbox in $mailboxes) {
                        $mailboxInfo = [pscustomobject]@{
                            Mailbox     = $mailbox.PrimarySmtpAddress
                            Type         = if ($mailbox.RecipientTypeDetails -eq "SharedMailbox") { "Shared" } else { "Regular" }
                            RulesStatus  = "No Rules"
                        }
                        $rules = Get-InboxRule -Mailbox $mailbox.PrimarySmtpAddress
                        if ($rules.Count -gt 0) {
                            $mailboxInfo.RulesStatus = "Has Rules"
                        }
                        $allMailboxesInfo += $mailboxInfo
                    }

                    if ($allMailboxesInfo.Count -gt 0) {
                        Add-Type -AssemblyName System.Windows.Forms
                        $fileDialog = New-Object System.Windows.Forms.SaveFileDialog
                        $fileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
                        $fileDialog.Title = "Save Excel File"
                        if ($fileDialog.ShowDialog() -eq "OK") {
                            $filePath = $fileDialog.FileName
                            $allMailboxesInfo | Export-Excel -Path $filePath -AutoSize -WorksheetName "Mailboxes"
                            Write-Host "Data exported to $filePath"
                        }
                    }
                }
            }
        }
        'S' {
            $sharedMailboxes = Get-Mailbox -ResultSize Unlimited -Filter { RecipientTypeDetails -eq "SharedMailbox" }
            if ($sharedMailboxes.Count -eq 0) {
                Write-Host "No shared mailboxes found."
            } else {
                Write-Host "Listing shared mailboxes..."
                foreach ($mailbox in $sharedMailboxes) {
                    Write-Host "--------------------------------------"
                    Write-Host "Mailbox: $($mailbox.PrimarySmtpAddress)"
                    Inspect-InboxRules -mailbox $mailbox.PrimarySmtpAddress
                }

                # Ask user if they want to save results to Excel
                $saveExcel = Read-Host "Do you want to save the mailbox data to an Excel file? (Y/N)"
                if ($saveExcel.ToUpper() -eq 'Y') {
                    # Export to Excel
                    $sharedMailboxesInfo = @()
                    foreach ($mailbox in $sharedMailboxes) {
                        $mailboxInfo = [pscustomobject]@{
                            Mailbox     = $mailbox.PrimarySmtpAddress
                            RulesStatus  = "No Rules"
                        }
                        $rules = Get-InboxRule -Mailbox $mailbox.PrimarySmtpAddress
                        if ($rules.Count -gt 0) {
                            $mailboxInfo.RulesStatus = "Has Rules"
                        }
                        $sharedMailboxesInfo += $mailboxInfo
                    }

                    if ($sharedMailboxesInfo.Count -gt 0) {
                        Add-Type -AssemblyName System.Windows.Forms
                        $fileDialog = New-Object System.Windows.Forms.SaveFileDialog
                        $fileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
                        $fileDialog.Title = "Save Excel File"
                        if ($fileDialog.ShowDialog() -eq "OK") {
                            $filePath = $fileDialog.FileName
                            $sharedMailboxesInfo | Export-Excel -Path $filePath -AutoSize -WorksheetName "SharedMailboxes"
                            Write-Host "Data exported to $filePath"
                        }
                    }
                }
            }
        }
        'O' {
            $specificMailbox = Read-Host "Enter the mailbox to inspect (e.g., sharedmailbox@example.com)"
            Inspect-InboxRules -mailbox $specificMailbox
        }
        'E' {
            Write-Host "Exiting script."
        }
        default {
            Write-Host "Invalid option. Please select again."
        }
    }
} while ($choice.ToUpper() -ne 'E')

# Disconnect from the Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false
