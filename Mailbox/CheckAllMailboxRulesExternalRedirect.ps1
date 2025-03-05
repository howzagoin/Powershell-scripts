# Author: Tim MacLatchy
# Date: 16/10/2024
# License: MIT License
# Description: This script connects to Exchange Online to inspect mailboxes with redirect rules to external addresses.

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
        [string]$mailbox,
        [array]$rules
    )

    Write-Host "Inspecting mailbox: $mailbox"
    
    if ($rules.Count -eq 0) {
        Write-Host "No inbox rules found for mailbox: $mailbox"
        return
    }

    # Display and handle rules
    $index = 1
    foreach ($rule in $rules) {
        Write-Host "--------------------------------------"
        Write-Host "Rule ${index}:"
        Write-Host "Name         : $($rule.Name)"
        Write-Host "Description  : $($rule.Description)"
        Write-Host "Enabled      : $($rule.Enabled)"
        Write-Host "Priority     : $($rule.Priority)"
        Write-Host "Conditions   : $($rule.Conditions)"
        Write-Host "Actions      : $($rule.Actions)"
        Write-Host "Created On   : $($rule.WhenCreated)"  # Include creation date
        Write-Host "RuleIdentity : $($rule.RuleIdentity)"
        Write-Host "--------------------------------------"

        $index++
    }

    # Ask user for action
    $action = Read-Host "Do you want to (D)elete all rules, (S)elect a specific rule to delete, or (L)eave them unchanged? (Enter D, S, or L)"
    switch ($action.ToUpper()) {
        'D' {
            foreach ($rule in $rules) {
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
            $ruleToDelete = $rules | Where-Object { $_.Index -eq [int]$ruleNumber }
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
            Write-Host "No valid option selected. Defaulting to leave unchanged."
            Write-Host "No changes made to rules."
        }
    }
}

# Main script logic
$totalMailboxesScanned = 0
$totalMailboxesWithRedirectRules = 0

$mailboxes = Get-Mailbox -ResultSize Unlimited

foreach ($mailbox in $mailboxes) {
    $totalMailboxesScanned++

    try {
        # Attempt to get inbox rules for the mailbox
        $rules = Get-InboxRule -Mailbox $mailbox.PrimarySmtpAddress
        
        # Check for redirect rules to external domains
        $redirectRules = $rules | Where-Object { $_.RedirectTo -and $_.RedirectTo -notlike "*@$(($mailbox.PrimarySmtpAddress -split '@')[1])" }

        # Display if any valid rules are found
        if ($redirectRules.Count -gt 0) {
            Write-Host "--------------------------------------"
            Write-Host "Mailbox: $($mailbox.PrimarySmtpAddress)"
            Inspect-InboxRules -mailbox $mailbox.PrimarySmtpAddress $redirectRules
            $totalMailboxesWithRedirectRules++
        }

    } catch {
        # Handle any errors during rule retrieval or processing
        Write-Warning "The Inbox rule for mailbox $($mailbox.PrimarySmtpAddress) contains errors. Please review the rule manually."
    }
}

# Summary output
Write-Host "--------------------------------------"
Write-Host "Summary:"
Write-Host "Total mailboxes scanned: $totalMailboxesScanned"
Write-Host "Total mailboxes with redirect rules to external addresses: $totalMailboxesWithRedirectRules"

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
