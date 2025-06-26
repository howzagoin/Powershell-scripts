# Author: Tim MacLatchy
# Date: 16/10/2024
# License: MIT License
# Description: This script connects to Exchange Online to inspect mailboxes with redirect rules. 

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
$mailboxesWithRedirectRules = Get-Mailbox -ResultSize Unlimited | Where-Object {
    $rules = Get-InboxRule -Mailbox $_.PrimarySmtpAddress
    $rules | Where-Object { $_.RedirectTo -and $_.RedirectTo -notlike "*@safecompanydomain.com" }
    #SET THE safecompanydomain.com TO CHECK FOR REDIRECT RULES to external domains
}
if ($mailboxesWithRedirectRules.Count -eq 0) {
    Write-Host "No mailboxes with redirect rules found."
} else {
    Write-Host "Listing mailboxes with redirect rules..."
    foreach ($mailbox in $mailboxesWithRedirectRules) {
        Write-Host "--------------------------------------"
        Write-Host "Mailbox: $($mailbox.PrimarySmtpAddress)"
        Inspect-InboxRules -mailbox $mailbox.PrimarySmtpAddress
    }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
