
# Metadata
<#
.SYNOPSIS
    A script to manage mailbox rules in Exchange Online.

.DESCRIPTION
    This script connects to Exchange Online and allows the user to inspect and manage mailbox rules.
    The user can choose to inspect all mailboxes, shared mailboxes only, or a specific mailbox.
    The user can then choose to delete all rules, a specific rule, or leave them unchanged.
    The script can also check for external redirect rules.

.AUTHOR
    Tim MacLatchy

.DATE
    02-07-2025

.LICENSE
    MIT License
#>

[CmdletBinding()]
param()

# Import the M365Utils module
Import-Module "C:\Users\Timothy.MacLatchy\OneDrive - Journe Brands\Documents\GitHubRepos\Powershell-scripts\Modules\M365Utils.psm1"

# Function to inspect and manage inbox rules
function Inspect-InboxRules {
    param (
        [string]$mailbox
    )

    Write-Log "Inspecting mailbox: $mailbox" -Level Info
    try {
        $rules = Get-InboxRule -Mailbox $mailbox
    } catch {
        Write-Log "Failed to retrieve inbox rules for mailbox: $mailbox" -Level Warning
        return
    }

    if ($rules.Count -eq 0) {
        Write-Log "No inbox rules found for mailbox: $mailbox" -Level Info
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
                    Write-Log "Deleted rule: $($rule.Name)" -Level Info
                } catch {
                    Write-Log "Failed to delete rule: $($rule.RuleIdentity)" -Level Error
                }
            }
            Write-Log "All inbox rules have been deleted." -Level Info
        }
        'S' {
            $ruleNumber = Read-Host "Enter the number of the rule you want to delete:"
            $ruleToDelete = $rulesList | Where-Object { $_.Index -eq [int]$ruleNumber }
            if ($ruleToDelete) {
                try {
                    Remove-InboxRule -Mailbox $mailbox -Identity $ruleToDelete.RuleIdentity
                    Write-Log "Deleted rule: $($ruleToDelete.Name)" -Level Info
                } catch {
                    Write-Log "Failed to delete rule: $($ruleToDelete.RuleIdentity)" -Level Error
                }
            } else {
                Write-Log "Invalid number. No rule deleted." -Level Warning
            }
        }
        'L' {
            Write-Log "No changes made to rules." -Level Info
        }
        default {
            Write-Log "Invalid option. Exiting." -Level Warning
            exit
        }
    }
}

# Function to check for external redirect rules
function Check-ExternalRedirectRules {
    param (
        [string]$SafeDomain
    )

    $mailboxesWithRedirectRules = Get-Mailbox -ResultSize Unlimited | Where-Object {
        $rules = Get-InboxRule -Mailbox $_.PrimarySmtpAddress
        $rules | Where-Object { $_.RedirectTo -and $_.RedirectTo -notlike "*@$SafeDomain" }
    }
    if ($mailboxesWithRedirectRules.Count -eq 0) {
        Write-Log "No mailboxes with redirect rules found." -Level Info
    } else {
        Write-Log "Listing mailboxes with redirect rules..." -Level Info
        foreach ($mailbox in $mailboxesWithRedirectRules) {
            Write-Host "--------------------------------------"
            Write-Host "Mailbox: $($mailbox.PrimarySmtpAddress)"
            Inspect-InboxRules -mailbox $mailbox.PrimarySmtpAddress
        }
    }
}

# Main script logic
Install-RequiredModules -ModuleNames @('ExchangeOnlineManagement', 'Microsoft.Graph', 'ImportExcel')
Connect-M365Services

do {
    Write-Host "Select an action:"
    Write-Host "1. Inspect all mailboxes"
    Write-Host "2. Inspect shared mailboxes only"
    Write-Host "3. Inspect a specific mailbox"
    Write-Host "4. Check for external redirect rules"
    Write-Host "5. Exit"
    $choice = Read-Host "Enter your choice"

    switch ($choice) {
        "1" {
            $mailboxes = Get-Mailbox -ResultSize Unlimited
            if ($mailboxes.Count -eq 0) {
                Write-Log "No mailboxes found." -Level Info
            } else {
                Write-Log "Listing all mailboxes..." -Level Info
                foreach ($mailbox in $mailboxes) {
                    Write-Host "--------------------------------------"
                    Write-Host "Mailbox: $($mailbox.PrimarySmtpAddress)"
                    Inspect-InboxRules -mailbox $mailbox.PrimarySmtpAddress
                }
            }
        }
        "2" {
            $sharedMailboxes = Get-Mailbox -ResultSize Unlimited -Filter { RecipientTypeDetails -eq "SharedMailbox" }
            if ($sharedMailboxes.Count -eq 0) {
                Write-Log "No shared mailboxes found." -Level Info
            } else {
                Write-Log "Listing shared mailboxes..." -Level Info
                foreach ($mailbox in $sharedMailboxes) {
                    Write-Host "--------------------------------------"
                    Write-Host "Mailbox: $($mailbox.PrimarySmtpAddress)"
                    Inspect-InboxRules -mailbox $mailbox.PrimarySmtpAddress
                }
            }
        }
        "3" {
            $specificMailbox = Read-Host "Enter the mailbox to inspect (e.g., sharedmailbox@example.com)"
            Inspect-InboxRules -mailbox $specificMailbox
        }
        "4" {
            $safeDomain = Read-Host "Enter your safe domain (e.g., contoso.com)"
            Check-ExternalRedirectRules -SafeDomain $safeDomain
        }
        "5" {
            Write-Log "Exiting script." -Level Info
        }
        default {
            Write-Log "Invalid choice. Please select a valid option." -Level Warning
        }
    }
} while ($choice -ne "5")

Disconnect-ExchangeOnline -Confirm:$false
