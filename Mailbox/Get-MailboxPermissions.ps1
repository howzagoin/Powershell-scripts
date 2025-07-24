
# Metadata
<#
.SYNOPSIS
    A script to get mailbox permissions in Exchange Online.

.DESCRIPTION
    This script connects to Exchange Online and allows the user to get mailbox permissions.
    The user can choose to get a list of all mailboxes and their delegates, or to get a list of all mailboxes and distribution lists a specific user has access to.

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

# Function to get a list of all mailboxes and their delegates
function Get-AllMailboxesAndDelegates {
    $results = @()
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
    return $results
}

# Function to get a list of all mailboxes and distribution lists a specific user has access to
function Get-UserMailboxAndDistributionListAccess {
    param (
        [string]$UserEmail
    )
    $mailboxes = Get-Mailbox | Get-MailboxPermission -User $UserEmail
    $dls = Get-DistributionGroup | where { (Get-DistributionGroupMember $_.Name | foreach {$_.PrimarySmtpAddress}) -contains "$UserEmail"}
    return [PSCustomObject]@{
        Mailboxes = $mailboxes
        DistributionLists = $dls
    }
}

# Main script logic
Install-RequiredModules -ModuleNames @('ExchangeOnlineManagement', 'Microsoft.Graph', 'ImportExcel')
Connect-M365Services

do {
    Write-Host "Select an action:"
    Write-Host "1. Get all mailboxes and their delegates"
    Write-Host "2. Get all mailboxes and distribution lists a specific user has access to"
    Write-Host "3. Exit"
    $choice = Read-Host "Enter your choice"

    switch ($choice) {
        "1" {
            $results = Get-AllMailboxesAndDelegates
            $exportPath = Prompt-SaveFileDialog -DefaultFileName "MailboxPermissionsReport_$(Get-Date -Format 'yyyyMMdd').xlsx"
            if ($exportPath) {
                Export-ResultsToExcel -Data $results -FilePath $exportPath
            }
        }
        "2" {
            $userEmail = Read-Host "Enter the user's email address"
            $results = Get-UserMailboxAndDistributionListAccess -UserEmail $userEmail
            $exportPath = Prompt-SaveFileDialog -DefaultFileName "UserMailboxAndDistributionListAccess_$(Get-Date -Format 'yyyyMMdd').xlsx"
            if ($exportPath) {
                Export-ResultsToExcel -Data $results -FilePath $exportPath
            }
        }
        "3" {
            Write-Log "Exiting script." -Level Info
        }
        default {
            Write-Log "Invalid choice. Please select a valid option." -Level Warning
        }
    }
} while ($choice -ne "3")

Disconnect-ExchangeOnline -Confirm:$false
