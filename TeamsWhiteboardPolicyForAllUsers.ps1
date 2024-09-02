<#
.SYNOPSIS
    This script retrieves and displays policy details for a specified user in Microsoft Teams and compares Whiteboard access across all users.

.DESCRIPTION
    The script performs the following actions:
    1. Connects to Exchange Online using Multi-Factor Authentication (MFA).
    2. Prompts for the email address of a reference user.
    3. Retrieves and displays Teams policy details for the specified user, including Teams Meeting Policy, Teams App Permission Policy, and Teams App Setup Policy.
    4. Displays whether Whiteboard is allowed during Teams calls with external and anonymous users for the reference user.
    5. Compares Whiteboard access settings for all users against the reference user and lists users with different settings.
    6. Disconnects from Exchange Online after completing the operations.

.NOTES
    File Name      : TeamsWhiteboardPolicyForAllUsers.ps1
    Author          : Timothy MacLatchy
    Date            : 20/08/2024
    Requires        : ExchangeOnlineManagement module
#>

# Function to get policy details for a specific user
function Get-PolicyDetails {
    param (
        [string]$userPrincipalName
    )
    
    try {
        # Retrieve user details using the user principal name
        $user = Get-CsOnlineUser -Identity $userPrincipalName
        
        if ($user) {
            Write-Output "User found: $userPrincipalName"
            
            # Get available policies
            $availableMeetingPolicies = Get-CsTeamsMeetingPolicy
            $availableAppPermissionPolicies = Get-CsTeamsAppPermissionPolicy
            $availableAppSetupPolicies = Get-CsTeamsAppSetupPolicy
            
            # Retrieve and display Teams Meeting Policy
            $userPolicy = $user.TeamsMeetingPolicy
            if ($userPolicy) {
                Write-Output "Teams Meeting Policy assigned to ${userPrincipalName}: ${userPolicy}"
                $policyDetails = $availableMeetingPolicies | Where-Object { $_.Identity -eq $userPolicy }
                if ($policyDetails) {
                    Write-Output "Teams Meeting Policy Details for ${userPrincipalName}:"
                    Write-Output "Identity: $($policyDetails.Identity)"
                    Write-Output "AllowWhiteboard: $($policyDetails.AllowWhiteboard)"
                } else {
                    Write-Output "No matching policy details found for the assigned policy."
                }
            } else {
                Write-Output "No Teams Meeting Policy assigned to ${userPrincipalName}."
            }
            
            # Retrieve and display Teams App Permission Policy
            $userAppPermissionPolicy = $user.TeamsAppPermissionPolicy
            if ($userAppPermissionPolicy) {
                Write-Output "Teams App Permission Policy assigned to ${userPrincipalName}: ${userAppPermissionPolicy}"
                $appPermissionPolicyDetails = $availableAppPermissionPolicies | Where-Object { $_.Identity -eq $userAppPermissionPolicy }
                if ($appPermissionPolicyDetails) {
                    Write-Output "Teams App Permission Policy Details for ${userPrincipalName}:"
                    Write-Output "Identity: $($appPermissionPolicyDetails.Identity)"
                } else {
                    Write-Output "No matching Teams App Permission Policy details found for the assigned policy."
                }
            } else {
                Write-Output "No Teams App Permission Policy assigned to ${userPrincipalName}."
            }
            
            # Retrieve and display Teams App Setup Policy
            $userAppSetupPolicy = $user.TeamsAppSetupPolicy
            if ($userAppSetupPolicy) {
                Write-Output "Teams App Setup Policy assigned to ${userPrincipalName}: ${userAppSetupPolicy}"
                $appSetupPolicyDetails = $availableAppSetupPolicies | Where-Object { $_.Identity -eq $userAppSetupPolicy }
                if ($appSetupPolicyDetails) {
                    Write-Output "Teams App Setup Policy Details for ${userPrincipalName}:"
                    Write-Output "Identity: $($appSetupPolicyDetails.Identity)"
                } else {
                    Write-Output "No matching Teams App Setup Policy details found for the assigned policy."
                }
            } else {
                Write-Output "No Teams App Setup Policy assigned to ${userPrincipalName}."
            }
            
            # Retrieve and display Teams Meeting Policy Details for External Users
            Write-Output "Teams Meeting Policy Details for External Users:"
            $externalUsersPolicy = $availableMeetingPolicies | Where-Object { $_.Identity -eq "Tag:RestrictedAnonymousAccess" }
            if ($externalUsersPolicy.AllowWhiteboard) {
                Write-Output "Whiteboard is allowed during Teams calls with external users."
            } else {
                Write-Output "Whiteboard is not allowed during Teams calls with external users."
            }
            
            # Retrieve and display Teams Meeting Policy Details for Anonymous Users
            Write-Output "Teams Meeting Policy Details for Anonymous Users:"
            $anonymousUsersPolicy = $availableMeetingPolicies | Where-Object { $_.Identity -eq "Tag:RestrictedAnonymousNoRecording" }
            if ($anonymousUsersPolicy.AllowWhiteboard) {
                Write-Output "Whiteboard is allowed during Teams calls with anonymous users."
            } else {
                Write-Output "Whiteboard is not allowed during Teams calls with anonymous users."
            }
        } else {
            Write-Output "User not found or unable to retrieve user details."
        }
    } catch {
        Write-Error "An error occurred: $_"
    }
}

# Function to compare Whiteboard access for all users against a reference user
function Compare-WhiteboardAccess {
    param (
        [string]$referenceUser
    )
    
    try {
        # Retrieve the reference user policy
        $referenceUserPolicy = Get-CsOnlineUser -Identity $referenceUser
        $referenceUserWhiteboardAccess = $referenceUserPolicy.TeamsMeetingPolicy.AllowWhiteboard
        
        # Get all users with a practical limit for ResultSize
        $allUsers = Get-CsOnlineUser -ResultSize 5000
        $differentAccessUsers = @()
        
        foreach ($user in $allUsers) {
            if ($user.Identity -ne $referenceUser) {
                $userWhiteboardAccess = $user.TeamsMeetingPolicy.AllowWhiteboard
                if ($userWhiteboardAccess -ne $referenceUserWhiteboardAccess) {
                    $differentAccessUsers += $user.Identity
                }
            }
        }
        
        $userCount = $differentAccessUsers.Count
        
        Write-Output "Users with different Whiteboard access compared to ${referenceUser}:"
        $differentAccessUsers | ForEach-Object { Write-Output $_ }
        Write-Output "Total number of users with different Whiteboard access: $userCount"
    } catch {
        Write-Error "An error occurred while comparing Whiteboard access: $_"
    }
}

# Connect to Exchange Online with MFA
Connect-ExchangeOnline -UserPrincipalName $null -ShowProgress $true

# Prompt for the reference user email (e.g., David)
$referenceUser = Read-Host "Enter the email address of the reference user (e.g., david.jackson@firstfinancial.com.au)"

# Get and display policy details for the specified user
Get-PolicyDetails -userPrincipalName $referenceUser

# Compare Whiteboard access for all users against the reference user
Compare-WhiteboardAccess -referenceUser $referenceUser

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false