
# Metadata
<#
.SYNOPSIS
    A script to manage Teams policies in Microsoft Teams.

.DESCRIPTION
    This script connects to Microsoft Teams and allows the user to get and set Teams policies.
    The user can choose to get policy details for a specific user, list users with a specific policy, or manage whiteboard and recording permissions.

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

# Function to get policy details for a user
function Get-PolicyDetails {
    param (
        [string]$userPrincipalName
    )

    # Retrieve user details
    $user = Get-MgUser -UserId $userPrincipalName
    Write-Log "User Details:" -Level Info
    Write-Host ""
    Write-Host "DisplayName       : $($user.DisplayName)"
    Write-Host "UserPrincipalName : $($user.UserPrincipalName)"
    Write-Host ""

    # Retrieve and display Teams meeting policy details
    $meetingPolicies = Get-MgTeamsMeetingPolicy
    Write-Log "Teams Meeting Policy Details:" -Level Info
    $userMeetingPolicy = Get-CsTeamsMeetingPolicy -Identity $userPrincipalName
    Write-Log "User Teams Meeting Policy:" -Level Info
    $userMeetingPolicy | Format-List
    Write-Host ""
    Write-Log "Available Teams Meeting Policies:" -Level Info
    $meetingPolicies | Format-Table -Property Identity, Description, AllowChannelMeetingScheduling, AllowMeetNow, AllowPrivateMeetNow, MeetingChatEnabled, Type

    # Retrieve and display Teams app permission policy details
    $permissionPolicies = Get-MgTeamsAppPermissionPolicy
    Write-Log "Teams App Permission Policy Details:" -Level Info
    $userPermissionPolicy = Get-CsTeamsAppPermissionPolicy -Identity $userPrincipalName
    Write-Log "User Teams App Permission Policy:" -Level Info
    $userPermissionPolicy | Format-List
    Write-Host ""
    Write-Log "Available Teams App Permission Policies:" -Level Info
    $permissionPolicies | Format-Table -Property Identity, DefaultCatalogApps, GlobalCatalogApps, PrivateCatalogApps, Description, DefaultCatalogAppsType, GlobalCatalogAppsType, PrivateCatalogAppsType

    # Retrieve and display Teams app setup policy details
    $setupPolicies = Get-MgTeamsAppSetupPolicy
    Write-Log "Teams App Setup Policy Details:" -Level Info
    $userSetupPolicy = Get-CsTeamsAppSetupPolicy -Identity $userPrincipalName
    Write-Log "User Teams App Setup Policy:" -Level Info
    $userSetupPolicy | Format-List
    Write-Host ""
    Write-Log "Available Teams App Setup Policies:" -Level Info
    $setupPolicies | Format-Table -Property Identity, AppPresetList, PinnedAppBarApps

    # List allowed and banned Teams apps for the user
    $userAppPermissionPolicy = Get-MgTeamsAppPermissionPolicy -UserId $userPrincipalName
    Write-Log "Teams Apps Allowed for User:" -Level Info
    $userAppPermissionPolicy.DefaultCatalogApps | Format-List
    Write-Host ""
    Write-Log "Teams Apps Banned for User:" -Level Info
    $userAppPermissionPolicy.BlockedAppList | Format-List
    Write-Host ""

    # List apps banned and allowed during Teams video calls
    $meetingPolicy = $userMeetingPolicy | Select-Object -ExpandProperty MeetingPolicies
    Write-Log "Teams Apps Allowed During Video Calls:" -Level Info
    $meetingPolicy.AllowPrivateMeetNow | Format-List
    Write-Host ""
    Write-Log "Teams Apps Banned During Video Calls:" -Level Info
    $meetingPolicy.BlockedAppList | Format-List
}

# Function to list users with a specific Teams meeting policy
function Get-UsersWithPolicy {
    param (
        [string]$PolicyName
    )
    $users = Get-CsUser -PolicyAssignment @{ TeamsMeetingPolicy = $PolicyName }
    $users | Select-Object DisplayName, UserPrincipalName
}

# Function to manage whiteboard access
function Manage-WhiteboardAccess {
    param (
        [string]$PolicyName,
        [boolean]$AllowWhiteboard
    )
    Set-CsTeamsMeetingPolicy -Identity $PolicyName -AllowWhiteboard $AllowWhiteboard
    Write-Log "Whiteboard access for policy '$PolicyName' set to '$AllowWhiteboard'." -Level Info
}

# Function to check user recording permissions
function Get-UserRecordingPermissions {
    param (
        [string]$UserPrincipalName
    )
    $user = Get-CsOnlineUser -Identity $UserPrincipalName
    $policy = Get-CsTeamsMeetingPolicy -Identity $user.TeamsMeetingPolicy
    return $policy.AllowCloudRecording
}

# Main script logic
Install-RequiredModules -ModuleNames @('MicrosoftTeams', 'Microsoft.Graph', 'ImportExcel')
Connect-M365Services

do {
    Write-Host "Select an action:"
    Write-Host "1. Get policy details for a specific user"
    Write-Host "2. List users with a specific Teams meeting policy"
    Write-Host "3. Manage whiteboard access"
    Write-Host "4. Check user recording permissions"
    Write-Host "5. Exit"
    $choice = Read-Host "Enter your choice"

    switch ($choice) {
        "1" {
            $userPrincipalName = Read-Host "Enter the user's email address"
            Get-PolicyDetails -userPrincipalName $userPrincipalName
        }
        "2" {
            $policyName = Read-Host "Enter the policy name"
            Get-UsersWithPolicy -PolicyName $policyName
        }
        "3" {
            $policyName = Read-Host "Enter the policy name"
            $allowWhiteboard = Read-Host "Allow whiteboard access? (true/false)"
            Manage-WhiteboardAccess -PolicyName $policyName -AllowWhiteboard $allowWhiteboard
        }
        "4" {
            $userPrincipalName = Read-Host "Enter the user's email address"
            $permission = Get-UserRecordingPermissions -UserPrincipalName $userPrincipalName
            Write-Log "User recording permission for '$userPrincipalName' is '$permission'." -Level Info
        }
        "5" {
            Write-Log "Exiting script." -Level Info
        }
        default {
            Write-Log "Invalid choice. Please select a valid option." -Level Warning
        }
    }
} while ($choice -ne "5")

Disconnect-MicrosoftTeams
