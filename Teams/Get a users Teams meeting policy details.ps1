# Import the required module
Import-Module Microsoft.Graph

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All", "Policy.Read.All", "App.Read.All"

# Function to get policy details for a user
function Get-PolicyDetails {
    param (
        [string]$userPrincipalName
    )

    # Retrieve user details
    $user = Get-MgUser -UserId $userPrincipalName
    Write-Output "User Details:"
    Write-Output ""
    Write-Output "DisplayName       : $($user.DisplayName)"
    Write-Output "UserPrincipalName : $($user.UserPrincipalName)"
    Write-Output ""

    # Retrieve and display Teams meeting policy details
    $meetingPolicies = Get-MgTeamsMeetingPolicy
    Write-Output "Teams Meeting Policy Details:"
    $userMeetingPolicy = Get-CsTeamsMeetingPolicy -Identity $userPrincipalName
    Write-Output "User Teams Meeting Policy:"
    $userMeetingPolicy | Format-List
    Write-Output ""
    Write-Output "Available Teams Meeting Policies:"
    $meetingPolicies | Format-Table -Property Identity, Description, AllowChannelMeetingScheduling, AllowMeetNow, AllowPrivateMeetNow, MeetingChatEnabled, Type

    # Retrieve and display Teams app permission policy details
    $permissionPolicies = Get-MgTeamsAppPermissionPolicy
    Write-Output "Teams App Permission Policy Details:"
    $userPermissionPolicy = Get-CsTeamsAppPermissionPolicy -Identity $userPrincipalName
    Write-Output "User Teams App Permission Policy:"
    $userPermissionPolicy | Format-List
    Write-Output ""
    Write-Output "Available Teams App Permission Policies:"
    $permissionPolicies | Format-Table -Property Identity, DefaultCatalogApps, GlobalCatalogApps, PrivateCatalogApps, Description, DefaultCatalogAppsType, GlobalCatalogAppsType, PrivateCatalogAppsType

    # Retrieve and display Teams app setup policy details
    $setupPolicies = Get-MgTeamsAppSetupPolicy
    Write-Output "Teams App Setup Policy Details:"
    $userSetupPolicy = Get-CsTeamsAppSetupPolicy -Identity $userPrincipalName
    Write-Output "User Teams App Setup Policy:"
    $userSetupPolicy | Format-List
    Write-Output ""
    Write-Output "Available Teams App Setup Policies:"
    $setupPolicies | Format-Table -Property Identity, AppPresetList, PinnedAppBarApps

    # List allowed and banned Teams apps for the user
    $userAppPermissionPolicy = Get-MgTeamsAppPermissionPolicy -UserId $userPrincipalName
    Write-Output "Teams Apps Allowed for User:"
    $userAppPermissionPolicy.DefaultCatalogApps | Format-List
    Write-Output ""
    Write-Output "Teams Apps Banned for User:"
    $userAppPermissionPolicy.BlockedAppList | Format-List
    Write-Output ""

    # List apps banned and allowed during Teams video calls
    $meetingPolicy = $userMeetingPolicy | Select-Object -ExpandProperty MeetingPolicies
    Write-Output "Teams Apps Allowed During Video Calls:"
    $meetingPolicy.AllowPrivateMeetNow | Format-List
    Write-Output ""
    Write-Output "Teams Apps Banned During Video Calls:"
    $meetingPolicy.BlockedAppList | Format-List
}

# Prompt for user email
$userPrincipalName = Read-Host "Enter the email address of the user (e.g., david.jackson@firstfinancial.com.au)"

# Get policy details for the specified user
Get-PolicyDetails -userPrincipalName $userPrincipalName
