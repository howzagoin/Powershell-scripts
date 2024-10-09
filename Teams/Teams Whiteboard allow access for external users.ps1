<#
.SYNOPSIS
    This PowerShell script is designed to update Microsoft Teams and SharePoint Online settings related to the use of whiteboards in meetings.

.DESCRIPTION
    This script performs the following actions:
    1. Connects to Microsoft Teams using interactive login (supports MFA).
    2. Retrieves and displays the current state of the 'AllowWhiteboard' setting for the Teams meeting policy.
    3. Updates the Teams meeting policy to enable the use of whiteboards for external and anonymous users.
    4. Retrieves and displays the updated state of the 'AllowWhiteboard' setting.
    5. Connects to SharePoint Online using interactive login (supports MFA).
    6. Retrieves and displays the current SharePoint tenant settings for 'AllowAnonymousMeetingParticipantsToAccessWhiteboards' and 'IsWBFluidEnabled'.
    7. Updates the SharePoint tenant settings to enable anonymous access to whiteboards and ensure that Fluid Whiteboard is enabled.
    8. Retrieves and displays the updated SharePoint tenant settings.
    9. Disconnects from both Microsoft Teams and SharePoint Online services.
    
.PARAMETER adminSiteURL
    The URL of the SharePoint Admin site. For example, https://yourdomain-admin.sharepoint.com.

.NOTES
    Ensure that you have the necessary administrative permissions to modify Teams and SharePoint settings.
    This script requires MFA-enabled accounts for Microsoft Teams and SharePoint Online.
    Make sure the MicrosoftTeams and SharePoint Online PowerShell modules are installed on your system.
#>

# Import required modules
Import-Module MicrosoftTeams
Import-Module Microsoft.Online.SharePoint.PowerShell

# Connect to Microsoft Teams interactively (supports MFA)
Write-Host "Connecting to Microsoft Teams..."
Connect-MicrosoftTeams

# Retrieve and display current Teams meeting policy settings
$policyName = "Global"  # Replace with your policy name if different
$policy = Get-CsTeamsMeetingPolicy -Identity $policyName
$initialWhiteboardState = $policy.AllowWhiteboard
Write-Host "Current state of 'AllowWhiteboard' for policy '$policyName': $initialWhiteboardState" -ForegroundColor Yellow

# Update the policy to allow Whiteboard
Set-CsTeamsMeetingPolicy -Identity $policyName -AllowWhiteboard $true

# Retrieve and display updated Teams meeting policy settings
$updatedPolicy = Get-CsTeamsMeetingPolicy -Identity $policyName
$updatedWhiteboardState = $updatedPolicy.AllowWhiteboard
Write-Host "Updated state of 'AllowWhiteboard' for policy '$policyName': $updatedWhiteboardState" -ForegroundColor Green

# Display change
Write-Host "The state of 'AllowWhiteboard' has been updated from $initialWhiteboardState to $updatedWhiteboardState." -ForegroundColor Cyan

# Disconnect from Microsoft Teams
Disconnect-MicrosoftTeams

# Connect to SharePoint Online interactively (supports MFA)
Write-Host "Connecting to SharePoint Online..."
$adminSiteURL = Read-Host "Enter your SharePoint Admin site URL (e.g., https://yourdomain-admin.sharepoint.com)"
Connect-SPOService -Url $adminSiteURL

# Retrieve and display current SharePoint settings
$tenantSettings = Get-SPOTenant
$initialWhiteboardAccess = $tenantSettings.AllowAnonymousMeetingParticipantsToAccessWhiteboards
$initialWBFluidState = $tenantSettings.IsWBFluidEnabled
Write-Host "Current state of 'AllowAnonymousMeetingParticipantsToAccessWhiteboards': $initialWhiteboardAccess" -ForegroundColor Yellow
Write-Host "Current state of 'IsWBFluidEnabled': $initialWBFluidState" -ForegroundColor Yellow

# Update SharePoint settings
Set-SPOTenant -AllowAnonymousMeetingParticipantsToAccessWhiteboards "On"
Set-SPOTenant -IsWBFluidEnabled $true

# Retrieve and display updated SharePoint settings
$updatedTenantSettings = Get-SPOTenant
$updatedWhiteboardAccess = $updatedTenantSettings.AllowAnonymousMeetingParticipantsToAccessWhiteboards
$updatedWBFluidState = $updatedTenantSettings.IsWBFluidEnabled
Write-Host "Updated state of 'AllowAnonymousMeetingParticipantsToAccessWhiteboards': $updatedWhiteboardAccess" -ForegroundColor Green
Write-Host "Updated state of 'IsWBFluidEnabled': $updatedWBFluidState" -ForegroundColor Green

# Display change
Write-Host "The state of 'AllowAnonymousMeetingParticipantsToAccessWhiteboards' has been updated from $initialWhiteboardAccess to $updatedWhiteboardAccess." -ForegroundColor Cyan
Write-Host "The state of 'IsWBFluidEnabled' has been updated from $initialWBFluidState to $updatedWBFluidState." -ForegroundColor Cyan

# Disconnect from SharePoint Online
Disconnect-SPOService

# End of script
Write-Host "SharePoint and Teams sessions closed. Script execution completed." -ForegroundColor Cyan
