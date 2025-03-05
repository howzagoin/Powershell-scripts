<#
    Author: Tim MacLatchy
    Date: 2024-09-26
    License: MIT License
    Description: This script sets delegate calendar access for a specified user and access level in Exchange Online.
#>

# Ensure required module is installed and imported
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing ExchangeOnlineManagement module..."
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
}

Import-Module ExchangeOnlineManagement

# Sign in to Exchange Online with MFA
Write-Host "Logging into Exchange Online with MFA..."
try {
    $adminUser = Read-Host "Enter admin username"
    Connect-ExchangeOnline -UserPrincipalName $adminUser -ShowProgress $true -ErrorAction Stop
    Write-Host "Successfully logged into Exchange Online."
} catch {
    Write-Host "Error logging into Exchange Online: $_"
    exit
}

# Prompt for user inputs
$calendarOwner = Read-Host "Enter the email of the user whose calendar to modify (calendar owner)"
$delegate = Read-Host "Enter the email of the user to give delegate access to"
$accessType = Read-Host "Enter the access level (e.g., Editor, Reviewer, Author)"

# Function to display current delegate access
function Get-DelegateAccess {
    param (
        [string]$calendarOwner
    )

    Write-Host "`nGetting current delegate access for $calendarOwner's calendar..."

    try {
        $permissions = Get-MailboxFolderPermission -Identity "${calendarOwner}:\Calendar" -ErrorAction Stop
        $permissions | ForEach-Object {
            Write-Host "User: $($_.User) | AccessRights: $($_.AccessRights)"
        }
    } catch {
        Write-Host "Error getting delegate access for $calendarOwner's calendar: $_"
    }
}

# Output current delegate access before changes
Write-Host "`nCurrent delegate access:"
Get-DelegateAccess -calendarOwner $calendarOwner

# Confirm changes before proceeding
$confirmation = Read-Host "Proceed with giving $delegate $accessType access to $calendarOwner's calendar? (yes/no)"
if ($confirmation -ne "yes") {
    Write-Host "Operation cancelled."
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# Function to add delegate access with error handling
function Add-DelegateAccess {
    param (
        [string]$calendarOwner,
        [string]$delegate,
        [string]$accessType
    )

    Write-Host "`nSetting delegate access for $delegate on $calendarOwner's calendar..."

    try {
        # Set delegate permissions
        Add-MailboxFolderPermission -Identity "${calendarOwner}:\Calendar" -User $delegate -AccessRights $accessType -ErrorAction Stop
        Write-Host "Delegate access successfully set for $delegate on $calendarOwner's calendar."
    } catch {
        Write-Host "Error setting delegate access for $delegate on $calendarOwner's calendar: $_"
    }
}

# Add the delegate access
Add-DelegateAccess -calendarOwner $calendarOwner -delegate $delegate -accessType $accessType

# Output current delegate access after changes
Write-Host "`nUpdated delegate access:"
Get-DelegateAccess -calendarOwner $calendarOwner

# Disconnect from Exchange Online
Write-Host "`nDisconnecting from Exchange Online..."
Disconnect-ExchangeOnline -Confirm:$false

Write-Host "Delegate access setup complete."
