# Author: Tim MacLatchy (Modified by Claude)
# Date: 14/08/2024
# License: MIT License
# Description: Enhanced script for managing Exchange Online calendar permissions with MFA authentication.
# Requirements: ExchangeOnlineManagement module

#Requires -Version 5.1
#Requires -Modules ExchangeOnlineManagement

[CmdletBinding()]
param()

# Define permission levels as an enumeration for better type safety
$script:PermissionLevels = @{
    Owner = "Owner"
    PublishingEditor = "PublishingEditor"
    Editor = "Editor"
    PublishingAuthor = "PublishingAuthor"
    Author = "Author"
    NoneditingAuthor = "NoneditingAuthor"
    Reviewer = "Reviewer"
    Contributor = "Contributor"
    AvailabilityOnly = "AvailabilityOnly"
    LimitedDetails = "LimitedDetails"
}

function Test-EmailAddress {
    param(
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress
    )
    
    $emailRegex = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return $EmailAddress -match $emailRegex
}

function Install-RequiredModules {
    [CmdletBinding()]
    param()
    
    Write-Verbose "Checking for required modules..."
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        try {
            Write-Host "Installing ExchangeOnlineManagement module..."
            Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser -ErrorAction Stop
            Write-Host "Module installed successfully." -ForegroundColor Green
        }
        catch {
            throw "Failed to install required module: $_"
        }
    }
}

function Connect-ExchangeOnlineMFA {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AdminEmail
    )
    
    try {
        Write-Host "Connecting to Exchange Online with MFA..."
        Connect-ExchangeOnline -UserPrincipalName $AdminEmail -ShowProgress $true -ErrorAction Stop
        Write-Host "Connected successfully to Exchange Online." -ForegroundColor Green
    }
    catch {
        throw "Failed to connect to Exchange Online: $_"
    }
}

function Get-ValidatedInput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,
        
        [Parameter(Mandatory = $false)]
        [scriptblock]$ValidationScript
    )
    
    do {
        $input = Read-Host -Prompt $Prompt
        if ($ValidationScript -and -not (& $ValidationScript $input)) {
            Write-Host "Invalid input. Please try again." -ForegroundColor Yellow
            continue
        }
        break
    } while ($true)
    
    return $input
}

function Select-PermissionLevel {
    [CmdletBinding()]
    param()
    
    Write-Host "`nAvailable permission levels:"
    $script:PermissionLevels.GetEnumerator() | Sort-Object Name | ForEach-Object {
        $i = [array]::IndexOf(($script:PermissionLevels.Keys | Sort-Object), $_.Key) + 1
        Write-Host "$i. $($_.Key)"
    }
    
    do {
        $selection = Read-Host "`nEnter the number corresponding to the desired permission level"
        $index = [int]$selection - 1
        
        if ($index -ge 0 -and $index -lt $script:PermissionLevels.Count) {
            return ($script:PermissionLevels.GetEnumerator() | Sort-Object Name)[$index].Value
        }
        Write-Host "Invalid selection. Please try again." -ForegroundColor Yellow
    } while ($true)
}

function Grant-CalendarPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OwnerEmail,
        
        [Parameter(Mandatory = $true)]
        [string]$DelegateEmail,
        
        [Parameter(Mandatory = $true)]
        [string]$Permission
    )
    
    try {
        Write-Host "`nGranting $Permission permission to $DelegateEmail on $OwnerEmail's calendar..."
        
        # Check if permission already exists
        $existingPermission = Get-MailboxFolderPermission -Identity "${OwnerEmail}:\Calendar" -User $DelegateEmail -ErrorAction SilentlyContinue
        
        if ($existingPermission) {
            Set-MailboxFolderPermission -Identity "${OwnerEmail}:\Calendar" -User $DelegateEmail -AccessRights $Permission -ErrorAction Stop
            Write-Host "Permissions successfully updated." -ForegroundColor Green
        }
        else {
            Add-MailboxFolderPermission -Identity "${OwnerEmail}:\Calendar" -User $DelegateEmail -AccessRights $Permission -ErrorAction Stop
            Write-Host "Permissions successfully granted." -ForegroundColor Green
        }
    }
    catch {
        throw "Failed to manage calendar permissions: $_"
    }
}

# Main script execution
try {
    # Install required modules
    Install-RequiredModules
    
    # Get admin email and connect
    $adminEmail = Get-ValidatedInput -Prompt "Enter the admin email address for MFA login" -ValidationScript { param($email) Test-EmailAddress $email }
    Connect-ExchangeOnlineMFA -AdminEmail $adminEmail
    
    # Get user emails
    $ownerEmail = Get-ValidatedInput -Prompt "Enter the email address of the user whose calendar is being shared" -ValidationScript { param($email) Test-EmailAddress $email }
    $delegateEmail = Get-ValidatedInput -Prompt "Enter the email address of the user who will be given access" -ValidationScript { param($email) Test-EmailAddress $email }
    
    # Get permission level
    $selectedPermission = Select-PermissionLevel
    
    # Grant permissions
    Grant-CalendarPermission -OwnerEmail $ownerEmail -DelegateEmail $delegateEmail -Permission $selectedPermission
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
}
finally {
    if (Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" }) {
        Write-Host "`nDisconnecting from Exchange Online..."
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "Session disconnected." -ForegroundColor Green
    }
}

Write-Host "`nScript execution completed." -ForegroundColor Blue