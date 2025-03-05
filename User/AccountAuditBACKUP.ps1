# Metadata
<#
.SYNOPSIS
Collects detailed user information, including account status, MFA status, mailbox access, and more.

.AUTHOR
Tim MacLatchy

.DATE
26-11-2024

.LICENSE
MIT License

.DESCRIPTION
Retrieves detailed user information from Microsoft 365, including user account details, group memberships, mailbox access, calendar permissions, MFA status, and enterprise applications linked to the user.
#>

[CmdletBinding()]
param()

# Function: Write-Log
Function Write-Log {
    [CmdletBinding()]
    Param (
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Critical')] 
        [string]$Level
    )
    
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    Write-Host $logMessage
}

# Function: Ensure-Modules
Function Ensure-Modules {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string[]]$RequiredModules
    )
    
    foreach ($module in $RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Log "Module $module is not installed. Installing..." -Level Info
            Install-Module -Name $module -Force -AllowClobber
        }
        else {
            Write-Log "Module $module is already installed." -Level Info
        }
    }
}

Function Test-RequiredRoles {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [array]$RequiredRoles
    )

    try {
        # Get the currently authenticated user's details using UserPrincipalName
        $currentUser = Get-MgUser -UserId "me" # "me" retrieves the authenticated user

        # Ensure the user was retrieved
        if ($null -eq $currentUser) {
            Write-Error "Failed to retrieve current user details."
            return $false
        }

        # Get the roles assigned to the authenticated user
        $roles = Get-MgUserAppRoleAssignment -UserId $currentUser.Id | Select-Object -ExpandProperty AppRoleId

        # Check if the required roles are assigned to the user
        foreach ($role in $RequiredRoles) {
            if ($roles -notcontains $role) {
                Write-Host "User does not have required role: $role"
                return $false
            }
        }

        Write-Host "User has all the required roles."
        return $true
    }
    catch {
        Write-Error "An error occurred while checking the roles: $_"
        return $false
    }
}

# Function: Test-GraphPermissions
Function Test-GraphPermissions {
    [CmdletBinding()]
    Param()
    
    try {
        Write-Log "Checking Microsoft Graph permissions..." -Level Info
        
        # Define required Graph permissions
        $requiredPermissions = @(
            "User.Read.All",
            "Group.Read.All",
            "Directory.Read.All",
            "Application.Read.All",
            "AuditLog.Read.All",
            "Sites.Read.All",
            "MailboxSettings.Read",
            "Calendars.Read"
        )
        
        # Get current permissions
        $context = Get-MgContext
        $currentPermissions = $context.Scopes
        
        # Check for missing permissions
        $missingPermissions = $requiredPermissions | Where-Object { $_ -notin $currentPermissions }
        
        if ($missingPermissions) {
            Write-Log "Missing required Graph permissions: $($missingPermissions -join ', ')" -Level Warning
            return $false
        }
        else {
            Write-Log "All required Graph permissions are granted." -Level Info
            return $true
        }
    }
    catch {
        Write-Log "Error checking Graph permissions: $_" -Level Error
        throw
    }
}

# Function: Connect-M365Services
Function Connect-M365Services {
    [CmdletBinding()]
    Param()
    
    try {
        $modules = @('ExchangeOnlineManagement', 'Microsoft.Graph')
        Ensure-Modules -RequiredModules $modules
        
        # Connect to Exchange Online with web login
        Write-Log "Connecting to Exchange Online using web login..." -Level Info
        Connect-ExchangeOnline -ShowProgress $true -ErrorAction Stop
        
        # Connect to Microsoft Graph with all required permissions
        Write-Log "Connecting to Microsoft Graph..." -Level Info
        $graphScopes = @(
            "User.Read.All",
            "Group.Read.All",
            "Directory.Read.All",
            "Application.Read.All",
            "AuditLog.Read.All",
            "Sites.Read.All",
            "MailboxSettings.Read",
            "Calendars.Read"
        )
        
        # Check if we're already connected to Microsoft Graph
        if (-not (Get-MgContext)) {
            Connect-MgGraph -Scopes $graphScopes
        }
        
        # Ensure the session is valid
        if (-not (Get-MgContext)) {
            Write-Log "Re-authentication required with additional permissions/roles." -Level Warning
            Disconnect-ExchangeOnline -Confirm:$false
            Disconnect-MgGraph
            return $false
        }
        
        Write-Log "Successfully connected to Microsoft 365 services" -Level Info
        return $true
    }
    catch {
        Write-Log "Failed to connect to Microsoft 365 services: $_" -Level Error
        throw
    }
}


# Function to get User Information including all requested fields
Function Get-UserInfo {
    [CmdletBinding()]
    Param(
        [string]$Username,
        [switch]$AllUsers,
        [switch]$ActiveMembersOnly,
        [int]$BatchSize = 100
    )

    Write-Log "Starting user information retrieval" -Level Info

    $baseFilter = switch ($true) {
        $ActiveMembersOnly { "accountEnabled eq true and userType eq 'Member'" }
        $Username { "userPrincipalName eq '$Username'" }
        $AllUsers { $null }
        default { 
            Write-Log "No valid filter criteria provided" -Level Warning
            return @()
        }
    }

    try {
        $results = [System.Collections.ArrayList]::new()
        $processedCount = 0
        
        # Get users in batches
        $users = if ($baseFilter) {
            Get-MgUser -Filter $baseFilter -All -PageSize $BatchSize -ErrorAction Stop
        }
        else {
            Get-MgUser -All -PageSize $BatchSize -ErrorAction Stop
        }

        foreach ($user in $users) {
            $processedCount++ 

            # Retrieve additional user information
            $groups = (Get-MgUserMemberOf -UserId $user.Id | Select-Object -ExpandProperty DisplayName) -join ", "
            $sharePointSite = (Get-MgUser -UserId $user.Id | Select-Object -ExpandProperty Mail) # Adjust according to SharePoint API
            $sharedMailboxes = (Get-Mailbox -Identity $user.UserPrincipalName | Select-Object -ExpandProperty SharedMailboxes) -join ", "
            $calendarAccess = (Get-MailboxFolderPermission -Identity "$($user.UserPrincipalName):\Calendar" | Select-Object -ExpandProperty User) -join ", "
            $assignedLicenses = ($user.AssignedLicenses | ForEach-Object { $_.SkuId }) -join ", "
            $mfaStatus = (Get-MfaStatus -UserPrincipalName $user.UserPrincipalName).Status
            $accountType = $user.UserType
            $mailboxType = (Get-Mailbox -Identity $user.UserPrincipalName).MailboxType
            $accountStatus = if ($user.AccountEnabled) { "Enabled" } else { "Disabled" }
            $enterpriseApps = (Get-MgUserAppRoleAssignment -UserId $user.Id | Select-Object -ExpandProperty AppDisplayName) -join ", "
            $linkedDevices = (Get-MgUserDevice -UserId $user.Id | Select-Object -ExpandProperty DisplayName) -join ", "
            $adminRoles = (Get-MgUserAppRoleAssignment -UserId $user.Id | Where-Object { $_.RoleTemplateId -match "admin" } | Select-Object -ExpandProperty RoleDisplayName) -join ", "
            $createdDate = $user.CreatedDateTime
            $lastSignInDate = (Get-MgAuditLogSignIn -Filter "userId eq '$($user.Id)'" -Top 1 | Select-Object -ExpandProperty CreatedDateTime)
            $lastPasswordChange = (Get-MgUserAuthentication -UserId $user.Id | Select-Object -ExpandProperty LastPasswordChangeDateTime)
            $accountDisableDate = if ($user.AccountEnabled -eq $false) { $user.LastSignInDateTime } else { $null }
            $mobilePhone = $user.MobilePhone
            $businessPhone = $user.BusinessPhones -join ", "
            $jobTitle = $user.JobTitle
            $manager = (Get-MgUserManager -UserId $user.Id).DisplayName
            $email = $user.Mail
            $city = $user.City
            $country = $user.Country
            
            # Add the information to results
            $results.Add([PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                AccountStatus = $accountStatus
                AccountType = $accountType
                AccountCreatedDate = $createdDate
                LastSignInDate = $lastSignInDate
                LastPasswordChangeDate = $lastPasswordChange
                AccountDisableDate = $accountDisableDate
                MFAStatus = $mfaStatus
                MailboxType = $mailboxType
                AssignedLicenses = $assignedLicenses
                MailboxOwner = $sharePointSite
                Groups = $groups
                SharedMailboxes = $sharedMailboxes
                CalendarAccess = $calendarAccess
                EnterpriseApps = $enterpriseApps
                LinkedDevices = $linkedDevices
                AdminRoles = $adminRoles
                MobilePhone = $mobilePhone
                BusinessPhone = $businessPhone
                JobTitle = $jobTitle
                Manager = $manager
                Email = $email
                City = $city
                Country = $country
            })
        }
        
        Write-Log "$processedCount users processed" -Level Info
        return $results
    }
    catch {
        Write-Log "An error occurred while retrieving user information: $_" -Level Error
        throw
    }
}

# Entry Point
Connect-M365Services
$results = Get-UserInfo -AllUsers

# Export to Excel (example)
$exportPath = "C:\Users\ExportedUserInfo.xlsx"
$results | Export-Excel -Path $exportPath -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WrapText

Write-Log "User information export complete." -Level Info
