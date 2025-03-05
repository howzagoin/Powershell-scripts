# Import required modules
Import-Module Microsoft.Graph -ErrorAction Stop
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Logging utility
Function Write-CustomLog {
    Param(
        [string]$Message,
        [string]$Level = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp][$Level] $Message"
}

# Function to connect to Microsoft 365 services
Function Connect-M365Services {
    [CmdletBinding()]
    Param()

    # Disconnect any existing sessions
    Try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-CustomLog "Disconnected any existing Microsoft 365 sessions." -Level "Info"
    } Catch {
        Write-CustomLog "No active sessions to disconnect or an error occurred during disconnect: $_" -Level "Warning"
    }

    Try {
        # Connect to Microsoft Graph using dynamic scopes
        Connect-MgGraph -Scopes @(
            "User.Read.All",
            "Directory.Read.All",
            "Group.Read.All",
            "Application.Read.All",
            "AuditLog.Read.All",
            "Organization.Read.All",
            "RoleManagement.Read.Directory"
        ) -ErrorAction Stop
        Write-CustomLog "Successfully authenticated to Microsoft Graph using browser authentication." -Level "Info"

        # Output and log the scopes for testing
        $activeScopes = (Get-MgContext).Scopes
        Write-CustomLog "Active Microsoft Graph scopes: $($activeScopes -join ', ')" -Level "Info"
        Write-Host "Active Microsoft Graph scopes: $($activeScopes -join ', ')" -ForegroundColor Cyan

        # Connect to Exchange Online
        Connect-ExchangeOnline -ShowProgress:$true -ErrorAction Stop
        Write-CustomLog "Successfully authenticated to Exchange Online using browser authentication." -Level "Info"

        return $true
    } Catch {
        Write-CustomLog "Authentication failed: $_" -Level "Error"
        Write-Host "Unable to authenticate. Please check your credentials and try again." -ForegroundColor Red
        return $false
    }
}

# Function to get user information
Function Get-UserInfo {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    try {
        Write-CustomLog "Starting user information retrieval for $UserPrincipalName" -Level "Info"

        # Check for active session and scopes
        $context = Get-MgContext
        if (-not $context) {
            Write-CustomLog "No active Microsoft Graph session. Connecting now..." -Level "Warning"
            if (-not (Connect-M365Services)) {
                Write-Host "Failed to connect to Microsoft Graph. Cannot proceed." -ForegroundColor Red
                return $null
            }
        }

        # Ensure required scopes are granted
        $requiredScopes = @("User.Read.All", "Directory.Read.All")
        $missingScopes = $requiredScopes | Where-Object { -not ($context.Scopes -contains $_) }
        if ($missingScopes) {
            Write-CustomLog "Missing required scopes: $($missingScopes -join ', '). Re-authenticating..." -Level "Warning"
            Disconnect-MgGraph
            Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop
        }

        # Define properties to retrieve
        $properties = @(
            "DisplayName",
            "UserPrincipalName",
            "JobTitle",
            "Department",
            "MobilePhone",
            "OfficeLocation",
            "AccountEnabled",
            "AssignedLicenses",
            "StrongAuthenticationMethods",
            "MailboxSettings",
            "CreatedDateTime",
            "LastSignInDateTime",
            "ProxyAddresses",
            "OtherMails",
            "MailNickname",
            "Drive",
            "Devices",
            "MemberOf",
            "LicenseDetails"
        )

        # Fetch user details
        $user = Get-MgUser -UserId $UserPrincipalName -Property $properties -ErrorAction Stop
        if (-not $user) {
            Write-CustomLog "No user data found for $UserPrincipalName" -Level "Warning"
            return $null
        }

        Write-CustomLog "User found: $($user.DisplayName)" -Level "Info"

        # Collect additional user details
        $userInfo = [PSCustomObject]@{
            DisplayName         = $user.DisplayName
            Email               = $user.UserPrincipalName
            Groups              = (Get-MgUserMemberOf -UserId $user.Id -All |
                                    Where-Object { $_.'@odata.type' -eq "#microsoft.graph.group" }).DisplayName -join ", "
            AssignedLicenses    = $user.AssignedLicenses.SkuId -join ", "
            MFAStatus           = if ($user.StrongAuthenticationMethods) { "Enabled" } else { "Disabled" }
            AccountStatus       = if ($user.AccountEnabled) { "Enabled" } else { "Disabled" }
            CreatedDate         = $user.CreatedDateTime
            LastSignInDate      = $user.LastSignInDateTime
            MobilePhone         = $user.MobilePhone
            JobTitle            = $user.JobTitle
            Department          = $user.Department
            ProxyAddresses      = $user.ProxyAddresses -join ", "
            OtherEmails         = $user.OtherMails -join ", "
            OneDriveUrl         = $user.Drive?.WebUrl
        }

        Write-CustomLog "User details retrieved successfully for $UserPrincipalName" -Level "Info"
        return $userInfo
    } Catch {
        Write-CustomLog "Error retrieving user info for ${UserPrincipalName}: $_" -Level "Error"
        Write-CustomLog "Detailed error retrieving user info: $($_.Exception.Message)" -Level "Error"
        throw
    }
}

# Main script to retrieve user info
$targetUser = "amy.clarke@firstfinancial.com.au"
$userInfo = Get-UserInfo -UserPrincipalName $targetUser
if ($userInfo) {
    Write-Host "`nUser Information for ${36735
    targetUser}:" -ForegroundColor Green
    $userInfo | Format-List
} else {
    Write-Host "Failed to retrieve user information for $targetUser." -ForegroundColor Red
}
