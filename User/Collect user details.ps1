# Custom Logging
Function Write-CustomLog {
    [CmdletBinding()]
    Param(
        [string]$Message,
        [string]$Level = 'Info',
        [string]$LogPath = "M365_Audit_$(Get-Date -Format 'yyyyMMdd').log"
    )

    $colorMap = @{
        'Info'    = 'Green'
        'Warning' = 'Yellow'
        'Error'   = 'Red'
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp][$Level] $Message"
    
    # Write to console with color
    Write-Host $logMessage -ForegroundColor $colorMap[$Level]
    
    # Write to log file
    try {
        Add-Content -Path $LogPath -Value $logMessage -ErrorAction Stop
    }
    catch {
        Write-Host "Failed to write to log file: $_" -ForegroundColor Red
    }
}
# Install Required Modules
Function Install-RequiredModules {
    [CmdletBinding()]
    Param(
        [string[]]$ModuleNames
    )
    
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
    $scope = if ($isAdmin) { "AllUsers" } else { "CurrentUser" }

    foreach ($moduleName in $ModuleNames) {
        try {
            Write-CustomLog "Checking module: $moduleName" -Level Info
            if (-not (Get-Module -Name $moduleName -ListAvailable)) {
                Write-CustomLog "Installing module: $moduleName" -Level Info
                Install-Module -Name $moduleName -Scope $scope -Force -AllowClobber
            }
            
            $currentVersion = (Get-Module -Name $moduleName -ListAvailable | Sort-Object Version -Descending)[0].Version
            $latestVersion = (Find-Module -Name $moduleName).Version
            
            if ($currentVersion -lt $latestVersion) {
                Write-CustomLog "Updating $moduleName from $currentVersion to $latestVersion" -Level Info
                Update-Module -Name $moduleName -Force
            }
            
            Import-Module -Name $moduleName -Force -ErrorAction Stop
            Write-CustomLog "Module $moduleName imported successfully" -Level Info
        }
        catch {
            Write-CustomLog "Error processing module ${moduleName}: $_" -Level Error
            throw
        }
    }
}
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
# Prompt for File Save Location
Function Prompt-SaveFileDialog {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$DefaultFileName,

        [Parameter()]
        [string]$InitialDirectory = ([Environment]::GetFolderPath("Desktop")),

        [Parameter()]
        [string]$Filter = "Excel Files (*.xlsx)|*.xlsx"
    )

    try {
        # Load Windows Forms if not already loaded
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

        # Create the save file dialog
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.InitialDirectory = $InitialDirectory
        $saveFileDialog.Filter = $Filter
        $saveFileDialog.FileName = $DefaultFileName
        $saveFileDialog.OverwritePrompt = $true

        # Show the dialog and return the selected file path
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Write-CustomLog "User selected file path: $($saveFileDialog.FileName)" -Level Info
            return $saveFileDialog.FileName
        }

        Write-CustomLog "File save operation canceled by user" -Level Warning
        return $null
    }
    catch {
        Write-CustomLog "Error in save file dialog: $_" -Level Error
        return $null
    }
}
Function Export-ResultsToExcel {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [array]$Data,

        [Parameter(Mandatory)]
        [string]$FilePath
    )

    try {
        # Ensure directory exists
        $directory = Split-Path -Parent $FilePath
        if (-not (Test-Path -Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }

        # Create Excel package with formatting
        $excelParams = @{
            Path           = $FilePath
            AutoSize       = $true
            AutoFilter     = $true
            FreezeTopRow   = $true
            BoldTopRow     = $true
            TableName      = "UserAuditResults"
            WorksheetName  = "User Audit"
            TableStyle     = "Medium2"
            ErrorAction    = "Stop"
        }

        $Data | Export-Excel @excelParams

        if (Test-Path $FilePath) {
            Write-CustomLog "Data exported successfully to $FilePath" -Level Info
            return $true
        } else {
            throw "Export file not found after operation"
        }
    } catch {
        Write-CustomLog "Excel export error: $_" -Level Error
        return $false
    }
}
# Get-UserInfo function
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

        Write-CustomLog "User found: $($userData.DisplayName)" -Level "Info"

        # Collect additional user details
        $userInfo = [PSCustomObject]@{
            DisplayName         = $user.DisplayName
            Email               = $user.UserPrincipalName
            Groups              = (Get-MgUserMemberOf -UserId $user.Id -All |
                                    Where-Object { $_.'@odata.type' -eq "#microsoft.graph.group" }).DisplayName -join ", "
            SharePointSite      = "Not Implemented"
            SharedMailboxes     = "Not Implemented"
            CalendarAccess      = "Not Implemented"
            AssignedLicenses    = $user.AssignedLicenses.SkuId -join ", "
            MFAStatus           = if ($user.StrongAuthenticationMethods) { "Enabled" } else { "Disabled" }
            AccountType         = "User Mailbox"
            MailboxType         = $user.MailboxSettings?.MailboxType
            AccountStatus       = if ($user.AccountEnabled) { "Enabled" } else { "Disabled" }
            EnterpriseApps      = "Not Implemented"
            LinkedDevices       = "Not Implemented"
            AdminRoles          = "Not Implemented"
            CreatedDate         = $user.CreatedDateTime
            LastSignInDate      = $user.LastSignInDateTime
            LastPasswordChange  = "Not Available"   # Can use `PasswordLastSet` property if available
            AccountDisableDate  = "Not Implemented"
            MobilePhone         = $user.MobilePhone
            BusinessPhone       = $user.BusinessPhones -join ", "
            JobTitle            = $user.JobTitle
            Department          = $user.Department
            ProxyAddresses      = $user.ProxyAddresses -join ", "
            OtherEmails         = $user.OtherMails -join ", "
            MailNickname        = $user.MailNickname
            OneDriveUrl         = $user.Drive?.WebUrl
        }

        Write-CustomLog "User details retrieved successfully for $UserPrincipalName" -Level "Info"
        return $userInfo
    } Catch {
        Write-CustomLog "Error retrieving user info for ${UserPrincipalName}: $_" -Level "Error"
        Write-CustomLog "Detailed error retrieving user info: $($_.Exception.Message)" -Level "Error"
        Write-CustomLog "Error Details: $($_.ScriptStackTrace)" -Level "Error"
        throw
    }
}
# Menu for Account Selection
Function Show-Menu {
    [CmdletBinding()]
    Param()

    Write-Host "Select the user details to fetch:"
    Write-Host "1: Specific User Account"
    Write-Host "2: All Active Member Accounts"
    Write-Host "3: All Accounts"
    $selection = Read-Host "Enter selection (1/2/3)"

    switch ($selection) {
        1 {
            $username = Read-Host "Enter the user principal name (email)"
            return @{Type = 'Username'; Value = $username}
        }
        2 {
            return @{Type = 'ActiveMembers'; Value = $true}
        }
        3 {
            return @{Type = 'AllUsers'; Value = $true}
        }
        default {
            Write-CustomLog "Invalid selection. Exiting..." -Level Error
            exit
        }
    }
}
# Function to Display User Info in Console
Function Display-UserInfo {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [PSCustomObject]$UserData
    )

    $output = @"
------------------------------------------
User Details
------------------------------------------
DisplayName         : $($UserData.DisplayName)
Email               : $($UserData.Email)
Groups              : $($UserData.Groups -ne $null ? $UserData.Groups : "None")
SharePointSite      : $($UserData.SharePointSite)
SharedMailboxes     : $($UserData.SharedMailboxes)
CalendarAccess      : $($UserData.CalendarAccess)
AssignedLicenses    : $($UserData.AssignedLicenses -ne $null ? $UserData.AssignedLicenses : "None")
MFAStatus           : $($UserData.MFAStatus)
AccountType         : $($UserData.AccountType)
MailboxType         : $($UserData.MailboxType)
AccountStatus       : $($UserData.AccountStatus)
EnterpriseApps      : $($UserData.EnterpriseApps)
LinkedDevices       : $($UserData.LinkedDevices)
AdminRoles          : $($UserData.AdminRoles -ne $null ? $UserData.AdminRoles : "None")
CreatedDate         : $($UserData.CreatedDate -ne $null ? $UserData.CreatedDate : "Unknown")
LastSignInDate      : $($UserData.LastSignInDate)
LastPasswordChange  : $($UserData.LastPasswordChange -ne $null ? $UserData.LastPasswordChange : "Unknown")
AccountDisableDate  : $($UserData.AccountDisableDate)
MobilePhone         : $($UserData.MobilePhone -ne $null ? $UserData.MobilePhone : "None")
BusinessPhone       : $($UserData.BusinessPhone -ne $null ? $UserData.BusinessPhone : "None")
JobTitle            : $($UserData.JobTitle -ne $null ? $UserData.JobTitle : "None")
Department          : $($UserData.Department -ne $null ? $UserData.Department : "None")
ProxyAddresses      : $($UserData.ProxyAddresses -ne $null ? $UserData.ProxyAddresses : "None")
OtherEmails         : $($UserData.OtherEmails -ne $null ? $UserData.OtherEmails : "None")
MailNickname        : $($UserData.MailNickname)
OneDriveUrl         : $($UserData.OneDriveUrl)
------------------------------------------
"@

    Write-Host $output -ForegroundColor Cyan
}
Function Main {
    [CmdletBinding()]
    Param()

    # Always prompt for a fresh session using Connect-M365Services
    if (-not (Connect-M365Services)) {
        Write-CustomLog "Authentication failed. Exiting script." -Level "Error"
        return
    }

    Write-Host "Select the user details to fetch:"
    Write-Host "1: Specific User"
    Write-Host "2: All Active Member Accounts"
    Write-Host "3: All Accounts"
    $choice = Read-Host "Enter your choice (1, 2, or 3)"

    $results = switch ($choice) {
        "1" {
            $upn = Read-Host "Enter the user principal name (email)"
            try {
                Get-UserInfo -UserPrincipalName $upn
            } Catch {
                Write-CustomLog "Error retrieving user info: $_" -Level "Error"
                $null
            }
        }
        "2" {
            try {
                Get-MgUser -Filter "accountEnabled eq true and userType eq 'Member'" -All
            } Catch {
                Write-CustomLog "Error retrieving active member accounts: $_" -Level "Error"
                $null
            }
        }
        "3" {
            try {
                Get-MgUser -All
            } Catch {
                Write-CustomLog "Error retrieving all accounts: $_" -Level "Error"
                $null
            }
        }
        default {
            Write-CustomLog "Invalid selection" -Level "Warning"
            $null
        }
    }

    if (-not $results) {
        Write-CustomLog "No valid data retrieved" -Level "Warning"
    } else {
        Display-UserInfo -UserData $results

        $saveResults = Read-Host "Do you want to save the results to an Excel file? (Y/N)"
        if ($saveResults -eq "Y") {
            $defaultFileName = "UserDetails_$(Get-Date -Format 'dd-MM-yyyy').xlsx"
            $filePath = Prompt-SaveFileDialog -DefaultFileName $defaultFileName
            if ($filePath) {
                Export-ResultsToExcel -Data @($results) -FilePath $filePath
            } else {
                Write-CustomLog "Export operation canceled by user" -Level "Warning"
            }
        }
    }
}
# Start the Main function
Main

#add mailbox size