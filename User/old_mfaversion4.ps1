<#
    .SYNOPSIS
        Retrieve MFA status for specified Azure AD users.

    .DESCRIPTION
        This script logs into an Azure AD tenant, retrieves MFA status (Enabled, Enforced, or Disabled) for users, and provides additional user information.
        Prompts allow the user to choose whether to view details for a single user, internal users, or all users, and whether to save the results to an Excel document.

    .AUTHOR
        Tim MacLatchy

    .DATE
        01/11/2024

    .LICENSE
        MIT License
#>

# --- Imported from AccountAuditBACKUP.ps1 ---
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
        } else {
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
        $currentUser = Get-MgUser -UserId "me"
        if ($null -eq $currentUser) {
            Write-Error "Failed to retrieve current user details."
            return $false
        }
        $roles = Get-MgUserAppRoleAssignment -UserId $currentUser.Id | Select-Object -ExpandProperty AppRoleId
        foreach ($role in $RequiredRoles) {
            if ($roles -notcontains $role) {
                Write-Host "User does not have required role: $role"
                return $false
            }
        }
        Write-Host "User has all the required roles."
        return $true
    } catch {
        Write-Error "An error occurred while checking the roles: $_"
        return $false
    }
}

Function Test-GraphPermissions {
    [CmdletBinding()]
    Param()
    try {
        Write-Log "Checking Microsoft Graph permissions..." -Level Info
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
        $context = Get-MgContext
        $currentPermissions = $context.Scopes
        $missingPermissions = $requiredPermissions | Where-Object { $_ -notin $currentPermissions }
        if ($missingPermissions) {
            Write-Log "Missing required Graph permissions: $($missingPermissions -join ', ')" -Level Warning
            return $false
        } else {
            Write-Log "All required Graph permissions are granted." -Level Info
            return $true
        }
    } catch {
        Write-Log "Error checking Graph permissions: $_" -Level Error
        throw
    }
}

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
        $users = if ($baseFilter) {
            Get-MgUser -Filter $baseFilter -All -PageSize $BatchSize -ErrorAction Stop
        } else {
            Get-MgUser -All -PageSize $BatchSize -ErrorAction Stop
        }
        foreach ($user in $users) {
            $processedCount++ 
            $groups = (Get-MgUserMemberOf -UserId $user.Id | Select-Object -ExpandProperty DisplayName) -join ", "
            $sharePointSite = (Get-MgUser -UserId $user.Id | Select-Object -ExpandProperty Mail)
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
                Email = $user.Mail
                City = $user.City
                Country = $user.Country
            })
        }
        Write-Log "$processedCount users processed" -Level Info
        return $results
    } catch {
        Write-Log "An error occurred while retrieving user information: $_" -Level Error
        throw
    }
}

# --- Imported from fetches licensed inactive users.ps1 ---
Function Get-LicensedInactiveUsers {
    param (
        [int]$InactiveDays = 90
    )
    $inactivePeriod = (Get-Date).AddDays(-$InactiveDays)
    $allInactiveUsers = @()
    $users = Get-MgUser -All
    foreach ($user in $users) {
        $signInActivity = Get-MgUserAuthenticationMethodSignInActivity -UserId $user.Id -ErrorAction SilentlyContinue
        if ($null -ne $signInActivity) {
            if ($signInActivity.LastSignInDateTime -lt $inactivePeriod) {
                $licenseDetails = Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction SilentlyContinue
                $userDetails = [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName       = $user.DisplayName
                    LastSignIn        = $signInActivity.LastSignInDateTime
                    Licenses          = $licenseDetails.SkuPartNumber -join ", "
                }
                $allInactiveUsers += $userDetails
            }
        } else {
            $licenseDetails = Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction SilentlyContinue
            $userDetails = [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName       = $user.DisplayName
                LastSignIn        = "No sign-in activity found"
                Licenses          = $licenseDetails.SkuPartNumber -join ", "
            }
            $allInactiveUsers += $userDetails
        }
    }
    return $allInactiveUsers
}

# --- Imported from Get-ListOfActiveUsers.ps1 ---
Function Get-ActiveUsers {
    $allUsers = Get-MgUser -All
    $activeUsers = $allUsers | Where-Object { $_.AccountEnabled -eq $true -and $_.UserPrincipalName -notlike "*#EXT#*" }
    return $activeUsers | Select-Object DisplayName, UserPrincipalName, AccountEnabled
}

# --- Imported from LicensedDisabledInactiveUnknownUsers.ps1 ---
Function Fetch-UserLicenseStatus {
    param (
        [hashtable]$LicenseMap
    )
    $results = @{
        "Deactivated Users with Licenses" = @()
        "Inactive Users with Licenses" = @()
        "Unknown Users with Licenses" = @()
    }
    $users = Get-MgUser -All -Property Id, UserPrincipalName, DisplayName, AccountEnabled, AssignedLicenses, SignInActivity
    foreach ($user in $users) {
        if ($user.AssignedLicenses.Count -eq 0) { continue }
        $lastSignIn = $user.SignInActivity?.LastSignInDateTime
        $lastSignInDate = if ($lastSignIn) { [datetime]$lastSignIn } else { $null }
        $licenseList = $user.AssignedLicenses | ForEach-Object {
            $LicenseMap[$_.SkuId] -ne $null ? $LicenseMap[$_.SkuId] : $_.SkuId
        }
        $licenses = $licenseList -join ", "
        $userObject = [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            DisplayName       = $user.DisplayName
            LastSignInDate    = if ($lastSignInDate) { $lastSignInDate } else { "N/A" }
            Status            = if ($user.AccountEnabled) { "Active" } else { "Disabled" }
            Licenses          = $licenses
        }
        if (-not $user.AccountEnabled) {
            $results["Deactivated Users with Licenses"] += $userObject
        } elseif ($lastSignInDate -and ($lastSignInDate -lt (Get-Date).AddDays(-90))) {
            $results["Inactive Users with Licenses"] += $userObject
        } elseif (-not $lastSignInDate) {
            $results["Unknown Users with Licenses"] += $userObject
        }
    }
    return $results
}

# --- Imported from List-MFA-Disabled-Users.ps1 ---
Function Get-MFADisabledUsersFromCSV {
    param (
        [string]$CsvFilePath
    )
    $tenants = Import-Csv -Path $CsvFilePath
    $allMfaDisabledUsers = @()
    foreach ($tenant in $tenants) {
        $tenantId = $tenant."Microsoft ID"
        $users = Get-AzureADUser -All $true -TenantId $tenantId
        foreach ($user in $users) {
            $mfaStatus = Get-MsolUser -UserPrincipalName $user.UserPrincipalName -TenantId $tenantId | Select-Object -ExpandProperty StrongAuthenticationRequirements
            if ($mfaStatus.Count -eq 0) {
                $allMfaDisabledUsers += [PSCustomObject]@{
                    TenantID          = $tenantId
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName       = $user.DisplayName
                }
            }
        }
    }
    return $allMfaDisabledUsers
}

# --- Imported from userdetailstest.ps1 ---
Function Write-CustomLog {
    Param(
        [string]$Message,
        [string]$Level = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp][$Level] $Message"
}

Function Get-UserInfoDetailed {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )
    try {
        Write-CustomLog "Starting user information retrieval for $UserPrincipalName" -Level "Info"
        $context = Get-MgContext
        if (-not $context) {
            Write-CustomLog "No active Microsoft Graph session. Cannot proceed." -Level "Warning"
            return $null
        }
        $requiredScopes = @("User.Read.All", "Directory.Read.All")
        $missingScopes = $requiredScopes | Where-Object { -not ($context.Scopes -contains $_) }
        if ($missingScopes) {
            Write-CustomLog "Missing required scopes: $($missingScopes -join ', ')." -Level "Warning"
            return $null
        }
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
        $user = Get-MgUser -UserId $UserPrincipalName -Property $properties -ErrorAction Stop
        if (-not $user) {
            Write-CustomLog "No user data found for $UserPrincipalName" -Level "Warning"
            return $null
        }
        Write-CustomLog "User found: $($user.DisplayName)" -Level "Info"
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

# Import Required Modules
Function Install-ModuleIfMissing {
    param (
        [string]$ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Output "Module $ModuleName is not found. Installing..."
        Install-Module -Name $ModuleName -Force -Scope CurrentUser -ErrorAction Stop
    }
}

# Connect to Microsoft Graph with Web-based MFA Authentication
Function Connect-MicrosoftGraphWithMFA {
    param (
        [string]$AdminEmail
    )

    try {
        Write-Output "Connecting to Microsoft Graph for Azure AD data..."
        Connect-MgGraph -Scopes "User.Read.All", "Policy.ReadWrite.AuthenticationMethod", "UserAuthenticationMethod.Read.All"
        Write-Output "Successfully connected to Microsoft Graph."
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph. Please check your network connection and credentials."
        throw
    }
}

# Get User MFA Status
Function Get-UserMFAStatus {
    param (
        [string]$UserEmail
    )
    Write-Output "Retrieving MFA status for ${UserEmail}..."
    
    try {
        $User = Get-MgUser -UserId $UserEmail -ErrorAction Stop
        if (-not $User) {
            Write-Output "User ${UserEmail} not found in the directory."
            return
        }

        $MFAStatus = "Disabled"
        $AuthRequirements = (Get-MgUser -UserId $UserEmail -Property "StrongAuthenticationRequirements").StrongAuthenticationRequirements

        if ($AuthRequirements.Count -gt 0) {
            foreach ($Requirement in $AuthRequirements) {
                if ($Requirement.State -eq "Enabled") {
                    $MFAStatus = "Enabled"
                }
                elseif ($Requirement.State -eq "Enforced") {
                    $MFAStatus = "Enforced"
                }
            }
        }
        
        Write-Output "MFA Status for ${UserEmail}: ${MFAStatus}"
    }
    catch {
        Write-Error "Error retrieving MFA status for ${UserEmail}: $_"
    }
}

# Save results to Excel with formatting
Function Save-ResultsToExcel {
    param ([array]$results)
    Write-Host "Saving results to Excel..."
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $filePath = [System.Windows.Forms.SaveFileDialog]::new()
        $filePath.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $filePath.Title = "Save Excel File"
        
        if ($filePath.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $results | Export-Excel -Path $filePath.FileName -AutoSize -Title "User MFA and License Status" -TableName "UserDetails"

            # Adjust columns to auto-size and wrap text where there are multiple comma-separated entries
            $excel = New-Object -ComObject Excel.Application
            $workbook = $excel.Workbooks.Open($filePath.FileName)
            $worksheet = $workbook.Sheets.Item(1)

            # Apply auto-sizing and wrap text in cells with multiple comma-separated entries
            foreach ($column in $worksheet.UsedRange.Columns) {
                $column.EntireColumn.AutoFit()
                foreach ($cell in $column.Cells) {
                    if ($cell.Value() -match ",") {
                        $cell.WrapText = $true
                    }
                }
            }

            $workbook.Save()
            $workbook.Close()
            $excel.Quit()
            
            Write-Host "Results saved to $($filePath.FileName)" -ForegroundColor Green
        } else {
            Write-Host "Save operation cancelled." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "Failed to save results to Excel: $($_.Exception.Message)"
    }
}

# Main Script Execution
Function Main {
    Install-ModuleIfMissing -ModuleName "Microsoft.Graph"
    Install-ModuleIfMissing -ModuleName "ImportExcel"  # If exporting to Excel

    # Prompt for Admin Email
    $AdminEmail = Read-Host "Please enter your admin email address"

    # Connect to Microsoft Graph
    Connect-MicrosoftGraphWithMFA -AdminEmail $AdminEmail

    # Ask for user selection
    $Selection = Read-Host "Enter '1' to view MFA for a specific user, '2' for internal users only, '3' for all users"

    # Initialize the report list
    $Report = [System.Collections.Generic.List[Object]]::new()
    
    if ($Selection -eq "1") {
        # Single User Mode
        $UserEmail = Read-Host "Please enter the user's email address"
        Get-UserMFAStatus -UserEmail $UserEmail
    }
    else {
        # Retrieve User Properties
        $Properties = @('Id', 'DisplayName', 'UserPrincipalName', 'UserType', 'Mail', 'ProxyAddresses', 'AccountEnabled', 'CreatedDateTime')
        [array]$Users = Get-MgUser -All -Property $Properties | Select-Object $Properties

        # Check if any users were retrieved
        if (-not $Users) {
            Write-Host "No users found. Exiting script." -ForegroundColor Red
            return
        }

        # Filter based on internal or all users
        if ($Selection -eq "2") {
            $Users = $Users | Where-Object { $_.UserType -eq "Member" }
        }

        # Loop through each user and get their MFA settings
        $counter = 0
        $totalUsers = $Users.Count

        foreach ($User in $Users) {
            $counter++
            $percentComplete = [math]::Round(($counter / $totalUsers) * 100)
            $progressParams = @{
                Activity        = "Processing Users"
                Status          = "User $($counter) of $totalUsers - $($User.UserPrincipalName) - $percentComplete% Complete"
                PercentComplete = $percentComplete
            }

            Write-Progress @progressParams

            # Get MFA settings
            $MFAStateUri = "https://graph.microsoft.com/beta/users/$($User.Id)/authentication/requirements"
            $Data = Invoke-MgGraphRequest -Uri $MFAStateUri -Method GET

            # Get the default MFA method
            $DefaultMFAUri = "https://graph.microsoft.com/beta/users/$($User.Id)/authentication/signInPreferences"
            $DefaultMFAMethod = Invoke-MgGraphRequest -Uri $DefaultMFAUri -Method GET

            $MFAMethod = if ($DefaultMFAMethod.userPreferredMethodForSecondaryAuthentication) {
                Switch ($DefaultMFAMethod.userPreferredMethodForSecondaryAuthentication) {
                    "push" { "Microsoft authenticator app" }
                    "oath" { "Authenticator app or hardware token" }
                    "voiceMobile" { "Mobile phone" }
                    "voiceAlternateMobile" { "Alternate mobile phone" }
                    "voiceOffice" { "Office phone" }
                    "sms" { "SMS" }
                    Default { "Unknown method" }
                }
            } else {
                "Not Enabled"
            }

            # Create a report line for each user
            $ReportLine = [PSCustomObject][ordered]@{
                UserPrincipalName = $User.UserPrincipalName
                DisplayName       = $User.DisplayName
                MFAState          = $Data.PerUserMfaState
                MFADefaultMethod  = $MFAMethod
                PrimarySMTP       = $User.Mail
                Aliases           = ($User.ProxyAddresses | Where-Object { $_ -clike "smtp*" } | ForEach-Object { $_ -replace "smtp:", "" }) -join ', '
                UserType          = $User.UserType
                AccountEnabled    = $User.AccountEnabled
                CreatedDateTime   = $User.CreatedDateTime
            }
            $Report.Add($ReportLine)
        }

        # Output the report to console
        $Report | Format-Table -AutoSize
    }

    # Ask if user wants to save to Excel
    $SaveToExcel = Read-Host "Do you want to save the report to an Excel file? (Y/N)"
    if ($SaveToExcel -eq "Y") {
        # Get tenant name for filename
        $tenantName = (Get-MgOrganization | Select-Object -ExpandProperty DisplayName) -replace '[^a-zA-Z0-9_-]', '_'
        $dateStr = Get-Date -Format yyyyMMdd_HHmmss
        $defaultFileName = "${tenantName}_User&GroupAudit_${dateStr}.xlsx"
        Add-Type -AssemblyName PresentationFramework
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Title = "Save Excel File"
        $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $saveDialog.FileName = $defaultFileName
        $saveDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
        if ($saveDialog.ShowDialog() -eq $true) {
            $filePath = $saveDialog.FileName
            $Report | Export-Excel -Path $filePath -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter -TableStyle Medium6 -Title "MFA and User Status Report" -WorksheetName "MFA Report"
            Write-Host "Results saved to $filePath" -ForegroundColor Green
            Start-Process -FilePath $filePath
        } else {
            Write-Host "Save operation cancelled." -ForegroundColor Yellow
        }
    }
}

# Execute Main Function
Main
