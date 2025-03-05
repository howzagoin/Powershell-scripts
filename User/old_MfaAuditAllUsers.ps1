<#
.SYNOPSIS
    Retrieve MFA status and related information for Azure AD users.

.DESCRIPTION
    This script retrieves detailed MFA-related user data, including default MFA methods, and allows scoped data collection. It forces an admin login each time via manual authentication.

.AUTHOR
    Tim MacLatchy

.DATE
    21-11-2024

.LICENSE
    MIT License
#>

# Function: Verify and Install Required Modules
Function Install-ModuleIfMissing {
    param ([string]$ModuleName)
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Module $ModuleName is not found. Installing..."
        Install-Module -Name $ModuleName -Force -Scope CurrentUser -ErrorAction Stop
    }
}

# Function: Connect to Microsoft Graph with Manual Login
Function Connect-MicrosoftGraphWithBrowserLogin {
    try {
        Write-Host "Connecting to Microsoft Graph with manual login..."
        # Manual login for Azure AD (MFA should be triggered by default)
        $UserCredential = Get-Credential
        Connect-MgGraph -Credential $UserCredential -Scopes "User.Read.All", "AuditLog.Read.All"
        Write-Host "Successfully connected to Microsoft Graph."
    } catch {
        Write-Error "Failed to connect to Microsoft Graph. Ensure proper permissions and MFA-enabled login."
        throw
    }
}

# Function: Save Results to Excel
Function Save-ResultsToExcel {
    param (
        [array]$Results,
        [string]$TenantName
    )
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $SaveFileDialog.Title = "Save Excel File"
        if ($SaveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $FilePath = $SaveFileDialog.FileName
            $Results | Export-Excel -Path $FilePath -AutoSize -Title "MFA Audit for $TenantName" -TableName "UserDetails"
            Write-Host "Results saved to $FilePath" -ForegroundColor Green
        } else {
            Write-Host "Save operation canceled." -ForegroundColor Yellow
        }
    } catch {
        Write-Error "Failed to save results to Excel: $($_.Exception.Message)"
    }
}

# Main Script Execution
Function Main {
    # Step 1: Verify and Install Modules
    Install-ModuleIfMissing -ModuleName "Microsoft.Graph"
    Install-ModuleIfMissing -ModuleName "ImportExcel"

    # Step 2: Connect to Microsoft Graph with Manual Login
    Connect-MicrosoftGraphWithBrowserLogin

    # Step 3: Fetch Users from Azure AD
    Write-Host "Fetching users from Azure AD..."
    $AllUsers = Get-MgUser -All -Property "Id,DisplayName,UserPrincipalName,UserType,AccountEnabled,CreatedDateTime" | Select-Object *

    # Step 4: Prompt for Scope Selection
    $ScopeOption = Read-Host "Select account scope: (1) Single User, (2) Active Members Only, (3) All Accounts"

    if ($ScopeOption -eq "1") {
        $Email = Read-Host "Enter the UserPrincipalName (email) of the user"
        $FilteredUsers = $AllUsers | Where-Object { $_.UserPrincipalName -eq $Email }
    } elseif ($ScopeOption -eq "2") {
        $FilteredUsers = $AllUsers | Where-Object { $_.AccountEnabled -eq $true -and $_.UserType -eq "Member" }
    } else {
        $FilteredUsers = $AllUsers
    }

    if (-not $FilteredUsers) {
        Write-Host "No users found for the selected criteria." -ForegroundColor Red
        return
    }

    # Step 5: Process Users and Retrieve MFA Data
    $Results = [System.Collections.Generic.List[Object]]::new()
    $Progress = 0
    foreach ($User in $FilteredUsers) {
        $Progress++
        Write-Progress -Activity "Processing Users" -Status "Processing $Progress of $($FilteredUsers.Count)" -PercentComplete (($Progress / $FilteredUsers.Count) * 100)

        try {
            $MFAStatus = "Disabled"
            $MFADefaultMethod = "None"
            $AuthMethods = (Get-MgUserAuthenticationMethod -UserId $User.Id)
            if ($AuthMethods) {
                $MFAStatus = "Enabled"
                $MFADefaultMethod = $AuthMethods | Where-Object { $_.IsDefault -eq $true } | Select-Object -ExpandProperty DisplayName
            }

            $Results.Add([PSCustomObject]@{
                UserPrincipalName = $User.UserPrincipalName
                DisplayName       = $User.DisplayName
                MFAStatus         = $MFAStatus
                MFADefaultMethod  = $MFADefaultMethod
                UserType          = $User.UserType
                AccountStatus     = if ($User.AccountEnabled) { "Enabled" } else { "Disabled" }
                CreatedDate       = $User.CreatedDateTime.ToString("dd-MM-yyyy")
            })
        } catch {
            Write-Warning "Failed to retrieve MFA data for $($User.UserPrincipalName): $($_.Exception.Message)"
        }
    }

    # Step 6: Output Results to Console
    $Results | Format-Table -Property UserPrincipalName, DisplayName, MFAStatus, MFADefaultMethod, UserType, AccountStatus, CreatedDate -AutoSize

    # Step 7: Prompt to Save Results to Excel
    $SaveToExcel = Read-Host "Do you want to save the results to an Excel file? (Y/N)"
    if ($SaveToExcel -eq "Y") {
        $TenantName = (Get-MgOrganization | Select-Object -ExpandProperty DisplayName)
        Save-ResultsToExcel -Results $Results -TenantName $TenantName
    }
}

# Execute Main Function
Main
