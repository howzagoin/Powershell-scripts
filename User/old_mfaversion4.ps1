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
        Save-ResultsToExcel -results $Report
    }
}

# Execute Main Function
Main
