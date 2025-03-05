# Author: Tim MacLatchy
# Date: 19-Nov-2024
# License: MIT License
# Description: This script retrieves inactive, unknown, and disabled accounts with licenses from a Microsoft 365 tenant
# and categorizes them into:
#    - Disabled accounts with licenses (classed as disabled)
#    - Active accounts with licenses, but no login >90 days (classed as inactive)
#    - Active accounts with licenses and no last sign-in date available (classed as unknown)
# The results are exported to an Excel file if needed.
# Steps:
#    1. Connect to Microsoft Graph using MFA.
#    2. Fetch users with license details.
#    3. Categorize users into disabled, inactive, or unknown accounts.
#    4. Prompt for file location to save results.
#    5. Export the categorized results to Excel with proper formatting.


# Check for required modules and install if necessary
$requiredModules = @("Microsoft.Graph", "ImportExcel")
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Module $module is not installed. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name $module -Force -Scope CurrentUser -ErrorAction Stop
            Write-Host "Successfully installed $module" -ForegroundColor Green
        } catch {
            throw "Failed to install $module. Error: $($_.Exception.Message)"
        }
    }
}

Function Connect-ToMicrosoftGraph {
    param (
        [int]$Retries = 3
    )

    $LoginSuccess = $false
    $Attempt = 0

    while (-not $LoginSuccess -and $Attempt -lt $Retries) {
        $Attempt++
        Write-Host "Attempt ${Attempt}: Connecting to Microsoft Graph..." -ForegroundColor Yellow
        try {
            # Explicitly prompt the user for interactive login
            Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All", "Organization.Read.All" -ErrorAction Stop
            $LoginSuccess = $true
            Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
        } catch {
            Write-Host "Attempt ${Attempt} failed. Error: $($_.Exception.Message)" -ForegroundColor Red
            if ($Attempt -lt $Retries) {
                Write-Host "Retrying in 5 seconds..."
                Start-Sleep -Seconds 5
            }
        }
    }

    if (-not $LoginSuccess) {
        throw "Failed to connect to Microsoft Graph after $Retries attempts."
    }
}

Function Save-ResultsToExcel {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Results
    )

    Write-Host "Saving results to Excel..."
    try {
        Add-Type -AssemblyName System.Windows.Forms

        # Initialize SaveFileDialog
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        if (-not $saveFileDialog) {
            throw "Failed to create SaveFileDialog object."
        }

        $saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $saveFileDialog.Title = "Save Excel File"

        # Set default filename
        $currentDate = Get-Date -Format "dd-MM-yyyy"
        $tenantName = (Get-MgOrganization).DisplayName -replace "\s", "_"
        $saveFileDialog.FileName = "${tenantName}_${currentDate}_disabled_inactive_unknown_licensed_accounts.xlsx"
        $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

        # Create a new form to enforce TopMost behavior
        $form = New-Object System.Windows.Forms.Form
        $form.TopMost = $true
        $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None

        # Show the form briefly to ensure it's on top, then show the SaveFileDialog
        $form.Show()

        # Ensure form is in the front and activated
        $form.BringToFront()
        $form.Activate()

        # Show the dialog with the form as the parent
        $result = $saveFileDialog.ShowDialog($form)

        # Close the form after the dialog has finished
        $form.Close()

        if ($result -ne [System.Windows.Forms.DialogResult]::OK -or [string]::IsNullOrWhiteSpace($saveFileDialog.FileName)) {
            Write-Host "Save operation cancelled by user." -ForegroundColor Yellow
            return
        }

        $filePath = $saveFileDialog.FileName

        # Log the file path for debugging
        Write-Host "File path to save: $filePath"

        # Prepare the data for export
        $dataTable = @()

        # Loop through the categories (Deactivated, Inactive, Unknown)
        foreach ($category in $Results.Keys) {
            # For each user list under a category, create a table row
            foreach ($user in $Results[$category]) {
                $row = New-Object PSObject -Property @{
                    "Category"          = $category
                    "UserPrincipalName" = $user.UserPrincipalName
                    "DisplayName"       = $user.DisplayName
                    "LastSignInDate"    = $user.LastSignInDate
                    "Licenses"          = $user.Licenses
                }
                $dataTable += $row
            }
        }

        # Log the data to be exported
        Write-Host "Data to be exported:"
        $dataTable | Format-Table -AutoSize

        # Export to Excel with added formatting
        $dataTable | Export-Excel -Path $filePath -WorksheetName "Results" -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter

        Write-Host "Results successfully saved to $filePath" -ForegroundColor Green
    } catch {
        Write-Host "Error saving to Excel: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Gray
        throw
    }
}



Function Fetch-UserData {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$LicenseMap
    )

    Write-Host "Retrieving user accounts..." -ForegroundColor Yellow
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

# Main script execution
try {
    Connect-ToMicrosoftGraph

    # Initialize license information
    Write-Host "Retrieving license information..." -ForegroundColor Yellow
    $subscriptions = Get-MgSubscribedSku
    $licenseMap = @{}
    foreach ($sub in $subscriptions) {
        $licenseMap[$sub.SkuId] = $sub.SkuPartNumber
    }
    Write-Host "Successfully retrieved license information" -ForegroundColor Green

    # Fetch user data
    $userResults = Fetch-UserData -LicenseMap $licenseMap

    # Display results
    foreach ($key in $userResults.Keys) {
        Write-Host "`n${key}:" -ForegroundColor Cyan
        $userResults[$key] | Format-Table -AutoSize
    }

    # Prompt for Excel export
    $proceedToSave = Read-Host "Do you want to save this data to an Excel file? (Y/N)"
    if ($proceedToSave -eq "Y") {
        Save-ResultsToExcel -Results $userResults
    }

} catch {
    Write-Host "Critical error: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not properly disconnect from Microsoft Graph" -ForegroundColor Yellow
    }
}
