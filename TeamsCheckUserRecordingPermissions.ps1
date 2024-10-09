# Author: Tim MacLatchy
# Date: 01 October 2024
# License: MIT License
# Copyright (c) 2024 Tim MacLatchy
# Description: This script checks if users have recording and transcription permissions 
# in Microsoft Teams, using Microsoft Graph exclusively to avoid compatibility issues with PowerShell Core.

# Enable verbose output
$VerbosePreference = "Continue"

# Function to check and install required modules (only Microsoft.Graph now)
function Install-RequiredModule {
    param (
        [string]$ModuleName
    )
    Write-Verbose "Checking for ${ModuleName} module..."
    if (-not (Get-Module -ListAvailable -Name ${ModuleName})) {
        Write-Verbose "${ModuleName} module not found. Installing..."
        try {
            Install-Module -Name ${ModuleName} -Force -AllowClobber -Scope CurrentUser
            Write-Verbose "${ModuleName} module installed successfully."
        } catch {
            Write-Error "Failed to install ${ModuleName} module: $_"
            exit
        }
    } else {
        Write-Verbose "${ModuleName} module is already installed."
    }
}

# Check and install required module
Install-RequiredModule -ModuleName "Microsoft.Graph"

# Import only Microsoft.Graph
Import-Module Microsoft.Graph

# Function to connect to Microsoft Graph with web-based MFA
function Connect-Services {
    Write-Verbose "Connecting to Microsoft Graph..."
    try {
        Connect-MgGraph -Scopes "User.Read.All", "TeamsAppInstallation.ReadForUser.All" -UseDeviceAuthentication
        Write-Verbose "Successfully connected to Microsoft Graph."
    } catch {
        Write-Error "Failed to connect to Microsoft Graph: $_"
        exit
    }
}

# Connect to services
Connect-Services

# Function to get user license details using Microsoft Graph
function Get-UserLicense {
    param ($userId)
    Write-Verbose "Checking license for user ${userId}..."
    try {
        $licenses = (Get-MgUserLicenseDetail -UserId $userId).SkuPartNumber
        return $licenses | Where-Object { $_ -in $validLicenseSkus }
    } catch {
        Write-Error "Error checking license for user ${userId}: $_"
        return $null
    }
}

# Function to get Teams recording and transcription policies using Microsoft Graph
function Get-UserPolicies {
    param ($userId)
    Write-Verbose "Checking policies for user ${userId}..."
    try {
        # Placeholder for Microsoft Graph equivalent to get Teams policies
        $policies = Get-MgUserTeamsPolicy -UserId $userId
        return @(
            TranscriptionEnabled = $policies.AllowTranscription,
            RecordingEnabled = $policies.AllowCloudRecording
        )
    } catch {
        Write-Error "Error retrieving policies for ${userId}: $_"
        return $null
    }
}

# Get all users using Microsoft Graph
Write-Verbose "Retrieving all users in the tenant..."
try {
    $users = Get-MgUser -All
    Write-Verbose "Retrieved ${users.Count} users."
} catch {
    Write-Error "Failed to retrieve users: $_"
    exit
}

# Processing users and checking licenses/policies
$results = @()
$totalUsers = $users.Count
$currentUser = 0

foreach ($user in $users) {
    $currentUser++
    $percentComplete = ($currentUser / $totalUsers) * 100
    Write-Progress -Activity "Processing Users" -Status "User ${currentUser} of ${totalUsers}" -PercentComplete $percentComplete
    
    $userPrincipalName = $user.UserPrincipalName
    Write-Verbose "Processing user: ${userPrincipalName}"
    
    $userLicense = Get-UserLicense -userId $user.Id
    
    if ($userLicense) {
        $policies = Get-UserPolicies -userId $user.Id
        if ($policies) {
            $results += [PSCustomObject]@{
                UserPrincipalName = $userPrincipalName
                License = $userLicense -join ", "
                TranscriptionEnabled = $policies.TranscriptionEnabled
                RecordingEnabled = $policies.RecordingEnabled
            }
        }
    } else {
        Write-Verbose "${userPrincipalName} does not have a valid license for Teams features."
    }
}

# Output results to the console
$results | Format-Table -AutoSize

# Optionally export to Excel
# Prompt for Excel file location
Add-Type -AssemblyName System.Windows.Forms
$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
$saveFileDialog.FileName = "${tenantName}_${date}_TeamsPermissionsCheck.xlsx"
$saveFileDialog.Title = "Save Teams Permissions Check Results"
if ($saveFileDialog.ShowDialog() -eq 'OK') {
    $excelPath = $saveFileDialog.FileName
    Write-Verbose "Saving results to ${excelPath}"
    
    try {
        $results | Export-Excel -Path $excelPath -AutoSize -TableName "TeamsPermissions" -WorksheetName "Permissions" -Title "Teams Permissions Check Results" -TitleSize 18 -TitleBold
        Write-Verbose "Results exported successfully to Excel."
    } catch {
        Write-Error "Failed to export results to Excel: $_"
    }
} else {
    Write-Verbose "Excel export cancelled by user."
}

# Disconnect from Microsoft Graph
Write-Verbose "Disconnecting from Microsoft Graph..."
try {
    Disconnect-MgGraph
    Write-Verbose "Successfully disconnected."
} catch {
    Write-Error "Error during disconnect: $_"
}

Write-Verbose "Script execution completed."
