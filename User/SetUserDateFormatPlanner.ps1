<#
.SYNOPSIS
Set Microsoft Planner default date format to en-GB for a specific user.
.DESCRIPTION
This script logs into Microsoft 365 as an administrator and sets the Planner default date format to en-GB for a selected user.
.AUTHOR
Your Name
.DATE
03-12-2024
.LICENSE
MIT License
#>

# Function to ensure required modules are installed
function Ensure-RequiredModule {
    param (
        [string]$ModuleName
    )
    
    try {
        if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
            Write-Host "Installing module: $ModuleName"
            Install-Module -Name $ModuleName -Force -Scope CurrentUser -AllowClobber
        }
        Import-Module -Name $ModuleName -Force
        Write-Host "Module $ModuleName is ready"
    }
    catch {
        Write-Error "Failed to install or import module ${ModuleName}: $_"
        exit 1
    }
}

# Connect to Microsoft Graph with admin credentials
function Connect-ToGraph {
    try {
        Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All", "Tasks.ReadWrite"
        Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to authenticate to Microsoft Graph: $_"
        exit 1
    }
}

# Function to update Planner regional settings
function Update-PlannerDateFormat {
    param (
        [string]$TargetUserEmail
    )
    
    try {
        # Get user
        $user = Get-MgUser -Filter "userPrincipalName eq '$TargetUserEmail'"
        
        if (-not $user) {
            Write-Error "User not found: $TargetUserEmail"
            return
        }

        # Attempt to update Planner settings
        $body = @{
            dateFormat = "dd/MM/yyyy"
        } | ConvertTo-Json

        # Construct URL for Planner settings
        $plannerSettingsUrl = "https://graph.microsoft.com/v1.0/users/$($user.Id)/planner/settings"

        # Make the request
        Write-Host "Updating Planner settings for $TargetUserEmail..."
        $response = Invoke-MgGraphRequest -Method PATCH -Uri $plannerSettingsUrl -Body $body
        
        if ($response -eq $null) {
            Write-Error "Failed to update Planner settings for $TargetUserEmail"
        } else {
            Write-Host "Successfully updated Planner date format for $TargetUserEmail" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Error updating Planner settings: $_"
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Main execution
try {
    # Install and import required module
    Ensure-RequiredModule -ModuleName "Microsoft.Graph"

    # Connect to Microsoft Graph
    Connect-ToGraph

    # Update Planner settings
    Update-PlannerDateFormat -TargetUserEmail "lylyna.im@firstfinancial.com.au"
}
catch {
    Write-Error "Script execution failed: $_"
    exit 1
}
