# Script Metadata
# Author: Tim MacLatchy
# Date: 17-12-2024
# License: MIT
# Description: Retrieves CVE vulnerabilities and recommendations from Microsoft Defender for Endpoint and exports them to Excel
# Steps:
#   1. Verify and install required modules
#   2. Authenticate to Microsoft Defender API
#   3. Retrieve CVE vulnerabilities data
#   4. Export data to formatted Excel file

#Requires -Version 7.0
using namespace System.Windows.Forms

function Log-And-Execute {
    param(
        [string]$FunctionName,
        [scriptblock]$ScriptBlock
    )
    try {
        Write-Host "$(Get-Date -Format 'dd-MM-yyyy HH:mm:ss'): Executing ${FunctionName}"
        & $ScriptBlock
        Write-Host "$(Get-Date -Format 'dd-MM-yyyy HH:mm:ss'): Completed ${FunctionName}"
    }
    catch {
        Write-Error "$(Get-Date -Format 'dd-MM-yyyy HH:mm:ss'): Error in ${FunctionName}: ${_}"
        throw
    }
}

function Initialize-RequiredModules {
    $requiredModules = @('Microsoft.Graph', 'ImportExcel')
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Host "Installing ${module} module..."
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
        }
    }
    Import-Module -Name $requiredModules -ErrorAction Stop
}

function Connect-ToDefenderAPI {
    try {
        # Define the required scopes for Defender API
        $scopes = @(
            "https://api.security.microsoft.com/.default"
        )
        
        Write-Host "Authenticating to Microsoft Defender API..."
        $authResult = Connect-MgGraph -Scopes $scopes
        if (-not $authResult) {
            throw "Failed to authenticate to Microsoft Defender API"
        }
        
        Write-Host "Successfully authenticated to Microsoft Defender API."
        return $authResult
    }
    catch {
        Write-Error "Failed to authenticate to Microsoft Defender API: $_"
        throw
    }
}

function Get-DefenderCVEData {
    try {
        Ensure-ActiveSession
        $endpoint = "https://api.security.microsoft.com/api/vulnerabilities/cves"

        Write-Host "Retrieving CVE vulnerabilities data from Defender API..."
        $response = Invoke-MgGraphRequest -Uri $endpoint -Method GET
        if (-not $response.value) {
            throw "No data retrieved from Microsoft Defender API"
        }

        # Extract relevant data
        $data = foreach ($item in $response.value) {
            [PSCustomObject]@{
                "CVE ID"             = $item.id
                "Affected Software"  = ($item.software | ForEach-Object { $_.name }) -join ', '
                "Severity Level"     = $item.severity
                "Exposed Devices"    = $item.exposedDevicesCount
                "User"               = $item.assignedTo
                "Recommended Action" = $item.remediation
            }
        }
        Write-Host "Successfully retrieved CVE vulnerabilities data."
        return $data
    }
    catch {
        Write-Error "Failed to retrieve CVE vulnerabilities data: $_"
        throw
    }
}

function Ensure-ActiveSession {
    try {
        $mgContext = Get-MgContext
        if (-not $mgContext -or -not $mgContext.AccessToken -or -not $mgContext.TokenExpiry -or ($mgContext.TokenExpiry -lt (Get-Date).AddMinutes(5))) {
            Write-Host "No active token found or token is near expiry. Reconnecting..."
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Connect-ToDefenderAPI
        }
    }
    catch {
        Write-Error "Failed to ensure an active session: $_"
        throw
    }
}

function Export-CVEDataToExcel {
    param (
        [Parameter(Mandatory)]
        [array]$Data
    )
    try {
        $saveDialog = [SaveFileDialog]::new()
        $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $saveDialog.FileName = "DefenderCVEData_$(Get-Date -Format 'dd-MM-yyyy').xlsx"
        
        if ($saveDialog.ShowDialog() -eq [DialogResult]::OK) {
            $excelParams = @{
                Path = $saveDialog.FileName
                AutoSize = $true
                AutoFilter = $true
                FreezeTopRow = $true
                BoldTopRow = $true
                TableStyle = 'Medium2'
                WorksheetName = 'CVE Vulnerabilities'
            }
            
            $Data | Export-Excel @excelParams
            Write-Host "Data exported successfully to $($saveDialog.FileName)"
        }
    }
    catch {
        Write-Error "Failed to export data: $_"
        throw
    }
}

function Main {
    try {
        Write-Host "Starting Microsoft Defender CVE Export Script..."
        Log-And-Execute -FunctionName "Initialize-RequiredModules" -ScriptBlock { Initialize-RequiredModules }
        Log-And-Execute -FunctionName "Connect-ToDefenderAPI" -ScriptBlock { Connect-ToDefenderAPI }
        Log-And-Execute -FunctionName "Get-DefenderCVEData" -ScriptBlock {
            $cveData = Get-DefenderCVEData
        }
        if ($cveData) {
            Log-And-Execute -FunctionName "Export-CVEDataToExcel" -ScriptBlock {
                Export-CVEDataToExcel -Data $cveData
            }
        }
        else {
            Write-Host "No data available to export."
        }
        Write-Host "Script execution completed successfully."
    }
    catch {
        Write-Error "Script execution failed: $_"
        exit 1
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
}

Main
