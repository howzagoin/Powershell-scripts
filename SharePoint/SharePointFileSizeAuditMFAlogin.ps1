# ==========================================
# Author: Tim MacLatchy (Enhanced by Claude)
# Date: 14-Nov-2024
# License: MIT License
# Copyright (c) 2024 Tim MacLatchy
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files.
# Description: Script to audit SharePoint file sizes using MFA-secure login and Microsoft Graph API
# ==========================================
# Requires -Version 5.1
using namespace System.Windows.Forms
using namespace System.Drawing

$VerbosePreference = "Continue"

# Configuration block for easy customization
$script:Config = @{
    MinFileSizeMB = 100
    BatchSize = 999  # Graph API typically limits to 999 items per request
    ExcludedLists = @(
        "Form Templates"
        "Preservation Hold Library"
        "Site Assets"
        "Pages"
        "Images"
        "Style Library"
    )
    RequiredModules = @{
        'Microsoft.Graph' = '1.9.0'
        'ImportExcel' = '7.0.0'
    }
}

function Initialize-ScriptRequirements {
    [CmdletBinding()]
    param()

    Write-Verbose "Initializing script requirements..."

    # Add error handling for PowerShell edition check
    try {
        if ($PSVersionTable.PSEdition -notin @('Core', 'Desktop')) {
            throw "This script requires PowerShell Core or Desktop Edition. Current edition: $($PSVersionTable.PSEdition)"
        }
    }
    catch {
        throw "Failed to verify PowerShell edition: $_"
    }

    # Loop through required modules and ensure they are installed and imported
    foreach ($module in $script:Config.RequiredModules.GetEnumerator()) {
        try {
            Write-Verbose "Checking for existing module: $($module.Key)"
            
            # Check if module is installed
            $existingModule = Get-Module -Name $module.Key -ListAvailable | 
                Sort-Object Version -Descending | 
                Select-Object -First 1

            if (-not $existingModule -or $existingModule.Version -lt [version]$module.Value) {
                Write-Verbose "Module $($module.Key) not found or version is outdated. Installing version $($module.Value)..."
                
                # Set a timeout for the installation process
                $timeout = 60  # Timeout in seconds
                $startTime = Get-Date

                # Install the module and handle potential blocking scenarios
                $installJob = Start-Job -ScriptBlock {
                    Install-Module -Name $using:module.Key -MinimumVersion $using:module.Value -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                }

                # Wait for the installation job to complete, checking for timeout
                while ($installJob.State -eq 'Running' -and ((Get-Date) - $startTime).TotalSeconds -lt $timeout) {
                    Start-Sleep -Seconds 5
                }

                # Check if the job has completed, otherwise kill it
                if ($installJob.State -eq 'Running') {
                    Write-Error "Module installation for $($module.Key) timed out after $timeout seconds."
                    Stop-Job -Job $installJob
                    throw "Module $($module.Key) installation failed due to timeout."
                }
                
                # Collect results from the job and clean up
                $installResult = Receive-Job -Job $installJob -ErrorAction Stop
                Remove-Job -Job $installJob

                if ($installResult) {
                    Write-Verbose "Module $($module.Key) installed successfully."
                }
                else {
                    Write-Warning "Module $($module.Key) failed to install."
                    throw "Module $($module.Key) could not be installed."
                }
            }
            else {
                Write-Verbose "Module $($module.Key) is already installed with the required version."
            }

            Write-Verbose "Importing module $($module.Key)..."
            Import-Module -Name $module.Key -MinimumVersion $module.Value -Force -ErrorAction Stop
        }
        catch {
            Write-Error "Failed to install/import ${module.Key}: $_"
            throw
        }
    }
}

function Connect-ToMicrosoftGraph {
    [CmdletBinding()]
    param()

    Write-Host "Starting Microsoft Graph connection..."
    Write-Verbose "Starting Microsoft Graph connection..."

    try {
        # Log the disconnect action to the console
        Write-Host "Disconnecting any previous Microsoft Graph sessions..."
        Write-Verbose "Disconnecting any previous Microsoft Graph sessions..."
        Disconnect-MgGraph -ErrorAction SilentlyContinue

        # Define required scopes
        $requiredScopes = @(
            "Sites.Read.All",
            "Files.Read.All",
            "Organization.Read.All"
        )
        
        Write-Host "Attempting to connect to Microsoft Graph with the following scopes: $($requiredScopes -join ', ')"
        Write-Verbose "Attempting to connect to Microsoft Graph with the following scopes: $($requiredScopes -join ', ')"

        # Attempt login (this will prompt for MFA automatically if configured)
        try {
            Write-Host "Attempting login with MFA..."
            Write-Verbose "Attempting login with MFA..."
            Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop

            Write-Host "Successfully connected using login with MFA."
            Write-Verbose "Successfully connected using login with MFA."
        }
        catch {
            Write-Error "Login with MFA failed. Ensure your credentials and MFA settings are correct. Error: $_"
            Write-Verbose "Login with MFA failed. Ensure your credentials and MFA settings are correct. Error: $_"
            return $null
        }

        # Fetch organization information to confirm the connection
        Write-Host "Fetching organization information..."
        Write-Verbose "Fetching organization information..."
        $orgInfo = Get-MgOrganization -ErrorAction Stop

        if (-not $orgInfo) {
            throw "Failed to retrieve organization information after successful connection"
        }

        Write-Host "Successfully connected to tenant: $($orgInfo.DisplayName)"
        Write-Verbose "Successfully connected to tenant: $($orgInfo.DisplayName)"
        
        return @{
            TenantId = $orgInfo.Id
            TenantName = $orgInfo.DisplayName
        }
    }
    catch {
        Write-Error "Microsoft Graph connection failed: $_"
        Write-Verbose "Microsoft Graph connection failed: $_"
        return $null
    }
}

function Get-SharePointSites {
    Write-Host "Retrieving SharePoint sites..."
    Write-Verbose "Retrieving SharePoint sites..."

    try {
        # Specify the regular tenant SharePoint domain, not the admin center URL
        $tenantDomain = "waverleygymnastcentre.sharepoint.com"  # Use the regular domain here
        $sites = Get-MgSite -Filter "siteCollection/hostname eq '$tenantDomain'" -Top 100

        if (-not $sites) {
            throw "No SharePoint sites were found for $tenantDomain."
        }

        Write-Host "Successfully retrieved SharePoint sites."
        Write-Verbose "Successfully retrieved SharePoint sites."

        return $sites
    }
    catch {
        Write-Error "Failed to retrieve SharePoint sites: $_"
        Write-Verbose "Failed to retrieve SharePoint sites: $_"

        if ($_.Exception.Response) {
            $response = $_.Exception.Response
            Write-Host "Raw error response: $($response.Content)"
        }

        return $null
    }
}



function Get-LargeFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Sites,

        [Parameter()]
        [int64]$MinSizeBytes = ($script:Config.MinFileSizeMB * 1MB),

        [Parameter()]
        [string[]]$ExcludedLists = $script:Config.ExcludedLists
    )

    $results = [System.Collections.ArrayList]::new()
    $totalSites = $Sites.Count
    $processedSites = 0
    
    Write-Verbose "Scanning $totalSites sites for files larger than $($MinSizeBytes/1MB) MB..."

    foreach ($site in $Sites) {
        $processedSites++
        $percentComplete = [math]::Round(($processedSites / $totalSites) * 100)
        
        Write-Progress -Activity "Scanning SharePoint Sites" `
            -Status "Site $processedSites of ${totalSites}: $($site.WebUrl)" `
            -PercentComplete $percentComplete

        try {
            $libraries = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/v1.0/sites/$($site.Id)/lists" `
                -Query "filter=baseType eq 'documentLibrary' and hidden eq false"

            foreach ($library in $libraries.value) {
                if ($library.ItemCount -eq 0 -or $ExcludedLists -contains $library.DisplayName) {
                    continue
                }
                
                Write-Verbose "Processing library: $($library.DisplayName)"
                $libraryId = $library.Id
                $files = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/sites/$($site.Id)/lists/$($libraryId)/items" `
                    -Query "select=id,fields/fileLeafRef,fields/fileSize"

                foreach ($file in $files.value) {
                    if ($file.fields.fileSize -ge $MinSizeBytes) {
                        $results.Add([PSCustomObject]@{
                            Site = $site.WebUrl
                            Library = $library.DisplayName
                            FileName = $file.fields.fileLeafRef
                            FileSizeMB = [math]::Round($file.fields.fileSize / 1MB, 2)
                        })
                    }
                }
            }
        }
        catch {
            Write-Warning "Failed to process site $($site.WebUrl): $_"
            continue
        }
    }

    return $results
}

function Export-ToExcel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Data,

        [string]$FilePath
    )

    Write-Verbose "Exporting data to Excel..."

    try {
        $Data | Export-Excel -Path $FilePath -AutoSize -AutoFilter -FreezeTopRow -TableName 'LargeFiles' -ErrorAction Stop
        Write-Host "Data successfully exported to $FilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to export data to Excel: $_"
    }
}

# Main execution script
try {
    Initialize-ScriptRequirements

    $graphInfo = Connect-ToMicrosoftGraph
    $tenantId = $graphInfo.TenantId

    $sites = Get-SharePointSites

    $largeFiles = Get-LargeFiles -Sites $sites

    if ($largeFiles.Count -gt 0) {
        $filePath = [System.IO.Path]::Combine(
            [System.IO.Path]::GetTempPath(),
            "LargeFilesReport_$tenantId_$(Get-Date -Format 'dd-MM-yyyy').xlsx"
        )

        Export-ToExcel -Data $largeFiles -FilePath $filePath
    }
    else {
        Write-Host "No large files found." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Microsoft Graph connection failed: $_" -ForegroundColor Red
    Write-Verbose "Microsoft Graph connection failed: $_"
    throw
}
