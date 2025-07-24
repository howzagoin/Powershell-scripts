
# M365Utils.psm1 - A module of utility functions for Microsoft 365 administration.

#region Logging
Function Write-Log {
    [CmdletBinding()]
    Param (
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Critical')]
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
#endregion

#region Module Management
Function Install-RequiredModules {
    [CmdletBinding()]
    Param(
        [string[]]$ModuleNames
    )

    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
    $scope = if ($isAdmin) { "AllUsers" } else { "CurrentUser" }

    foreach ($moduleName in $ModuleNames) {
        try {
            Write-Log "Checking module: $moduleName" -Level Info
            if (-not (Get-Module -Name $moduleName -ListAvailable)) {
                Write-Log "Installing module: $moduleName" -Level Info
                Install-Module -Name $moduleName -Scope $scope -Force -AllowClobber
            }

            $currentVersion = (Get-Module -Name $moduleName -ListAvailable | Sort-Object Version -Descending)[0].Version
            $latestVersion = (Find-Module -Name $moduleName).Version

            if ($currentVersion -lt $latestVersion) {
                Write-Log "Updating $moduleName from $currentVersion to $latestVersion" -Level Info
                Update-Module -Name $moduleName -Force
            }

            Import-Module -Name $moduleName -Force -ErrorAction Stop
            Write-Log "Module $moduleName imported successfully" -Level Info
        }
        catch {
            Write-Log "Error processing module ${moduleName}: $_" -Level Error
            throw
        }
    }
}
#endregion

#region Connection Management
Function Connect-M365Services {
    [CmdletBinding()]
    Param()

    # Disconnect any existing sessions
    Try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Log "Disconnected any existing Microsoft 365 sessions." -Level "Info"
    } Catch {
        Write-Log "No active sessions to disconnect or an error occurred during disconnect: $_" -Level "Warning"
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
        Write-Log "Successfully authenticated to Microsoft Graph using browser authentication." -Level "Info"

        # Output and log the scopes for testing
        $activeScopes = (Get-MgContext).Scopes
        Write-Log "Active Microsoft Graph scopes: $($activeScopes -join ', ')" -Level "Info"
        Write-Host "Active Microsoft Graph scopes: $($activeScopes -join ', ')" -ForegroundColor Cyan

        # Connect to Exchange Online
        Connect-ExchangeOnline -ShowProgress:$true -ErrorAction Stop
        Write-Log "Successfully authenticated to Exchange Online using browser authentication." -Level "Info"

        return $true
    } Catch {
        Write-Log "Authentication failed: $_" -Level "Error"
        Write-Host "Unable to authenticate. Please check your credentials and try again." -ForegroundColor Red
        return $false
    }
}
#endregion

#region File Dialog
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
            Write-Log "User selected file path: $($saveFileDialog.FileName)" -Level Info
            return $saveFileDialog.FileName
        }

        Write-Log "File save operation canceled by user" -Level Warning
        return $null
    }
    catch {
        Write-Log "Error in save file dialog: $_" -Level Error
        return $null
    }
}
#endregion

#region Excel Export
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
            Write-Log "Data exported successfully to $FilePath" -Level Info
            return $true
        } else {
            throw "Export file not found after operation"
        }
    } catch {
        Write-Log "Excel export error: $_" -Level Error
        return $false
    }
}
#endregion

Export-ModuleMember -Function Write-Log, Install-RequiredModules, Connect-M365Services, Prompt-SaveFileDialog, Export-ResultsToExcel
