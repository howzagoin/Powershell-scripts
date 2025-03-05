# .NET Runtime Installation Manager
# Author: Tim MacLatchy
# Date: 14-Nov-2024
# License: MIT License
# Description: This script checks if .NET Runtime is installed and installs the latest available version.
# Steps: 
# - Check for installed .NET runtime versions
# - Download and install the latest version (version 9.0), only if not already installed.

# Required minimum version
$VERBOSE_LOGGING = $true
$LATEST_RUNTIME_VERSION = "9.0.0"  # Target version to install (latest version)

# Function to enable verbose logging
function Write-VerboseLog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [string]$Color = "White"
    )
    if ($VERBOSE_LOGGING) {
        Write-Host $Message -ForegroundColor $Color
    }
}

# Function to get the installed .NET Runtime versions
function Get-DotNetRuntimeVersions {
    [CmdletBinding()]
    param()
    
    try {
        Write-VerboseLog "Checking installed .NET Runtime versions..." "Cyan"
        $dotnetPath = (Get-Command "dotnet" -ErrorAction SilentlyContinue)
        
        if (-not $dotnetPath) {
            Write-VerboseLog "The .NET Runtime is not installed." "Yellow"
            return @()
        }

        # Get runtime versions and clean up the output
        $versions = & $dotnetPath.Source --list-runtimes | ForEach-Object {
            if ($_ -match "Microsoft\.NETCore\.App\s+(\d+\.\d+\.\d+)") {
                $matches[1]
            }
        } | Sort-Object -Unique

        if (-not $versions) {
            Write-VerboseLog "No .NET Runtime versions were detected." "Yellow"
            return @()
        }

        Write-VerboseLog "Found .NET Runtime versions: $($versions -join ', ')" "Green"
        return $versions
    } 
    catch {
        Write-Error "Failed to retrieve installed .NET Runtime versions: $_"
        Write-VerboseLog "Error stack trace: $($_.ScriptStackTrace)" "Red"
        return @()
    }
}

# Function to download and install the .NET Runtime using the provided link
function Install-DotNetRuntime {
    [CmdletBinding()]
    param (
        [string]$InstallerPath = "$env:TEMP\dotnet-runtime-installer.exe"
    )
    
    try {
        # Verify admin privileges
        $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (-not $isAdmin) {
            throw "This installation requires administrator privileges. Please run as administrator."
        }

        # Detect system architecture
        $systemType = if ([System.Environment]::Is64BitOperatingSystem) { "x64" } else { "x86" }
        Write-VerboseLog "System Architecture: $systemType" "Green"

        # Use the provided fixed download URL for the .NET Runtime 9.0
        $downloadUrl = "https://download.visualstudio.microsoft.com/download/pr/685792b6-4827-4dca-a971-bce5d7905170/1bf61b02151bc56e763dc711e45f0e1e/windowsdesktop-runtime-9.0.0-win-x64.exe"
        Write-VerboseLog "Downloading .NET Runtime installer from: $downloadUrl" "Green"

        # Clean up existing installer if present
        if (Test-Path $InstallerPath) { 
            Remove-Item -Path $InstallerPath -Force
            Write-VerboseLog "Removed existing installer file" "Gray"
        }

        # Download the installer using Invoke-WebRequest for better error handling
        try {
            $ProgressPreference = 'SilentlyContinue'  # Significantly improves download speed
            Invoke-WebRequest -Uri $downloadUrl -OutFile $InstallerPath -UseBasicParsing
            $ProgressPreference = 'Continue'
        }
        catch {
            throw "Failed to download the .NET Runtime installer. Please check the URL or try again later."
        }

        if (-not (Test-Path $InstallerPath)) {
            throw "Installer download failed. File not found at: $InstallerPath"
        }

        # Get installer information
        $installerInfo = Get-Item $InstallerPath
        $installerVersion = if ($installerInfo.VersionInfo.FileVersion) { $installerInfo.VersionInfo.FileVersion } else { $installerInfo.Name }
        Write-VerboseLog "Downloaded .NET Runtime installer version: $installerVersion" "Green"

        # Install the runtime
        Write-VerboseLog "Starting .NET Runtime installation..." "Yellow"
        $processParams = @{
            FilePath = $InstallerPath
            ArgumentList = "/quiet /norestart /log $env:TEMP\dotnet-install.log"
            Wait = $true
            NoNewWindow = $true
            PassThru = $true
        }
        
        $process = Start-Process @processParams
        
        if ($process.ExitCode -eq 0) {
            Write-VerboseLog ".NET Runtime installation completed successfully." "Green"
        } else {
            throw "Installation failed with exit code: $($process.ExitCode). Check logs at: $env:TEMP\dotnet-install.log"
        }
    }
    catch {
        Write-Error "Failed to download or install .NET Runtime: $_"
        Write-VerboseLog "Error stack trace: $($_.ScriptStackTrace)" "Red"
        throw
    }
    finally {
        # Cleanup
        if (Test-Path $InstallerPath) {
            Remove-Item -Path $InstallerPath -Force -ErrorAction SilentlyContinue
        }
    }
}

# Main script logic
function Start-DotNetInstallation {
    [CmdletBinding()]
    param(
        [switch]$Force
    )

    try {
        Write-VerboseLog "Starting .NET Runtime installation..." "White"
        
        $installedVersions = Get-DotNetRuntimeVersions
        Write-VerboseLog "Currently installed .NET Runtime versions: $($installedVersions -join ', ')" "Yellow"

        # Check if the latest version is already installed
        if ($installedVersions -contains $LATEST_RUNTIME_VERSION) {
            Write-VerboseLog ".NET Runtime version $LATEST_RUNTIME_VERSION is already installed. Skipping installation." "Green"
            return $true
        }

        Write-VerboseLog "Installing latest .NET Runtime version ($LATEST_RUNTIME_VERSION)..." "Yellow"
        Install-DotNetRuntime
    }
    catch {
        Write-Error "Installation process failed: $_"
        Write-VerboseLog "Error stack trace: $($_.ScriptStackTrace)" "Red"
        return $false
    }
    return $true
}

# Entry point with error handling
try {
    $result = Start-DotNetInstallation
    if ($result) {
        Write-VerboseLog "Script completed successfully." "Green"
    } else {
        Write-VerboseLog "Script completed with errors." "Red"
    }
}
catch {
    Write-Error "Script failed: $_"
    Write-VerboseLog "Error stack trace: $($_.ScriptStackTrace)" "Red"
    exit 1
}
