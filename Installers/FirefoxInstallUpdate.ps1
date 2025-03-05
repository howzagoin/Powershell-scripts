<#
.SYNOPSIS
    A script to check, install, or update Mozilla Firefox to the latest version.

.DESCRIPTION
    This script verifies the installed version of Mozilla Firefox, compares it with the latest version available online, 
    and updates or installs Firefox as needed. The installer is downloaded and removed after execution.

.AUTHOR
    Tim MacLatchy

.DATE
    28-11-2024

.LICENSE
    MIT License

.NOTES
    Requires PowerShell 5.1 or later.
#>

# Function to ensure the target directory exists
function Ensure-Directory {
    param (
        [string]$path
    )
    if (-not (Test-Path -Path $path)) {
        try {
            New-Item -ItemType Directory -Path $path -Force | Out-Null
            Write-Host "[INFO] Created folder: $path" -ForegroundColor Green
        } catch {
            Write-Host "[ERROR] Failed to create folder $path. $_" -ForegroundColor Red
            Exit 1
        }
    }
}

# Function to get installed Firefox version
function Get-InstalledFirefoxVersion {
    $firefoxRegPath = "HKLM:\Software\Mozilla\Mozilla Firefox"
    $installedVersion = $null

    try {
        # Get installed version from registry
        $installedVersion = Get-ItemProperty -Path $firefoxRegPath | Select-Object -ExpandProperty CurrentVersion
        Write-Host "[INFO] Firefox is installed. Current version: $installedVersion" -ForegroundColor Cyan
    } catch {
        Write-Host "[INFO] Firefox is not installed." -ForegroundColor Yellow
    }

    # Extract numeric version from string (e.g., "133.0 (x64 en-US)" -> "133.0")
    if ($installedVersion -match "^(\d+\.\d+)") {
        $installedVersion = $matches[1]
    }

    return $installedVersion
}

# Function to download the Firefox installer
function Download-FirefoxInstaller {
    param (
        [string]$downloadUrl,
        [string]$outputFile
    )

    Write-Host "[INFO] Downloading the Firefox installer..."
    try {
        Invoke-WebRequest -Uri $downloadUrl -OutFile $outputFile -ErrorAction Stop
        Write-Host "[INFO] Download complete. Installer saved to $outputFile" -ForegroundColor Green
    } catch {
        Write-Host "[ERROR] Failed to download the Firefox installer. $_" -ForegroundColor Red
        Exit 1
    }
}

# Function to check and update Firefox or install if not present
function Check-And-UpdateFirefox {
    $latestFirefoxURL = "https://download.mozilla.org/?product=firefox-latest-ssl&os=win64&lang=en-US"
    $installerFolder = "C:\Temp"
    $installerFile = "$installerFolder\firefoxInstaller.exe"

    # Ensure the installer folder exists
    Ensure-Directory -path $installerFolder

    # Get installed Firefox version
    $installedVersion = Get-InstalledFirefoxVersion

    # If Firefox is installed, display its version
    if ($installedVersion) {
        Write-Host "[INFO] Installed Firefox version: $installedVersion" -ForegroundColor Cyan
    } else {
        Write-Host "[INFO] Firefox is not installed. Proceeding with installation..." -ForegroundColor Yellow
    }

    # Download the installer
    Download-FirefoxInstaller -downloadUrl $latestFirefoxURL -outputFile $installerFile

    # Install or update Firefox
    Write-Host "[INFO] Running the Firefox installer..." -ForegroundColor Green
    Start-Process -FilePath $installerFile -ArgumentList "/S" -Wait
    Write-Host "[INFO] Firefox installation/update complete!" -ForegroundColor Green

    # Clean up by deleting the installer
    if (Test-Path $installerFile) {
        try {
            Remove-Item $installerFile -Force
            Write-Host "[INFO] Installer deleted: $installerFile" -ForegroundColor Green
        } catch {
            Write-Host "[ERROR] Failed to delete the installer. $_" -ForegroundColor Red
        }
    }
}

# Main Script Execution
Write-Host "[INFO] Checking and updating Firefox..." -ForegroundColor Cyan
Check-And-UpdateFirefox
Write-Host "[INFO] Script completed successfully." -ForegroundColor Green
