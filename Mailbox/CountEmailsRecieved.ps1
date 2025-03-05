<#
.SYNOPSIS
Script to scan a mailbox in Microsoft 365 using Microsoft Graph.

.AUTHOR
Timothy MacLatchy

.DATE
2024-11-25

.LICENSE
MIT License

.DESCRIPTION
This script scans a specified mailbox for emails within a date range using Microsoft Graph and exports results to an Excel file.

#>

# Function to ensure required modules are installed
function Ensure-RequiredModules {
    param (
        [string[]]$RequiredModules = @('Microsoft.Graph', 'ImportExcel')
    )
    foreach ($module in $RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Host "Installing module: $module..." -ForegroundColor Yellow
            Install-Module -Name $module -Force -Scope CurrentUser -AllowClobber
        }
        # Import the module
        Import-Module $module -ErrorAction SilentlyContinue
        Write-Host "Module $module is ready." -ForegroundColor Green
    }
}

# Function to authenticate with Microsoft Graph
function Connect-MicrosoftGraph {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    try {
        Connect-MgGraph -Scopes "Mail.Read" -ErrorAction Stop
        Write-Host "Connected successfully!" -ForegroundColor Green
    } catch {
        throw "Failed to authenticate with Microsoft Graph. Check your credentials and permissions."
    }
}

# Function to calculate predefined date ranges
function Get-DateRange {
    param (
        [string]$RangeOption
    )
    $today = Get-Date
    switch ($RangeOption) {
        "1" { return @{ StartDate = ($today).AddDays(-7); EndDate = $today } } # Last week
        "2" { return @{ StartDate = ($today).AddMonths(-1); EndDate = $today } } # Last month
        "3" { return @{ StartDate = ($today).AddMonths(-6); EndDate = $today } } # Last 6 months
        "4" { return @{ StartDate = (Get-Date -Year $today.Year -Month 1 -Day 1); EndDate = $today } } # Year to date
        "5" { return @{ StartDate = (Get-Date -Year ($today.Year - 1) -Month 1 -Day 1); EndDate = (Get-Date -Year ($today.Year - 1) -Month 12 -Day 31) } } # Last year
        "6" { 
            Write-Host "Enter custom date range:" -ForegroundColor Cyan
            $StartDate = [DateTime]::ParseExact((Read-Host "Enter start date (yyyy-MM-dd)"), 'yyyy-MM-dd', $null)
            $EndDate = [DateTime]::ParseExact((Read-Host "Enter end date (yyyy-MM-dd)"), 'yyyy-MM-dd', $null)
            if ($EndDate -lt $StartDate) {
                throw "End date cannot be earlier than start date."
            }
            return @{ StartDate = $StartDate; EndDate = $EndDate }
        }
        default { throw "Invalid option selected." }
    }
}

# Function to retrieve mailbox messages using Microsoft Graph
function Get-MailboxMessages {
    param (
        [string]$Mailbox,
        [datetime]$StartDate,
        [datetime]$EndDate
    )

    Write-Host "Fetching emails for $Mailbox..." -ForegroundColor Cyan

    # Format dates for Graph API filter
    $startDateFilter = $StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $endDateFilter = $EndDate.ToString("yyyy-MM-ddTHH:mm:ssZ")

    try {
        # Query messages using Microsoft Graph
        $messages = Get-MgUserMessage -UserId $Mailbox -Filter "receivedDateTime ge $startDateFilter and receivedDateTime le $endDateFilter" -All

        if ($messages.Count -eq 0) {
            Write-Warning "No messages found in the specified date range."
            return @()
        }

        return $messages
    } catch {
        throw "Error retrieving messages: $_"
    }
}

# Function to export messages to Excel
function Export-MessagesToExcel {
    param (
        [array]$Messages,
        [string]$FilePath
    )

    try {
        # Prepare data for Excel
        $data = $Messages | ForEach-Object {
            @{
                Subject      = $_.Subject
                Sender       = $_.From.EmailAddress.Address
                ReceivedDate = $_.ReceivedDateTime
                BodyPreview  = $_.BodyPreview
            }
        }

        # Export to Excel
        $data | Export-Excel -Path $FilePath -AutoSize -Title "Mailbox Search Results" -BoldTopRow -FreezeTopRow -AutoFilter
        Write-Host "Exported results to $FilePath" -ForegroundColor Green
    } catch {
        throw "Failed to export to Excel: $_"
    }
}

# Main Execution Block
function Main {
    try {
        Ensure-RequiredModules
        Connect-MicrosoftGraph

        # Prompt for mailbox and date range
        $Mailbox = Read-Host "Enter the mailbox to scan"
        Write-Host "Select date range:"
        Write-Host "1. Last week | 2. Last month | 3. Last 6 months | 4. Year to date | 5. Last year | 6. Custom"
        $option = Read-Host "Choose an option (1-6)"

        # Get the date range
        $DateRange = Get-DateRange -RangeOption $option

        # Fetch messages
        $messages = Get-MailboxMessages -Mailbox $Mailbox -StartDate $DateRange.StartDate -EndDate $DateRange.EndDate
        if ($messages.Count -eq 0) {
            Write-Warning "No messages to export. Exiting script."
            return
        }

        # Export to Excel
        $savePath = (Read-Host "Choose a location to save the report").Replace('"', '')
        $fileName = "MailboxSearchResults_{0}.xlsx" -f (Get-Date -Format 'yyyyMMddHHmmss')
        $filePath = Join-Path -Path $savePath -ChildPath $fileName
        Export-MessagesToExcel -Messages $messages -FilePath $filePath
    } catch {
        Write-Host "Error: $_" -ForegroundColor Red
    }
}

# Execute the script
Main
