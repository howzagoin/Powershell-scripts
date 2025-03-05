<#
MIT License

Copyright (c) 2024 Timothy MacLatchy

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Author: Timothy MacLatchy
Date: September 6, 2024

Script Description:
This PowerShell script performs the following actions:
1. Ensures that the Exchange Online Management module is installed and imported.
2. Prompts the user for an admin username to log into Exchange Online with MFA.
3. Asks for a GUID and retrieves the recipient object details associated with that GUID.
4. Displays the details of the object in a formatted list.
5. Handles any errors during login or recipient lookup and disconnects after execution.

#>

# Ensure Exchange Online Management Module is installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Force
}

# Import the module
Import-Module ExchangeOnlineManagement

# Prompt for admin login and connect to Exchange Online with MFA
try {
    Connect-ExchangeOnline -UserPrincipalName (Read-Host "Enter Admin Username") -ShowProgress $true
} catch {
    Write-Host "Failed to connect to Exchange Online. Please check your credentials." -ForegroundColor Red
    exit
}

# Prompt for GUID input
$guid = Read-Host "Enter the GUID (e.g., f3cf8d34-1a45-4740-9429-42e2a79f0d1e)"

# Try to get recipient details from the provided GUID
try {
    $recipient = Get-Recipient -Identity $guid
    if ($recipient) {
        Write-Host "Details for GUID: $guid"
        $recipient | Format-List
    } else {
        Write-Host "No recipient found for the provided GUID." -ForegroundColor Yellow
    }
} catch {
    Write-Host "Error retrieving recipient details. Please verify the GUID and try again." -ForegroundColor Red
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
